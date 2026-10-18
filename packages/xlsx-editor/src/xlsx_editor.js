// XlsxEditor: shows an Excel workbook on a canvas, as Excel draws it. The workbook is read and
// formatted by the Rust/WebAssembly engine of @dimarborda/xlsx; this class lays out the grid,
// scrolls it virtually (only the visible cells are drawn), keeps frozen panes in place and
// handles selection and sheet tabs.
import { createWorkbook, initEngine, openWorkbook } from '@dimarborda/xlsx';
import { canvasFont, drawHeaders, drawRegion, key } from './draw.js';
import { columnName, parseRange, sheetAxes } from './grid.js';

const COL_HEADER_H = 22;
const SELECTION = '#217346';

/** The language the cells are formatted in: the page's, else the browser's */
function defaultLocale() {
  if (typeof document !== 'undefined' && document.documentElement.lang) return document.documentElement.lang;
  if (typeof navigator !== 'undefined' && navigator.language) return navigator.language;
  return 'en';
}

async function toBytes(source) {
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  if (ArrayBuffer.isView(source)) return new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  if (typeof Blob !== 'undefined' && source instanceof Blob) return new Uint8Array(await source.arrayBuffer());
  if (typeof source === 'string' || source instanceof URL) {
    const res = await fetch(source);
    if (!res.ok) throw new Error(`No se pudo descargar ${source}: ${res.status} ${res.statusText}`);
    return new Uint8Array(await res.arrayBuffer());
  }
  throw new TypeError('open() espera una URL, un File/Blob, un ArrayBuffer o un Uint8Array.');
}

export class XlsxEditor extends EventTarget {
  /**
   * Creates an editor inside `container` (which should have a height)
   * @param {HTMLElement} container
   * @param {{locale?: string, zoom?: number, tabs?: boolean, formulaBar?: boolean, wasmUrl?: string | URL}} [options]
   */
  static async create(container, options = {}) {
    await initEngine(options.wasmUrl);
    return new XlsxEditor(container, options);
  }

  #container;
  #root;
  #bar;
  #nameBox;
  #valueBox;
  #body;
  #canvas;
  #ctx;
  #scroller;
  #spacer;
  #tabsEl;
  #observer;
  #frame = 0;
  #workbook = null;
  #styles = null;
  #sheet = null; // drawing model of the active sheet
  #sheetIndex = -1;
  #selection = null; // { first, last, active } (1-based rows and columns)
  #mdw = 7;
  #fontsListener = null;

  constructor(container, { locale, zoom = 1, tabs = true, formulaBar = true } = {}) {
    super();
    this.#container = container;
    this.locale = locale || defaultLocale();
    this.zoom = zoom;

    const root = document.createElement('div');
    root.className = 'xlsx-editor';
    root.innerHTML = `
      <div class="xlsx-editor-bar"${formulaBar ? '' : ' hidden'}>
        <div class="xlsx-editor-name" aria-label="Celda"></div>
        <div class="xlsx-editor-value" aria-label="Contenido"></div>
      </div>
      <div class="xlsx-editor-body">
        <canvas class="xlsx-editor-canvas"></canvas>
        <div class="xlsx-editor-scroller" tabindex="0" role="grid" aria-label="Hoja de cálculo"><div class="xlsx-editor-spacer"></div></div>
      </div>
      <div class="xlsx-editor-tabs" role="tablist"${tabs ? '' : ' hidden'}></div>`;
    container.appendChild(root);
    this.#root = root;
    this.#bar = root.querySelector('.xlsx-editor-bar');
    this.#nameBox = root.querySelector('.xlsx-editor-name');
    this.#valueBox = root.querySelector('.xlsx-editor-value');
    this.#body = root.querySelector('.xlsx-editor-body');
    this.#canvas = root.querySelector('canvas');
    this.#ctx = this.#canvas.getContext('2d');
    this.#scroller = root.querySelector('.xlsx-editor-scroller');
    this.#spacer = root.querySelector('.xlsx-editor-spacer');
    this.#tabsEl = root.querySelector('.xlsx-editor-tabs');

    this.#scroller.addEventListener('scroll', () => this.render());
    this.#scroller.addEventListener('pointerdown', e => this.#onPointerDown(e));
    this.#scroller.addEventListener('keydown', e => this.#onKeyDown(e));
    this.#observer = new ResizeObserver(() => this.render());
    this.#observer.observe(this.#body);
    if (document.fonts) {
      // Web fonts that arrive after the first paint change widths: measure and draw again
      this.#fontsListener = () => {
        if (this.#sheet) this.#loadSheet(this.#sheetIndex, { keepScroll: true });
      };
      document.fonts.addEventListener('loadingdone', this.#fontsListener);
    }
  }

  /** The open workbook (`@dimarborda/xlsx`), or null */
  get workbook() {
    return this.#workbook;
  }

  /** Position of the active sheet among the workbook's tabs */
  get sheetIndex() {
    return this.#sheetIndex;
  }

  get sheetName() {
    return this.#sheetIndex >= 0 ? this.#workbook.tabs[this.#sheetIndex].name : null;
  }

  /**
   * Opens a workbook
   * @param {string | URL | Blob | ArrayBuffer | Uint8Array} source
   */
  async open(source) {
    const workbook = await openWorkbook(await toBytes(source));
    this.#setWorkbook(workbook);
  }

  /** Starts a new, empty workbook */
  async createBlank() {
    this.#setWorkbook(await createWorkbook());
  }

  #setWorkbook(workbook) {
    this.#workbook?.close();
    this.#workbook = workbook;
    this.#styles = workbook.styleView();
    this.#sheetIndex = -1; // a new workbook starts at its first sheet, scrolled to the top
    const tabs = workbook.tabs;
    const first = tabs.findIndex(t => t.kind === 'worksheet' && t.state === 'visible');
    this.#renderTabs();
    this.#loadSheet(first >= 0 ? first : tabs.findIndex(t => t.kind === 'worksheet'));
    this.dispatchEvent(new CustomEvent('load', { detail: { sheets: tabs.map(t => t.name) } }));
  }

  /**
   * Shows another sheet
   * @param {string | number} nameOrIndex
   */
  setSheet(nameOrIndex) {
    const tabs = this.#workbook?.tabs ?? [];
    let index = nameOrIndex;
    if (typeof nameOrIndex === 'string') index = tabs.findIndex(t => t.name.toLowerCase() === nameOrIndex.toLowerCase());
    if (!tabs[index] || tabs[index].kind !== 'worksheet') throw new Error(`No hay ninguna hoja de cálculo '${nameOrIndex}'.`);
    if (index === this.#sheetIndex) return;
    this.#loadSheet(index);
    this.dispatchEvent(new CustomEvent('sheetchange', { detail: { index, name: tabs[index].name } }));
  }

  /** Reads the active sheet again from the workbook, after changing cells through `workbook` */
  reload() {
    if (!this.#workbook || this.#sheetIndex < 0) return;
    this.#styles = this.#workbook.styleView();
    this.#loadSheet(this.#sheetIndex, { keepScroll: true });
  }

  /** @param {number} zoom 0.25–4 */
  setZoom(zoom) {
    const previous = this.zoom;
    this.zoom = Math.min(4, Math.max(0.25, Number(zoom) || 1));
    this.#updateSpacer();
    // Keep the same cells in view
    const ratio = this.zoom / previous;
    this.#scroller.scrollLeft *= ratio;
    this.#scroller.scrollTop *= ratio;
    this.render();
  }

  /** Selects a cell or range (`B3`, `A1:C4`) and scrolls it into view */
  select(ref) {
    const range = parseRange(ref);
    this.#setSelection(range.first, range.last, range.first);
  }

  /** The selected range: `{ref, range, active, text}` or null */
  get selection() {
    if (!this.#selection) return null;
    const { first, last, active } = this.#selection;
    const ref = `${columnName(first.col)}${first.row}`;
    const range = first.row === last.row && first.col === last.col ? ref : `${ref}:${columnName(last.col)}${last.row}`;
    const cell = this.#sheet?.cells.get(key(active.row, active.col));
    return { ref: `${columnName(active.col)}${active.row}`, range, active: { ...active }, text: cell?.text ?? '' };
  }

  /** The workbook as .xlsx bytes */
  toBytes() {
    if (!this.#workbook) throw new Error('No hay ningún libro abierto.');
    return this.#workbook.toBytes();
  }

  /** Saves the workbook as a file in the browser */
  download(filename = 'libro.xlsx') {
    const blob = new Blob([this.toBytes()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  destroy() {
    cancelAnimationFrame(this.#frame);
    this.#observer.disconnect();
    if (this.#fontsListener) document.fonts.removeEventListener('loadingdone', this.#fontsListener);
    this.#workbook?.close();
    this.#workbook = null;
    this.#root.remove();
  }

  // ---------- Sheet model ----------

  #measureMdw() {
    const font = this.#styles.defaultFont;
    this.#ctx.font = canvasFont({ name: font.name, size: font.size });
    let widest = 0;
    for (const d of '0123456789') widest = Math.max(widest, this.#ctx.measureText(d).width);
    return Math.max(1, Math.round(widest));
  }

  #loadSheet(index, { keepScroll = false } = {}) {
    if (index < 0) {
      this.#sheet = null;
      this.#sheetIndex = -1;
      this.render();
      return;
    }
    const sheet = this.#workbook.sheet(index);
    const layout = sheet.layout();
    const cells = new Map();
    let maxRow = 1;
    let maxCol = 1;
    for (const [row, col, style, text, kind, color] of sheet.cellViews(this.locale)) {
      cells.set(key(row, col), { style, text, kind, color });
      if (text) {
        maxRow = Math.max(maxRow, row);
        maxCol = Math.max(maxCol, col);
      } else {
        maxRow = Math.max(maxRow, Math.min(row, maxRow + 1000));
        maxCol = Math.max(maxCol, Math.min(col, maxCol + 100));
      }
    }
    const merges = layout.merges.map(parseRange);
    const mergeAt = new Map();
    const covered = new Set();
    for (const m of merges) {
      mergeAt.set(key(m.first.row, m.first.col), m);
      for (let r = m.first.row; r <= m.last.row; r++) {
        for (let c = m.first.col; c <= m.last.col; c++) if (r !== m.first.row || c !== m.first.col) covered.add(key(r, c));
      }
    }
    this.#mdw = this.#measureMdw();
    const { rows, cols } = sheetAxes(layout, { maxRow, maxCol }, this.#mdw);
    const frozen = layout.frozen ?? { rows: 0, cols: 0 };
    this.#sheet = {
      cells, merges, mergeAt, covered, rows, cols,
      gridlines: layout.gridlines,
      headings: layout.headings,
      frozen: { rows: Math.min(frozen.rows, rows.count - 1), cols: Math.min(frozen.cols, cols.count - 1) },
    };
    const changed = this.#sheetIndex !== index;
    this.#sheetIndex = index;
    this.#updateSpacer();
    if (changed && !keepScroll) {
      this.#scroller.scrollTo(0, 0);
      const start = { row: this.#sheet.frozen.rows + 1, col: this.#sheet.frozen.cols + 1 };
      this.#selection = { first: { row: 1, col: 1 }, last: { row: 1, col: 1 }, active: { row: 1, col: 1 } };
      if (start.row > 1 || start.col > 1) this.#selection = { first: start, last: start, active: start };
    }
    this.#updateTabs();
    this.#updateBar();
    this.render();
  }

  #rowHeaderW() {
    if (!this.#sheet?.headings) return 0;
    const digits = String(this.#sheet.rows.count).length;
    return Math.round(Math.max(34, 12 + digits * 7.5) * this.zoom);
  }

  #colHeaderH() {
    return this.#sheet?.headings ? Math.round(COL_HEADER_H * this.zoom) : 0;
  }

  #updateSpacer() {
    if (!this.#sheet) {
      this.#spacer.style.width = '0px';
      this.#spacer.style.height = '0px';
      return;
    }
    const { rows, cols } = this.#sheet;
    this.#spacer.style.width = `${this.#rowHeaderW() + cols.total * this.zoom}px`;
    this.#spacer.style.height = `${this.#colHeaderH() + rows.total * this.zoom}px`;
  }

  // ---------- Drawing ----------

  /** Draws on the next animation frame (many calls in a frame draw once) */
  render() {
    if (this.#frame) return;
    this.#frame = requestAnimationFrame(() => {
      this.#frame = 0;
      this.#draw();
    });
  }

  /** Visible panes: rows and columns of each and where they start on the canvas */
  #panes(width, height) {
    const { rows, cols, frozen } = this.#sheet;
    const zoom = this.zoom;
    const gx = this.#rowHeaderW();
    const gy = this.#colHeaderH();
    const frozenW = (cols.offset(frozen.cols + 1) - cols.offset(1)) * zoom;
    const frozenH = (rows.offset(frozen.rows + 1) - rows.offset(1)) * zoom;
    const contentX = cols.offset(frozen.cols + 1) + this.#scroller.scrollLeft / zoom;
    const contentY = rows.offset(frozen.rows + 1) + this.#scroller.scrollTop / zoom;
    const c0 = cols.indexAt(contentX);
    const r0 = rows.indexAt(contentY);
    const mainX = gx + frozenW + (cols.offset(c0) - contentX) * zoom;
    const mainY = gy + frozenH + (rows.offset(r0) - contentY) * zoom;
    const c1 = Math.min(cols.count, cols.indexAt(contentX + (width - gx - frozenW) / zoom));
    const r1 = Math.min(rows.count, rows.indexAt(contentY + (height - gy - frozenH) / zoom));
    const colSegs = [];
    const rowSegs = [];
    if (frozen.cols > 0) colSegs.push({ from: 1, to: frozen.cols, start: gx, clipStart: gx, clipEnd: gx + frozenW });
    colSegs.push({ from: c0, to: c1, start: mainX, clipStart: gx + frozenW, clipEnd: width });
    if (frozen.rows > 0) rowSegs.push({ from: 1, to: frozen.rows, start: gy, clipStart: gy, clipEnd: gy + frozenH });
    rowSegs.push({ from: r0, to: r1, start: mainY, clipStart: gy + frozenH, clipEnd: height });
    return { colSegs, rowSegs, gx, gy, frozenW, frozenH };
  }

  #draw() {
    const rect = this.#body.getBoundingClientRect();
    // The scroller's scrollbars cover part of the body: draw only where cells are visible
    const width = this.#scroller.clientWidth || rect.width;
    const height = this.#scroller.clientHeight || rect.height;
    const dpr = window.devicePixelRatio || 1;
    const canvas = this.#canvas;
    if (canvas.width !== Math.round(width * dpr) || canvas.height !== Math.round(height * dpr)) {
      canvas.width = Math.round(width * dpr);
      canvas.height = Math.round(height * dpr);
      canvas.style.width = `${width}px`;
      canvas.style.height = `${height}px`;
    }
    const ctx = this.#ctx;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);
    if (!this.#sheet) return;

    const { rows, cols } = this.#sheet;
    const zoom = this.zoom;
    const { colSegs, rowSegs, gx, gy, frozenW, frozenH } = this.#panes(width, height);
    for (const rs of rowSegs) {
      for (const cs of colSegs) {
        const clip = { x: cs.clipStart, y: rs.clipStart, w: cs.clipEnd - cs.clipStart, h: rs.clipEnd - rs.clipStart };
        if (clip.w <= 0 || clip.h <= 0) continue;
        drawRegion(ctx, this.#sheet, this.#styles, { r0: rs.from, r1: rs.to, c0: cs.from, c1: cs.to, x0: cs.start, y0: rs.start, clip, zoom });
        this.#drawSelection(ctx, rs, cs, clip);
      }
    }
    // Lines that mark frozen panes
    ctx.fillStyle = '#9ca3af';
    if (frozenW > 0) ctx.fillRect(gx + frozenW - 1, gy, 1, height - gy);
    if (frozenH > 0) ctx.fillRect(gx, gy + frozenH - 1, width - gx, 1);
    if (this.#sheet.headings) {
      drawHeaders(ctx, { rows, cols, colSegs, rowSegs, rowHeaderW: gx, colHeaderH: gy, width, height, zoom, selection: this.#selection, columnName });
    }
  }

  #drawSelection(ctx, rs, cs, clip) {
    const sel = this.#selection;
    if (!sel) return;
    const { rows, cols } = this.#sheet;
    const zoom = this.zoom;
    const xOf = c => cs.start + (cols.offset(c) - cols.offset(cs.from)) * zoom;
    const yOf = r => rs.start + (rows.offset(r) - rows.offset(rs.from)) * zoom;
    const range = this.#expandToMerges(sel.first, sel.last);
    const x = xOf(range.first.col);
    const y = yOf(range.first.row);
    const w = xOf(range.last.col + 1) - x;
    const h = yOf(range.last.row + 1) - y;
    ctx.save();
    ctx.beginPath();
    ctx.rect(clip.x, clip.y, clip.w, clip.h);
    ctx.clip();
    if (range.first.row !== range.last.row || range.first.col !== range.last.col) {
      // Tint the range, leaving the active cell clear as Excel does
      const a = this.#expandToMerges(sel.active, sel.active);
      const ax = xOf(a.first.col);
      const ay = yOf(a.first.row);
      ctx.beginPath();
      ctx.rect(x, y, w, h);
      ctx.rect(ax, ay, xOf(a.last.col + 1) - ax, yOf(a.last.row + 1) - ay);
      ctx.fillStyle = 'rgba(33, 115, 70, 0.12)';
      ctx.fill('evenodd');
    }
    ctx.strokeStyle = SELECTION;
    ctx.lineWidth = Math.max(2, Math.round(2 * zoom));
    ctx.strokeRect(Math.round(x) - 1, Math.round(y) - 1, Math.round(w), Math.round(h));
    ctx.fillStyle = SELECTION;
    const handle = Math.max(5, Math.round(5 * zoom));
    ctx.fillRect(Math.round(x + w) - handle / 2 - 1, Math.round(y + h) - handle / 2 - 1, handle, handle);
    ctx.restore();
  }

  /** Grows a range until it holds every merged range it touches */
  #expandToMerges(first, last) {
    let f = { ...first };
    let l = { ...last };
    let grown = true;
    while (grown) {
      grown = false;
      for (const m of this.#sheet.merges) {
        if (m.last.row < f.row || m.first.row > l.row || m.last.col < f.col || m.first.col > l.col) continue;
        if (m.first.row < f.row || m.first.col < f.col || m.last.row > l.row || m.last.col > l.col) {
          f = { row: Math.min(f.row, m.first.row), col: Math.min(f.col, m.first.col) };
          l = { row: Math.max(l.row, m.last.row), col: Math.max(l.col, m.last.col) };
          grown = true;
        }
      }
    }
    return { first: f, last: l };
  }

  // ---------- Interaction ----------

  /** The cell under a point of the canvas, or null over the headers */
  #cellAt(x, y) {
    const { rows, cols, frozen } = this.#sheet;
    const zoom = this.zoom;
    const { gx, gy, frozenW, frozenH } = this.#panes(this.#scroller.clientWidth, this.#scroller.clientHeight);
    if (x < gx || y < gy) return null;
    const col = x < gx + frozenW
      ? cols.indexAt((x - gx) / zoom + cols.offset(1))
      : cols.indexAt(cols.offset(frozen.cols + 1) + this.#scroller.scrollLeft / zoom + (x - gx - frozenW) / zoom);
    const row = y < gy + frozenH
      ? rows.indexAt((y - gy) / zoom + rows.offset(1))
      : rows.indexAt(rows.offset(frozen.rows + 1) + this.#scroller.scrollTop / zoom + (y - gy - frozenH) / zoom);
    return { row, col };
  }

  #onPointerDown(e) {
    if (!this.#sheet || e.button !== 0) return;
    const box = this.#scroller.getBoundingClientRect();
    const x = e.clientX - box.left;
    const y = e.clientY - box.top;
    // Clicks on the scrollbars scroll
    if (x > this.#scroller.clientWidth || y > this.#scroller.clientHeight) return;
    const at = this.#cellAt(x, y);
    if (!at) return;
    this.#scroller.focus({ preventScroll: true });
    const anchor = e.shiftKey && this.#selection ? this.#selection.active : at;
    this.#setSelection(anchor, at, anchor, { scroll: false });
    const move = ev => {
      const cell = this.#cellAt(ev.clientX - box.left, ev.clientY - box.top);
      if (cell) this.#setSelection(anchor, cell, anchor, { scroll: false });
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  }

  #onKeyDown(e) {
    if (!this.#sheet || !this.#selection) return;
    const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Enter: [1, 0], Tab: [0, 1] };
    const delta = moves[e.key];
    if (!delta) return;
    e.preventDefault();
    const { rows, cols } = this.#sheet;
    let [dr, dc] = delta;
    if (e.key === 'Enter' && e.shiftKey) dr = -1;
    if (e.key === 'Tab' && e.shiftKey) dc = -1;
    const extend = e.shiftKey && e.key.startsWith('Arrow');
    const from = extend ? this.#selection.cursor : this.#selection.active;
    // Step over merged ranges and hidden rows or columns
    const range = this.#expandToMerges(from, from);
    let row = dr > 0 ? range.last.row + 1 : dr < 0 ? range.first.row - 1 : from.row;
    let col = dc > 0 ? range.last.col + 1 : dc < 0 ? range.first.col - 1 : from.col;
    while (dr && row >= 1 && row <= rows.count && rows.size(row) === 0) row += dr;
    while (dc && col >= 1 && col <= cols.count && cols.size(col) === 0) col += dc;
    row = Math.min(Math.max(row, 1), rows.count);
    col = Math.min(Math.max(col, 1), cols.count);
    if (extend) this.#setSelection(this.#selection.active, { row, col }, this.#selection.active);
    else this.#setSelection({ row, col }, { row, col }, { row, col });
  }

  #setSelection(a, b, active, { scroll = true } = {}) {
    const first = { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) };
    const last = { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) };
    const prev = this.#selection;
    // `cursor` is the corner that moves when the range is extended with Shift
    this.#selection = { first, last, active: { ...active }, cursor: { ...b } };
    if (scroll) this.#scrollIntoView(b);
    this.#updateBar();
    this.render();
    if (!prev || prev.first.row !== first.row || prev.first.col !== first.col || prev.last.row !== last.row || prev.last.col !== last.col) {
      this.dispatchEvent(new CustomEvent('selectionchange', { detail: this.selection }));
    }
  }

  #scrollIntoView(cell) {
    const { rows, cols, frozen } = this.#sheet;
    const zoom = this.zoom;
    const { gx, gy, frozenW, frozenH } = this.#panes(this.#scroller.clientWidth, this.#scroller.clientHeight);
    const viewW = this.#scroller.clientWidth - gx - frozenW;
    const viewH = this.#scroller.clientHeight - gy - frozenH;
    if (cell.col > frozen.cols) {
      const left = (cols.offset(cell.col) - cols.offset(frozen.cols + 1)) * zoom;
      const right = left + cols.size(cell.col) * zoom;
      if (left < this.#scroller.scrollLeft) this.#scroller.scrollLeft = left;
      else if (right > this.#scroller.scrollLeft + viewW) this.#scroller.scrollLeft = right - viewW;
    }
    if (cell.row > frozen.rows) {
      const top = (rows.offset(cell.row) - rows.offset(frozen.rows + 1)) * zoom;
      const bottom = top + rows.size(cell.row) * zoom;
      if (top < this.#scroller.scrollTop) this.#scroller.scrollTop = top;
      else if (bottom > this.#scroller.scrollTop + viewH) this.#scroller.scrollTop = bottom - viewH;
    }
  }

  #updateBar() {
    const sel = this.selection;
    this.#nameBox.textContent = sel ? sel.range : '';
    let content = '';
    if (sel && this.#workbook && this.#sheetIndex >= 0) {
      const cell = this.#workbook.sheet(this.#sheetIndex).getCell(sel.ref);
      content = cell.formula ? `=${cell.formula}` : sel.text;
    }
    this.#valueBox.textContent = content;
    this.#valueBox.title = content;
  }

  // ---------- Tabs ----------

  #renderTabs() {
    this.#tabsEl.textContent = '';
    this.#workbook.tabs.forEach((tab, index) => {
      if (tab.state !== 'visible') return;
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'xlsx-editor-tab';
      button.textContent = tab.name;
      button.setAttribute('role', 'tab');
      button.dataset.index = String(index);
      if (tab.kind !== 'worksheet') {
        button.disabled = true;
        button.title = 'Hoja de gráfico (todavía no se muestra)';
      }
      button.addEventListener('click', () => this.setSheet(index));
      this.#tabsEl.appendChild(button);
    });
  }

  #updateTabs() {
    for (const button of this.#tabsEl.querySelectorAll('button')) {
      const active = Number(button.dataset.index) === this.#sheetIndex;
      button.classList.toggle('is-active', active);
      button.setAttribute('aria-selected', String(active));
    }
  }
}
