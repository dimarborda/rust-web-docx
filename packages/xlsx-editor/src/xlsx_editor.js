// XlsxEditor: shows an Excel workbook on a canvas, as Excel draws it, and edits its cells. The
// workbook is read, formatted and written by the Rust/WebAssembly engine of @dimarborda/xlsx;
// this class lays out the grid, scrolls it virtually (only the visible cells are drawn), keeps
// frozen panes in place and handles selection, the in-cell editor, the clipboard, undo and tabs.
import { createWorkbook, initEngine, openWorkbook } from '@dimarborda/xlsx';
import { canvasFont, drawHeaders, drawRegion, key } from './draw.js';
import { columnName, parseRange, sheetAxes } from './grid.js';
import { editText, parseInput, parseTSV, toTSV } from './input.js';

const COL_HEADER_H = 22;
const SELECTION = '#217346';
const MS_PER_DAY = 86_400_000;

const refOf = cell => `${columnName(cell.col)}${cell.row}`;
const rangeOf = (first, last) => (first.row === last.row && first.col === last.col ? refOf(first) : `${refOf(first)}:${refOf(last)}`);

/** A value read from a cell, ready to be written again (times without date are day fractions) */
function rewritable(value) {
  if (value instanceof Date && value.getUTCFullYear() < 1900) return (((value.getTime() % MS_PER_DAY) + MS_PER_DAY) % MS_PER_DAY) / MS_PER_DAY;
  return value;
}

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
   * @param {{locale?: string, zoom?: number, tabs?: boolean, formulaBar?: boolean, readOnly?: boolean, wasmUrl?: string | URL}} [options]
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
  #input;
  #sink;
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
  #clipboardListeners = [];
  #readOnly = false;
  #edit = null; // { row, col, mode: 'enter' | 'edit', original } while a cell is being edited
  #clip = null; // { tsv, values, sheet, first, last, cut? } of the last copy, to paste losslessly
  #lastChange = null; // range of the last change, until the selection after it is recorded

  constructor(container, { locale, zoom = 1, tabs = true, formulaBar = true, readOnly = false } = {}) {
    super();
    this.#container = container;
    this.locale = locale || defaultLocale();
    this.zoom = zoom;

    const root = document.createElement('div');
    root.className = 'xlsx-editor';
    root.innerHTML = `
      <div class="xlsx-editor-bar"${formulaBar ? '' : ' hidden'}>
        <div class="xlsx-editor-name" aria-label="Celda"></div>
        <input class="xlsx-editor-value" aria-label="Contenido" spellcheck="false" autocomplete="off" />
      </div>
      <div class="xlsx-editor-body">
        <canvas class="xlsx-editor-canvas"></canvas>
        <div class="xlsx-editor-scroller" tabindex="0" role="grid" aria-label="Hoja de cálculo"><div class="xlsx-editor-spacer"></div></div>
        <textarea class="xlsx-editor-input" hidden spellcheck="false" autocomplete="off" aria-label="Editar celda"></textarea>
        <textarea class="xlsx-editor-clipboard" tabindex="-1" aria-hidden="true"></textarea>
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
    this.#input = root.querySelector('.xlsx-editor-input');
    this.#sink = root.querySelector('.xlsx-editor-clipboard');

    this.#scroller.addEventListener('scroll', () => this.render());
    this.#scroller.addEventListener('pointerdown', e => this.#onPointerDown(e));
    this.#scroller.addEventListener('keydown', e => this.#onKeyDown(e));
    this.#scroller.addEventListener('dblclick', () => this.#beginEdit({ mode: 'edit' }));
    this.#input.addEventListener('keydown', e => this.#onEditKey(e));
    this.#input.addEventListener('input', () => {
      this.#valueBox.value = this.#input.value;
      this.#placeInput();
    });
    this.#input.addEventListener('blur', e => {
      if (e.relatedTarget !== this.#valueBox) this.#commitEdit(null, { focus: false });
    });
    this.#valueBox.addEventListener('focus', () => this.#beginEdit({ mode: 'edit', fromBar: true }));
    this.#valueBox.addEventListener('input', () => {
      if (!this.#edit) return;
      this.#input.value = this.#valueBox.value;
      this.#placeInput();
    });
    this.#valueBox.addEventListener('keydown', e => this.#onEditKey(e, { bar: true }));
    this.#valueBox.addEventListener('blur', e => {
      if (e.relatedTarget !== this.#input) this.#commitEdit(null, { focus: false });
    });
    // The native `change`, `input` and `selectionchange` of the inner text fields are not the
    // editor's events: keep them from reaching listeners on <xlsx-editor>
    for (const field of [this.#input, this.#valueBox, this.#sink]) {
      for (const type of ['change', 'input', 'selectionchange']) field.addEventListener(type, e => e.stopPropagation());
    }
    for (const type of ['copy', 'cut', 'paste']) {
      // Ctrl/Cmd+C, X and V focus the sink first (see #onKeyDown); the Edit menu of the browser
      // fires them on <body> while the grid has the focus
      this.#sink.addEventListener(type, e => {
        e.stopPropagation();
        this.#onClipboard(e);
        this.#scroller.focus({ preventScroll: true });
      });
      const listener = e => {
        if (document.activeElement === this.#scroller) this.#onClipboard(e);
      };
      document.addEventListener(type, listener);
      this.#clipboardListeners.push([type, listener]);
    }
    this.readOnly = readOnly;
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

  /** Without editing from the grid: a viewer (cells can still be selected and copied) */
  get readOnly() {
    return this.#readOnly;
  }

  set readOnly(value) {
    this.#readOnly = !!value;
    if (this.#readOnly) this.#cancelEdit();
    this.#valueBox.readOnly = this.#readOnly;
    this.#root.classList.toggle('is-readonly', this.#readOnly);
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
    this.#cancelEdit();
    this.#clip = null;
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
    this.#commitEdit(null, { focus: false });
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
    this.#commitEdit(null, { focus: false });
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

  /** Undoes the last change made from the grid (or since the last `workbook.checkpoint()`) */
  undo() {
    return this.#moveInHistory('undo');
  }

  /** Redoes the last undone change */
  redo() {
    return this.#moveInHistory('redo');
  }

  get canUndo() {
    return !!this.#workbook?.canUndo;
  }

  get canRedo() {
    return !!this.#workbook?.canRedo;
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
    for (const [type, listener] of this.#clipboardListeners) document.removeEventListener(type, listener);
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
    if (changed) this.#cancelEdit();
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
        this.#drawCopied(ctx, rs, cs, clip);
      }
    }
    // Lines that mark frozen panes
    ctx.fillStyle = '#9ca3af';
    if (frozenW > 0) ctx.fillRect(gx + frozenW - 1, gy, 1, height - gy);
    if (frozenH > 0) ctx.fillRect(gx, gy + frozenH - 1, width - gx, 1);
    if (this.#edit) this.#placeInput();
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

  /** Dashed border around the range last copied, as Excel's marching ants */
  #drawCopied(ctx, rs, cs, clip) {
    const c = this.#clip;
    if (!c || c.sheet !== this.#sheetIndex) return;
    const { rows, cols } = this.#sheet;
    const zoom = this.zoom;
    const x = cs.start + (cols.offset(c.first.col) - cols.offset(cs.from)) * zoom;
    const y = rs.start + (rows.offset(c.first.row) - rows.offset(rs.from)) * zoom;
    const w = (cols.offset(c.last.col + 1) - cols.offset(c.first.col)) * zoom;
    const h = (rows.offset(c.last.row + 1) - rows.offset(c.first.row)) * zoom;
    ctx.save();
    ctx.beginPath();
    ctx.rect(clip.x, clip.y, clip.w, clip.h);
    ctx.clip();
    ctx.lineWidth = 1;
    ctx.setLineDash([4, 3]);
    ctx.strokeStyle = SELECTION;
    ctx.strokeRect(Math.round(x) + 0.5, Math.round(y) + 0.5, Math.round(w) - 2, Math.round(h) - 2);
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
    this.#commitEdit(null, { focus: false });
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
    if (!this.#sheet || !this.#selection || e.isComposing) return;
    if ((e.ctrlKey || e.metaKey) && !e.altKey) {
      const k = e.key.toLowerCase();
      if (k === 'z' || k === 'y') {
        e.preventDefault();
        if (k === 'y' || e.shiftKey) this.redo();
        else this.undo();
      } else if (k === 'c' || k === 'x' || k === 'v') {
        // Browsers fire clipboard events only where text can be selected or edited (Safari
        // needs a selection): let the shortcut act on a hidden textarea, then come back
        this.#sink.value = ' ';
        this.#sink.focus({ preventScroll: true });
        this.#sink.select();
        setTimeout(() => {
          if (document.activeElement === this.#sink) this.#scroller.focus({ preventScroll: true });
        });
      }
      return;
    }
    if (e.key === 'Escape' && this.#clip) {
      this.#clip = null;
      this.render();
      return;
    }
    if (!this.#readOnly) {
      if (e.key === 'F2') {
        e.preventDefault();
        this.#beginEdit({ mode: 'edit' });
        return;
      }
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        this.#clearSelection();
        return;
      }
      if (e.key.length === 1 || e.key === 'Dead' || e.key === 'Process') {
        // Typing replaces the cell; dead keys and IMEs compose in the editor
        const plain = e.key.length === 1;
        if (plain) e.preventDefault();
        this.#beginEdit({ mode: 'enter', text: plain ? e.key : '' });
        return;
      }
    }
    const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Enter: [1, 0], Tab: [0, 1] };
    const delta = moves[e.key];
    if (!delta) return;
    e.preventDefault();
    let [dr, dc] = delta;
    if (e.key === 'Enter' && e.shiftKey) dr = -1;
    if (e.key === 'Tab' && e.shiftKey) dc = -1;
    this.#move(dr, dc, e.shiftKey && e.key.startsWith('Arrow'));
  }

  /** Moves the active cell (or, with `extend`, the moving corner of the range) one step */
  #move(dr, dc, extend = false) {
    const { rows, cols } = this.#sheet;
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
    if (this.#edit) return; // the bar shows what is being typed
    let content = '';
    if (sel && this.#workbook && this.#sheetIndex >= 0) {
      const at = this.#expandToMerges(sel.active, sel.active).first;
      const cell = this.#activeSheet().getCell(refOf(at));
      content = editText(cell, this.locale, this.#sheet?.cells.get(key(at.row, at.col))?.text ?? '');
    }
    this.#valueBox.value = content;
    this.#valueBox.title = content;
  }

  // ---------- Editing ----------

  #activeSheet() {
    return this.#workbook.sheet(this.#sheetIndex);
  }

  /**
   * Opens the editor on the active cell: `enter` mode (started by typing) moves with the arrow
   * keys as Excel does; `edit` mode (F2, double click, the bar) moves the caret
   */
  #beginEdit({ mode, text, fromBar = false }) {
    if (this.#readOnly || this.#edit || !this.#sheet || !this.#selection) return;
    const at = this.#expandToMerges(this.#selection.active, this.#selection.active).first;
    const cell = this.#activeSheet().getCell(refOf(at));
    const shown = this.#sheet.cells.get(key(at.row, at.col))?.text ?? '';
    const original = editText(cell, this.locale, shown);
    this.#edit = { row: at.row, col: at.col, mode, original };
    this.#clip = null;
    this.#scrollIntoView(at);
    const value = text ?? original;
    this.#input.value = value;
    this.#valueBox.value = value;
    this.#styleInput();
    this.#input.hidden = false;
    this.#placeInput();
    if (!fromBar) {
      this.#input.focus({ preventScroll: true });
      this.#input.setSelectionRange(value.length, value.length);
    }
    this.render();
  }

  /** Font and colors of the cell being edited */
  #styleInput() {
    const view = this.#sheet.cells.get(key(this.#edit.row, this.#edit.col));
    const xf = this.#styles.xfs[view?.style ?? 0] || this.#styles.xfs[0];
    const font = xf?.font ?? this.#styles.defaultFont;
    const style = this.#input.style;
    style.font = canvasFont(font, this.zoom);
    style.lineHeight = '1.2';
    style.color = font.color || '#000000';
    style.background = xf?.fill || '#ffffff';
    style.textDecoration = [font.underline && 'underline', font.strike && 'line-through'].filter(Boolean).join(' ') || 'none';
    style.whiteSpace = xf?.wrap ? 'pre-wrap' : 'pre';
    style.textAlign = xf?.h === 'center' || xf?.h === 'right' ? xf.h : 'left';
  }

  /** Puts the editor over its cell, wide and tall enough for the text */
  #placeInput() {
    if (!this.#edit || !this.#sheet) return;
    const { rows, cols, frozen } = this.#sheet;
    const zoom = this.zoom;
    const width = this.#scroller.clientWidth;
    const height = this.#scroller.clientHeight;
    const { gx, gy, frozenW, frozenH } = this.#panes(width, height);
    const xOf = c => (c <= frozen.cols ? gx + (cols.offset(c) - cols.offset(1)) * zoom : gx + frozenW + (cols.offset(c) - cols.offset(frozen.cols + 1)) * zoom - this.#scroller.scrollLeft);
    const yOf = r => (r <= frozen.rows ? gy + (rows.offset(r) - rows.offset(1)) * zoom : gy + frozenH + (rows.offset(r) - rows.offset(frozen.rows + 1)) * zoom - this.#scroller.scrollTop);
    const at = { row: this.#edit.row, col: this.#edit.col };
    const range = this.#expandToMerges(at, at);
    const x = Math.round(xOf(range.first.col)) - 1;
    const y = Math.round(yOf(range.first.row)) - 1;
    const cellW = Math.round(xOf(range.last.col + 1)) - x + 1;
    const cellH = Math.round(yOf(range.last.row + 1)) - y + 1;
    const input = this.#input;
    this.#ctx.font = input.style.font;
    const lines = input.value.split('\n');
    const textW = Math.max(...lines.map(l => this.#ctx.measureText(l).width));
    const lineH = (parseFloat(input.style.fontSize) || 15) * 1.2;
    const maxW = Math.max(cellW, width - x - 2);
    const w = input.style.whiteSpace === 'pre-wrap' ? cellW : Math.min(maxW, Math.max(cellW, Math.ceil(textW) + 14 * zoom));
    input.style.left = `${x}px`;
    input.style.top = `${y}px`;
    input.style.width = `${w}px`;
    input.style.height = `${cellH}px`;
    // Tall enough for every line (scrollHeight counts the wrapped ones)
    const needed = Math.max(cellH, Math.ceil(lines.length * lineH + 6), input.scrollHeight);
    input.style.height = `${Math.min(Math.max(cellH, height - y - 2), needed)}px`;
  }

  #closeEditor() {
    this.#edit = null;
    this.#input.hidden = true;
    this.#updateBar();
  }

  /** Closes the editor without writing */
  #cancelEdit({ focus = false } = {}) {
    if (!this.#edit) return;
    this.#closeEditor();
    if (focus) this.#scroller.focus({ preventScroll: true });
    this.render();
  }

  /**
   * Writes what was typed and closes the editor; `move` is the step to take afterwards and
   * `fill` writes it in every cell of the selection (Ctrl+Enter)
   */
  #commitEdit(move, { focus = true, fill = false } = {}) {
    const edit = this.#edit;
    if (!edit) return;
    const text = this.#input.value;
    this.#closeEditor();
    if (text !== edit.original || fill) {
      const { value } = parseInput(text, this.locale);
      const sel = this.#selection;
      const first = fill ? sel.first : { row: edit.row, col: edit.col };
      const last = fill ? sel.last : first;
      const rows = Array.from({ length: last.row - first.row + 1 }, () => new Array(last.col - first.col + 1).fill(value));
      this.#write(first, rows, 'edit');
    }
    if (focus) this.#scroller.focus({ preventScroll: true });
    if (move) this.#move(...move);
    this.#recordSelectionAfter();
    this.render();
  }

  #onEditKey(e, { bar = false } = {}) {
    if (!this.#edit || e.isComposing) return;
    const mod = e.ctrlKey || e.metaKey;
    switch (e.key) {
      case 'Enter':
        e.preventDefault();
        if (e.altKey) {
          // Alt+Enter breaks the line inside the cell
          const field = bar ? this.#valueBox : this.#input;
          if (bar) return;
          field.setRangeText('\n', field.selectionStart, field.selectionEnd, 'end');
          field.dispatchEvent(new Event('input'));
          return;
        }
        if (mod) this.#commitEdit(null, { fill: true });
        else this.#commitEdit([e.shiftKey ? -1 : 1, 0]);
        return;
      case 'Tab':
        e.preventDefault();
        this.#commitEdit([0, e.shiftKey ? -1 : 1]);
        return;
      case 'Escape':
        e.preventDefault();
        this.#cancelEdit({ focus: true });
        return;
      case 'F2':
        e.preventDefault();
        this.#edit.mode = this.#edit.mode === 'enter' ? 'edit' : 'enter';
        return;
      default:
        if (!bar && this.#edit.mode === 'enter' && !mod && e.key.startsWith('Arrow')) {
          e.preventDefault();
          const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
          this.#commitEdit(moves[e.key]);
        }
    }
  }

  /** What the undo history remembers about a change: the sheet, the selection, the cells */
  #historyToken(changed) {
    const sel = this.selection;
    return JSON.stringify({ sheet: this.#sheetIndex, range: sel?.range, active: sel?.ref, changed });
  }

  /** Once the selection has moved after a change, redo puts it back there */
  #recordSelectionAfter() {
    if (this.#lastChange) this.#workbook.setSelectionAfter(this.#historyToken(this.#lastChange));
    this.#lastChange = null;
  }

  /**
   * Writes rows of values from `first` as one undo step and draws the sheet again; returns the
   * range written, or null when the engine refused (the workbook is then unchanged)
   */
  #write(first, rows, source) {
    const width = Math.max(...rows.map(r => r.length));
    const last = { row: first.row + rows.length - 1, col: first.col + width - 1 };
    const changed = rangeOf(first, last);
    try {
      this.#workbook.checkpoint(this.#historyToken(changed));
      this.#activeSheet().setValues(refOf(first), rows);
    } catch (error) {
      this.dispatchEvent(new CustomEvent('error', { detail: { error } }));
      return null;
    }
    this.#lastChange = changed;
    this.#workbook.setSelectionAfter(this.#historyToken(changed));
    this.reload();
    this.#emitChange(changed, source);
    return { first, last };
  }

  #emitChange(range, source) {
    const sheet = this.#activeSheet();
    const { first } = parseRange(range);
    const cells = [];
    sheet.getValues(range).forEach((row, r) => row.forEach((value, c) => {
      cells.push({ ref: refOf({ row: first.row + r, col: first.col + c }), value });
    }));
    this.dispatchEvent(new CustomEvent('change', { detail: { sheet: sheet.name, index: this.#sheetIndex, range, cells, source } }));
  }

  /** Delete and Backspace: empties the selected cells, keeping their format */
  #clearSelection(source = 'clear') {
    const { first, last } = this.#expandToMerges(this.#selection.first, this.#selection.last);
    const rows = Array.from({ length: last.row - first.row + 1 }, () => new Array(last.col - first.col + 1).fill(null));
    this.#clip = null;
    this.#write(first, rows, source);
    this.#recordSelectionAfter();
  }

  #moveInHistory(direction) {
    if (this.#readOnly || !this.#workbook) return false;
    this.#cancelEdit();
    const done = direction === 'undo' ? this.#workbook.undo() : this.#workbook.redo();
    if (!done) return false;
    let token = null;
    try {
      token = done.selection ? JSON.parse(done.selection) : null;
    } catch {
      token = null; // a checkpoint made from code with a selection of its own
    }
    this.#clip = null;
    this.#styles = this.#workbook.styleView();
    const tabs = this.#workbook.tabs;
    if (token && token.sheet !== this.#sheetIndex && tabs[token.sheet]?.kind === 'worksheet') {
      this.#loadSheet(token.sheet);
      this.dispatchEvent(new CustomEvent('sheetchange', { detail: { index: token.sheet, name: tabs[token.sheet].name } }));
    } else {
      this.#loadSheet(this.#sheetIndex, { keepScroll: true });
    }
    if (token?.range) {
      const range = parseRange(token.range);
      const active = token.active ? parseRange(token.active).first : range.first;
      this.#setSelection(range.first, range.last, active);
    }
    if (token?.changed) this.#emitChange(token.changed, direction);
    return true;
  }

  // ---------- Clipboard ----------

  #onClipboard(e) {
    if (!this.#sheet || !this.#selection) return;
    if (e.type === 'paste') {
      if (this.#readOnly) return;
      e.preventDefault();
      this.#paste(e.clipboardData?.getData('text/plain') ?? '');
      return;
    }
    if (e.type === 'cut' && this.#readOnly) return;
    e.preventDefault();
    const { first, last } = this.#expandToMerges(this.#selection.first, this.#selection.last);
    const texts = [];
    for (let r = first.row; r <= last.row; r++) {
      const row = [];
      for (let c = first.col; c <= last.col; c++) row.push(this.#sheet.cells.get(key(r, c))?.text ?? '');
      texts.push(row);
    }
    // Excel and Google Sheets get the texts the cells show; pasting here again keeps the values
    const tsv = toTSV(texts);
    e.clipboardData?.setData('text/plain', tsv);
    const values = this.#activeSheet().getValues(rangeOf(first, last)).map(row => row.map(rewritable));
    this.#clip = { tsv, values, sheet: this.#sheetIndex, first, last };
    if (e.type === 'cut') {
      this.#clearSelection('cut');
      this.#clip = { tsv, values, sheet: this.#sheetIndex, first, last, cut: true };
    }
    this.render();
  }

  /** Writes TSV from the selection; a block that fits the selection a whole number of times repeats */
  #paste(text) {
    const normalize = t => t.replace(/\r\n?/g, '\n').replace(/\n$/, '');
    const own = this.#clip && normalize(this.#clip.tsv) === normalize(text);
    let rows = own ? this.#clip.values : parseTSV(text).map(row => row.map(t => parseInput(t, this.locale).value));
    if (!rows.length || (rows.length === 1 && rows[0].length === 1 && !text)) return;
    const width = Math.max(...rows.map(r => r.length));
    rows = rows.map(r => Array.from({ length: width }, (_, i) => (i < r.length ? r[i] : null)));
    const { first, last } = this.#selection;
    const selH = last.row - first.row + 1;
    const selW = last.col - first.col + 1;
    if ((selH > rows.length || selW > width) && selH % rows.length === 0 && selW % width === 0) {
      rows = Array.from({ length: selH }, (_, r) => Array.from({ length: selW }, (_, c) => rows[r % rows.length][c % width]));
    }
    if (this.#clip?.cut) this.#clip = null;
    const written = this.#write(first, rows, 'paste');
    if (written) {
      this.#setSelection(written.first, written.last, written.first);
      this.#recordSelectionAfter();
    }
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
