// XlsxEditor: shows an Excel workbook on a canvas, as Excel draws it, and edits its cells. The
// workbook is read, formatted and written by the Rust/WebAssembly engine (engine.js);
// this class lays out the grid, scrolls it virtually (only the visible cells are drawn), keeps
// frozen panes in place and handles selection, the in-cell editor, the clipboard, undo, tabs,
// moving rows and columns from their headers and the status bar.
import { createWorkbook, initEngine, openWorkbook } from './engine.js';
import { PAD, canvasFont, drawHeaders, drawRegion, key, wrapLines } from './draw.js';
import { Axis, MAX_COLS, MAX_ROWS, colWidthToPx, columnName, dataEdge, numberRun, parseRange, ptToPx, pxToColWidth, rangeStats, sheetAxes } from './grid.js';
import { matches, replaceMatches, searchPattern } from './find.js';
import { editText, parseInput, parseTSV, toTSV } from './input.js';

const COL_HEADER_H = 22;
const SELECTION = '#217346';
const MS_PER_DAY = 86_400_000;
/** Pixels the pointer travels on a selected header before it counts as dragging it */
const DRAG_START_PX = 4;
/** How close to the edge of the view a drag scrolls it, and how fast (pixels every 40 ms) */
const DRAG_SCROLL_EDGE_PX = 24;
const DRAG_SCROLL_STEP_PX = 24;
/** Above this many cells a `change` event leaves out their values, to stay fast */
const MAX_CHANGE_CELLS = 200_000;
/** How near the edge of a header a press resizes its row or column */
const RESIZE_EDGE_PX = 4;
/** Smallest size a drag leaves a row or column with */
const MIN_LINE_PX = 4;
/** How near the corner of the selection a press takes its fill handle */
const FILL_HANDLE_PX = 6;
/** Pause after typing in the search panel before looking (a search reads the whole sheet) */
const FIND_DELAY_MS = 150;
/** Largest row Excel allows, in points */
const MAX_ROW_PT = 409.5;
/** Largest column Excel allows, in characters */
const MAX_COL_CHARS = 255;
/** Seconds a message of the status bar stays */
const MESSAGE_SECONDS = 8;

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
   * @param {{locale?: string, zoom?: number, tabs?: boolean, formulaBar?: boolean, statusBar?: boolean, readOnly?: boolean, wasmUrl?: string | URL}} [options]
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
  #statusEl;
  #messageEl;
  #statsEl;
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
  #drag = null; // { axis, first, count, before, cancel } while a header is dragged to move lines
  #statusDirty = false;
  #statusShown = null; // what the statistics of the status bar were computed from
  #messageTimer = 0;
  #gesture = null; // { cancel } while the edge of a header or the fill handle is dragged
  #fillDrag = null; // { first, last } of the range the fill handle would fill
  #find = null; // the search panel: its fields and the hits found, once it has been opened
  #menu = null; // the open context menu: { close }

  constructor(container, { locale, zoom = 1, tabs = true, formulaBar = true, statusBar = true, readOnly = false } = {}) {
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
      <div class="xlsx-editor-tabs" role="tablist"${tabs ? '' : ' hidden'}></div>
      <div class="xlsx-editor-status"${statusBar ? '' : ' hidden'}>
        <span class="xlsx-editor-status-message" role="status"></span>
        <span class="xlsx-editor-status-stats"></span>
      </div>`;
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
    this.#statusEl = root.querySelector('.xlsx-editor-status');
    this.#messageEl = root.querySelector('.xlsx-editor-status-message');
    this.#statsEl = root.querySelector('.xlsx-editor-status-stats');
    this.#input = root.querySelector('.xlsx-editor-input');
    this.#sink = root.querySelector('.xlsx-editor-clipboard');

    this.#scroller.addEventListener('scroll', () => {
      this.#growOnScroll();
      this.render();
    });
    this.#scroller.addEventListener('pointerdown', e => this.#onPointerDown(e));
    this.#scroller.addEventListener('pointermove', e => this.#updateCursor(e));
    this.#scroller.addEventListener('keydown', e => this.#onKeyDown(e));
    this.#scroller.addEventListener('dblclick', e => this.#onDoubleClick(e));
    this.#scroller.addEventListener('contextmenu', e => this.#onContextMenu(e));
    this.addEventListener('change', () => this.#staleFind());
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

  /** The open workbook (see `./engine`), or null */
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
    this.#staleFind();
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
    this.#staleFind();
    this.#styles = this.#workbook.styleView();
    // Sheets added or renamed through `workbook` show up in the tabs
    this.#renderTabs();
    this.#loadSheet(this.#validSheet(this.#sheetIndex), { keepScroll: true });
  }

  /** `index` if it is still a worksheet (after undoing an added sheet), else the first one */
  #validSheet(index) {
    const tabs = this.#workbook.tabs;
    return tabs[index]?.kind === 'worksheet' ? index : tabs.findIndex(t => t.kind === 'worksheet');
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
    this.#extend(range.last.row, range.last.col);
    this.#setSelection(range.first, range.last, range.first);
  }

  /**
   * Selects an area and shows its sheet: a table or defined name, an address with its sheet
   * (`Ventas!A1:H40`) or a range of the active sheet. Returns the area (see
   * `workbook.areas()`).
   */
  selectArea(target) {
    if (!this.#workbook) throw new Error('No hay ningún libro abierto.');
    const { rows, columnsInfo, markdown, totals, merges, rowCount, rowsShown, truncated, firstRow, ...area } = this.#workbook.area(target, { sheet: this.#sheetIndex, maxRows: 0 });
    if (area.sheet !== this.#sheetIndex) this.setSheet(area.sheet);
    this.select(area.ref);
    return area;
  }

  /**
   * Context of what is selected, for an agent (see `workbook.area()`): the selected range, or
   * with a single cell selected the region around it (its table, or the block Ctrl+A selects)
   * @param {{maxRows?: number, expand?: boolean}} [options] `expand: false` keeps a single cell
   */
  selectionContext({ maxRows = 30, expand = true } = {}) {
    const selection = this.selection;
    if (!selection || !this.#workbook) return null;
    let target = selection.range;
    if (!target.includes(':') && expand) target = this.#workbook.sheet(this.#sheetIndex).regionAt(selection.ref).ref;
    return this.#workbook.area(target, { sheet: this.#sheetIndex, maxRows, locale: this.locale });
  }

  /** The selected range: `{ref, range, active, text, whole}` or null */
  get selection() {
    if (!this.#selection) return null;
    const { first, last, active, whole = null } = this.#selection;
    const ref = `${columnName(first.col)}${first.row}`;
    const range = first.row === last.row && first.col === last.col ? ref : `${ref}:${columnName(last.col)}${last.row}`;
    const cell = this.#sheet?.cells.get(key(active.row, active.col));
    return { ref: `${columnName(active.col)}${active.row}`, range, active: { ...active }, text: cell?.text ?? '', whole };
  }

  /**
   * Count, sum, average, smallest and largest of the selected cells, as the status bar shows
   * them; null when they are empty
   */
  get selectionStats() {
    if (!this.#selection || !this.#sheet) return null;
    const stats = rangeStats(this.#sheet.cells, this.#selection, this.#sheet);
    if (!stats) return null;
    const { style, ...numbers } = stats;
    return numbers;
  }

  /**
   * Moves `count` whole rows from `first` before the row `before` (numbered as they are now), as
   * Excel's "cut and insert", in one undo step, and selects them. False when nothing changed: it
   * stays where it is, the editor is read-only, or the engine refused (an `error` event says why).
   */
  moveRows(first, count, before) {
    return this.#moveLines('row', first, count, before);
  }

  /** Moves whole columns, by number (see `moveRows`) */
  moveColumns(first, count, before) {
    return this.#moveLines('col', first, count, before);
  }

  /** Inserts `count` rows before the row `before`, as one undo step, and selects them; false if refused */
  insertRows(before, count = 1) {
    return this.#insertLines('row', before, count);
  }

  /** Inserts columns, by number (see `insertRows`) */
  insertColumns(before, count = 1) {
    return this.#insertLines('col', before, count);
  }

  /** Deletes `count` rows from `first`, as one undo step; false if refused */
  deleteRows(first, count = 1) {
    return this.#deleteLines('row', first, count);
  }

  /** Deletes columns, by number (see `deleteRows`) */
  deleteColumns(first, count = 1) {
    return this.#deleteLines('col', first, count);
  }

  /** Hides `count` rows from `first`, as one undo step */
  hideRows(first, count = 1) {
    return this.#setLinesHidden('row', first, first + count - 1, true);
  }

  /** Shows the hidden rows among the `count` from `first` */
  showRows(first, count = 1) {
    return this.#setLinesHidden('row', first, first + count - 1, false);
  }

  /** Hides columns, by number */
  hideColumns(first, count = 1) {
    return this.#setLinesHidden('col', first, first + count - 1, true);
  }

  /** Shows hidden columns, by number */
  showColumns(first, count = 1) {
    return this.#setLinesHidden('col', first, first + count - 1, false);
  }

  /**
   * Fills `target` (`A1:A20`) from the selected cells, which must be at one of its ends, as one
   * undo step, and selects it. False if refused (an `error` event says why).
   * @param {string} target
   * @param {{series?: boolean}} [options]
   */
  fill(target, { series = true } = {}) {
    if (!this.#selection) return false;
    const range = this.#selection;
    return this.#fillRange(rangeOf(range.first, range.last), parseRange(target), series);
  }

  /**
   * Where a text is, row by row: `*` is any text, `?` one character and `~` takes them literally
   * @param {string} query
   * @param {{matchCase?: boolean, wholeCell?: boolean, scope?: 'sheet' | 'workbook', inFormulas?: boolean}} [options]
   */
  find(query, options = {}) {
    return this.#findHits(query, options).map(({ cell, ...hit }) => hit);
  }

  /**
   * Replaces a text in cells with text and in formulas (never in numbers or dates), as one undo
   * step; returns how many cells changed. Formatted runs of a rich text become plain.
   */
  replaceAll(query, replacement, options = {}) {
    if (this.#readOnly || !this.#workbook) return 0;
    const pattern = searchPattern(query, options);
    return pattern ? this.#replaceHits(this.#findHits(query, { ...options, inFormulas: true }), pattern, replacement) : 0;
  }

  /** Opens the search panel, with the row to replace when asked */
  openFind({ replace = false } = {}) {
    const find = this.#findPanel();
    find.el.hidden = false;
    find.replaceRow.hidden = !replace || this.#readOnly;
    find.replacing = !find.replaceRow.hidden;
    find.hits = null;
    find.query.focus({ preventScroll: true });
    find.query.select();
    this.#findStatus('');
  }

  closeFind() {
    if (!this.#find || this.#find.el.hidden) return;
    this.#find.el.hidden = true;
    clearTimeout(this.#find.timer);
    this.#scroller.focus({ preventScroll: true });
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
    this.#drag?.cancel();
    this.#gesture?.cancel();
    this.#menu?.close();
    if (this.#find) clearTimeout(this.#find.timer);
    clearTimeout(this.#messageTimer);
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
    for (const [row, col, style, text, kind, color, number] of sheet.cellViews(this.locale)) {
      cells.set(key(row, col), { style, text, kind, color, number });
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
    let { rows, cols } = sheetAxes(layout, { maxRow, maxCol }, this.#mdw);
    // Reading the same sheet again keeps the room scrolled into past its cells
    if (this.#sheetIndex === index && this.#sheet) {
      rows = rows.grown(this.#sheet.rows.count, MAX_ROWS);
      cols = cols.grown(this.#sheet.cols.count, MAX_COLS);
    }
    const frozen = layout.frozen ?? { rows: 0, cols: 0 };
    this.#sheet = {
      cells, merges, mergeAt, covered, rows, cols,
      used: { maxRow, maxCol },
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
    this.#fitWholeSelection();
    this.#statusDirty = true;
    this.#statusShown = null;
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

  /**
   * Makes room up to row `row` and column `col` and some more, as Excel does when scrolling or
   * moving past the end of the shown grid
   */
  #extend(row, col) {
    const sheet = this.#sheet;
    if (!sheet) return;
    const rows = row + 50 > sheet.rows.count ? sheet.rows.grown(row + 200, MAX_ROWS) : sheet.rows;
    const cols = col + 5 > sheet.cols.count ? sheet.cols.grown(col + 20, MAX_COLS) : sheet.cols;
    if (rows === sheet.rows && cols === sheet.cols) return;
    sheet.rows = rows;
    sheet.cols = cols;
    this.#fitWholeSelection();
    this.#updateSpacer();
  }

  /** Whole rows or columns reach the end of the grid, also after it grows */
  #fitWholeSelection() {
    const sel = this.#selection;
    if (!sel?.whole || !this.#sheet) return;
    if (sel.whole !== 'row') {
      sel.last = { ...sel.last, row: this.#sheet.rows.count };
      sel.cursor = { ...sel.cursor, row: this.#sheet.rows.count };
    }
    if (sel.whole !== 'col') {
      sel.last = { ...sel.last, col: this.#sheet.cols.count };
      sel.cursor = { ...sel.cursor, col: this.#sheet.cols.count };
    }
  }

  /** Grows the grid when the view comes near its bottom or right end */
  #growOnScroll() {
    const s = this.#scroller;
    if (!this.#sheet) return;
    const nearBottom = s.scrollTop + s.clientHeight * 2 >= s.scrollHeight;
    const nearRight = s.scrollLeft + s.clientWidth * 2 >= s.scrollWidth;
    if (nearBottom || nearRight) this.#extend(nearBottom ? this.#sheet.rows.count : 1, nearRight ? this.#sheet.cols.count : 1);
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
      if (this.#statusDirty) this.#updateStatus();
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
    if (this.#drag) this.#drawLineDrag(ctx, width, height, gx, gy);
    if (this.#fillDrag) this.#drawFillDrag(ctx, width, height, gx, gy);
  }

  /** While the fill handle is dragged: the range it would fill */
  #drawFillDrag(ctx, width, height, gx, gy) {
    const { first, last } = this.#fillDrag;
    const x = this.#linePos('col', first.col);
    const y = this.#linePos('row', first.row);
    ctx.save();
    ctx.beginPath();
    ctx.rect(gx, gy, width - gx, height - gy);
    ctx.clip();
    ctx.setLineDash([4, 3]);
    ctx.lineWidth = 2;
    ctx.strokeStyle = SELECTION;
    ctx.strokeRect(Math.round(x), Math.round(y), Math.round(this.#linePos('col', last.col + 1) - x) - 1, Math.round(this.#linePos('row', last.row + 1) - y) - 1);
    ctx.restore();
  }

  /** While a header is dragged: the lines it takes and the place they would go */
  #drawLineDrag(ctx, width, height, gx, gy) {
    const { axis, first, count, before } = this.#drag;
    const isRow = axis === 'row';
    const start = this.#linePos(axis, first);
    const end = this.#linePos(axis, first + count);
    ctx.save();
    ctx.beginPath();
    if (isRow) ctx.rect(0, gy, width, height - gy);
    else ctx.rect(gx, 0, width - gx, height);
    ctx.clip();
    ctx.fillStyle = 'rgba(33, 115, 70, 0.18)';
    if (isRow) ctx.fillRect(0, start, width, end - start);
    else ctx.fillRect(start, 0, end - start, height);
    if (before !== null) {
      const at = Math.round(this.#linePos(axis, before));
      ctx.fillStyle = SELECTION;
      if (isRow) ctx.fillRect(0, at - 1, width, 3);
      else ctx.fillRect(at - 1, 0, 3, height);
    }
    ctx.restore();
  }

  #drawSelection(ctx, rs, cs, clip) {
    const sel = this.#selection;
    if (!sel) return;
    const { rows, cols } = this.#sheet;
    const zoom = this.zoom;
    const xOf = c => cs.start + (cols.offset(c) - cols.offset(cs.from)) * zoom;
    const yOf = r => rs.start + (rows.offset(r) - rows.offset(rs.from)) * zoom;
    // Whole rows or columns are what was asked for, even under a merged title that goes beyond them
    const range = sel.whole ? sel : this.#expandToMerges(sel.first, sel.last);
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
      const a = sel.whole ? { first: sel.active, last: sel.active } : this.#expandToMerges(sel.active, sel.active);
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
    if (!sel.whole) ctx.fillRect(Math.round(x + w) - handle / 2 - 1, Math.round(y + h) - handle / 2 - 1, handle, handle);
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
    const { gx, gy } = this.#panes(this.#scroller.clientWidth, this.#scroller.clientHeight);
    if (x < gx || y < gy) return null;
    return { row: this.#lineAt('row', y), col: this.#lineAt('col', x) };
  }

  /** The row or column at a position of the canvas (clamped to the grid) */
  #lineAt(axis, pos) {
    const { rows, cols, frozen } = this.#sheet;
    const zoom = this.zoom;
    const { gx, gy, frozenW, frozenH } = this.#panes(this.#scroller.clientWidth, this.#scroller.clientHeight);
    if (axis === 'col') {
      return pos < gx + frozenW
        ? cols.indexAt((pos - gx) / zoom + cols.offset(1))
        : cols.indexAt(cols.offset(frozen.cols + 1) + this.#scroller.scrollLeft / zoom + (pos - gx - frozenW) / zoom);
    }
    return pos < gy + frozenH
      ? rows.indexAt((pos - gy) / zoom + rows.offset(1))
      : rows.indexAt(rows.offset(frozen.rows + 1) + this.#scroller.scrollTop / zoom + (pos - gy - frozenH) / zoom);
  }

  /** Where a row or column starts on the canvas */
  #linePos(axis, line) {
    const { rows, cols, frozen } = this.#sheet;
    const zoom = this.zoom;
    const { gx, gy, frozenW, frozenH } = this.#panes(this.#scroller.clientWidth, this.#scroller.clientHeight);
    if (axis === 'col') {
      return line <= frozen.cols
        ? gx + (cols.offset(line) - cols.offset(1)) * zoom
        : gx + frozenW + (cols.offset(line) - cols.offset(frozen.cols + 1)) * zoom - this.#scroller.scrollLeft;
    }
    return line <= frozen.rows
      ? gy + (rows.offset(line) - rows.offset(1)) * zoom
      : gy + frozenH + (rows.offset(line) - rows.offset(frozen.rows + 1)) * zoom - this.#scroller.scrollTop;
  }

  /** The header under a point of the canvas: `{axis: 'row' | 'col', line}`, `{axis: 'all'}` or null */
  #headerAt(x, y) {
    const { gx, gy } = this.#panes(this.#scroller.clientWidth, this.#scroller.clientHeight);
    if (x >= gx && y >= gy) return null;
    if (x < gx && y < gy) return { axis: 'all' };
    return y < gy ? { axis: 'col', line: this.#lineAt('col', x) } : { axis: 'row', line: this.#lineAt('row', y) };
  }

  /** The row or column whose end is at a point of a header: `{axis, line}`, or null */
  #resizeEdge(x, y) {
    if (this.#readOnly) return null;
    const header = this.#headerAt(x, y);
    if (!header || header.axis === 'all') return null;
    const { axis, line } = header;
    const sizes = axis === 'col' ? this.#sheet.cols : this.#sheet.rows;
    const pos = axis === 'col' ? x : y;
    const start = this.#linePos(axis, line);
    if (start + sizes.size(line) * this.zoom - pos <= RESIZE_EDGE_PX) return { axis, line };
    if (pos - start > RESIZE_EDGE_PX) return null;
    // The start of a line is the end of the one before it that can be seen
    let before = line - 1;
    while (before >= 1 && sizes.size(before) === 0) before--;
    return before >= 1 ? { axis, line: before } : null;
  }

  /** The lines a resize of `line` acts on: the selected ones when it is one of them */
  #linesToResize(axis, line) {
    const [a, b] = this.#selectedLines(axis);
    return this.#selection.whole === axis && line >= a && line <= b ? [a, b] : [line, line];
  }

  /** Follows the pointer from the edge of a header, showing the new size as it changes */
  #startResize({ axis, line }, e, box) {
    const isRow = axis === 'row';
    const sheet = this.#sheet;
    const original = isRow ? sheet.rows : sheet.cols;
    const [a, b] = this.#linesToResize(axis, line);
    const posOf = ev => (isRow ? ev.clientY - box.top : ev.clientX - box.left);
    const startPos = posOf(e);
    const startSize = original.size(line);
    const max = isRow ? Math.round(ptToPx(MAX_ROW_PT)) : colWidthToPx(MAX_COL_CHARS, this.#mdw);
    let size = startSize;
    const finish = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      window.removeEventListener('pointercancel', cancel);
      this.#gesture = null;
      this.#scroller.style.cursor = '';
    };
    const cancel = () => {
      finish();
      if (isRow) sheet.rows = original;
      else sheet.cols = original;
      this.#updateSpacer();
      this.#updateBar();
      this.render();
    };
    const move = ev => {
      size = Math.min(max, Math.max(MIN_LINE_PX, Math.round(startSize + (posOf(ev) - startPos) / this.zoom)));
      const custom = new Map(original.custom);
      for (let i = a; i <= b; i++) custom.set(i, size);
      const axisNow = new Axis(original.count, original.defaultSize, custom);
      if (isRow) sheet.rows = axisNow;
      else sheet.cols = axisNow;
      this.#updateSpacer();
      this.#nameBox.textContent = isRow ? `Alto: ${(size * 0.75).toFixed(2)}` : `Ancho: ${pxToColWidth(size, this.#mdw).toFixed(2)}`;
      this.render();
    };
    const up = () => {
      finish();
      if (size === startSize) {
        cancel();
        return;
      }
      const sizes = [];
      for (let i = a; i <= b; i++) sizes.push([i, size]);
      this.#applySizes(axis, sizes);
    };
    this.#gesture = { cancel };
    this.#scroller.style.cursor = isRow ? 'row-resize' : 'col-resize';
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
    window.addEventListener('pointercancel', cancel);
  }

  /** Sets rows or columns to sizes in pixels, as one undo step: `[[line, px or null]]` */
  #applySizes(axis, sizes) {
    if (!sizes.length) return;
    try {
      this.#workbook.checkpoint(this.#historyToken(null));
      const sheet = this.#activeSheet();
      // Runs of lines of one size are one call
      for (let i = 0; i < sizes.length;) {
        const [first, size] = sizes[i];
        let j = i;
        while (j + 1 < sizes.length && sizes[j + 1][1] === size && sizes[j + 1][0] === sizes[j][0] + 1) j++;
        const last = sizes[j][0];
        if (axis === 'row') sheet.setRowHeight(`${first}:${last}`, size === null ? null : Math.round(size * 0.75 * 4) / 4);
        else if (size !== null) sheet.setColumnWidth(`${columnName(first)}:${columnName(last)}`, pxToColWidth(size, this.#mdw));
        i = j + 1;
      }
    } catch (error) {
      this.#fail(error);
    }
    this.#workbook.setSelectionAfter(this.#historyToken(null));
    this.reload();
  }

  /** Double click: on the edge of a header fits its lines to their content, in a cell edits it */
  #onDoubleClick(e) {
    if (!this.#sheet) return;
    const box = this.#scroller.getBoundingClientRect();
    const x = e.clientX - box.left;
    const y = e.clientY - box.top;
    if (!this.#headerAt(x, y)) {
      if (this.#fillHandleAt(x, y)) this.#fillDownToData();
      else this.#beginEdit({ mode: 'edit' });
      return;
    }
    const edge = this.#resizeEdge(x, y);
    if (edge) this.#autofit(edge.axis, ...this.#linesToResize(edge.axis, edge.line));
  }

  /** Each of the lines `a..=b` as wide or tall as its content asks for */
  #autofit(axis, a, b) {
    const { cells, covered, mergeAt, rows, cols } = this.#sheet;
    const needed = new Map();
    const ctx = this.#ctx;
    ctx.save();
    for (const [at, cell] of cells) {
      if (!cell.text) continue;
      const r = Math.floor(at / 16_385);
      const c = at - r * 16_385;
      const line = axis === 'row' ? r : c;
      if (line < a || line > b || covered.has(at) || mergeAt.has(at)) continue;
      const xf = this.#styles.xfs[cell.style] || this.#styles.xfs[0];
      const font = xf?.font ?? this.#styles.defaultFont;
      ctx.font = canvasFont(font);
      const padding = PAD * 2 + (xf?.indent || 0) * 9;
      let size;
      if (axis === 'col') {
        // Wrapped text fits whatever the width: it does not ask for any
        if (xf?.wrap) continue;
        size = Math.ceil(ctx.measureText(cell.text).width + padding + 2);
      } else {
        const lines = xf?.wrap && cell.kind !== 'n' ? wrapLines(ctx, cell.text, Math.max(cols.size(c) - padding, 1)).length : 1;
        size = Math.ceil(lines * ptToPx(font.size || 11) * 1.22 + 4);
      }
      needed.set(line, Math.max(needed.get(line) ?? 0, size));
    }
    ctx.restore();
    const sizes = [];
    for (let line = a; line <= b; line++) {
      if ((axis === 'row' ? rows : cols).size(line) === 0) continue;
      const size = needed.get(line);
      // A row with nothing goes back to the automatic height; a column stays as it is
      if (size !== undefined) sizes.push([line, axis === 'row' ? Math.max(size, rows.defaultSize) : size]);
      else if (axis === 'row') sizes.push([line, null]);
    }
    this.#applySizes(axis, sizes);
  }

  /** The first and last lines of the selection along an axis */
  #selectedLines(axis) {
    const { first, last } = this.#selection;
    return axis === 'row' ? [first.row, last.row] : [first.col, last.col];
  }

  #updateCursor(e) {
    if (!this.#sheet || e.buttons || !this.#selection) return;
    const box = this.#scroller.getBoundingClientRect();
    const x = e.clientX - box.left;
    const y = e.clientY - box.top;
    const header = x > this.#scroller.clientWidth || y > this.#scroller.clientHeight ? null : this.#headerAt(x, y);
    let cursor = '';
    const edge = header ? this.#resizeEdge(x, y) : null;
    if (edge) {
      cursor = edge.axis === 'col' ? 'col-resize' : 'row-resize';
    } else if (header) {
      cursor = 'pointer';
      if (!this.#readOnly && this.#selection.whole === header.axis && header.axis !== 'all') {
        const [a, b] = this.#selectedLines(header.axis);
        if (header.line >= a && header.line <= b) cursor = 'grab';
      }
    }
    if (!header && this.#fillHandleAt(x, y)) cursor = 'crosshair';
    this.#scroller.style.cursor = cursor;
  }

  /** Whether a point of the canvas is on the handle at the corner of the selection */
  #fillHandleAt(x, y) {
    const sel = this.#selection;
    if (this.#readOnly || !sel || sel.whole || this.#edit) return false;
    const { last } = this.#expandToMerges(sel.first, sel.last);
    return Math.abs(x - this.#linePos('col', last.col + 1)) <= FILL_HANDLE_PX && Math.abs(y - this.#linePos('row', last.row + 1)) <= FILL_HANDLE_PX;
  }

  /** Follows the pointer from the fill handle: the range grows along one axis, whichever it left by more */
  #startFillDrag(box) {
    const source = this.#expandToMerges(this.#selection.first, this.#selection.last);
    const { first, last } = source;
    const finish = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      window.removeEventListener('pointercancel', cancel);
      this.#gesture = null;
      this.#scroller.style.cursor = '';
      const target = this.#fillDrag;
      this.#fillDrag = null;
      return target;
    };
    const cancel = () => {
      finish();
      this.render();
    };
    const move = ev => {
      const row = this.#lineAt('row', ev.clientY - box.top);
      const col = this.#lineAt('col', ev.clientX - box.left);
      const rows = row > last.row ? row - last.row : row < first.row ? row - first.row : 0;
      const cols = col > last.col ? col - last.col : col < first.col ? col - first.col : 0;
      let target = null;
      if (rows && Math.abs(rows) >= Math.abs(cols)) target = rows > 0 ? { first, last: { row, col: last.col } } : { first: { row, col: first.col }, last };
      else if (cols) target = cols > 0 ? { first, last: { row: last.row, col } } : { first: { row: first.row, col }, last };
      this.#fillDrag = target;
      if (target) {
        this.#extend(target.last.row, target.last.col);
        this.#scrollIntoView(rows && Math.abs(rows) >= Math.abs(cols) ? { row } : { col });
      }
      this.render();
    };
    const up = () => {
      const target = finish();
      if (target) this.#fillRange(rangeOf(first, last), target, true);
      this.render();
    };
    this.#gesture = { cancel };
    this.#scroller.style.cursor = 'crosshair';
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
    window.addEventListener('pointercancel', cancel);
  }

  /** Fills `target` from the range `source`, as one undo step, and selects it */
  #fillRange(source, target, series) {
    if (this.#readOnly || !this.#sheet) return false;
    this.#commitEdit(null, { focus: false });
    const changed = rangeOf(target.first, target.last);
    try {
      this.#workbook.checkpoint(this.#historyToken(changed));
      this.#activeSheet().fill(source, changed, { series });
    } catch (error) {
      this.#fail(error);
      return false;
    }
    this.#clip = null;
    this.#lastChange = changed;
    this.reload();
    const { active } = this.#selection;
    this.#setSelection(target.first, target.last, active, { scroll: false });
    this.#recordSelectionAfter();
    this.#emitChange(changed, 'fill');
    return true;
  }

  /** Double click on the handle: down as far as the data in the column beside the selection goes */
  #fillDownToData() {
    const { cells, used } = this.#sheet;
    const { first, last } = this.#expandToMerges(this.#selection.first, this.#selection.last);
    const filled = (row, col) => col >= 1 && !!cells.get(key(row, col))?.text;
    const beside = filled(last.row + 1, first.col - 1) ? first.col - 1 : last.col + 1;
    if (!filled(last.row + 1, beside)) return;
    let end = last.row + 1;
    while (end < used.maxRow && filled(end + 1, beside)) end++;
    this.#fillRange(rangeOf(first, last), { first, last: { row: end, col: last.col } }, true);
  }

  /** Ctrl+D and Ctrl+R: copies the first row or column of the selection over the rest (with one, the one before) */
  #fillSelection(axis) {
    if (this.#readOnly || this.#selection.whole) return;
    const { first, last } = this.#selection;
    const down = axis === 'row';
    const single = down ? first.row === last.row : first.col === last.col;
    if (single && (down ? first.row : first.col) === 1) return;
    const from = single
      ? { first: down ? { row: first.row - 1, col: first.col } : { row: first.row, col: first.col - 1 }, last: down ? { row: first.row - 1, col: last.col } : { row: last.row, col: first.col - 1 } }
      : { first, last: down ? { row: first.row, col: last.col } : { row: last.row, col: first.col } };
    this.#fillRange(rangeOf(from.first, from.last), { first: from.first, last }, false);
    if (single) this.#setSelection(first, last, first, { scroll: false });
  }

  /** Alt+=: the sum of the numbers next to the cell, to confirm; below or beside a selected range, written */
  #autosum() {
    if (this.#readOnly || this.#selection.whole) return;
    this.#commitEdit(null, { focus: false });
    const { cells, rows, cols } = this.#sheet;
    const isNumber = (row, col) => cells.get(key(row, col))?.number !== undefined;
    const isEmpty = (row, col) => !cells.get(key(row, col))?.kind;
    const { first, last, active } = this.#selection;
    if (first.row === last.row && first.col === last.col) {
      const { row, col } = active;
      const up = numberRun(r => isNumber(r, col), row, -1, rows.count);
      const left = up === null ? numberRun(c => isNumber(row, c), col, -1, cols.count) : null;
      const range = up !== null ? rangeOf({ row: up, col }, { row: row - 1, col }) : left !== null ? rangeOf({ row, col: left }, { row, col: col - 1 }) : '';
      this.#beginEdit({ mode: 'edit', text: `=SUM(${range})` });
      if (!range) this.#input.setSelectionRange(5, 5);
      return;
    }
    const sideways = first.row === last.row;
    const lastIsEmpty = sideways
      ? isEmpty(last.row, last.col)
      : Array.from({ length: last.col - first.col + 1 }, (_, i) => isEmpty(last.row, first.col + i)).every(Boolean);
    const lanes = sideways ? [first.row] : Array.from({ length: last.col - first.col + 1 }, (_, i) => first.col + i);
    const at = (sideways ? last.col : last.row) + (lastIsEmpty ? 0 : 1);
    const start = sideways ? first.col : first.row;
    const values = lanes.map(lane => {
      const cell = line => (sideways ? { row: lane, col: line } : { row: line, col: lane });
      let any = false;
      for (let line = start; line < at; line++) any ||= isNumber(cell(line).row, cell(line).col);
      if (!any || !isEmpty(cell(at).row, cell(at).col)) return undefined;
      return { formula: `SUM(${rangeOf(cell(start), cell(at - 1))})` };
    });
    if (values.every(v => v === undefined)) {
      this.#showMessage('No hay números que sumar, o el lugar de la suma ya tiene datos.');
      return;
    }
    const written = this.#write(sideways ? { row: first.row, col: at } : { row: at, col: first.col }, sideways ? [values] : [values], 'edit');
    if (written) {
      this.#setSelection(first, sideways ? { row: last.row, col: at } : { row: at, col: last.col }, active, { scroll: false });
      this.#recordSelectionAfter();
    }
  }

  /** Hides or shows the lines `a..=b`, as one undo step */
  #setLinesHidden(axis, a, b, hidden) {
    if (this.#readOnly || !this.#sheet) return false;
    this.#commitEdit(null, { focus: false });
    try {
      this.#workbook.checkpoint(this.#historyToken(null));
      const sheet = this.#activeSheet();
      if (axis === 'row') sheet.setRowsHidden(`${a}:${b}`, hidden);
      else sheet.setColumnsHidden(`${columnName(a)}:${columnName(b)}`, hidden);
      this.#workbook.setSelectionAfter(this.#historyToken(null));
    } catch (error) {
      this.#fail(error);
      return false;
    }
    this.reload();
    return true;
  }

  /** Whether any of the lines `a..=b` is hidden */
  #hasHidden(axis, a, b) {
    const { custom } = axis === 'row' ? this.#sheet.rows : this.#sheet.cols;
    for (const [line, size] of custom) if (size === 0 && line >= a && line <= b) return true;
    return false;
  }

  // ---------- Search ----------

  /** The hits found by the search panel are out of date */
  #staleFind() {
    if (this.#find) this.#find.hits = null;
  }

  /**
   * Cells whose text matches, sheet by sheet and row by row. With `inFormulas` the formula counts
   * instead of the text it shows, and `cell` has the formula or text to replace in; numbers and
   * dates are then left out.
   */
  #findHits(query, { matchCase, wholeCell, scope = 'sheet', inFormulas = false } = {}) {
    const pattern = searchPattern(query, { matchCase, wholeCell });
    if (!pattern || !this.#workbook) return [];
    const tabs = this.#workbook.tabs;
    const indexes = scope === 'workbook' ? tabs.flatMap((t, i) => (t.kind === 'worksheet' && t.state === 'visible' ? [i] : [])) : [this.#sheetIndex];
    const hits = [];
    for (const index of indexes) {
      const sheet = this.#workbook.sheet(index);
      const found = [];
      if (inFormulas) {
        let maxRow = 0;
        let maxCol = 0;
        for (const [row, col, , text] of sheet.cellViews(this.locale)) {
          if (!text) continue;
          maxRow = Math.max(maxRow, row);
          maxCol = Math.max(maxCol, col);
        }
        if (!maxRow) continue;
        sheet.getCells(`A1:${columnName(maxCol)}${maxRow}`).forEach((cells, r) => cells.forEach((cell, c) => {
          const text = cell.formula ? `=${cell.formula}` : cell.type === 'string' ? cell.value : null;
          if (text && matches(pattern, text)) found.push({ row: r + 1, col: c + 1, text, cell });
        }));
      } else if (index === this.#sheetIndex) {
        for (const [at, cell] of this.#sheet.cells) {
          if (!cell.text || !matches(pattern, cell.text)) continue;
          const row = Math.floor(at / 16_385);
          found.push({ row, col: at - row * 16_385, text: cell.text });
        }
      } else {
        for (const [row, col, , text] of sheet.cellViews(this.locale)) if (text && matches(pattern, text)) found.push({ row, col, text });
      }
      found.sort((a, b) => a.row - b.row || a.col - b.col);
      for (const hit of found) hits.push({ sheet: index, sheetName: tabs[index].name, ref: refOf(hit), ...hit });
    }
    return hits;
  }

  /** Writes the replacements, one block per sheet, as one undo step; returns how many cells changed */
  #replaceHits(hits, pattern, replacement) {
    if (!hits.length) return 0;
    this.#commitEdit(null, { focus: false });
    const bySheet = new Map();
    for (const hit of hits) {
      if (!bySheet.has(hit.sheet)) bySheet.set(hit.sheet, []);
      bySheet.get(hit.sheet).push(hit);
    }
    const blocks = [...bySheet].map(([index, found]) => {
      const first = { row: Math.min(...found.map(h => h.row)), col: Math.min(...found.map(h => h.col)) };
      const last = { row: Math.max(...found.map(h => h.row)), col: Math.max(...found.map(h => h.col)) };
      return { index, first, last, found };
    });
    const only = blocks.length === 1 && blocks[0].index === this.#sheetIndex ? rangeOf(blocks[0].first, blocks[0].last) : null;
    let written = false;
    try {
      this.#workbook.checkpoint(this.#historyToken(only));
      for (const { index, first, last, found } of blocks) {
        // Holes leave the cells between the hits as they are
        const rows = Array.from({ length: last.row - first.row + 1 }, () => new Array(last.col - first.col + 1));
        for (const hit of found) {
          const text = replaceMatches(pattern, hit.text, replacement);
          const value = hit.cell.formula ? { formula: text } : parseInput(text, this.locale).value;
          rows[hit.row - first.row][hit.col - first.col] = value === '' ? null : value;
        }
        this.#workbook.sheet(index).setValues(refOf(first), rows);
        written = true;
      }
      this.#workbook.setSelectionAfter(this.#historyToken(only));
    } catch (error) {
      // All or nothing, also when a later sheet refuses
      if (written) this.#workbook.undo();
      this.reload();
      this.#fail(error);
      return 0;
    }
    this.#clip = null;
    this.reload();
    for (const { index, first, last } of blocks) this.#emitChange(rangeOf(first, last), 'replace', this.#workbook.sheet(index));
    return hits.length;
  }

  #findOptions() {
    const { matchCase, wholeCell, scope, inFormulas, replacing } = this.#find;
    return { matchCase: matchCase.checked, wholeCell: wholeCell.checked, scope: scope.value, inFormulas: inFormulas.checked || replacing };
  }

  #findStatus(text) {
    this.#find.count.textContent = text;
  }

  /** Goes to the next hit after the active cell (or that one, with `here`), or the one before */
  #findStep(step, { here = false } = {}) {
    const find = this.#find;
    const query = find.query.value;
    if (!query) {
      this.#findStatus('');
      return;
    }
    try {
      find.hits ??= this.#findHits(query, this.#findOptions());
    } catch (error) {
      this.#fail(error);
      return;
    }
    const { hits } = find;
    if (!hits.length) {
      this.#findStatus('Sin resultados');
      return;
    }
    const { row, col } = this.#selection.active;
    const compare = hit => hit.sheet - this.#sheetIndex || hit.row - row || hit.col - col;
    let at;
    if (step > 0) {
      at = hits.findIndex(hit => (here ? compare(hit) >= 0 : compare(hit) > 0));
      if (at < 0) at = 0;
    } else {
      at = hits.findLastIndex(hit => compare(hit) < 0);
      if (at < 0) at = hits.length - 1;
    }
    const hit = hits[at];
    if (hit.sheet !== this.#sheetIndex) this.setSheet(hit.sheet);
    this.select(hit.ref);
    this.#findStatus(`${at + 1} de ${hits.length}`);
  }

  /** The replacement of the search panel on the active cell if it is a hit, and on to the next */
  #findReplaceOne() {
    const find = this.#find;
    const pattern = searchPattern(find.query.value, this.#findOptions());
    if (!pattern) return;
    find.hits ??= this.#findHits(find.query.value, this.#findOptions());
    const { row, col } = this.#selection.active;
    const hit = find.hits.find(h => h.sheet === this.#sheetIndex && h.row === row && h.col === col);
    if (hit) this.#replaceHits([hit], pattern, find.replacement.value);
    this.#findStep(1);
  }

  #findReplaceAll() {
    const find = this.#find;
    const pattern = searchPattern(find.query.value, this.#findOptions());
    if (!pattern) return;
    const hits = this.#findHits(find.query.value, this.#findOptions());
    const done = this.#replaceHits(hits, pattern, find.replacement.value);
    this.#findStatus(done === 1 ? '1 reemplazo' : `${done} reemplazos`);
  }

  /** The search panel, made the first time it is asked for */
  #findPanel() {
    if (this.#find) return this.#find;
    const el = document.createElement('div');
    el.className = 'xlsx-editor-find';
    el.setAttribute('role', 'search');
    el.innerHTML = `
      <div class="xlsx-editor-find-row">
        <input type="text" data-field="query" placeholder="Buscar" aria-label="Buscar" spellcheck="false" autocomplete="off" />
        <button type="button" data-action="previous" aria-label="Anterior" title="Anterior (Mayús+Enter)">‹</button>
        <button type="button" data-action="next" aria-label="Siguiente" title="Siguiente (Enter)">›</button>
        <button type="button" data-action="close" aria-label="Cerrar" title="Cerrar (Escape)">×</button>
      </div>
      <div class="xlsx-editor-find-row" data-field="replaceRow" hidden>
        <input type="text" data-field="replacement" placeholder="Reemplazar por" aria-label="Reemplazar por" spellcheck="false" autocomplete="off" />
        <button type="button" data-action="replace">Reemplazar</button>
        <button type="button" data-action="replaceAll">Todos</button>
      </div>
      <div class="xlsx-editor-find-row">
        <label><input type="checkbox" data-field="matchCase" /> Mayús/minús</label>
        <label><input type="checkbox" data-field="wholeCell" /> Celda completa</label>
        <label><input type="checkbox" data-field="inFormulas" /> En fórmulas</label>
        <select data-field="scope" aria-label="Dónde buscar">
          <option value="sheet">Esta hoja</option>
          <option value="workbook">Todo el libro</option>
        </select>
        <span class="xlsx-editor-find-count" role="status"></span>
      </div>`;
    const field = name => el.querySelector(`[data-field="${name}"]`);
    const find = {
      el,
      hits: null,
      timer: 0,
      replacing: false,
      query: field('query'),
      replacement: field('replacement'),
      replaceRow: field('replaceRow'),
      matchCase: field('matchCase'),
      wholeCell: field('wholeCell'),
      inFormulas: field('inFormulas'),
      scope: field('scope'),
      count: el.querySelector('.xlsx-editor-find-count'),
    };
    this.#find = find;
    // Native events of the inner fields are not the editor's, and its keys are not the grid's
    for (const type of ['change', 'input', 'selectionchange', 'keydown', 'copy', 'cut', 'paste']) el.addEventListener(type, e => e.stopPropagation());
    const search = () => {
      clearTimeout(find.timer);
      find.hits = null;
      find.timer = setTimeout(() => this.#findStep(1, { here: true }), FIND_DELAY_MS);
    };
    find.query.addEventListener('input', search);
    for (const option of [find.matchCase, find.wholeCell, find.inFormulas, find.scope]) option.addEventListener('change', search);
    const actions = {
      next: () => this.#findStep(1),
      previous: () => this.#findStep(-1),
      close: () => this.closeFind(),
      replace: () => this.#findReplaceOne(),
      replaceAll: () => this.#findReplaceAll(),
    };
    el.addEventListener('click', e => actions[e.target.closest?.('[data-action]')?.dataset.action]?.());
    el.addEventListener('keydown', e => {
      const mod = e.ctrlKey || e.metaKey;
      if (e.key === 'Escape') {
        e.preventDefault();
        this.closeFind();
      } else if (mod && e.key.toLowerCase() === 'f') {
        e.preventDefault();
        find.query.focus();
        find.query.select();
      } else if (e.key === 'Enter' && e.target === find.query) {
        e.preventDefault();
        clearTimeout(find.timer);
        this.#findStep(e.shiftKey ? -1 : 1);
      } else if (e.key === 'Enter' && e.target === find.replacement) {
        e.preventDefault();
        this.#findReplaceOne();
      }
    });
    this.#body.appendChild(el);
    return find;
  }

  #onPointerDown(e) {
    // Ctrl+click is the right button on a Mac: `contextmenu` takes it
    if (!this.#sheet || e.button !== 0 || e.ctrlKey) return;
    const box = this.#scroller.getBoundingClientRect();
    const x = e.clientX - box.left;
    const y = e.clientY - box.top;
    // Clicks on the scrollbars scroll
    if (x > this.#scroller.clientWidth || y > this.#scroller.clientHeight) return;
    this.#commitEdit(null, { focus: false });
    this.#scroller.focus({ preventScroll: true });
    const header = this.#headerAt(x, y);
    if (header) {
      const edge = this.#resizeEdge(x, y);
      if (edge) this.#startResize(edge, e, box);
      else this.#onHeaderDown(e, header, box);
      return;
    }
    if (!e.shiftKey && this.#fillHandleAt(x, y)) {
      this.#startFillDrag(box);
      return;
    }
    const at = this.#cellAt(x, y);
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

  /**
   * A press on a header: selects its row or column (with Shift, up to it; dragging, the ones
   * passed over), or on a line already selected starts moving the selected ones
   */
  #onHeaderDown(e, { axis, line }, box) {
    if (axis === 'all') {
      this.#selectLines('all', 1, 1, 1);
      return;
    }
    const sel = this.#selection;
    const [a, b] = this.#selectedLines(axis);
    if (!this.#readOnly && !e.shiftKey && sel.whole === axis && line >= a && line <= b) {
      this.#startLineDrag(axis, a, b - a + 1, line, e, box);
      return;
    }
    const anchor = e.shiftKey ? (axis === 'row' ? sel.active.row : sel.active.col) : line;
    this.#selectLines(axis, anchor, line, anchor);
    const move = ev => this.#selectLines(axis, anchor, this.#lineAt(axis, axis === 'row' ? ev.clientY - box.top : ev.clientX - box.left), anchor);
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  }

  /**
   * Follows the pointer from a selected header: past a few pixels the selected lines are being
   * moved (near the edges the view scrolls); a press that stays selects only the line pressed
   */
  #startLineDrag(axis, first, count, line, e, box) {
    const isRow = axis === 'row';
    const posOf = ev => (isRow ? ev.clientY - box.top : ev.clientX - box.left);
    const startPos = posOf(e);
    let lastPos = startPos;
    const aim = () => {
      // The nearest line between two: after the line under the pointer past its middle
      const under = this.#lineAt(axis, lastPos);
      const size = (isRow ? this.#sheet.rows : this.#sheet.cols).size(under) * this.zoom;
      const before = lastPos > this.#linePos(axis, under) + size / 2 ? under + 1 : under;
      this.#drag.before = before >= first && before <= first + count ? null : before;
      this.render();
    };
    const scrollTimer = setInterval(() => {
      if (!this.#drag) return;
      const { gx, gy, frozenW, frozenH } = this.#panes(this.#scroller.clientWidth, this.#scroller.clientHeight);
      const start = isRow ? gy + frozenH : gx + frozenW;
      const end = isRow ? this.#scroller.clientHeight : this.#scroller.clientWidth;
      let step = 0;
      if (lastPos < start + DRAG_SCROLL_EDGE_PX) step = -DRAG_SCROLL_STEP_PX;
      else if (lastPos > end - DRAG_SCROLL_EDGE_PX) step = DRAG_SCROLL_STEP_PX;
      if (!step) return;
      if (isRow) this.#scroller.scrollTop += step;
      else this.#scroller.scrollLeft += step;
      aim();
    }, 40);
    const finish = () => {
      clearInterval(scrollTimer);
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      window.removeEventListener('pointercancel', cancel);
      this.#scroller.style.cursor = '';
      const drag = this.#drag;
      this.#drag = null;
      return drag;
    };
    const cancel = () => {
      finish();
      this.render();
    };
    const move = ev => {
      lastPos = posOf(ev);
      if (!this.#drag) {
        if (Math.abs(lastPos - startPos) < DRAG_START_PX) return;
        this.#drag = { axis, first, count, before: null, cancel };
        this.#scroller.style.cursor = 'grabbing';
      }
      aim();
    };
    const up = () => {
      const drag = finish();
      if (!drag) this.#selectLines(axis, line, line, line);
      else if (drag.before !== null) this.#moveLines(axis, first, count, drag.before);
      this.render();
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
    window.addEventListener('pointercancel', cancel);
  }

  /** Selects whole rows or columns from `a` to `b` (or the whole sheet, with `axis` `all`) */
  #selectLines(axis, a, b, activeLine) {
    const { rows, cols } = this.#sheet;
    const options = { scroll: false, whole: axis };
    if (axis === 'row') this.#setSelection({ row: a, col: 1 }, { row: b, col: cols.count }, { row: activeLine, col: 1 }, options);
    else if (axis === 'col') this.#setSelection({ row: 1, col: a }, { row: rows.count, col: b }, { row: 1, col: activeLine }, options);
    else this.#setSelection({ row: 1, col: 1 }, { row: rows.count, col: cols.count }, { row: 1, col: 1 }, options);
  }

  /** The selected cells to act on: whole rows or columns only as far as the sheet has cells */
  #contentBounds() {
    const { first, last, whole } = this.#selection;
    if (!whole) return { first, last };
    const { maxRow, maxCol } = this.#sheet.used;
    return {
      first,
      last: {
        row: whole === 'row' ? last.row : Math.max(first.row, Math.min(last.row, maxRow)),
        col: whole === 'col' ? last.col : Math.max(first.col, Math.min(last.col, maxCol)),
      },
    };
  }

  /** The cells copying or clearing act on: a plain range takes in the merged cells it touches */
  #actionBounds() {
    const bounds = this.#contentBounds();
    return this.#selection.whole ? bounds : this.#expandToMerges(bounds.first, bounds.last);
  }

  /** Alt+arrow: the rows or columns of the selection one place up, down, left or right */
  #nudgeLines(axis, step) {
    const { whole } = this.#selection;
    if (whole && whole !== axis) return;
    const [first, last] = this.#selectedLines(axis);
    const before = step < 0 ? first - 1 : last + 2;
    if (before < 1 || before > (axis === 'row' ? MAX_ROWS : MAX_COLS) + 1) return;
    if (!this.#moveLines(axis, first, last - first + 1, before)) return;
    const [a, b] = this.#selectedLines(axis);
    const line = step < 0 ? a : b;
    this.#scrollIntoView(axis === 'row' ? { row: line } : { col: line });
  }

  #moveLines(axis, first, count, before) {
    const last = first + count - 1;
    if (before >= first && before <= last + 1) return false;
    const isRow = axis === 'row';
    return this.#editLines(axis, Math.min(first, before), Math.max(last, before - 1), count, 'move', sheet => (isRow ? sheet.moveRows(first, count, before) : sheet.moveColumns(first, count, before)));
  }

  #insertLines(axis, before, count) {
    const isRow = axis === 'row';
    return this.#editLines(axis, before, null, count, 'insert', sheet => (isRow ? sheet.insertRows(before, count) : sheet.insertColumns(before, count)));
  }

  #deleteLines(axis, first, count) {
    const isRow = axis === 'row';
    return this.#editLines(axis, first, null, count, 'delete', sheet => (isRow ? sheet.deleteRows(first, count) : sheet.deleteColumns(first, count)));
  }

  /**
   * One edit of whole lines as an undo step: `run` does it and returns where the lines start.
   * It changes the lines `from..=to`; `to` null reaches the end of what the sheet has (and of what
   * is pushed there). The lines end up selected.
   */
  #editLines(axis, from, to, count, source, run) {
    if (this.#readOnly || !this.#sheet) return false;
    this.#commitEdit(null, { focus: false });
    const isRow = axis === 'row';
    const { maxRow, maxCol } = this.#sheet.used;
    const end = Math.min(isRow ? MAX_ROWS : MAX_COLS, to ?? Math.max((isRow ? maxRow : maxCol) + count, from + count - 1));
    const changed = isRow ? rangeOf({ row: from, col: 1 }, { row: end, col: maxCol }) : rangeOf({ row: 1, col: from }, { row: maxRow, col: end });
    let start;
    try {
      this.#workbook.checkpoint(this.#historyToken(changed));
      start = run(this.#activeSheet());
    } catch (error) {
      this.#fail(error);
      return false;
    }
    this.#clip = null;
    this.#lastChange = changed;
    this.reload();
    this.#extend(isRow ? start + count : 1, isRow ? 1 : start + count);
    this.#selectLines(axis, start, start + count - 1, start);
    this.#recordSelectionAfter();
    this.#emitChange(changed, source);
    return true;
  }

  // ---------- Context menu ----------

  /** Right click: acts on what is selected, or on what is under the pointer when it is outside */
  #onContextMenu(e) {
    if (!this.#sheet || !this.#selection) return;
    const box = this.#scroller.getBoundingClientRect();
    const x = e.clientX - box.left;
    const y = e.clientY - box.top;
    if (x > this.#scroller.clientWidth || y > this.#scroller.clientHeight) return;
    e.preventDefault();
    this.#commitEdit(null, { focus: false });
    this.#scroller.focus({ preventScroll: true });
    const header = this.#headerAt(x, y);
    const { first, last, whole } = this.#selection;
    if (header?.axis === 'all') {
      this.#selectLines('all', 1, 1, 1);
    } else if (header) {
      const [a, b] = this.#selectedLines(header.axis);
      if (whole !== header.axis || header.line < a || header.line > b) this.#selectLines(header.axis, header.line, header.line, header.line);
    } else {
      const at = this.#cellAt(x, y);
      if (at.row < first.row || at.row > last.row || at.col < first.col || at.col > last.col) this.#setSelection(at, at, at, { scroll: false });
    }
    this.#openMenu(e.clientX, e.clientY);
  }

  /** Shift+F10 or the menu key: the menu at the active cell */
  #openMenuAtCell() {
    const { active } = this.#selection;
    const box = this.#scroller.getBoundingClientRect();
    this.#openMenu(box.left + this.#linePos('col', active.col) + 8, box.top + this.#linePos('row', active.row) + 8);
  }

  #menuItems() {
    const { whole } = this.#selection;
    const [r1, r2] = this.#selectedLines('row');
    const [c1, c2] = this.#selectedLines('col');
    const [rows, cols] = [r2 - r1 + 1, c2 - c1 + 1];
    const many = (n, one, several) => (n === 1 ? `una ${one}` : `${n} ${several}`);
    const shortcut = key => `${/Mac|iPhone|iPad/.test(globalThis.navigator?.platform ?? '') ? '⌘' : 'Ctrl+'}${key}`;
    const items = [];
    if (!this.#readOnly) items.push({ label: 'Cortar', keys: shortcut('X'), run: () => this.#copyToSystem(true) });
    items.push({ label: 'Copiar', keys: shortcut('C'), run: () => this.#copyToSystem(false) });
    if (this.#readOnly) return items;
    items.push(
      { label: 'Pegar', keys: shortcut('V'), run: () => this.#pasteFromSystem() },
      { label: 'Borrar contenido', keys: 'Supr', run: () => this.#clearSelection() },
    );
    if (whole === 'all') {
      if (this.#hasHidden('row', r1, r2) || this.#hasHidden('col', c1, c2)) {
        items.push(null, {
          label: 'Mostrar filas y columnas ocultas',
          run: () => {
            this.#setLinesHidden('row', r1, r2, false);
            this.#setLinesHidden('col', c1, c2, false);
          },
        });
      }
      return items;
    }
    items.push(null);
    if (whole !== 'col') {
      items.push(
        { label: `Insertar ${many(rows, 'fila', 'filas')} arriba`, run: () => this.#insertLines('row', r1, rows) },
        { label: `Insertar ${many(rows, 'fila', 'filas')} abajo`, run: () => this.#insertLines('row', r2 + 1, rows) },
      );
    }
    if (whole !== 'row') {
      items.push(
        { label: `Insertar ${many(cols, 'columna', 'columnas')} a la izquierda`, run: () => this.#insertLines('col', c1, cols) },
        { label: `Insertar ${many(cols, 'columna', 'columnas')} a la derecha`, run: () => this.#insertLines('col', c2 + 1, cols) },
      );
    }
    items.push(null);
    if (whole !== 'col') items.push({ label: rows === 1 ? 'Eliminar la fila' : `Eliminar ${rows} filas`, run: () => this.#deleteLines('row', r1, rows) });
    if (whole !== 'row') items.push({ label: cols === 1 ? 'Eliminar la columna' : `Eliminar ${cols} columnas`, run: () => this.#deleteLines('col', c1, cols) });
    if (whole === 'col') items.push(null, { label: 'Autoajustar el ancho', run: () => this.#autofit('col', c1, c2) });
    if (whole === 'row') items.push(null, { label: 'Autoajustar el alto', run: () => this.#autofit('row', r1, r2) });
    if (whole === 'col') {
      items.push({ label: cols === 1 ? 'Ocultar la columna' : `Ocultar ${cols} columnas`, run: () => this.#setLinesHidden('col', c1, c2, true) });
      if (this.#hasHidden('col', c1, c2)) items.push({ label: 'Mostrar columnas ocultas', run: () => this.#setLinesHidden('col', c1, c2, false) });
    }
    if (whole === 'row') {
      items.push({ label: rows === 1 ? 'Ocultar la fila' : `Ocultar ${rows} filas`, run: () => this.#setLinesHidden('row', r1, r2, true) });
      if (this.#hasHidden('row', r1, r2)) items.push({ label: 'Mostrar filas ocultas', run: () => this.#setLinesHidden('row', r1, r2, false) });
    }
    if (!whole) items.push(null, { label: 'Autosuma', keys: 'Alt+=', run: () => this.#autosum() });
    return items;
  }

  #openMenu(clientX, clientY) {
    this.#menu?.close();
    const menu = document.createElement('div');
    menu.className = 'xlsx-editor-menu';
    menu.setAttribute('role', 'menu');
    menu.setAttribute('aria-label', 'Acciones de la selección');
    const buttons = [];
    let close = () => {};
    for (const item of this.#menuItems()) {
      if (!item) {
        const line = document.createElement('div');
        line.className = 'xlsx-editor-menu-line';
        line.setAttribute('role', 'separator');
        menu.appendChild(line);
        continue;
      }
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'xlsx-editor-menu-item';
      button.setAttribute('role', 'menuitem');
      const label = document.createElement('span');
      label.textContent = item.label;
      button.appendChild(label);
      if (item.keys) {
        const keys = document.createElement('span');
        keys.className = 'xlsx-editor-menu-keys';
        keys.textContent = item.keys;
        button.appendChild(keys);
      }
      button.addEventListener('click', () => {
        close();
        this.#scroller.focus({ preventScroll: true });
        item.run();
      });
      menu.appendChild(button);
      buttons.push(button);
    }
    menu.addEventListener('contextmenu', e => e.preventDefault());
    menu.addEventListener('keydown', e => {
      e.stopPropagation();
      const at = buttons.indexOf(document.activeElement);
      const go = { ArrowDown: at + 1, ArrowUp: at - 1, Home: 0, End: buttons.length - 1 }[e.key];
      if (go !== undefined) {
        e.preventDefault();
        buttons[(go + buttons.length) % buttons.length].focus();
      } else if (e.key === 'Escape' || e.key === 'Tab') {
        e.preventDefault();
        close();
        this.#scroller.focus({ preventScroll: true });
      }
    });
    const outside = e => {
      if (!menu.contains(e.target)) close();
    };
    close = () => {
      document.removeEventListener('pointerdown', outside, true);
      window.removeEventListener('blur', close);
      window.removeEventListener('resize', close);
      this.#scroller.removeEventListener('scroll', close);
      menu.remove();
      this.#menu = null;
    };
    document.addEventListener('pointerdown', outside, true);
    window.addEventListener('blur', close);
    window.addEventListener('resize', close);
    this.#scroller.addEventListener('scroll', close);
    this.#root.appendChild(menu);
    this.#menu = { close };
    // Inside the window, also near its edges
    const { width, height } = menu.getBoundingClientRect();
    menu.style.left = `${Math.max(4, Math.min(clientX, window.innerWidth - width - 4))}px`;
    menu.style.top = `${Math.max(4, Math.min(clientY, window.innerHeight - height - 4))}px`;
    buttons[0]?.focus({ preventScroll: true });
  }

  /** Copy and cut from the menu: no clipboard event comes with them */
  #copyToSystem(cut) {
    let text = '';
    this.#onClipboard({ type: cut ? 'cut' : 'copy', preventDefault() {}, clipboardData: { setData: (_, t) => { text = t; } } });
    if (!navigator.clipboard?.writeText) {
      this.#showMessage('Este navegador no deja copiar desde el menú: usa el atajo de teclado.');
      return;
    }
    navigator.clipboard.writeText(text).catch(() => this.#showMessage('El navegador no dejó copiar desde el menú: usa el atajo de teclado.'));
  }

  #pasteFromSystem() {
    if (!navigator.clipboard?.readText) {
      this.#showMessage('Este navegador no deja pegar desde el menú: usa el atajo de teclado.');
      return;
    }
    navigator.clipboard.readText().then(text => this.#paste(text), () => this.#showMessage('El navegador no dejó pegar desde el menú: usa el atajo de teclado.'));
  }

  /** The engine refused a change: says why in the status bar and in an `error` event */
  #fail(error) {
    this.#showMessage(error.message);
    this.dispatchEvent(new CustomEvent('error', { detail: { error } }));
  }

  #showMessage(text) {
    clearTimeout(this.#messageTimer);
    this.#messageEl.textContent = text;
    this.#messageEl.title = text;
    if (text) this.#messageTimer = setTimeout(() => this.#showMessage(''), MESSAGE_SECONDS * 1000);
  }

  /** The count, average and sum of the selection, as many cells as it has */
  #updateStatus() {
    this.#statusDirty = false;
    if (this.#statusEl.hidden) return;
    const sel = this.#selection;
    const several = sel && this.#sheet && (sel.first.row !== sel.last.row || sel.first.col !== sel.last.col);
    const stats = several ? rangeStats(this.#sheet.cells, sel, this.#sheet) : null;
    const shown = stats ? `${stats.count}|${stats.sum}|${stats.average}|${stats.style}` : '';
    if (shown === this.#statusShown) return;
    this.#statusShown = shown;
    const parts = [];
    if (stats) {
      let count = String(stats.count);
      try {
        count = new Intl.NumberFormat(this.locale).format(stats.count);
      } catch {
        // A language the browser does not know: plain digits
      }
      if (stats.sum === null) {
        parts.push(`Recuento: ${count}`);
      } else {
        const [average, sum] = this.#workbook.formatNumbers(stats.style, [stats.average, stats.sum], this.locale);
        parts.push(`Promedio: ${average}`, `Recuento: ${count}`, `Suma: ${sum}`);
      }
    }
    this.#statsEl.replaceChildren(...parts.map(text => {
      const part = document.createElement('span');
      part.textContent = text;
      return part;
    }));
  }

  #onKeyDown(e) {
    if (this.#drag || this.#gesture) {
      if (e.key === 'Escape') (this.#drag ?? this.#gesture).cancel();
      return;
    }
    if (!this.#sheet || !this.#selection || e.isComposing) return;
    if (e.key === 'ContextMenu' || (e.key === 'F10' && e.shiftKey)) {
      e.preventDefault();
      this.#openMenuAtCell();
      return;
    }
    if (this.#navigate(e)) {
      e.preventDefault();
      return;
    }
    if (e.altKey && e.code === 'Equal' && !e.ctrlKey && !e.metaKey) {
      // By the physical key: on a Mac, Option+= types another character
      e.preventDefault();
      this.#autosum();
      return;
    }
    if (e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey && !this.#readOnly) {
      const nudge = { ArrowUp: ['row', -1], ArrowDown: ['row', 1], ArrowLeft: ['col', -1], ArrowRight: ['col', 1] }[e.key];
      if (nudge) {
        e.preventDefault();
        this.#nudgeLines(...nudge);
        return;
      }
    }
    if (e.key === ' ' && (e.ctrlKey || e.shiftKey) && !e.metaKey && !e.altKey) {
      // Ctrl+Space selects the columns of the selection and Shift+Space its rows, as in Excel
      e.preventDefault();
      const axis = e.ctrlKey && e.shiftKey ? 'all' : e.ctrlKey ? 'col' : 'row';
      const [a, b] = axis === 'all' ? [1, 1] : this.#selectedLines(axis);
      const { active } = this.#selection;
      const anchor = axis === 'row' ? active.row : active.col;
      this.#selectLines(axis, a === anchor ? a : b, a === anchor ? b : a, anchor);
      return;
    }
    if ((e.ctrlKey || e.metaKey) && !e.altKey) {
      const k = e.key.toLowerCase();
      const { whole } = this.#selection;
      if (k === 'a') {
        e.preventDefault();
        this.#selectLines('all', 1, 1, 1);
      } else if (k === 'f') {
        e.preventDefault();
        this.openFind();
      } else if (k === 'h' && e.ctrlKey && !this.#readOnly) {
        e.preventDefault();
        this.openFind({ replace: true });
      } else if (k === 'd' && !this.#readOnly) {
        e.preventDefault();
        this.#fillSelection('row');
      } else if (k === 'r' && e.ctrlKey && !e.shiftKey && !this.#readOnly) {
        // Cmd+R stays the browser's reload
        e.preventDefault();
        this.#fillSelection('col');
      } else if (e.ctrlKey && (e.code === 'Digit9' || e.code === 'Digit0') && !this.#readOnly) {
        // Ctrl+9 hides rows and Ctrl+0 columns; with Shift they show them again
        e.preventDefault();
        const axis = e.code === 'Digit9' ? 'row' : 'col';
        const [a, b] = this.#selectedLines(axis);
        if (this.#selection.whole === (axis === 'row' ? 'col' : 'row') && !e.shiftKey) return;
        this.#setLinesHidden(axis, a, b, !e.shiftKey);
      } else if ((k === '-' || k === '+' || (k === '=' && e.shiftKey)) && (whole === 'row' || whole === 'col') && !this.#readOnly) {
        // Ctrl+- and Ctrl++ on whole lines, as in Excel (on anything else they are the browser's zoom)
        e.preventDefault();
        const [a, b] = this.#selectedLines(whole);
        if (k === '-') this.#deleteLines(whole, a, b - a + 1);
        else this.#insertLines(whole, a, b - a + 1);
      } else if (k === 'z' || k === 'y') {
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

  /**
   * Keys that go far: Ctrl/Cmd+arrow to the edge of the data, Home to the start of the row,
   * Ctrl/Cmd+Home and End to the corners of the data, PageUp and PageDown a screen (with Alt,
   * sideways); with Shift they extend the selection. Returns whether the key was one of them.
   */
  #navigate(e) {
    const mod = e.ctrlKey || e.metaKey;
    const extend = e.shiftKey;
    const { rows, cols, frozen, cells, used } = this.#sheet;
    const from = extend ? this.#selection.cursor : this.#selection.active;
    let target = null;
    const arrow = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
    if (arrow && mod && !e.altKey) {
      const filled = (row, col) => !!cells.get(key(row, col))?.kind && rows.size(row) !== 0 && cols.size(col) !== 0;
      target = arrow[0]
        ? { row: dataEdge(r => filled(r, from.col), from.row, arrow[0], MAX_ROWS), col: from.col }
        : { row: from.row, col: dataEdge(c => filled(from.row, c), from.col, arrow[1], MAX_COLS) };
    } else if (e.key === 'Home' && !e.altKey) {
      target = mod ? { row: frozen.rows + 1, col: frozen.cols + 1 } : { row: from.row, col: frozen.cols + 1 };
    } else if (e.key === 'End' && mod && !e.altKey) {
      target = { row: used.maxRow, col: used.maxCol };
    } else if ((e.key === 'PageDown' || e.key === 'PageUp') && !mod) {
      const step = e.key === 'PageDown' ? 1 : -1;
      const { gx, gy, frozenW, frozenH } = this.#panes(this.#scroller.clientWidth, this.#scroller.clientHeight);
      if (e.altKey) {
        const page = (this.#scroller.clientWidth - gx - frozenW) / this.zoom;
        // Going back rounds the other way, to come back to the line a page forward started from
        const col = cols.indexAt(cols.offset(from.col) + step * page + (step < 0 ? cols.defaultSize - 1 : 0));
        this.#extend(1, col);
        this.#scroller.scrollLeft += (this.#sheet.cols.offset(col) - cols.offset(from.col)) * this.zoom;
        target = { row: from.row, col };
      } else {
        const page = (this.#scroller.clientHeight - gy - frozenH) / this.zoom;
        // Past the shown grid when there is room to grow
        this.#extend(rows.indexAt(rows.offset(from.row) + page) + 1, 1);
        const axis = this.#sheet.rows;
        const row = axis.indexAt(axis.offset(from.row) + step * page + (step < 0 ? axis.defaultSize - 1 : 0));
        this.#scroller.scrollTop += (axis.offset(row) - axis.offset(from.row)) * this.zoom;
        target = { row, col: from.col };
      }
    }
    if (!target) return false;
    this.#extend(target.row, target.col);
    const anchor = extend ? this.#selection.active : target;
    this.#setSelection(anchor, target, anchor);
    return true;
  }

  /** Moves the active cell (or, with `extend`, the moving corner of the range) one step */
  #move(dr, dc, extend = false) {
    const { whole } = this.#selection;
    if (extend && (whole === 'row' || whole === 'col')) {
      // Mayús+arrow along the lines selected takes one more or one less
      const step = whole === 'row' ? dr : dc;
      if (!step) return;
      const { active, cursor } = this.#selection;
      const anchor = whole === 'row' ? active.row : active.col;
      const limit = whole === 'row' ? this.#sheet.rows.count : this.#sheet.cols.count;
      const line = Math.min(Math.max((whole === 'row' ? cursor.row : cursor.col) + step, 1), limit);
      this.#selectLines(whole, anchor, line, anchor);
      this.#extend(whole === 'row' ? line : 1, whole === 'col' ? line : 1);
      this.#scrollIntoView(whole === 'row' ? { row: line } : { col: line });
      return;
    }
    const from = extend ? this.#selection.cursor : this.#selection.active;
    this.#extend(from.row + Math.max(dr, 0), from.col + Math.max(dc, 0));
    const { rows, cols } = this.#sheet;
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

  #setSelection(a, b, active, { scroll = true, whole = null } = {}) {
    const first = { row: Math.min(a.row, b.row), col: Math.min(a.col, b.col) };
    const last = { row: Math.max(a.row, b.row), col: Math.max(a.col, b.col) };
    const prev = this.#selection;
    // `cursor` is the corner that moves when the range is extended with Shift
    this.#selection = { first, last, active: { ...active }, cursor: { ...b }, whole };
    if (scroll) this.#scrollIntoView(b);
    this.#statusDirty = true;
    this.#updateBar();
    this.render();
    if (!prev || prev.first.row !== first.row || prev.first.col !== first.col || prev.last.row !== last.row || prev.last.col !== last.col) {
      if (this.#messageEl.textContent) this.#showMessage('');
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
    this.#nameBox.textContent = sel ? this.#selectionName(sel) : '';
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

  /** What the name box shows: `B:D` and `2:5` for whole columns and rows */
  #selectionName(sel) {
    const { first, last } = this.#selection;
    if (sel.whole === 'col') return `${columnName(first.col)}:${columnName(last.col)}`;
    if (sel.whole === 'row') return `${first.row}:${last.row}`;
    return sel.range;
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
    const zoom = this.zoom;
    const width = this.#scroller.clientWidth;
    const height = this.#scroller.clientHeight;
    const xOf = c => this.#linePos('col', c);
    const yOf = r => this.#linePos('row', r);
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
      const bounds = this.#contentBounds();
      const first = fill ? bounds.first : { row: edit.row, col: edit.col };
      const last = fill ? bounds.last : first;
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
    return JSON.stringify({ sheet: this.#sheetIndex, range: sel?.range, active: sel?.ref, whole: sel?.whole ?? null, changed });
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
      this.#fail(error);
      return null;
    }
    this.#lastChange = changed;
    this.#workbook.setSelectionAfter(this.#historyToken(changed));
    this.reload();
    this.#emitChange(changed, source);
    return { first, last };
  }

  #emitChange(range, source, sheet = this.#activeSheet()) {
    const { first, last } = parseRange(range);
    const cells = [];
    if ((last.row - first.row + 1) * (last.col - first.col + 1) <= MAX_CHANGE_CELLS) {
      sheet.getValues(range).forEach((row, r) => row.forEach((value, c) => {
        cells.push({ ref: refOf({ row: first.row + r, col: first.col + c }), value });
      }));
    }
    this.dispatchEvent(new CustomEvent('change', { detail: { sheet: sheet.name, index: sheet.index, range, cells, source } }));
  }

  /** Delete and Backspace: empties the selected cells, keeping their format */
  #clearSelection(source = 'clear') {
    const { first, last } = this.#actionBounds();
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
    this.#staleFind();
    this.#styles = this.#workbook.styleView();
    const tabs = this.#workbook.tabs;
    this.#renderTabs();
    if (token && token.sheet !== this.#sheetIndex && tabs[token.sheet]?.kind === 'worksheet') {
      this.#loadSheet(token.sheet);
      this.dispatchEvent(new CustomEvent('sheetchange', { detail: { index: token.sheet, name: tabs[token.sheet].name } }));
    } else {
      const index = this.#validSheet(this.#sheetIndex);
      this.#loadSheet(index, { keepScroll: index === this.#sheetIndex });
      if (index !== token?.sheet && tabs[index]) this.dispatchEvent(new CustomEvent('sheetchange', { detail: { index, name: tabs[index].name } }));
    }
    if (token?.range) {
      const range = parseRange(token.range);
      const active = token.active ? parseRange(token.active).first : range.first;
      const whole = token.whole ?? null;
      this.#setSelection(range.first, range.last, active, { whole, scroll: !whole });
      this.#fitWholeSelection();
      // Whole lines end far away: only their first one needs to be in view
      if (whole === 'row') this.#scrollIntoView({ row: active.row });
      if (whole === 'col') this.#scrollIntoView({ col: active.col });
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
    const { first, last } = this.#actionBounds();
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
    const { first, whole } = this.#selection;
    // Whole rows or columns take the block from their first cell, without repeating it
    const last = whole ? first : this.#selection.last;
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
