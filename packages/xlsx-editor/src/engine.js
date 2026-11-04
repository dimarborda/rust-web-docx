// Excel (.xlsx) workbooks backed by the Rust/WebAssembly engine, without any interface: works in
// Node (`@dimarborda/xlsx-editor/engine`) and in the browser. In Node the engine is read
// from the package itself; in the browser bundlers serve the .wasm next to the JS, or pass its
// location with `initEngine(url)`.
import init, { initSync, XlsxWorkbook } from '../wasm/xlsx_engine.js';

let loading = null;

const isNode = typeof process !== 'undefined' && !!process.versions?.node;

/**
 * Starts the engine. Safe to call many times: later calls wait for the first one.
 * @param {string | URL | BufferSource} [wasm] where `xlsx_engine_bg.wasm` is served from, or its bytes
 */
export function initEngine(wasm) {
  loading ??= (async () => {
    if (wasm && (ArrayBuffer.isView(wasm) || wasm instanceof ArrayBuffer)) {
      initSync({ module: wasm });
    } else if (!wasm && isNode) {
      // Hidden from bundlers: this branch only runs in Node
      const fs = 'node:fs/promises';
      const { readFile } = await import(/* @vite-ignore */ /* webpackIgnore: true */ fs);
      initSync({ module: await readFile(new URL('../wasm/xlsx_engine_bg.wasm', import.meta.url)) });
    } else {
      await init(wasm ? { module_or_path: wasm } : undefined);
    }
  })().catch(err => {
    loading = null; // a later call may retry (e.g. after fixing the location)
    throw err;
  });
  return loading;
}

/** Runs an engine call, turning the engine's error messages into `Error`s */
function call(fn) {
  try {
    return fn();
  } catch (err) {
    throw typeof err === 'string' ? new Error(err) : err;
  }
}

/** A cell as the engine returns it, with dates turned into `Date` objects */
function toCell(raw) {
  const cell = { ref: raw.ref, type: raw.type, value: raw.type === 'date' ? new Date(raw.value) : raw.value };
  if (raw.formula !== undefined) cell.formula = raw.formula;
  return cell;
}

/** A value to write, in the engine's JSON form; `undefined` leaves the cell as it is */
function toInput(value, where) {
  if (value === undefined) return { skip: true };
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return value;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new TypeError(`El valor de ${where} no es un número finito: ${value}`);
    return value;
  }
  if (value instanceof Date) {
    const ms = value.getTime();
    if (Number.isNaN(ms)) throw new TypeError(`La fecha de ${where} no es válida.`);
    return { date: ms };
  }
  if (value && typeof value === 'object' && typeof value.formula === 'string') return { formula: value.formula };
  throw new TypeError(`Valor no admitido en ${where}: usa texto, número, booleano, Date, { formula } o null.`);
}

/** Column letters (`'B'`, `'b'`) or number (2) → column number */
function columnNumber(col) {
  if (Number.isInteger(col) && col >= 1) return col;
  if (typeof col === 'string' && /^[A-Za-z]{1,3}$/.test(col.trim())) {
    return [...col.trim().toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  }
  throw new TypeError(`Columna no válida: ${col} (usa una letra como 'B' o un número desde 1).`);
}

/** A row or column number, or a count, for the engine */
function lineNumber(value, what) {
  if (!Number.isInteger(value) || value < 0) throw new TypeError(`${what} debe ser un entero positivo: ${value}`);
  return value;
}

/** One worksheet of a workbook */
export class Sheet {
  /**
   * @param {Workbook} workbook
   * @param {number} index
   * @param {string} name
   */
  constructor(workbook, index, name) {
    this.workbook = workbook;
    this.index = index;
    this.name = name;
  }

  /** The cell at `ref` (`B3`): its type, its value and its formula, if any */
  getCell(ref) {
    return toCell(JSON.parse(call(() => this.workbook._engine.cell_json(this.index, ref))));
  }

  /** The value of the cell at `ref`; `null` when it is empty */
  getValue(ref) {
    return this.getCell(ref).value;
  }

  /** The cells of a range (`A1:C3`), by row: their type, value and formula, if any */
  getCells(range) {
    return JSON.parse(call(() => this.workbook._engine.range_json(this.index, range))).map(row => row.map(toCell));
  }

  /** The values of a range (`A1:C3`), by row */
  getValues(range) {
    return JSON.parse(call(() => this.workbook._engine.range_json(this.index, range))).map(row => row.map(c => toCell(c).value));
  }

  /**
   * Writes one value; `null` clears the cell (its format stays) and `{ formula: '=SUM(A1:A3)' }`
   * writes a formula
   */
  setValue(ref, value) {
    this.setValues(ref, [[value]]);
  }

  /**
   * Writes rows of values from the cell `start` rightwards and down. `undefined` (or a hole)
   * leaves a cell as it is; `null` clears it; `{ formula }` writes a formula. The formulas that
   * depend on the cells written are recalculated.
   */
  setValues(start, rows) {
    if (!Array.isArray(rows) || !rows.every(Array.isArray)) {
      throw new TypeError('setValues espera una lista de filas, por ejemplo [[1, "a"], [2, "b"]].');
    }
    const json = rows.map((row, r) => Array.from(row, (v, c) => toInput(v, `${start} +${r},${c}`)));
    call(() => this.workbook._engine.set_values_json(this.index, start, JSON.stringify(json)));
  }

  /**
   * Drawing API: the grid of the sheet (column widths in characters, row heights in points,
   * merged ranges, frozen panes, gridlines)
   */
  layout() {
    return JSON.parse(call(() => this.workbook._engine.sheet_layout_json(this.index)));
  }

  /**
   * Drawing API: every cell that shows something, with the text Excel shows for it in the
   * language `locale`, as `[row, col, style, text, kind, color?, number?]` (`kind`: `n`, `s`,
   * `b`, `e` or `''` for an empty cell with a format; `style` indexes
   * `workbook.styleView().xfs`; `number` is the value of `n` cells, dates as serial numbers)
   * @param {string} [locale]
   */
  cellViews(locale = 'en') {
    return JSON.parse(call(() => this.workbook._engine.sheet_cells_json(this.index, locale)));
  }

  /** Areas of this sheet (see `Workbook.areas`) */
  areas() {
    return this.workbook.areas({ sheet: this.index });
  }

  /**
   * The region around a cell, as Excel's Ctrl+A selects it (or the table, or the cell alone):
   * an area like those of `areas()`
   */
  regionAt(ref) {
    return JSON.parse(call(() => this.workbook._engine.region_at_json(this.index, ref)));
  }

  /** Context of a range of this sheet or of a named area (see `Workbook.area`) */
  area(target, options = {}) {
    return this.workbook.area(target, { ...options, sheet: this.index });
  }

  /**
   * Formats a range (`A1:C3`): `{bold, italic, underline, strike, color, fontSize, fontName,
   * fill, numberFormat, align, valign, wrap, border}`. What is not given stays as it is.
   */
  setFormat(range, format) {
    if (!format || typeof format !== 'object') throw new TypeError('setFormat espera un objeto, por ejemplo { bold: true }.');
    call(() => this.workbook._engine.set_format_json(this.index, range, JSON.stringify(format)));
  }

  /** Back to the default format for the cells of a range (values stay) */
  clearFormat(range) {
    call(() => this.workbook._engine.clear_format(this.index, range));
  }

  /**
   * Width of a column or columns (`'B'`, `'B:D'`, `2`) in characters, as Excel measures it
   * @param {string | number} columns
   * @param {number} width
   */
  setColumnWidth(columns, width) {
    const [a, b = a] = typeof columns === 'string' ? columns.split(':') : [columns];
    if (typeof width !== 'number' || !Number.isFinite(width)) throw new TypeError('El ancho debe ser un número de caracteres.');
    call(() => this.workbook._engine.set_column_width(this.index, columnNumber(a), columnNumber(b), width));
  }

  /**
   * Moves whole rows, as Excel's "cut and insert": `count` rows from `first` go before the row
   * `before` (numbered as they are now). Formats, heights, merged cells, formulas of the whole
   * workbook, defined names, charts, tables, validations, conditional formats, images and
   * comments follow. Throws, and changes nothing, when it would tear merged cells, an array
   * formula, a table or a filter. Returns where the first row ends up.
   * @param {number} first
   * @param {number} count
   * @param {number} before
   */
  moveRows(first, count, before) {
    const [a, n, b] = [lineNumber(first, 'La fila'), lineNumber(count, 'La cantidad'), lineNumber(before, 'El destino')];
    return call(() => this.workbook._engine.move_rows(this.index, a, n, b));
  }

  /**
   * Moves whole columns (see `moveRows`); `first` and `before` are letters (`'B'`) or numbers
   * @param {string | number} first
   * @param {number} count
   * @param {string | number} before
   */
  moveColumns(first, count, before) {
    const [a, n, b] = [columnNumber(first), lineNumber(count, 'La cantidad'), columnNumber(before)];
    return call(() => this.workbook._engine.move_columns(this.index, a, n, b));
  }

  /**
   * Inserts `count` empty rows before the row `before`, with the height and format of the one
   * above. Formulas, names, charts and everything that `moveRows` follows are updated; ranges and
   * merged cells that the new rows fall inside grow, also tables. Returns `before`.
   * @param {number} before
   * @param {number} [count]
   */
  insertRows(before, count = 1) {
    const [b, n] = [lineNumber(before, 'La fila'), lineNumber(count, 'La cantidad')];
    return call(() => this.workbook._engine.insert_rows(this.index, b, n));
  }

  /** Inserts empty columns before `before` (letters or number), like the one to their left */
  insertColumns(before, count = 1) {
    const [b, n] = [columnNumber(before), lineNumber(count, 'La cantidad')];
    return call(() => this.workbook._engine.insert_columns(this.index, b, n));
  }

  /**
   * Deletes `count` rows from `first`. What pointed only at them becomes `#REF!`; ranges and
   * merged cells shrink, also tables. Returns `first`.
   * @param {number} first
   * @param {number} [count]
   */
  deleteRows(first, count = 1) {
    const [a, n] = [lineNumber(first, 'La fila'), lineNumber(count, 'La cantidad')];
    return call(() => this.workbook._engine.delete_rows(this.index, a, n));
  }

  /** Deletes columns from `first` (letters or number), see `deleteRows` */
  deleteColumns(first, count = 1) {
    const [a, n] = [columnNumber(first), lineNumber(count, 'La cantidad')];
    return call(() => this.workbook._engine.delete_columns(this.index, a, n));
  }

  /**
   * Height of a row or rows (`3`, `'3:5'`) in points; `null` goes back to the automatic height
   * @param {string | number} rows
   * @param {number | null} height
   */
  setRowHeight(rows, height) {
    const [a, b = a] = typeof rows === 'string' ? rows.split(':').map(Number) : [rows];
    if (height !== null && (typeof height !== 'number' || !Number.isFinite(height))) throw new TypeError('El alto debe ser un número de puntos o null.');
    const [first, last] = [lineNumber(a, 'La fila'), lineNumber(b, 'La fila')];
    call(() => this.workbook._engine.set_row_height(this.index, first, last, height ?? undefined));
  }

  /**
   * Hides or shows rows (`3`, `'3:5'`); they keep their cells and height
   * @param {string | number} rows
   * @param {boolean} hidden
   */
  setRowsHidden(rows, hidden) {
    const [a, b = a] = typeof rows === 'string' ? rows.split(':').map(Number) : [rows];
    const [first, last] = [lineNumber(a, 'La fila'), lineNumber(b, 'La fila')];
    call(() => this.workbook._engine.set_rows_hidden(this.index, first, last, !!hidden));
  }

  /**
   * Hides or shows columns (`'B'`, `'B:D'`, `2`); they keep their cells and width
   * @param {string | number} columns
   * @param {boolean} hidden
   */
  setColumnsHidden(columns, hidden) {
    const [a, b = a] = typeof columns === 'string' ? columns.split(':') : [columns];
    const [first, last] = [columnNumber(a), columnNumber(b)];
    call(() => this.workbook._engine.set_columns_hidden(this.index, first, last, !!hidden));
  }

  /**
   * Fills `target` from `source`, the cells at one of its ends, as Excel's fill handle: formats
   * repeat, formulas move their references, and with `series` (by default) numbers continue their
   * trend, one date goes on day by day, texts ending in a number count on and names of months and
   * days follow their order. Refuses to fill over merged cells.
   * @param {string} source e.g. `A1:A2`
   * @param {string} target e.g. `A1:A20`
   * @param {{series?: boolean}} [options]
   */
  fill(source, target, { series = true } = {}) {
    call(() => this.workbook._engine.fill_range(this.index, source, target, !!series));
  }

  /** Merges a range into one cell; only its first cell may hold a value */
  mergeCells(range) {
    call(() => this.workbook._engine.merge_cells(this.index, range));
  }

  /** Splits the merged cells that overlap a range; returns how many */
  unmergeCells(range) {
    return call(() => this.workbook._engine.unmerge_cells(this.index, range));
  }
}

/** An open workbook */
export class Workbook {
  /** @param {XlsxWorkbook} engine */
  constructor(engine) {
    this._engine = engine;
  }

  /** Sheet names, in tab order */
  get sheetNames() {
    return this._engine.sheet_names();
  }

  /** The worksheets (chart sheets have no cells and are left out) */
  get sheets() {
    return JSON.parse(this._engine.sheets_json())
      .map((s, index) => (s.kind === 'worksheet' ? new Sheet(this, index, s.name) : null))
      .filter(Boolean);
  }

  /** Every tab in order: `{name, kind: 'worksheet' | 'chartsheet' | 'other', state}` */
  get tabs() {
    return JSON.parse(call(() => this._engine.sheets_json()));
  }

  /** Drawing API: the cell formats resolved to fonts, fills, borders and alignment */
  styleView() {
    return JSON.parse(call(() => this._engine.styles_view_json()));
  }

  /**
   * Numbers as a cell format shows them: `style` indexes `styleView().xfs`
   * @param {number} style
   * @param {number[]} values
   * @param {string} [locale]
   */
  formatNumbers(style, values, locale = 'en') {
    return call(() => this._engine.format_numbers(style, Float64Array.from(values), locale));
  }

  /**
   * A worksheet by name (ignoring case) or by tab position, from 0
   * @param {string | number} nameOrIndex
   */
  sheet(nameOrIndex) {
    const all = JSON.parse(this._engine.sheets_json());
    let index = nameOrIndex;
    if (typeof nameOrIndex === 'string') {
      index = all.findIndex(s => s.name === nameOrIndex);
      if (index < 0) index = all.findIndex(s => s.name.toLowerCase() === nameOrIndex.toLowerCase());
      if (index < 0) throw new Error(`El libro no tiene ninguna hoja llamada '${nameOrIndex}'.`);
    } else if (!Number.isInteger(index) || index < 0 || index >= all.length) {
      throw new RangeError(`El libro tiene ${all.length} hoja(s); no existe la hoja ${nameOrIndex}.`);
    }
    if (all[index].kind !== 'worksheet') throw new Error(`La hoja '${all[index].name}' no tiene celdas.`);
    return new Sheet(this, index, all[index].name);
  }

  /**
   * Adds an empty worksheet after the tab `after` (name or position; at the end without it).
   * Names follow Excel's rules: 1–31 characters, none of `[ ] : * ? / \`, unique ignoring case.
   * Returns the new sheet.
   * @param {string} name
   * @param {string | number} [after]
   */
  addSheet(name, after) {
    const index = call(() => this._engine.add_sheet(String(name), after === undefined || after === null ? undefined : this.#tabIndex(after)));
    return this.sheet(index);
  }

  /**
   * Renames a tab (name or position) and the references to it in formulas, defined names,
   * data validations, conditional formats and charts; returns how many of those changed
   * @param {string | number} sheet
   * @param {string} name
   */
  renameSheet(sheet, name) {
    return call(() => this._engine.rename_sheet(this.#tabIndex(sheet), String(name)));
  }

  /** Position of a tab by name (ignoring case) or position */
  #tabIndex(nameOrIndex) {
    const all = JSON.parse(this._engine.sheets_json());
    if (typeof nameOrIndex === 'number') {
      if (!Number.isInteger(nameOrIndex) || nameOrIndex < 0 || nameOrIndex >= all.length) {
        throw new RangeError(`El libro tiene ${all.length} hoja(s); no existe la hoja ${nameOrIndex}.`);
      }
      return nameOrIndex;
    }
    let index = all.findIndex(s => s.name === nameOrIndex);
    if (index < 0) index = all.findIndex(s => s.name.toLowerCase() === String(nameOrIndex).toLowerCase());
    if (index < 0) throw new Error(`El libro no tiene ninguna hoja llamada '${nameOrIndex}'.`);
    return index;
  }

  /**
   * Replaces `{{FIELD}}` placeholders (also `{{ FIELD }}` and `{{FIELD: hint}}`) in every text
   * of the workbook; returns how many
   * @param {Record<string, string | number | boolean>} values
   */
  fillTemplate(values) {
    const strings = {};
    for (const [key, value] of Object.entries(values ?? {})) {
      if (value === null || value === undefined) continue;
      if (!['string', 'number', 'boolean'].includes(typeof value)) {
        throw new TypeError(`El valor de {{${key}}} debe ser texto, número o booleano.`);
      }
      strings[key] = String(value);
    }
    return call(() => this._engine.fill_template_json(JSON.stringify(strings)));
  }

  /**
   * Markers of the standard lexicon (shared with @dimarborda/docx-editor) in the text cells of
   * every worksheet, sheet by sheet and row by row: `field` for `{{KEY}}`, `{{KEY: hint}}` or
   * `{key}`, `pending` for `[ … ]` content still to be written. `offset` / `length` are in
   * code points within the cell's text.
   * @param {{kind?: 'field' | 'pending'}} [filter]
   * @returns {Array<{sheet: number, sheetName: string, ref: string, row: number, col: number, kind: 'field'|'pending', syntax: 'double'|'single'|'bracket', raw: string, key: string|null, hint: string|null, offset: number, length: number}>}
   */
  markers({ kind } = {}) {
    const all = JSON.parse(call(() => this._engine.markers_json()));
    return kind ? all.filter(m => m.kind === kind) : all;
  }

  /**
   * Starts an undo step: the edits made until the next checkpoint undo together.
   * `selection` (any string, e.g. JSON) comes back from `undo()` to restore the selection.
   * @param {string} [selection]
   */
  checkpoint(selection) {
    this._engine.checkpoint(selection ?? undefined);
  }

  /** Where the selection ended up after the current step; `redo()` returns it */
  setSelectionAfter(selection) {
    this._engine.set_selection_after(selection ?? undefined);
  }

  /** Undoes the last step: `{selection, parts}`, or `null` when there is nothing to undo */
  undo() {
    return JSON.parse(this._engine.undo_json());
  }

  /** Redoes the last undone step: `{selection, parts}` or `null` */
  redo() {
    return JSON.parse(this._engine.redo_json());
  }

  get canUndo() {
    return this._engine.can_undo();
  }

  get canRedo() {
    return this._engine.can_redo();
  }

  /**
   * Areas an agent can point at, as the file marks them: tables ("Format as Table"), defined
   * names and print areas, filters, and blocks of adjacent filled cells with their title and
   * header (`kind: 'region'`). Each is `{kind, name, sheet, sheetName, ref, address, title,
   * headerRow, totalsRow, body, columns}`.
   * @param {{sheet?: string | number}} [options] only those of one sheet
   */
  areas({ sheet } = {}) {
    const index = sheet === undefined ? undefined : this.sheet(sheet).index;
    return JSON.parse(call(() => this._engine.areas_json(index)));
  }

  /**
   * What an agent needs to work on an area without reading the whole sheet: its description
   * (as in `areas()`), each column with its type, formulas and number totals, the first
   * `maxRows` rows of data as the cells show them, and the same as a Markdown table.
   * @param {string} target a table or defined name, an address (`'Ventas'!A1:H40`) or, with
   *   `sheet`, a range of that sheet
   * @param {{sheet?: string | number, maxRows?: number, locale?: string}} [options]
   */
  area(target, { sheet, maxRows = 30, locale = 'en' } = {}) {
    const index = sheet === undefined ? undefined : this.sheet(sheet).index;
    if (!Number.isInteger(maxRows) || maxRows < 0) throw new TypeError('maxRows debe ser un entero mayor o igual que 0.');
    return JSON.parse(call(() => this._engine.area_context_json(String(target), index, maxRows, locale)));
  }

  /**
   * Recalculates every formula and stores the new results; returns how many cells changed.
   * Edits already recalculate what depends on them: this is for a workbook just opened whose
   * saved results may be stale. Formulas the engine does not evaluate keep their saved value.
   */
  recalculate() {
    return call(() => this._engine.recalculate());
  }

  /** The workbook as .xlsx bytes */
  toBytes() {
    return call(() => this._engine.export_bytes());
  }

  /** Frees the engine memory of this workbook; it cannot be used afterwards */
  close() {
    this._engine.free();
  }
}

/**
 * Opens an .xlsx
 * @param {Uint8Array | ArrayBuffer} bytes
 */
export async function openWorkbook(bytes) {
  await initEngine();
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  return new Workbook(call(() => new XlsxWorkbook(data)));
}

/** A new workbook with one empty sheet */
export async function createWorkbook() {
  await initEngine();
  return new Workbook(XlsxWorkbook.new_blank());
}
