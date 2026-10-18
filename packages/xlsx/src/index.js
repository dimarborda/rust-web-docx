// Excel (.xlsx) workbooks backed by the Rust/WebAssembly engine. In Node the engine is read
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
  throw new TypeError(`Valor no admitido en ${where}: usa texto, número, booleano, Date o null.`);
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

  /** The values of a range (`A1:C3`), by row */
  getValues(range) {
    return JSON.parse(call(() => this.workbook._engine.range_json(this.index, range))).map(row => row.map(c => toCell(c).value));
  }

  /** Writes one value; `null` clears the cell (its format stays) */
  setValue(ref, value) {
    this.setValues(ref, [[value]]);
  }

  /**
   * Writes rows of values from the cell `start` rightwards and down. `undefined` (or a hole)
   * leaves a cell as it is; `null` clears it.
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
   * language `locale`, as `[row, col, style, text, kind, color?]` (`kind`: `n`, `s`, `b`, `e`
   * or `''` for an empty cell with a format; `style` indexes `workbook.styleView().xfs`)
   * @param {string} [locale]
   */
  cellViews(locale = 'en') {
    return JSON.parse(call(() => this.workbook._engine.sheet_cells_json(this.index, locale)));
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
   * Replaces `{{FIELD}}` placeholders in every text of the workbook; returns how many
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
