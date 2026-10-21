/** Starts the engine; `wasm` is where `xlsx_engine_bg.wasm` is served from, or its bytes */
export function initEngine(wasm?: string | URL | BufferSource): Promise<void>;

/**
 * A cell value. Dates are `Date` objects holding the wall-clock time Excel shows as UTC
 * (read them with `getUTCFullYear()`, `toISOString()`…); `null` is an empty cell.
 */
export type CellValue = string | number | boolean | Date | null;

/**
 * A formula to write, in Excel's English syntax (`SUM`, `IF`, `,` between arguments; `;` is
 * also accepted); the leading `=` is optional
 */
export interface FormulaInput {
  formula: string;
}

/** What a cell holds; `error` values are codes such as `#DIV/0!` */
export type CellType = 'empty' | 'number' | 'string' | 'boolean' | 'date' | 'error';

export interface Cell {
  /** Reference, e.g. `B3` */
  ref: string;
  type: CellType;
  /**
   * For formulas, the result: recalculated after every edit, or the value saved in the file
   * for formulas the engine does not evaluate
   */
  value: CellValue;
  /** Formula without `=` (cells of a shared formula get their own, shifted) */
  formula?: string;
}

/** What `undo()` or `redo()` changed */
export interface HistoryMove {
  /** The selection given to `checkpoint()` (undo) or `setSelectionAfter()` (redo) */
  selection: string | null;
  /** Package parts that changed */
  parts: string[];
}

/** Where an area comes from in the file */
export type AreaKind = 'table' | 'name' | 'printArea' | 'filter' | 'region';

/** A range that holds something as a unit */
export interface Area {
  kind: AreaKind;
  /** Table or defined name; null for filters, print areas and regions */
  name: string | null;
  sheet: number;
  sheetName: string;
  /** Range, e.g. `A1:H40` */
  ref: string;
  /** With the sheet, e.g. `Ventas!A1:H40`: usable as `target` anywhere */
  address: string;
  /** Text of the title rows above the data (regions) */
  title: string | null;
  headerRow: number | null;
  totalsRow: number | null;
  /** Rows of data, below the header and above the totals */
  body: string | null;
  /** Header texts or table column names, one per column */
  columns: string[];
}

export interface AreaColumn {
  /** Column letters */
  col: string;
  name: string;
  type: 'number' | 'text' | 'date' | 'boolean' | 'error' | 'mixed' | 'empty';
  /** Cells with a value */
  filled: number;
  numbers?: { count: number; sum: number; min: number; max: number; average: number };
  /** Cells with a formula, and the first one as `D4: =B4*C4` */
  formulas?: number;
  exampleFormula?: string;
}

/** Context of an area for an agent */
export interface AreaContext extends Area {
  /** Rows of data in the area */
  rowCount: number;
  rowsShown: number;
  truncated: boolean;
  columnsInfo: AreaColumn[];
  /** The first rows of data as the cells show them (formatted text) */
  rows: string[][];
  /** Row number of `rows[0]` */
  firstRow: number | null;
  /** Totals row as shown, for tables with one */
  totals: string[] | null;
  /** Merged cells that overlap the area */
  merges: string[];
  /** Header, rows and totals as a Markdown table, with row numbers and column letters */
  markdown: string;
}

/** A border: a style (`'thin'`), null to remove borders, or the full form */
export type BorderInput =
  | 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair'
  | null
  | { style: string | null; color?: string; edges?: 'all' | 'outline' | 'top' | 'bottom' | 'left' | 'right' | 'inside' };

/** What to change in the look of cells; what is not given stays as it is */
export interface CellFormat {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Font color, `#RRGGBB` */
  color?: string;
  /** Points */
  fontSize?: number;
  fontName?: string;
  /** Background `#RRGGBB`; null removes it */
  fill?: string | null;
  /** Excel format code: `#,##0.00`, `0%`, `dd/mm/yyyy`, `"$"#,##0`… */
  numberFormat?: string;
  align?: 'general' | 'left' | 'center' | 'right' | 'justify' | 'fill' | 'centerContinuous' | 'distributed';
  valign?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
  wrap?: boolean;
  border?: BorderInput;
}

/** One worksheet */
export class Sheet {
  private constructor();
  readonly workbook: Workbook;
  /** Position of the tab, from 0 */
  readonly index: number;
  readonly name: string;
  /** The cell at `ref` (`B3`) */
  getCell(ref: string): Cell;
  /** The value of the cell at `ref` */
  getValue(ref: string): CellValue;
  /** The values of a range (`A1:C3`), by row */
  getValues(range: string): CellValue[][];
  /** Writes one value; `null` clears the cell and keeps its format */
  setValue(ref: string, value: CellValue | FormulaInput): void;
  /**
   * Writes rows of values from `start` rightwards and down, creating rows and cells as needed.
   * `undefined` leaves a cell as it is; `null` clears it. Writing a value replaces a formula.
   * The formulas that depend on the cells written are recalculated.
   */
  setValues(start: string, rows: ReadonlyArray<ReadonlyArray<CellValue | FormulaInput | undefined>>): void;
  /** Drawing API: the grid of the sheet */
  layout(): SheetLayout;
  /**
   * Drawing API: every cell that shows something, with the text Excel shows for it in the
   * language `locale` (`es`, `en-US`…; default `en`)
   */
  cellViews(locale?: string): CellViewRow[];
  /** Areas of this sheet */
  areas(): Area[];
  /** The region around a cell as Excel's Ctrl+A selects it (or its table, or the cell alone) */
  regionAt(ref: string): Area;
  /** Context of a range of this sheet or of a named area */
  area(target: string, options?: { maxRows?: number; locale?: string }): AreaContext;
  /** Formats the cells of a range; what is not given stays as it is */
  setFormat(range: string, format: CellFormat): void;
  /** Back to the default format (values stay) */
  clearFormat(range: string): void;
  /** Width of a column or columns (`'B'`, `'B:D'`, `2`) in characters */
  setColumnWidth(columns: string | number, width: number): void;
  /** Merges a range into one cell; only its first cell may hold a value */
  mergeCells(range: string): void;
  /** Splits the merged cells that overlap a range; returns how many */
  unmergeCells(range: string): number;
}

/** Grid of a sheet, as Excel stores it */
export interface SheetLayout {
  /** Points */
  defaultRowHeight: number | null;
  /** Characters of the default font's digit width, padding included */
  defaultColWidth: number | null;
  baseColWidth: number | null;
  /** `[firstColumn, lastColumn, width | null, hidden]`, columns counted from 1 */
  cols: [number, number, number | null, boolean][];
  /** `[row, heightInPoints | null, hidden]`, rows counted from 1 */
  rows: [number, number | null, boolean][];
  /** Merged ranges such as `A1:D1` */
  merges: string[];
  gridlines: boolean;
  headings: boolean;
  rightToLeft: boolean;
  frozen: { rows: number; cols: number } | null;
}

/**
 * `[row, col, style, text, kind, color?]`: `kind` is `n` (number or date), `s` (text), `b`,
 * `e` (error) or `''` (empty cell with a format); `color` comes from the number format
 */
export type CellViewRow = [number, number, number, string, 'n' | 's' | 'b' | 'e' | '', string?];

export interface FontView {
  name: string;
  /** Points */
  size: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  /** `#RRGGBB`; `null` is automatic (black) */
  color: string | null;
}

/** A cell format resolved for drawing */
export interface XfView {
  font: FontView;
  /** Background `#RRGGBB` */
  fill: string | null;
  /** Left, right, top, bottom: `[style, color]` or `null` */
  border: ([string, string | null] | null)[];
  h: string | null;
  v: string | null;
  wrap: boolean;
  shrink: boolean;
  indent: number;
  rotation: number;
}

export interface StyleView {
  defaultFont: { name: string; size: number };
  /** By style index (a cell's `style`) */
  xfs: XfView[];
}

/** An open workbook */
export class Workbook {
  private constructor(engine: unknown);
  /** Sheet names, in tab order (chart sheets included) */
  readonly sheetNames: string[];
  /** The worksheets, in tab order (chart sheets left out) */
  readonly sheets: Sheet[];
  /** Every tab in order, chart sheets included */
  readonly tabs: { name: string; kind: 'worksheet' | 'chartsheet' | 'other'; state: 'visible' | 'hidden' | 'veryHidden' }[];
  /** Drawing API: the cell formats resolved to fonts, fills, borders and alignment */
  styleView(): StyleView;
  /** A worksheet by name (ignoring case) or tab position from 0 */
  sheet(nameOrIndex: string | number): Sheet;
  /**
   * Replaces `{{FIELD}}` placeholders (also `{{ FIELD }}` and `{{FIELD: hint}}`) in the shared and inline strings of the workbook and
   * returns how many were replaced. Unknown placeholders are left as they are.
   */
  fillTemplate(values: Record<string, string | number | boolean | null | undefined>): number;
  /**
   * Markers of the standard lexicon (shared with @dimarborda/docx-editor) in the text cells of
   * every worksheet: `field` for `{{KEY}}`, `{{KEY: hint}}` or `{key}`, `pending` for `[ … ]`
   * content still to be written. `offset` / `length` are in code points within the cell text.
   */
  markers(filter?: { kind?: 'field' | 'pending' }): WorkbookMarker[];
  /**
   * Starts an undo step: the edits made until the next checkpoint undo together. `selection`
   * (any string, e.g. JSON) comes back from `undo()` so the caller can restore its selection.
   */
  checkpoint(selection?: string): void;
  /** Where the selection ended up after the current step; `redo()` returns it */
  setSelectionAfter(selection?: string): void;
  /** Undoes the last step; `null` when there is nothing to undo */
  undo(): HistoryMove | null;
  /** Redoes the last undone step; `null` when there is nothing to redo */
  redo(): HistoryMove | null;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  /**
   * Areas as the file marks them (tables, defined names, print areas, filters) and blocks of
   * adjacent filled cells with their title and header
   */
  areas(options?: { sheet?: string | number }): Area[];
  /**
   * Context of an area for an agent: `target` is a table or defined name, an address
   * (`'Ventas'!A1:H40`) or, with `sheet`, a range of that sheet. `maxRows` (default 30) limits
   * the rows shown.
   */
  area(target: string, options?: { sheet?: string | number; maxRows?: number; locale?: string }): AreaContext;
  /**
   * Recalculates every formula and stores the results that changed; returns how many cells
   * changed. Formulas the engine does not evaluate keep the value saved in the file.
   */
  recalculate(): number;
  /** The workbook as .xlsx bytes */
  toBytes(): Uint8Array;
  /** Frees the engine memory of this workbook; it cannot be used afterwards */
  close(): void;
}

/** Opens an .xlsx */
export function openWorkbook(bytes: Uint8Array | ArrayBuffer): Promise<Workbook>;

/** A new workbook with one empty sheet */
export function createWorkbook(): Promise<Workbook>;

/** A marker found in a text cell by `Workbook.markers()` */
export interface WorkbookMarker {
  sheet: number;
  sheetName: string;
  /** A1 reference of the cell */
  ref: string;
  /** Row and column, counted from 1 */
  row: number;
  col: number;
  kind: 'field' | 'pending';
  syntax: 'double' | 'single' | 'bracket';
  /** The marker as written */
  raw: string;
  /** Field name; null for pending content */
  key: string | null;
  /** Text after the field's colon, or the text inside `[ … ]`; null when there is none */
  hint: string | null;
  offset: number;
  length: number;
}
