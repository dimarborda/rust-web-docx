/** Starts the engine; `wasm` is where `xlsx_engine_bg.wasm` is served from, or its bytes */
export function initEngine(wasm?: string | URL | BufferSource): Promise<void>;

/**
 * A cell value. Dates are `Date` objects holding the wall-clock time Excel shows as UTC
 * (read them with `getUTCFullYear()`, `toISOString()`…); `null` is an empty cell.
 */
export type CellValue = string | number | boolean | Date | null;

/** What a cell holds; `error` values are codes such as `#DIV/0!` */
export type CellType = 'empty' | 'number' | 'string' | 'boolean' | 'date' | 'error';

export interface Cell {
  /** Reference, e.g. `B3` */
  ref: string;
  type: CellType;
  /** For formulas, the value cached when the file was last calculated */
  value: CellValue;
  /** Formula without `=`; empty for cells that reuse a shared formula of another cell */
  formula?: string;
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
  setValue(ref: string, value: CellValue): void;
  /**
   * Writes rows of values from `start` rightwards and down, creating rows and cells as needed.
   * `undefined` leaves a cell as it is; `null` clears it. Writing a value replaces a formula.
   */
  setValues(start: string, rows: ReadonlyArray<ReadonlyArray<CellValue | undefined>>): void;
}

/** An open workbook */
export class Workbook {
  private constructor(engine: unknown);
  /** Sheet names, in tab order (chart sheets included) */
  readonly sheetNames: string[];
  /** The worksheets, in tab order (chart sheets left out) */
  readonly sheets: Sheet[];
  /** A worksheet by name (ignoring case) or tab position from 0 */
  sheet(nameOrIndex: string | number): Sheet;
  /**
   * Replaces `{{FIELD}}` placeholders in the shared and inline strings of the workbook and
   * returns how many were replaced. Unknown placeholders are left as they are.
   */
  fillTemplate(values: Record<string, string | number | boolean | null | undefined>): number;
  /** The workbook as .xlsx bytes */
  toBytes(): Uint8Array;
  /** Frees the engine memory of this workbook; it cannot be used afterwards */
  close(): void;
}

/** Opens an .xlsx */
export function openWorkbook(bytes: Uint8Array | ArrayBuffer): Promise<Workbook>;

/** A new workbook with one empty sheet */
export function createWorkbook(): Promise<Workbook>;
