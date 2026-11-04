import type { Area, AreaContext, CellValue, Workbook } from './engine';

export * from './engine';

export interface XlsxEditorOptions {
  /** Language the cells are formatted in (`es`, `en-US`…); default: the page's or the browser's */
  locale?: string;
  /** 0.25–4; default 1 */
  zoom?: number;
  /** Sheet tabs at the bottom; default true */
  tabs?: boolean;
  /** Bar above the grid with the selected cell and its content; default true */
  formulaBar?: boolean;
  /** Bar below the tabs with the count, sum and average of the selection; default true */
  statusBar?: boolean;
  /** A viewer: cells are selected and copied but not edited; default false */
  readOnly?: boolean;
  /** Where `xlsx_engine_bg.wasm` is served from, when the bundler does not serve it */
  wasmUrl?: string | URL;
}

export interface CellPosition {
  row: number;
  col: number;
}

export interface Selection {
  /** Active cell, e.g. `B3` */
  ref: string;
  /** Selected range, e.g. `B3:D7` (or the cell alone) */
  range: string;
  active: CellPosition;
  /** Text the active cell shows */
  text: string;
  /** Whole rows, whole columns or the whole sheet selected from the headers */
  whole: 'row' | 'col' | 'all' | null;
}

/** What the status bar says about the selection */
export interface SelectionStats {
  /** Cells that have something */
  count: number;
  /** Of them, numbers (dates count, as serial numbers) */
  numbers: number;
  /** `null` without numbers, or when an error is selected */
  sum: number | null;
  average: number | null;
  min: number | null;
  max: number | null;
}

export interface FindOptions {
  matchCase?: boolean;
  /** Only cells whose whole text matches */
  wholeCell?: boolean;
  /** The active sheet (default) or every worksheet */
  scope?: 'sheet' | 'workbook';
  /** Look in the formula of formula cells instead of the text they show */
  inFormulas?: boolean;
}

export interface FindHit {
  sheet: number;
  sheetName: string;
  ref: string;
  row: number;
  col: number;
  /** The text that matched */
  text: string;
}

/** Cells changed from the grid (typing, deleting, pasting, cutting) or by undo and redo */
export interface ChangeDetail {
  /** Name and tab position of the sheet */
  sheet: string;
  index: number;
  /** Range changed, e.g. `B3` or `A1:C4` */
  range: string;
  /** Every cell of `range` with its new value (`null` when empty); empty above 200 000 cells */
  cells: { ref: string; value: CellValue }[];
  source: 'edit' | 'clear' | 'paste' | 'cut' | 'move' | 'insert' | 'delete' | 'fill' | 'replace' | 'undo' | 'redo';
}

export interface XlsxEditorEventMap {
  load: CustomEvent<{ sheets: string[] }>;
  sheetchange: CustomEvent<{ index: number; name: string }>;
  selectionchange: CustomEvent<Selection>;
  change: CustomEvent<ChangeDetail>;
  /** A change the engine refused (e.g. overwriting the source of a shared formula) */
  error: CustomEvent<{ error: Error }>;
}

/** Excel workbook editor drawn on canvas */
export class XlsxEditor extends EventTarget {
  /** Creates an editor inside `container` (which should have a height) */
  static create(container: HTMLElement, options?: XlsxEditorOptions): Promise<XlsxEditor>;
  readonly locale: string;
  readonly zoom: number;
  /** The open workbook, to read or change cells from code */
  readonly workbook: Workbook | null;
  /** Position of the active sheet among the workbook's tabs */
  readonly sheetIndex: number;
  readonly sheetName: string | null;
  readonly selection: Selection | null;
  /** Count, sum, average, smallest and largest of the selected cells; `null` when they are empty */
  readonly selectionStats: SelectionStats | null;
  /** Without editing from the grid (selecting and copying still work) */
  readOnly: boolean;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  /**
   * Moves `count` whole rows from `first` before the row `before` (numbered as they are now), as
   * one undo step, and selects them. False when nothing changed: it stays where it is, it is
   * read-only, or the engine refused (an `error` event says why).
   */
  moveRows(first: number, count: number, before: number): boolean;
  /** Moves whole columns, by number (see `moveRows`) */
  moveColumns(first: number, count: number, before: number): boolean;
  /** Inserts rows before the row `before`, as one undo step, and selects them; false if refused */
  insertRows(before: number, count?: number): boolean;
  /** Inserts columns, by number (see `insertRows`) */
  insertColumns(before: number, count?: number): boolean;
  /** Deletes rows from `first`, as one undo step; false if refused */
  deleteRows(first: number, count?: number): boolean;
  /** Deletes columns, by number (see `deleteRows`) */
  deleteColumns(first: number, count?: number): boolean;
  /** Hides `count` rows from `first`, as one undo step */
  hideRows(first: number, count?: number): boolean;
  /** Shows hidden rows among the `count` from `first` */
  showRows(first: number, count?: number): boolean;
  /** Hides columns, by number */
  hideColumns(first: number, count?: number): boolean;
  /** Shows hidden columns, by number */
  showColumns(first: number, count?: number): boolean;
  /**
   * Fills `target` (`A1:A20`) from the selected cells, which must be at one of its ends, as one
   * undo step, and selects it. False if refused.
   */
  fill(target: string, options?: { series?: boolean }): boolean;
  /** Where a text is, row by row: `*` is any text, `?` one character and `~` takes them literally */
  find(query: string, options?: FindOptions): FindHit[];
  /**
   * Replaces a text in cells with text and in formulas (never in numbers or dates), as one undo
   * step; returns how many cells changed. Formatted runs of a rich text become plain.
   */
  replaceAll(query: string, replacement: string, options?: Omit<FindOptions, 'inFormulas'>): number;
  /** Opens the search panel, with the row to replace when asked */
  openFind(options?: { replace?: boolean }): void;
  closeFind(): void;
  /** Undoes the last change; false when there is nothing to undo */
  undo(): boolean;
  /** Redoes the last undone change */
  redo(): boolean;
  /** Opens a workbook from a URL, a File/Blob or its bytes */
  open(source: string | URL | Blob | ArrayBuffer | Uint8Array): Promise<void>;
  /** Starts a new, empty workbook */
  createBlank(): Promise<void>;
  /** Shows a sheet by name or tab position */
  setSheet(nameOrIndex: string | number): void;
  setZoom(zoom: number): void;
  /** Selects a cell or range (`B3`, `A1:C4`) and scrolls it into view */
  select(ref: string): void;
  /**
   * Selects an area and shows its sheet: a table or defined name, an address with its sheet
   * (`Ventas!A1:H40`) or a range of the active sheet
   */
  selectArea(target: string): Area;
  /**
   * Context of the selection for an agent: the selected range, or with one cell selected the
   * region around it (`expand: false` keeps the cell)
   */
  selectionContext(options?: { maxRows?: number; expand?: boolean }): AreaContext | null;
  /** Draws again (after changing cells through `workbook`, call `reload()` instead) */
  render(): void;
  /** Reads the active sheet again from the workbook, after changing it from code */
  reload(): void;
  /** The workbook as .xlsx bytes */
  toBytes(): Uint8Array;
  /** Saves the workbook as a file */
  download(filename?: string): void;
  destroy(): void;
  addEventListener<K extends keyof XlsxEditorEventMap>(type: K, listener: (event: XlsxEditorEventMap[K]) => void, options?: boolean | AddEventListenerOptions): void;
  addEventListener(type: string, listener: EventListenerOrEventListenerObject | null, options?: boolean | AddEventListenerOptions): void;
}

/** `<xlsx-editor src="…" locale zoom sheet tabs formula-bar status-bar readonly wasm-url blank quiet>` */
export class XlsxEditorElement extends HTMLElement {
  /** The editor, once the `ready` event fired */
  editor: XlsxEditor | null;
}

/** What a typed text means in the language `locale`, as the editor decides it (see the README) */
export function parseInput(text: string, locale?: string, options?: { today?: Date }): {
  type: 'empty' | 'string' | 'number' | 'percent' | 'boolean' | 'date' | 'time' | 'formula';
  value: CellValue;
};

/** Registers `<xlsx-editor>` (done on import where `customElements` exists) */
export function defineXlsxElements(): void;

declare global {
  interface HTMLElementTagNameMap {
    'xlsx-editor': XlsxEditorElement;
  }
}
