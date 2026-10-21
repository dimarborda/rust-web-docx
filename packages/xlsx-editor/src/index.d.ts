import type { CellValue, Workbook } from '@dimarborda/xlsx';

export { initEngine } from '@dimarborda/xlsx';

export interface XlsxEditorOptions {
  /** Language the cells are formatted in (`es`, `en-US`…); default: the page's or the browser's */
  locale?: string;
  /** 0.25–4; default 1 */
  zoom?: number;
  /** Sheet tabs at the bottom; default true */
  tabs?: boolean;
  /** Bar above the grid with the selected cell and its content; default true */
  formulaBar?: boolean;
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
}

/** Cells changed from the grid (typing, deleting, pasting, cutting) or by undo and redo */
export interface ChangeDetail {
  /** Name and tab position of the sheet */
  sheet: string;
  index: number;
  /** Range changed, e.g. `B3` or `A1:C4` */
  range: string;
  /** Every cell of `range` with its new value (`null` when empty) */
  cells: { ref: string; value: CellValue }[];
  source: 'edit' | 'clear' | 'paste' | 'cut' | 'undo' | 'redo';
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
  /** Without editing from the grid (selecting and copying still work) */
  readOnly: boolean;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
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

/** `<xlsx-editor src="…" locale zoom sheet tabs formula-bar readonly wasm-url blank quiet>` */
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
