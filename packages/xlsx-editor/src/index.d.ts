import type { Workbook } from '@dimarborda/xlsx';

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

export interface XlsxEditorEventMap {
  load: CustomEvent<{ sheets: string[] }>;
  sheetchange: CustomEvent<{ index: number; name: string }>;
  selectionchange: CustomEvent<Selection>;
}

/** Excel workbook viewer drawn on canvas */
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

/** `<xlsx-editor src="…" locale zoom sheet tabs formula-bar wasm-url blank>` */
export class XlsxEditorElement extends HTMLElement {
  /** The editor, once the `ready` event fired */
  editor: XlsxEditor | null;
}

/** Registers `<xlsx-editor>` (done on import where `customElements` exists) */
export function defineXlsxElements(): void;

declare global {
  interface HTMLElementTagNameMap {
    'xlsx-editor': XlsxEditorElement;
  }
}
