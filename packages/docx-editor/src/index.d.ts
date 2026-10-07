// Type definitions for @dimarborda/docx-editor

export type Alignment = 'left' | 'center' | 'right' | 'both';

export interface TextPosition {
  /** Paragraph index in document order (body and table cells) */
  paragraph: number;
  /** Offset in Unicode code points within the paragraph text */
  offset: number;
}

export interface Selection {
  anchor: TextPosition;
  focus: TextPosition;
  start: TextPosition;
  end: TextPosition;
  collapsed: boolean;
}

export interface Format {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  /** "#RRGGBB" or null when the text uses the automatic color */
  color: string | null;
  align: Alignment;
}

export interface ReplaceResult {
  occurrences_replaced: number;
  affected_files: string[];
  message: string;
}

export interface DocumentStats {
  word_count: number;
  char_count: number;
  paragraph_count: number;
  table_count: number;
  page_count: number;
  page_setup?: { orientation: 'portrait' | 'landscape'; [key: string]: unknown };
  [key: string]: unknown;
}

export interface DocumentFont {
  name: string;
  /** True when a metric-compatible substitute is drawn instead of the real font */
  substitute: boolean;
}

export interface DocxEditorLabels {
  editor: string;
  page: (n: number, total: number) => string;
  nothingToUndo: string;
  nothingToRedo: string;
  clickToFormat: string;
  /** Default file name for `openBlank()` */
  untitled: string;
}

export type PageSize = 'a4' | 'letter' | 'legal';

export interface DocxEditorOptions {
  /** 1 = 100 % (default) */
  zoom?: number;
  /** Dashed guides where table cells have no border (default true) */
  gridlines?: boolean;
  /** "Page 1 of 3" above each page (default true) */
  pageLabels?: boolean;
  /** Default 'es' */
  locale?: 'es' | 'en';
  labels?: Partial<DocxEditorLabels>;
  /** Where docx_engine_bg.wasm is served from, when the bundler does not handle it */
  wasmUrl?: string | URL;
}

export interface DocxEditorEventMap {
  load: CustomEvent<{ fileName: string }>;
  change: CustomEvent<Record<string, never>>;
  selectionchange: CustomEvent<{ selection: Selection | null; format: Format | null }>;
  message: CustomEvent<{ message: string; error: boolean }>;
}

export class DocxEditor extends EventTarget {
  /** Loads the engine if needed and mounts an editor in `container` */
  static create(container: HTMLElement, options?: DocxEditorOptions): Promise<DocxEditor>;
  /** Requires the engine to be initialized already (see initEngine) */
  constructor(container: HTMLElement, options?: DocxEditorOptions);

  /** The element passed to `create()`; the editor never changes it */
  readonly container: HTMLElement;
  /** The editor's own element (class `docx-editor`), created inside `container` and removed by `destroy()` */
  readonly root: HTMLElement;
  readonly options: Required<Omit<DocxEditorOptions, 'labels' | 'wasmUrl'>> & DocxEditorOptions;
  readonly labels: DocxEditorLabels;
  readonly hasDocument: boolean;
  readonly fileName: string;
  readonly zoom: number;
  readonly selection: Selection | null;

  open(source: Uint8Array | ArrayBuffer | Blob, options?: { fileName?: string }): Promise<void>;
  openSample(): void;
  /** New empty document (A4 by default) with the caret at its start */
  openBlank(options?: { fileName?: string; pageSize?: PageSize; focus?: boolean }): void;
  close(): void;
  /** The edited document as .docx bytes */
  save(): Uint8Array;
  saveBlob(): Blob;
  /** Triggers a browser download */
  download(fileName?: string): void;

  text(): string;
  /** Placeholders such as "{{CLIENTE}}" or "{fecha}" */
  variables(): string[];
  /** Keys with or without braces: { CLIENTE: 'Acme' } fills {{CLIENTE}} */
  replaceVariables(values: Record<string, string> | { key: string; value: string }[]): ReplaceResult;
  findReplace(search: string, replacement: string, options?: { matchCase?: boolean; regex?: boolean }): ReplaceResult;
  stats(): DocumentStats | null;
  fonts(): DocumentFont[];

  undo(): void;
  redo(): void;
  toggleBold(): void;
  toggleItalic(): void;
  toggleUnderline(): void;
  setColor(hex: string, options?: { refocus?: boolean }): void;
  setAlignment(align: Alignment): void;
  insertTable(rows: number, cols: number, headers?: string[]): void;
  setBackgroundColor(hex: string): void;
  /** View-only diagonal watermark; pass null to remove it */
  setWatermark(text: string | null, options?: { opacity?: number }): void;

  /** Places a caret (or a selection from `anchor` to `focus`) and scrolls it into view */
  select(anchor: TextPosition, focus?: TextPosition): void;

  setZoom(zoom: number): void;
  /** Focuses the editor; when there is no caret yet it goes to the start of the document */
  focus(): void;
  destroy(): void;

  addEventListener<K extends keyof DocxEditorEventMap>(
    type: K, listener: (event: DocxEditorEventMap[K]) => void, options?: boolean | AddEventListenerOptions): void;
  addEventListener(type: string, listener: EventListenerOrEventListenerObject | null, options?: boolean | AddEventListenerOptions): void;
  removeEventListener<K extends keyof DocxEditorEventMap>(
    type: K, listener: (event: DocxEditorEventMap[K]) => void, options?: boolean | EventListenerOptions): void;
  removeEventListener(type: string, listener: EventListenerOrEventListenerObject | null, options?: boolean | EventListenerOptions): void;
}

export type ToolbarItem =
  | 'undo' | 'redo' | 'bold' | 'italic' | 'underline' | 'color'
  | 'left' | 'center' | 'right' | 'both' | 'zoom' | '|';

export const DEFAULT_TOOLBAR_ITEMS: ToolbarItem[];

export function createDocxToolbar(
  editor: DocxEditor,
  container: HTMLElement,
  options?: { items?: ToolbarItem[]; locale?: 'es' | 'en'; labels?: Record<string, string> },
): { destroy(): void };

export class DocxEditorElement extends HTMLElement {
  /** The DocxEditor, null until the `ready` event */
  editor: DocxEditor | null;
  whenReady(): Promise<DocxEditor>;
}

export class DocxToolbarElement extends HTMLElement {}

/** Registers <docx-editor> and <docx-toolbar>; importing the package already does it */
export function defineDocxElements(): void;

export function initEngine(wasmUrl?: string | URL): Promise<void>;
/** Bytes of the built-in demo contract */
export function sampleDocx(): Promise<Uint8Array>;

export interface StoredFont {
  id: string;
  family: string;
  fileName: string;
  size: number;
  weight: string;
  style: string;
  buffer: ArrayBuffer;
  addedAt: number;
}

export function registerFont(file: File | { name: string; buffer: ArrayBuffer }, options?: { persist?: boolean }): Promise<StoredFont>;
export function loadStoredFonts(): Promise<StoredFont[]>;
export function listStoredFonts(): Promise<StoredFont[]>;
export function removeStoredFont(id: string): Promise<void>;

declare global {
  interface HTMLElementTagNameMap {
    'docx-editor': DocxEditorElement;
    'docx-toolbar': DocxToolbarElement;
  }
}
