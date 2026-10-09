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

/** `cursor` falls back to the end of the document when there is no caret */
export type InsertPosition = 'cursor' | 'start' | 'end';

export interface NewParagraph {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** Points */
  fontSize?: number;
  align?: Alignment;
}

export interface NewTable {
  /** Cells as text; "\n" is a line break inside the cell. Short rows are padded. */
  rows: string[][];
  /** First row bold, shaded and repeated on every page (default true) */
  header?: boolean;
  /** Relative column widths, e.g. [2, 1, 1]; equal columns when unset */
  widths?: number[];
  /** Per-column alignment */
  align?: Array<'left' | 'center' | 'right'>;
}

/**
 * How text flows around a picture: in line with it, wrapped around its box ('square') or
 * contour ('tight', 'through'; drawn as its box), above and below only, or floating behind /
 * in front of the text (which then ignores it)
 */
export type ImageWrap = 'inline' | 'square' | 'tight' | 'through' | 'topAndBottom' | 'behind' | 'inFront';

/** Side(s) text may take next to a wrapped picture; 'bothSides' uses the wider side */
export type ImageWrapSide = 'bothSides' | 'left' | 'right' | 'largest';

export type ImageHorizontalFrame =
  | 'margin' | 'page' | 'column' | 'character' | 'leftMargin' | 'rightMargin' | 'insideMargin' | 'outsideMargin';
export type ImageVerticalFrame =
  | 'margin' | 'page' | 'paragraph' | 'line' | 'topMargin' | 'bottomMargin' | 'insideMargin' | 'outsideMargin';

/** Position of a floating picture on one axis: an offset from the frame's edge, or an alignment in it */
export interface ImageAxis<Frame extends string, Align extends string> {
  relativeTo?: Frame;
  /** CSS px from the frame's left/top edge (replaces `align`) */
  offset?: number;
  /** Replaces `offset` */
  align?: Align;
}

export type ImageHorizontal = ImageAxis<ImageHorizontalFrame, 'left' | 'center' | 'right' | 'inside' | 'outside'>;
export type ImageVertical = ImageAxis<ImageVerticalFrame, 'top' | 'center' | 'bottom' | 'inside' | 'outside'>;

/** A picture: its paragraph and its index among that paragraph's pictures */
export interface ImageRef {
  paragraph: number;
  index: number;
}

export interface ImageInfo extends ImageRef {
  /** A picture, a text box or a shape without text; all three are edited the same way */
  kind: 'picture' | 'textbox' | 'shape';
  /** CSS px at 100 % zoom */
  width: number;
  height: number;
  wrap: ImageWrap;
  wrapSide: ImageWrapSide | null;
  /** Floating (any wrap but 'inline') */
  anchored: boolean;
  /** null for pictures in line with the text */
  horizontal: { relativeTo: ImageHorizontalFrame; offset: number; align: string | null } | null;
  vertical: { relativeTo: ImageVerticalFrame; offset: number; align: string | null } | null;
  /** Space kept free between the picture and the text */
  distance: { top: number; bottom: number; left: number; right: number };
  alt: string;
  /** Text of a text box, paragraphs separated by "\n" ("" otherwise) */
  text: string;
  /** Outline and fill of a text box or shape ("#RRGGBB"; null for pictures) */
  shape: { geometry: 'rect' | 'roundRect' | 'ellipse' | 'line' | string; fill: string | null; stroke: string | null; strokeWidth: number } | null;
  /** Character offset in the paragraph text where the picture sits */
  textOffset: number;
  /** Where the picture is drawn (page number, page px), or null before it is laid out */
  bounds: { page: number; x: number; y: number; width: number; height: number } | null;
}

export interface ImageChanges {
  width?: number;
  height?: number;
  /** With only width or height, scale the other side (default true) */
  keepRatio?: boolean;
  wrap?: ImageWrap;
  wrapSide?: ImageWrapSide;
  horizontal?: ImageHorizontal;
  vertical?: ImageVertical;
  /** Space between the picture and the text, on every side or per side */
  distance?: number | { top?: number; bottom?: number; left?: number; right?: number };
  alt?: string;
}

export interface InsertImageOptions {
  at?: InsertPosition;
  /** CSS px; with only one of width/height the aspect ratio is kept */
  width?: number;
  height?: number;
  /** Alignment of an inline picture's paragraph */
  align?: 'left' | 'center' | 'right';
  /** Alternative text */
  alt?: string;
  /**
   * 'inline' (default) puts the picture in a paragraph of its own; any other wrap anchors it
   * to the paragraph at the position, which keeps its text and the caret
   */
  wrap?: ImageWrap;
  wrapSide?: ImageWrapSide;
  horizontal?: ImageHorizontal;
  vertical?: ImageVertical;
  distance?: number;
}

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
  /** Start in read-only mode (default false: editable); see `setReadOnly` */
  readOnly?: boolean;
}

export interface DocxEditorEventMap {
  load: CustomEvent<{ fileName: string }>;
  change: CustomEvent<Record<string, never>>;
  selectionchange: CustomEvent<{ selection: Selection | null; format: Format | null }>;
  /** A picture was selected (mouse or `selectImage()`), changed with the mouse, or deselected (null) */
  imageselect: CustomEvent<{ image: ImageInfo | null }>;
  /** `setReadOnly()` changed the mode */
  readonlychange: CustomEvent<{ readOnly: boolean }>;
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
  /** True while the user cannot edit (see `setReadOnly`) */
  readonly readOnly: boolean;

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
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  redo(): void;
  toggleBold(): void;
  toggleItalic(): void;
  toggleUnderline(): void;
  setColor(hex: string, options?: { refocus?: boolean }): void;
  setAlignment(align: Alignment): void;
  /**
   * Inserts paragraphs as one undo step and leaves the caret after them. Existing text before
   * and after the insertion point keeps its own paragraphs.
   */
  insertParagraphs(paragraphs: string | Array<string | NewParagraph>, options?: { at?: InsertPosition }): { first: number; count: number };
  /** Inserts plain text as one undo step; each line becomes a paragraph */
  insertText(text: string, options?: { at?: InsertPosition }): { first: number; count: number };
  /**
   * Inserts a table as one undo step at the caret (or start/end); the caret goes to the
   * paragraph after it. Throws when the position is inside another table.
   */
  insertTable(table: NewTable, options?: { at?: InsertPosition }): { first: number };
  /** @deprecated Appends an empty table at the end; use `insertTable({ rows })` */
  insertTable(rows: number, cols: number, headers?: string[]): { first: number };
  /**
   * Inserts a PNG, JPEG or GIF picture as one undo step; it never exceeds the text width of the
   * page. Inline (default): in its own paragraph, caret after it. Floating (`wrap`): anchored to
   * the paragraph at the position, placed with `horizontal` / `vertical`.
   */
  insertImage(image: Uint8Array | ArrayBuffer | Blob | string, options?: InsertImageOptions): Promise<ImageRef>;

  /** Every picture, text box and shape of the body (table cells included) in document order */
  images(): ImageInfo[];
  /** The picture selected with the mouse or `selectImage()`, or null */
  readonly selectedImage: ImageInfo | null;
  /** Selects a picture as if it had been clicked; null goes back to the text caret */
  selectImage(ref: ImageRef | null): void;
  /**
   * Changes size, wrapping, position, distance to the text or alternative text as one undo
   * step; unset fields keep their value. Returns the picture as it is now.
   * Positions need a floating picture (set `wrap` first or in the same call).
   */
  updateImage(ref: ImageRef, changes: ImageChanges): ImageInfo;
  /** One side keeps the aspect ratio unless `keepRatio` is false; `scale` multiplies the size */
  resizeImage(ref: ImageRef, size: { width?: number; height?: number; keepRatio?: boolean; scale?: number }): ImageInfo;
  /**
   * Moves a floating picture to page coordinates `{x, y}` (top-left corner) or by `{dx, dy}`,
   * keeping its frame when it has an offset. An inline picture first becomes 'square'.
   */
  moveImage(ref: ImageRef, to: { x?: number; y?: number; dx?: number; dy?: number }): ImageInfo;
  setImageWrap(ref: ImageRef, wrap: ImageWrap, options?: { side?: ImageWrapSide; distance?: ImageChanges['distance'] }): ImageInfo;
  /** A floating picture aligns within the margins; an inline one aligns its paragraph */
  alignImage(ref: ImageRef, align: 'left' | 'center' | 'right'): ImageInfo;
  deleteImage(ref: ImageRef): void;
  setBackgroundColor(hex: string): void;
  /** View-only diagonal watermark; pass null to remove it */
  setWatermark(text: string | null, options?: { opacity?: number }): void;

  /** Places a caret (or a selection from `anchor` to `focus`) and scrolls it into view */
  select(anchor: TextPosition, focus?: TextPosition): void;

  setZoom(zoom: number): void;
  /**
   * Read-only mode (editable by default): the user can move the caret, select, copy, zoom and
   * scroll, but not type, delete, paste, cut, format, undo or edit pictures; the toolbar
   * disables its editing buttons. Calls from code still change the document.
   */
  setReadOnly(readOnly?: boolean): void;
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
  | 'left' | 'center' | 'right' | 'both' | 'imageWrap' | 'zoom' | '|';

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
