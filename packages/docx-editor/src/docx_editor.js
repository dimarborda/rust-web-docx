// DocxEditor: a Word document editor mounted in any element. Rust (WebAssembly) parses the
// .docx, lays out the pages and applies lossless edits; this class renders the pages to
// canvases, wires the caret/selection editor and exposes a small API plus DOM events.
import { DocxSession, engineReady, initEngine } from './engine.js';
import { createCanvasEditor } from './canvas_editor.js';
import { onFontsChanged } from './fonts_registry.js';
import { disableLigatures, drawPageItems, fontWidthsSignature, layoutFontFaces, measureText, usesSubstitute } from './render.js';

const LABELS = {
  es: {
    editor: 'Editor del documento',
    page: (n, total) => `Página ${n} de ${total}`,
    nothingToUndo: 'No hay nada para deshacer.',
    nothingToRedo: 'No hay nada para rehacer.',
    clickToFormat: 'Haz clic en el texto para aplicar formato.',
    untitled: 'documento.docx',
  },
  en: {
    editor: 'Document editor',
    page: (n, total) => `Page ${n} of ${total}`,
    nothingToUndo: 'Nothing to undo.',
    nothingToRedo: 'Nothing to redo.',
    clickToFormat: 'Click in the text to apply formatting.',
    untitled: 'document.docx',
  },
};

// {{VARIABLE}} placeholders, plus VEAD-style {variable}
const VARIABLE_PATTERN = /\{\{[^{}]+\}\}|\{[^{}\s]+\}/g;
const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/**
 * @typedef {object} DocxEditorOptions
 * @property {number} [zoom=1]            1 = 100 %
 * @property {boolean} [gridlines=true]   dashed guides where table cells have no border
 * @property {boolean} [pageLabels=true]  "Page 1 of 3" above each page
 * @property {'es'|'en'} [locale='es']
 * @property {object} [labels]            overrides for the locale's texts
 * @property {string|URL} [wasmUrl]       where docx_engine_bg.wasm is served from
 *
 * Events (dispatched on the editor; `<docx-editor>` re-dispatches them on the element):
 * - `load`             a document was opened
 * - `change`           the document changed (typing, formatting, replacements…)
 * - `selectionchange`  detail: { selection, format } with format = { bold, italic, underline, color, align }
 * - `imageselect`      detail: { image } the picture selected with the mouse or `selectImage()` (null when none)
 * - `message`          detail: { message, error } hints and errors meant for the user
 */
export class DocxEditor extends EventTarget {
  /** Loads the engine if needed and mounts an editor in `container` */
  static async create(container, options = {}) {
    await initEngine(options.wasmUrl);
    return new DocxEditor(container, options);
  }

  #session = null;
  #fileName = 'documento.docx';
  #elements = [];
  #layout = null;
  #rendered = { session: null, zoom: null, cards: new Map() };
  #paragraphCache = { source: null, list: [], byIndex: new Map() };
  #requestedFaces = new Set();
  #watermark = null;
  #zoom;
  #canvasEditor;
  #pages;
  #unsubscribeFonts;
  #onFontsLoaded;
  #fontsRelayout = 0;
  /** Probe widths of the fonts the last layout was measured with */
  #fontWidths = '';
  #fontChecks = [];
  #onKeyDown;

  /** Use `DocxEditor.create()` unless the engine is already initialized (`initEngine()`) */
  constructor(container, options = {}) {
    super();
    if (!engineReady()) throw new Error('docx-editor: await initEngine() or use DocxEditor.create()');
    this.options = { zoom: 1, gridlines: true, pageLabels: true, locale: 'es', ...options };
    this.labels = { ...(LABELS[this.options.locale] || LABELS.en), ...options.labels };
    this.#zoom = this.options.zoom;

    // The editor lives in its own element inside `container` and only ever touches that
    // element: several editors (or a destroyed one and its replacement, as with React's
    // StrictMode) can share a container without breaking each other.
    this.container = container;
    this.root = document.createElement('div');
    this.root.className = 'docx-editor';
    this.#pages = document.createElement('div');
    this.#pages.className = 'docx-editor-pages';
    this.root.appendChild(this.#pages);
    container.appendChild(this.root);

    this.#canvasEditor = createCanvasEditor({
      root: this.root,
      labels: this.labels,
      session: () => this.#session,
      paragraphs: () => this.#paragraphs(),
      zoom: () => this.#zoom,
      measure: measureText,
      documentChanged: () => this.#changed(),
      selectionChanged: selection => this.#emit('selectionchange', { selection, format: this.#formatAt(selection) }),
      hint: message => this.#emit('message', { message, error: false }),
      imageSelected: ref => this.#emit('imageselect', { image: ref ? this.#imageInfo(ref) : null }),
      imageEdited: (ref, change) => this.#userImageEdit(ref, change),
      imageDeleted: ref => this.#userImageEdit(ref, null),
    });

    // Formatting shortcuts while the caret is in this editor
    this.#onKeyDown = e => {
      if (!(e.metaKey || e.ctrlKey) || e.altKey || !this.#session) return;
      const action = { b: 'toggleBold', i: 'toggleItalic', u: 'toggleUnderline' }[e.key.toLowerCase()];
      if (action && !e.shiftKey) {
        e.preventDefault();
        this[action]();
      }
    };
    this.root.addEventListener('keydown', this.#onKeyDown);

    // Uploaded fonts change text widths: measure and lay out again
    this.#unsubscribeFonts = onFontsChanged(() => this.#relayoutForFonts());

    // Any web font that finishes loading (also those a canvas requested on its own, which
    // Safari does not report to document.fonts.load) may change text widths
    this.#onFontsLoaded = () => this.#relayoutForFonts();
    document.fonts?.addEventListener?.('loadingdone', this.#onFontsLoaded);
  }

  // ---------- Documents ----------

  /**
   * Opens a .docx.
   * @param {Uint8Array | ArrayBuffer | Blob} source
   * @param {{fileName?: string}} [options]
   */
  async open(source, { fileName } = {}) {
    const bytes = source instanceof Uint8Array ? source
      : source instanceof ArrayBuffer ? new Uint8Array(source)
      : new Uint8Array(await source.arrayBuffer());
    this.#replaceSession(new DocxSession(bytes), fileName || source.name || 'documento.docx');
  }

  /** Opens the built-in demo contract */
  openSample() {
    this.#replaceSession(DocxSession.new_sample(), 'contrato_ejemplo.docx');
  }

  /**
   * Opens a new empty document (one empty paragraph, Calibri 11) and puts the caret in it.
   * @param {{fileName?: string, pageSize?: 'a4'|'letter'|'legal', focus?: boolean}} [options]
   */
  openBlank({ fileName = this.labels.untitled, pageSize = 'a4', focus = true } = {}) {
    this.#replaceSession(DocxSession.new_blank(pageSize), fileName);
    this.#canvasEditor.select({ paragraph: 0, offset: 0 });
    if (focus) this.focus();
  }

  /** Closes the current document */
  close() {
    this.#session?.free();
    this.#session = null;
    this.#layout = null;
    this.#elements = [];
    this.#canvasEditor.clear();
    this.#pages.replaceChildren();
    this.#rendered = { session: null, zoom: null, cards: new Map() };
  }

  get hasDocument() {
    return !!this.#session;
  }

  get fileName() {
    return this.#fileName;
  }

  /** The edited document as .docx bytes */
  save() {
    this.#require();
    return this.#session.export_bytes();
  }

  saveBlob() {
    return new Blob([this.save()], { type: DOCX_MIME });
  }

  /** Downloads the edited document (browsers; in Tauri write `save()` with the fs plugin) */
  download(fileName = this.#fileName) {
    const url = URL.createObjectURL(this.saveBlob());
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  // ---------- Content ----------

  /** Plain text of the document, paragraphs separated by line breaks */
  text() {
    this.#require();
    return this.#session.get_raw_text();
  }

  /** Placeholders found in the document, e.g. ["{{CLIENTE}}", "{fecha}"] */
  variables() {
    if (!this.#session) return [];
    return [...new Set(this.#session.get_raw_text().match(VARIABLE_PATTERN) || [])];
  }

  /**
   * Fills placeholders, keeping each one's formatting even when Word split it across runs.
   * Keys may be written with or without braces: `{ CLIENTE: 'Acme' }` fills `{{CLIENTE}}`.
   * @param {Record<string, string> | {key: string, value: string}[]} values
   * @returns {{occurrences_replaced: number, message: string}}
   */
  replaceVariables(values) {
    this.#require();
    const pairs = (Array.isArray(values) ? values : Object.entries(values).map(([key, value]) => ({ key, value })))
      .filter(p => p.key)
      .map(({ key, value }) => ({ key: key.includes('{') ? key : `{{${key}}}`, value: String(value ?? '') }));
    const result = JSON.parse(this.#session.batch_replace(JSON.stringify(pairs)));
    this.#changed();
    return result;
  }

  /**
   * @param {string} search
   * @param {string} replacement
   * @param {{matchCase?: boolean, regex?: boolean}} [options]
   * @returns {{occurrences_replaced: number, message: string}}
   */
  findReplace(search, replacement, { matchCase = false, regex = false } = {}) {
    this.#require();
    const result = JSON.parse(this.#session.find_and_replace(search, replacement, matchCase, regex));
    if (result.occurrences_replaced > 0) this.#changed();
    return result;
  }

  /** Word, character, paragraph, table and page counts plus the page setup */
  stats() {
    if (!this.#session) return null;
    return { ...JSON.parse(this.#session.get_stats_json()), page_count: this.#layout?.total_pages ?? 0 };
  }

  /** Fonts the document asks for and whether each is drawn with a substitute */
  fonts() {
    const names = new Set();
    const add = f => f && f.trim() && names.add(f);
    this.#paragraphs().list.forEach(p => {
      add(p.font_family);
      (p.runs || []).forEach(r => add(r.font_family));
    });
    this.#elements.forEach(el => el.type === 'table' && (el.rich_rows || []).forEach(row => row.cells.forEach(c => add(c.font_family))));
    return [...names].map(name => ({ name, substitute: usesSubstitute(name) }));
  }

  // ---------- Editing ----------

  undo() {
    this.#canvasEditor.undo();
  }

  /** Whether `undo()` / `redo()` have a step to apply */
  get canUndo() {
    return !!this.#session?.can_undo();
  }

  get canRedo() {
    return !!this.#session?.can_redo();
  }

  redo() {
    this.#canvasEditor.redo();
  }

  toggleBold() {
    this.#toggleRunFlag('bold');
  }

  toggleItalic() {
    this.#toggleRunFlag('italic');
  }

  toggleUnderline() {
    this.#toggleRunFlag('underline');
  }

  /** Text color of the selection (or of the paragraph at the caret), e.g. "#1E3A8A" */
  setColor(hex, { refocus = true } = {}) {
    const clean = hex.replace('#', '').toUpperCase();
    this.#applyRunFormat(range => range.forEach(c => { c.color = clean; }), refocus);
  }

  /** @param {'left'|'center'|'right'|'both'} align  `both` is justified */
  setAlignment(align) {
    const items = this.#selectedRanges();
    if (!items) return this.#emit('message', { message: this.labels.clickToFormat, error: false });
    this.#applyUpdates(items.map(it => ({ index: it.p.index, runs: it.p.runs || [], align })), true);
  }

  /**
   * Inserts paragraphs as one undo step and leaves the caret after the last one. Each item is
   * a string or `{ text, bold, italic, underline, fontSize, align }` (unset properties are
   * inherited; `fontSize` in points). Text before and after the insertion point keeps its own
   * paragraphs; a "\n" inside an item is a line break within that paragraph.
   * @param {string | Array<string | {text: string, bold?: boolean, italic?: boolean, underline?: boolean, fontSize?: number, align?: 'left'|'center'|'right'|'both'}>} paragraphs
   * @param {{at?: 'cursor'|'start'|'end'}} [options]  `cursor` (default) falls back to the end
   *        of the document when there is no caret
   * @returns {{first: number, count: number}} index of the first new paragraph and how many
   */
  insertParagraphs(paragraphs, { at = 'cursor' } = {}) {
    this.#require();
    const items = (Array.isArray(paragraphs) ? paragraphs : [paragraphs])
      .map(p => (typeof p === 'string' ? { text: p } : p))
      .map(({ text = '', bold, italic, underline, fontSize, align }) => ({
        text: String(text), bold, italic, underline, font_size: fontSize, align,
      }));
    if (!items.length) return { first: -1, count: 0 };

    const pos = this.#insertionPoint(at);
    const result = JSON.parse(this.#session.insert_paragraphs(
      pos.paragraph, pos.offset, JSON.stringify(items), this.#selectionJSON(),
    ));
    this.#changed();
    this.#canvasEditor.select(result.caret);
    return { first: result.first, count: result.count };
  }

  /** Inserts plain text as one undo step; each line becomes a paragraph, as when pasting */
  insertText(text, options) {
    return this.insertParagraphs(String(text).replace(/\r\n?/g, '\n').split('\n'), options);
  }

  /**
   * Inserts a table as one undo step and leaves the caret in the paragraph after it. Text
   * before the insertion point stays above the table and text after it goes below.
   * `insertTable(rows, cols, headers)` (earlier versions) still appends an empty table at the end.
   * @param {{rows: string[][], header?: boolean, widths?: number[], align?: Array<'left'|'center'|'right'>} | number} table
   *        cells as text ("\n" = line break in the cell); `header` (default true) makes the first
   *        row bold, shaded and repeated on every page; `widths` are relative column widths
   * @param {{at?: 'cursor'|'start'|'end'}} [options]
   * @returns {{first: number}} paragraph index of the first cell
   */
  insertTable(table, options, legacyHeaders) {
    this.#require();
    if (typeof table === 'number') {
      const cols = Number(options) || 1;
      const headers = legacyHeaders ?? Array.from({ length: cols }, (_, i) => `${i + 1}`);
      this.#session.add_table(table, cols, JSON.stringify(headers));
      this.#changed();
      return { first: -1 };
    }
    const { at = 'cursor' } = options ?? {};
    const { rows, header = true, widths, align } = table ?? {};
    if (!Array.isArray(rows) || !rows.length) throw new Error('docx-editor: insertTable needs rows');
    const payload = { rows: rows.map(row => (Array.isArray(row) ? row : [row]).map(c => (c == null ? '' : String(c)))), header, widths, align };
    const pos = this.#insertionPoint(at);
    const result = JSON.parse(this.#session.insert_table(
      pos.paragraph, pos.offset, JSON.stringify(payload), this.#selectionJSON(),
    ));
    this.#changed();
    this.#canvasEditor.select(result.caret);
    return { first: result.first };
  }

  /**
   * Inserts a PNG, JPEG or GIF picture as one undo step. Without a size the picture keeps its
   * pixel size; it never exceeds the text width of the page.
   * - In line with the text (default `wrap`): in a paragraph of its own, with the caret after it.
   * - Floating (`wrap` = 'square', 'tight', 'through', 'topAndBottom', 'behind' or 'inFront'):
   *   anchored to the paragraph at the position, which keeps its text and the caret. Place it
   *   with `horizontal` / `vertical` (default: at the paragraph's start).
   * @param {Uint8Array | ArrayBuffer | Blob | string} image bytes, a Blob/File or a `data:` URL
   * @param {{at?: 'cursor'|'start'|'end', width?: number, height?: number, align?: 'left'|'center'|'right', alt?: string,
   *          wrap?: ImageWrap, wrapSide?: 'bothSides'|'left'|'right'|'largest', horizontal?: object, vertical?: object,
   *          distance?: number}} [options]
   *        `width` / `height` in CSS px (one is enough: the aspect ratio is kept)
   * @returns {Promise<{paragraph: number, index: number}>} the new picture (see `images()`)
   */
  async insertImage(image, { at = 'cursor', width, height, align, alt, wrap, wrapSide, horizontal, vertical, distance } = {}) {
    this.#require();
    const bytes = await toBytes(image);
    this.#require();
    const pos = this.#insertionPoint(at);
    const floating = wrap && wrap !== 'inline';
    const payload = {
      ...(floating ? imageUpdatePayload({ wrap, wrapSide, horizontal, vertical, distance }) : {}),
      width, height, align, alt,
    };
    const result = JSON.parse(this.#session.insert_image(
      pos.paragraph, pos.offset, bytes, JSON.stringify(payload), this.#selectionJSON(),
    ));
    this.#changed();
    this.#canvasEditor.select(result.caret);
    return { paragraph: result.paragraph, index: 0 };
  }

  // ---------- Pictures ----------

  /**
   * Every picture of the document body (table cells included) in document order. A picture is
   * addressed by `{ paragraph, index }` (its paragraph and its position among that paragraph's
   * pictures); any object with those two fields, such as an item of this list, works as `ref`.
   * Lengths are CSS px at 100 % zoom.
   * @returns {ImageInfo[]}
   */
  images() {
    if (!this.#session) return [];
    return JSON.parse(this.#session.list_images()).map(imageInfo);
  }

  /** The picture selected with the mouse (or `selectImage()`), or null */
  get selectedImage() {
    const ref = this.#canvasEditor.selectedImage();
    return ref ? this.#imageInfo(ref) : null;
  }

  /** Selects a picture as if it had been clicked (null goes back to the text caret) */
  selectImage(ref) {
    this.#require();
    if (ref && !this.#imageInfo(ref)) throw new Error('docx-editor: no such picture');
    this.#canvasEditor.selectImage(ref ? { paragraph: ref.paragraph, index: ref.index } : null);
  }

  /**
   * Changes a picture as one undo step and returns it as it is now. Unset fields keep their
   * value. Moving needs a floating picture: set `wrap` first (or in the same call).
   * @param {{paragraph: number, index: number}} ref
   * @param {{width?: number, height?: number, keepRatio?: boolean, wrap?: ImageWrap,
   *          wrapSide?: 'bothSides'|'left'|'right'|'largest',
   *          horizontal?: {relativeTo?: string, offset?: number, align?: 'left'|'center'|'right'|'inside'|'outside'},
   *          vertical?: {relativeTo?: string, offset?: number, align?: 'top'|'center'|'bottom'|'inside'|'outside'},
   *          distance?: number | {top?: number, bottom?: number, left?: number, right?: number}, alt?: string}} changes
   *        `keepRatio` (default true) scales the other side when only width or height is given.
   *        `horizontal.relativeTo`: 'margin' | 'page' | 'column' | 'character' | 'leftMargin' | 'rightMargin' | …
   *        `vertical.relativeTo`: 'margin' | 'page' | 'paragraph' | 'line' | 'topMargin' | 'bottomMargin' | …
   *        An `offset` is the distance from the frame's left/top edge; `align` replaces it.
   * @returns {ImageInfo}
   */
  updateImage(ref, changes = {}) {
    this.#require();
    const payload = imageUpdatePayload(changes);
    this.#session.update_image(ref.paragraph, ref.index, JSON.stringify(payload), this.#selectionJSON());
    this.#changed();
    return this.#imageChanged(ref);
  }

  /**
   * Resizes a picture. With only one side the other follows the aspect ratio.
   * @param {{paragraph: number, index: number}} ref
   * @param {{width?: number, height?: number, keepRatio?: boolean, scale?: number}} size
   *        `scale` multiplies the current size (e.g. 0.5)
   */
  resizeImage(ref, { width, height, keepRatio = true, scale } = {}) {
    if (scale > 0) {
      const info = this.#requireImage(ref);
      return this.updateImage(ref, { width: info.width * scale, height: info.height * scale });
    }
    return this.updateImage(ref, { width, height, keepRatio });
  }

  /**
   * Moves a floating picture: to page coordinates `{x, y}` (its top-left corner, CSS px from
   * the page's corner) or by `{dx, dy}`. Its position stays relative to the same frame
   * (margin, paragraph…) when it had an offset; an aligned axis becomes relative to the page.
   * A picture in line with the text first becomes `square` (text wraps around it).
   * @param {{paragraph: number, index: number}} ref
   * @param {{x?: number, y?: number, dx?: number, dy?: number}} to
   */
  moveImage(ref, { x, y, dx, dy } = {}) {
    let info = this.#requireImage(ref);
    if (!info.anchored) info = this.updateImage(ref, { wrap: 'square' });
    if (!info.bounds) throw new Error('docx-editor: the picture has not been laid out yet');
    const deltaX = x != null ? x - info.bounds.x : dx || 0;
    const deltaY = y != null ? y - info.bounds.y : dy || 0;
    if (!deltaX && !deltaY) return info;
    return this.updateImage(ref, moveChanges(info, deltaX, deltaY));
  }

  /**
   * How text flows around a picture:
   * - 'inline'       in line with the text, like a big character
   * - 'square'       text wraps around its box; 'tight' / 'through' wrap around its contour
   *                  (drawn as its box)
   * - 'topAndBottom' text above and below only
   * - 'behind' / 'inFront'  floats behind or over the text, which ignores it
   * @param {{paragraph: number, index: number}} ref
   * @param {ImageWrap} wrap
   * @param {{side?: 'bothSides'|'left'|'right'|'largest', distance?: number | object}} [options]
   */
  setImageWrap(ref, wrap, { side, distance } = {}) {
    return this.updateImage(ref, { wrap, wrapSide: side, distance });
  }

  /**
   * Aligns a picture horizontally: a floating one within the margins, an inline one by
   * aligning its paragraph.
   * @param {{paragraph: number, index: number}} ref
   * @param {'left'|'center'|'right'} align
   */
  alignImage(ref, align) {
    const info = this.#requireImage(ref);
    if (info.anchored) return this.updateImage(ref, { horizontal: { relativeTo: 'margin', align } });
    const p = this.#paragraphs().byIndex.get(ref.paragraph);
    this.#applyUpdates([{ index: ref.paragraph, runs: p?.runs || [], align }], false);
    return this.#imageChanged(ref);
  }

  /** Deletes a picture as one undo step */
  deleteImage(ref) {
    this.#require();
    if (sameRef(ref, this.#canvasEditor.selectedImage())) this.#canvasEditor.selectImage(null);
    this.#session.delete_image(ref.paragraph, ref.index, this.#selectionJSON());
    this.#changed();
  }

  #imageInfo(ref) {
    return this.images().find(img => sameRef(img, ref)) || null;
  }

  /** The picture after a change; tells listeners when it is the selected one */
  #imageChanged(ref) {
    const info = this.#imageInfo(ref);
    if (sameRef(ref, this.#canvasEditor.selectedImage())) this.#emit('imageselect', { image: info });
    return info;
  }

  #requireImage(ref) {
    this.#require();
    const info = this.#imageInfo(ref);
    if (!info) throw new Error('docx-editor: no such picture');
    return info;
  }

  /** Resize/move with the mouse or keyboard (`change` = null deletes) */
  #userImageEdit(ref, change) {
    try {
      if (!change) {
        this.deleteImage(ref);
        return;
      }
      const info = this.#requireImage(ref);
      const changes = {};
      let { dx = 0, dy = 0 } = change;
      if (change.width != null) {
        Object.assign(changes, { width: change.width, height: change.height });
        // Resizing keeps an aligned picture aligned (as Word does); offsets keep the far edge
        if (info.horizontal?.align != null) dx = 0;
        if (info.vertical?.align != null) dy = 0;
      }
      if ((dx || dy) && info.anchored) Object.assign(changes, moveChanges(info, dx, dy));
      if (Object.keys(changes).length) this.updateImage(ref, changes);
    } catch (err) {
      this.#emit('message', { message: String(err), error: true });
    }
  }

  /** Current selection as stored with an undo step */
  #selectionJSON() {
    const selection = this.#canvasEditor.selection();
    return selection ? JSON.stringify({ anchor: selection.anchor, focus: selection.focus }) : undefined;
  }

  setBackgroundColor(hex) {
    this.#require();
    this.#session.set_background_color(hex);
    this.#changed();
  }

  /** Draws a diagonal watermark on every page (view only; not saved into the .docx) */
  setWatermark(text, { opacity = 0.2 } = {}) {
    this.#watermark = text ? { text, opacity } : null;
    this.#render({ full: true });
  }

  // ---------- View ----------

  get zoom() {
    return this.#zoom;
  }

  setZoom(zoom) {
    this.#zoom = Math.min(3, Math.max(0.25, zoom));
    this.#render();
  }

  get selection() {
    return this.#canvasEditor.selection();
  }

  focus() {
    this.#canvasEditor.focus();
  }

  /**
   * Places the selection (a collapsed caret when `focus` is omitted) and scrolls it into view.
   * Positions are `{ paragraph, offset }` as in `selection`.
   */
  select(anchor, focus = anchor) {
    this.#require();
    this.#canvasEditor.select(anchor, focus);
  }

  /** Removes the editor's element from the container; the container itself is left as it was */
  destroy() {
    this.close();
    this.#canvasEditor.destroy();
    this.#unsubscribeFonts();
    document.fonts?.removeEventListener?.('loadingdone', this.#onFontsLoaded);
    cancelAnimationFrame(this.#fontsRelayout);
    this.#fontChecks.forEach(clearTimeout);
    this.root.removeEventListener('keydown', this.#onKeyDown);
    this.root.remove();
  }

  // ---------- Internals ----------

  #require() {
    if (!this.#session) throw new Error('docx-editor: no document is open');
  }

  #emit(type, detail) {
    this.dispatchEvent(new CustomEvent(type, { detail }));
  }

  #replaceSession(session, fileName) {
    this.close();
    this.#session = session;
    this.#fileName = fileName;
    this.#refreshElements();
    this.#render({ full: true });
    this.#emit('load', { fileName });
  }

  #refreshElements() {
    this.#elements = JSON.parse(this.#session.get_document_elements_json());
  }

  #changed() {
    this.#render();
    this.#emit('change', {});
  }

  /** Every paragraph (body and table cells) in document order, tagged with its container */
  #paragraphs() {
    if (this.#paragraphCache.source !== this.#elements) {
      const list = [];
      this.#elements.forEach(el => {
        if (el.type === 'paragraph') {
          list.push({ ...el, container: 'body' });
        } else if (el.type === 'table') {
          (el.rich_rows || []).forEach((row, r) => row.cells.forEach((cell, c) => {
            (cell.paragraphs || []).forEach(p => list.push({ ...p, container: `${el.index}:${r}:${c}` }));
          }));
        }
      });
      list.sort((a, b) => a.index - b.index);
      this.#paragraphCache = { source: this.#elements, list, byIndex: new Map(list.map(p => [p.index, p])) };
    }
    return this.#paragraphCache;
  }

  #render({ full = false } = {}) {
    const session = this.#session;
    if (!session) return;
    try {
      // Edits made through the caret editor change the elements too
      this.#refreshElements();
      const redrawAll = full || this.#rendered.session !== session || this.#rendered.zoom !== this.#zoom;
      const previous = this.#layout;
      this.#fontWidths = fontWidthsSignature(this.#requestedFaces);
      const layout = JSON.parse(session.compute_canvas_layout_json(
        this.#watermark?.text ?? null, this.#watermark?.opacity ?? 0.2, measureText, !redrawAll,
      ));
      const fresh = new Set(layout.pages.filter(p => !p.unchanged).map(p => p.page_number));
      layout.pages = layout.pages.map((page, i) => (page.unchanged ? previous.pages[i] : page));
      this.#layout = layout;

      this.#canvasEditor.beginRender();
      if (redrawAll) {
        this.#pages.replaceChildren();
        this.#rendered = { session, zoom: this.#zoom, cards: new Map() };
      }
      const cards = this.#rendered.cards;
      layout.pages.forEach(page => {
        let card = cards.get(page.page_number);
        if (!card || fresh.has(page.page_number)) {
          const next = this.#pageCard(page, layout.total_pages);
          if (card) card.replaceWith(next);
          else this.#pages.appendChild(next);
          card = next;
          cards.set(page.page_number, card);
        }
        this.#canvasEditor.attachPage(page, card, card.querySelector('canvas'));
      });
      cards.forEach((card, number) => {
        if (number > layout.total_pages) {
          card.remove();
          cards.delete(number);
        }
      });
      this.#loadFonts(layout);
      this.#canvasEditor.endRender();
    } catch (err) {
      console.error('docx-editor: render failed', err);
      this.#emit('message', { message: String(err), error: true });
    }
  }

  #pageCard(page, totalPages) {
    const card = document.createElement('div');
    card.className = 'docx-editor-page';
    card.dataset.pageNumber = page.page_number;
    if (this.options.pageLabels) {
      const label = document.createElement('div');
      label.className = 'docx-editor-page-label';
      label.textContent = this.labels.page(page.page_number, totalPages);
      card.appendChild(label);
    }

    const dpr = window.devicePixelRatio || 1;
    const canvas = document.createElement('canvas');
    canvas.className = 'docx-editor-canvas';
    canvas.dataset.pageNum = page.page_number;
    const w = page.width * this.#zoom;
    const h = page.height * this.#zoom;
    canvas.width = Math.round(w * dpr);
    canvas.height = Math.round(h * dpr);
    canvas.style.width = `${w}px`;
    canvas.style.height = `${h}px`;

    const draw = () => {
      const ctx = canvas.getContext('2d');
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      disableLigatures(ctx);
      ctx.scale(dpr * this.#zoom, dpr * this.#zoom);
      ctx.fillStyle = page.bg_color || '#FFFFFF';
      ctx.fillRect(0, 0, page.width, page.height);
      return drawPageItems(ctx, page.items, { gridlines: this.options.gridlines });
    };
    const pending = draw();
    if (pending.length) Promise.all(pending).then(() => canvas.isConnected && draw());

    card.appendChild(canvas);
    return card;
  }

  // Web fonts load lazily: request every face the layout uses and lay out again once they
  // arrive, since widths measured with a fallback font are wrong
  #loadFonts(layout) {
    const pending = [...layoutFontFaces(layout)].filter(face => !this.#requestedFaces.has(face));
    if (!pending.length) return;
    pending.forEach(face => this.#requestedFaces.add(face));
    // One request per family: WebKit only loads the first family of a list ("Calibri", which
    // usually does not exist), never the substitute behind it
    const requests = pending.flatMap(splitFontFamilies);
    Promise.all(requests.map(face => document.fonts.load(face).catch(() => []))).then(results => {
      if (results.some(loaded => loaded.length > 0)) this.#relayoutForFonts();
    });
    this.#watchFontWidths();
  }

  /** A font reported as loaded may still draw with the fallback for a moment (Safari): check
   *  the measured widths for a few seconds and lay out again as soon as they change */
  #watchFontWidths() {
    this.#fontChecks.forEach(clearTimeout);
    this.#fontChecks = [100, 300, 700, 1500, 3000, 6000].map(ms => setTimeout(() => {
      if (this.#session && fontWidthsSignature(this.#requestedFaces) !== this.#fontWidths) this.#relayoutForFonts();
    }, ms));
  }

  /** Measures and draws everything again with the fonts available now (once per frame) */
  #relayoutForFonts() {
    cancelAnimationFrame(this.#fontsRelayout);
    this.#fontsRelayout = requestAnimationFrame(() => {
      if (!this.#session) return;
      this.#session.reset_measurements();
      this.#render({ full: true });
      this.#watchFontWidths();
    });
  }

  // ---------- Formatting ----------

  /** Paragraphs touched by the selection, each with its selected character range */
  #selectedRanges() {
    const sel = this.#canvasEditor.selection();
    if (!sel) return null;
    const items = [];
    for (let i = sel.start.paragraph; i <= sel.end.paragraph; i++) {
      const p = this.#paragraphs().byIndex.get(i);
      if (!p) continue;
      const chars = paragraphChars(p);
      const start = sel.collapsed || i !== sel.start.paragraph ? 0 : sel.start.offset;
      const end = sel.collapsed || i !== sel.end.paragraph ? chars.length : sel.end.offset;
      items.push({ p, chars, start, end });
    }
    return items;
  }

  #applyUpdates(updates, refocus) {
    if (!updates.length) return;
    try {
      this.#session.update_paragraphs_runs(JSON.stringify(updates));
    } catch (err) {
      this.#emit('message', { message: String(err), error: true });
      return;
    }
    this.#changed();
    if (refocus) this.#canvasEditor.focus();
  }

  /** Where `insertParagraphs` puts new content: the caret, or the start/end of the body */
  #insertionPoint(at) {
    if (at === 'cursor') {
      const selection = this.#canvasEditor.selection();
      if (selection) return selection.end;
    }
    const list = this.#paragraphs().list;
    const body = list.filter(p => p.container === 'body');
    if (at === 'start' && body.length) return { paragraph: body[0].index, offset: 0 };
    const last = body[body.length - 1] ?? list[list.length - 1];
    if (!last) throw new Error('docx-editor: the document has no paragraphs');
    return { paragraph: last.index, offset: Array.from(last.text || '').length };
  }

  /** Formatting applies to the selected characters, or to the paragraph at the caret */
  #applyRunFormat(change, refocus = true) {
    const items = this.#selectedRanges();
    if (!items) return this.#emit('message', { message: this.labels.clickToFormat, error: false });
    const range = items.flatMap(it => it.chars.slice(it.start, it.end));
    if (!range.length) return;
    change(range);
    this.#applyUpdates(items
      .filter(it => it.end > it.start)
      .map(it => ({ index: it.p.index, runs: mergeRuns(it.chars), align: it.p.align || 'left' })), refocus);
  }

  #toggleRunFlag(flag) {
    this.#applyRunFormat(range => {
      const on = !range.every(c => c[flag]);
      range.forEach(c => { c[flag] = on; });
    });
  }

  /** Formatting at the caret (the character before it) or of the selection's start */
  #formatAt(sel) {
    if (!sel) return null;
    const p = this.#paragraphs().byIndex.get(sel.start.paragraph);
    if (!p) return null;
    const chars = paragraphChars(p);
    const range = sel.collapsed
      ? chars.slice(Math.max(0, sel.start.offset - 1), Math.max(1, sel.start.offset))
      : chars.slice(sel.start.offset, sel.start.paragraph === sel.end.paragraph ? sel.end.offset : chars.length);
    const all = flag => range.length > 0 && range.every(c => c[flag]);
    const color = (range[0]?.color || p.color || '').replace('#', '');
    return {
      bold: all('bold'),
      italic: all('italic'),
      underline: all('underline'),
      color: /^[0-9a-fA-F]{6}$/.test(color) ? `#${color.toUpperCase()}` : null,
      align: p.align === 'justify' ? 'both' : p.align || 'left',
    };
  }
}

/** One entry per character, carrying the formatting of its run */
function paragraphChars(p) {
  const chars = [];
  (p.runs || []).forEach(run => Array.from(run.text).forEach(ch => chars.push({ ...run, text: ch })));
  return chars;
}

const sameFormat = (a, b) => a.bold === b.bold && a.italic === b.italic && a.underline === b.underline
  && (a.color || '') === (b.color || '') && a.font_size === b.font_size && a.font_family === b.font_family;

function mergeRuns(chars) {
  const runs = [];
  chars.forEach(c => {
    const last = runs[runs.length - 1];
    if (last && sameFormat(last, c)) last.text += c.text;
    else runs.push({ ...c });
  });
  return runs;
}

const GENERIC_FAMILIES = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui']);

/** `700 16px "Calibri", "Carlito", sans-serif` → ['700 16px "Calibri"', '700 16px "Carlito"'] */
function splitFontFamilies(face) {
  const match = face.match(/^(.*?\d+(?:\.\d+)?px)\s+(.*)$/);
  if (!match) return [face];
  return match[2]
    .split(',')
    .map(f => f.trim())
    .filter(f => f && !GENERIC_FAMILIES.has(f.replace(/["']/g, '')))
    .map(f => `${match[1]} ${f}`);
}

const sameRef = (a, b) => !!a && !!b && a.paragraph === b.paragraph && a.index === b.index;

/**
 * @typedef {'inline'|'square'|'tight'|'through'|'topAndBottom'|'behind'|'inFront'} ImageWrap
 * @typedef {object} ImageInfo
 * @property {number} paragraph
 * @property {number} index         among the paragraph's pictures
 * @property {number} width
 * @property {number} height
 * @property {ImageWrap} wrap
 * @property {string|null} wrapSide
 * @property {boolean} anchored     floating (any wrap but 'inline')
 * @property {{relativeTo: string, offset: number, align: string|null}|null} horizontal  null when inline
 * @property {{relativeTo: string, offset: number, align: string|null}|null} vertical
 * @property {{top: number, bottom: number, left: number, right: number}} distance  space kept from the text
 * @property {string} alt
 * @property {number} textOffset    character offset in the paragraph where it sits
 * @property {{page: number, x: number, y: number, width: number, height: number}|null} bounds  where it is drawn
 */
function imageInfo(raw) {
  const axis = (relativeTo, offset, align) => (raw.anchored ? { relativeTo, offset, align: align ?? null } : null);
  return {
    paragraph: raw.paragraph,
    index: raw.index,
    width: raw.width,
    height: raw.height,
    wrap: raw.wrap || (raw.anchored ? (raw.behind_text ? 'behind' : 'inFront') : 'inline'),
    wrapSide: raw.wrap_side || null,
    anchored: raw.anchored,
    horizontal: axis(raw.h_relative, raw.h_offset, raw.h_align),
    vertical: axis(raw.v_relative, raw.v_offset, raw.v_align),
    distance: { top: raw.dist_top, bottom: raw.dist_bottom, left: raw.dist_left, right: raw.dist_right },
    alt: raw.alt || '',
    textOffset: raw.offset,
    bounds: raw.bounds ?? null,
  };
}

/** `updateImage` changes as the engine's JSON (snake_case, undefined fields left out) */
function imageUpdatePayload({ width, height, keepRatio, wrap, wrapSide, horizontal, vertical, distance, alt } = {}) {
  const payload = { width, height, keep_ratio: keepRatio, wrap, wrap_side: wrapSide, alt };
  if (horizontal) Object.assign(payload, { h_relative: horizontal.relativeTo, h_offset: horizontal.offset, h_align: horizontal.align ?? undefined });
  if (vertical) Object.assign(payload, { v_relative: vertical.relativeTo, v_offset: vertical.offset, v_align: vertical.align ?? undefined });
  if (typeof distance === 'number') payload.distance = distance;
  else if (distance) Object.assign(payload, { dist_top: distance.top, dist_bottom: distance.bottom, dist_left: distance.left, dist_right: distance.right });
  return payload;
}

/** Position changes that move a floating picture by (dx, dy) page px: an offset grows in its
 *  own frame; an aligned axis becomes an offset from the page edge */
function moveChanges(info, dx, dy) {
  const axis = (pos, delta, absolute) => (pos.align == null
    ? { relativeTo: pos.relativeTo, offset: pos.offset + delta }
    : { relativeTo: 'page', offset: absolute });
  const changes = {};
  if (dx) changes.horizontal = axis(info.horizontal, dx, (info.bounds?.x ?? 0) + dx);
  if (dy) changes.vertical = axis(info.vertical, dy, (info.bounds?.y ?? 0) + dy);
  return changes;
}

/** Bytes of an image given as bytes, a Blob/File or a `data:` URL */
async function toBytes(image) {
  if (image instanceof Uint8Array) return image;
  if (image instanceof ArrayBuffer) return new Uint8Array(image);
  if (typeof Blob !== 'undefined' && image instanceof Blob) return new Uint8Array(await image.arrayBuffer());
  if (typeof image === 'string' && image.startsWith('data:')) {
    const comma = image.indexOf(',');
    const meta = image.slice(0, comma);
    const data = image.slice(comma + 1);
    if (meta.endsWith(';base64')) {
      const bin = atob(data);
      return Uint8Array.from(bin, c => c.charCodeAt(0));
    }
    return new TextEncoder().encode(decodeURIComponent(data));
  }
  throw new Error('docx-editor: insertImage needs a Uint8Array, ArrayBuffer, Blob or data: URL');
}
