// DocxEditor: a Word document editor mounted in any element. Rust (WebAssembly) parses the
// .docx, lays out the pages and applies lossless edits; this class renders the pages to
// canvases, wires the caret/selection editor and exposes a small API plus DOM events.
import { DocxSession, engineReady, initEngine, scanMarkers } from './engine.js';
import { createCanvasEditor } from './canvas_editor.js';
import { checkIndex, engineCall, tableRowsPayload } from './table_api.js';
import { checkSectionKey, planMissingSections, schemaPayload, sectionContent, sectionIsEmpty, validateSections } from './section_api.js';
import { createSectionLayer } from './section_layer.js';
import { collectMarkers, innermostSection, relocateMarker, variablePairs } from './marker_api.js';
import { createActivityLayer, textDiff } from './activity.js';
import { onFontsChanged } from './fonts_registry.js';
import { disableLigatures, drawPageItems, fontWidthsSignature, layoutFontFaces, measureText, usesSubstitute } from './render.js';

/** First index of text box paragraphs (the engine numbers them apart from the body) */
const TEXT_BOX_BASE = 1 << 24;

const LABELS = {
  es: {
    editor: 'Editor del documento',
    page: (n, total) => `Página ${n} de ${total}`,
    nothingToUndo: 'No hay nada para deshacer.',
    nothingToRedo: 'No hay nada para rehacer.',
    clickToFormat: 'Haz clic en el texto para aplicar formato.',
    untitled: 'documento.docx',
    header: 'Encabezado',
    footer: 'Pie de página',
    agent: 'Agente',
    goToActivity: 'Ir a donde se está trabajando',
    sectionContentLocked: title => `La sección «${title}» está bloqueada: su contenido no se puede editar.`,
    sectionLocked: title => `La sección «${title}» está protegida: no se puede borrar.`,
    sectionLockedMark: '🛡',
    sectionContentLockedMark: '🔒',
    sectionIssues: {
      missing: title => `Falta la sección «${title}».`,
      duplicated: (title, { count }) => `La sección «${title}» aparece ${count} veces.`,
      parent: (title, { expected }) => (expected ? `La sección «${title}» debe ir dentro de «${expected}».` : `La sección «${title}» no debe ir dentro de otra.`),
      empty: title => `La sección «${title}» está vacía.`,
      order: title => `La sección «${title}» está fuera de orden.`,
      unexpected: title => `La sección «${title}» no está en el esquema.`,
    },
  },
  en: {
    editor: 'Document editor',
    page: (n, total) => `Page ${n} of ${total}`,
    nothingToUndo: 'Nothing to undo.',
    nothingToRedo: 'Nothing to redo.',
    clickToFormat: 'Click in the text to apply formatting.',
    untitled: 'document.docx',
    header: 'Header',
    footer: 'Footer',
    agent: 'Agent',
    goToActivity: 'Go to where the work is happening',
    sectionContentLocked: title => `The section “${title}” is locked: its content cannot be edited.`,
    sectionLocked: title => `The section “${title}” is protected: it cannot be deleted.`,
    sectionLockedMark: '🛡',
    sectionContentLockedMark: '🔒',
    sectionIssues: {
      missing: title => `The section “${title}” is missing.`,
      duplicated: (title, { count }) => `The section “${title}” appears ${count} times.`,
      parent: (title, { expected }) => (expected ? `The section “${title}” must be inside “${expected}”.` : `The section “${title}” must not be inside another one.`),
      empty: title => `The section “${title}” is empty.`,
      order: title => `The section “${title}” is out of order.`,
      unexpected: title => `The section “${title}” is not in the schema.`,
    },
  },
};

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/**
 * @typedef {object} DocxEditorOptions
 * @property {number} [zoom=1]            1 = 100 %
 * @property {boolean} [gridlines=true]   dashed guides where table cells have no border
 * @property {boolean} [pageLabels=true]  "Page 1 of 3" above each page
 * @property {'es'|'en'} [locale='es']
 * @property {object} [labels]            overrides for the locale's texts
 * @property {string|URL} [wasmUrl]       where docx_engine_bg.wasm is served from
 * @property {boolean} [readOnly=false]   the user can navigate, select and copy but not edit
 *                                        (see `setReadOnly()`)
 * @property {number|false} [flash=1800]  ms a change made from code stays highlighted (see
 *                                        `markActivity()`); false: no highlight
 * @property {boolean} [showSections=false] outline the logical sections in the margin (see
 *                                        `setShowSections()`)
 *
 * Events (dispatched on the editor; `<docx-editor>` re-dispatches them on the element):
 * - `load`             a document was opened
 * - `change`           the document changed (typing, formatting, replacements…)
 * - `selectionchange`  detail: { selection, format } with format = { bold, italic, underline, color, align }
 * - `imageselect`      detail: { image } the picture selected with the mouse or `selectImage()` (null when none)
 * - `readonlychange`   detail: { readOnly } after `setReadOnly()` changed the mode
 * - `message`          detail: { message, error } hints and errors meant for the user
 * - `activity`         detail: { activities } the activity marks changed (see `activities()`)
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
  /** Header and footer paragraphs (numbered from HEADER_FOOTER_BASE), edited in header mode */
  #partParagraphs = [];
  #requestedFaces = new Set();
  #watermark = null;
  #zoom;
  #canvasEditor;
  #pages;
  #unsubscribeFonts;
  #readOnly = false;
  #onFontsLoaded;
  #fontsRelayout = 0;
  /** Probe widths of the fonts the last layout was measured with */
  #fontWidths = '';
  #fontChecks = [];
  #onKeyDown;
  #activity;
  #activityJSON = '[]';
  #sectionLayer;
  /** `sections()` until the document changes */
  #sectionCache = null;

  /** Use `DocxEditor.create()` unless the engine is already initialized (`initEngine()`) */
  constructor(container, options = {}) {
    super();
    if (!engineReady()) throw new Error('docx-editor: await initEngine() or use DocxEditor.create()');
    this.options = { zoom: 1, gridlines: true, pageLabels: true, locale: 'es', readOnly: false, flash: 1800, showSections: false, ...options };
    this.#readOnly = !!this.options.readOnly;
    const base = LABELS[this.options.locale] || LABELS.en;
    this.labels = { ...base, ...options.labels, sectionIssues: { ...base.sectionIssues, ...options.labels?.sectionIssues } };
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
      readOnly: () => this.#readOnly,
      editBlocked: (start, end) => this.#lockedMessage(start.paragraph, end.paragraph),
    });
    this.root.classList.toggle('read-only', this.#readOnly);
    this.#canvasEditor.readOnlyChanged();

    this.#activity = createActivityLayer({
      root: this.root,
      labels: this.labels,
      cards: () => this.#rendered.cards,
      zoom: () => this.#zoom,
      sections: () => this.sections(),
      sectionAt: index => innermostSection(this.sections(), index),
      lengthOf: index => paragraphChars(this.#paragraphs().byIndex.get(index) || {}).length,
      rects: (a, b) => JSON.parse(this.#session?.selection_rects_range(
        a.paragraph, a.offset, b.paragraph, b.offset, measureText, undefined,
      ) || '[]'),
      changed: activities => {
        const json = JSON.stringify(activities);
        if (json === this.#activityJSON) return;
        this.#activityJSON = json;
        this.#emit('activity', { activities });
      },
    });

    this.#sectionLayer = createSectionLayer({
      labels: this.labels,
      cards: () => this.#rendered.cards,
      zoom: () => this.#zoom,
      sections: () => this.sections(),
      lengthOf: index => paragraphChars(this.#paragraphs().byIndex.get(index) || {}).length,
      rects: (a, b) => JSON.parse(this.#session?.selection_rects_range(
        a.paragraph, a.offset, b.paragraph, b.offset, measureText, undefined,
      ) || '[]'),
      select: section => this.#canvasEditor.select(
        { paragraph: section.first, offset: 0 },
        { paragraph: section.last, offset: paragraphChars(this.#paragraphs().byIndex.get(section.last) || {}).length },
      ),
    });
    if (this.options.showSections) this.#sectionLayer.setVisible(true);

    // Formatting shortcuts while the caret is in this editor
    this.#onKeyDown = e => {
      if (!(e.metaKey || e.ctrlKey) || e.altKey || !this.#session || this.#readOnly) return;
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
    this.#activity?.clear();
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

  /**
   * The document as a vector PDF of the pages as laid out here: selectable text in the fonts
   * the page shows (uploaded fonts or the bundled substitutes), pictures, tables and shapes.
   * pdf-lib is loaded the first time this is called.
   * @param {{watermark?: boolean}} [options] `watermark` (default true) includes the view
   *        watermark set with `setWatermark()`
   * @returns {Promise<Uint8Array>}
   */
  async exportPdf({ watermark = true } = {}) {
    this.#require();
    // The PDF uses the layout's measurements: make sure they were taken with the fonts the
    // page has now (a font that finished loading may not have been laid out yet)
    await document.fonts?.ready;
    if (!this.#layout || fontWidthsSignature(this.#requestedFaces) !== this.#fontWidths) {
      this.#session.reset_measurements();
      this.#render({ full: true });
    }
    const { exportPdf } = await import('./pdf_export.js');
    return exportPdf(this.#layout, { title: this.#fileName.replace(/\.docx$/i, ''), watermark });
  }

  /** `exportPdf()` as a Blob (application/pdf) */
  async exportPdfBlob(options) {
    return new Blob([await this.exportPdf(options)], { type: 'application/pdf' });
  }

  /** Downloads the document as PDF (browsers; in Tauri write `exportPdf()` with the fs plugin) */
  async downloadPdf(fileName = this.#fileName.replace(/\.docx$/i, '') + '.pdf', options) {
    const url = URL.createObjectURL(await this.exportPdfBlob(options));
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
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

  /** Fields found in the document as written, e.g. ["{{CLIENTE}}", "{fecha}"] (see `markers()`) */
  variables() {
    return [...new Set(this.markers({ kind: 'field' }).map(m => m.raw))];
  }

  /**
   * Fills placeholders, keeping each one's formatting even when Word split it across runs.
   * A key with braces is replaced as written; a bare key fills every spelling of that field:
   * `{ CLIENTE: 'Acme' }` fills `{{CLIENTE}}`, `{{ CLIENTE }}`, `{CLIENTE}` and `{{CLIENTE: hint}}`.
   * @param {Record<string, string> | {key: string, value: string}[]} values
   * @returns {{occurrences_replaced: number, message: string}}
   */
  replaceVariables(values) {
    this.#require();
    const fields = this.markers({ kind: 'field' });
    const pairs = variablePairs(values, fields);
    const result = JSON.parse(this.#session.batch_replace(JSON.stringify(pairs)));
    this.#changed();
    const filled = new Set(pairs.map(p => p.key));
    new Set(fields.filter(m => filled.has(m.raw)).map(m => m.paragraph)).forEach(p => this.#flash({ from: p, to: p }));
    return result;
  }

  // ---------- Markers ----------

  /**
   * Markers of the standard lexicon in document order (body, table cells, text boxes, headers
   * and footers):
   * - `field`: `{{KEY}}`, `{{KEY: hint}}` or `{key}`, a value an automation fills
   * - `pending`: `[ … ]`, content still to be written (`[…]`, `[Completar: alcance]`);
   *   citations (`[1]`, `[2-4]`), checkboxes (`[ ]`, `[x]`) and `[sic]` are not markers
   * `offset` / `length` are in code points, as `selection`; `place` is `body`, `table`,
   * `textbox` or `headerFooter`; `section` the key of the deepest logical section holding it.
   * @param {{kind?: 'field'|'pending', section?: string}} [filter]
   * @returns {Array<{id: string, kind: 'field'|'pending', syntax: 'double'|'single'|'bracket', raw: string, key: string|null, hint: string|null, paragraph: number, offset: number, length: number, place: string, section: string|null}>}
   */
  markers({ kind, section } = {}) {
    if (!this.#session) return [];
    const all = collectMarkers(this.#paragraphs().list, this.sections(), scanMarkers, { kind });
    return section == null ? all : all.filter(m => m.section === section);
  }

  /**
   * Replaces one marker (from `markers()`, or its `id`) with `text` as one undo step, keeping
   * the formatting of the marker's text ("\n" starts a new paragraph). Text typed before the
   * marker since it was listed is taken into account; throws when the marker is gone.
   * The change is highlighted for a moment (see the `flash` option).
   * @returns {{paragraph: number, offset: number, end: {paragraph: number, offset: number}}}
   */
  fillMarker(marker, text) {
    this.#require();
    const ref = typeof marker === 'string' ? this.markers().find(m => m.id === marker) : marker;
    if (!ref || !Number.isInteger(ref.paragraph) || typeof ref.raw !== 'string') {
      throw new Error(`docx-editor: fillMarker: unknown marker ${typeof marker === 'string' ? marker : JSON.stringify(marker)}`);
    }
    const p = this.#paragraphs().byIndex.get(ref.paragraph);
    const at = p && relocateMarker(ref, p.text);
    if (!at) throw new Error(`docx-editor: fillMarker: ${ref.raw} is no longer in paragraph ${ref.paragraph}`);
    const value = text == null ? '' : String(text).replace(/\r\n?/g, '\n');
    const end = JSON.parse(engineCall(() => this.#session.edit(
      ref.paragraph, at.offset, ref.paragraph, at.offset + at.length, value, this.#selectionJSON(), false,
    )));
    this.#changed();
    this.#flash(end.paragraph === ref.paragraph
      ? { paragraph: ref.paragraph, offset: at.offset, length: end.offset - at.offset }
      : { from: ref.paragraph, to: end.paragraph });
    return { paragraph: ref.paragraph, offset: at.offset, end };
  }

  // ---------- Activity ----------

  /**
   * Shows where an agent (or any code) is working, so the person reading can find it: a tinted
   * area with a bar and a label in the margin, plus a chip at the top of the editor that
   * scrolls there on click. View only: nothing is saved into the .docx, the caret and the
   * scroll position are left alone. The mark follows its target through later edits.
   * @param {string | {from: number, to?: number} | {first: number, count: number} | {paragraph: number, offset?: number, length?: number}} target
   *        a section key, a paragraph range, what `insertParagraphs` / `replaceSection`…
   *        return, or a span such as a marker
   * @param {{label?: string, color?: string, ttl?: number}} [options] `ttl`: ms until it goes
   *        away by itself (by default it stays until `done()`)
   * @returns {{id: string, update: (target?: object|string, options?: {label?: string}) => void, done: () => void, reveal: () => boolean}}
   */
  markActivity(target, { label, color, ttl } = {}) {
    const id = this.#activity.add('work', target, { label, color, ttl });
    return {
      id,
      update: (next, opts) => this.#activity.update(id, next, opts),
      done: () => this.#activity.remove(id),
      reveal: () => this.#activity.reveal(id),
    };
  }

  /**
   * Activity marks shown now, in document order (`kind`: `work` from `markActivity()`, `flash`
   * for a change made from code a moment ago); `page` is where the mark starts
   * @returns {Array<{id: string, kind: 'work'|'flash', label: string, from: number|null, to: number|null, section: string|null, page: number|null}>}
   */
  activities() {
    return this.#activity.list();
  }

  /** Scrolls an activity mark (the latest `work` one by default) into view; false when none */
  revealActivity(id) {
    return this.#activity.reveal(id);
  }

  /** Removes one activity mark, or every one when `id` is omitted */
  clearActivity(id) {
    if (id == null) this.#activity.clear();
    else this.#activity.remove(id);
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
      .map(toNewParagraph);
    if (!items.length) return { first: -1, count: 0 };

    const pos = this.#insertionPoint(at);
    const result = JSON.parse(this.#session.insert_paragraphs(
      pos.paragraph, pos.offset, JSON.stringify(items), this.#selectionJSON(),
    ));
    this.#changed();
    this.#canvasEditor.select(result.caret);
    this.#flash(result);
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
   * @param {{rows: Array<Array<string | {text?: string, image?: Uint8Array|ArrayBuffer|string, width?: number, height?: number, alt?: string}>>,
   *          header?: boolean, widths?: number[], align?: Array<'left'|'center'|'right'>} | number} table
   *        cells as text ("\n" = line break in the cell) or `{ text, image }`: a PNG, JPEG or GIF
   *        (bytes or a `data:` URL; read a Blob/File first with `await file.arrayBuffer()`) that
   *        fills the column width unless `width` / `height` (CSS px) are given, with the text
   *        below it; `header` (default true) makes the first row bold, shaded and repeated on
   *        every page; `widths` are relative column widths
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
    const payload = { rows: rows.map(row => (Array.isArray(row) ? row : [row]).map(tableCell)), header, widths, align };
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
   * Every picture, text box and shape of the document body (table cells included) in document
   * order (`kind` tells them apart). Each is addressed by `{ paragraph, index }` (its paragraph
   * and its position among that paragraph's drawings); any object with those two fields, such
   * as an item of this list, works as `ref`, and the methods below work on all three kinds.
   * Lengths are CSS px at 100 % zoom. Legacy VML text boxes are drawn but not listed.
   * @returns {ImageInfo[]}
   */
  /**
   * Blocks of the body in order: paragraphs with their heading level (1 = Heading 1, null when
   * not in the outline) and tables with the range of their cells' paragraphs
   * @returns {Array<{kind: 'paragraph', paragraph: number, text: string, headingLevel: number|null} | {kind: 'table', index: number|null, first: number, last: number, rows: number, cols: number}>}
   */
  structure() {
    if (!this.#session) return [];
    const tableIndex = new Map(this.tables().map(t => [t.first, t.index]));
    return (this.#elements || []).flatMap(el => {
      if (el.type === 'paragraph') {
        const level = el.outline_level;
        return [{ kind: 'paragraph', paragraph: el.index, text: el.text || '', headingLevel: level == null ? null : level + 1 }];
      }
      if (el.type === 'table') {
        const indices = (el.rich_rows || []).flatMap(row => row.cells.flatMap(cell => (cell.paragraphs || []).map(p => p.index)));
        if (!indices.length) return [];
        const rows = (el.rich_rows || []).length;
        const cols = Math.max(0, ...(el.rich_rows || []).map(row => row.cells.length));
        const first = Math.min(...indices);
        const index = tableIndex.get(first) ?? null;
        return [{ kind: 'table', index, first, last: Math.max(...indices), rows, cols }];
      }
      return [];
    });
  }

  // ---------- Existing tables ----------

  /**
   * Tables of the body in order (tables nested in cells or text boxes are not listed). `index`
   * is what `setTableCell`, `insertTableRows` and `deleteTableRows` take; `cols` counts grid
   * columns; `header` is the text of the first row by column; `first` / `last` are the range
   * of the cells' paragraphs (the same numbering as `structure()` and `selection`).
   * @returns {Array<{index: number, rows: number, cols: number, header: string[], first: number, last: number}>}
   */
  tables() {
    if (!this.#session) return [];
    return JSON.parse(this.#session.tables_json());
  }

  /**
   * Replaces the text of a cell as one undo step, keeping the cell's paragraph properties and
   * the formatting of its first run ("\n" = line break inside the cell). A cell merged across
   * columns is written through its first column; a cell continuing a vertical merge is refused.
   */
  setTableCell(table, row, col, text) {
    this.#require();
    checkIndex('setTableCell', { table, row, col });
    const value = text == null ? '' : String(text).replace(/\r\n?/g, '\n');
    engineCall(() => this.#session.set_table_cell(table, row, col, value, this.#selectionJSON()));
    this.#changed();
    this.#flashTable(table, row, row, col);
  }

  /**
   * Inserts rows before row `at` as one undo step (`at` = 0: at the top; `at` = row count: at
   * the end). Each new row copies the formatting of row `copyFormatFrom` (borders, shading,
   * cell widths, paragraph properties and first run); by default the row before `at`, or the
   * next one at the top. Short rows are completed with empty cells.
   * @param {number} table
   * @param {number} at
   * @param {Array<Array<string>>} rows cell texts ("\n" = line break)
   * @param {{copyFormatFrom?: number}} [options]
   * @returns {{first: number, count: number}} the first new row and how many were inserted
   */
  insertTableRows(table, at, rows, { copyFormatFrom } = {}) {
    this.#require();
    checkIndex('insertTableRows', { table, at, ...(copyFormatFrom == null ? {} : { copyFormatFrom }) });
    const payload = tableRowsPayload(rows);
    const result = JSON.parse(engineCall(() => this.#session.insert_table_rows(
      table, at, JSON.stringify(payload), copyFormatFrom ?? undefined, this.#selectionJSON(),
    )));
    this.#changed();
    if (result.count > 0) this.#flashTable(table, result.first, result.first + result.count - 1);
    return { first: result.first, count: result.count };
  }

  /** Deletes `count` rows from row `from` as one undo step; a table cannot be left without rows */
  deleteTableRows(table, from, count) {
    this.#require();
    checkIndex('deleteTableRows', { table, from, count });
    engineCall(() => this.#session.delete_table_rows(table, from, count, this.#selectionJSON()));
    this.#changed();
  }

  /**
   * Replaces the blocks from paragraph `from` to `to` (both included, tables between them too)
   * with new paragraphs, as one undo step; the caret goes after the last one. Throws when the
   * range cuts a table or holds a section break.
   * @returns {{first: number, count: number}}
   */
  replaceParagraphs({ from, to }, paragraphs) {
    this.#require();
    const items = (Array.isArray(paragraphs) ? paragraphs : [paragraphs])
      .map(p => (typeof p === 'string' ? { text: p } : p))
      .map(toNewParagraph);
    if (!items.length) throw new Error('docx-editor: replaceParagraphs needs at least one paragraph');
    const result = JSON.parse(this.#session.replace_paragraphs(from, to, JSON.stringify(items), this.#selectionJSON()));
    this.#changed();
    this.#canvasEditor.select(result.caret);
    this.#flash(result);
    return { first: result.first, count: result.count };
  }

  // ---------- Logical sections ----------

  /**
   * Logical sections (Introducción, Conclusiones…) in document order, outer before inner.
   * They are content controls tagged `section:<key>` in the .docx, so other tools find the
   * same boundaries. `first` / `last` use the numbering of `structure()` (null when empty).
   * `locked`: it cannot be deleted; `contentLocked`: its content cannot be edited by hand.
   * @returns {Array<{key: string, title: string, first: number|null, last: number|null, depth: number, parent: string|null, locked: boolean, contentLocked: boolean}>}
   */
  sections() {
    if (!this.#session) return [];
    this.#sectionCache ??= JSON.parse(this.#session.sections_json());
    return this.#sectionCache.map(s => ({ ...s }));
  }

  /** A section with its blocks (as in `structure()`) and plain text, or null when missing */
  sectionContent(key) {
    checkSectionKey('sectionContent', key);
    const section = this.sections().find(s => s.key === key);
    return section ? sectionContent(section, this.structure(), this.#elements || []) : null;
  }

  /**
   * Wraps the blocks from paragraph `from` to `to` (both included, tables between them too)
   * in a new section, as one undo step. A table is taken whole when `from` is its first
   * paragraph or `to` its last one. `locked`: the section cannot be deleted;
   * `contentLocked`: its content cannot be edited by hand, in Word or here (code still can).
   */
  defineSection(key, { from, to }, { title, locked = false, contentLocked = false } = {}) {
    this.#require();
    checkSectionKey('defineSection', key);
    checkIndex('defineSection', { from, to });
    const info = JSON.parse(engineCall(() => this.#session.define_section(
      key, from, to, title == null ? undefined : String(title), Boolean(locked), Boolean(contentLocked), this.#selectionJSON(),
    )));
    this.#changed();
    return info;
  }

  /**
   * Replaces the content of a section with new paragraphs as one undo step, keeping the
   * section itself. `keepHeading`: a heading opening the section stays, the paragraphs go
   * after it. The caret goes after the last new paragraph.
   * @returns {{first: number, count: number}}
   */
  replaceSection(key, paragraphs, { keepHeading = false } = {}) {
    this.#require();
    checkSectionKey('replaceSection', key);
    const items = (Array.isArray(paragraphs) ? paragraphs : [paragraphs])
      .map(p => (typeof p === 'string' ? { text: p } : p))
      .map(toNewParagraph);
    if (!items.length) throw new Error('docx-editor: replaceSection needs at least one paragraph');
    const result = JSON.parse(engineCall(() => this.#session.replace_section(
      key, JSON.stringify(items), Boolean(keepHeading), this.#selectionJSON(),
    )));
    this.#changed();
    this.#canvasEditor.select(result.caret);
    this.#flash(result);
    return { first: result.first, count: result.count };
  }

  /** Removes a section's marker as one undo step; its content (and inner sections) stay */
  unwrapSection(key) {
    this.#require();
    checkSectionKey('unwrapSection', key);
    engineCall(() => this.#session.unwrap_section(key, this.#selectionJSON()));
    this.#changed();
  }

  /**
   * Sets a section's locks as one undo step: `locked`, it cannot be deleted; `contentLocked`,
   * its content cannot be edited by hand (typing, formatting, pictures), in Word or here.
   * Calls from code (`replaceSection`, `insertParagraphs`…) still change it.
   */
  setSectionLock(key, { locked = false, contentLocked = false } = {}) {
    this.#require();
    checkSectionKey('setSectionLock', key);
    const info = JSON.parse(engineCall(() => this.#session.set_section_lock(
      key, Boolean(locked), Boolean(contentLocked), this.#selectionJSON(),
    )));
    this.#changed();
    return info;
  }

  /**
   * Inserts a new section as one undo step: by default a heading with its title (Word's
   * Heading 1–3 by nesting depth) and an empty paragraph. Place it with `after`, `before`
   * or `inside` (a section key); without them it goes at the end of the document.
   * @returns {{section: object, first: number, count: number}}
   */
  insertSection(key, { title, after, before, inside, heading = true, paragraphs = [''], locked = false, contentLocked = false } = {}) {
    this.#require();
    checkSectionKey('insertSection', key);
    const given = [['after', after], ['before', before], ['inside', inside]].filter(([, v]) => v != null);
    if (given.length > 1) throw new Error('docx-editor: insertSection takes only one of after, before or inside');
    const [place, relativeTo] = given[0] ?? ['end', null];
    const sections = this.sections();
    const other = sections.find(s => s.key === relativeTo);
    const depth = place === 'end' ? 0 : place === 'inside' ? (other?.depth ?? -1) + 1 : other?.depth ?? 0;
    const [result] = this.#insertSections([{ key, title, place, relativeTo, depth, heading, paragraphs, locked, contentLocked }]);
    return result;
  }

  /**
   * Creates sections from the document's headings as one undo step: each heading of level
   * `fromLevel`–`toLevel` (1 = Heading 1) with what follows it up to the next heading of the
   * same or a higher level; lower levels nest. Keys come from the heading text
   * ("2. Caracterización" → "caracterizacion"). Headings whose part is already a section, or
   * that would cut a table or hold a section break, are reported in `skipped`.
   * @returns {{created: Array<object>, skipped: Array<{paragraph: number, text: string, reason: string}>}}
   */
  sectionsFromHeadings({ fromLevel = 1, toLevel = fromLevel } = {}) {
    this.#require();
    checkIndex('sectionsFromHeadings', { fromLevel, toLevel });
    const result = JSON.parse(engineCall(() => this.#session.sections_from_headings(fromLevel, toLevel, this.#selectionJSON())));
    if (result.created.length) this.#changed();
    return result;
  }

  /** The section schema stored in the document (`setSectionSchema()`), or null */
  sectionSchema() {
    if (!this.#session) return null;
    return JSON.parse(this.#session.section_schema_json());
  }

  /**
   * Stores in the .docx which sections this type of document must have, in order (null
   * removes it), as one undo step. It travels with the file as a custom XML part, so other
   * tools can read it. `required` defaults to true; `parent` nests a section in an earlier one.
   * @param {{type: string, title?: string, sections: Array<{key: string, title?: string, required?: boolean, parent?: string}>} | null} schema
   */
  setSectionSchema(schema) {
    this.#require();
    const payload = schemaPayload(schema);
    engineCall(() => this.#session.set_section_schema(JSON.stringify(payload), this.#selectionJSON()));
    this.#changed();
  }

  /**
   * Checks the sections against the schema: missing, repeated, empty, out of order or in the
   * wrong parent (`valid` is false), and sections the schema does not list (`unexpected`,
   * which does not make it invalid). Without a schema everything is valid.
   * @returns {{valid: boolean, type: string|null, issues: Array<{kind: 'missing'|'duplicated'|'parent'|'empty'|'order'|'unexpected', key: string, title: string, message: string}>}}
   */
  validateSections() {
    const sections = this.sections();
    const structure = this.structure();
    const isEmpty = s => sectionIsEmpty(sectionContent(s, structure, this.#elements || []));
    return validateSections(this.sectionSchema(), sections, isEmpty, this.labels.sectionIssues);
  }

  /**
   * Inserts the schema's required sections that the document lacks (with `optional`, every
   * one), each in its place by the schema's order, as one undo step: a heading with its
   * title and an empty paragraph, or `paragraphs`.
   * @returns {Array<{section: object, first: number, count: number}>}
   */
  addMissingSections({ optional = false, heading = true, paragraphs = [''] } = {}) {
    this.#require();
    const plan = planMissingSections(this.sectionSchema(), this.sections(), { optional });
    if (!plan.length) return [];
    return this.#insertSections(plan.map(p => ({ ...p, heading, paragraphs })));
  }

  /** Whether the logical sections are outlined in the margin */
  get sectionsVisible() {
    return this.#sectionLayer.visible;
  }

  /** Outlines the logical sections in the margin (a bracket and the title), or hides them */
  setShowSections(show = true) {
    this.#sectionLayer.setVisible(show);
  }

  #insertSections(list) {
    const payload = list.map(({ key, title, place, relativeTo, depth, heading, paragraphs, locked, contentLocked }) => {
      checkSectionKey('insertSection', key);
      const name = title == null ? key : String(title);
      const items = (Array.isArray(paragraphs) ? paragraphs : [paragraphs])
        .map(p => (typeof p === 'string' ? { text: p } : p))
        .map(toNewParagraph);
      const style = typeof heading === 'string' ? heading : `Heading${Math.min(3, Math.max(0, depth) + 1)}`;
      if (heading) items.unshift(toNewParagraph({ text: name, style }));
      if (!items.length) throw new Error('docx-editor: insertSection needs a heading or at least one paragraph');
      return { key, title: name, place, relativeTo, paragraphs: items, locked: Boolean(locked), contentLocked: Boolean(contentLocked) };
    });
    const results = JSON.parse(engineCall(() => this.#session.insert_sections(JSON.stringify(payload), this.#selectionJSON())));
    this.#changed();
    results.forEach(r => this.#flash(r));
    return results;
  }

  /**
   * Why a change by hand to paragraphs `from`–`to` is not allowed, or null: it touches a
   * section whose content is locked or, when `deleting`, would remove a section that cannot
   * be deleted
   */
  #lockedMessage(from, to, { deleting = true } = {}) {
    if (from >= TEXT_BOX_BASE) return null;
    for (const s of this.sections()) {
      if (s.first == null) continue;
      if (s.contentLocked && from <= s.last && to >= s.first) return this.labels.sectionContentLocked(s.title);
      const covers = from <= s.first && to >= s.last && (from < s.first || to > s.last);
      if (deleting && s.locked && covers) return this.labels.sectionLocked(s.title);
    }
    return null;
  }

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
    const blocked = this.#lockedMessage(ref.paragraph, ref.paragraph, { deleting: false });
    if (blocked) return this.#emit('message', { message: blocked, error: false });
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

  /** True while the user cannot edit the document (see `setReadOnly`) */
  get readOnly() {
    return this.#readOnly;
  }

  /**
   * Read-only mode: the user can still move the caret, select and copy text, zoom and scroll,
   * but typing, deleting, pasting, cutting, formatting shortcuts, undo/redo and moving or
   * resizing pictures are ignored, and the optional toolbar disables its editing buttons.
   * Calls from code (`replaceVariables`, `insertText`, `updateImage`…) still work, so an app
   * can fill a document that its users only view. Editable by default.
   * @param {boolean} readOnly
   */
  setReadOnly(readOnly = true) {
    const next = !!readOnly;
    if (next === this.#readOnly) return;
    this.#readOnly = next;
    this.root.classList.toggle('read-only', next);
    this.#canvasEditor.readOnlyChanged();
    this.#emit('readonlychange', { readOnly: next });
  }

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
    this.#activity.destroy();
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
    this.#sectionCache = null;
    this.#elements = JSON.parse(this.#session.get_document_elements_json());
    this.#partParagraphs = JSON.parse(this.#session.header_footer_paragraphs_json());
  }

  #changed() {
    this.#render();
    this.#emit('change', {});
  }

  /** Highlights for a moment what a call from code just changed (option `flash`) */
  #flash(target) {
    const ms = this.options.flash;
    if (!(ms > 0)) return;
    try {
      this.#activity.add('flash', target, { label: '', ttl: ms });
    } catch {
      // nothing to highlight (e.g. an empty result)
    }
  }

  /** Highlights rows `fromRow`–`toRow` of a table (one column when `col` is given) */
  #flashTable(table, fromRow, toRow, col) {
    const first = this.tables().find(t => t.index === table)?.first;
    const el = this.#elements.find(e => e.type === 'table'
      && Math.min(...(e.rich_rows || []).flatMap(r => r.cells.flatMap(c => (c.paragraphs || []).map(p => p.index)))) === first);
    const indices = (el?.rich_rows || []).slice(fromRow, toRow + 1)
      .flatMap(r => r.cells.filter((_, c) => col == null || c === col).flatMap(c => (c.paragraphs || []).map(p => p.index)));
    if (indices.length) this.#flash({ from: Math.min(...indices), to: Math.max(...indices) });
  }

  /** Texts of the body and table cell paragraphs by index, to follow edits */
  #bodyTexts() {
    const texts = [];
    for (const p of this.#paragraphs().list) if (p.index < TEXT_BOX_BASE) texts[p.index] = p.text || '';
    return texts;
  }

  /** Every paragraph (body and table cells) in document order, tagged with its container */
  #paragraphs() {
    if (this.#paragraphCache.source !== this.#elements) {
      const list = [];
      // Text box paragraphs (numbered from TEXT_BOX_BASE) belong to "tb:<anchor>:<index>"
      const addTextBoxes = p => (p.images || []).forEach((img, k) => {
        (img.text_box?.paragraphs || []).forEach(inner => {
          if (inner.index >= TEXT_BOX_BASE) list.push({ ...inner, container: `tb:${p.index}:${k}` });
        });
      });
      this.#elements.forEach(el => {
        if (el.type === 'paragraph') {
          list.push({ ...el, container: 'body' });
          addTextBoxes(el);
        } else if (el.type === 'table') {
          (el.rich_rows || []).forEach((row, r) => row.cells.forEach((cell, c) => {
            (cell.paragraphs || []).forEach(p => {
              list.push({ ...p, container: `${el.index}:${r}:${c}` });
              addTextBoxes(p);
            });
          }));
        }
      });
      this.#partParagraphs.forEach(p => list.push({ ...p, container: `hf:${p.part_slot}` }));
      list.sort((a, b) => a.index - b.index);
      this.#paragraphCache = { source: this.#elements, list, byIndex: new Map(list.map(p => [p.index, p])) };
    }
    return this.#paragraphCache;
  }

  #render({ full = false } = {}) {
    const session = this.#session;
    if (!session) return;
    try {
      // Edits made through the caret editor change the elements too; activity marks on
      // paragraph ranges follow the paragraphs that moved
      const before = this.#activity.size ? this.#bodyTexts() : null;
      this.#refreshElements();
      if (before) this.#activity.remap(textDiff(before, this.#bodyTexts()));
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
      this.#activity.paint();
      this.#sectionLayer.paint();
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
    const indices = updates.map(u => u.index);
    const blocked = this.#lockedMessage(Math.min(...indices), Math.max(...indices), { deleting: false });
    if (blocked) return this.#emit('message', { message: blocked, error: false });
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
    if (at && typeof at === 'object') {
      if (!Number.isInteger(at.paragraph) || at.paragraph < 0) throw new Error('docx-editor: at.paragraph must be a paragraph index');
      return { paragraph: at.paragraph, offset: Math.max(0, at.offset ?? 0) };
    }
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
 * @property {number} index         among the paragraph's pictures, text boxes and shapes
 * @property {'picture'|'textbox'|'shape'} kind
 * @property {number} width
 * @property {number} height
 * @property {ImageWrap} wrap
 * @property {string|null} wrapSide
 * @property {boolean} anchored     floating (any wrap but 'inline')
 * @property {{relativeTo: string, offset: number, align: string|null}|null} horizontal  null when inline
 * @property {{relativeTo: string, offset: number, align: string|null}|null} vertical
 * @property {{top: number, bottom: number, left: number, right: number}} distance  space kept from the text
 * @property {string} alt
 * @property {string} text          text of a text box (paragraphs separated by "\n")
 * @property {{geometry: string, fill: string|null, stroke: string|null, strokeWidth: number}|null} shape  outline and fill of a text box or shape
 * @property {number} textOffset    character offset in the paragraph where it sits
 * @property {{page: number, x: number, y: number, width: number, height: number}|null} bounds  where it is drawn
 */
const hexColor = c => (c ? `#${c}` : null);

function imageInfo(raw) {
  const axis = (relativeTo, offset, align) => (raw.anchored ? { relativeTo, offset, align: align ?? null } : null);
  return {
    paragraph: raw.paragraph,
    index: raw.index,
    kind: raw.kind || 'picture',
    width: raw.width,
    height: raw.height,
    wrap: raw.wrap || (raw.anchored ? (raw.behind_text ? 'behind' : 'inFront') : 'inline'),
    wrapSide: raw.wrap_side || null,
    anchored: raw.anchored,
    horizontal: axis(raw.h_relative, raw.h_offset, raw.h_align),
    vertical: axis(raw.v_relative, raw.v_offset, raw.v_align),
    distance: { top: raw.dist_top, bottom: raw.dist_bottom, left: raw.dist_left, right: raw.dist_right },
    alt: raw.alt || '',
    text: raw.text || '',
    shape: raw.shape ? { geometry: raw.shape.geometry, fill: hexColor(raw.shape.fill), stroke: hexColor(raw.shape.stroke), strokeWidth: raw.shape.stroke_width } : null,
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

/** A cell for the engine: its text, or `{ text, image: { data: base64 } }` */
function tableCell(cell) {
  if (cell == null) return '';
  if (typeof cell !== 'object') return String(cell);
  const { text, image, width, height, alt } = cell;
  if (image == null) return text == null ? '' : String(text);
  if (typeof Blob !== 'undefined' && image instanceof Blob) {
    throw new Error('docx-editor: insertTable needs image bytes or a data: URL; read the Blob first with await blob.arrayBuffer()');
  }
  const data = typeof image === 'string' ? image : base64(image instanceof ArrayBuffer ? new Uint8Array(image) : image);
  return { text: text == null ? undefined : String(text), image: { data, width, height, alt } };
}

function base64(bytes) {
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
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

/** A paragraph for the engine (snake_case, optional Word style) */
function toNewParagraph({ text = '', bold, italic, underline, fontSize, align, style }) {
  return { text: String(text), bold, italic, underline, font_size: fontSize, align, style };
}
