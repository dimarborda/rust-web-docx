// DocxEditor: a Word document editor mounted in any element. Rust (WebAssembly) parses the
// .docx, lays out the pages and applies lossless edits; this class renders the pages to
// canvases, wires the caret/selection editor and exposes a small API plus DOM events.
import { DocxSession, engineReady, initEngine } from './engine.js';
import { createCanvasEditor } from './canvas_editor.js';
import { onFontsChanged } from './fonts_registry.js';
import { disableLigatures, drawPageItems, layoutFontFaces, measureText, usesSubstitute } from './render.js';

const LABELS = {
  es: {
    editor: 'Editor del documento',
    page: (n, total) => `Página ${n} de ${total}`,
    nothingToUndo: 'No hay nada para deshacer.',
    nothingToRedo: 'No hay nada para rehacer.',
    clickToFormat: 'Haz clic en el texto para aplicar formato.',
  },
  en: {
    editor: 'Document editor',
    page: (n, total) => `Page ${n} of ${total}`,
    nothingToUndo: 'Nothing to undo.',
    nothingToRedo: 'Nothing to redo.',
    clickToFormat: 'Click in the text to apply formatting.',
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
  #onKeyDown;

  /** Use `DocxEditor.create()` unless the engine is already initialized (`initEngine()`) */
  constructor(container, options = {}) {
    super();
    if (!engineReady()) throw new Error('docx-editor: await initEngine() or use DocxEditor.create()');
    this.root = container;
    this.options = { zoom: 1, gridlines: true, pageLabels: true, locale: 'es', ...options };
    this.labels = { ...(LABELS[this.options.locale] || LABELS.en), ...options.labels };
    this.#zoom = this.options.zoom;

    container.classList.add('docx-editor');
    this.#pages = document.createElement('div');
    this.#pages.className = 'docx-editor-pages';
    container.appendChild(this.#pages);

    this.#canvasEditor = createCanvasEditor({
      root: container,
      labels: this.labels,
      session: () => this.#session,
      paragraphs: () => this.#paragraphs(),
      zoom: () => this.#zoom,
      measure: measureText,
      documentChanged: () => this.#changed(),
      selectionChanged: selection => this.#emit('selectionchange', { selection, format: this.#formatAt(selection) }),
      hint: message => this.#emit('message', { message, error: false }),
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
    container.addEventListener('keydown', this.#onKeyDown);

    // Uploaded fonts change text widths: measure and lay out again
    this.#unsubscribeFonts = onFontsChanged(() => {
      if (!this.#session) return;
      this.#session.reset_measurements();
      this.#render({ full: true });
    });
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

  /** Appends a table at the end of the document */
  insertTable(rows, cols, headers = Array.from({ length: cols }, (_, i) => `${i + 1}`)) {
    this.#require();
    this.#session.add_table(rows, cols, JSON.stringify(headers));
    this.#changed();
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

  destroy() {
    this.close();
    this.#canvasEditor.destroy();
    this.#unsubscribeFonts();
    this.root.removeEventListener('keydown', this.#onKeyDown);
    this.#pages.remove();
    this.root.classList.remove('docx-editor');
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
    Promise.all(pending.map(face => document.fonts.load(face).catch(() => []))).then(results => {
      if (results.some(loaded => loaded.length > 0) && this.#session) {
        this.#session.reset_measurements();
        this.#render({ full: true });
      }
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
