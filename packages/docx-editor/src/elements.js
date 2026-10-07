// Custom elements for HTML-first and framework projects (React, Vue, Svelte, Angular...):
//
//   <docx-editor id="doc" src="/plantilla.docx" locale="es"></docx-editor>
//   <docx-toolbar for="doc"></docx-toolbar>
//
// `element.editor` is the DocxEditor (null until the `ready` event). Editor events (load,
// change, selectionchange, message) are re-dispatched on the element.
import { DocxEditor } from './docx_editor.js';
import { createDocxToolbar } from './toolbar.js';

const FORWARDED = ['load', 'change', 'selectionchange', 'message'];

export class DocxEditorElement extends HTMLElement {
  static observedAttributes = ['src', 'zoom'];

  editor = null;
  #creating = null;

  connectedCallback() {
    this.#creating ??= this.#create();
  }

  disconnectedCallback() {
    // Moving the element around the DOM disconnects and reconnects it in the same task
    queueMicrotask(() => {
      if (this.isConnected || !this.editor) return;
      this.editor.destroy();
      this.editor = null;
      this.#creating = null;
    });
  }

  attributeChangedCallback(name, _old, value) {
    if (!this.editor) return;
    if (name === 'src' && value) this.#openUrl(value);
    if (name === 'zoom' && value) this.editor.setZoom(Number(value));
  }

  async #create() {
    const editor = await DocxEditor.create(this, {
      locale: this.getAttribute('locale') || undefined,
      zoom: Number(this.getAttribute('zoom')) || 1,
      gridlines: this.getAttribute('gridlines') !== 'false',
      pageLabels: this.getAttribute('page-labels') !== 'false',
      wasmUrl: this.getAttribute('wasm-url') || undefined,
    });
    FORWARDED.forEach(type => editor.addEventListener(type, e => {
      this.dispatchEvent(new CustomEvent(type, { detail: e.detail, bubbles: true }));
    }));
    this.editor = editor;
    this.dispatchEvent(new CustomEvent('ready', { detail: { editor } }));
    const src = this.getAttribute('src');
    if (src) await this.#openUrl(src);
    else if (this.hasAttribute('sample')) editor.openSample();
  }

  async #openUrl(url) {
    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
      await this.editor.open(await res.arrayBuffer(), { fileName: decodeURIComponent(url.split('/').pop().split('?')[0]) });
    } catch (err) {
      this.dispatchEvent(new CustomEvent('message', { detail: { message: `${url}: ${err.message}`, error: true }, bubbles: true }));
    }
  }

  /** Resolves with the DocxEditor once it is ready */
  async whenReady() {
    if (!this.editor) await (this.#creating ??= this.#create());
    return this.editor;
  }
}

export class DocxToolbarElement extends HTMLElement {
  #toolbar = null;

  async connectedCallback() {
    const target = this.getAttribute('for')
      ? document.getElementById(this.getAttribute('for'))
      : this.closest('docx-editor') || this.parentElement?.querySelector('docx-editor');
    if (!target) return console.warn('docx-toolbar: no <docx-editor> found; set for="editor-id"');
    await customElements.whenDefined('docx-editor');
    const editor = await target.whenReady();
    if (!this.isConnected || this.#toolbar) return;
    const items = this.getAttribute('items')?.split(/[\s,]+/).filter(Boolean);
    this.#toolbar = createDocxToolbar(editor, this, { ...(items && { items }), locale: this.getAttribute('locale') || undefined });
  }

  disconnectedCallback() {
    this.#toolbar?.destroy();
    this.#toolbar = null;
  }
}

/** Registers <docx-editor> and <docx-toolbar> (once per page) */
export function defineDocxElements() {
  if (!customElements.get('docx-editor')) customElements.define('docx-editor', DocxEditorElement);
  if (!customElements.get('docx-toolbar')) customElements.define('docx-toolbar', DocxToolbarElement);
}
