// Custom element for HTML-first and framework projects (React, Vue, Svelte, Angular...):
//
//   <xlsx-editor src="/plantilla.xlsx" style="height: 80vh"></xlsx-editor>
//   <xlsx-editor blank locale="es" zoom="1.25" tabs="false"></xlsx-editor>
//   <xlsx-editor src="/informe.xlsx" readonly></xlsx-editor>
//
// `element.editor` is the XlsxEditor (null until the `ready` event). Editor events (load,
// sheetchange, selectionchange, change, error) are re-dispatched on the element; a file that
// cannot be opened dispatches `error` too.
import { XlsxEditor } from './xlsx_editor.js';

const FORWARDED = ['load', 'sheetchange', 'selectionchange', 'change', 'error'];

// Outside the browser (SSR) there is no HTMLElement: the class still loads, and the element is
// only registered where customElements exists (see index.js)
const ElementBase = typeof HTMLElement === 'undefined' ? class {} : HTMLElement;

export class XlsxEditorElement extends ElementBase {
  static observedAttributes = ['src', 'zoom', 'sheet', 'readonly'];

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
    if (name === 'src' && value) this.#open(value);
    if (name === 'zoom' && value) this.editor.setZoom(Number(value));
    if (name === 'readonly') this.editor.readOnly = value !== null;
    if (name === 'sheet' && value) this.#run(() => this.editor.setSheet(/^\d+$/.test(value) ? Number(value) : value));
  }

  async #create() {
    const editor = await XlsxEditor.create(this, {
      locale: this.getAttribute('locale') || undefined,
      zoom: Number(this.getAttribute('zoom')) || 1,
      tabs: this.getAttribute('tabs') !== 'false',
      formulaBar: this.getAttribute('formula-bar') !== 'false',
      readOnly: this.hasAttribute('readonly'),
      wasmUrl: this.getAttribute('wasm-url') || undefined,
    });
    FORWARDED.forEach(type => editor.addEventListener(type, e => {
      this.dispatchEvent(new CustomEvent(type, { detail: e.detail, bubbles: true }));
      if (type === 'error' && !this.hasAttribute('quiet')) console.error('xlsx-editor:', e.detail.error);
    }));
    this.editor = editor;
    this.dispatchEvent(new CustomEvent('ready', { detail: { editor } }));
    const src = this.getAttribute('src');
    if (src) await this.#open(src);
    else if (this.hasAttribute('blank')) await editor.createBlank();
  }

  async #open(src) {
    await this.#run(async () => {
      await this.editor.open(src);
      const sheet = this.getAttribute('sheet');
      if (sheet) this.editor.setSheet(/^\d+$/.test(sheet) ? Number(sheet) : sheet);
    });
  }

  async #run(fn) {
    try {
      await fn();
    } catch (error) {
      this.dispatchEvent(new CustomEvent('error', { detail: { error }, bubbles: true }));
      if (!this.hasAttribute('quiet')) console.error('xlsx-editor:', error);
    }
  }
}

export function defineXlsxElements() {
  if (!customElements.get('xlsx-editor')) customElements.define('xlsx-editor', XlsxEditorElement);
}
