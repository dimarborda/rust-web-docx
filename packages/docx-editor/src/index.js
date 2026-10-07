// @dimarborda/docx-editor: Word (.docx) editor for the browser, powered by a Rust/WebAssembly
// layout engine. Importing this module registers <docx-editor> and <docx-toolbar>.
import { defineDocxElements } from './elements.js';

export { DocxEditor } from './docx_editor.js';
export { createDocxToolbar, DEFAULT_TOOLBAR_ITEMS } from './toolbar.js';
export { DocxEditorElement, DocxToolbarElement, defineDocxElements } from './elements.js';
export { initEngine, sampleDocx } from './engine.js';
export { registerFont, loadStoredFonts, listStoredFonts, removeStoredFont } from './fonts_registry.js';

if (typeof customElements !== 'undefined') defineDocxElements();
