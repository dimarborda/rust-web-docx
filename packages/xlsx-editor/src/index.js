// @dimarborda/xlsx-editor: Excel (.xlsx) editor for the browser, drawn on canvas, powered by the
// Rust/WebAssembly engine of @dimarborda/xlsx. Importing this module registers <xlsx-editor>.
import { defineXlsxElements } from './elements.js';

export { XlsxEditor } from './xlsx_editor.js';
export { XlsxEditorElement, defineXlsxElements } from './elements.js';
export { parseInput } from './input.js';
export { initEngine } from '@dimarborda/xlsx';

if (typeof customElements !== 'undefined') defineXlsxElements();
