// @dimarborda/xlsx-editor: Excel (.xlsx) editor for the browser, drawn on canvas, powered by the
// Rust/WebAssembly engine in engine.js. Importing this module registers <xlsx-editor>.
import { defineXlsxElements } from './elements.js';

export { XlsxEditor } from './xlsx_editor.js';
export { XlsxEditorElement, defineXlsxElements } from './elements.js';
export { parseInput } from './input.js';
// The engine API too, so one import gives both; in Node import '@dimarborda/xlsx-editor/engine'
export * from './engine.js';

if (typeof customElements !== 'undefined') defineXlsxElements();
