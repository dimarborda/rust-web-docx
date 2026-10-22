# @dimarborda/docx-editor

[![npm](https://img.shields.io/npm/v/@dimarborda/docx-editor?color=cb3837&label=npm)](https://www.npmjs.com/package/@dimarborda/docx-editor)

Word (`.docx`) editor for the browser. The engine is written in Rust and compiled to WebAssembly: it reads the document, lays out the pages like Word and applies edits losslessly. Pages are drawn on canvas, with their own caret and selection.

Works with any frontend (no framework, React, Vue, Svelte, Angular) and in desktop apps with a webview, such as Tauri. No server needed: documents never leave the device.

[Demo](https://rust-web-docx.dimarborda.workers.dev/?demo) · [Full documentation (Spanish)](https://github.com/dimarborda/rust-web-docx/tree/master/docs/docx-editor) · [Repository](https://github.com/dimarborda/rust-web-docx)

```bash
npm install @dimarborda/docx-editor
```

## Web component

```html
<docx-toolbar for="doc"></docx-toolbar>
<docx-editor id="doc" src="/templates/contract.docx" style="height: 80vh"></docx-editor>

<script type="module">
  import '@dimarborda/docx-editor';          // registers <docx-editor> and <docx-toolbar>
  import '@dimarborda/docx-editor/style.css';
  import '@dimarborda/docx-editor/fonts';    // optional: stand-ins for Calibri, Cambria, Arial…

  const editor = await document.getElementById('doc').whenReady();
  editor.replaceVariables({ CLIENT: 'Acme Inc.', DATE: 'October 7, 2026' });
</script>
```

## JavaScript

```js
import { DocxEditor } from '@dimarborda/docx-editor';
import '@dimarborda/docx-editor/style.css';

const editor = await DocxEditor.create(document.querySelector('#editor'), { locale: 'en' });

await editor.open(file);                       // File, Blob, ArrayBuffer or Uint8Array
editor.replaceVariables({ CLIENT: 'Acme' });   // fills {{CLIENT}}, {CLIENT}…
editor.insertParagraphs(['The parties agree as follows.'], { at: 'end' });
await editor.downloadPdf('contract.pdf');      // vector PDF with selectable text
const bytes = editor.save();                   // Uint8Array of the edited .docx
```

The container needs a height (for example `height: 80vh`). TypeScript types included.

## What you can do

- **Templates and markers**: `markers()` lists `{{KEY}}` / `{key}` fields and pending `[…]` content with their position; `fillMarker()` and `replaceVariables()` fill them keeping the formatting. [Guide](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/guias.md#marcadores-campos-y-pendientes)
- **Agents**: `markActivity()` shows where an agent is working, and edits made through the API are highlighted for a moment without moving the user's caret. [Guide](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/guias.md#actividad-dónde-está-trabajando-el-agente)
- **Content**: insert paragraphs with heading styles, tables and images (inline or floating), and read the document structure.
- **Logical sections and existing tables**: read and replace parts of the document by key, and write cells or rows copying their formatting. [Guides](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/guias.md)
- **Export**: lossless `.docx` and vector PDF.
- **View**: read-only mode, zoom, watermark and an optional toolbar.

Every method, option and event is in the [API reference](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/api.md). For React, Next.js, Tauri, fonts, styles and bundlers, see the [integrations](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/integraciones.md).

## Content Security Policy

If your app sends a `Content-Security-Policy`, add `'wasm-unsafe-eval'` to the production `script-src`, or the engine will not start. [Details](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/integraciones.md#content-security-policy)

## Security

Documents never leave the browser, their content is never executed (it is drawn on canvas) and the reader rejects external XML entities and ZIP bombs. [Security and limitations](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/seguridad.md) · [Report a vulnerability](https://github.com/dimarborda/rust-web-docx/blob/master/SECURITY.md)

## License

MIT
