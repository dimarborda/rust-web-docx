# @dimarborda/docx-editor

[![npm](https://img.shields.io/npm/v/@dimarborda/docx-editor?color=cb3837&label=npm)](https://www.npmjs.com/package/@dimarborda/docx-editor)

Word (`.docx`) editor for the browser. A Rust and WebAssembly engine reads the document, lays out the pages like Word and applies edits losslessly. It works with any frontend and in desktop apps such as Tauri. No server needed: documents never leave the device.

[Demo](https://rust-web-docx.dimarborda.workers.dev/?demo) · **Full documentation on GitHub:** [English](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/en/api.md) · [Español](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/es/api.md)

## Install

```bash
npm install @dimarborda/docx-editor
```

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

The container needs a height. TypeScript types are included. If your app sends a `Content-Security-Policy`, the production `script-src` needs `'wasm-unsafe-eval'`.

## Use it from code

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

## Features

- **Templates and markers:** list and fill `{{KEY}}`, `{key}` and pending `[…]` content, keeping the formatting.
- **AI agents:** show where an agent is working and highlight its edits without moving the user's caret.
- **Content:** paragraphs, tables, images, text boxes, logical sections and existing tables, from code.
- **Formatting:** styles, fonts, character formatting, alignment, spacing, indent and lists, written the way Word does.
- **Export:** lossless `.docx` and vector PDF with working links.
- **View:** read-only mode, zoom, watermark and an optional toolbar.

## Documentation

| Guide | English | Español |
| :--- | :--- | :--- |
| API: every method, option and event | [api](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/en/api.md) | [api](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/es/api.md) |
| Guides: markers, agents, sections, tables, PDF, images | [guides](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/en/guides.md) | [guides](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/es/guides.md) |
| React, Next.js, Tauri, fonts, styles, bundlers | [integrations](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/en/integrations.md) | [integrations](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/es/integrations.md) |
| Security and limitations | [security](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/en/security.md) | [security](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/es/security.md) |

Documents never leave the browser and their content is never executed. [Report a vulnerability](https://github.com/dimarborda/rust-web-docx/blob/master/SECURITY.md).

## License

MIT
