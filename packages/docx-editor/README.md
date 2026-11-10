# @dimarborda/docx-editor

[![npm](https://img.shields.io/npm/v/@dimarborda/docx-editor?color=cb3837&label=npm)](https://www.npmjs.com/package/@dimarborda/docx-editor)

Word (`.docx`) editor for the browser. A Rust and WebAssembly engine reads the document, lays out the pages like Word and applies edits losslessly. It works with React, Vue, Svelte, Angular or plain HTML, and in desktop apps such as Tauri. It needs no server: the library sends your documents nowhere.

For AI agents, `markers()` lists what is left to fill in, and edits made through the API are highlighted without moving the user's caret.

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
- **Agent activity:** `markActivity()` shows where an agent is working.
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

The library sends your documents nowhere and never executes their content. [Report a vulnerability](https://github.com/dimarborda/rust-web-docx/blob/master/SECURITY.md).

## For AI coding agents

The package ships [`AGENTS.md`](https://github.com/dimarborda/rust-web-docx/blob/master/packages/docx-editor/AGENTS.md): a compact guide to the API, its conventions and pitfalls, written for coding agents. To have your agent load it in a project that uses the package, reference it from your `AGENTS.md` or `CLAUDE.md`:

```md
@node_modules/@dimarborda/docx-editor/AGENTS.md
```

With the [Context7](https://context7.com) MCP server, ask for the library by its ID, `/dimarborda/rust-web-docx`, rather than by name: other projects are also called `docx-editor`.

Both libraries are also summarized at [llms.txt](https://rust-web-docx.dimarborda.workers.dev/llms.txt) and in full at [llms-full.txt](https://rust-web-docx.dimarborda.workers.dev/llms-full.txt).

## License

MIT
