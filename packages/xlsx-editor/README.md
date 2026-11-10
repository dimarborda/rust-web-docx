# @dimarborda/xlsx-editor

View and edit Excel (`.xlsx`) workbooks in the browser, drawn on canvas the way Excel shows them, and work with them from Node.js. A Rust and WebAssembly engine reads the file, recalculates formulas and saves losslessly: what it does not touch stays byte by byte. It works with React, Vue, Svelte, Angular or plain HTML, and in desktop apps such as Tauri. It needs no server: the library sends your files nowhere.

For AI agents, `wb.area()` gives a table's columns, types and totals without reading the whole sheet, and a change that would break something is refused with the reason, leaving the workbook as it was.

**Full documentation on GitHub:** [English](https://github.com/dimarborda/rust-web-docx/blob/master/docs/xlsx-editor/en/component.md) · [Español](https://github.com/dimarborda/rust-web-docx/blob/master/docs/xlsx-editor/es/component.md)

## Install

```bash
npm install @dimarborda/xlsx-editor
```

```html
<xlsx-editor src="/report.xlsx" locale="en" style="height: 80vh"></xlsx-editor>
<script type="module">
  import '@dimarborda/xlsx-editor';
  import '@dimarborda/xlsx-editor/fonts';      // optional: substitutes for Calibri, Arial…
  import '@dimarborda/xlsx-editor/style.css';
</script>
```

With WebAssembly, the production `script-src` needs `'wasm-unsafe-eval'`.

## Features

- **Faithful display:** fonts, colors, borders, number formats, merged cells, frozen panes and hidden lines. Large sheets scroll smoothly.
- **Editing:** values and formulas, copy and paste with Excel and Google Sheets, undo, and a status bar with average, count and sum.
- **Rows and columns:** move, insert, delete, resize, autofit, hide and show, updating formulas, names, charts and tables across the workbook.
- **Productivity:** fill handle with series, autosum, find and replace, keyboard navigation and a context menu.
- **Formulas:** about 90 functions recalculated instantly; for anything else, Excel's saved value stays and Excel recalculates when opening the file.
- **Headless engine:** the same API in Node.js.

Charts, images, conditional formats and comments are kept in the file but not drawn yet.

## Use it from code

```js
import { readFile, writeFile } from 'node:fs/promises';
import { openWorkbook } from '@dimarborda/xlsx-editor/engine';

const wb = await openWorkbook(await readFile('template.xlsx'));
const sheet = wb.sheet('Sales');

wb.fillTemplate({ CLIENT: 'Acme Inc.' });
sheet.setValues('A2', [['Keyboard', 3, 29.9], ['Mouse', 5, 12.5]]);
sheet.setValue('D2', { formula: '=B2*C2' });
sheet.insertRows(2);
sheet.fill('D3', 'D3:D4');

await writeFile('output.xlsx', wb.toBytes());
```

In the browser, `element.editor` is the editor and `editor.workbook` this same API:

```js
document.querySelector('xlsx-editor').addEventListener('ready', async ({ detail: { editor } }) => {
  await editor.open(file);
  editor.moveRows(4, 2, 9);
  editor.download('report.xlsx');
});
```

## Documentation

| Guide | English | Español |
| :--- | :--- | :--- |
| Component: keyboard, attributes, events, editor API | [component](https://github.com/dimarborda/rust-web-docx/blob/master/docs/xlsx-editor/en/component.md) | [component](https://github.com/dimarborda/rust-web-docx/blob/master/docs/xlsx-editor/es/component.md) |
| Move, insert, delete, fill, find and replace | [operations](https://github.com/dimarborda/rust-web-docx/blob/master/docs/xlsx-editor/en/operations.md) | [operations](https://github.com/dimarborda/rust-web-docx/blob/master/docs/xlsx-editor/es/operations.md) |
| Engine: full API, formulas, formats, dates | [engine](https://github.com/dimarborda/rust-web-docx/blob/master/docs/xlsx-editor/en/engine.md) | [engine](https://github.com/dimarborda/rust-web-docx/blob/master/docs/xlsx-editor/es/engine.md) |
| Areas and context for AI agents | [agents](https://github.com/dimarborda/rust-web-docx/blob/master/docs/xlsx-editor/en/agents.md) | [agents](https://github.com/dimarborda/rust-web-docx/blob/master/docs/xlsx-editor/es/agents.md) |

## For AI coding agents

The package ships [`AGENTS.md`](https://github.com/dimarborda/rust-web-docx/blob/master/packages/xlsx-editor/AGENTS.md): a compact guide to the API, its conventions and pitfalls, written for coding agents. To have your agent load it in a project that uses the package, reference it from your `AGENTS.md` or `CLAUDE.md`:

```md
@node_modules/@dimarborda/xlsx-editor/AGENTS.md
```

With the [Context7](https://context7.com) MCP server, ask for the library by its ID, `/dimarborda/rust-web-docx`, rather than by name: other projects are also called `docx-editor`.

Both libraries are also summarized at [llms.txt](https://rust-web-docx.dimarborda.workers.dev/llms.txt) and in full at [llms-full.txt](https://rust-web-docx.dimarborda.workers.dev/llms-full.txt).

## License

MIT
