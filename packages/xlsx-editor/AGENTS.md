# @dimarborda/xlsx-editor — guide for AI coding agents

Read this before writing code that uses this package. It is the short, complete map of the
library; exact signatures and JSDoc are in `src/engine.d.ts` (workbook API) and `src/index.d.ts`
(browser editor) next to this file.

## What it is

A Rust + WebAssembly engine for Excel (`.xlsx`) workbooks, with two entry points:

| Import | Where | What |
| :--- | :--- | :--- |
| `@dimarborda/xlsx-editor/engine` | Node.js ≥ 20 and browser | Headless `Workbook` / `Sheet` API: read, write, recalculate, format, restructure |
| `@dimarborda/xlsx-editor` | Browser | `<xlsx-editor>` element and `XlsxEditor` class (canvas grid), plus everything from `/engine` |

Saving is lossless: only the parts and cells that changed are rewritten; charts, images,
conditional formats, validations and the rest stay byte for byte. No server, no network calls.

## Headless engine (scripts, servers, agent tools)

```js
import { readFile, writeFile } from 'node:fs/promises';
import { openWorkbook, createWorkbook } from '@dimarborda/xlsx-editor/engine';

const wb = await openWorkbook(await readFile('template.xlsx'));   // or await createWorkbook()
const sheet = wb.sheet('Sales');                                   // name (ignores case) or 0-based tab index

sheet.getValue('B10');                          // formula result, already recalculated
sheet.getCell('D2');                            // { ref, type, value, formula? }
sheet.getValues('A1:C3');                       // matrix; empty cells are null

sheet.setValue('B3', 42);                       // text, number, boolean, Date, null (clears value, keeps format)
sheet.setValue('D2', { formula: '=B2*C2' });    // formulas: English syntax, leading '='
sheet.setValues('A2', [['Keyboard', 3, 29.9], ['Mouse', 5, 12.5]]); // undefined leaves a cell as is
wb.fillTemplate({ CLIENT: 'Acme Inc.' });       // {{CLIENT}}, {{ CLIENT }}, {{CLIENT: hint}}

await writeFile('output.xlsx', wb.toBytes());
wb.close();                                     // frees WASM memory; the workbook is unusable afterwards
```

In Node the engine loads by itself. In the browser bundlers serve the `.wasm`; otherwise call
`initEngine(url)` first. In production with a CSP, `script-src` needs `'wasm-unsafe-eval'`.

## Understand a sheet without reading all of it

Do not dump whole sheets into a prompt. Ask for areas, then the context of the one you need:

```js
wb.areas();                      // [{ kind, name, sheetName, ref, address, title, columns, … }]
const ctx = wb.area('Sales');    // table / defined name, 'Sheet!A1:H40', or a range with { sheet }
ctx.columnsInfo;                 // per column: name, type, filled, numbers { sum, min, max, average }, exampleFormula
ctx.markdown;                    // header + first rows (maxRows, default 30) as a Markdown table
ctx.truncated; ctx.rowCount;
sheet.regionAt('B5');            // the block Ctrl+A would select around a cell
```

`kind` is `table` (Excel tables), `name` (defined names), `printArea`, `filter` or `region`
(adjacent filled cells, with detected title and header). `address` works as a target anywhere.

## Addresses and numbering

- Cells and ranges are A1 strings: `'B3'`, `'A1:C4'`, `'B:D'`, `'3:5'`.
- **Rows and column numbers are 1-based, as in Excel** (`insertRows(2)` inserts before row 2;
  column `2` is `B`). Columns also accept letters. **Sheet indexes are 0-based.**
- Move/insert/delete use numbering *as it is now*; afterwards re-read areas or cells rather than
  reusing old references.

## Formulas

- Written in **English syntax** (`SUM`, `IF`, `,` between arguments; `;` accepted). Spanish names
  like `SUMA` are not translated.
- About 90 functions are evaluated (math, conditional `*IFS`, logic, `VLOOKUP`/`XLOOKUP`/`INDEX`/
  `MATCH`, text, dates). Anything else (`OFFSET`, `INDIRECT`, external links, structured table
  references, array/spill formulas, circular references) keeps the value Excel saved, and Excel
  recalculates on open (`fullCalcOnLoad` is set). So `getValue` may return a stale value for those.
- Every write recalculates dependents; `wb.recalculate()` refreshes a freshly opened workbook.

## Dates

Excel dates have no time zone: they map to the **UTC** fields of a `Date`. Write
`new Date(Date.UTC(2024, 0, 15))`, read with `getUTCFullYear()` / `toISOString()`. Writing a `Date`
gives the cell a date format if it had none.

## Format and structure

```js
sheet.setFormat('A1:D1', { bold: true, fill: '#DDEBF7', border: { style: 'thin', edges: 'outline' } });
sheet.setFormat('D2:D50', { numberFormat: '#,##0.00', align: 'right' });
sheet.setColumnWidth('B:D', 14);          // characters
sheet.setRowHeight('3:5', 30);            // points; null = automatic
sheet.mergeCells('A1:H1');
sheet.insertRows(2, 3);  sheet.deleteRows(10, 2);  sheet.moveRows(4, 2, 9);
sheet.insertColumns('C'); sheet.moveColumns('B', 1, 'E');
sheet.fill('D2', 'D2:D20');               // fill handle: formulas shift, series continue
sheet.setRowsHidden('3:5', true);
wb.addSheet('Notes'); wb.renameSheet('Sheet1', 'Data');   // rename updates references
```

`setFormat` changes only the given properties. Structural changes update formulas, names, charts,
tables, validations and conditional formats across the whole workbook, like Excel.

**Refusals are safe:** an operation that would tear merged cells, a multi-cell array formula, a
table or a filter, or touch a sheet with pivot tables, **throws and changes nothing**. Catch the
error and tell the user why (engine messages are in Spanish).

## Markers

Same lexicon as `@dimarborda/docx-editor`: `{{KEY}}` / `{key}` are `field`s an automation fills;
`[…]`, `[____]`, `[Complete: scope]` are `pending` content to write.

```js
for (const m of wb.markers({ kind: 'pending' })) {         // m: { sheet, sheetName, ref, row, col, raw, key, hint, offset, length }
  const s = wb.sheet(m.sheet);                              // m.sheet is the 0-based tab index
  const text = String(s.getValue(m.ref));
  s.setValue(m.ref, text.replace(m.raw, await draft(m.hint))); // the marker may be only part of the cell
}
```

For `{{FIELD}}`s, `wb.fillTemplate(values)` is simpler and keeps the rest of each text.

## Undo history (engine)

`wb.checkpoint(selection?)` opens an undo step; edits until the next checkpoint undo together.
`wb.undo()` / `wb.redo()` return `{ selection, parts }` or `null`.

## Browser editor

```html
<xlsx-editor src="/report.xlsx" locale="en" style="height: 80vh"></xlsx-editor>
<script type="module">
  import '@dimarborda/xlsx-editor';
  import '@dimarborda/xlsx-editor/style.css';
  import '@dimarborda/xlsx-editor/fonts';     // optional font substitutes
</script>
```

```js
el.addEventListener('ready', async ({ detail: { editor } }) => {
  await editor.open(file);                  // URL, File/Blob, ArrayBuffer, Uint8Array
  editor.setSheet('Summary'); editor.select('B3:D7');

  // Changes through the workbook API: checkpoint first (undo step), reload after (redraw)
  editor.workbook.checkpoint();
  editor.workbook.sheet('Summary').setValues('B2', [[1200, 0.25]]);
  editor.reload();

  editor.moveRows(4, 2, 9);                 // editor methods: one undo step each, false if refused
  editor.selectionContext();                // AreaContext of what the user selected
  editor.download('report.xlsx');           // or editor.toBytes()
});
el.addEventListener('change', ({ detail }) => save(detail)); // { sheet, range, source, cells }
```

Without the element: `const editor = await XlsxEditor.create(container, { locale: 'en' })`. The
container needs a height. Events: `ready`, `load`, `sheetchange`, `selectionchange`, `change`,
`error`. Editor methods return `false` on refusal and emit `error` with the reason.

## Do and don't

- Do use `areas()` / `area()` to get structure, types and totals instead of reading every cell.
- Do write numbers as numbers and dates as `Date` — not preformatted strings — so formulas and
  formats keep working. `parseInput(text, locale)` converts user-typed text the way the grid does.
- Do call `editor.reload()` after changing `editor.workbook` from code, and `checkpoint()` before.
- Don't unzip or rewrite the `.xlsx` XML yourself: the engine keeps it lossless.
- Don't expect charts, images, conditional formats or comments to be drawn in the grid: they are
  preserved in the file but not rendered yet.
- A `{{FIELD}}` split across differently formatted runs inside one cell is not detected.

## More

Context7 library ID: `/dimarborda/rust-web-docx` (both packages; use the ID, other libraries share the `docx-editor` name).

Full docs: https://github.com/dimarborda/rust-web-docx/tree/master/docs/xlsx-editor/en
(component, operations, engine, agents).
