[← Index](../README.md) · **English** · [Español](../es/engine.md)

# The headless engine

`@dimarborda/xlsx-editor/engine` is the same engine without grid or DOM: it reads, writes, recalculates and formats workbooks in Node.js (scripts, servers, agent tools) or in the browser.

```bash
npm install @dimarborda/xlsx-editor
```

```js
import { readFile, writeFile } from 'node:fs/promises';
import { openWorkbook } from '@dimarborda/xlsx-editor/engine';

const wb = await openWorkbook(await readFile('template.xlsx'));

// {{FIELD}} variables in any text of the workbook
wb.fillTemplate({ CLIENT: 'Acme Inc.', DATE: '01/15/2024' });

// Rows of data from a cell
const sheet = wb.sheet('Sales');
sheet.setValues('A2', [
  ['Keyboard', 3, 29.9, new Date(Date.UTC(2024, 0, 15))],
  ['Mouse', 5, 12.5, new Date(Date.UTC(2024, 0, 16))],
]);

console.log(sheet.getValue('B10'));   // result of a formula, already recalculated
await writeFile('output.xlsx', wb.toBytes());
```

In Node the engine loads by itself. In the browser, bundlers (Vite, webpack 5, Rollup) serve the `.wasm` next to the JavaScript; if you serve it from elsewhere, call `initEngine(url)` first. The package's main entry also exports all this, so in the browser `import { openWorkbook } from '@dimarborda/xlsx-editor'` is enough; `editor.workbook` is a `Workbook` of this same API.

## API

| | |
|---|---|
| `openWorkbook(bytes)` / `createWorkbook()` | Opens an `.xlsx` or creates a workbook with an empty sheet. |
| `wb.sheetNames` | Names of the sheets in tab order. |
| `wb.sheet(nameOrIndex)` / `wb.sheets` | One worksheet (by name, ignoring case, or by position from 0) or all. Chart sheets have no cells. |
| `sheet.getCell('B3')` | `{ ref, type, value, formula? }`. `type` is `empty`, `number`, `string`, `boolean`, `date` or `error` (`value` is then the code, e.g. `#DIV/0!`). In formulas, `value` is the result and `formula` the text without `=` (also in the cells of a shared formula). |
| `sheet.getValue('B3')` / `sheet.getValues('A1:C3')` | Only the values; empty cells give `null`. |
| `sheet.setValue('B3', value)` | Writes text, a number, a boolean, a `Date` or a formula (`{ formula: '=SUM(A1:A3)' }`); `null` clears the value and leaves the format. Formulas depending on the cell are recalculated. |
| `sheet.setValues('A2', rows)` | Writes a matrix from the given cell, creating rows and cells in order. `undefined` leaves the cell as it is. |
| `wb.fillTemplate(values)` | Replaces `{{FIELD}}` (spaces and a hint are accepted: `{{ FIELD }}`, `{{FIELD: hint}}`) in shared and inline strings; returns how many it replaced. Variables without a value stay as they are. |
| `wb.addSheet(name, after?)` | Adds an empty sheet behind the tab `after` (name or position; at the end without it) and returns it. The name follows Excel's rules: 1 to 31 characters, without `[ ] : * ? / \`, without an apostrophe at the start or end, different from `History` and the other sheets (ignoring case). |
| `wb.renameSheet(sheet, name)` | Renames a tab (name or position) and the references to it in formulas, defined names, data validations, conditional formats and charts; returns how many it changed. References inside texts (`INDIRECT("Sheet1!A1")`) and to other workbooks stay as they are, as in Excel. |
| `wb.markers({ kind })` | Markers of the standard lexicon (the same as `@dimarborda/docx-editor`) in the text cells of all sheets: `field` for `{{KEY}}` / `{key}`, `pending` for `[…]` (content to complete). Each with `sheet`, `sheetName`, `ref` (`B3`), `row`, `col`, `raw`, `key`, `hint`, `offset` and `length`. |
| `wb.checkpoint(selection?)` | Opens an undo step: changes until the next `checkpoint` undo together. `selection` is a free text that `undo()` returns. `wb.setSelectionAfter(selection)` stores the one after, which `redo()` returns. |
| `wb.undo()` / `wb.redo()` | Undoes or redoes a step: `{ selection, parts }`, or `null` if there is nothing. `wb.canUndo` / `wb.canRedo` say whether it can. |
| `wb.recalculate()` | Recalculates all formulas and returns how many cells changed value. |
| `wb.areas({ sheet })` / `sheet.areas()` | The areas of the workbook (see [Areas](agents.md#areas)). |
| `wb.area(target, { sheet, maxRows, locale })` / `sheet.area(range)` | Context of an area for an agent. |
| `sheet.regionAt('B5')` | The region around a cell, as Ctrl+A selects it in Excel (or its table). |
| `sheet.setFormat('A1:C3', format)` / `sheet.clearFormat(range)` | Applies a format (see [Format](#format)) or removes it. |
| `sheet.setColumnWidth('B:D', 14)` | Column width in characters. |
| `sheet.mergeCells('A1:H1')` / `sheet.unmergeCells(range)` | Merges or splits cells; when merging, only the first may have a value. |
| `sheet.moveRows(first, count, before)` / `sheet.moveColumns('B', count, 'E')` | Moves whole rows or columns before another, numbered as it is now (see [Move rows and columns](operations.md#move-rows-and-columns)), and returns where the first ends up. `editor.moveRows` and `editor.moveColumns` do the same in one undo step and select the result. |
| `sheet.insertRows(before, count?)` / `sheet.insertColumns('C', count?)` | Inserts empty rows or columns, with the format of the previous one (see [Insert and delete](operations.md#insert-and-delete)). Also `editor.insertRows` and `editor.insertColumns`, in one undo step. |
| `sheet.deleteRows(first, count?)` / `sheet.deleteColumns('C', count?)` | Deletes rows or columns. Also `editor.deleteRows` and `editor.deleteColumns`. |
| `sheet.setRowHeight('3:5', 30)` | Row height in points; `null` goes back to the automatic height. |
| `sheet.setRowsHidden('3:5', true)` / `sheet.setColumnsHidden('B:D', true)` | Hides or shows lines. Also `editor.hideRows`, `showRows`, `hideColumns` and `showColumns`. |
| `sheet.fill('A1:A2', 'A1:A20', { series })` | Fills a range from the cells at one of its ends (see [Fill](operations.md#fill)). `editor.fill('A1:A20')` does it from the selection, in one undo step. |
| `sheet.getCells('A1:C3')` | The cells of a range with their type, value and formula. |
| `wb.toBytes()` | The workbook as `.xlsx`. |
| `wb.close()` | Frees the engine's memory. |

## Formulas

Formulas are written with `{ formula: '=SUM(B2:B10)*TAX' }` and calculated right away. Each time a cell is written (or a template filled) all the workbook's formulas are recalculated, in dependency order, and only the `<v>` of the cells whose result changed is rewritten. `wb.recalculate()` does the same without editing anything, for example to refresh a workbook just opened.

It uses the English syntax Excel saves (`SUM`, `IF`, `,` between arguments; `;` is also accepted). Relative and absolute references, ranges, whole columns (`A:A`), other sheets (`'My data'!B2`), defined names, shared formulas and constant arrays (`{1,2;3,4}`) are supported, with Excel's conversions between text, number and boolean and its error codes. Functions:

- **Math:** `SUM`, `PRODUCT`, `AVERAGE`, `MIN`, `MAX`, `MEDIAN`, `COUNT`, `COUNTA`, `COUNTBLANK`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `TRUNC`, `ABS`, `SIGN`, `SQRT`, `POWER`, `MOD`, `PI`, `CEILING`, `FLOOR`, `CEILING.MATH`, `FLOOR.MATH`, `SUMPRODUCT`.
- **Conditional:** `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS` (criteria such as `">5"`, `"<>x"`, `"n*"`).
- **Logic and information:** `IF`, `IFS`, `SWITCH`, `IFERROR`, `IFNA`, `AND`, `OR`, `XOR`, `NOT`, `TRUE`, `FALSE`, `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISNONTEXT`, `ISLOGICAL`, `ISERROR`, `ISERR`, `ISNA`, `NA`.
- **Lookup:** `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `MATCH`, `INDEX`, `CHOOSE`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`.
- **Text:** `CONCATENATE`, `CONCAT`, `TEXTJOIN`, `LEFT`, `RIGHT`, `MID`, `LEN`, `UPPER`, `LOWER`, `PROPER`, `TRIM`, `SUBSTITUTE`, `REPT`, `EXACT`, `FIND`, `SEARCH`, `VALUE`, `TEXT`.
- **Dates:** `TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, `DAY`, `WEEKDAY`, `EDATE`, `EOMONTH`, `DAYS`.

A formula that uses anything else (a function not on the list, `OFFSET`, `INDIRECT`, references to other workbooks or to tables, array formulas, results that spill over several cells) or that forms a circular reference keeps the value Excel saved, and those depending on it use that value. When the file opens, Excel calculates it (see [Lossless](#lossless)).

## Format

`sheet.setFormat(range, format)` changes only what is given and keeps the rest of each cell's look:

| Property | Value |
|---|---|
| `bold`, `italic`, `underline`, `strike` | `true` / `false` |
| `color` | Font color, `#RRGGBB` |
| `fontSize`, `fontName` | Points; font name |
| `fill` | Background `#RRGGBB`; `null` removes it |
| `numberFormat` | Excel code: `#,##0.00`, `0%`, `dd/mm/yyyy`, `#,##0.00 "€"`… |
| `align`, `valign`, `wrap` | `left`, `center`, `right`…; `top`, `center`, `bottom`; text wrapping |
| `border` | A style (`'thin'`, `'medium'`, `'thick'`, `'dashed'`, `'dotted'`, `'double'`) on all sides, `null` to remove them, or `{ style, color, edges }` with `edges`: `all`, `outline`, `inside`, `top`, `bottom`, `left` or `right` |

New formats are added to `xl/styles.xml` and identical existing ones are reused; nothing already there is rewritten.

## Dates

Excel stores dates as serial numbers (days since 1900, or since 1904 in workbooks with that system) that look like a date through the cell's format. When reading, a numeric cell with a date format comes back as a `Date`; when writing a `Date`, the cell gets a date format based on the one it had (same font, fill and borders), unless it already showed dates.

Excel dates have no time zone: the time Excel shows is represented as the **UTC** time of the `Date`. Write with `new Date(Date.UTC(2024, 0, 15))` and read with `getUTCFullYear()`, `getUTCHours()` or `toISOString()`.

## Lossless

Only the parts of the file that change are rewritten (the edited sheet, shared strings and, if needed, styles), and within them only the written cells: charts, images, conditional formats, validations, defined names and everything else are saved byte by byte as they were.

If the first cell of a shared formula is written over (the one Excel keeps for a whole group of filled cells), the other cells of the group each get a formula of their own, with the same result.

When a workbook changes, `xl/calcChain.xml` is removed (Excel rebuilds it) and `fullCalcOnLoad="1"` is set so that Excel recalculates formulas when opening it: that way those this engine does not calculate end up right too.

Current limitations:

- Formulas outside what is supported (see [Formulas](#formulas)) are not recalculated: `getValue` returns their last saved result.
- Formulas are written in English syntax (`SUM`, not `SUMA`); function names in Spanish are not translated.
- A `{{FIELD}}` variable split across text fragments with different formats is not recognized.

## Content Security Policy

As the engine is WebAssembly, the production `script-src` needs `'wasm-unsafe-eval'`.

## Development

From the root of the repository:

```bash
npm run build:wasm:xlsx
npm test -w packages/xlsx-editor
npm run dev        # http://localhost:5173/xlsx.html
```
