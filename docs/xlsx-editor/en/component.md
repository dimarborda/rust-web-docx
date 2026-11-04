[← Index](../README.md) · **English** · [Español](../es/component.md)

# The component

A component to view and edit Excel workbooks (`.xlsx`) in the browser, drawn on canvas the way Excel shows them. Inside is an engine written in Rust and compiled to WebAssembly, which reads the file, resolves styles and number formats, recalculates formulas and saves losslessly. The same engine works without an interface, in Node.js, from [`@dimarborda/xlsx-editor/engine`](engine.md). It shares the OOXML package layer (`crates/ooxml-core`) with [`@dimarborda/docx-editor`](../../docx-editor/README.md). It works in any frontend (HTML, React, Vue, Svelte, Angular, Tauri) and nothing leaves the browser.

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

## What it shows

- **Cells as in Excel:** fonts, bold, italic, underline, strikethrough, colors (including theme colors and their tints), fills and borders in every style.
- **Number formats:** separators by language, currency, percentages, scientific notation, fractions, sections and conditions, colors such as `[Red]`, and dates and times with month and day names in Spanish or English.
- **Alignment:** Excel's defaults (numbers right, text left), horizontal and vertical, indent, text wrapping, text overflowing onto empty cells and `####` for numbers that do not fit.
- **Grid:** column widths and row heights from the file, hidden rows and columns, merged cells, frozen panes and gridlines (or their absence).
- **Navigation:** smooth scrolling on large sheets (only what is visible is drawn), zoom, sheet tabs, selection with mouse and keyboard, and a bar with the active cell and its content or formula.

Charts, images, conditional formats and comments are not drawn yet, and rotated text is drawn horizontally. Formulas are recalculated after each change by the package's engine; those that use something it does not support show the last value Excel saved.

## Editing

| Action | How |
| :--- | :--- |
| Type in a cell | Start typing (replaces the content), press **F2** or double click (edits the current content), or type in the cell bar. |
| Confirm | **Enter** (down), **Shift+Enter** (up), **Tab** / **Shift+Tab** (right / left). If you started typing without F2, the arrows also confirm and move; **F2** switches between that mode and moving the caret within the text. |
| Fill the selection | **Ctrl/Cmd+Enter** writes the same thing in every selected cell. |
| Line break in the cell | **Alt+Enter**. |
| Cancel | **Escape**. |
| Clear | **Delete** or **Backspace** empty the selected range; the cell format stays. |
| Copy, cut and paste | **Ctrl/Cmd+C**, **X** and **V**, as tab-separated text (TSV), the same that Excel and Google Sheets exchange. Copying takes the texts the cells show; pasting into the same editor what was just copied pastes the exact values. A block that fits a whole number of times in the selection repeats. Cut empties the cells right away (not on paste, as Excel does). |
| Select rows and columns | Click their header (**Shift+click** or drag for several); click the corner or **Ctrl/Cmd+A** for the whole sheet. **Ctrl+Space** selects the columns of the selection and **Shift+Space** its rows. Copying and clearing only reach as far as the sheet has cells. |
| Move rows and columns | Select them and drag their header to where the green line shows (**Escape** cancels), or use **Alt+↑ ↓** for the rows and **Alt+← →** for the columns of the selection. See [Move rows and columns](operations.md#move-rows-and-columns). |
| Insert and delete rows and columns | Right click a cell or a header (also **Shift+F10** or the menu key). With whole rows or columns selected, **Ctrl/Cmd+Shift++** inserts and **Ctrl/Cmd+-** deletes. See [Insert and delete](operations.md#insert-and-delete). |
| Change width and height | Drag the right edge of a column header or the bottom edge of a row header (**Escape** cancels); double click that edge to fit the content. With several selected, it applies to all of them. The context menu also offers autofit. |
| Hide and show | Right click on headers, or **Ctrl+9** (rows) and **Ctrl+0** (columns); with **Shift** they are shown. To show them, select the lines around them. A thicker line in the header marks where hidden ones are. They keep their cells and size. |
| Fill | Drag the small square at the corner of the selection. Double clicking it fills down as far as the data in the neighboring column goes. **Ctrl/Cmd+D** copies down and **Ctrl+R** to the right. See [Fill](operations.md#fill). |
| Autosum | **Alt+=** or the context menu. In a cell it proposes `=SUM(…)` with the numbers above or, if there are none, to the left, to confirm with **Enter**. With a range selected it writes the sum below each column (or to the right of a single row), in its last line if empty or in the next one, without overwriting data. |
| Find and replace | **Ctrl/Cmd+F** finds and **Ctrl+H** replaces; **Enter** and **Shift+Enter** go through the results. See [Find and replace](operations.md#find-and-replace). |
| Move around | **Ctrl/Cmd+arrows** go to the edge of the data, **Home** to the start of the row, **Ctrl/Cmd+Home** and **Ctrl/Cmd+End** to the corners of the data, **PageUp** and **PageDown** one screen (with **Alt**, sideways). With **Shift** they extend the selection. |
| Undo and redo | **Ctrl/Cmd+Z**, and **Ctrl/Cmd+Y** or **Cmd+Shift+Z**. Each edit, clear, paste or move is one step; undo goes back to the sheet and selection of that moment. |

### What is saved when typing

What is typed is read the way Excel does in the editor's language (`locale`; `es` and its variants use the Spanish conventions and the rest the English ones):

| You type (`es`) | You type (`en`) | Saved as |
| :--- | :--- | :--- |
| `1234,5` · `1.234,5` · `-0,5` · `(300)` · `1,5E+3` | `1234.5` · `1,234.5` · `-0.5` · `(300)` · `1.5E+3` | Number |
| `1.234,50 €` | `$1,234.50` | Number (the symbol is dropped) |
| `15%` · `12,5 %` | `15%` · `12.5 %` | Number divided by 100 (0.15) |
| `15/03/2024` · `15-3-24` · `15/3` · `2024-03-15` | `3/15/2024` · `3-15-24` · `3/15` · `2024-03-15` | Date (without a year, the current one; two-digit years: 00–29 → 2000, 30–99 → 1900) |
| `15/03/2024 14:30` | `3/15/2024 2:30 PM` | Date and time |
| `14:30` · `14:30:15` | `2:30 PM` | Time, as a fraction of the day |
| `VERDADERO` · `FALSO` | `TRUE` · `FALSE` | Boolean (ignoring case) |
| `'0012` | `'0012` | Text `0012`: the leading apostrophe forces text |
| `=SUM(A1:A3)` | `=SUM(A1:A3)` | Formula (see Limitations) |
| Anything else | | Text |
| (empty) | | Empty cell |

Values are written with the engine's `sheet.setValues`, which keeps the cell format: a number written in a cell with a currency format looks like currency. Dates get a date format based on the cell's if it did not have one. Percentages and times are stored as their value (0.15 or 0.604…) and look as the cell's format makes them. The same function is available as `parseInput(text, locale)`.

### Limitations

- **Formulas:** what starts with `=` is saved as a formula and calculated, as are the formulas that depend on it. They are written in English syntax (`=SUM(A1:A3)`, `,` or `;` between arguments) in any editor language. The supported functions are in the [engine](engine.md#formulas); with others, the cell shows the last saved value and Excel calculates it when opening the file (`fullCalcOnLoad`).
- **Format:** typing does not change the number format (except for dates), and columns do not widen by themselves: a date in a narrow column looks like `####`, as in Excel.

## Attributes and events

| Attribute | What for |
| :--- | :--- |
| `src` | URL of the `.xlsx` |
| `blank` | Starts with an empty workbook |
| `locale` | Language of the formats (`es`, `en-US`…); by default, the page's |
| `sheet` | Initial sheet, by name or position |
| `zoom` | 0.25–4 |
| `tabs="false"` | Hides the tabs |
| `formula-bar="false"` | Hides the cell bar |
| `status-bar="false"` | Hides the status bar |
| `readonly` | Viewer mode: cells are selected and copied but not edited (can be set and removed at any time) |
| `quiet` | Does not write errors to the console |
| `wasm-url` | Where `xlsx_engine_bg.wasm` is served from, if the bundler does not do it |

Events on the element: `ready` (with `detail.editor`), `load`, `sheetchange`, `selectionchange`, `change` and `error`.

`change` comes after each edit, clear, paste, cut, move, fill, replacement, undo or redo:

```js
element.addEventListener('change', ({ detail }) => {
  // detail: { sheet: 'Sales', index: 0, range: 'B3:C4',
  //           source: 'edit' | 'clear' | 'paste' | 'cut' | 'move' | 'insert' | 'delete' | 'fill' | 'replace' | 'undo' | 'redo',
  //           cells: [{ ref: 'B3', value: 1200 }, …] }  // empty above 200,000 cells
});
```

## API

```js
const element = document.querySelector('xlsx-editor');
element.addEventListener('ready', async ({ detail: { editor } }) => {
  await editor.open(file);             // URL, File/Blob, ArrayBuffer or Uint8Array
  editor.setSheet('Summary');
  editor.select('B3:D7');
  editor.setZoom(1.25);

  // Change cells with the engine's API and draw again
  editor.workbook.fillTemplate({ CLIENT: 'Acme' });
  editor.workbook.sheet('Summary').setValues('B2', [[1200, 0.25]]);
  editor.workbook.addSheet('Notes');   // reload() also redraws the tabs
  editor.reload();

  editor.undo();                       // also editor.redo(), editor.canUndo, editor.canRedo
  editor.readOnly = true;              // the same as the readonly attribute

  editor.download('report.xlsx');
});
```

Changes made from code through `editor.workbook` are not an undo step of their own: call `editor.workbook.checkpoint()` first for them to be.

Without the element: `const editor = await XlsxEditor.create(container, { locale: 'en', readOnly: false })`.

## Editor methods

Besides those in the example, each as one undo step:

| Method | What it does |
| :--- | :--- |
| `editor.moveRows(first, count, before)` / `moveColumns` | Moves rows or columns, by number, and selects them ([details](operations.md#move-rows-and-columns)). |
| `editor.insertRows(before, count?)` / `insertColumns` | Inserts lines and selects them ([details](operations.md#insert-and-delete)). |
| `editor.deleteRows(first, count?)` / `deleteColumns` | Deletes lines. |
| `editor.hideRows(first, count?)` / `showRows` / `hideColumns` / `showColumns` | Hides or shows lines. |
| `editor.autofitColumns(first, count?)` / `autofitRows` | Fits to the content. Only on the editor, because it measures the text as drawn. |
| `editor.fill(target, { series })` | Fills from the selection ([details](operations.md#fill)). |
| `editor.find(text, options)` / `editor.replaceAll(text, replacement, options)` | Finds and replaces ([details](operations.md#find-and-replace)). |
| `editor.openFind({ replace })` / `editor.closeFind()` | Opens or closes the search panel. |
| `editor.selectionStats` | Count, sum, average, minimum and maximum of the selection. |

They return `false` when nothing changed: the editor is read-only or the engine refused, and then an `error` event and the status bar say why. `find` returns the results and `replaceAll` how many cells changed.
