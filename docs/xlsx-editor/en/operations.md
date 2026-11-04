[← Index](../README.md) · **English** · [Español](../es/operations.md)

# Rows and columns

## Move rows and columns

Moving works like "cut and insert" in Excel. Cells, formats, height or width, merged cells, comments and anchored images go with the row or column, and the formulas of the whole workbook, defined names, charts, tables, filters, data validations, conditional formats and hyperlinks are updated. Formulas are rewritten as text, so those this engine does not evaluate follow the move too; only the affected references change.

- A reference to a cell follows that cell.
- A range keeps its rows: moving inside `SUM(B2:B10)`, also to its ends, does not change it. If a row leaves the range it shrinks; if one enters, it grows. A range contained in what moves goes with it.
- References to other workbooks (`[1]Sheet!A1`), 3D references (`Sheet1:Sheet3!A1`) and those inside texts (`INDIRECT("A1")`) stay as they are.
- Shared formulas that change get a formula of their own in each cell.

What would tear something that must stay whole is refused, changing nothing:

- **Merged cells** cut by the move. Reordering under a merged range is allowed without touching its first line, which holds the value (for example, columns under a title).
- **Array formulas** of several cells cut by the move.
- **Tables and filters**, unless they move whole or lines are reordered inside them: their data rows, without header or totals, or their columns.
- **Sheets with pivot tables**.

A conditional format or validation rule follows its cells, even if it ends up split over several ranges. Its relative references keep their distance to each cell, and absolute ones follow their target.

## Insert and delete

Inserting and deleting update the same things as moving, across the whole workbook.

**When inserting**, new lines take the height or width and the format of the previous one, without its values. Ranges, merged cells and rules they fall inside grow, as do tables if they are data rows. If inserted inside the frozen area, it grows. It is refused if the last lines of the sheet have data, because they would fall off.

**When deleting**, what pointed only at those lines becomes `#REF!`, also in other sheets, names and charts, as in Excel. Ranges and merged cells shrink. Comments, hyperlinks and rules that apply to nothing else disappear with their cells. Images anchored in them are not deleted: they move to the line that takes their place.

Refused, changing nothing:

- Inserting or deleting **columns inside a table**, deleting **a whole table**, its header, its totals or all its data rows.
- Inserting or deleting across an **array formula** of several cells.
- Deleting a column with an **applied filter**.
- Any of these on a **sheet with pivot tables**.

## Fill

Along each filled row or column the source cells repeat with their format, and:

- **Formulas** are rewritten for each cell: relative references shift and absolute ones do not.
- Two or more **numbers** continue their trend (the line that fits best, as Excel does). A single number is copied. A single **date** goes on day by day.
- A **text ending in a number** keeps counting (`Lot 9`, `Lot 10`; `L08`, `L09`).
- **Months and days**, full or abbreviated, in Spanish or English, follow their order and respect capitals.
- Anything else is copied as it is, and an empty source cell empties its target.

With `{ series: false }`, and with **Ctrl/Cmd+D** and **Ctrl+R**, it only copies (formulas still shift). Merged cells cannot be filled over. Shrinking the selection with the small square deletes nothing, unlike Excel.

## Find and replace

`*` is any text, `?` one character and `~` takes them literally. By default it searches the text the cells of the active sheet show; the panel and the options allow matching case, requiring the whole cell, looking in formulas or covering the whole workbook (without hidden sheets). From code: `editor.find(text, options)` and `editor.replaceAll(text, replacement, options)`.

Replacing acts on text cells and on the text of formulas, never on numbers or dates. The result is read as if typed by hand, so `12` becomes a number, and "All" is one undo step. Text with several formats within a cell ends up with a single format.

## Status bar

With more than one cell selected, the **status bar** shows the average, count and sum, as Excel does: numbers in the format of the first of them, without counting hidden rows and columns, and only the count if there is an error in the selection. `editor.selectionStats` gives that data (and the minimum and maximum) from code, and `status-bar="false"` hides the bar. The reason for a refused change also appears there.
