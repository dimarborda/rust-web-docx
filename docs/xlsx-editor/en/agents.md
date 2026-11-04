[← Index](../README.md) · **English** · [Español](../es/agents.md)

# Areas and context for agents

An agent does not need to read the whole sheet: it can ask for the areas and then the context of the one it is interested in.

## Areas

`wb.areas()` uses what the file itself marks, in this order:

1. Excel **tables** (Format as Table, `xl/tables/tableN.xml`): name, range, column names and totals row.
2. **Defined names** that point to a range (`Sales = Data!$A$3:$H$40`) and the **print area**.
3. **Filters** (`autoFilter`) of the sheet.
4. **Regions**: blocks of adjacent filled cells, as Ctrl+A selects them, where the file marks nothing. Initial rows with a single cell (or a merged one) are the **title**, as is loose text right above; the first row of texts followed by data (or in bold) is the **header**.

Each area is `{ kind, name, sheet, sheetName, ref, address, title, headerRow, totalsRow, body, columns }`; `address` (`Sales!A1:H64`) works as a target in any call.

`wb.area(target)` accepts the name of a table or named range, an address with its sheet or, with `{ sheet }`, a range of that sheet. It returns the area's description and also:

- `columnsInfo`: per column, its name, type (`number`, `text`, `date`, `mixed`…), how many cells have a value, sum, minimum, maximum and mean of the numbers, and how many formulas there are with an example (`D4: =B4*C4`).
- `rows`: the first `maxRows` rows of data (30 by default) with the text the cells show, `firstRow`, `rowCount` and `truncated`.
- `totals`, `merges` and `markdown`: the area as a Markdown table with row numbers and column letters, ready to pass to a model.

## From the editor

- `editor.selectArea(target)`: selects an area and shows its sheet. It accepts a table or a named range, an address (`Sales!A1:H40`) or a range of the active sheet. It returns the area (see [Areas](#areas)).
- `editor.selectionContext({ maxRows })`: context of what is selected for an agent: the range or, with a single cell, the region around it (its table, or the block Ctrl+A would select). It returns title, header, types and totals per column, the first rows and a Markdown table.

To format from code use `editor.workbook.sheet(…).setFormat(…)` and then `editor.reload()`.
