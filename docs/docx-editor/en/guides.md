[← Index](../README.md) · **English** · [Español](../es/guides.md)

# Guides for @dimarborda/docx-editor

[API](api.md) · [Guides](guides.md) · [Integrations](integrations.md) · [Security and limitations](security.md)

Logical sections, markers, agent activity, tables, PDF, read-only mode and images.

## Logical sections

Logical sections mark where each part of a document starts and ends (for example "Introduction", "Characterization" and "Conclusions" in a technical report), so you can read, fill in or review it section by section. They do not change how the document looks and have nothing to do with Word section breaks.

```js
// Mark paragraphs 2 to 6 as "introduction" (tables between them are included whole)
editor.defineSection('introduction', { from: 2, to: 6 }, { title: 'Introduction' });

editor.sections();
// [{ key: 'introduction', title: 'Introduction', first: 2, last: 6, depth: 0, parent: null,
//    locked: false, contentLocked: false }]

editor.sectionContent('introduction').text;   // plain text of the section
editor.replaceSection('introduction', ['New paragraph.', 'One more.'], { keepHeading: true });
editor.unwrapSection('introduction');         // removes the mark, keeps the content
```

In the `.docx`, each section is a block-level content control (`w:sdt`) with `w:tag="section:<key>"` and the title in `w:alias`, so Word keeps it when saving and any other tool (python-docx, Open XML SDK, docx4j) can find it by searching for that tag. Sections can be nested (`depth`, `parent`). Each call that modifies the document is one undo step and throws an `Error`, leaving the document untouched, if the key already exists, is missing or is repeated, or if the range cuts through a table or another section, falls inside a table or contains a section break.

### Create sections from headings

If the document already uses Word's heading styles, you do not need to mark each section by hand:

```js
const { created, skipped } = editor.sectionsFromHeadings({ fromLevel: 1, toLevel: 2 });
// created: "1. Introduction" → 'introduction', with "1.1 Scope" → 'scope' nested inside it
```

Each heading of the requested level spans up to the next heading of its own level or a higher one; lower levels are nested. The key comes from the heading text, without the numbering and without accents ("2. Site characterization" → `site-characterization`); if it is repeated, `-2`, `-3`… is appended. Headings that are already a section, or whose span would cut through a table, another content control or a section break, are returned in `skipped` with the reason. Calling it again does not duplicate anything.

### Insert new sections

```js
editor.insertSection('annexes', { title: 'Annexes' });                         // at the end of the document
editor.insertSection('summary', { title: 'Summary', before: 'introduction' });
editor.insertSection('method', { title: 'Method', inside: 'characterization', paragraphs: ['12 samples were taken.'] });
```

By default the new section gets a heading with the Word style that matches its depth (Heading 1, 2 or 3) and an empty paragraph. With `heading: false` no heading is added; with `heading: 'Heading2'` you choose the style.

### Locks

```js
editor.setSectionLock('introduction', { locked: true });         // cannot be deleted
editor.setSectionLock('legal-framework', { contentLocked: true });   // its content cannot be edited by hand
```

- `locked`: the section cannot be deleted (`w:lock="sdtLocked"`). The editor rejects any selection that would delete it entirely.
- `contentLocked`: its content cannot be edited by hand (`w:lock="contentLocked"`; with both locks, `sdtContentLocked`). The editor rejects typing, deleting, formatting and image changes inside it, and fires a `message` event that explains why. Word applies the same locks.

Locks limit what the person does, not what code does: `replaceSection`, `insertParagraphs` and the other calls keep working, so a tool can fill in a section that the user can only read. They can also be set when the section is created: `defineSection(key, range, { locked, contentLocked })`.

### Schema by document type

A schema says which sections a type of document must have and in what order. It is stored inside the `.docx`, so it travels with the file:

```js
editor.setSectionSchema({
  type: 'technical-report',
  title: 'Technical report',
  sections: [
    { key: 'introduction', title: 'Introduction' },
    { key: 'scope', title: 'Scope', parent: 'introduction', required: false },
    { key: 'characterization', title: 'Characterization' },
    { key: 'conclusions', title: 'Conclusions' },
  ],
});

editor.validateSections();
// { valid: false, type: 'technical-report', issues: [
//   { kind: 'missing', key: 'conclusions', title: 'Conclusions', message: 'Section "Conclusions" is missing.' } ] }

editor.addMissingSections();   // inserts the required ones that are missing, each in its place
```

`required` is `true` by default, and `parent` nests a section inside another one that appears earlier in the list. `validateSections()` reports sections that are missing (`missing`), repeated (`duplicated`), empty (`empty`: no text or tables besides the heading), out of order (`order`), inside a section where they do not belong (`parent`) and not mentioned by the schema (`unexpected`). The last one does not make `valid` `false`. Without a schema, everything is valid. `addMissingSections({ optional: true })` also adds the optional ones. `sectionSchema()` returns the stored schema and `setSectionSchema(null)` removes it.

In the file, the schema is a custom XML part (`customXml/itemN.xml`, namespace `urn:rust-web-docx:sections:1`) related from `word/document.xml`, just like the ones Word creates:

```xml
<sections xmlns="urn:rust-web-docx:sections:1" type="technical-report" title="Technical report">
  <section key="introduction" title="Introduction" required="true"/>
  <section key="scope" title="Scope" required="false" parent="introduction"/>
</sections>
```

### Show sections in the editor

```js
const editor = await DocxEditor.create(container, { showSections: true });
editor.setShowSections(false);   // or true; editor.sectionsVisible tells you whether they are shown
```

Each section is drawn with a bracket in the left margin, from its first line to its last, with the title on top; nested ones sit further out. 🛡 marks sections that cannot be deleted and 🔒 marks those with locked content, which also get a dotted bracket. Clicking the title selects the section. This is a view only: it is not saved in the `.docx` and is not printed. If the margin is very narrow, only the bracket is shown and the title appears on hover.

## Markers: fields and pending items

A standard vocabulary, shared with [`@dimarborda/xlsx-editor`](../../../packages/xlsx-editor), so an agent or an automation can tell what is missing in a document without guessing:

| Written as | Type | Meaning |
| :--- | :--- | :--- |
| `{{KEY}}`, `{{ KEY }}`, `{{KEY: hint}}` | `field` | Value that an automation fills in (canonical form). The hint after the colon says what is expected (`{{DATE: dd/mm/yyyy}}`) |
| `{key}`, `{key:hint}` | `field` | Short form, no spaces inside |
| `[…]`, `[...]`, `[____]`, `[Complete: project scope]` | `pending` | Content that someone (a person or an agent) still has to write. The text inside is the hint |

Citations (`[1]`, `[2, 3]`, `[4-6]`), checkboxes (`[ ]`, `[x]`) and `[sic]` are not markers, and no marker spans a line break.

```js
editor.markers({ kind: 'pending' });
// [{ id: '12:9', kind: 'pending', syntax: 'bracket', raw: '[Complete: scope]', key: null,
//    hint: 'Complete: scope', paragraph: 12, offset: 9, length: 17, place: 'body', section: 'purpose' }]

const [first] = editor.markers({ kind: 'pending', section: 'purpose' });
editor.fillMarker(first, 'Design and implementation of the payments module.'); // one undo step
```

`place` is `body`, `table`, `textbox`, `headerFooter` or `headerFooterTextbox` (a floating text box inside a header or footer, e.g. an institutional letterhead logo); `section` is the innermost [logical section](#logical-sections) that contains it — `null` for the last three values, as was already the case with `textbox`. `fillMarker` finds the marker again even if text has been typed before it since it was listed, and throws an `Error` if it is no longer there. To analyze a loose piece of text (for example, what an agent is about to insert) without a document: `import { findMarkers } from '@dimarborda/docx-editor'`.

### Highlight the pending items

In a long document, whoever is filling it in cannot tell at a glance what is left. With `highlightPending`, each pending `[…]` is shown with a soft background; with `{ fields: true }`, fields are shown too, in another color:

```js
const editor = await DocxEditor.create(container, { highlightPending: true });
editor.setHighlightPending({ fields: true });  // pending items and fields; false turns it off
editor.highlightedMarkers;                     // ['pending', 'field']

// "Next pending" button: from the cursor, wrapping back to the start at the end
const next = editor.nextMarker({ kind: 'pending', wrap: true });
if (next) editor.select({ paragraph: next.paragraph, offset: next.offset }, { paragraph: next.paragraph, offset: next.offset + next.length });
```

This is a view only: it is not saved in the `.docx`, not printed and not included in `downloadPdf`. The highlight updates as you type: when a marker is filled in or deleted, its highlight disappears. Colors are changed with the CSS variables `--docx-pending-highlight` and `--docx-field-highlight`.

## Activity: where the agent is working

When an agent modifies the document through the API, the person reading it must be able to find their bearings. There are two levels, both view-only (nothing is saved in the `.docx`, they do not move the user's cursor or scroll, and they cost nothing if unused):

- **Automatic flash**: every edit made from code (`insertParagraphs`, `replaceParagraphs`, `moveParagraphs`, `replaceSection`, `setTableCell`, `insertTableRows`, `replaceVariables`, `fillMarker`) highlights what changed for a moment. Duration is set with the `flash` option (ms, default `1800`; `false` disables it).
- **Work mark**: `markActivity()` leaves a tinted area with a bar and a label in the margin, plus a fixed chip above the editor that shows where it is (↑ / ↓ when it is out of view) and takes you there with a click.

```js
const work = editor.markActivity('conclusions', { label: 'Agent · drafting' }); // a section: it follows it as it grows
const added = editor.replaceSection('conclusions', ['First draft…']);
work.update(added, { label: 'Agent · reviewing' }); // also { from, to } or { paragraph, offset, length }
work.done();

editor.addEventListener('activity', e => console.log(e.detail.activities));
// [{ id: 'a1', kind: 'work', label: 'Agent · drafting', from: 14, to: 18, section: 'conclusions', page: 2 }]
```

The target can be a section key (recommended: it follows the section), a `{ from, to }` range, what `insertParagraphs` / `replaceSection` return (`{ first, count }`) or a `{ paragraph, offset, length }` span like a marker. Paragraph ranges stay in place when the user or code inserts or deletes paragraphs before them. `ttl` (ms) removes it automatically; `color` changes its color (useful with several agents). If your application has a fixed bar at the top, move the chip down with `--docx-activity-dock-top`.

## Existing tables

To fill in a table from a template (for example, the rows of a risk matrix), locate it with `tables()` and write to it:

```js
// Each body table: { index, rows, cols, header, first, last }
const matrix = editor.tables().find(t => t.header[0] === 'Risk');
if (!matrix) throw new Error('The template has no risk matrix');

// Two rows at the end, with the format of the last row (borders, shading, font…)
const { first } = editor.insertTableRows(matrix.index, matrix.rows, [
  ['Server outage', 'Medium', 'High'],
  ['Data leak', 'Low', 'Critical\nreview every month'],
]);

// Fix a cell: keeps the format of the text that was there
editor.setTableCell(matrix.index, first, 1, 'High');

editor.undo();   // undoes only the fix
editor.undo();   // undoes both rows at once
```

- **Indexes.** Tables are numbered from 0 in document order, not counting nested ones or those inside text boxes. Rows and columns also start at 0. Columns are grid columns: to write to a horizontally merged cell, use its first column. `first`/`last` are the paragraph range of its cells, using the same numbering as `structure()` and `selection`.
- **`setTableCell`** replaces the text of the cell and keeps its paragraph properties and the format of its first *run* (font, size, color, bold). `\n` is a line break inside the cell.
- **`insertTableRows(table, at, rows, { copyFormatFrom })`** inserts before row `at`: `0` at the start, `rows` (the row count) at the end. Each new row copies the format of row `copyFormatFrom` (by default the one before `at`, or the next one if you insert at the start): height, borders, shading, cell widths, paragraph properties and the format of the first *run*. Rows with fewer values are completed with empty cells, which keep the format for whatever is written later; rows with more values than cells produce an error. It returns `{ first, count }`: the first new row and how many were inserted. A repeated header row is only copied as a header if you insert at the start. Be careful when inserting right below the header: the previous row is the header and its shading would be copied; pass `copyFormatFrom` with a data row. The same applies if the template's empty rows have gray or italic placeholder text: copy a row that is already filled in.
- **`deleteTableRows(table, from, count)`** deletes rows; the table cannot be left without rows (to remove it entirely use `replaceParagraphs` with its `first`–`last` range).
- Each call is a single undo step, fires `change`, redraws the document and also works in read-only mode.

### Duplicate a table

To add another copy of a table that is already in the document (a record card with shaded labels, a signature block…), duplicate it instead of creating it with `insertTable`, which uses the editor's style. The copy is identical: borders, shading, widths, merged cells, format and text. Then adjust it with `setTableCell`, `insertTableRows` and `deleteTableRows`:

```js
const card = editor.tables().find(t => t.header[0] === 'Objective');
const { table } = editor.duplicateTable(card.index, {
  at: 'end',
  values: [[null, 'Reduce delivery times'], [null, 'Ana Pérez']], // null keeps the text
});
editor.insertTableRows(table, 2, [['Deadline', '3 months']]);
```

`at` works as in `insertTable` (`cursor`, `start` or `end`; text after the cursor ends up below the table). `values` changes the text of the cells by row and column, like `setTableCell`, keeping their format. It returns `{ table, first }`: the index of the copy in `tables()` (the original shifts if the copy lands before it) and the paragraph of its first cell. The copy and the values are a single undo step.

The copy does not include Word bookmarks, comment marks, footnote references or content control ids, which must be unique in the document; tracked changes are accepted and images get new ids. A table that contains a logical section cannot be duplicated.
- They throw a clear `Error`, without modifying anything, if an index does not exist, if the cell continues a vertical merge (the message says which row to write to), if the row being copied or the position has vertically merged cells, or if the cell has a nested table.

## PDF

```js
const pdf = await editor.exportPdf();            // Uint8Array
await editor.downloadPdf('contract.pdf');         // or download it directly
const blob = await editor.exportPdfBlob({ watermark: false });
```

It generates a vector PDF of the pages as they look in the editor: real text (selectable and searchable) with fonts embedded and subset to the characters used, images, tables, borders, shapes, text boxes, headers and footers with their page number. The watermark from `setWatermark()` is included unless you pass `{ watermark: false }`. The fonts are those uploaded by the user with `registerFont()` or the package's metric-compatible substitutes (`@dimarborda/docx-editor/fonts`); without them, the standard PDF fonts are used, and the text is fitted to the width measured on screen so nothing shifts. [pdf-lib](https://pdf-lib.js.org/) is downloaded only the first time you export.

## Read-only mode

The editor is editable by default. To show a document without letting the user change it:

```js
const editor = await DocxEditor.create(el, { readOnly: true }); // from the start
editor.setReadOnly(true);   // or at any time
editor.setReadOnly(false);  // editable again
```

```html
<docx-editor src="/contract.docx" readonly></docx-editor>
```

In read-only mode the user can scroll, zoom, select and copy text, but typing, deleting, pasting, cutting, formatting shortcuts, undo and redo, and image selection and editing are ignored. The toolbar disables its editing buttons (zoom stays active). Calls from code (`replaceVariables`, `insertText`, `updateImage`…) keep working, so your application can fill in a document that its users only view.

## Images

With the mouse: one click selects the image and shows eight handles. Corners resize while keeping the aspect ratio (`Shift` releases it) and sides change a single dimension. A floating image is moved by dragging it or with the arrow keys (`Shift` + arrow = 10 px). `Del` deletes it and `Esc` returns to the text. An image behind the text can only be selected where there is no text on top of it, so clicking on the text of a letterhead still places the cursor.

Text boxes and shapes (rectangles, ellipses…) work the same way: they appear in `images()` with `kind: 'textbox'` or `kind: 'shape'`, with their text in `text` and their fill and border in `shape`, and they are selected, moved, resized and have their wrapping changed with the same methods.

From code, an image is identified by `{ paragraph, index }`: its paragraph and its position among the images of that paragraph. Any element of `images()` works as a reference. Measurements are in CSS px at 100% zoom.

```js
const [logo] = editor.images();
// { paragraph, index, width, height, wrap, wrapSide, anchored, horizontal, vertical,
//   distance: { top, bottom, left, right }, alt, textOffset, bounds: { page, x, y, width, height } }

editor.resizeImage(logo, { width: 180 });            // the height follows the aspect ratio
editor.resizeImage(logo, { scale: 0.5 });
editor.setImageWrap(logo, 'square', { side: 'bothSides', distance: 12 });
editor.alignImage(logo, 'right');                    // floating: within the margins
editor.moveImage(logo, { dx: 0, dy: 40 });           // or { x, y } in page coordinates
editor.updateImage(logo, {                           // several changes in a single undo step
  wrap: 'topAndBottom',
  horizontal: { relativeTo: 'margin', align: 'center' },
  vertical: { relativeTo: 'paragraph', offset: 10 },
  alt: 'Company logo',
});
editor.deleteImage(logo);

await editor.insertImage(file, {                     // floating from the start
  width: 160, wrap: 'square',
  horizontal: { relativeTo: 'margin', align: 'right' },
  vertical: { relativeTo: 'paragraph', offset: 0 },
});
```

| `wrap` | Behavior with the text |
| :--- | :--- |
| `inline` | In line with the text: behaves like a large character that rests on the baseline and moves to the next line if it does not fit |
| `square` | The text wraps around the image's rectangle. `wrapSide`: `bothSides` (the text uses the wider side), `left`, `right` or `largest` |
| `tight` / `through` | Like `square`; saved with Word's outline, but laid out using the rectangle |
| `topAndBottom` | The text goes only above and below the image |
| `behind` / `inFront` | Floats behind or in front of the text, which ignores it |

`horizontal.relativeTo` can be `margin`, `page`, `column`, `character`, `leftMargin`, `rightMargin`, `insideMargin` or `outsideMargin`; `vertical.relativeTo`, `margin`, `page`, `paragraph`, `line`, `topMargin`, `bottomMargin`, `insideMargin` or `outsideMargin`. On each axis you use either an `offset` (distance from the left or top edge of the frame) or an `align` (`left`/`center`/`right` or `top`/`center`/`bottom`). When moving an image that uses an offset, its frame is kept; an aligned axis becomes page-relative. To move an inline image, `moveImage` first converts it to `square`.
