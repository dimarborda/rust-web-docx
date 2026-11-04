[← Index](../README.md) · **English** · [Español](../es/api.md)

# @dimarborda/docx-editor API

[API](api.md) · [Guides](guides.md) · [Integrations](integrations.md) · [Security and limitations](security.md)

Methods, options and events of `DocxEditor`. The full types are in the package's `src/index.d.ts`.

## Web component

Attributes of `<docx-editor>`: `src` (URL of the .docx), `sample` (opens the sample contract), `blank` (new empty document, with optional `page-size="a4|letter|legal"` and `autofocus`), `locale` (`es` or `en`), `zoom`, `gridlines="false"`, `page-labels="false"`, `readonly` (read-only; without the attribute it is editable) and `wasm-url`.

`<docx-toolbar>` is optional. It accepts `for="editor-id"` and `items` with the controls separated by spaces and `|` as a separator; by default it shows all of them: `undo redo | style font fontSize | bold italic underline strike superscript subscript color highlight clearFormat | left center right both lineSpacing | bulletList numberedList outdent indent | imageWrap | zoom`. `style` is the list of paragraph styles, `font` and `fontSize` accept any typed name or size, `highlight` applies the last highlight color used (the arrow opens the Word colors) and `imageWrap` is a text-wrapping selector that is enabled when an image is selected; with an image selected, the alignment buttons align the image.

## DocxEditor

```js
import { DocxEditor } from '@dimarborda/docx-editor';
import '@dimarborda/docx-editor/style.css';

const editor = await DocxEditor.create(document.querySelector('#editor'), { locale: 'es' });

await editor.open(file);                 // File, Blob, ArrayBuffer or Uint8Array
editor.openBlank({ pageSize: 'letter' }); // or a new empty document, with the cursor ready
editor.variables();                      // ['{{CLIENTE}}', '{{FECHA}}']
editor.replaceVariables({ CLIENTE: 'Acme' });
editor.findReplace('Bogota', 'Bogotá', { matchCase: true });
editor.insertParagraphs([
  { text: 'SERVICES AGREEMENT', bold: true, fontSize: 16, align: 'center' },
  'The parties agree to the following.',
], { at: 'end' });                        // a single ⌘Z undoes it
editor.insertTable({                     // table with content at the cursor
  rows: [['Item', 'Amount'], ['Design', '$1,200,000']],
  widths: [2, 1], align: ['left', 'right'],
});
await editor.insertImage(file, { width: 200, align: 'center', alt: 'Logo' }); // PNG, JPEG or GIF
editor.insertTable({ rows: [['Company', 'Logo'], ['Acme', { image: logoBytes, text: 'Seal' }]] });
await editor.downloadPdf('contrato.pdf');  // vector PDF with selectable text
const bytes = editor.save();             // Uint8Array of the edited .docx

editor.addEventListener('change', () => console.log(editor.stats()));
editor.addEventListener('selectionchange', e => console.log(e.detail.format));
```

| Method | What it does |
| :--- | :--- |
| `open(source, { fileName })` / `openBlank({ fileName, pageSize, focus })` / `openSample()` / `close()` | Open, create and close documents. `pageSize`: `a4` (default), `letter` or `legal` |
| `save()` / `saveBlob()` / `download(name)` | Export the edited `.docx` |
| `exportPdf({ watermark })` / `exportPdfBlob()` / `downloadPdf(name)` | Export to vector PDF (see [PDF](guides.md#pdf)). They return promises |
| `text()` / `variables()` / `stats()` / `fonts()` | Read the content. `variables()` lists the fields exactly as they are written |
| `replaceVariables(values)` / `findReplace(search, replacement, options)` | Fill templates and replace text. A key without braces fills every form of the field (`{{KEY}}`, `{{ KEY }}`, `{KEY}`, `{{KEY: hint}}`) |
| `markers({ kind, section })` / `fillMarker(marker, text)` | `{{…}}` / `{…}` fields and `[…]` pending content from the standard lexicon, with their position; fill one while keeping its formatting (see [Markers](guides.md#markers-fields-and-pending-items)) |
| `nextMarker({ kind, from, wrap })` | The first marker after a position (by default, the end of the selection); with `wrap`, it goes back to the first one when none is left after it. `null` if there are none |
| `setHighlightPending(true \| { fields: true } \| false)` / `highlightedMarkers` | Highlights the `[…]` pending items on screen (and with `fields`, the fields too). View only: it is not saved, not printed and does not appear in the PDF (also available as the `highlightPending` option; see [Highlight pending items](guides.md#highlight-the-pending-items)) |
| `markActivity(target, { label, color, ttl })` / `activities()` / `revealActivity(id)` / `clearActivity(id)` | Show where an agent or your code is working (see [Activity](guides.md#activity-where-the-agent-is-working)) |
| `insertParagraphs(paragraphs, { at })` / `insertText(text, { at })` | Insert content at the cursor (`at: 'cursor'`, the default), at the start, at the end or at an exact position (`at: { paragraph, offset }`), as a single undo step. Each paragraph is text or `{ text, bold, italic, underline, fontSize, align, style }`; `style` (`'Title'`, `'Heading1'`, `'Heading2'`, `'Heading3'`) applies the Word style and creates it if the document does not have it. Paragraphs without `style` inserted before a heading do not inherit its style |
| `structure()` | Body blocks in order: `{ kind: 'paragraph', paragraph, text, headingLevel }` (1 = Heading 1; `null` outside the outline) and `{ kind: 'table', index, first, last, rows, cols }` (paragraphs of its cells; `index` is the one from `tables()`). The basis for working by sections |
| `replaceParagraphs({ from, to }, paragraphs)` | Replaces the blocks from paragraph `from` to `to` (including the tables between them) as a single undo step. Fails if the range cuts through a table or includes a section break |
| `deleteParagraphs({ from, to })` | Deletes the blocks from paragraph `from` to `to` (including the tables between them) as a single undo step. Same errors as `replaceParagraphs`; the body keeps at least one paragraph. Returns `{ count }` |
| `moveParagraphs({ from, to }, before)` | Moves those blocks in front of paragraph `before` (numbered before the move; the paragraph count carries them to the end) without rewriting them: they keep their formatting, images, links and notes. One undo step. Fails like `deleteParagraphs` and when `before` falls inside the range or inside a table. Returns `{ first, count }` |
| `insertTable({ rows, header, widths, align }, { at })` | Insert a table with content at the cursor, at the start or at the end, as a single undo step. Each cell is text (`\n` = line break in the cell) or `{ text, image, width, height, alt }` with a PNG, JPEG or GIF image (`Uint8Array`, `ArrayBuffer` or `data:` URL) that fills the column width unless you specify `width`/`height`, with the text below it; `header` (default `true`) makes the first row bold, shaded and repeated on every page; `widths` are relative widths. Tables inside tables are not allowed |
| `sections()` / `sectionContent(key)` / `defineSection(key, { from, to }, { title, locked })` / `replaceSection(key, paragraphs, { keepHeading })` / `unwrapSection(key)` | Logical sections (Introduction, Conclusions…) marked in the `.docx` to read or fill the document section by section (see [Logical sections](guides.md#logical-sections)) |
| `sectionsFromHeadings({ fromLevel, toLevel })` / `insertSection(key, { title, after, before, inside, heading, paragraphs })` | Create sections from the document's headings or insert new sections (see [Create sections from headings](guides.md#create-sections-from-headings)) |
| `setSectionLock(key, { locked, contentLocked })` | Prevents a section from being deleted or its content from being edited by hand, in Word and in the editor (see [Locks](guides.md#locks)) |
| `setSectionSchema(schema)` / `sectionSchema()` / `validateSections()` / `addMissingSections({ optional })` | Section schema for a document type, stored in the `.docx`, its validation and the creation of the missing sections (see [Schema by document type](guides.md#schema-by-document-type)) |
| `setShowSections(true \| false)` / `sectionsVisible` | Shows the sections in the margin (also available as the `showSections` option; see [View sections in the editor](guides.md#show-sections-in-the-editor)) |
| `tables()` / `setTableCell(table, row, col, text)` / `insertTableRows(table, at, rows, { copyFormatFrom })` / `deleteTableRows(table, from, count)` | Fill in tables that are already in the document: write cells and add or remove rows, copying the formatting of an existing row, each call as one undo step (see [Existing tables](guides.md#existing-tables)) |
| `duplicateTable(table, { at, values })` | Insert a copy of a table in the document (same formatting, merged cells and text) at the cursor, at the start or at the end, and optionally change the text of its cells, as a single undo step. Returns `{ table, first }` (see [Duplicate a table](guides.md#duplicate-a-table)) |
| `insertImage(image, { at, width, height, align, alt, wrap, wrapSide, horizontal, vertical, distance })` | Insert a PNG, JPEG or GIF image (`Uint8Array`, `ArrayBuffer`, `Blob`/`File` or `data:` URL) as a single undo step. Without a size it keeps its pixels; it never exceeds the text width. Inline (default) it goes in its own paragraph; with any other `wrap` it floats, anchored to the cursor's paragraph (see [Images](guides.md#images)). Returns a promise with `{ paragraph, index }` |
| `images()` / `selectedImage` / `selectImage(ref)` | The document's images and the selected one (see [Images](guides.md#images)) |
| `updateImage(ref, changes)` / `resizeImage` / `moveImage` / `setImageWrap` / `alignImage` / `deleteImage` | Size, position and text wrapping of an image, each change as one undo step |
| `undo()` / `redo()` / `canUndo` / `canRedo` | History |
| `toggleBold()` / `toggleItalic()` / `toggleUnderline()` / `toggleStrikethrough()` / `toggleSuperscript()` / `toggleSubscript()` / `setColor(hex)` | Formatting of the selection (or of the cursor's paragraph) |
| `setFontFamily(name)` / `setFontSize(pt)` / `changeFontSize(±1)` / `fontFamilies()` | Font and size. `changeFontSize` moves to the next or previous size in Word's list; `fontFamilies()` returns the document's fonts and the usual Office ones |
| `setHighlight(color \| null)` | Highlight with one of the Word colors (any other color uses the nearest one); `null` removes it |
| `clearFormatting()` | Removes direct character formatting (font, size, bold, italic, underline, strikethrough, superscript, color and highlight) so that the paragraph style's formatting shows |
| `paragraphStyles()` / `setParagraphStyle(id)` | The document's paragraph styles for a menu (`{ id, name }`, with the Word names in the editor's language) and applying them. Title and Heading 1–3 always appear and are added to the document the first time they are used |
| `setAlignment(align)` / `setLineSpacing(n)` / `setParagraphSpacing({ before, after })` | Alignment, line spacing (multiple: 1, 1.15, 1.5, 2) and space before and after in points |
| `toggleBulletList()` / `toggleNumberedList()` / `setList('bullet' \| 'number' \| 'none')` | Bulleted or numbered lists. Right after a list of the same type they continue it; otherwise numbering starts at 1. If the document had no numbering, it is created |
| `indent()` / `outdent()` | Increase or decrease indentation: in a list they change the level; in any other paragraph they move the indent by 1.27 cm (0.5") |
| `setBackgroundColor(hex)` / `setWatermark(text, { opacity })` | Appearance. `insertTable(rows, cols, headers)` from earlier versions still works (empty table at the end) |
| `select(anchor, focus)` | Places the cursor or a selection (`{ paragraph, offset }`) |
| `setZoom(z)` / `focus()` / `destroy()` | View and lifecycle. With no previous cursor, `focus()` puts it at the start of the document |
| `setReadOnly(true \| false)` / `readOnly` | Read-only mode (see [Read-only](guides.md#read-only-mode)). By default the editor is editable |

The editor creates its own element (`editor.root`, class `docx-editor`) inside the container, and `destroy()` removes it without touching the container (`editor.container`).

| Event | `detail` |
| :--- | :--- |
| `load` | `{ fileName }` |
| `change` | `{}` |
| `selectionchange` | `{ selection, format: { bold, italic, underline, strike, superscript, subscript, color, highlight, fontFamily, fontSize, align, style, list, listLevel, lineSpacing } }` (`fontFamily` and `fontSize` are `null` if the selection mixes several) |
| `imageselect` | `{ image }`: the image selected with the mouse or `selectImage()`, or changed while selected (`null` when deselected) |
| `readonlychange` | `{ readOnly }`: `setReadOnly()` changed the mode |
| `message` | `{ message, error }`: notices to show to the user |
| `activity` | `{ activities }`: activity marks appeared, moved or ended (see `activities()`) |

Includes TypeScript types.
