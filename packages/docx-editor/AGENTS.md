# @dimarborda/docx-editor — guide for AI coding agents

Read this before writing code that uses this package. It is the short, complete map of the
library; the exact signatures and JSDoc are in `src/index.d.ts` next to this file.

## What it is

A Word (`.docx`) editor for the browser. A Rust + WebAssembly engine reads the file, lays out the
pages like Word and draws them on a `<canvas>`; edits are written back losslessly (what is not
touched stays as it was). Framework-agnostic: a class (`DocxEditor`) and two custom elements
(`<docx-editor>`, `<docx-toolbar>`). No server, no network calls of its own.

- **Browser only.** There is no headless/Node API for `.docx` (unlike `@dimarborda/xlsx-editor`,
  which has `/engine`). Importing is SSR-safe; creating the editor needs the DOM.
- The editor's container **must have a height** (`height: 80vh`, or a sized flex/grid child).
- In production with a CSP, `script-src` needs `'wasm-unsafe-eval'`.

## Setup

```js
import { DocxEditor } from '@dimarborda/docx-editor';
import '@dimarborda/docx-editor/style.css';
import '@dimarborda/docx-editor/fonts';   // optional: metric-compatible Calibri, Cambria, Arial…

const editor = await DocxEditor.create(container, { locale: 'en' }); // locale: 'es' (default) | 'en'
await editor.open(fileOrBytes);   // File, Blob, ArrayBuffer or Uint8Array
// or: editor.openBlank({ pageSize: 'a4' | 'letter' | 'legal' })

const docx = editor.save();                 // Uint8Array of the edited .docx
const pdf  = await editor.exportPdf();      // Uint8Array, vector PDF (pdf-lib loads lazily)
editor.destroy();                           // removes editor.root, leaves the container alone
```

Web component alternative:

```html
<docx-toolbar for="doc"></docx-toolbar>
<docx-editor id="doc" src="/contract.docx" locale="en" style="height: 80vh"></docx-editor>
<script type="module">
  import '@dimarborda/docx-editor';
  import '@dimarborda/docx-editor/style.css';
  const editor = await document.getElementById('doc').whenReady();
</script>
```

React/Next.js: create it in `useEffect`, and in StrictMode destroy the instance that resolves after
cleanup (`if (disposed) return e.destroy()`). If the bundler does not serve the `.wasm`, pass
`{ wasmUrl }` (file: `node_modules/@dimarborda/docx-editor/wasm/docx_engine_bg.wasm`).

## Mental model (get these right)

- **Positions are paragraph indexes** in body order, from 0, as in `structure()` and `selection`.
  Table cells count as paragraphs (`tables()[i].first/last`). Text boxes use indexes ≥ 2²⁴ and
  headers/footers ≥ 2²⁵ — never compute those by hand.
- **Indexes shift** after inserting or deleting. Re-read `structure()`, `tables()` or `markers()`
  after each structural change instead of reusing old numbers. Exceptions: `fillMarker` relocates
  its marker, and activity marks and section keys follow their content.
- **Prefer stable handles over indexes:** section keys (`sections()`), markers (`markers()`), table
  header text (`tables().find(t => t.header[0] === 'Risk')`).
- **Every API edit is one undo step**, fires `change`, flashes the changed area for the user and
  does not move the user's caret or scroll. API edits also work in read-only mode and inside
  locked sections (locks restrict the person, not code).
- **Errors are atomic:** invalid calls throw an `Error` with a clear message and leave the document
  untouched. Catch, re-read state, and retry differently.

## Reading a document

| Call | Returns |
| :--- | :--- |
| `text()` | Plain text of the whole document |
| `structure()` | Body blocks in order: `{ kind: 'paragraph', paragraph, text, headingLevel }` and `{ kind: 'table', index, first, last, rows, cols }` |
| `markers({ kind, section })` | Fields and pending items (see below), with `paragraph`, `offset`, `length`, `place`, `section` |
| `variables()` | Field texts exactly as written, e.g. `['{{CLIENT}}', '{DATE}']` |
| `sections()` / `sectionContent(key)` | Logical sections and their text |
| `tables()` | `{ index, rows, cols, header, first, last }` per body table |
| `images()` | Images, text boxes and shapes with size, wrap and position |
| `stats()` / `fonts()` / `paragraphStyles()` | Counts, fonts used, available paragraph styles |

## Markers: what is left to fill in

Standard lexicon shared with `@dimarborda/xlsx-editor`:

| Written as | `kind` | Meaning |
| :--- | :--- | :--- |
| `{{KEY}}`, `{{ KEY }}`, `{{KEY: hint}}`, `{key}`, `{key:hint}` | `field` | A value an automation fills in |
| `[…]`, `[...]`, `[____]`, `[Complete: project scope]` | `pending` | Prose someone still has to write; the inner text is the hint |

Not markers: citations (`[1]`, `[2, 3]`), checkboxes (`[ ]`, `[x]`), `[sic]`. No marker spans a
line break.

```js
editor.replaceVariables({ CLIENT: 'Acme Inc.', DATE: 'October 7, 2026' }); // key without braces fills every form
for (const m of editor.markers({ kind: 'pending' })) {
  editor.fillMarker(m, await draft(m.hint, m.section));   // keeps the marker's formatting
}
editor.nextMarker({ kind: 'pending', wrap: true });        // first marker after the selection, or null
```

`findMarkers(text)` (exported) checks a loose string — e.g. text you are about to insert — for
leftover markers.

## Writing content

```js
editor.insertParagraphs([
  { text: 'Scope', style: 'Heading1' },                 // 'Title' | 'Heading1' | 'Heading2' | 'Heading3'
  { text: 'Important', bold: true, align: 'center', fontSize: 14 },
  'A plain paragraph.',
], { at: 'end' });                                       // 'cursor' (default) | 'start' | 'end' | { paragraph, offset }

editor.insertText('inline text', { at: { paragraph: 4, offset: 10 } });
editor.replaceParagraphs({ from: 3, to: 7 }, ['New text']);   // whole blocks, tables in between included
editor.deleteParagraphs({ from: 3, to: 7 });
editor.moveParagraphs({ from: 3, to: 7 }, 12);                // keeps formatting, images, links
editor.findReplace('Bogota', 'Bogotá', { matchCase: true });
```

Ranges fail if they cut through a table or include a Word section break.

## Logical sections (work section by section)

Sections are block content controls in the `.docx` (`w:sdt`, `w:tag="section:<key>"`), so they
survive Word round-trips.

```js
editor.sectionsFromHeadings({ fromLevel: 1, toLevel: 2 }); // create from Heading styles; idempotent
editor.defineSection('introduction', { from: 2, to: 6 }, { title: 'Introduction' });
editor.sectionContent('introduction').text;
editor.replaceSection('introduction', ['New paragraph.'], { keepHeading: true });
editor.insertSection('annexes', { title: 'Annexes', after: 'conclusions' });
editor.setSectionLock('legal', { contentLocked: true });   // user cannot edit; code still can

editor.setSectionSchema({ type: 'report', sections: [{ key: 'introduction', title: 'Introduction' }] });
editor.validateSections();     // { valid, issues: [{ kind: 'missing' | 'empty' | 'order' | … }] }
editor.addMissingSections();
```

## Tables

```js
editor.insertTable({ rows: [['Item', 'Amount'], ['Design', '$1,200']], widths: [2, 1], align: ['left', 'right'] }, { at: 'end' });

const t = editor.tables().find(t => t.header[0] === 'Risk');       // existing template table
const { first } = editor.insertTableRows(t.index, t.rows, [['Outage', 'Medium']], { copyFormatFrom: 1 });
editor.setTableCell(t.index, first, 1, 'High');                     // keeps the cell's text format
editor.deleteTableRows(t.index, first, 1);
editor.duplicateTable(t.index, { at: 'end', values: [[null, 'New value']] }); // null keeps text
```

- Rows, columns and tables are 0-based; for a merged cell write to its first grid column.
- When inserting right under a header row, pass `copyFormatFrom` with a **data** row, otherwise the
  header shading is copied.
- To reproduce a styled template table (record card, signature block) use `duplicateTable`, not
  `insertTable` (which uses the editor's own table style). Nested tables are not supported.

## Images

```js
await editor.insertImage(bytes, { width: 200, align: 'center', alt: 'Logo' }); // PNG, JPEG, GIF; ≤ 15 MB
const [logo] = editor.images();                  // ref = { paragraph, index }
editor.updateImage(logo, { wrap: 'square', horizontal: { relativeTo: 'margin', align: 'right' } });
```

Sizes are CSS px at 100 % zoom. `wrap`: `inline` (default) | `square` | `tight` | `through` |
`topAndBottom` | `behind` | `inFront`.

## Showing the user what an agent does

```js
const work = editor.markActivity('conclusions', { label: 'Agent · drafting' }); // section key, range or span
const added = editor.replaceSection('conclusions', ['First draft…']);
work.update(added, { label: 'Agent · reviewing' });
work.done();
```

View-only: never saved or printed. Use `color` to distinguish several agents and `ttl` to expire.
`setHighlightPending({ fields: true })` highlights what is still pending.

## Events

`load`, `change`, `selectionchange` (`detail.format`), `imageselect`, `readonlychange`, `message`
(user-facing notices), `activity`. Save on `change`.

## Do and don't

- Do read the state (`structure()`, `markers()`, `sections()`) before editing, and again after.
- Do fill templates with `replaceVariables` / `fillMarker` rather than `findReplace` on braces:
  they keep the original formatting and handle every field syntax.
- Don't build `.docx` XML yourself or unzip the file to edit it: use the API, which keeps the
  document lossless and undoable.
- Don't assume Word section breaks, comments, tracked changes, multiple columns or grouped shapes
  are editable: they are preserved but not drawn yet. Tables and images cannot go inside headers,
  footers or text boxes.
- Don't send the document anywhere from library code; if your app sends content to a model, that
  is the app's decision.

## More

Full docs: https://github.com/dimarborda/rust-web-docx/tree/master/docs/docx-editor/en
(API, guides, integrations, security). Demo: https://rust-web-docx.dimarborda.workers.dev
