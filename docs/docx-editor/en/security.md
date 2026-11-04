[← Index](../README.md) · **English** · [Español](../es/security.md)

# @dimarborda/docx-editor security and limitations

[API](api.md) · [Guides](guides.md) · [Integrations](integrations.md) · [Security and limitations](security.md)

What the editor protects and which parts of Word it does not draw yet.

## Security and privacy

- **The library sends your documents nowhere.** Opening, editing, saving and exporting to PDF all happen in the page. There is no server of its own, no telemetry and no requests to third parties. The only downloads are from your own site: the `.wasm` engine, the fonts and the `.docx` from the `src` attribute, if used. Fonts the user registers with `registerFont` are stored only in that browser's IndexedDB. If your app sends content to another service, such as an AI model, that is your app's decision and outside what the library does.
- **Document content is never executed.** Pages are drawn on a `<canvas>` and text is never inserted as HTML, so a malicious `.docx` cannot inject scripts (XSS). Images are read only from the file itself; links to external resources do not trigger requests.
- **XML without external entities.** The parser (quick-xml) does not resolve DTDs or external entities, which prevents XXE attacks.
- **Limits against ZIP bombs.** Files with more than 10,000 parts, with a part larger than 256 MB, or with more than 512 MB in total when decompressed are rejected. The size is measured while decompressing, without trusting the size the ZIP declares. Images inserted through the API have a 15 MB maximum.
- **Audited dependencies.** Every deployment of the repository runs `cargo audit` and `npm audit` and stops if there are known vulnerabilities. At runtime the package depends only on `pdf-lib`, `@pdf-lib/fontkit` (MIT) and the Carlito, Caladea, Arimo, Tinos and Cousine fonts from Fontsource (SIL OFL 1.1).
- **Strict CSP.** Adding `'wasm-unsafe-eval'` to `script-src` is enough; `'unsafe-eval'` is not needed for scripts (see [Content Security Policy](integrations.md#content-security-policy)).

Vulnerabilities are reported privately, as explained in [SECURITY.md](https://github.com/dimarborda/rust-web-docx/blob/master/SECURITY.md).

## Limitations

Grouped shapes, comments, tracked changes and multiple columns are not drawn yet. All of that is preserved intact when the document is saved. Headers and footers are drawn (text, tables, images, page number and different first page) and are edited by double-clicking them, as in Word: a dotted line appears with the label "Header" or "Footer", and the text is typed, deleted and formatted like body text; Esc or a click in the body returns to it. Tables and images cannot be inserted in them. The header and footer of the last section are used (no variant for even pages), and on a line with a page number the cursor can drift if the displayed number has a different number of digits than the stored one. In the selection, their paragraphs use indices starting at 33,554,432 (2²⁵).

Text boxes: a click inside places the cursor in their text (typing, Enter, delete, formatting, undo) and a click on their border selects the box to move it, resize it or change its wrapping like an image. The selection and the arrow keys do not leave the box, and ⌘A selects all of its text. Tables and floating images cannot be inserted inside a box. Tables inside a box and box rotation are not drawn, and text that does not fit is not clipped. VML boxes from old documents are drawn but cannot be edited. In the selection (`selection`), the paragraphs of the boxes use indices starting at 16,777,216 (2²⁴), so as not to shift the body's numbering.

Floating images: text wraps around them on both sides when there is room (`bothSides`), on the indicated side (`left`, `right`, `largest`) and following their outline in tight and through wrapping (`wrapPolygon`); a table row that would collide with them moves down until it clears them. Overlapping objects are drawn in Word's order (`relativeHeight`). They only displace text that comes after their anchor paragraph on the same page, and images in headers and footers are displayed but cannot be selected.
