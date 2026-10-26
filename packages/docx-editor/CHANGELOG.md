# Changelog

Cambios de `@dimarborda/docx-editor`. El formato sigue [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/).

## [Unreleased]

### Corregido

- `markers()`, `fillMarker`/`replaceVariables` y `fillTextBox` ya reconocen y rellenan marcadores dentro de un cuadro de texto flotante que está DENTRO de un encabezado o pie de página (membretes institucionales con logos en cuadros de texto del encabezado, típicos de plantillas de VEAD); antes se perdían en silencio, aunque el mismo tipo de cuadro de texto en el cuerpo del documento ya funcionaba. Nuevo valor de `Marker.place`: `"headerFooterTextbox"`.

### Añadido

- `duplicateTable(table, { at, values })`: inserta una copia de una tabla del documento con su mismo formato (sombreado, bordes, anchos, celdas combinadas) y, opcionalmente, cambia el texto de sus celdas, como un solo paso de deshacer ([#2](https://github.com/dimarborda/rust-web-docx/issues/2)).
- Opción `highlightPending` y `setHighlightPending()`: resaltan en pantalla los marcadores pendientes `[…]` (con `{ fields: true }`, también los campos `{{…}}`). Es solo de vista: no se guarda, no se imprime ni sale en el PDF. Los colores se cambian con `--docx-pending-highlight` y `--docx-field-highlight` ([#3](https://github.com/dimarborda/rust-web-docx/issues/3)).
- `nextMarker({ kind, from, wrap })`: el siguiente marcador después de una posición, para ofrecer «ir al siguiente pendiente» ([#3](https://github.com/dimarborda/rust-web-docx/issues/3)).
- `exportPdf()`/`exportPdfBlob()`/`downloadPdf()` conservan los hipervínculos de Word como anotaciones `/Link` reales del PDF (acción URI), en la posición exacta del texto enlazado — igual que LibreOffice o el propio exportador de Word. Necesario para flujos de firma electrónica que leen la posición de la firma de un hipervínculo del PDF (p. ej. VEAD).
- `fillTextBox(marker, ooxml, images?)`: escribe contenido (párrafos, tablas, imágenes) dentro de un cuadro de texto que ya existe en el documento, ubicado por su marcador (`{{...}}` con `place: "textbox"`), reemplazando lo que tenía, como un solo paso de deshacer. Las imágenes se referencian con un token `{{img:placeholder}}` en el `ooxml` y se registran en el paquete igual que `insertImage`. Pensado para plantillas con campos `html` (p. ej. VEAD), cuyo contenido ya llega convertido a OOXML.

## [0.12.1] - 2026-10-22

### Corregido

- `defineSection` y `sectionsFromHeadings` ya envuelven una sección existente que comparte el primer o el último párrafo con el rango nuevo; antes fallaban con «El rango corta otra sección…» ([#1](https://github.com/dimarborda/rust-web-docx/issues/1)).
