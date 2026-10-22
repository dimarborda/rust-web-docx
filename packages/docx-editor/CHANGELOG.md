# Changelog

Cambios de `@dimarborda/docx-editor`. El formato sigue [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/).

## [Unreleased]

### Añadido

- `duplicateTable(table, { at, values })`: inserta una copia de una tabla del documento con su mismo formato (sombreado, bordes, anchos, celdas combinadas) y, opcionalmente, cambia el texto de sus celdas, como un solo paso de deshacer ([#2](https://github.com/dimarborda/rust-web-docx/issues/2)).
- Opción `highlightPending` y `setHighlightPending()`: resaltan en pantalla los marcadores pendientes `[…]` (con `{ fields: true }`, también los campos `{{…}}`). Es solo de vista: no se guarda, no se imprime ni sale en el PDF. Los colores se cambian con `--docx-pending-highlight` y `--docx-field-highlight` ([#3](https://github.com/dimarborda/rust-web-docx/issues/3)).
- `nextMarker({ kind, from, wrap })`: el siguiente marcador después de una posición, para ofrecer «ir al siguiente pendiente» ([#3](https://github.com/dimarborda/rust-web-docx/issues/3)).

## [0.12.1] - 2026-10-22

### Corregido

- `defineSection` y `sectionsFromHeadings` ya envuelven una sección existente que comparte el primer o el último párrafo con el rango nuevo; antes fallaban con «El rango corta otra sección…» ([#1](https://github.com/dimarborda/rust-web-docx/issues/1)).
