# API de @dimarborda/docx-editor

[← Índice](README.md) · [API](api.md) · [Guías](guias.md) · [Integraciones](integraciones.md) · [Seguridad y limitaciones](seguridad.md)

Métodos, opciones y eventos de `DocxEditor`. Los tipos completos están en `src/index.d.ts` del paquete.

## Web component

Atributos de `<docx-editor>`: `src` (URL del .docx), `sample` (abre el contrato de ejemplo), `blank` (documento nuevo vacío, con `page-size="a4|letter|legal"` y `autofocus` opcionales), `locale` (`es` o `en`), `zoom`, `gridlines="false"`, `page-labels="false"`, `readonly` (solo lectura; sin el atributo es editable) y `wasm-url`.

`<docx-toolbar>` es opcional. Acepta `for="id-del-editor"` e `items` con los controles separados por espacios y `|` como separador; por defecto muestra todos: `undo redo | style font fontSize | bold italic underline strike superscript subscript color highlight clearFormat | left center right both lineSpacing | bulletList numberedList outdent indent | imageWrap | zoom`. `style` es la lista de estilos de párrafo, `font` y `fontSize` aceptan cualquier nombre o tamaño escrito, `highlight` aplica el último color de resaltado (la flecha abre los colores de Word) y `imageWrap` es un selector del ajuste de texto que se activa al seleccionar una imagen; con una imagen seleccionada, los botones de alineación alinean la imagen.

## DocxEditor

```js
import { DocxEditor } from '@dimarborda/docx-editor';
import '@dimarborda/docx-editor/style.css';

const editor = await DocxEditor.create(document.querySelector('#editor'), { locale: 'es' });

await editor.open(file);                 // File, Blob, ArrayBuffer o Uint8Array
editor.openBlank({ pageSize: 'letter' }); // o un documento nuevo vacío, con el cursor listo
editor.variables();                      // ['{{CLIENTE}}', '{{FECHA}}']
editor.replaceVariables({ CLIENTE: 'Acme' });
editor.findReplace('Bogota', 'Bogotá', { matchCase: true });
editor.insertParagraphs([
  { text: 'CONTRATO DE SERVICIOS', bold: true, fontSize: 16, align: 'center' },
  'Entre las partes se acuerda lo siguiente.',
], { at: 'end' });                        // un solo ⌘Z lo deshace
editor.insertTable({                     // tabla con contenido en el cursor
  rows: [['Ítem', 'Valor'], ['Diseño', '$1.200.000']],
  widths: [2, 1], align: ['left', 'right'],
});
await editor.insertImage(file, { width: 200, align: 'center', alt: 'Logo' }); // PNG, JPEG o GIF
editor.insertTable({ rows: [['Empresa', 'Logo'], ['Acme', { image: logoBytes, text: 'Sello' }]] });
await editor.downloadPdf('contrato.pdf');  // PDF vectorial con texto seleccionable
const bytes = editor.save();             // Uint8Array del .docx editado

editor.addEventListener('change', () => console.log(editor.stats()));
editor.addEventListener('selectionchange', e => console.log(e.detail.format));
```

| Método | Qué hace |
| :--- | :--- |
| `open(source, { fileName })` / `openBlank({ fileName, pageSize, focus })` / `openSample()` / `close()` | Abrir, crear y cerrar documentos. `pageSize`: `a4` (predeterminado), `letter` o `legal` |
| `save()` / `saveBlob()` / `download(name)` | Exportar el `.docx` editado |
| `exportPdf({ watermark })` / `exportPdfBlob()` / `downloadPdf(name)` | Exportar a PDF vectorial (ver [PDF](guias.md#pdf)). Devuelven promesas |
| `text()` / `variables()` / `stats()` / `fonts()` | Leer el contenido. `variables()` lista los campos tal como están escritos |
| `replaceVariables(values)` / `findReplace(search, replacement, options)` | Rellenar plantillas y reemplazar texto. Una clave sin llaves rellena todas las formas del campo (`{{CLAVE}}`, `{{ CLAVE }}`, `{CLAVE}`, `{{CLAVE: pista}}`) |
| `markers({ kind, section })` / `fillMarker(marker, text)` | Campos `{{…}}` / `{…}` y contenido pendiente `[…]` del léxico estándar, con su posición; rellenar uno conservando su formato (ver [Marcadores](guias.md#marcadores-campos-y-pendientes)) |
| `nextMarker({ kind, from, wrap })` | El primer marcador después de una posición (por defecto, el final de la selección); con `wrap`, vuelve al primero cuando no queda ninguno después. `null` si no hay |
| `setHighlightPending(true \| { fields: true } \| false)` / `highlightedMarkers` | Resalta en pantalla los pendientes `[…]` (y con `fields`, también los campos). Solo de vista: no se guarda, no se imprime ni sale en el PDF (también con la opción `highlightPending`; ver [Resaltar los pendientes](guias.md#resaltar-los-pendientes)) |
| `markActivity(target, { label, color, ttl })` / `activities()` / `revealActivity(id)` / `clearActivity(id)` | Mostrar dónde está trabajando un agente o tu código (ver [Actividad](guias.md#actividad-dónde-está-trabajando-el-agente)) |
| `insertParagraphs(paragraphs, { at })` / `insertText(text, { at })` | Insertar contenido en el cursor (`at: 'cursor'`, predeterminado), al inicio, al final o en una posición exacta (`at: { paragraph, offset }`), como un solo paso de deshacer. Cada párrafo es texto o `{ text, bold, italic, underline, fontSize, align, style }`; `style` (`'Title'`, `'Heading1'`, `'Heading2'`, `'Heading3'`) aplica el estilo de Word y lo crea si el documento no lo tiene. Los párrafos sin `style` insertados antes de un título no heredan su estilo |
| `structure()` | Bloques del cuerpo en orden: `{ kind: 'paragraph', paragraph, text, headingLevel }` (1 = Título 1; `null` fuera del esquema) y `{ kind: 'table', index, first, last, rows, cols }` (párrafos de sus celdas; `index` es el de `tables()`). Base para trabajar por secciones |
| `replaceParagraphs({ from, to }, paragraphs)` | Reemplaza los bloques del párrafo `from` al `to` (incluidas las tablas entre ellos) como un solo paso de deshacer. Falla si el rango corta una tabla o incluye un salto de sección |
| `insertTable({ rows, header, widths, align }, { at })` | Insertar una tabla con contenido en el cursor, al inicio o al final, como un solo paso de deshacer. Cada celda es un texto (`\n` = salto de línea en la celda) o `{ text, image, width, height, alt }` con una imagen PNG, JPEG o GIF (`Uint8Array`, `ArrayBuffer` o URL `data:`) que ocupa el ancho de la columna salvo que indiques `width`/`height`, con el texto debajo; `header` (predeterminado `true`) pone la primera fila en negrita, sombreada y repetida en cada página; `widths` son anchos relativos. No se permiten tablas dentro de tablas |
| `sections()` / `sectionContent(key)` / `defineSection(key, { from, to }, { title, locked })` / `replaceSection(key, paragraphs, { keepHeading })` / `unwrapSection(key)` | Secciones lógicas (Introducción, Conclusiones…) marcadas en el `.docx` para leer o llenar el documento sección por sección (ver [Secciones lógicas](guias.md#secciones-lógicas)) |
| `sectionsFromHeadings({ fromLevel, toLevel })` / `insertSection(key, { title, after, before, inside, heading, paragraphs })` | Crear secciones a partir de los títulos del documento o insertar secciones nuevas (ver [Crear secciones a partir de los títulos](guias.md#crear-secciones-a-partir-de-los-títulos)) |
| `setSectionLock(key, { locked, contentLocked })` | Impide borrar una sección o editar su contenido a mano, en Word y en el editor (ver [Bloqueos](guias.md#bloqueos)) |
| `setSectionSchema(schema)` / `sectionSchema()` / `validateSections()` / `addMissingSections({ optional })` | Esquema de secciones de un tipo de documento, guardado en el `.docx`, su validación y la creación de las secciones que falten (ver [Esquema por tipo de documento](guias.md#esquema-por-tipo-de-documento)) |
| `setShowSections(true \| false)` / `sectionsVisible` | Muestra las secciones en el margen (también con la opción `showSections`; ver [Ver las secciones en el editor](guias.md#ver-las-secciones-en-el-editor)) |
| `tables()` / `setTableCell(table, row, col, text)` / `insertTableRows(table, at, rows, { copyFormatFrom })` / `deleteTableRows(table, from, count)` | Completar tablas que ya están en el documento: escribir celdas y agregar o quitar filas copiando el formato de una fila existente, cada llamada como un paso de deshacer (ver [Tablas existentes](guias.md#tablas-existentes)) |
| `duplicateTable(table, { at, values })` | Insertar una copia de una tabla del documento (mismo formato, celdas combinadas y texto) en el cursor, al inicio o al final, y opcionalmente cambiar el texto de sus celdas, como un solo paso de deshacer. Devuelve `{ table, first }` (ver [Duplicar una tabla](guias.md#duplicar-una-tabla)) |
| `insertImage(image, { at, width, height, align, alt, wrap, wrapSide, horizontal, vertical, distance })` | Insertar una imagen PNG, JPEG o GIF (`Uint8Array`, `ArrayBuffer`, `Blob`/`File` o URL `data:`) como un solo paso de deshacer. Sin tamaño conserva sus píxeles; nunca supera el ancho del texto. En línea (predeterminado) va en su propio párrafo; con otro `wrap` queda flotante, anclada al párrafo del cursor (ver [Imágenes](guias.md#imágenes)). Devuelve una promesa con `{ paragraph, index }` |
| `images()` / `selectedImage` / `selectImage(ref)` | Imágenes del documento y la seleccionada (ver [Imágenes](guias.md#imágenes)) |
| `updateImage(ref, changes)` / `resizeImage` / `moveImage` / `setImageWrap` / `alignImage` / `deleteImage` | Tamaño, posición y ajuste del texto de una imagen, cada cambio como un paso de deshacer |
| `undo()` / `redo()` / `canUndo` / `canRedo` | Historial |
| `toggleBold()` / `toggleItalic()` / `toggleUnderline()` / `toggleStrikethrough()` / `toggleSuperscript()` / `toggleSubscript()` / `setColor(hex)` | Formato de la selección (o del párrafo del cursor) |
| `setFontFamily(name)` / `setFontSize(pt)` / `changeFontSize(±1)` / `fontFamilies()` | Fuente y tamaño. `changeFontSize` pasa al tamaño siguiente o anterior de la lista de Word; `fontFamilies()` da las fuentes del documento y las habituales de Office |
| `setHighlight(color \| null)` | Resaltado con uno de los colores de Word (otro color usa el más cercano); `null` lo quita |
| `clearFormatting()` | Quita el formato directo de carácter (fuente, tamaño, negrita, cursiva, subrayado, tachado, superíndice, color y resaltado) para que se vea el del estilo del párrafo |
| `paragraphStyles()` / `setParagraphStyle(id)` | Estilos de párrafo del documento para un menú (`{ id, name }`, con los nombres de Word en el idioma del editor) y aplicarlos. Título y Título 1–3 siempre aparecen y se añaden al documento la primera vez que se usan |
| `setAlignment(align)` / `setLineSpacing(n)` / `setParagraphSpacing({ before, after })` | Alineación, interlineado (múltiplo: 1, 1.15, 1.5, 2) y espacio antes y después en puntos |
| `toggleBulletList()` / `toggleNumberedList()` / `setList('bullet' \| 'number' \| 'none')` | Listas con viñetas o numeradas. Justo después de una lista del mismo tipo la continúan; si no, la numeración empieza en 1. Si el documento no tenía numeración, se crea |
| `indent()` / `outdent()` | Aumentar o disminuir sangría: en una lista cambian el nivel; en otro párrafo mueven la sangría 1,27 cm (0,5") |
| `setBackgroundColor(hex)` / `setWatermark(text, { opacity })` | Apariencia. `insertTable(rows, cols, headers)` de versiones anteriores sigue funcionando (tabla vacía al final) |
| `select(anchor, focus)` | Coloca el cursor o una selección (`{ paragraph, offset }`) |
| `setZoom(z)` / `focus()` / `destroy()` | Vista y ciclo de vida. Sin cursor previo, `focus()` lo pone al inicio del documento |
| `setReadOnly(true \| false)` / `readOnly` | Modo de solo lectura (ver [Solo lectura](guias.md#solo-lectura)). Por defecto el editor es editable |

El editor crea su propio elemento (`editor.root`, clase `docx-editor`) dentro del contenedor y `destroy()` lo elimina sin tocar el contenedor (`editor.container`).

| Evento | `detail` |
| :--- | :--- |
| `load` | `{ fileName }` |
| `change` | `{}` |
| `selectionchange` | `{ selection, format: { bold, italic, underline, strike, superscript, subscript, color, highlight, fontFamily, fontSize, align, style, list, listLevel, lineSpacing } }` (`fontFamily` y `fontSize` son `null` si la selección mezcla varios) |
| `imageselect` | `{ image }`: la imagen seleccionada con el ratón o `selectImage()`, o cambiada mientras está seleccionada (`null` al deseleccionar) |
| `readonlychange` | `{ readOnly }`: `setReadOnly()` cambió el modo |
| `message` | `{ message, error }`: avisos para mostrar al usuario |
| `activity` | `{ activities }`: aparecieron, se movieron o terminaron marcas de actividad (ver `activities()`) |

Incluye tipos de TypeScript.
