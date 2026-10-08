# @dimarborda/docx-editor

[![npm](https://img.shields.io/npm/v/@dimarborda/docx-editor?color=cb3837&label=npm)](https://www.npmjs.com/package/@dimarborda/docx-editor)

Editor de documentos Word (`.docx`) para el navegador. El motor está escrito en Rust y compilado a WebAssembly: lee el documento, maqueta las páginas como Word y aplica ediciones sin pérdidas. Las páginas se dibujan en canvas, con cursor y selección propios.

Funciona en cualquier frontend (sin framework, React, Vue, Svelte, Angular) y dentro de apps de escritorio con webview, como Tauri. No necesita servidor: los documentos nunca salen del equipo.

[Demo](https://rust-web-docx.dimarborda.workers.dev/?demo) · [Repositorio](https://github.com/dimarborda/rust-web-docx)

```bash
npm install @dimarborda/docx-editor
```

## Uso rápido: web component

```html
<docx-toolbar for="doc"></docx-toolbar>
<docx-editor id="doc" src="/plantillas/contrato.docx" style="height: 80vh"></docx-editor>

<script type="module">
  import '@dimarborda/docx-editor';          // registra <docx-editor> y <docx-toolbar>
  import '@dimarborda/docx-editor/style.css';
  import '@dimarborda/docx-editor/fonts';    // opcional: sustitutos de Calibri, Cambria, Arial…

  const editor = await document.getElementById('doc').whenReady();
  editor.replaceVariables({ CLIENTE: 'Acme S.A.S.', FECHA: '7 de octubre de 2026' });
</script>
```

Atributos de `<docx-editor>`: `src` (URL del .docx), `sample` (abre el contrato de ejemplo), `blank` (documento nuevo vacío, con `page-size="a4|letter|legal"` y `autofocus` opcionales), `locale` (`es` o `en`), `zoom`, `gridlines="false"`, `page-labels="false"` y `wasm-url`.

`<docx-toolbar>` es opcional. Acepta `for="id-del-editor"` e `items="undo redo | bold italic underline color | left center right both | zoom"`.

## Uso desde JavaScript

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
const bytes = editor.save();             // Uint8Array del .docx editado

editor.addEventListener('change', () => console.log(editor.stats()));
editor.addEventListener('selectionchange', e => console.log(e.detail.format));
```

| Método | Qué hace |
| :--- | :--- |
| `open(source, { fileName })` / `openBlank({ fileName, pageSize, focus })` / `openSample()` / `close()` | Abrir, crear y cerrar documentos. `pageSize`: `a4` (predeterminado), `letter` o `legal` |
| `save()` / `saveBlob()` / `download(name)` | Exportar el `.docx` editado |
| `text()` / `variables()` / `stats()` / `fonts()` | Leer el contenido |
| `replaceVariables(values)` / `findReplace(search, replacement, options)` | Rellenar plantillas y reemplazar texto |
| `insertParagraphs(paragraphs, { at })` / `insertText(text, { at })` | Insertar contenido en el cursor (`at: 'cursor'`, predeterminado), al inicio o al final, como un solo paso de deshacer. Cada párrafo es texto o `{ text, bold, italic, underline, fontSize, align }` |
| `insertTable({ rows, header, widths, align }, { at })` | Insertar una tabla con contenido en el cursor, al inicio o al final, como un solo paso de deshacer. `rows` son las celdas como texto (`\n` = salto de línea en la celda); `header` (predeterminado `true`) pone la primera fila en negrita, sombreada y repetida en cada página; `widths` son anchos relativos. No se permiten tablas dentro de tablas |
| `insertImage(image, { at, width, height, align, alt, wrap, wrapSide, horizontal, vertical, distance })` | Insertar una imagen PNG, JPEG o GIF (`Uint8Array`, `ArrayBuffer`, `Blob`/`File` o URL `data:`) como un solo paso de deshacer. Sin tamaño conserva sus píxeles; nunca supera el ancho del texto. En línea (predeterminado) va en su propio párrafo; con otro `wrap` queda flotante, anclada al párrafo del cursor (ver [Imágenes](#imágenes)). Devuelve una promesa con `{ paragraph, index }` |
| `images()` / `selectedImage` / `selectImage(ref)` | Imágenes del documento y la seleccionada (ver [Imágenes](#imágenes)) |
| `updateImage(ref, changes)` / `resizeImage` / `moveImage` / `setImageWrap` / `alignImage` / `deleteImage` | Tamaño, posición y ajuste del texto de una imagen, cada cambio como un paso de deshacer |
| `undo()` / `redo()` / `canUndo` / `canRedo` | Historial |
| `toggleBold()` / `toggleItalic()` / `toggleUnderline()` / `setColor(hex)` / `setAlignment(align)` | Formato de la selección |
| `setBackgroundColor(hex)` / `setWatermark(text, { opacity })` | Apariencia. `insertTable(rows, cols, headers)` de versiones anteriores sigue funcionando (tabla vacía al final) |
| `select(anchor, focus)` | Coloca el cursor o una selección (`{ paragraph, offset }`) |
| `setZoom(z)` / `focus()` / `destroy()` | Vista y ciclo de vida. Sin cursor previo, `focus()` lo pone al inicio del documento |

El editor crea su propio elemento (`editor.root`, clase `docx-editor`) dentro del contenedor y `destroy()` lo elimina sin tocar el contenedor (`editor.container`).

| Evento | `detail` |
| :--- | :--- |
| `load` | `{ fileName }` |
| `change` | `{}` |
| `selectionchange` | `{ selection, format: { bold, italic, underline, color, align } }` |
| `imageselect` | `{ image }`: la imagen seleccionada con el ratón o `selectImage()`, o cambiada mientras está seleccionada (`null` al deseleccionar) |
| `message` | `{ message, error }`: avisos para mostrar al usuario |

Incluye tipos de TypeScript.

### Imágenes

Con el ratón: un clic selecciona la imagen y muestra ocho asas. Las esquinas cambian el tamaño manteniendo la proporción (`Shift` la libera) y los lados cambian una sola medida. Una imagen flotante se mueve arrastrándola o con las flechas (`Shift` + flecha = 10 px). `Supr` la elimina y `Esc` vuelve al texto. Una imagen detrás del texto solo se selecciona donde no hay texto encima, así que hacer clic sobre el texto de un membrete sigue colocando el cursor.

Desde código, una imagen se identifica con `{ paragraph, index }`: su párrafo y su posición entre las imágenes de ese párrafo. Cualquier elemento de `images()` sirve como referencia. Las medidas están en px CSS a zoom 100 %.

```js
const [logo] = editor.images();
// { paragraph, index, width, height, wrap, wrapSide, anchored, horizontal, vertical,
//   distance: { top, bottom, left, right }, alt, textOffset, bounds: { page, x, y, width, height } }

editor.resizeImage(logo, { width: 180 });            // el alto sigue la proporción
editor.resizeImage(logo, { scale: 0.5 });
editor.setImageWrap(logo, 'square', { side: 'bothSides', distance: 12 });
editor.alignImage(logo, 'right');                    // flotante: dentro de los márgenes
editor.moveImage(logo, { dx: 0, dy: 40 });           // o { x, y } en coordenadas de la página
editor.updateImage(logo, {                           // varios cambios en un solo paso de deshacer
  wrap: 'topAndBottom',
  horizontal: { relativeTo: 'margin', align: 'center' },
  vertical: { relativeTo: 'paragraph', offset: 10 },
  alt: 'Logo de la empresa',
});
editor.deleteImage(logo);

await editor.insertImage(file, {                     // flotante desde el inicio
  width: 160, wrap: 'square',
  horizontal: { relativeTo: 'margin', align: 'right' },
  vertical: { relativeTo: 'paragraph', offset: 0 },
});
```

| `wrap` | Comportamiento con el texto |
| :--- | :--- |
| `inline` | En línea con el texto: se comporta como un carácter grande que se apoya en la línea base y pasa a la línea siguiente si no cabe |
| `square` | El texto rodea el rectángulo de la imagen. `wrapSide`: `bothSides` (el texto usa el lado más ancho), `left`, `right` o `largest` |
| `tight` / `through` | Como `square`; se guarda con el contorno de Word, pero se maqueta usando el rectángulo |
| `topAndBottom` | El texto solo va arriba y abajo de la imagen |
| `behind` / `inFront` | Flota detrás o delante del texto, que la ignora |

`horizontal.relativeTo` puede ser `margin`, `page`, `column`, `character`, `leftMargin`, `rightMargin`, `insideMargin` u `outsideMargin`; `vertical.relativeTo`, `margin`, `page`, `paragraph`, `line`, `topMargin`, `bottomMargin`, `insideMargin` u `outsideMargin`. En cada eje se usa un `offset` (distancia desde el borde izquierdo o superior del marco) o un `align` (`left`/`center`/`right` o `top`/`center`/`bottom`). Al mover una imagen con offset se conserva su marco; un eje alineado pasa a ser relativo a la página. Para mover una imagen en línea, `moveImage` la convierte antes en `square`.

### React

```jsx
import { useEffect, useRef } from 'react';
import { DocxEditor } from '@dimarborda/docx-editor';
import '@dimarborda/docx-editor/style.css';

export function DocxView({ file, onChange }) {
  const ref = useRef(null);
  useEffect(() => {
    let editor;
    let disposed = false;
    DocxEditor.create(ref.current).then(async e => {
      // En StrictMode el efecto se monta dos veces: descarta el editor que llega tarde
      if (disposed) return e.destroy();
      editor = e;
      editor.addEventListener('change', () => onChange?.(editor));
      if (file) await editor.open(file);
      else editor.openBlank();
    });
    return () => {
      disposed = true;
      editor?.destroy();
    };
  }, [file]);
  return <div ref={ref} style={{ height: '80vh' }} />;
}
```

### Tauri

Dentro de la ventana de la app se usa igual que en la web. Lo único que cambia es abrir y guardar con los plugins `dialog` y `fs`:

```js
import { open, save } from '@tauri-apps/plugin-dialog';
import { readFile, writeFile } from '@tauri-apps/plugin-fs';

const path = await open({ filters: [{ name: 'Word', extensions: ['docx'] }] });
await editor.open(await readFile(path), { fileName: path.split(/[\\/]/).pop() });

const target = await save({ defaultPath: editor.fileName });
if (target) await writeFile(target, editor.save());
```

## Fuentes

Si el documento usa Calibri, Cambria, Arial, Times New Roman o Courier y esas fuentes no están instaladas, el editor dibuja sustitutos con las mismas medidas: Carlito, Caladea, Arimo, Tinos y Cousine. Para que estén disponibles sin depender del sistema, importa `@dimarborda/docx-editor/fonts`; incluye solo los subconjuntos latinos.

Para que un documento se vea exactamente igual que en Word, registra la fuente real:

```js
import { registerFont, loadStoredFonts } from '@dimarborda/docx-editor';

await loadStoredFonts();        // al iniciar: las fuentes guardadas antes en este navegador
await registerFont(fileInput.files[0]);  // .ttf/.otf/.woff/.woff2; queda guardada en IndexedDB
```

## Estilos

Los estilos se personalizan con variables CSS, definidas en el contenedor del editor o en cualquier ancestro:
- **Editor:** `--docx-page-gap`, `--docx-page-shadow`, `--docx-page-border`, `--docx-label-color`, `--docx-caret-color`, `--docx-selection`, `--docx-image-frame` (marco de la imagen seleccionada).
- **Barra de herramientas:** `--docx-toolbar-bg`, `--docx-toolbar-border`, `--docx-toolbar-color`, `--docx-toolbar-hover`, `--docx-toolbar-active-bg`, `--docx-toolbar-active-color`.

El contenedor del editor necesita una altura (por ejemplo `height: 80vh`, o ser un elemento flex/grid con altura definida); el editor ocupa el 100 % y dentro de esa altura hace su propio scroll.

## Bundlers

El paquete carga `docx_engine_bg.wasm` con `new URL(..., import.meta.url)`. Vite, webpack 5, Rollup y esbuild copian ese archivo junto al JavaScript sin configuración extra. Si lo sirves desde otro lugar, pásalo con `DocxEditor.create(el, { wasmUrl })` o con el atributo `wasm-url`.

## Limitaciones

Es un prototipo de laboratorio. Todavía no se dibujan el texto de encabezados y pies de página, las celdas combinadas, los cuadros de texto, los comentarios, el control de cambios ni las columnas múltiples. Todo eso se conserva intacto al guardar el documento.

Imágenes flotantes: el texto las rodea por un solo lado (el más ancho con `bothSides`) y usando su rectángulo, no su contorno. Solo desplazan el texto que viene después de su párrafo de anclaje en la misma página, y las tablas no las rodean. Las imágenes de encabezados y pies de página se muestran, pero todavía no se pueden editar.

## Licencia

MIT
