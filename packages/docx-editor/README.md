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
| `insertImage(image, { at, width, height, align, alt })` | Insertar una imagen PNG, JPEG o GIF (`Uint8Array`, `ArrayBuffer`, `Blob`/`File` o URL `data:`) en su propio párrafo, como un solo paso de deshacer. Sin tamaño conserva sus píxeles; nunca supera el ancho del texto. Devuelve una promesa |
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
| `message` | `{ message, error }`: avisos para mostrar al usuario |

Incluye tipos de TypeScript.

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
- **Editor:** `--docx-page-gap`, `--docx-page-shadow`, `--docx-page-border`, `--docx-label-color`, `--docx-caret-color`, `--docx-selection`.
- **Barra de herramientas:** `--docx-toolbar-bg`, `--docx-toolbar-border`, `--docx-toolbar-color`, `--docx-toolbar-hover`, `--docx-toolbar-active-bg`, `--docx-toolbar-active-color`.

El contenedor del editor necesita una altura (por ejemplo `height: 80vh`, o ser un elemento flex/grid con altura definida); el editor ocupa el 100 % y dentro de esa altura hace su propio scroll.

## Bundlers

El paquete carga `docx_engine_bg.wasm` con `new URL(..., import.meta.url)`. Vite, webpack 5, Rollup y esbuild copian ese archivo junto al JavaScript sin configuración extra. Si lo sirves desde otro lugar, pásalo con `DocxEditor.create(el, { wasmUrl })` o con el atributo `wasm-url`.

## Limitaciones

Es un prototipo de laboratorio. Todavía no se dibujan el texto de encabezados y pies de página, las celdas combinadas, los cuadros de texto, los comentarios, el control de cambios ni las columnas múltiples. Todo eso se conserva intacto al guardar el documento.

## Licencia

MIT
