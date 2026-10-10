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

Atributos de `<docx-editor>`: `src` (URL del .docx), `sample` (abre el contrato de ejemplo), `blank` (documento nuevo vacío, con `page-size="a4|letter|legal"` y `autofocus` opcionales), `locale` (`es` o `en`), `zoom`, `gridlines="false"`, `page-labels="false"`, `readonly` (solo lectura; sin el atributo es editable) y `wasm-url`.

`<docx-toolbar>` es opcional. Acepta `for="id-del-editor"` e `items="undo redo | bold italic underline color | left center right both | imageWrap | zoom"`. `imageWrap` es un selector del ajuste de texto que se activa al seleccionar una imagen; con una imagen seleccionada, los botones de alineación alinean la imagen.

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
| `exportPdf({ watermark })` / `exportPdfBlob()` / `downloadPdf(name)` | Exportar a PDF vectorial (ver [PDF](#pdf)). Devuelven promesas |
| `text()` / `variables()` / `stats()` / `fonts()` | Leer el contenido |
| `replaceVariables(values)` / `findReplace(search, replacement, options)` | Rellenar plantillas y reemplazar texto |
| `insertParagraphs(paragraphs, { at })` / `insertText(text, { at })` | Insertar contenido en el cursor (`at: 'cursor'`, predeterminado), al inicio o al final, como un solo paso de deshacer. Cada párrafo es texto o `{ text, bold, italic, underline, fontSize, align }` |
| `insertTable({ rows, header, widths, align }, { at })` | Insertar una tabla con contenido en el cursor, al inicio o al final, como un solo paso de deshacer. Cada celda es un texto (`\n` = salto de línea en la celda) o `{ text, image, width, height, alt }` con una imagen PNG, JPEG o GIF (`Uint8Array`, `ArrayBuffer` o URL `data:`) que ocupa el ancho de la columna salvo que indiques `width`/`height`, con el texto debajo; `header` (predeterminado `true`) pone la primera fila en negrita, sombreada y repetida en cada página; `widths` son anchos relativos. No se permiten tablas dentro de tablas |
| `insertImage(image, { at, width, height, align, alt, wrap, wrapSide, horizontal, vertical, distance })` | Insertar una imagen PNG, JPEG o GIF (`Uint8Array`, `ArrayBuffer`, `Blob`/`File` o URL `data:`) como un solo paso de deshacer. Sin tamaño conserva sus píxeles; nunca supera el ancho del texto. En línea (predeterminado) va en su propio párrafo; con otro `wrap` queda flotante, anclada al párrafo del cursor (ver [Imágenes](#imágenes)). Devuelve una promesa con `{ paragraph, index }` |
| `images()` / `selectedImage` / `selectImage(ref)` | Imágenes del documento y la seleccionada (ver [Imágenes](#imágenes)) |
| `updateImage(ref, changes)` / `resizeImage` / `moveImage` / `setImageWrap` / `alignImage` / `deleteImage` | Tamaño, posición y ajuste del texto de una imagen, cada cambio como un paso de deshacer |
| `undo()` / `redo()` / `canUndo` / `canRedo` | Historial |
| `toggleBold()` / `toggleItalic()` / `toggleUnderline()` / `setColor(hex)` / `setAlignment(align)` | Formato de la selección |
| `setBackgroundColor(hex)` / `setWatermark(text, { opacity })` | Apariencia. `insertTable(rows, cols, headers)` de versiones anteriores sigue funcionando (tabla vacía al final) |
| `select(anchor, focus)` | Coloca el cursor o una selección (`{ paragraph, offset }`) |
| `setZoom(z)` / `focus()` / `destroy()` | Vista y ciclo de vida. Sin cursor previo, `focus()` lo pone al inicio del documento |
| `setReadOnly(true \| false)` / `readOnly` | Modo de solo lectura (ver [Solo lectura](#solo-lectura)). Por defecto el editor es editable |

El editor crea su propio elemento (`editor.root`, clase `docx-editor`) dentro del contenedor y `destroy()` lo elimina sin tocar el contenedor (`editor.container`).

| Evento | `detail` |
| :--- | :--- |
| `load` | `{ fileName }` |
| `change` | `{}` |
| `selectionchange` | `{ selection, format: { bold, italic, underline, color, align } }` |
| `imageselect` | `{ image }`: la imagen seleccionada con el ratón o `selectImage()`, o cambiada mientras está seleccionada (`null` al deseleccionar) |
| `readonlychange` | `{ readOnly }`: `setReadOnly()` cambió el modo |
| `message` | `{ message, error }`: avisos para mostrar al usuario |

Incluye tipos de TypeScript.

### PDF

```js
const pdf = await editor.exportPdf();            // Uint8Array
await editor.downloadPdf('contrato.pdf');         // o descargarlo directamente
const blob = await editor.exportPdfBlob({ watermark: false });
```

Genera un PDF vectorial de las páginas tal como se ven en el editor: texto real (seleccionable y buscable) con las fuentes incrustadas y recortadas a los caracteres usados, imágenes, tablas, bordes, formas, cuadros de texto, encabezados y pies con su número de página. La marca de agua de `setWatermark()` se incluye salvo `{ watermark: false }`. Las fuentes son las que sube el usuario con `registerFont()` o las sustitutas métricas del paquete (`@dimarborda/docx-editor/fonts`); sin ellas se usan las fuentes estándar del PDF, y el texto se ajusta al ancho medido en pantalla para que nada se desplace. [pdf-lib](https://pdf-lib.js.org/) se descarga solo la primera vez que se exporta.

### Solo lectura

El editor es editable por defecto. Para mostrar un documento sin que el usuario pueda cambiarlo:

```js
const editor = await DocxEditor.create(el, { readOnly: true }); // desde el inicio
editor.setReadOnly(true);   // o en cualquier momento
editor.setReadOnly(false);  // vuelve a ser editable
```

```html
<docx-editor src="/contrato.docx" readonly></docx-editor>
```

En solo lectura el usuario puede desplazarse, hacer zoom, seleccionar y copiar texto, pero se ignoran la escritura, el borrado, pegar, cortar, los atajos de formato, deshacer y rehacer, y la selección y edición de imágenes. La barra de herramientas desactiva sus botones de edición (el zoom sigue activo). Las llamadas desde código (`replaceVariables`, `insertText`, `updateImage`…) siguen funcionando, para que tu aplicación pueda rellenar un documento que sus usuarios solo ven.

### Imágenes

Con el ratón: un clic selecciona la imagen y muestra ocho asas. Las esquinas cambian el tamaño manteniendo la proporción (`Shift` la libera) y los lados cambian una sola medida. Una imagen flotante se mueve arrastrándola o con las flechas (`Shift` + flecha = 10 px). `Supr` la elimina y `Esc` vuelve al texto. Una imagen detrás del texto solo se selecciona donde no hay texto encima, así que hacer clic sobre el texto de un membrete sigue colocando el cursor.

Los cuadros de texto y las formas (rectángulos, elipses…) funcionan igual: aparecen en `images()` con `kind: 'textbox'` o `kind: 'shape'`, su texto en `text` y su relleno y borde en `shape`, y se seleccionan, mueven, redimensionan y cambian de ajuste con los mismos métodos.

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

### Next.js (App Router, Turbopack)

Probado con Next 16 y Turbopack, tanto en `next dev` como en `next build` y `next start`. El paquete se importa con un `import` normal: no toca el DOM al cargarse, así que el renderizado en servidor no falla. El editor se crea en el cliente, dentro de `useEffect`.

```jsx
// app/editor.jsx
'use client';
import { useEffect, useRef } from 'react';
import { DocxEditor } from '@dimarborda/docx-editor';
import '@dimarborda/docx-editor/style.css';
import '@dimarborda/docx-editor/fonts';

export default function Editor({ file }) {
  const ref = useRef(null);
  useEffect(() => {
    let editor;
    let disposed = false;
    DocxEditor.create(ref.current, { locale: 'es' }).then(async e => {
      if (disposed) return e.destroy();
      editor = e;
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

`app/page.jsx` (un componente de servidor) puede renderizar `<Editor />` directamente.

- **El `.wasm` (≈2 MB, ≈800 KB comprimido) se descarga solo en el cliente** y solo al crear el editor, no al importar el paquete. Turbopack lo copia a `/_next/static/media/` y Next lo sirve como `application/wasm`.
- **pdf-lib** se descarga únicamente al llamar a `exportPdf()`.
- **Si sirves el `.wasm` desde otro lugar** (un CDN, o `/public` en un despliegue con reglas propias), copia `node_modules/@dimarborda/docx-editor/wasm/docx_engine_bg.wasm` y pasa su ruta: `DocxEditor.create(el, { wasmUrl: '/docx_engine_bg.wasm' })`. Si actualizas el paquete, vuelve a copiarlo.
- Si tu aplicación usa una cabecera CSP (por ejemplo desde `proxy.ts`), revisa la sección [Content Security Policy](#content-security-policy).
- Versiones anteriores a la 0.7.0 accedían a `document` al importarse y había que cargarlas con `dynamic(() => import(...), { ssr: false })`; ya no hace falta.

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

El paquete carga `docx_engine_bg.wasm` con `new URL(..., import.meta.url)`. Vite, webpack 5, Turbopack (probado con Next.js 16), Rollup y esbuild copian ese archivo junto al JavaScript sin configuración extra. Importar el paquete no toca el DOM, así que también se puede importar en código que se renderiza en el servidor (SSR); el editor se crea en el navegador. Si lo sirves desde otro lugar, pásalo con `DocxEditor.create(el, { wasmUrl })` o con el atributo `wasm-url`.

## Content Security Policy

Si tu aplicación envía una cabecera `Content-Security-Policy`, el `script-src` de **producción** debe incluir `'wasm-unsafe-eval'`. Sin ella, el navegador bloquea la compilación del motor y el editor no arranca: `DocxEditor.create()` / `initEngine()` fallan con el error *"docx-editor: the page's Content-Security-Policy blocks WebAssembly. Add 'wasm-unsafe-eval' to script-src"* (el error original del navegador queda en `error.cause`).

El fallo suele aparecer solo al desplegar. En desarrollo muchas configuraciones (Next.js, Vite) incluyen `'unsafe-eval'` para la recarga en caliente, y eso también permite el WebAssembly. Al quitarlo en producción, que es lo correcto, hay que añadir `'wasm-unsafe-eval'`, que solo autoriza WebAssembly y no habilita `eval()`.

| Directiva | Valor que necesita el editor | Para qué |
| :--- | :--- | :--- |
| `script-src` | `'self' 'wasm-unsafe-eval'` | Cargar el JavaScript del paquete y compilar el motor WebAssembly |
| `connect-src` | `'self'` (o el origen desde donde sirvas `docx_engine_bg.wasm` y las fuentes) | Descargar el `.wasm`, y leer las fuentes al exportar a PDF |
| `img-src` | `'self' data:` | Las imágenes del documento se dibujan desde URLs `data:` (no hace falta `blob:`) |
| `font-src` | `'self'` (o el origen de las fuentes) | Las fuentes sustitutas de `@dimarborda/docx-editor/fonts`. Las que sube el usuario con `registerFont()` se cargan desde memoria y no dependen de esta directiva |
| `style-src` | `'self'` | El CSS del paquete. Si tu empaquetador inyecta el CSS con etiquetas `<style>` (por ejemplo Vite en desarrollo), añade `'unsafe-inline'` o un *nonce* |

`download()` y `downloadPdf()` descargan desde una URL `blob:`, que no depende de `img-src` y no necesita nada extra con políticas habituales. El editor no usa `eval()`, workers, iframes ni recursos de terceros.

Ejemplo para Next.js (`proxy.ts` o `middleware.ts`), donde solo desarrollo añade `'unsafe-eval'`:

```ts
const isDev = process.env.NODE_ENV === 'development';
const csp = [
  "default-src 'self'",
  `script-src 'self' 'wasm-unsafe-eval'${isDev ? " 'unsafe-eval'" : ''}`,
  "connect-src 'self'",
  "img-src 'self' data:",
  "font-src 'self'",
  `style-src 'self'${isDev ? " 'unsafe-inline'" : ''}`,
].join('; ');
```

El servidor también debe entregar `docx_engine_bg.wasm` con el tipo `application/wasm`. Con otro tipo el motor igual arranca, pero más despacio, y la consola muestra un aviso de `instantiateStreaming`.

`'wasm-unsafe-eval'` lo admiten Chrome y Edge 97+, Firefox 102+ y Safari 16+. En navegadores más antiguos solo funciona `'unsafe-eval'`. La demo publicada usa exactamente esta política (en [`apps/demo/public/_headers`](../../apps/demo/public/_headers)).

## Seguridad y privacidad

- **Los documentos no salen del navegador.** Abrir, editar, guardar y exportar a PDF ocurren en la página. No hay servidor propio, telemetría ni peticiones a terceros. Las únicas descargas son las del propio sitio: el motor `.wasm`, las fuentes y el `.docx` del atributo `src`, si se usa. Las fuentes que registra el usuario con `registerFont` se guardan solo en el IndexedDB de ese navegador.
- **El contenido del documento no se ejecuta.** Las páginas se dibujan en un `<canvas>` y el texto nunca se inserta como HTML, así que un `.docx` malicioso no puede inyectar scripts (XSS). Las imágenes se leen solo del propio archivo; los enlaces a recursos externos no generan peticiones.
- **XML sin entidades externas.** El parser (quick-xml) no resuelve DTD ni entidades externas, lo que evita ataques XXE.
- **Límites contra bombas ZIP.** Se rechazan los archivos con más de 10 000 partes, con una parte de más de 256 MB o con más de 512 MB en total al descomprimirse. El tamaño se mide mientras se descomprime, sin fiarse del que declara el ZIP. Las imágenes insertadas por API tienen un máximo de 15 MB.
- **Dependencias auditadas.** Cada despliegue del repositorio ejecuta `cargo audit` y `npm audit` y se detiene si hay vulnerabilidades conocidas. En tiempo de ejecución el paquete solo depende de `pdf-lib`, `@pdf-lib/fontkit` (MIT) y de las fuentes Carlito, Caladea, Arimo, Tinos y Cousine de Fontsource (SIL OFL 1.1).
- **CSP estricta.** Basta con añadir `'wasm-unsafe-eval'` al `script-src`; no hace falta `'unsafe-eval'` en scripts (ver [Content Security Policy](#content-security-policy)).

Las vulnerabilidades se reportan en privado, como se explica en [SECURITY.md](https://github.com/dimarborda/rust-web-docx/blob/master/SECURITY.md).

## Limitaciones

Todavía no se dibujan las formas agrupadas, los comentarios, el control de cambios ni las columnas múltiples. Todo eso se conserva intacto al guardar el documento. Los encabezados y pies de página se dibujan (texto, tablas, imágenes, número de página y primera página distinta) y se editan con doble clic sobre ellos, como en Word: aparece una línea punteada con la etiqueta "Encabezado" o "Pie de página" y el texto se escribe, se borra y se formatea como el del cuerpo; Esc o un clic en el cuerpo vuelven a él. No se insertan tablas ni imágenes en ellos. Se usan el encabezado y el pie de la última sección (sin variante para páginas pares), y en una línea con número de página el cursor puede desfasarse si el número mostrado tiene otra cantidad de cifras que el guardado. En la selección, sus párrafos usan índices a partir de 33 554 432 (2²⁵).

Cuadros de texto: un clic dentro coloca el cursor en su texto (escribir, Enter, borrar, formato, deshacer) y un clic en su borde selecciona el cuadro para moverlo, redimensionarlo o cambiar su ajuste como una imagen. La selección y las flechas no salen del cuadro, y ⌘A selecciona todo su texto. No se pueden insertar tablas ni imágenes flotantes dentro de un cuadro. Las tablas dentro de un cuadro y el giro del cuadro no se dibujan, y el texto que no cabe no se recorta. Los cuadros VML de documentos antiguos se dibujan, pero no se pueden editar. En la selección (`selection`), los párrafos de los cuadros usan índices a partir de 16 777 216 (2²⁴), para no desplazar la numeración del cuerpo.

Imágenes flotantes: el texto las rodea por ambos lados cuando hay sitio (`bothSides`), por el lado indicado (`left`, `right`, `largest`) y siguiendo su contorno en el ajuste estrecho y transparente (`wrapPolygon`); una fila de tabla que chocaría con ellas baja hasta pasarlas. Los objetos que se pisan se dibujan en el orden de Word (`relativeHeight`). Solo desplazan el texto que viene después de su párrafo de anclaje en la misma página, y las imágenes de encabezados y pies de página se muestran, pero no se pueden seleccionar.

## Licencia

MIT
