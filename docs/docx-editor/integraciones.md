# Integraciones de @dimarborda/docx-editor

[← Índice](README.md) · [API](api.md) · [Guías](guias.md) · [Integraciones](integraciones.md) · [Seguridad y limitaciones](seguridad.md)

React, Next.js, Tauri, fuentes, estilos, bundlers y Content Security Policy.

## React

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

## Next.js (App Router, Turbopack)

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

## Tauri

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
