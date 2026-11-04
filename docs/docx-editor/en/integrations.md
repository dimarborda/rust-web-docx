[← Index](../README.md) · **English** · [Español](../es/integrations.md)

# @dimarborda/docx-editor integrations

[API](api.md) · [Guides](guides.md) · [Integrations](integrations.md) · [Security and limitations](security.md)

React, Next.js, Tauri, fonts, styles, bundlers and Content Security Policy.

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
      // In StrictMode the effect mounts twice: discard the editor that arrives late
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

Tested with Next 16 and Turbopack, in `next dev`, `next build` and `next start`. The package is imported with a normal `import`: it does not touch the DOM when loaded, so server-side rendering does not fail. The editor is created on the client, inside `useEffect`.

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

`app/page.jsx` (a server component) can render `<Editor />` directly.

- **The `.wasm` (≈2 MB, ≈800 KB compressed) is downloaded only on the client** and only when the editor is created, not when the package is imported. Turbopack copies it to `/_next/static/media/` and Next serves it as `application/wasm`.
- **pdf-lib** is downloaded only when `exportPdf()` is called.
- **If you serve the `.wasm` from somewhere else** (a CDN, or `/public` in a deployment with its own rules), copy `node_modules/@dimarborda/docx-editor/wasm/docx_engine_bg.wasm` and pass its path: `DocxEditor.create(el, { wasmUrl: '/docx_engine_bg.wasm' })`. If you update the package, copy it again.
- If your application uses a CSP header (for example from `proxy.ts`), see the [Content Security Policy](#content-security-policy) section.
- Versions before 0.7.0 accessed `document` when imported and had to be loaded with `dynamic(() => import(...), { ssr: false })`; this is no longer necessary.

## Tauri

Inside the app window it is used just like on the web. The only thing that changes is opening and saving with the `dialog` and `fs` plugins:

```js
import { open, save } from '@tauri-apps/plugin-dialog';
import { readFile, writeFile } from '@tauri-apps/plugin-fs';

const path = await open({ filters: [{ name: 'Word', extensions: ['docx'] }] });
await editor.open(await readFile(path), { fileName: path.split(/[\\/]/).pop() });

const target = await save({ defaultPath: editor.fileName });
if (target) await writeFile(target, editor.save());
```

## Fonts

If the document uses Calibri, Cambria, Arial, Times New Roman or Courier and those fonts are not installed, the editor draws substitutes with the same metrics: Carlito, Caladea, Arimo, Tinos and Cousine. To make them available without depending on the system, import `@dimarborda/docx-editor/fonts`; it includes only the Latin subsets.

For a document to look exactly like it does in Word, register the real font:

```js
import { registerFont, loadStoredFonts } from '@dimarborda/docx-editor';

await loadStoredFonts();        // on startup: the fonts previously stored in this browser
await registerFont(fileInput.files[0]);  // .ttf/.otf/.woff/.woff2; stored in IndexedDB
```

## Styles

Styles are customized with CSS variables, defined on the editor container or on any ancestor:
- **Editor:** `--docx-page-gap`, `--docx-page-shadow`, `--docx-page-border`, `--docx-label-color`, `--docx-caret-color`, `--docx-selection`, `--docx-image-frame` (frame of the selected image).
- **Toolbar:** `--docx-toolbar-bg`, `--docx-toolbar-border`, `--docx-toolbar-color`, `--docx-toolbar-hover`, `--docx-toolbar-active-bg`, `--docx-toolbar-active-color`.

The editor container needs a height (for example `height: 80vh`, or being a flex/grid element with a defined height); the editor fills 100 % and scrolls on its own within that height.

## Bundlers

The package loads `docx_engine_bg.wasm` with `new URL(..., import.meta.url)`. Vite, webpack 5, Turbopack (tested with Next.js 16), Rollup and esbuild copy that file next to the JavaScript with no extra configuration. Importing the package does not touch the DOM, so it can also be imported in code that is rendered on the server (SSR); the editor is created in the browser. If you serve it from somewhere else, pass it with `DocxEditor.create(el, { wasmUrl })` or with the `wasm-url` attribute.

## Content Security Policy

If your application sends a `Content-Security-Policy` header, the **production** `script-src` must include `'wasm-unsafe-eval'`. Without it, the browser blocks compilation of the engine and the editor does not start: `DocxEditor.create()` / `initEngine()` fail with the error *"docx-editor: the page's Content-Security-Policy blocks WebAssembly. Add 'wasm-unsafe-eval' to script-src"* (the browser's original error is kept in `error.cause`).

The failure usually appears only on deployment. In development many setups (Next.js, Vite) include `'unsafe-eval'` for hot reloading, and that also allows WebAssembly. When you remove it in production, which is the right thing to do, you need to add `'wasm-unsafe-eval'`, which authorizes only WebAssembly and does not enable `eval()`.

| Directive | Value the editor needs | What for |
| :--- | :--- | :--- |
| `script-src` | `'self' 'wasm-unsafe-eval'` | Load the package's JavaScript and compile the WebAssembly engine |
| `connect-src` | `'self'` (or the origin you serve `docx_engine_bg.wasm` and the fonts from) | Download the `.wasm`, and read the fonts when exporting to PDF |
| `img-src` | `'self' data:` | The document's images are drawn from `data:` URLs (`blob:` is not needed) |
| `font-src` | `'self'` (or the fonts' origin) | The substitute fonts from `@dimarborda/docx-editor/fonts`. Those the user uploads with `registerFont()` are loaded from memory and do not depend on this directive |
| `style-src` | `'self'` | The package's CSS. If your bundler injects the CSS with `<style>` tags (for example Vite in development), add `'unsafe-inline'` or a *nonce* |

`download()` and `downloadPdf()` download from a `blob:` URL, which does not depend on `img-src` and needs nothing extra with common policies. The editor does not use `eval()`, workers, iframes or third-party resources.

Example for Next.js (`proxy.ts` or `middleware.ts`), where only development adds `'unsafe-eval'`:

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

The server must also deliver `docx_engine_bg.wasm` with the `application/wasm` type. With any other type the engine still starts, but more slowly, and the console shows an `instantiateStreaming` warning.

`'wasm-unsafe-eval'` is supported by Chrome and Edge 97+, Firefox 102+ and Safari 16+. In older browsers only `'unsafe-eval'` works. The published demo uses exactly this policy (in [`apps/demo/public/_headers`](../../../apps/demo/public/_headers)).
