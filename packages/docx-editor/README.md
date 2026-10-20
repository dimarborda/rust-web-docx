# @dimarborda/docx-editor

[![npm](https://img.shields.io/npm/v/@dimarborda/docx-editor?color=cb3837&label=npm)](https://www.npmjs.com/package/@dimarborda/docx-editor)

Editor de documentos Word (`.docx`) para el navegador. El motor está escrito en Rust y compilado a WebAssembly: lee el documento, maqueta las páginas como Word y aplica ediciones sin pérdidas. Las páginas se dibujan en canvas, con cursor y selección propios.

Funciona en cualquier frontend (sin framework, React, Vue, Svelte, Angular) y en apps de escritorio con webview, como Tauri. No necesita servidor: los documentos nunca salen del equipo.

[Demo](https://rust-web-docx.dimarborda.workers.dev/?demo) · [Documentación completa](https://github.com/dimarborda/rust-web-docx/tree/master/docs/docx-editor) · [Repositorio](https://github.com/dimarborda/rust-web-docx)

```bash
npm install @dimarborda/docx-editor
```

## Web component

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

## JavaScript

```js
import { DocxEditor } from '@dimarborda/docx-editor';
import '@dimarborda/docx-editor/style.css';

const editor = await DocxEditor.create(document.querySelector('#editor'), { locale: 'es' });

await editor.open(file);                       // File, Blob, ArrayBuffer o Uint8Array
editor.replaceVariables({ CLIENTE: 'Acme' });  // rellena {{CLIENTE}}, {CLIENTE}…
editor.insertParagraphs(['Entre las partes se acuerda lo siguiente.'], { at: 'end' });
await editor.downloadPdf('contrato.pdf');      // PDF vectorial con texto seleccionable
const bytes = editor.save();                   // Uint8Array del .docx editado
```

El contenedor necesita una altura (por ejemplo `height: 80vh`). Incluye tipos de TypeScript.

## Qué puedes hacer

- **Plantillas y marcadores**: `markers()` lista los campos `{{CLAVE}}` / `{clave}` y el contenido pendiente `[…]` con su posición; `fillMarker()` y `replaceVariables()` los rellenan conservando el formato. [Guía](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/guias.md#marcadores-campos-y-pendientes)
- **Agentes**: `markActivity()` muestra dónde está trabajando un agente y las ediciones hechas por API se resaltan un momento, sin mover el cursor del usuario. [Guía](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/guias.md#actividad-dónde-está-trabajando-el-agente)
- **Contenido**: insertar párrafos con estilos de título, tablas, imágenes (en línea o flotantes), y leer la estructura del documento.
- **Secciones lógicas y tablas existentes**: leer y reemplazar partes del documento por clave, y escribir celdas o filas copiando su formato. [Guías](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/guias.md)
- **Exportar**: `.docx` sin pérdidas y PDF vectorial.
- **Vista**: solo lectura, zoom, marca de agua y barra de herramientas opcional.

Todos los métodos, opciones y eventos están en la [referencia de la API](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/api.md). Para React, Next.js, Tauri, fuentes, estilos y bundlers, mira las [integraciones](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/integraciones.md).

## Content Security Policy

Si tu aplicación envía `Content-Security-Policy`, añade `'wasm-unsafe-eval'` al `script-src` de producción, o el motor no arranca. [Detalles](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/integraciones.md#content-security-policy)

## Seguridad

Los documentos no salen del navegador, su contenido no se ejecuta (se dibuja en canvas) y el lector rechaza entidades XML externas y bombas ZIP. [Seguridad y limitaciones](https://github.com/dimarborda/rust-web-docx/blob/master/docs/docx-editor/seguridad.md) · [Reportar una vulnerabilidad](https://github.com/dimarborda/rust-web-docx/blob/master/SECURITY.md)

## Licencia

MIT
