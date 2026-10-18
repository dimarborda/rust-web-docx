# @dimarborda/xlsx

Librería para libros de Excel (`.xlsx`) en Node.js y el navegador, con un motor escrito en Rust y compilado a WebAssembly. Comparte con [`@dimarborda/docx-editor`](../docx-editor) la capa de paquete OOXML (`crates/ooxml-core`): lectura del ZIP con límites de seguridad y guardado sin pérdidas, que deja intacto todo lo que no se edita.

> **Estado: en desarrollo temprano.** Por ahora abre, crea y guarda libros y lista sus hojas. La lectura y escritura de celdas y el relleno de plantillas llegarán en las próximas versiones.

```bash
npm install @dimarborda/xlsx
```

```js
import { readFile, writeFile } from 'node:fs/promises';
import { openWorkbook } from '@dimarborda/xlsx';

const wb = await openWorkbook(await readFile('plantilla.xlsx'));
console.log(wb.sheetNames);
await writeFile('salida.xlsx', wb.toBytes());
```

En Node el motor se carga solo. En el navegador, los bundlers (Vite, webpack 5, Rollup) sirven el `.wasm` junto al JavaScript; si lo sirves desde otro sitio, llama antes a `initEngine(url)`.

## Desarrollo

Desde la raíz del repositorio:

```bash
npm run build:wasm:xlsx
npm test -w packages/xlsx
```
