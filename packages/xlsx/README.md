# @dimarborda/xlsx

Librería para libros de Excel (`.xlsx`) en Node.js y el navegador, con un motor escrito en Rust y compilado a WebAssembly. Comparte con [`@dimarborda/docx-editor`](../docx-editor) la capa de paquete OOXML (`crates/ooxml-core`): lectura del ZIP con límites de seguridad y guardado sin pérdidas, que deja intacto todo lo que no se edita.

> **Estado: en desarrollo temprano.** Lee y escribe valores de celda y rellena plantillas con variables `{{CAMPO}}`. Todavía no calcula fórmulas: al guardar un libro modificado se pide a Excel que las recalcule al abrirlo.

```bash
npm install @dimarborda/xlsx
```

```js
import { readFile, writeFile } from 'node:fs/promises';
import { openWorkbook } from '@dimarborda/xlsx';

const wb = await openWorkbook(await readFile('plantilla.xlsx'));

// Variables {{CAMPO}} en cualquier texto del libro
wb.fillTemplate({ CLIENTE: 'Acme S.A.', FECHA: '15/01/2024' });

// Filas de datos a partir de una celda
const hoja = wb.sheet('Ventas');
hoja.setValues('A2', [
  ['Teclado', 3, 29.9, new Date(Date.UTC(2024, 0, 15))],
  ['Ratón', 5, 12.5, new Date(Date.UTC(2024, 0, 16))],
]);

console.log(hoja.getValue('B10'));   // valor en caché de una fórmula, p. ej. 8
await writeFile('salida.xlsx', wb.toBytes());
```

En Node el motor se carga solo. En el navegador, los bundlers (Vite, webpack 5, Rollup) sirven el `.wasm` junto al JavaScript; si lo sirves desde otro sitio, llama antes a `initEngine(url)`.

## API

| | |
|---|---|
| `openWorkbook(bytes)` / `createWorkbook()` | Abre un `.xlsx` o crea un libro con una hoja vacía. |
| `wb.sheetNames` | Nombres de las hojas en el orden de las pestañas. |
| `wb.sheet(nombreOÍndice)` / `wb.sheets` | Una hoja de cálculo (por nombre, sin distinguir mayúsculas, o por posición desde 0) o todas. Las hojas de gráfico no tienen celdas. |
| `hoja.getCell('B3')` | `{ ref, type, value, formula? }`. `type` es `empty`, `number`, `string`, `boolean`, `date` o `error` (`value` es entonces el código, p. ej. `#DIV/0!`). En las fórmulas, `value` es el resultado guardado en el archivo. |
| `hoja.getValue('B3')` / `hoja.getValues('A1:C3')` | Solo los valores; las celdas vacías dan `null`. |
| `hoja.setValue('B3', valor)` | Escribe texto, número, booleano o `Date`; `null` borra el valor y deja el formato. |
| `hoja.setValues('A2', filas)` | Escribe una matriz desde la celda indicada, creando filas y celdas en orden. `undefined` deja la celda como está. |
| `wb.fillTemplate(valores)` | Sustituye `{{CAMPO}}` (se admiten espacios: `{{ CAMPO }}`) en las cadenas compartidas y en las cadenas en línea; devuelve cuántas sustituyó. Las variables sin valor se quedan como están. |
| `wb.toBytes()` | El libro como `.xlsx`. |
| `wb.close()` | Libera la memoria del motor. |

### Fechas

Excel guarda las fechas como números de serie (días desde 1900, o desde 1904 en libros con ese sistema) que se ven como fecha por el formato de la celda. Al leer, una celda numérica con formato de fecha se devuelve como `Date`; al escribir un `Date`, la celda recibe un formato de fecha basado en el que tenía (misma fuente, relleno y bordes), salvo que ya mostrara fechas.

Las fechas de Excel no tienen zona horaria: la hora que muestra Excel se representa como hora **UTC** del `Date`. Escribe con `new Date(Date.UTC(2024, 0, 15))` y lee con `getUTCFullYear()`, `getUTCHours()` o `toISOString()`.

## Sin pérdidas

Solo se reescriben las partes del archivo que cambian (la hoja editada, las cadenas compartidas y, si hace falta, los estilos), y dentro de ellas solo las celdas escritas: gráficos, imágenes, formatos condicionales, validaciones, nombres definidos y todo lo demás se guarda byte a byte tal como estaba.

Cuando un libro cambia, se elimina `xl/calcChain.xml` (Excel lo reconstruye) y se marca `fullCalcOnLoad="1"` para que Excel recalcule las fórmulas al abrirlo, ya que sus resultados guardados pueden depender de las celdas escritas.

Limitaciones actuales:

- No se calculan fórmulas; `getValue` devuelve el último resultado guardado.
- Escribir en la celda de origen de una fórmula compartida (la que usan otras celdas) da error, para no romper esas celdas.
- Una variable `{{CAMPO}}` partida entre fragmentos de texto con formatos distintos no se reconoce.

## Desarrollo

Desde la raíz del repositorio:

```bash
npm run build:wasm:xlsx
npm test -w packages/xlsx
```
