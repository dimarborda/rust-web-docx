[← Índice](../README.md) · **Español** · [English](../en/engine.md)

# El motor sin interfaz

`@dimarborda/xlsx-editor/engine` es el mismo motor sin cuadrícula ni DOM: lee, escribe, recalcula y da formato a libros en Node.js (scripts, servidores, herramientas de agentes) o en el navegador.

```bash
npm install @dimarborda/xlsx-editor
```

```js
import { readFile, writeFile } from 'node:fs/promises';
import { openWorkbook } from '@dimarborda/xlsx-editor/engine';

const wb = await openWorkbook(await readFile('plantilla.xlsx'));

// Variables {{CAMPO}} en cualquier texto del libro
wb.fillTemplate({ CLIENTE: 'Acme S.A.', FECHA: '15/01/2024' });

// Filas de datos a partir de una celda
const hoja = wb.sheet('Ventas');
hoja.setValues('A2', [
  ['Teclado', 3, 29.9, new Date(Date.UTC(2024, 0, 15))],
  ['Ratón', 5, 12.5, new Date(Date.UTC(2024, 0, 16))],
]);

console.log(hoja.getValue('B10'));   // resultado de una fórmula, ya recalculado
await writeFile('salida.xlsx', wb.toBytes());
```

En Node el motor se carga solo. En el navegador, los bundlers (Vite, webpack 5, Rollup) sirven el `.wasm` junto al JavaScript; si lo sirves desde otro sitio, llama antes a `initEngine(url)`. La entrada principal del paquete también exporta todo esto, así que en el navegador basta con `import { openWorkbook } from '@dimarborda/xlsx-editor'`; `editor.workbook` es un `Workbook` de esta misma API.

## API

| | |
|---|---|
| `openWorkbook(bytes)` / `createWorkbook()` | Abre un `.xlsx` o crea un libro con una hoja vacía. |
| `wb.sheetNames` | Nombres de las hojas en el orden de las pestañas. |
| `wb.sheet(nombreOÍndice)` / `wb.sheets` | Una hoja de cálculo (por nombre, sin distinguir mayúsculas, o por posición desde 0) o todas. Las hojas de gráfico no tienen celdas. |
| `hoja.getCell('B3')` | `{ ref, type, value, formula? }`. `type` es `empty`, `number`, `string`, `boolean`, `date` o `error` (`value` es entonces el código, p. ej. `#DIV/0!`). En las fórmulas, `value` es su resultado y `formula` su texto sin `=` (también en las celdas de una fórmula compartida). |
| `hoja.getValue('B3')` / `hoja.getValues('A1:C3')` | Solo los valores; las celdas vacías dan `null`. |
| `hoja.setValue('B3', valor)` | Escribe texto, número, booleano, `Date` o una fórmula (`{ formula: '=SUM(A1:A3)' }`); `null` borra el valor y deja el formato. Las fórmulas que dependen de la celda se recalculan. |
| `hoja.setValues('A2', filas)` | Escribe una matriz desde la celda indicada, creando filas y celdas en orden. `undefined` deja la celda como está. |
| `wb.fillTemplate(valores)` | Sustituye `{{CAMPO}}` (se admiten espacios y una pista: `{{ CAMPO }}`, `{{CAMPO: pista}}`) en las cadenas compartidas y en las cadenas en línea; devuelve cuántas sustituyó. Las variables sin valor se quedan como están. |
| `wb.addSheet(nombre, después?)` | Agrega una hoja vacía detrás de la pestaña `después` (nombre o posición; al final sin ella) y la devuelve. El nombre sigue las reglas de Excel: de 1 a 31 caracteres, sin `[ ] : * ? / \`, sin apóstrofo al principio ni al final, distinto de `History` y de las demás hojas (sin distinguir mayúsculas). |
| `wb.renameSheet(hoja, nombre)` | Renombra una pestaña (nombre o posición) y las referencias a ella en fórmulas, nombres definidos, validaciones de datos, formato condicional y gráficos; devuelve cuántas cambió. Las referencias dentro de textos (`INDIRECT("Hoja1!A1")`) y a otros libros se quedan como están, igual que en Excel. |
| `wb.markers({ kind })` | Marcadores del léxico estándar (el mismo de `@dimarborda/docx-editor`) en las celdas de texto de todas las hojas: `field` para `{{CLAVE}}` / `{clave}`, `pending` para `[…]` (contenido por completar). Cada uno con `sheet`, `sheetName`, `ref` (`B3`), `row`, `col`, `raw`, `key`, `hint`, `offset` y `length`. |
| `wb.checkpoint(selección?)` | Abre un paso de deshacer: los cambios hasta el siguiente `checkpoint` se deshacen juntos. `selección` es un texto libre que `undo()` devuelve. `wb.setSelectionAfter(selección)` guarda la de después, que devuelve `redo()`. |
| `wb.undo()` / `wb.redo()` | Deshace o rehace un paso: `{ selection, parts }`, o `null` si no hay nada. `wb.canUndo` / `wb.canRedo` dicen si se puede. |
| `wb.recalculate()` | Recalcula todas las fórmulas y devuelve cuántas celdas cambiaron de valor. |
| `wb.areas({ sheet })` / `hoja.areas()` | Las áreas del libro (ver [Áreas](agents.md#áreas)). |
| `wb.area(destino, { sheet, maxRows, locale })` / `hoja.area(rango)` | Contexto de un área para un agente. |
| `hoja.regionAt('B5')` | La región alrededor de una celda, como la selecciona Ctrl+A en Excel (o su tabla). |
| `hoja.setFormat('A1:C3', formato)` / `hoja.clearFormat(rango)` | Da formato (ver [Formato](#formato)) o lo quita. |
| `hoja.setColumnWidth('B:D', 14)` | Ancho de columnas en caracteres. |
| `hoja.mergeCells('A1:H1')` / `hoja.unmergeCells(rango)` | Combina o separa celdas; al combinar, solo la primera puede tener valor. |
| `hoja.moveRows(primera, cantidad, antesDe)` / `hoja.moveColumns('B', cantidad, 'E')` | Mueve filas o columnas enteras delante de otra, numerada como está ahora (ver [Mover filas y columnas](operations.md#mover-filas-y-columnas)), y devuelve dónde queda la primera. `editor.moveRows` y `editor.moveColumns` hacen lo mismo en un paso de deshacer y seleccionan el resultado. |
| `hoja.insertRows(antesDe, cantidad?)` / `hoja.insertColumns('C', cantidad?)` | Inserta filas o columnas vacías, con el formato de la anterior (ver [Insertar y eliminar](operations.md#insertar-y-eliminar)). También `editor.insertRows` y `editor.insertColumns`, en un paso de deshacer. |
| `hoja.deleteRows(primera, cantidad?)` / `hoja.deleteColumns('C', cantidad?)` | Elimina filas o columnas. También `editor.deleteRows` y `editor.deleteColumns`. |
| `hoja.setRowHeight('3:5', 30)` | Alto de filas en puntos; `null` vuelve al alto automático. |
| `hoja.setRowsHidden('3:5', true)` / `hoja.setColumnsHidden('B:D', true)` | Oculta o muestra líneas. También `editor.hideRows`, `showRows`, `hideColumns` y `showColumns`. |
| `hoja.fill('A1:A2', 'A1:A20', { series })` | Rellena un rango desde las celdas de uno de sus extremos (ver [Relleno](operations.md#relleno)). `editor.fill('A1:A20')` lo hace desde la selección, en un paso de deshacer. |
| `hoja.getCells('A1:C3')` | Las celdas de un rango con su tipo, valor y fórmula. |
| `wb.toBytes()` | El libro como `.xlsx`. |
| `wb.close()` | Libera la memoria del motor. |

## Fórmulas

Las fórmulas se escriben con `{ formula: '=SUM(B2:B10)*IVA' }` y se calculan en el momento. Cada vez que se escribe una celda (o se rellena una plantilla) se recalculan, en orden de dependencias, todas las fórmulas del libro, y solo se reescribe el `<v>` de las celdas cuyo resultado cambió. `wb.recalculate()` hace lo mismo sin editar nada, por ejemplo para refrescar un libro recién abierto.

Se usa la sintaxis en inglés que guarda Excel (`SUM`, `IF`, `,` entre argumentos; también se acepta `;`). Se admiten referencias relativas y absolutas, rangos, columnas completas (`A:A`), otras hojas (`'Mis datos'!B2`), nombres definidos, fórmulas compartidas y matrices constantes (`{1,2;3,4}`), con las conversiones de Excel entre texto, número y booleano y sus códigos de error. Funciones:

- **Matemáticas:** `SUM`, `PRODUCT`, `AVERAGE`, `MIN`, `MAX`, `MEDIAN`, `COUNT`, `COUNTA`, `COUNTBLANK`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `TRUNC`, `ABS`, `SIGN`, `SQRT`, `POWER`, `MOD`, `PI`, `CEILING`, `FLOOR`, `CEILING.MATH`, `FLOOR.MATH`, `SUMPRODUCT`.
- **Condicionales:** `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS` (criterios como `">5"`, `"<>x"`, `"n*"`).
- **Lógicas e información:** `IF`, `IFS`, `SWITCH`, `IFERROR`, `IFNA`, `AND`, `OR`, `XOR`, `NOT`, `TRUE`, `FALSE`, `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISNONTEXT`, `ISLOGICAL`, `ISERROR`, `ISERR`, `ISNA`, `NA`.
- **Búsqueda:** `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `MATCH`, `INDEX`, `CHOOSE`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`.
- **Texto:** `CONCATENATE`, `CONCAT`, `TEXTJOIN`, `LEFT`, `RIGHT`, `MID`, `LEN`, `UPPER`, `LOWER`, `PROPER`, `TRIM`, `SUBSTITUTE`, `REPT`, `EXACT`, `FIND`, `SEARCH`, `VALUE`, `TEXT`.
- **Fechas:** `TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, `DAY`, `WEEKDAY`, `EDATE`, `EOMONTH`, `DAYS`.

Una fórmula que use otra cosa (una función que no está en la lista, `OFFSET`, `INDIRECT`, referencias a otros libros o a tablas, fórmulas matriciales, resultados que se desbordan a varias celdas) o que forme una referencia circular conserva el valor que guardó Excel, y las que dependen de ella lo usan. Al abrirse el archivo, Excel la calcula (ver [Sin pérdidas](#sin-pérdidas)).

## Formato

`hoja.setFormat(rango, formato)` cambia solo lo que se indica y conserva el resto del aspecto de cada celda:

| Propiedad | Valor |
|---|---|
| `bold`, `italic`, `underline`, `strike` | `true` / `false` |
| `color` | Color de la letra, `#RRGGBB` |
| `fontSize`, `fontName` | Puntos; nombre de fuente |
| `fill` | Fondo `#RRGGBB`; `null` lo quita |
| `numberFormat` | Código de Excel: `#,##0.00`, `0%`, `dd/mm/yyyy`, `#,##0.00 "€"`… |
| `align`, `valign`, `wrap` | `left`, `center`, `right`…; `top`, `center`, `bottom`; ajuste de texto |
| `border` | Un estilo (`'thin'`, `'medium'`, `'thick'`, `'dashed'`, `'dotted'`, `'double'`) en todos los lados, `null` para quitarlos, o `{ style, color, edges }` con `edges`: `all`, `outline`, `inside`, `top`, `bottom`, `left` o `right` |

Los formatos nuevos se añaden a `xl/styles.xml` y se reutilizan los iguales que ya existan; nada de lo que ya estaba se reescribe.

## Fechas

Excel guarda las fechas como números de serie (días desde 1900, o desde 1904 en libros con ese sistema) que se ven como fecha por el formato de la celda. Al leer, una celda numérica con formato de fecha se devuelve como `Date`; al escribir un `Date`, la celda recibe un formato de fecha basado en el que tenía (misma fuente, relleno y bordes), salvo que ya mostrara fechas.

Las fechas de Excel no tienen zona horaria: la hora que muestra Excel se representa como hora **UTC** del `Date`. Escribe con `new Date(Date.UTC(2024, 0, 15))` y lee con `getUTCFullYear()`, `getUTCHours()` o `toISOString()`.

## Sin pérdidas

Solo se reescriben las partes del archivo que cambian (la hoja editada, las cadenas compartidas y, si hace falta, los estilos), y dentro de ellas solo las celdas escritas: gráficos, imágenes, formatos condicionales, validaciones, nombres definidos y todo lo demás se guarda byte a byte tal como estaba.

Si se escribe sobre la celda de origen de una fórmula compartida (la que Excel guarda para todo un grupo de celdas rellenadas), las demás celdas del grupo pasan a tener cada una su propia fórmula, con el mismo resultado.

Cuando un libro cambia, se elimina `xl/calcChain.xml` (Excel lo reconstruye) y se marca `fullCalcOnLoad="1"` para que Excel recalcule las fórmulas al abrirlo: así quedan bien también las que este motor no calcula.

Limitaciones actuales:

- Las fórmulas fuera de lo admitido (ver [Fórmulas](#fórmulas)) no se recalculan: `getValue` devuelve su último resultado guardado.
- Las fórmulas se escriben en la sintaxis en inglés (`SUM`, no `SUMA`); los nombres de función en español no se traducen.
- Una variable `{{CAMPO}}` partida entre fragmentos de texto con formatos distintos no se reconoce.

## Content Security Policy

Como el motor es WebAssembly, el `script-src` de producción necesita `'wasm-unsafe-eval'`.

## Desarrollo

Desde la raíz del repositorio:

```bash
npm run build:wasm:xlsx
npm test -w packages/xlsx-editor
npm run dev        # http://localhost:5173/xlsx.html
```
