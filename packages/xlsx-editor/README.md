# @dimarborda/xlsx-editor

Componente para ver y editar libros de Excel (`.xlsx`) en el navegador, dibujados en canvas como los muestra Excel. Por dentro usa un motor escrito en Rust y compilado a WebAssembly, que lee el archivo, resuelve estilos y formatos de número, recalcula fórmulas y guarda sin pérdidas. Ese motor también se usa sin interfaz, en Node.js, desde [`@dimarborda/xlsx-editor/engine`](#motor-sin-interfaz-nodejs). Comparte con [`@dimarborda/docx-editor`](../docx-editor) la capa de paquete OOXML (`crates/ooxml-core`). Funciona en cualquier frontend (HTML, React, Vue, Svelte, Angular, Tauri) y nada sale del navegador.

> **Estado: fase 2, edición de celdas.** Se escriben valores y fórmulas en la cuadrícula, se borran, se copian y pegan, y se deshacen. Las fórmulas habituales se recalculan al momento (ver [Limitaciones](#limitaciones)). Con el atributo `readonly` sigue siendo un visor.

```bash
npm install @dimarborda/xlsx-editor
```

```html
<xlsx-editor src="/informe.xlsx" locale="es" style="height: 80vh"></xlsx-editor>
<script type="module">
  import '@dimarborda/xlsx-editor';
  import '@dimarborda/xlsx-editor/fonts';      // opcional: sustitutos de Calibri, Arial…
  import '@dimarborda/xlsx-editor/style.css';
</script>
```

## Qué muestra

- **Celdas como en Excel:** fuentes, negrita, cursiva, subrayado, tachado, colores (incluidos los del tema y sus tonos), rellenos y bordes con todos sus estilos.
- **Formatos de número:** separadores según el idioma, moneda, porcentajes, notación científica, fracciones, secciones y condiciones, colores como `[Red]`, y fechas y horas con nombres de meses y días en español o inglés.
- **Alineación:** la de Excel por defecto (números a la derecha, texto a la izquierda), horizontal y vertical, sangría, ajuste de texto, texto que se desborda sobre celdas vacías y `####` en los números que no caben.
- **Cuadrícula:** anchos de columna y altos de fila del archivo, filas y columnas ocultas, celdas combinadas, paneles inmovilizados y líneas de cuadrícula (o su ausencia).
- **Navegación:** desplazamiento fluido en hojas grandes (solo se dibuja lo visible), zoom, pestañas de hojas, selección con ratón y teclado, y una barra con la celda activa y su contenido o fórmula.

Aún no se dibujan gráficos, imágenes, formato condicional ni comentarios, y el texto girado se dibuja horizontal. Las fórmulas se recalculan tras cada cambio con el motor del paquete; las que usan algo que no admite muestran el último valor que guardó Excel.

## Edición

| Acción | Cómo |
| :--- | :--- |
| Escribir en una celda | Empieza a escribir (sustituye el contenido), pulsa **F2** o haz doble clic (edita el contenido actual), o escribe en la barra de celda. |
| Confirmar | **Enter** (baja), **Mayús+Enter** (sube), **Tab** / **Mayús+Tab** (derecha / izquierda). Si empezaste a escribir sin F2, las flechas también confirman y mueven; **F2** alterna entre ese modo y mover el cursor dentro del texto. |
| Rellenar la selección | **Ctrl/Cmd+Enter** escribe lo mismo en todas las celdas seleccionadas. |
| Salto de línea en la celda | **Alt+Enter**. |
| Cancelar | **Escape**. |
| Borrar | **Supr** o **Retroceso** vacían el rango seleccionado; el formato de las celdas se queda. |
| Copiar, cortar y pegar | **Ctrl/Cmd+C**, **X** y **V**, como texto separado por tabulaciones (TSV), el mismo que intercambian Excel y Google Sheets. Al copiar se copian los textos que muestran las celdas; si se pega en el mismo editor lo que se acaba de copiar, se pegan los valores exactos. Un bloque que cabe un número entero de veces en la selección se repite. Cortar vacía las celdas al momento (no al pegar, como hace Excel). |
| Deshacer y rehacer | **Ctrl/Cmd+Z**, y **Ctrl/Cmd+Y** o **Cmd+Mayús+Z**. Cada edición, borrado o pegado es un paso; deshacer vuelve a la hoja y a la selección de entonces. |

### Qué se guarda al escribir

Lo escrito se interpreta como lo hace Excel en el idioma del editor (`locale`; `es` y sus variantes usan las convenciones españolas y el resto las inglesas):

| Escribes (`es`) | Escribes (`en`) | Se guarda |
| :--- | :--- | :--- |
| `1234,5` · `1.234,5` · `-0,5` · `(300)` · `1,5E+3` | `1234.5` · `1,234.5` · `-0.5` · `(300)` · `1.5E+3` | Número |
| `1.234,50 €` | `$1,234.50` | Número (el símbolo se descarta) |
| `15%` · `12,5 %` | `15%` · `12.5 %` | Número dividido por 100 (0,15) |
| `15/03/2024` · `15-3-24` · `15/3` · `2024-03-15` | `3/15/2024` · `3-15-24` · `3/15` · `2024-03-15` | Fecha (sin año, el actual; años de dos cifras: 00–29 → 2000, 30–99 → 1900) |
| `15/03/2024 14:30` | `3/15/2024 2:30 PM` | Fecha y hora |
| `14:30` · `14:30:15` | `2:30 PM` | Hora, como fracción del día |
| `VERDADERO` · `FALSO` | `TRUE` · `FALSE` | Booleano (sin distinguir mayúsculas) |
| `'0012` | `'0012` | Texto `0012`: el apóstrofo inicial fuerza texto |
| `=SUM(A1:A3)` | `=SUM(A1:A3)` | Fórmula (ver Limitaciones) |
| Cualquier otra cosa | | Texto |
| (vacío) | | Celda vacía |

Los valores se escriben con `sheet.setValues` del motor, que conserva el formato de la celda: un número escrito en una celda con formato de moneda se ve como moneda. A las fechas se les da un formato de fecha basado en el de la celda si no lo tenía. Los porcentajes y las horas se guardan como su valor (0,15 o 0,604…) y se ven con el formato que ya tenga la celda. La misma función está disponible como `parseInput(texto, locale)`.

### Limitaciones

- **Fórmulas:** lo que empieza por `=` se guarda como fórmula y se calcula, igual que las fórmulas que dependen de ella. Se escriben en la sintaxis en inglés (`=SUM(A1:A3)`, `,` o `;` entre argumentos) en cualquier idioma del editor. Las funciones admitidas están en el [motor](#fórmulas); con otras, la celda muestra el último valor guardado y Excel la calcula al abrir el archivo (`fullCalcOnLoad`).
- **Formato:** escribir no cambia el formato de número (salvo en fechas), y las columnas no se ensanchan solas: una fecha en una columna estrecha se ve como `####`, como en Excel.
- No se puede sobrescribir la celda de origen de una fórmula compartida (la que usan otras celdas); el editor emite `error` y deja el libro como estaba.

### Áreas y contexto para agentes

- `editor.selectArea(destino)`: selecciona un área y muestra su hoja. Acepta una tabla o un rango con nombre, una dirección (`Ventas!A1:H40`) o un rango de la hoja activa. Devuelve el área (ver [Áreas para agentes](#áreas-para-agentes)).
- `editor.selectionContext({ maxRows })`: contexto de lo seleccionado para un agente: el rango, o con una sola celda la región que la rodea (su tabla, o el bloque que seleccionaría Ctrl+A). Devuelve título, cabecera, tipos y totales por columna, las primeras filas y una tabla Markdown.

Para dar formato desde código se usa `editor.workbook.sheet(…).setFormat(…)` y después `editor.reload()`.

## Atributos

| Atributo | Para qué |
| :--- | :--- |
| `src` | URL del `.xlsx` |
| `blank` | Empieza con un libro vacío |
| `locale` | Idioma de los formatos (`es`, `en-US`…); por defecto, el de la página |
| `sheet` | Hoja inicial, por nombre o posición |
| `zoom` | 0.25–4 |
| `tabs="false"` | Oculta las pestañas |
| `formula-bar="false"` | Oculta la barra de celda |
| `readonly` | Modo visor: se selecciona y se copia, pero no se edita (se puede quitar y poner en cualquier momento) |
| `quiet` | No escribe los errores en la consola |
| `wasm-url` | Dónde se sirve `xlsx_engine_bg.wasm`, si el bundler no lo hace |

Eventos en el elemento: `ready` (con `detail.editor`), `load`, `sheetchange`, `selectionchange`, `change` y `error`.

`change` llega tras cada edición, borrado, pegado, corte, deshacer o rehacer:

```js
element.addEventListener('change', ({ detail }) => {
  // detail: { sheet: 'Ventas', index: 0, range: 'B3:C4', source: 'edit' | 'clear' | 'paste' | 'cut' | 'undo' | 'redo',
  //           cells: [{ ref: 'B3', value: 1200 }, …] }
});
```

## API

```js
const element = document.querySelector('xlsx-editor');
element.addEventListener('ready', async ({ detail: { editor } }) => {
  await editor.open(file);             // URL, File/Blob, ArrayBuffer o Uint8Array
  editor.setSheet('Resumen');
  editor.select('B3:D7');
  editor.setZoom(1.25);

  // Cambiar celdas con la API del motor y volver a dibujar
  editor.workbook.fillTemplate({ CLIENTE: 'Acme' });
  editor.workbook.sheet('Resumen').setValues('B2', [[1200, 0.25]]);
  editor.reload();

  editor.undo();                       // también editor.redo(), editor.canUndo, editor.canRedo
  editor.readOnly = true;              // lo mismo que el atributo readonly

  editor.download('informe.xlsx');
});
```

Los cambios hechos desde código con `editor.workbook` no forman un paso de deshacer propio: llama antes a `editor.workbook.checkpoint()` para que lo sean.

Sin el elemento: `const editor = await XlsxEditor.create(contenedor, { locale: 'es', readOnly: false })`.

## Motor sin interfaz (Node.js)

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

### API

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
| `wb.markers({ kind })` | Marcadores del léxico estándar (el mismo de `@dimarborda/docx-editor`) en las celdas de texto de todas las hojas: `field` para `{{CLAVE}}` / `{clave}`, `pending` para `[…]` (contenido por completar). Cada uno con `sheet`, `sheetName`, `ref` (`B3`), `row`, `col`, `raw`, `key`, `hint`, `offset` y `length`. |
| `wb.checkpoint(selección?)` | Abre un paso de deshacer: los cambios hasta el siguiente `checkpoint` se deshacen juntos. `selección` es un texto libre que `undo()` devuelve. `wb.setSelectionAfter(selección)` guarda la de después, que devuelve `redo()`. |
| `wb.undo()` / `wb.redo()` | Deshace o rehace un paso: `{ selection, parts }`, o `null` si no hay nada. `wb.canUndo` / `wb.canRedo` dicen si se puede. |
| `wb.recalculate()` | Recalcula todas las fórmulas y devuelve cuántas celdas cambiaron de valor. |
| `wb.areas({ sheet })` / `hoja.areas()` | Las áreas del libro (ver [Áreas](#áreas-para-agentes)). |
| `wb.area(destino, { sheet, maxRows, locale })` / `hoja.area(rango)` | Contexto de un área para un agente. |
| `hoja.regionAt('B5')` | La región alrededor de una celda, como la selecciona Ctrl+A en Excel (o su tabla). |
| `hoja.setFormat('A1:C3', formato)` / `hoja.clearFormat(rango)` | Da formato (ver [Formato](#formato)) o lo quita. |
| `hoja.setColumnWidth('B:D', 14)` | Ancho de columnas en caracteres. |
| `hoja.mergeCells('A1:H1')` / `hoja.unmergeCells(rango)` | Combina o separa celdas; al combinar, solo la primera puede tener valor. |
| `wb.toBytes()` | El libro como `.xlsx`. |
| `wb.close()` | Libera la memoria del motor. |

#### Fórmulas

Las fórmulas se escriben con `{ formula: '=SUM(B2:B10)*IVA' }` y se calculan en el momento. Cada vez que se escribe una celda (o se rellena una plantilla) se recalculan, en orden de dependencias, todas las fórmulas del libro, y solo se reescribe el `<v>` de las celdas cuyo resultado cambió. `wb.recalculate()` hace lo mismo sin editar nada, por ejemplo para refrescar un libro recién abierto.

Se usa la sintaxis en inglés que guarda Excel (`SUM`, `IF`, `,` entre argumentos; también se acepta `;`). Se admiten referencias relativas y absolutas, rangos, columnas completas (`A:A`), otras hojas (`'Mis datos'!B2`), nombres definidos, fórmulas compartidas y matrices constantes (`{1,2;3,4}`), con las conversiones de Excel entre texto, número y booleano y sus códigos de error. Funciones:

- **Matemáticas:** `SUM`, `PRODUCT`, `AVERAGE`, `MIN`, `MAX`, `MEDIAN`, `COUNT`, `COUNTA`, `COUNTBLANK`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `TRUNC`, `ABS`, `SIGN`, `SQRT`, `POWER`, `MOD`, `PI`, `CEILING`, `FLOOR`, `CEILING.MATH`, `FLOOR.MATH`, `SUMPRODUCT`.
- **Condicionales:** `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS`, `MINIFS` (criterios como `">5"`, `"<>x"`, `"n*"`).
- **Lógicas e información:** `IF`, `IFS`, `SWITCH`, `IFERROR`, `IFNA`, `AND`, `OR`, `XOR`, `NOT`, `TRUE`, `FALSE`, `ISBLANK`, `ISNUMBER`, `ISTEXT`, `ISNONTEXT`, `ISLOGICAL`, `ISERROR`, `ISERR`, `ISNA`, `NA`.
- **Búsqueda:** `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `MATCH`, `INDEX`, `CHOOSE`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS`.
- **Texto:** `CONCATENATE`, `CONCAT`, `TEXTJOIN`, `LEFT`, `RIGHT`, `MID`, `LEN`, `UPPER`, `LOWER`, `PROPER`, `TRIM`, `SUBSTITUTE`, `REPT`, `EXACT`, `FIND`, `SEARCH`, `VALUE`, `TEXT`.
- **Fechas:** `TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, `DAY`, `WEEKDAY`, `EDATE`, `EOMONTH`, `DAYS`.

Una fórmula que use otra cosa (una función que no está en la lista, `OFFSET`, `INDIRECT`, referencias a otros libros o a tablas, fórmulas matriciales, resultados que se desbordan a varias celdas) o que forme una referencia circular conserva el valor que guardó Excel, y las que dependen de ella lo usan. Al abrirse el archivo, Excel la calcula (ver abajo).

#### Áreas para agentes

Un agente no necesita leer toda la hoja: puede pedir las áreas y luego el contexto de la que le interesa. `wb.areas()` usa lo que el propio archivo marca, en este orden:

1. **Tablas** de Excel (Dar formato como tabla, `xl/tables/tableN.xml`): nombre, rango, nombres de columna y fila de totales.
2. **Nombres definidos** que apuntan a un rango (`Ventas = Datos!$A$3:$H$40`) y el **área de impresión**.
3. **Filtros** (`autoFilter`) de la hoja.
4. **Regiones**: bloques de celdas llenas contiguas, como los selecciona Ctrl+A, donde el archivo no marca nada. Las filas iniciales con una sola celda (o una combinada) son el **título**, igual que un texto suelto justo encima; la primera fila de textos seguida de datos (o en negrita) es la **cabecera**.

Cada área es `{ kind, name, sheet, sheetName, ref, address, title, headerRow, totalsRow, body, columns }`; `address` (`Ventas!A1:H64`) sirve como destino en cualquier llamada.

`wb.area(destino)` acepta el nombre de una tabla o de un rango con nombre, una dirección con hoja o, con `{ sheet }`, un rango de esa hoja. Devuelve la descripción del área y además:

- `columnsInfo`: por columna, su nombre, tipo (`number`, `text`, `date`, `mixed`…), cuántas celdas tienen valor, suma, mínimo, máximo y media de los números, y cuántas fórmulas hay con un ejemplo (`D4: =B4*C4`).
- `rows`: las primeras `maxRows` filas de datos (30 por defecto) con el texto que muestran las celdas, `firstRow`, `rowCount` y `truncated`.
- `totals`, `merges` y `markdown`: el área como tabla Markdown con números de fila y letras de columna, lista para pasarla a un modelo.

#### Formato

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

#### Fechas

Excel guarda las fechas como números de serie (días desde 1900, o desde 1904 en libros con ese sistema) que se ven como fecha por el formato de la celda. Al leer, una celda numérica con formato de fecha se devuelve como `Date`; al escribir un `Date`, la celda recibe un formato de fecha basado en el que tenía (misma fuente, relleno y bordes), salvo que ya mostrara fechas.

Las fechas de Excel no tienen zona horaria: la hora que muestra Excel se representa como hora **UTC** del `Date`. Escribe con `new Date(Date.UTC(2024, 0, 15))` y lee con `getUTCFullYear()`, `getUTCHours()` o `toISOString()`.

### Sin pérdidas

Solo se reescriben las partes del archivo que cambian (la hoja editada, las cadenas compartidas y, si hace falta, los estilos), y dentro de ellas solo las celdas escritas: gráficos, imágenes, formatos condicionales, validaciones, nombres definidos y todo lo demás se guarda byte a byte tal como estaba.

Cuando un libro cambia, se elimina `xl/calcChain.xml` (Excel lo reconstruye) y se marca `fullCalcOnLoad="1"` para que Excel recalcule las fórmulas al abrirlo: así quedan bien también las que este motor no calcula.

Limitaciones actuales:

- Las fórmulas fuera de lo admitido (ver [Fórmulas](#fórmulas)) no se recalculan: `getValue` devuelve su último resultado guardado.
- Las fórmulas se escriben en la sintaxis en inglés (`SUM`, no `SUMA`); los nombres de función en español no se traducen.
- Escribir en la celda de origen de una fórmula compartida (la que usan otras celdas) da error, para no romper esas celdas.
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
