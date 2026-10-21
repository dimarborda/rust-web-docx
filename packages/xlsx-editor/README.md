# @dimarborda/xlsx-editor

Componente para ver y editar libros de Excel (`.xlsx`) en el navegador, dibujados en canvas como los muestra Excel. Por dentro usa el motor Rust/WebAssembly de [`@dimarborda/xlsx`](../xlsx), que lee el archivo, resuelve estilos y formatos de número y guarda sin pérdidas. Funciona en cualquier frontend (HTML, React, Vue, Svelte, Angular, Tauri) y nada sale del navegador.

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

Aún no se dibujan gráficos, imágenes, formato condicional ni comentarios, y el texto girado se dibuja horizontal. Las fórmulas se recalculan tras cada cambio con el motor de `@dimarborda/xlsx`; las que usan algo que no admite muestran el último valor que guardó Excel.

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

Los valores se escriben con `sheet.setValues` de `@dimarborda/xlsx`, que conserva el formato de la celda: un número escrito en una celda con formato de moneda se ve como moneda. A las fechas se les da un formato de fecha basado en el de la celda si no lo tenía. Los porcentajes y las horas se guardan como su valor (0,15 o 0,604…) y se ven con el formato que ya tenga la celda. La misma función está disponible como `parseInput(texto, locale)`.

### Limitaciones

- **Fórmulas:** lo que empieza por `=` se guarda como fórmula y se calcula, igual que las fórmulas que dependen de ella. Se escriben en la sintaxis en inglés (`=SUM(A1:A3)`, `,` o `;` entre argumentos) en cualquier idioma del editor. Las funciones admitidas están en el [README de `@dimarborda/xlsx`](../xlsx#fórmulas); con otras, la celda muestra el último valor guardado y Excel la calcula al abrir el archivo (`fullCalcOnLoad`).
- **Formato:** escribir no cambia el formato de número (salvo en fechas), y las columnas no se ensanchan solas: una fecha en una columna estrecha se ve como `####`, como en Excel.
- No se puede sobrescribir la celda de origen de una fórmula compartida (la que usan otras celdas); el editor emite `error` y deja el libro como estaba.

### Áreas y contexto para agentes

- `editor.selectArea(destino)`: selecciona un área y muestra su hoja. Acepta una tabla o un rango con nombre, una dirección (`Ventas!A1:H40`) o un rango de la hoja activa. Devuelve el área (ver `workbook.areas()` en [`@dimarborda/xlsx`](../xlsx#áreas-para-agentes)).
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

  // Cambiar celdas con la API de @dimarborda/xlsx y volver a dibujar
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

## Content Security Policy

Como el motor es WebAssembly, el `script-src` de producción necesita `'wasm-unsafe-eval'`.

## Desarrollo

Desde la raíz del repositorio:

```bash
npm run build:wasm:xlsx
npm test -w packages/xlsx-editor
npm run dev        # http://localhost:5173/xlsx.html
```
