[← Índice](../README.md) · **Español** · [English](../en/component.md)

# El componente

Componente para ver y editar libros de Excel (`.xlsx`) en el navegador, dibujados en canvas como los muestra Excel. Por dentro usa un motor escrito en Rust y compilado a WebAssembly, que lee el archivo, resuelve estilos y formatos de número, recalcula fórmulas y guarda sin pérdidas. Ese motor también se usa sin interfaz, en Node.js, desde [`@dimarborda/xlsx-editor/engine`](engine.md). Comparte con [`@dimarborda/docx-editor`](../../docx-editor/README.md) la capa de paquete OOXML (`crates/ooxml-core`). Funciona con React, Vue, Svelte, Angular o HTML puro, y en apps de escritorio como Tauri. No necesita servidor: la librería no envía tus archivos a ninguna parte.

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
| Seleccionar filas y columnas | Clic en su encabezado (**Mayús+clic** o arrastrar para varias); clic en la esquina o **Ctrl/Cmd+A** para toda la hoja. **Ctrl+Espacio** selecciona las columnas de la selección y **Mayús+Espacio** sus filas. Al copiar o borrar solo cuentan hasta donde la hoja tiene celdas. |
| Mover filas y columnas | Selecciónalas y arrastra su encabezado hasta el lugar que indica la línea verde (**Escape** cancela), o usa **Alt+↑ ↓** para las filas y **Alt+← →** para las columnas de la selección. Ver [Mover filas y columnas](operations.md#mover-filas-y-columnas). |
| Insertar y eliminar filas y columnas | Clic derecho en una celda o un encabezado (también **Mayús+F10** o la tecla de menú). Con filas o columnas enteras seleccionadas, **Ctrl/Cmd+Mayús++** inserta y **Ctrl/Cmd+-** elimina. Ver [Insertar y eliminar](operations.md#insertar-y-eliminar). |
| Cambiar ancho y alto | Arrastra el borde derecho de un encabezado de columna o el inferior de uno de fila (**Escape** cancela); doble clic en ese borde lo ajusta al contenido. Con varias seleccionadas, se aplica a todas. El menú contextual también ofrece el autoajuste. Desde código: `editor.autofitColumns(primera, cantidad)` y `editor.autofitRows(primera, cantidad)`. |
| Ocultar y mostrar | Clic derecho en encabezados, o **Ctrl+9** (filas) y **Ctrl+0** (columnas); con **Mayús** se muestran. Para mostrar, selecciona las líneas de alrededor. Una línea más gruesa en el encabezado señala dónde hay ocultas. Conservan sus celdas y su tamaño. |
| Rellenar | Arrastra el cuadrito de la esquina de la selección. Doble clic en él rellena hacia abajo hasta donde lleguen los datos de la columna vecina. **Ctrl/Cmd+D** copia hacia abajo y **Ctrl+R** hacia la derecha. Ver [Relleno](operations.md#relleno). |
| Autosuma | **Alt+=** o el menú contextual. En una celda propone `=SUM(…)` con los números de arriba o, si no hay, de la izquierda, para confirmar con **Enter**. Con un rango seleccionado escribe la suma debajo de cada columna (o a la derecha de una sola fila), en su última línea si está vacía o en la siguiente, sin sobrescribir datos. |
| Buscar y reemplazar | **Ctrl/Cmd+F** busca y **Ctrl+H** reemplaza; **Enter** y **Mayús+Enter** recorren los resultados. Ver [Buscar y reemplazar](operations.md#buscar-y-reemplazar). |
| Moverse | **Ctrl/Cmd+flechas** va al borde de los datos, **Inicio** al principio de la fila, **Ctrl/Cmd+Inicio** y **Ctrl/Cmd+Fin** a las esquinas de los datos, **RePág** y **AvPág** una pantalla (con **Alt**, hacia los lados). Con **Mayús** extienden la selección. |
| Deshacer y rehacer | **Ctrl/Cmd+Z**, y **Ctrl/Cmd+Y** o **Cmd+Mayús+Z**. Cada edición, borrado, pegado o movimiento es un paso; deshacer vuelve a la hoja y a la selección de entonces. |

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

- **Fórmulas:** lo que empieza por `=` se guarda como fórmula y se calcula, igual que las fórmulas que dependen de ella. Se escriben en la sintaxis en inglés (`=SUM(A1:A3)`, `,` o `;` entre argumentos) en cualquier idioma del editor. Las funciones admitidas están en el [motor](engine.md#fórmulas); con otras, la celda muestra el último valor guardado y Excel la calcula al abrir el archivo (`fullCalcOnLoad`).
- **Formato:** escribir no cambia el formato de número (salvo en fechas), y las columnas no se ensanchan solas: una fecha en una columna estrecha se ve como `####`, como en Excel.

## Atributos y eventos

| Atributo | Para qué |
| :--- | :--- |
| `src` | URL del `.xlsx` |
| `blank` | Empieza con un libro vacío |
| `locale` | Idioma de los formatos (`es`, `en-US`…); por defecto, el de la página |
| `sheet` | Hoja inicial, por nombre o posición |
| `zoom` | 0.25–4 |
| `tabs="false"` | Oculta las pestañas |
| `formula-bar="false"` | Oculta la barra de celda |
| `status-bar="false"` | Oculta la barra de estado |
| `readonly` | Modo visor: se selecciona y se copia, pero no se edita (se puede quitar y poner en cualquier momento) |
| `quiet` | No escribe los errores en la consola |
| `wasm-url` | Dónde se sirve `xlsx_engine_bg.wasm`, si el bundler no lo hace |

Eventos en el elemento: `ready` (con `detail.editor`), `load`, `sheetchange`, `selectionchange`, `change` y `error`.

`change` llega tras cada edición, borrado, pegado, corte, movimiento, relleno, reemplazo, deshacer o rehacer:

```js
element.addEventListener('change', ({ detail }) => {
  // detail: { sheet: 'Ventas', index: 0, range: 'B3:C4',
  //           source: 'edit' | 'clear' | 'paste' | 'cut' | 'move' | 'insert' | 'delete' | 'fill' | 'replace' | 'undo' | 'redo',
  //           cells: [{ ref: 'B3', value: 1200 }, …] }  // vacío por encima de 200 000 celdas
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
  editor.workbook.addSheet('Notas');  // reload() también redibuja las pestañas
  editor.reload();

  editor.undo();                       // también editor.redo(), editor.canUndo, editor.canRedo
  editor.readOnly = true;              // lo mismo que el atributo readonly

  editor.download('informe.xlsx');
});
```

Los cambios hechos desde código con `editor.workbook` no forman un paso de deshacer propio: llama antes a `editor.workbook.checkpoint()` para que lo sean.

Sin el elemento: `const editor = await XlsxEditor.create(contenedor, { locale: 'es', readOnly: false })`.

## Métodos del editor

Además de los del ejemplo, cada uno como un paso de deshacer:

| Método | Qué hace |
| :--- | :--- |
| `editor.moveRows(primera, cantidad, antesDe)` / `moveColumns` | Mueve filas o columnas, por número, y las selecciona ([detalle](operations.md#mover-filas-y-columnas)). |
| `editor.insertRows(antesDe, cantidad?)` / `insertColumns` | Inserta líneas y las selecciona ([detalle](operations.md#insertar-y-eliminar)). |
| `editor.deleteRows(primera, cantidad?)` / `deleteColumns` | Elimina líneas. |
| `editor.hideRows(primera, cantidad?)` / `showRows` / `hideColumns` / `showColumns` | Oculta o muestra líneas. |
| `editor.autofitColumns(primera, cantidad?)` / `autofitRows` | Ajusta al contenido. Solo existe en el editor, porque mide el texto tal como se dibuja. |
| `editor.fill(destino, { series })` | Rellena desde la selección ([detalle](operations.md#relleno)). |
| `editor.find(texto, opciones)` / `editor.replaceAll(texto, nuevo, opciones)` | Busca y reemplaza ([detalle](operations.md#buscar-y-reemplazar)). |
| `editor.openFind({ replace })` / `editor.closeFind()` | Abre o cierra el panel de búsqueda. |
| `editor.selectionStats` | Recuento, suma, promedio, mínimo y máximo de la selección. |

Devuelven `false` si no cambió nada: el editor es de solo lectura o el motor lo rechazó, y entonces un evento `error` y la barra de estado dicen por qué. `find` devuelve los resultados y `replaceAll` cuántas celdas cambió.
