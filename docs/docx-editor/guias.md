# Guías de @dimarborda/docx-editor

[← Índice](README.md) · [API](api.md) · [Guías](guias.md) · [Integraciones](integraciones.md) · [Seguridad y limitaciones](seguridad.md)

Secciones lógicas, marcadores, actividad del agente, tablas, PDF, solo lectura e imágenes.

## Secciones lógicas

Las secciones lógicas marcan dónde empieza y dónde termina cada parte de un documento (por ejemplo «Introducción», «Caracterización» y «Conclusiones» de un informe técnico), para leerlo, llenarlo o revisarlo sección por sección. No cambian el aspecto del documento ni tienen que ver con los saltos de sección de Word.

```js
// Marcar como «introduccion» los párrafos 2 a 6 (las tablas entre ellos van completas)
editor.defineSection('introduccion', { from: 2, to: 6 }, { title: 'Introducción', locked: true });

editor.sections();
// [{ key: 'introduccion', title: 'Introducción', first: 2, last: 6, depth: 0, parent: null, locked: true }]

editor.sectionContent('introduccion').text;   // texto plano de la sección
editor.replaceSection('introduccion', ['Nuevo párrafo.', 'Otro más.'], { keepHeading: true });
editor.unwrapSection('introduccion');         // quita la marca, deja el contenido
```

En el `.docx` cada sección es un control de contenido de bloque (`w:sdt`) con `w:tag="section:<clave>"` y el título en `w:alias`, así que Word la conserva al guardar y cualquier otra herramienta (python-docx, Open XML SDK, docx4j) la encuentra buscando esa etiqueta. Las secciones pueden anidarse (`depth`, `parent`). Con `locked: true`, en Word no se puede borrar la sección, pero su contenido sigue siendo editable. Cada llamada que modifica es un paso de deshacer y lanza un `Error` sin tocar el documento si la clave ya existe, falta o está repetida, o si el rango corta una tabla u otra sección, queda dentro de una tabla o contiene un salto de sección.

## Marcadores: campos y pendientes

Un léxico estándar, compartido con [`@dimarborda/xlsx`](../../packages/xlsx), para que un agente o una automatización sepa qué falta en un documento sin adivinar:

| Escritura | Tipo | Significado |
| :--- | :--- | :--- |
| `{{CLAVE}}`, `{{ CLAVE }}`, `{{CLAVE: pista}}` | `field` | Valor que rellena una automatización (forma canónica). La pista tras los dos puntos dice qué se espera (`{{FECHA: dd/mm/aaaa}}`) |
| `{clave}`, `{clave:pista}` | `field` | Forma corta, sin espacios dentro |
| `[…]`, `[...]`, `[____]`, `[Completar: alcance del proyecto]` | `pending` | Contenido que alguien (persona o agente) todavía tiene que escribir. El texto de dentro es la pista |

No son marcadores las citas (`[1]`, `[2, 3]`, `[4-6]`), las casillas (`[ ]`, `[x]`) ni `[sic]`, y ningún marcador cruza un salto de línea.

```js
editor.markers({ kind: 'pending' });
// [{ id: '12:9', kind: 'pending', syntax: 'bracket', raw: '[Completar: alcance]', key: null,
//    hint: 'Completar: alcance', paragraph: 12, offset: 9, length: 20, place: 'body', section: 'objeto' }]

const [first] = editor.markers({ kind: 'pending', section: 'objeto' });
editor.fillMarker(first, 'Diseño e implementación del módulo de pagos.'); // un paso de deshacer
```

`place` es `body`, `table`, `textbox` o `headerFooter`; `section` es la [sección lógica](#secciones-lógicas) más interna que lo contiene. `fillMarker` vuelve a encontrar el marcador aunque se haya escrito texto antes de él desde que se listó, y lanza un `Error` si ya no está. Para analizar un texto suelto (por ejemplo, lo que va a insertar un agente) sin documento: `import { findMarkers } from '@dimarborda/docx-editor'`.

## Actividad: dónde está trabajando el agente

Cuando un agente modifica el documento por API, la persona que lo está leyendo debe poder ubicarse. Hay dos niveles, ambos solo de vista (no se guarda nada en el `.docx`, no mueven el cursor ni el scroll del usuario, y no cuestan nada si no se usan):

- **Destello automático**: toda edición hecha desde código (`insertParagraphs`, `replaceParagraphs`, `replaceSection`, `setTableCell`, `insertTableRows`, `replaceVariables`, `fillMarker`) resalta lo que cambió durante un momento. Duración con la opción `flash` (ms, predeterminado `1800`; `false` lo desactiva).
- **Marca de trabajo**: `markActivity()` deja un área tintada con una barra y una etiqueta en el margen, más un chip fijo arriba del editor que indica dónde está (↑ / ↓ cuando queda fuera de la vista) y lleva hasta allí con un clic.

```js
const work = editor.markActivity('conclusiones', { label: 'Agente · redactando' }); // una sección: la sigue aunque crezca
const added = editor.replaceSection('conclusiones', ['Primer borrador…']);
work.update(added, { label: 'Agente · revisando' }); // también { from, to } o { paragraph, offset, length }
work.done();

editor.addEventListener('activity', e => console.log(e.detail.activities));
// [{ id: 'a1', kind: 'work', label: 'Agente · redactando', from: 14, to: 18, section: 'conclusiones', page: 2 }]
```

El objetivo puede ser la clave de una sección (recomendado: sigue a la sección), un rango `{ from, to }`, lo que devuelven `insertParagraphs` / `replaceSection` (`{ first, count }`) o un tramo `{ paragraph, offset, length }` como un marcador. Los rangos de párrafos se mantienen en su sitio cuando el usuario o el código insertan o borran párrafos antes de ellos. `ttl` (ms) la quita sola; `color` cambia su color (útil con varios agentes). Si tu aplicación tiene una barra fija arriba, baja el chip con `--docx-activity-dock-top`.

## Tablas existentes

Para completar una tabla de una plantilla (por ejemplo, las filas de una matriz de riesgos), ubícala con `tables()` y escribe en ella:

```js
// Cada tabla del cuerpo: { index, rows, cols, header, first, last }
const matriz = editor.tables().find(t => t.header[0] === 'Riesgo');
if (!matriz) throw new Error('La plantilla no tiene la matriz de riesgos');

// Dos filas al final, con el formato de la última fila (bordes, sombreado, fuente…)
const { first } = editor.insertTableRows(matriz.index, matriz.rows, [
  ['Caída del servidor', 'Media', 'Alto'],
  ['Fuga de datos', 'Baja', 'Crítico\nrevisar cada mes'],
]);

// Corregir una celda: conserva el formato del texto que había
editor.setTableCell(matriz.index, first, 1, 'Alta');

editor.undo();   // deshace solo la corrección
editor.undo();   // deshace las dos filas de una vez
```

- **Índices.** Las tablas se numeran desde 0 en el orden del documento, sin contar las anidadas ni las de cuadros de texto. Las filas y columnas también empiezan en 0. Las columnas son las de la cuadrícula: a una celda combinada horizontalmente se le escribe por su primera columna. `first`/`last` son el rango de párrafos de sus celdas, con la misma numeración de `structure()` y de `selection`.
- **`setTableCell`** reemplaza el texto de la celda y conserva sus propiedades de párrafo y el formato de su primer *run* (fuente, tamaño, color, negrita). `\n` es un salto de línea dentro de la celda.
- **`insertTableRows(table, at, rows, { copyFormatFrom })`** inserta antes de la fila `at`: `0` al principio, `rows` (la cantidad de filas) al final. Cada fila nueva copia la fila `copyFormatFrom` (por defecto la anterior a `at`, o la siguiente si insertas al principio): alto, bordes, sombreado, ancho de celdas, propiedades de párrafo y formato del primer *run*. Las filas con menos valores se completan con celdas vacías, que conservan el formato para lo que se escriba después; las filas con más valores que celdas producen un error. Devuelve `{ first, count }`: la primera fila nueva y cuántas se insertaron. Una fila de encabezado repetido solo se copia como encabezado si la insertas al principio. Ojo al insertar justo debajo del encabezado: la fila anterior es el encabezado y se copiaría su sombreado; pasa `copyFormatFrom` con una fila de datos. Lo mismo si las filas vacías de la plantilla tienen texto de relleno en gris o cursiva: copia una fila ya completada.
- **`deleteTableRows(table, from, count)`** elimina filas; la tabla no puede quedar sin filas (para quitarla entera usa `replaceParagraphs` con su rango `first`–`last`).
- Cada llamada es un solo paso de deshacer, dispara `change`, vuelve a dibujar el documento y funciona también en solo lectura.
- Lanzan un `Error` claro, sin modificar nada, si un índice no existe, si la celda continúa una combinación vertical (el mensaje dice en qué fila escribir), si la fila que se copia o la posición tienen celdas combinadas verticalmente, o si la celda tiene una tabla anidada.

## PDF

```js
const pdf = await editor.exportPdf();            // Uint8Array
await editor.downloadPdf('contrato.pdf');         // o descargarlo directamente
const blob = await editor.exportPdfBlob({ watermark: false });
```

Genera un PDF vectorial de las páginas tal como se ven en el editor: texto real (seleccionable y buscable) con las fuentes incrustadas y recortadas a los caracteres usados, imágenes, tablas, bordes, formas, cuadros de texto, encabezados y pies con su número de página. La marca de agua de `setWatermark()` se incluye salvo `{ watermark: false }`. Las fuentes son las que sube el usuario con `registerFont()` o las sustitutas métricas del paquete (`@dimarborda/docx-editor/fonts`); sin ellas se usan las fuentes estándar del PDF, y el texto se ajusta al ancho medido en pantalla para que nada se desplace. [pdf-lib](https://pdf-lib.js.org/) se descarga solo la primera vez que se exporta.

## Solo lectura

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

## Imágenes

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
