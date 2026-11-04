[← Índice](../README.md) · **Español** · [English](../en/operations.md)

# Operaciones sobre filas y columnas

## Mover filas y columnas

Mover funciona como "cortar e insertar" en Excel. Van con la fila o columna sus celdas, formatos, alto o ancho, celdas combinadas, comentarios e imágenes ancladas, y se actualizan las fórmulas de todo el libro, los nombres definidos, los gráficos, las tablas, los filtros, las validaciones de datos, el formato condicional y los hipervínculos. Las fórmulas se reescriben como texto, así que también siguen el movimiento las que este motor no evalúa; solo cambian las referencias afectadas.

- Una referencia a una celda sigue a esa celda.
- Un rango conserva sus filas: mover dentro de `SUM(B2:B10)`, también a sus extremos, no lo cambia. Si una fila sale del rango, este se encoge; si entra, crece. Un rango contenido en lo que se mueve se va con ello.
- Las referencias a otros libros (`[1]Hoja!A1`), las 3D (`Hoja1:Hoja3!A1`) y las que están dentro de textos (`INDIRECT("A1")`) se quedan como están.
- Las fórmulas compartidas que cambian pasan a tener cada celda su propia fórmula.

Se rechaza, sin cambiar nada, lo que rompería algo que debe seguir entero:

- **Celdas combinadas** cortadas por el movimiento. Sí se puede reordenar por debajo de una combinada sin tocar su primera línea, que es la que guarda el valor (por ejemplo, columnas bajo un título).
- **Fórmulas matriciales** de varias celdas cortadas por el movimiento.
- **Tablas y filtros**, salvo que se muevan enteros o que se reordenen dentro de ellos: sus filas de datos, sin encabezado ni totales, o sus columnas.
- **Hojas con tablas dinámicas**.

Una regla de formato condicional o de validación sigue a sus celdas, aunque quede repartida en varios rangos. Sus referencias relativas conservan la distancia a cada celda, y las absolutas siguen a su destino.

## Insertar y eliminar

Insertar y eliminar actualizan lo mismo que mover, en todo el libro.

**Al insertar**, las líneas nuevas toman el alto o ancho y el formato de la anterior, sin sus valores. Los rangos, las celdas combinadas y las reglas en cuyo interior caen crecen, igual que las tablas si son filas de datos. Si se insertan dentro de la zona inmovilizada, esta crece. Se rechaza si las últimas líneas de la hoja tienen datos, porque se saldrían.

**Al eliminar**, lo que apuntaba solo a esas líneas pasa a `#REF!`, también en otras hojas, nombres y gráficos, como en Excel. Los rangos y las celdas combinadas se encogen. Desaparecen con sus celdas los comentarios, los hipervínculos y las reglas que no aplican a nada más. Las imágenes ancladas en ellas no se borran: pasan a la línea que ocupa su lugar.

Se rechaza, sin cambiar nada:

- Insertar o eliminar **columnas dentro de una tabla**, eliminar **una tabla entera**, su encabezado, sus totales o todas sus filas de datos.
- Insertar o eliminar cortando una **fórmula matricial** de varias celdas.
- Eliminar una columna con un **filtro aplicado**.
- Cualquiera de estas operaciones en una **hoja con tablas dinámicas**.

## Relleno

A lo largo de cada fila o columna rellenada se repiten las celdas de origen con su formato, y:

- Las **fórmulas** se reescriben para cada celda: las referencias relativas se desplazan y las absolutas no.
- Dos o más **números** siguen su tendencia (la recta que mejor se ajusta, como Excel). Un número solo se copia. Una **fecha** sola avanza de día en día.
- Un **texto que acaba en número** sigue contando (`Lote 9`, `Lote 10`; `L08`, `L09`).
- Los **meses y días**, completos o abreviados, en español o inglés, siguen su orden y respetan las mayúsculas.
- Lo demás se copia tal cual, y una celda de origen vacía vacía su destino.

Con `{ series: false }`, y con **Ctrl/Cmd+D** y **Ctrl+R**, solo se copia (las fórmulas siguen desplazándose). No se puede rellenar sobre celdas combinadas. Reducir la selección con el cuadrito no borra nada, a diferencia de Excel.

## Buscar y reemplazar

`*` es cualquier texto, `?` un carácter y `~` los toma literalmente. Por defecto se busca en el texto que muestran las celdas de la hoja activa; el panel y las opciones permiten distinguir mayúsculas, exigir la celda completa, mirar las fórmulas o abarcar todo el libro (sin las hojas ocultas). Desde código: `editor.find(texto, opciones)` y `editor.replaceAll(texto, nuevo, opciones)`.

Reemplazar actúa sobre las celdas de texto y sobre el texto de las fórmulas, nunca sobre números ni fechas. El resultado se interpreta como lo escrito a mano, así que `12` queda como número, y "Todos" es un solo paso de deshacer. El texto con varios formatos dentro de una celda queda con un solo formato.

## Barra de estado

Con más de una celda seleccionada, la **barra de estado** muestra el promedio, el recuento y la suma, como Excel: los números con el formato del primero de ellos, sin contar filas y columnas ocultas, y solo el recuento si hay un error en la selección. `editor.selectionStats` da esos datos (y el mínimo y el máximo) desde código, y `status-bar="false"` oculta la barra. Ahí mismo aparece también el motivo de un cambio rechazado.
