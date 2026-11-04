[← Índice](../README.md) · **Español** · [English](../en/agents.md)

# Áreas y contexto para agentes

Un agente no necesita leer toda la hoja: puede pedir las áreas y luego el contexto de la que le interesa.

## Áreas

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

## Desde el editor

- `editor.selectArea(destino)`: selecciona un área y muestra su hoja. Acepta una tabla o un rango con nombre, una dirección (`Ventas!A1:H40`) o un rango de la hoja activa. Devuelve el área (ver [Áreas](#áreas)).
- `editor.selectionContext({ maxRows })`: contexto de lo seleccionado para un agente: el rango, o con una sola celda la región que la rodea (su tabla, o el bloque que seleccionaría Ctrl+A). Devuelve título, cabecera, tipos y totales por columna, las primeras filas y una tabla Markdown.

Para dar formato desde código se usa `editor.workbook.sheet(…).setFormat(…)` y después `editor.reload()`.
