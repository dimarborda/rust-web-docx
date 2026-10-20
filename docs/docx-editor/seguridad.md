# Seguridad y limitaciones de @dimarborda/docx-editor

[← Índice](README.md) · [API](api.md) · [Guías](guias.md) · [Integraciones](integraciones.md) · [Seguridad y limitaciones](seguridad.md)

Qué protege el editor y qué partes de Word todavía no dibuja.

## Seguridad y privacidad

- **Los documentos no salen del navegador.** Abrir, editar, guardar y exportar a PDF ocurren en la página. No hay servidor propio, telemetría ni peticiones a terceros. Las únicas descargas son las del propio sitio: el motor `.wasm`, las fuentes y el `.docx` del atributo `src`, si se usa. Las fuentes que registra el usuario con `registerFont` se guardan solo en el IndexedDB de ese navegador.
- **El contenido del documento no se ejecuta.** Las páginas se dibujan en un `<canvas>` y el texto nunca se inserta como HTML, así que un `.docx` malicioso no puede inyectar scripts (XSS). Las imágenes se leen solo del propio archivo; los enlaces a recursos externos no generan peticiones.
- **XML sin entidades externas.** El parser (quick-xml) no resuelve DTD ni entidades externas, lo que evita ataques XXE.
- **Límites contra bombas ZIP.** Se rechazan los archivos con más de 10 000 partes, con una parte de más de 256 MB o con más de 512 MB en total al descomprimirse. El tamaño se mide mientras se descomprime, sin fiarse del que declara el ZIP. Las imágenes insertadas por API tienen un máximo de 15 MB.
- **Dependencias auditadas.** Cada despliegue del repositorio ejecuta `cargo audit` y `npm audit` y se detiene si hay vulnerabilidades conocidas. En tiempo de ejecución el paquete solo depende de `pdf-lib`, `@pdf-lib/fontkit` (MIT) y de las fuentes Carlito, Caladea, Arimo, Tinos y Cousine de Fontsource (SIL OFL 1.1).
- **CSP estricta.** Basta con añadir `'wasm-unsafe-eval'` al `script-src`; no hace falta `'unsafe-eval'` en scripts (ver [Content Security Policy](integraciones.md#content-security-policy)).

Las vulnerabilidades se reportan en privado, como se explica en [SECURITY.md](https://github.com/dimarborda/rust-web-docx/blob/master/SECURITY.md).

## Limitaciones

Todavía no se dibujan las formas agrupadas, los comentarios, el control de cambios ni las columnas múltiples. Todo eso se conserva intacto al guardar el documento. Los encabezados y pies de página se dibujan (texto, tablas, imágenes, número de página y primera página distinta) y se editan con doble clic sobre ellos, como en Word: aparece una línea punteada con la etiqueta "Encabezado" o "Pie de página" y el texto se escribe, se borra y se formatea como el del cuerpo; Esc o un clic en el cuerpo vuelven a él. No se insertan tablas ni imágenes en ellos. Se usan el encabezado y el pie de la última sección (sin variante para páginas pares), y en una línea con número de página el cursor puede desfasarse si el número mostrado tiene otra cantidad de cifras que el guardado. En la selección, sus párrafos usan índices a partir de 33 554 432 (2²⁵).

Cuadros de texto: un clic dentro coloca el cursor en su texto (escribir, Enter, borrar, formato, deshacer) y un clic en su borde selecciona el cuadro para moverlo, redimensionarlo o cambiar su ajuste como una imagen. La selección y las flechas no salen del cuadro, y ⌘A selecciona todo su texto. No se pueden insertar tablas ni imágenes flotantes dentro de un cuadro. Las tablas dentro de un cuadro y el giro del cuadro no se dibujan, y el texto que no cabe no se recorta. Los cuadros VML de documentos antiguos se dibujan, pero no se pueden editar. En la selección (`selection`), los párrafos de los cuadros usan índices a partir de 16 777 216 (2²⁴), para no desplazar la numeración del cuerpo.

Imágenes flotantes: el texto las rodea por ambos lados cuando hay sitio (`bothSides`), por el lado indicado (`left`, `right`, `largest`) y siguiendo su contorno en el ajuste estrecho y transparente (`wrapPolygon`); una fila de tabla que chocaría con ellas baja hasta pasarlas. Los objetos que se pisan se dibujan en el orden de Word (`relativeHeight`). Solo desplazan el texto que viene después de su párrafo de anclaje en la misma página, y las imágenes de encabezados y pies de página se muestran, pero no se pueden seleccionar.
