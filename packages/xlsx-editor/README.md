# @dimarborda/xlsx-editor

Componente para ver libros de Excel (`.xlsx`) en el navegador, dibujados en canvas como los muestra Excel. Por dentro usa el motor Rust/WebAssembly de [`@dimarborda/xlsx`](../xlsx), que lee el archivo, resuelve estilos y formatos de número y guarda sin pérdidas. Funciona en cualquier frontend (HTML, React, Vue, Svelte, Angular, Tauri) y nada sale del navegador.

> **Estado: fase 1, visor.** Todavía no se editan celdas desde la cuadrícula: eso llegará en la fase 2. Desde código sí se pueden cambiar valores con `editor.workbook` y descargar el resultado.

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

Aún no se dibujan gráficos, imágenes, formato condicional ni comentarios, y el texto girado se dibuja horizontal. Las fórmulas muestran el último valor que guardó Excel; no se recalculan.

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
| `wasm-url` | Dónde se sirve `xlsx_engine_bg.wasm`, si el bundler no lo hace |

Eventos en el elemento: `ready` (con `detail.editor`), `load`, `sheetchange`, `selectionchange` y `error`.

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

  editor.download('informe.xlsx');
});
```

Sin el elemento: `const editor = await XlsxEditor.create(contenedor, { locale: 'es' })`.

## Content Security Policy

Como el motor es WebAssembly, el `script-src` de producción necesita `'wasm-unsafe-eval'`.

## Desarrollo

Desde la raíz del repositorio:

```bash
npm run build:wasm:xlsx
npm test -w packages/xlsx-editor
npm run dev        # http://localhost:5173/xlsx.html
```
