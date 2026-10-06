# 📄 Rust Web DOCX (`rust-web-docx`)

Un motor de renderizado y editor visual interactivo de documentos **`.docx` (Microsoft Word)** que se ejecuta **100% en el navegador web** utilizando **Rust compilado a WebAssembly (WASM)** y renderizado en **Canvas 2D acelerado por GPU**.

---

## 🌟 Visión y Evolución del Proyecto

El proyecto fue concebido originalmente como una herramienta ligera de búsqueda y reemplazo de texto en documentos DOCX sin servidores. 

Conforme evolucionó, se transformó en un **motor completo de procesamiento, maquetación tipográfica y edición visual en Canvas** (siguiendo la arquitectura moderna utilizada por herramientas de alto rendimiento como *Google Docs* o *Figma*). Hoy en día ofrece:

- **Maquetación multipágina precisa** con cálculo de métricas de fuentes, saltos de página, sangrías francesas/primera línea, interlineados y tablas.
- **Edición interactiva en tiempo real sobre Canvas** con cursor nativo (*caret*), selección de texto con ratón y teclado, y soporte completo de atajos (`Ctrl+Z`, `Ctrl+Y`, `Ctrl+S`, `Ctrl+A`).
- **Manipulación OpenXML sin pérdidas (*lossless*)**: Las modificaciones preservan intacta la jerarquía de estilos, metadatos, temas, imágenes, saltos de sección y relaciones del archivo original.
- **Privacidad y Seguridad Absoluta**: Ningún documento ni byte viaja a servidores externos; todo el ciclo (lectura ZIP, análisis XML, layout, edición y reempaquetado) ocurre en la memoria del navegador del cliente.

---

## 🚀 Características Principales

### 🎨 Motor de Renderizado en Canvas 2D
- **Paginación Física Multipage**: Simulación exacta de hojas (A4 / Carta) con márgenes superior, inferior, izquierdo y derecho.
- **Cero sobrecarga de DOM (*Zero DOM Bloat*)**: Renderizado fluido a 60 FPS sin saturar el árbol HTML, incluso en documentos de cientos de páginas.
- **Soporte Retina / HiDPI**: Escala adaptativa al `devicePixelRatio` para una nitidez tipográfica total.
- **Zoom Vectorial Fluido**: Visualización ajustable (75%, 100%, 125%, 150%, 200%).
- **Marcas de Agua y Fondos**: Soporte para colores de fondo de página y marcas de agua de texto o imagen con control de opacidad.

### 🎯 Cursor Interactivo y Selección en Canvas (`Caret Engine`)
- **Hit-Testing Bidireccional**: Conversión exacta de coordenadas de clic en pantalla a párrafos, líneas y desplazamientos de caracteres (*offsets*).
- **Selección de Rango Completa**: Soporte de selección multilínea y multipárrafo mediante arrastre de ratón o teclado (`Shift + Flechas`, `Home`, `End`, `⌘A / Ctrl+A`).
- **Edición Estructural en Vivo**: Inserción de texto, saltos de párrafo (`Enter`), fusión de bloques (`Backspace`, `Delete`) y pegado de texto multilínea manteniendo los estilos del contexto.
- **Historial Deshacer / Rehacer**: Sistema de historial en memoria para revertir o repetir cambios instantáneamente (`⌘Z` / `⌘Y`).

### 📑 Cascada de Estilos y Tipografía OpenXML
- **Resolución en Cascada de 4 Niveles**:
  $$\text{docDefaults} \longrightarrow \text{styles.xml} \longrightarrow \text{numbering.xml} \longrightarrow \text{direct w:pPr / w:rPr}$$
- **Manejo Preciso de Espaciados**: Espaciado anterior/posterior (`space before/after`) e interlineados automáticos, exactos o mínimos (`lineRule`).
- **Sangrías y Tabuladores**: Sangría izquierda, derecha, primera línea (`firstLine`) y sangría francesa (`hanging`).
- **Listas Multinivel**: Renderizado y cálculo de etiquetas automáticas de numeración (`1.`, `a)`, `i.`, viñetas) y tabulación según la definición del documento.

### 📊 Tablas OpenXML
- Soporte para estructuras `<w:tbl>`, filas `<w:tr>` y celdas `<w:tc>`.
- Renderizado de bordes personalizados (estilo, grosor, color), márgenes de celda y cálculo de anchos de columna.
- Edición de texto celda por celda integrada en el motor de cursor.

### 🔍 Búsqueda y Reemplazo Avanzado
- Búsqueda simple, insensible a mayúsculas/minúsculas y soporte de expresiones regulares (Regex).
- Reemplazo seguro de variables distribuidas a través de múltiples fragmentos de texto (`<w:r>`).

### 🧪 Generador de Documento Demo (1-Click)
- Generación instantánea en memoria desde Rust de un documento DOCX de prueba completo con tablas, listas numeradas y encabezados para probar la aplicación sin necesidad de subir archivos.

### 💾 Exportación Instantánea
- Empaquetado ZIP en memoria generando un `.docx` 100% válido bajo la norma ISO/IEC 29500, listo para abrirse en Microsoft Word, Google Docs o LibreOffice.

---

## 🛠️ Arquitectura del Proyecto

```text
.
├── Cargo.toml               # Dependencias de Rust (wasm-bindgen, zip, quick-xml, serde, regex)
├── src/                     # Núcleo del motor en Rust (compilado a WebAssembly)
│   ├── lib.rs               # API pública WebAssembly (DocxSession, DocxModifier) y puente JS
│   ├── docx_parser.rs       # Extracción ZIP, parser OpenXML, tablas y jerarquía de estilos
│   ├── layout_engine.rs     # Motor de maquetación geométrica, saltos de línea y render commands
│   ├── caret.rs             # Geometría de cursor, hit-testing y selección de texto en canvas
│   ├── paragraph_edit.rs    # Mutación estructural de OpenXML y edición de runs sin pérdidas
│   ├── styles.rs            # Árbol de herencia de estilos (styles.xml) y numeración (numbering.xml)
│   └── sample_generator.rs  # Generador de documentos DOCX de prueba en memoria
├── tests/                   # Suite completa de pruebas unitarias y de integración en Rust
├── src_web/                 # Frontend web moderno (Vanilla JS + CSS)
│   ├── main.js              # Controlador de UI, carga de archivos, barra de herramientas y estado
│   ├── canvas_editor.js     # Manejador de eventos Canvas (puntero, teclado, caret blinking, IME)
│   └── style.css            # Estilos del visor de páginas, toolbar y tema oscuro
├── index.html               # Punto de entrada de la aplicación web
├── package.json             # Scripts de npm y dependencias de Vite / wasm-pack
├── vite.config.js           # Configuración de Vite con soporte para WebAssembly
└── README.md                # Documentación del proyecto
```

---

## 💻 Requisitos Previos

- **Rust y Cargo** (edición 2021 o superior) con el target `wasm32-unknown-unknown`:
  ```bash
  rustup target add wasm32-unknown-unknown
  ```
- **wasm-pack** (opcional, para compilación optimizada de WASM):
  ```bash
  cargo install wasm-pack
  ```
- **Node.js** (v18 o superior) y **npm**.

---

## 🏃‍♂️ Instrucciones de Instalación y Uso

### 1. Clonar e Instalar Dependencias

```bash
git clone <URL_DEL_REPOSITORIO>
cd rust-web-docx
npm install
```

### 2. Iniciar en Modo Desarrollo (con Hot-Reload)

```bash
npm run dev
```
Abre tu navegador en `http://localhost:5173`.

### 3. Compilar el Módulo WebAssembly

Si realizas cambios en el código de Rust (`src/`):
```bash
npm run build:wasm
```

### 4. Ejecutar la Suite de Pruebas en Rust

El proyecto cuenta con más de 70 pruebas unitarias y de integración que validan el parser OpenXML, la geometría del layout, las mutaciones de párrafos y las tablas:

```bash
cargo test
```

### 5. Compilar para Producción

```bash
npm run build
```
Los archivos finales optimizados para distribución se generarán en la carpeta `dist/`.

---

## ⌨️ Atajos de Teclado en el Editor Canvas

| Atajo | Acción |
| :--- | :--- |
| `Click` | Posicionar el cursor en el texto |
| `Arrastrar Ratón` | Seleccionar un rango de texto |
| `Shift + Flechas` | Expandir o contraer selección de texto |
| `Home` / `End` | Ir al inicio / fin de la línea actual |
| `Ctrl + A` / `⌘A` | Seleccionar todo el documento |
| `Ctrl + Z` / `⌘Z` | Deshacer (*Undo*) |
| `Ctrl + Y` / `⌘Shift+Z` | Rehacer (*Redo*) |
| `Ctrl + S` / `⌘S` | Descargar / Exportar documento `.docx` |
| `Enter` | Nuevo párrafo manteniendo el formato |
| `Backspace` / `Delete` | Borrar carácter o unir párrafos |

---

## 👨‍💻 Autoría

<p align="center">
  <img src="./src_web/assets/firma-qr.png" alt="Código QR de Dimar Borda" width="130" />
</p>

<p align="center">
  <strong>Desarrollado por Dimar Borda</strong><br>
  <sub>Escanea el código QR para conocer más sobre el autor y proyectos relacionados.</sub>
</p>

---

## 🛡️ Licencia

Distribuido bajo la Licencia **MIT**. Consulta el archivo `LICENSE` para más información.
