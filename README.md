# 📄 Rust DOCX WebAssembly Text Modifier (`rust-web-docx`)

Una aplicación web moderna y de ultra alto rendimiento que permite **modificar el texto de documentos `.docx` directamente desde el navegador** utilizando **WebAssembly (WASM) programado en Rust**, sin herramientas pesadas de edición (WYSIWYG) ni dependencias de backend/servidores.

---

## 🚀 Características Principales

- **🎨 Motor de Renderizado en Canvas 2D Acelerado por GPU (Opción Google Docs / Figma)**:
  - Rust + WebAssembly calcula la geometría de fuentes, márgenes, tablas, saltos de página y dibuja directamente en lienzos `<canvas>` de alta definición (Retina DPI).
  - **Cero sobrecarga de DOM (*Zero DOM Bloat*)**: Renderizado continuo a 60 FPS sin importar si el documento tiene 5 o 500 páginas.
  - **Paginación Física A4 Multipage**: Simulación exacta de hojas físicas (`Página 1 de N`, `Página 2 de N`).
  - **Selector de Motor Dual**: Conmutador instantáneo entre **Motor Canvas 2D** y **Modo Vista HTML**.
  - **Zoom Vectorial**: Ajuste de escala al 75%, 100%, 125% y 150% con renderizado ultra nítido.
- **⚡ 100% Client-Side con Rust + WebAssembly**:
  - Todo el procesamiento de desempaquetado ZIP, análisis de OpenXML (`word/document.xml`), sustitución de texto y re-empaquetado ZIP se realiza en memoria en el navegador.
  - **Máxima privacidad y seguridad**: Ningún byte o documento sale del dispositivo del usuario.
- **📊 Soporte Completo de Tablas OpenXML**:
  - Renderizado fiel de tablas con bordes, cabeceras y celdas interactivas editables en vivo.
- **📐 Formato Enriquecido**:
  - Alineaciones (Izquierda, Centro, Derecha, Justificado), colores de texto (`<w:color>`), negrita y cursiva.
- **🖼️ Imágenes de Fondo y Marcas de Agua**:
  - Soporte de color de página y marcas de agua de texto (*CONFIDENCIAL*, *BORRADOR*) o imágenes personalizadas con control de opacidad.
- **🧪 Generador de Documento Demo (1-Click)**:
  - Generador de contratos `.docx` válido en memoria desde Rust para probar la aplicación al instante.
- **💾 Exportación Instantánea (`⌘S` / `Ctrl+S`)**:
  - Descarga directa del archivo `.docx` modificado listo para abrirse en Microsoft Word, Google Docs o LibreOffice.

---

## 🛠️ Estructura del Proyecto

```
/Volumes/DirWork/2026/rust-web-docx/
├── Cargo.toml               # Configuración del crate en Rust (wasm-bindgen, zip, regex, serde)
├── src/
│   ├── lib.rs               # API expuesta a WebAssembly (DocxSession, funciones wasm_bindgen)
│   ├── docx_parser.rs       # Motor de manipulación OpenXML, búsqueda/reemplazo y zip
│   └── sample_generator.rs  # Generador de documentos DOCX de prueba
├── pkg/                     # Paquete compilado WebAssembly generado por wasm-pack
├── index.html               # Interfaz de usuario principal
├── src_web/
│   ├── main.js              # Lógica de frontend, eventos y puente con WebAssembly
│   └── style.css            # Estilos modernos con paleta oscura, glassmorphism y diseño responsivo
├── package.json             # Scripts de npm y configuración de dependencias Vite
├── vite.config.js           # Configuración de Vite con soporte nativo Wasm y Top-Level Await
└── README.md                # Documentación del proyecto
```

---

## 💻 Requisitos Previos

1. **Rust y Cargo** (con el target `wasm32-unknown-unknown`):
   ```bash
   rustup target add wasm32-unknown-unknown
   ```
2. **Node.js** (v18+) y **npm**.

---

## 🏃‍♂️ Instrucciones de Ejecución

### 1. Iniciar en Modo Desarrollo (Hot Reload)

```bash
cd /Volumes/DirWork/2026/rust-web-docx
npm run dev
```
Abre en tu navegador: **`http://localhost:5173`**

### 2. Compilar WebAssembly manualmente

Si modificas el código en Rust (`src/`):
```bash
npm run build:wasm
```

### 3. Compilar para Producción

```bash
npm run build
```
Los archivos finales optimizados se generarán en la carpeta `dist/`.

### 4. Ejecutar Pruebas Unitarias en Rust

```bash
cargo test
```

---

## ⚙️ Cómo Funciona la Manipulación de DOCX en Rust

Un archivo `.docx` es un archivo comprimido en formato ZIP que contiene especificaciones OpenXML (`ISO/IEC 29500`):
1. **Extracción**: El crate `zip` lee los bytes del búfer en memoria (`Uint8Array`).
2. **Normalización XML**: El parser inspecciona `word/document.xml` y los archivos de encabezado/pie (`word/header*.xml`, `word/footer*.xml`).
3. **Manejo de Runs (`<w:r><w:t>`)**: En DOCX, el texto está dentro de etiquetas `<w:t>`. Rust localiza estas etiquetas, realiza el reemplazo de caracteres respetando entidades XML (`&amp;`, `&lt;`, `&gt;`, `&quot;`, `&apos;`) y sincroniza los fragmentos.
4. **Re-empaquetado**: Se genera un nuevo archivo ZIP manteniendo intactos todos los metadatos, temas, imágenes y relaciones originales, garantizando que el documento resultante sea 100% compatible con Microsoft Word y Google Docs.

---

## 🛡️ Licencia

MIT - Desarrollado para modificación ligera y segura de documentos DOCX en navegador con Rust & WebAssembly.
