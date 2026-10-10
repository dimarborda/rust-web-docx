# Política de seguridad

## Versiones con soporte

Solo la última versión menor publicada en npm recibe correcciones de seguridad.

| Versión | Soporte |
| :--- | :--- |
| 0.7.x | ✅ |
| < 0.7 | ❌ (actualiza a la última versión) |

## Cómo reportar una vulnerabilidad

**No abras un issue público.** Usa el reporte privado de GitHub: en este repositorio, pestaña **Security → Report a vulnerability**. El reporte solo lo ven los mantenedores.

Incluye, si puedes:

- la versión de `@dimarborda/docx-editor` (o el commit) y el navegador;
- los pasos para reproducirlo y, si aplica, un `.docx` de ejemplo **sin datos personales**;
- el impacto que observas (por ejemplo: bloqueo de la pestaña, ejecución de código, lectura de datos).

Cuando el problema esté confirmado, se corrige en una nueva versión del paquete y se describe en las notas de esa versión, reconociendo a quien lo reportó si lo desea.

## Modelo de seguridad

El editor procesa documentos que pueden venir de cualquier parte, así que trata cada `.docx` como no confiable:

- **Todo ocurre en el navegador.** El documento no se envía a ningún servidor, no hay telemetría y las fuentes van incluidas en el paquete. Las únicas peticiones de red son al propio sitio: el motor `.wasm`, las fuentes y, si se usa el atributo `src`, el `.docx` indicado.
- **Sin contenido activo.** El documento se dibuja en un `<canvas>`, nunca se inserta como HTML, así que su contenido no puede ejecutar código (XSS). Las imágenes se leen solo del propio archivo: un documento con enlaces externos no provoca peticiones de red.
- **XML seguro.** El parser no resuelve entidades externas (XXE) ni definiciones DTD.
- **Límites al descomprimir.** Un `.docx` es un ZIP: el editor rechaza archivos con más de 10 000 partes, partes de más de 256 MB o un total de más de 512 MB descomprimidos, sin fiarse de los tamaños declarados en el archivo, para que una "bomba ZIP" no agote la memoria del navegador. Las imágenes insertadas desde código tienen un máximo de 15 MB.
- **Content Security Policy.** Funciona con una CSP estricta: solo necesita `'wasm-unsafe-eval'` en `script-src`, sin `'unsafe-eval'` (ver la sección *Content Security Policy* del [README del paquete](packages/docx-editor/README.md#content-security-policy)).
- **Dependencias auditadas.** Cada despliegue ejecuta `cargo audit` (base de datos RustSec) y `npm audit` sobre las dependencias del paquete, y se detiene si encuentra vulnerabilidades conocidas.

## English

Report vulnerabilities privately through **Security → Report a vulnerability** in this repository; please do not open public issues. Only the latest minor version (currently 0.7.x) receives security fixes. The editor runs entirely in the browser, never sends documents anywhere, draws them on a canvas (no HTML injection), does not resolve external XML entities, caps what a .docx may decompress to, works under a strict CSP with only `'wasm-unsafe-eval'`, and audits its Rust and npm dependencies on every deploy.
