# Documentos de prueba

Coloca aquí los archivos `.docx` con los que quieras probar el editor. La carpeta está
ignorada por git (salvo este README), así que cada quien usa sus propios documentos y
ninguno —contratos firmados, datos personales— termina en el repositorio.

- **App web:** en `npm run dev`, cada `.docx` de esta carpeta aparece como tarjeta en la
  pantalla de inicio; al agregar o quitar archivos basta con recargar la página. También
  puedes usar **"Elegir carpeta…"** para ver los `.docx` de cualquier carpeta de tu equipo
  (funciona en producción y nada se sube: los archivos se leen en el navegador).
- **Tests de Rust:** algunos tests usan documentos reales por nombre. Si el archivo no
  está, el test se omite con un aviso en lugar de fallar. Para ejecutarlos completos se
  usan estos nombres:
  - `plantilla-ejemplo-vead.docx` — plantilla de certificado con variables `{name}`, `{email}`…
  - `costo-eficiencia-modelos-llm-rag.docx` — informe con tablas y listas con viñetas
  - `CONTRATO DE PROMESA DE COMPRAVENTA DE BIEN INMUEBLE_firma_….docx` — contrato con listas numeradas
