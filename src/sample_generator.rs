use std::io::{Cursor, Write};
use zip::write::{SimpleFileOptions, ZipWriter};
use zip::CompressionMethod;

pub fn generate_sample_docx() -> Result<Vec<u8>, String> {
    let mut buffer = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(&mut buffer);

    let options = SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated);

    // 1. [Content_Types].xml
    zip.start_file("[Content_Types].xml", options)
        .map_err(|e| format!("Failed to create [Content_Types].xml: {}", e))?;
    zip.write_all(CONTENT_TYPES_XML.as_bytes())
        .map_err(|e| format!("Failed to write [Content_Types].xml: {}", e))?;

    // 2. _rels/.rels
    zip.start_file("_rels/.rels", options)
        .map_err(|e| format!("Failed to create _rels/.rels: {}", e))?;
    zip.write_all(ROOT_RELS_XML.as_bytes())
        .map_err(|e| format!("Failed to write _rels/.rels: {}", e))?;

    // 3. word/_rels/document.xml.rels
    zip.start_file("word/_rels/document.xml.rels", options)
        .map_err(|e| format!("Failed to create word/_rels/document.xml.rels: {}", e))?;
    zip.write_all(DOC_RELS_XML.as_bytes())
        .map_err(|e| format!("Failed to write word/_rels/document.xml.rels: {}", e))?;

    // 4. word/document.xml
    zip.start_file("word/document.xml", options)
        .map_err(|e| format!("Failed to create word/document.xml: {}", e))?;
    zip.write_all(SAMPLE_DOCUMENT_XML.as_bytes())
        .map_err(|e| format!("Failed to write word/document.xml: {}", e))?;

    zip.finish()
        .map_err(|e| format!("Failed to finalize sample docx zip: {}", e))?;

    Ok(buffer.into_inner())
}

const CONTENT_TYPES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#;

const ROOT_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

const DOC_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#;

const SAMPLE_DOCUMENT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:sz w:val="48"/>
          <w:color w:val="2B579A"/>
        </w:rPr>
        <w:t>Acuerdo de Servicios Profesionales</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr>
          <w:i/>
        </w:rPr>
        <w:t>Documento de prueba generado automáticamente con Rust y WebAssembly.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Este contrato se celebra el día </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>{{FECHA_CONTRATO}}</w:t>
      </w:r>
      <w:r>
        <w:t> entre </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>{{NOMBRE_CLIENTE}}</w:t>
      </w:r>
      <w:r>
        <w:t> (en adelante "El Cliente") y </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>{{NOMBRE_PROVEEDOR}}</w:t>
      </w:r>
      <w:r>
        <w:t> (en adelante "El Proveedor").</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading2"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:sz w:val="32"/>
        </w:rPr>
        <w:t>Cláusula Primera: Objeto del Contrato</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>El Proveedor se compromete a prestar servicios de desarrollo de software para el proyecto </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>{{NOMBRE_PROYECTO}}</w:t>
      </w:r>
      <w:r>
        <w:t>, con un importe acordado de </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>{{VALOR_PROYECTO}}</w:t>
      </w:r>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading2"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:sz w:val="32"/>
        </w:rPr>
        <w:t>Cláusula Segunda: Plazos y Entregas</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Las entregas se realizarán en la ciudad de </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>{{CIUDAD}}</w:t>
      </w:r>
      <w:r>
        <w:t> antes del plazo fijado.</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Firma del Cliente: _________________________</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:t>Firma del Proveedor: _______________________</w:t>
      </w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>"#;
