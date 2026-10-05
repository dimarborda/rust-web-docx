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
  <w:background w:color="FFFFFF"/>
  <w:body>
    <!-- Document Title Centered with Deep Blue Color -->
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:sz w:val="48"/>
          <w:color w:val="1E3A8A"/>
        </w:rPr>
        <w:t>CONTRATO DE PRESTACIÓN DE SERVICIOS</w:t>
      </w:r>
    </w:p>
    
    <!-- Subtitle Centered with Italic Accent -->
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:i/>
          <w:color w:val="64748B"/>
        </w:rPr>
        <w:t>Documento generado y gestionado en tiempo real con Rust y WebAssembly</w:t>
      </w:r>
    </w:p>
    
    <!-- Introductory Paragraph with Template Variables -->
    <w:p>
      <w:pPr>
        <w:jc w:val="both"/>
      </w:pPr>
      <w:r>
        <w:t>En la ciudad de </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:color w:val="0284C7"/>
        </w:rPr>
        <w:t>{{CIUDAD}}</w:t>
      </w:r>
      <w:r>
        <w:t>, a los </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>{{FECHA_CONTRATO}}</w:t>
      </w:r>
      <w:r>
        <w:t>, se celebra el presente acuerdo entre </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:color w:val="1E40AF"/>
        </w:rPr>
        <w:t>{{NOMBRE_CLIENTE}}</w:t>
      </w:r>
      <w:r>
        <w:t> (en adelante "El Cliente") y </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:color w:val="1E40AF"/>
        </w:rPr>
        <w:t>{{NOMBRE_PROVEEDOR}}</w:t>
      </w:r>
      <w:r>
        <w:t> (en adelante "El Proveedor"), para el desarrollo del proyecto </w:t>
      </w:r>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:color w:val="059669"/>
        </w:rPr>
        <w:t>{{NOMBRE_PROYECTO}}</w:t>
      </w:r>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>

    <!-- Heading 2 -->
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading2"/>
        <w:jc w:val="left"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:sz w:val="32"/>
          <w:color w:val="1E3A8A"/>
        </w:rPr>
        <w:t>Cláusula Primera: Hitos y Cronograma de Entregas</w:t>
      </w:r>
    </w:p>

    <!-- Table of Deliverables -->
    <w:tbl>
      <w:tblPr>
        <w:tblBorders>
          <w:top w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>
          <w:left w:val="none"/>
          <w:bottom w:val="single" w:sz="8" w:space="0" w:color="94A3B8"/>
          <w:right w:val="none"/>
          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>
          <w:insideV w:val="none"/>
        </w:tblBorders>
      </w:tblPr>
      
      <!-- Table Header Row -->
      <w:tr>
        <w:tc>
          <w:p>
            <w:pPr><w:jc w:val="left"/></w:pPr>
            <w:r>
              <w:rPr><w:b/><w:color w:val="1E3A8A"/></w:rPr>
              <w:t>Hito / Fase</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:pPr><w:jc w:val="left"/></w:pPr>
            <w:r>
              <w:rPr><w:b/><w:color w:val="1E3A8A"/></w:rPr>
              <w:t>Descripción del Entregable</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r>
              <w:rPr><w:b/><w:color w:val="1E3A8A"/></w:rPr>
              <w:t>Plazo Estimado</w:t>
            </w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:pPr><w:jc w:val="right"/></w:pPr>
            <w:r>
              <w:rPr><w:b/><w:color w:val="1E3A8A"/></w:rPr>
              <w:t>Valor / Importe</w:t>
            </w:r>
          </w:p>
        </w:tc>
      </w:tr>

      <!-- Row 1 -->
      <w:tr>
        <w:tc>
          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Fase 1: Arquitectura</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>Diseño de base de datos y especificación técnica</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Semana 2</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>$5,000,000 COP</w:t></w:r></w:p>
        </w:tc>
      </w:tr>

      <!-- Row 2 -->
      <w:tr>
        <w:tc>
          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Fase 2: Motor Rust WASM</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>Implementación de algoritmos de procesamiento local</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Semana 6</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>$12,000,000 COP</w:t></w:r></w:p>
        </w:tc>
      </w:tr>

      <!-- Row 3 -->
      <w:tr>
        <w:tc>
          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Fase 3: Interfaz &amp; QA</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>Editor visual directo y pruebas de compatibilidad</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Semana 8</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="059669"/></w:rPr><w:t>{{VALOR_PROYECTO}}</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>

    <!-- Heading 2 -->
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading2"/>
        <w:jc w:val="left"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:sz w:val="32"/>
          <w:color w:val="1E3A8A"/>
        </w:rPr>
        <w:t>Cláusula Segunda: Firmas de Conformidad</w:t>
      </w:r>
    </w:p>

    <!-- Signatures Table -->
    <w:tbl>
      <w:tblPr>
        <w:tblBorders>
          <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/>
          <w:insideH w:val="none"/><w:insideV w:val="none"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tr>
        <w:tc>
          <w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:t>__________________________________</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:b/></w:rPr><w:t>Por El Cliente</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:color w:val="64748B"/></w:rPr><w:t>{{NOMBRE_CLIENTE}}</w:t></w:r>
          </w:p>
        </w:tc>
        <w:tc>
          <w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:t>__________________________________</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:b/></w:rPr><w:t>Por El Proveedor</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:color w:val="64748B"/></w:rPr><w:t>{{NOMBRE_PROVEEDOR}}</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>

    <w:sectPr/>
  </w:body>
</w:document>"#;
