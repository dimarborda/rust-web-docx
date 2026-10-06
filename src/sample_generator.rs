//! Demo contract generated in memory: a small but complete Word document (styles, table
//! style, numbered grid, section) so it renders the same here and in Microsoft Word.

use std::io::{Cursor, Write};
use zip::write::{SimpleFileOptions, ZipWriter};
use zip::CompressionMethod;

pub fn generate_sample_docx() -> Result<Vec<u8>, String> {
    let mut buffer = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(&mut buffer);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

    let document = document_xml();
    for (name, content) in [
        ("[Content_Types].xml", CONTENT_TYPES_XML),
        ("_rels/.rels", ROOT_RELS_XML),
        ("word/_rels/document.xml.rels", DOC_RELS_XML),
        ("word/styles.xml", STYLES_XML),
        ("word/document.xml", document.as_str()),
    ] {
        zip.start_file(name, options)
            .map_err(|e| format!("Failed to create {}: {}", name, e))?;
        zip.write_all(content.as_bytes())
            .map_err(|e| format!("Failed to write {}: {}", name, e))?;
    }

    zip.finish()
        .map_err(|e| format!("Failed to finalize sample docx zip: {}", e))?;
    Ok(buffer.into_inner())
}

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// Text column of an A4 page with 1" margins: 11906 − 2 × 1440 twips
const TEXT_WIDTH: u32 = 9026;

/// A run: `props` is the inner XML of `w:rPr` (empty for none)
fn run(props: &str, text: &str) -> String {
    let rpr = if props.is_empty() { String::new() } else { format!("<w:rPr>{}</w:rPr>", props) };
    format!(r#"<w:r>{}<w:t xml:space="preserve">{}</w:t></w:r>"#, rpr, text)
}

/// A paragraph: `props` is the inner XML of `w:pPr`
fn paragraph(props: &str, runs: &[String]) -> String {
    let ppr = if props.is_empty() { String::new() } else { format!("<w:pPr>{}</w:pPr>", props) };
    format!("<w:p>{}{}</w:p>", ppr, runs.concat())
}

/// A table cell of `width` twips holding one paragraph; cell paragraphs have no spacing
fn cell(width: u32, shading: Option<&str>, align: &str, runs: &[String]) -> String {
    let shd = shading
        .map(|fill| format!(r#"<w:shd w:val="clear" w:color="auto" w:fill="{}"/>"#, fill))
        .unwrap_or_default();
    format!(
        r#"<w:tc><w:tcPr><w:tcW w:w="{}" w:type="dxa"/>{}</w:tcPr>{}</w:tc>"#,
        width,
        shd,
        paragraph(&format!(r#"<w:spacing w:before="0" w:after="0"/><w:jc w:val="{}"/>"#, align), runs)
    )
}

fn document_xml() -> String {
    // Template variables keep their own color so they stand out while editing
    let var = |color: &str, name: &str| run(&format!(r#"<w:b/><w:color w:val="{}"/>"#, color), name);
    let text = |t: &str| run("", t);
    let bold = |t: &str| run("<w:b/>", t);

    let intro = paragraph(
        r#"<w:jc w:val="both"/>"#,
        &[
            text("En la ciudad de "),
            var("0284C7", "{{CIUDAD}}"),
            text(", a los "),
            bold("{{FECHA_CONTRATO}}"),
            text(", se celebra el presente acuerdo entre "),
            var("1E40AF", "{{NOMBRE_CLIENTE}}"),
            text(" (en adelante \"El Cliente\") y "),
            var("1E40AF", "{{NOMBRE_PROVEEDOR}}"),
            text(" (en adelante \"El Proveedor\"), para el desarrollo del proyecto "),
            var("059669", "{{NOMBRE_PROYECTO}}"),
            text("."),
        ],
    );

    // Deliverables table: a header row repeated on each page, then one row per phase
    let widths = [2100, 3326, 1600, TEXT_WIDTH - 2100 - 3326 - 1600];
    let header = ["Hito / Fase", "Descripción del Entregable", "Plazo Estimado", "Valor / Importe"];
    let aligns = ["left", "left", "center", "right"];
    let header_row = format!(
        r#"<w:tr><w:trPr><w:tblHeader/></w:trPr>{}</w:tr>"#,
        (0..4)
            .map(|i| cell(widths[i], Some("EEF2FF"), aligns[i], &[run(r#"<w:b/><w:color w:val="1E3A8A"/>"#, header[i])]))
            .collect::<String>()
    );
    let phases = [
        ("Fase 1: Arquitectura", "Diseño de base de datos y especificación técnica", "Semana 2", text("$5,000,000 COP")),
        ("Fase 2: Motor Rust WASM", "Implementación de algoritmos de procesamiento local", "Semana 6", text("$12,000,000 COP")),
        ("Fase 3: Interfaz &amp; QA", "Editor visual directo y pruebas de compatibilidad", "Semana 8", var("059669", "{{VALOR_PROYECTO}}")),
    ];
    let phase_rows: String = phases
        .iter()
        .map(|(phase, description, due, value)| {
            format!(
                "<w:tr>{}{}{}{}</w:tr>",
                cell(widths[0], None, aligns[0], &[bold(phase)]),
                cell(widths[1], None, aligns[1], &[text(description)]),
                cell(widths[2], None, aligns[2], &[text(due)]),
                cell(widths[3], None, aligns[3], &[value.clone()]),
            )
        })
        .collect();
    let grid: String = widths.iter().map(|w| format!(r#"<w:gridCol w:w="{}"/>"#, w)).collect();
    let deliverables = format!(
        r#"<w:tbl><w:tblPr><w:tblStyle w:val="TablaContrato"/><w:tblW w:w="{}" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr><w:tblGrid>{}</w:tblGrid>{}{}</w:tbl>"#,
        TEXT_WIDTH, grid, header_row, phase_rows
    );

    // Signatures: a borderless two-column table
    let half = TEXT_WIDTH / 2;
    let signer = |role: &str, variable: &str| {
        format!(
            r#"<w:tc><w:tcPr><w:tcW w:w="{half}" w:type="dxa"/></w:tcPr>{}{}{}</w:tc>"#,
            paragraph(r#"<w:spacing w:before="480" w:after="0"/><w:jc w:val="center"/>"#, &[text("______________________________")]),
            paragraph(r#"<w:spacing w:after="0"/><w:jc w:val="center"/>"#, &[bold(role)]),
            paragraph(r#"<w:spacing w:after="0"/><w:jc w:val="center"/>"#, &[run(r#"<w:color w:val="64748B"/>"#, variable)]),
        )
    };
    let signatures = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="{}" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="{half}"/><w:gridCol w:w="{half}"/></w:tblGrid><w:tr>{}{}</w:tr></w:tbl>"#,
        TEXT_WIDTH,
        signer("Por El Cliente", "{{NOMBRE_CLIENTE}}"),
        signer("Por El Proveedor", "{{NOMBRE_PROVEEDOR}}"),
    );

    let body = [
        paragraph(r#"<w:pStyle w:val="Titulo"/>"#, &[text("CONTRATO DE PRESTACIÓN DE SERVICIOS")]),
        paragraph(
            r#"<w:pStyle w:val="Subtitulo"/>"#,
            &[text("Documento generado y gestionado en tiempo real con Rust y WebAssembly")],
        ),
        intro,
        paragraph(r#"<w:pStyle w:val="Ttulo1"/>"#, &[text("Cláusula Primera: Hitos y Cronograma de Entregas")]),
        deliverables,
        paragraph(r#"<w:pStyle w:val="Ttulo1"/>"#, &[text("Cláusula Segunda: Firmas de Conformidad")]),
        signatures,
        paragraph("", &[]),
    ]
    .concat();

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr></w:body></w:document>"#
    )
}

const STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
      <w:lang w:val="es-CO" w:eastAsia="es-CO" w:bidi="ar-SA"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:qFormat/>
  </w:style>
  <w:style w:type="character" w:default="1" w:styleId="Fuentedeprrafopredeter">
    <w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="Tablanormal">
    <w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>
    <w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar>
      <w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar></w:tblPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Titulo">
    <w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:spacing w:after="60"/><w:jc w:val="center"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Subtitulo">
    <w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:spacing w:after="360"/><w:jc w:val="center"/></w:pPr>
    <w:rPr><w:i/><w:color w:val="64748B"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Ttulo1">
    <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="table" w:styleId="TablaContrato">
    <w:name w:val="Tabla Contrato"/><w:basedOn w:val="Tablanormal"/>
    <w:tblPr><w:tblBorders>
      <w:top w:val="single" w:sz="8" w:space="0" w:color="1E3A8A"/>
      <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="8" w:space="0" w:color="1E3A8A"/>
      <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
      <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
    </w:tblBorders><w:tblCellMar>
      <w:top w:w="60" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>
    </w:tblCellMar></w:tblPr>
  </w:style>
</w:styles>"#;

const CONTENT_TYPES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#;

const ROOT_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

const DOC_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx_parser::DocxModifier;

    #[test]
    fn test_sample_is_a_complete_word_document() {
        let m = DocxModifier::from_bytes(&generate_sample_docx().unwrap()).unwrap();
        assert!(m.get_file_string("word/styles.xml").unwrap().contains(r#"w:lang w:val="es-CO""#));

        let paragraphs = m.extract_paragraphs().unwrap();
        let title = &paragraphs[0];
        assert_eq!(title.style, "Titulo");
        assert_eq!((title.align.as_str(), title.runs[0].bold, title.runs[0].font_size), ("center", true, Some(20.0)));
        assert_eq!(title.runs[0].font_family.as_deref(), Some("Calibri"), "fonts come from docDefaults");

        let table = m.extract_tables().unwrap().remove(0);
        assert!(table.rich_rows[0].is_header && !table.rich_rows[1].is_header);
        assert!(table.borders.top.as_ref().is_some_and(|b| b.is_visible()), "borders from the table style");
        assert!(!table.borders.inside_v.as_ref().unwrap().is_visible());
        let cell_paragraph = &table.rich_rows[1].cells[0].paragraphs[0];
        assert_eq!((cell_paragraph.space_before, cell_paragraph.space_after), (0.0, 0.0));
        assert_eq!(table.grid_cols.iter().sum::<f64>(), TEXT_WIDTH as f64);
    }
}
