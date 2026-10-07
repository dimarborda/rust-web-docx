//! Empty Word document generated in memory: one empty paragraph, Word's default styles and
//! the chosen page size, so "new document" needs no file shipped with the app.

use std::io::{Cursor, Write};
use zip::write::{SimpleFileOptions, ZipWriter};
use zip::CompressionMethod;

use crate::sample_generator::{CONTENT_TYPES_XML, DOC_RELS_XML, ROOT_RELS_XML};

/// Paper sizes offered for new documents
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PageSize {
    A4,
    Letter,
    Legal,
}

impl PageSize {
    /// Parses "a4", "letter" or "legal" (any case); `None` means A4
    pub fn parse(name: Option<&str>) -> Result<PageSize, String> {
        match name.map(|n| n.trim().to_ascii_lowercase()).as_deref() {
            None | Some("") | Some("a4") => Ok(PageSize::A4),
            Some("letter") | Some("carta") => Ok(PageSize::Letter),
            Some("legal") | Some("oficio") => Ok(PageSize::Legal),
            Some(other) => Err(format!("Unknown page size '{}': use a4, letter or legal", other)),
        }
    }

    /// Width and height in twips (1/1440 inch)
    pub fn twips(self) -> (u32, u32) {
        match self {
            PageSize::A4 => (11906, 16838),
            PageSize::Letter => (12240, 15840),
            PageSize::Legal => (12240, 20160),
        }
    }
}

pub fn generate_blank_docx(page: PageSize) -> Result<Vec<u8>, String> {
    let mut buffer = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(&mut buffer);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

    let document = document_xml(page);
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
        .map_err(|e| format!("Failed to finalize blank docx zip: {}", e))?;
    Ok(buffer.into_inner())
}

fn document_xml(page: PageSize) -> String {
    let (w, h) = page.twips();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr><w:pgSz w:w="{w}" w:h="{h}"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr></w:body></w:document>"#
    )
}

/// Word's defaults: Calibri 11, 8 pt after each paragraph, 1.08 line spacing
const STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
      <w:lang w:val="es-CO" w:eastAsia="es-CO" w:bidi="ar-SA"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
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
</w:styles>"#;

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx_parser::DocxModifier;

    #[test]
    fn test_blank_has_one_empty_normal_paragraph() {
        let m = DocxModifier::from_bytes(&generate_blank_docx(PageSize::A4).unwrap()).unwrap();
        let paragraphs = m.extract_paragraphs().unwrap();
        assert_eq!(paragraphs.len(), 1);
        assert_eq!(paragraphs[0].text, "");
        assert_eq!(paragraphs[0].style, "Normal");
    }

    #[test]
    fn test_blank_page_sizes() {
        for (page, (w, h)) in [
            (PageSize::A4, (11906.0, 16838.0)),
            (PageSize::Letter, (12240.0, 15840.0)),
            (PageSize::Legal, (12240.0, 20160.0)),
        ] {
            let m = DocxModifier::from_bytes(&generate_blank_docx(page).unwrap()).unwrap();
            let setup = &m.layout_inputs().unwrap().page_setup;
            assert_eq!((setup.width, setup.height), (w / 15.0, h / 15.0), "{:?}", page);
        }
    }

    #[test]
    fn test_page_size_parse() {
        assert_eq!(PageSize::parse(None).unwrap(), PageSize::A4);
        assert_eq!(PageSize::parse(Some("Letter")).unwrap(), PageSize::Letter);
        assert_eq!(PageSize::parse(Some("carta")).unwrap(), PageSize::Letter);
        assert!(PageSize::parse(Some("a3")).is_err());
    }
}
