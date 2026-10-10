use rust_web_docx::blank_generator::{generate_blank_docx, PageSize};
use rust_web_docx::docx_parser::{DocumentElement, DocxModifier, NewParagraph};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

fn blank() -> DocxModifier {
    DocxModifier::from_bytes(&generate_blank_docx(PageSize::Letter).unwrap()).unwrap()
}

/// .docx with the given body and styles (as Word writes them in Spanish: "Ttulo1" named "heading 1")
fn docx_with_styles(body: &str, styles: &str) -> DocxModifier {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#,
        body
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>{}</w:styles>"#,
        styles
    );
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        for (name, content) in [
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#),
            ("word/document.xml", document.as_str()),
            ("word/styles.xml", styles.as_str()),
        ] {
            zip.start_file(name, SimpleFileOptions::default()).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
    }
    DocxModifier::from_bytes(&buf.into_inner()).unwrap()
}

fn p(text: &str, style: Option<&str>) -> NewParagraph {
    NewParagraph { text: text.into(), style: style.map(String::from), ..Default::default() }
}

fn styles_xml(m: &DocxModifier) -> String {
    m.get_file_string("word/styles.xml").unwrap()
}

/// (text, outline level) of the body paragraphs
fn outline(m: &DocxModifier) -> Vec<(String, Option<u32>)> {
    m.extract_paragraphs().unwrap().into_iter().map(|p| (p.text, p.outline_level)).collect()
}

#[test]
fn test_headings_become_real_word_styles_with_outline_levels() {
    let mut m = blank();
    let items = [
        p("Informe de investigación", Some("Title")),
        p("1. Introducción", Some("Heading1")),
        p("Texto de la introducción.", None),
        p("1.1 Antecedentes", Some("Heading2")),
        p("1.1.1 En Colombia", Some("Heading3")),
        p("2. Metodología", Some("Heading1")),
    ];
    m.insert_paragraphs(0, 0, &items).unwrap();

    let got = outline(&m);
    assert_eq!(got[0], ("Informe de investigación".into(), None), "the title is not a section");
    assert_eq!(got[1], ("1. Introducción".into(), Some(0)));
    assert_eq!(got[2], ("Texto de la introducción.".into(), None));
    assert_eq!(got[3].1, Some(1));
    assert_eq!(got[4].1, Some(2));
    assert_eq!(got[5].1, Some(0));

    let styles = styles_xml(&m);
    for id in ["Title", "Heading1", "Heading2", "Heading3"] {
        assert_eq!(styles.matches(&format!(r#"w:styleId="{}""#, id)).count(), 1, "{} defined once", id);
    }
    assert!(!styles.contains("2E74B5") && styles.contains("262626"), "gray headings, not Word's blue");

    // Inserting more headings reuses the styles
    let end = m.extract_paragraphs().unwrap().len() - 1;
    m.insert_paragraphs(end, 14, &[p("3. Resultados", Some("Heading1"))]).unwrap();
    assert_eq!(styles_xml(&m).matches(r#"w:styleId="Heading1""#).count(), 1);

    // The saved file keeps them
    let reopened = DocxModifier::from_bytes(&m.to_bytes().unwrap()).unwrap();
    assert_eq!(outline(&reopened).last().unwrap(), &("3. Resultados".to_string(), Some(0)));
}

#[test]
fn test_existing_word_heading_styles_are_reused() {
    let mut m = docx_with_styles(
        r#"<w:p><w:r><w:t>hola</w:t></w:r></w:p>"#,
        r#"<w:style w:type="paragraph" w:styleId="Ttulo1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>"#,
    );
    m.insert_paragraphs(0, 4, &[p("Capítulo", Some("Heading1"))]).unwrap();
    assert!(m.get_file_string("word/document.xml").unwrap().contains(r#"<w:pStyle w:val="Ttulo1"/>"#));
    assert!(!styles_xml(&m).contains(r#"w:styleId="Heading1""#), "no duplicate style");
    assert_eq!(outline(&m)[1], ("Capítulo".into(), Some(0)));
    assert!(m.insert_paragraphs(0, 0, &[p("x", Some("Heading7"))]).unwrap_err().contains("Estilo no válido"));
}

#[test]
fn test_replace_paragraphs_with_a_table_inside_is_one_undo_step() {
    let table = r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>celda</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
    let body = format!(
        r#"<w:p><w:r><w:t>Antes</w:t></w:r></w:p><w:p><w:r><w:t>Sección vieja</w:t></w:r></w:p>{}<w:p><w:r><w:t>Cierre viejo</w:t></w:r></w:p><w:p><w:r><w:t>Después</w:t></w:r></w:p>"#,
        table
    );
    let mut m = docx_with_styles(&body, "");
    let original = m.get_file_string("word/document.xml").unwrap();

    // Paragraphs: 0 Antes, 1 Sección vieja, 2 celda, 3 Cierre viejo, 4 Después
    m.checkpoint(Some("sel".into()));
    let (first, caret) = m.replace_paragraphs(1, 3, &[p("Sección nueva", Some("Heading1")), p("Contenido nuevo.", None)]).unwrap();
    assert_eq!((first, caret), (1, (2, "Contenido nuevo.".chars().count())));
    let texts: Vec<String> = outline(&m).into_iter().map(|(t, _)| t).collect();
    assert_eq!(texts, vec!["Antes", "Sección nueva", "Contenido nuevo.", "Después"]);
    assert!(!m.extract_elements().unwrap().iter().any(|e| matches!(e, DocumentElement::Table(_))), "the table inside the range is gone");

    assert_eq!(m.undo(), Some(Some("sel".into())));
    assert_eq!(m.get_file_string("word/document.xml").unwrap(), original, "one step restores everything");
    // Undo also removed the style it had added; inserting again re-creates it
    assert!(!styles_xml(&m).contains(r#"w:styleId="Heading1""#));
    m.insert_paragraphs(0, 0, &[p("Otra vez", Some("Heading1"))]).unwrap();
    assert!(styles_xml(&m).contains(r#"w:styleId="Heading1""#));
}

#[test]
fn test_replace_paragraphs_refuses_unsafe_ranges() {
    let table = r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>celda</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
    let body = format!(
        r#"<w:p><w:r><w:t>Uno</w:t></w:r></w:p>{}<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>Fin de sección</w:t></w:r></w:p><w:p><w:r><w:t>Tres</w:t></w:r></w:p>"#,
        table
    );
    let mut m = docx_with_styles(&body, "");
    let original = m.get_file_string("word/document.xml").unwrap();
    // 0 Uno, 1 celda, 2 Fin de sección, 3 Tres
    assert!(m.replace_paragraphs(0, 1, &[p("x", None)]).unwrap_err().contains("corta una tabla"));
    assert!(m.replace_paragraphs(2, 3, &[p("x", None)]).unwrap_err().contains("salto de sección"));
    assert!(m.replace_paragraphs(0, 9, &[p("x", None)]).is_err());
    assert!(m.replace_paragraphs(0, 0, &[]).is_err());
    assert_eq!(m.get_file_string("word/document.xml").unwrap(), original, "nothing changes on error");
    // Inside one cell is fine
    m.replace_paragraphs(1, 1, &[p("celda nueva", None)]).unwrap();
    assert!(m.get_file_string("word/document.xml").unwrap().contains("celda nueva</w:t></w:r></w:p></w:tc>"));
}

#[test]
fn test_text_inserted_before_a_heading_is_not_a_heading() {
    let mut m = blank();
    m.insert_paragraphs(0, 0, &[p("1. Introducción", Some("Heading1")), p("Texto.", None), p("2. Métodos", Some("Heading1"))]).unwrap();
    // Before "2. Métodos" (paragraph 2, offset 0): the split keeps the heading's properties
    m.insert_paragraphs(2, 0, &[p("Cierre de la introducción.", None), p("1.2 Alcance", Some("Heading2"))]).unwrap();
    let got = outline(&m);
    assert_eq!(got[2], ("Cierre de la introducción.".into(), None), "{:?}", got);
    assert_eq!(got[3], ("1.2 Alcance".into(), Some(1)));
    assert_eq!(got[4], ("2. Métodos".into(), Some(0)));
}

