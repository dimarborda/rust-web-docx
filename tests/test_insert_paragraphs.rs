use rust_web_docx::blank_generator::{generate_blank_docx, PageSize};
use rust_web_docx::docx_parser::{DocxModifier, NewParagraph};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

/// Minimal .docx around the given body XML
fn docx(body: &str) -> DocxModifier {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>"#,
        body
    );
    let styles = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"#;
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let opts = SimpleFileOptions::default();
        for (name, content) in [
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#),
            ("word/document.xml", document.as_str()),
            ("word/styles.xml", styles),
        ] {
            zip.start_file(name, opts).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
    }
    DocxModifier::from_bytes(&buf.into_inner()).unwrap()
}

fn texts(m: &DocxModifier) -> Vec<String> {
    m.extract_paragraphs().unwrap().into_iter().map(|p| p.text).collect()
}

fn para(text: &str) -> NewParagraph {
    NewParagraph { text: text.to_string(), ..Default::default() }
}

#[test]
fn test_insert_into_blank_document_with_heading_format() {
    let mut m = DocxModifier::from_bytes(&generate_blank_docx(PageSize::Letter).unwrap()).unwrap();
    let title = NewParagraph {
        text: "CONTRATO DE SERVICIOS".into(),
        bold: Some(true),
        font_size: Some(16.0),
        align: Some("center".into()),
        ..Default::default()
    };
    let (first, caret) = m.insert_paragraphs(0, 0, &[title, para("Entre las partes."), para("Primera cláusula.")]).unwrap();

    assert_eq!(first, 0);
    assert_eq!(caret, (2, "Primera cláusula.".chars().count()));
    assert_eq!(texts(&m), vec!["CONTRATO DE SERVICIOS", "Entre las partes.", "Primera cláusula."]);

    let paragraphs = m.extract_paragraphs().unwrap();
    let heading = &paragraphs[0];
    assert_eq!(heading.align, "center");
    assert!(heading.runs.iter().all(|r| r.bold && r.font_size == Some(16.0)), "{:?}", heading.runs);
    let body = &paragraphs[1];
    assert!(body.runs.iter().all(|r| !r.bold), "unformatted paragraphs keep the document defaults");
    assert_ne!(body.align, "center");
}

#[test]
fn test_insert_in_the_middle_keeps_existing_text_in_its_own_paragraphs() {
    let mut m = docx(r#"<w:p><w:r><w:t>antes después</w:t></w:r></w:p>"#);
    let (first, caret) = m.insert_paragraphs(0, 6, &[para("uno"), para("dos")]).unwrap();

    assert_eq!(texts(&m), vec!["antes ", "uno", "dos", "después"]);
    assert_eq!(first, 1);
    assert_eq!(caret, (2, 3));
}

#[test]
fn test_insert_at_paragraph_end_and_start() {
    let mut m = docx(r#"<w:p><w:r><w:t>hola</w:t></w:r></w:p>"#);
    let (first, _) = m.insert_paragraphs(0, 4, &[para("al final")]).unwrap();
    assert_eq!((first, texts(&m)), (1, vec!["hola".to_string(), "al final".to_string()]));

    let (first, _) = m.insert_paragraphs(0, 0, &[para("al inicio")]).unwrap();
    assert_eq!(first, 0);
    assert_eq!(texts(&m), vec!["al inicio", "hola", "al final"]);
}

#[test]
fn test_line_breaks_stay_inside_one_paragraph() {
    let mut m = docx(r#"<w:p/>"#);
    m.insert_paragraphs(0, 0, &[para("Calle 1\nBogotá"), para("Otro")]).unwrap();
    let t = texts(&m);
    assert_eq!(t.len(), 2, "{:?}", t);
    assert!(t[0].contains("Calle 1") && t[0].contains("Bogotá"));
}

#[test]
fn test_insert_is_one_undo_step() {
    let mut m = docx(r#"<w:p><w:r><w:t>hola</w:t></w:r></w:p>"#);
    let original = texts(&m);

    m.checkpoint(Some("before".into()));
    let heading = NewParagraph { text: "Título".into(), bold: Some(true), align: Some("center".into()), ..Default::default() };
    m.insert_paragraphs(0, 4, &[heading, para("texto")]).unwrap();
    m.set_selection_after(Some("after".into()));
    assert_eq!(texts(&m).len(), 3);

    assert_eq!(m.undo(), Some(Some("before".into())));
    assert_eq!(texts(&m), original, "text and formatting are undone together");
    assert_eq!(m.undo(), None);
}

#[test]
fn test_insert_rejects_nothing_and_bad_positions() {
    let mut m = docx(r#"<w:p><w:r><w:t>hola</w:t></w:r></w:p>"#);
    assert!(m.insert_paragraphs(0, 0, &[]).is_err());
    assert!(m.insert_paragraphs(9, 0, &[para("x")]).is_err());
    // An offset past the end is clamped to the end of the paragraph
    m.insert_paragraphs(0, 99, &[para("x")]).unwrap();
    assert_eq!(texts(&m), vec!["hola", "x"]);
}
