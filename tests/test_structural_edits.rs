use rust_web_docx::docx_parser::{DocumentElement, DocxModifier};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
    <w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
</w:styles>"#;

/// Builds a minimal .docx around the given body XML
fn docx(body: &str) -> DocxModifier {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>{}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>"#,
        body
    );
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let opts = SimpleFileOptions::default();
        for (name, content) in [
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#),
            ("word/document.xml", document.as_str()),
            ("word/styles.xml", STYLES),
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

fn document(m: &DocxModifier) -> String {
    m.get_file_string("word/document.xml").unwrap()
}

#[test]
fn test_enter_splits_paragraph_keeping_formatting() {
    let mut m = docx(r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>hola mundo</w:t></w:r></w:p>"#);
    assert_eq!(m.replace_range(0, 4, 0, 4, "\n").unwrap(), (1, 0));
    assert_eq!(texts(&m), vec!["hola", " mundo"]);
    for p in m.extract_paragraphs().unwrap() {
        assert_eq!(p.align, "center", "both halves keep paragraph properties");
        assert!(p.runs[0].bold && p.runs[0].color == "FF0000", "and run formatting");
    }
}

#[test]
fn test_enter_at_end_of_heading_continues_with_next_style() {
    let mut m = docx(r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Título</w:t></w:r></w:p>"#);
    m.replace_range(0, 6, 0, 6, "\n").unwrap();
    let ps = m.extract_paragraphs().unwrap();
    assert_eq!(ps[0].style, "Heading1");
    assert_eq!((ps[1].style.as_str(), ps[1].text.as_str()), ("Normal", ""));

    // Enter in the middle keeps the heading on both halves
    let mut m = docx(r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Título</w:t></w:r></w:p>"#);
    m.replace_range(0, 3, 0, 3, "\n").unwrap();
    assert!(m.extract_paragraphs().unwrap().iter().all(|p| p.style == "Heading1"));
}

#[test]
fn test_split_keeps_unique_things_unique() {
    let drawing = r#"<w:drawing><wp:inline><wp:docPr id="1" name="img"/></wp:inline></w:drawing>"#;
    let mut m = docx(&format!(
        r#"<w:p w14:paraId="1A2B3C4D" w14:textId="77777777"><w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr><w:bookmarkStart w:id="0" w:name="marca"/><w:r><w:t>antes</w:t></w:r><w:r>{}</w:r><w:r><w:t>después</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>"#,
        drawing
    ));
    m.replace_range(0, 2, 0, 2, "\n").unwrap();
    let xml = document(&m);
    assert_eq!(xml.matches("<w:drawing>").count(), 1, "the image is not duplicated");
    assert_eq!(xml.matches("<w:bookmarkStart").count(), 1, "bookmark ids stay unique");
    assert_eq!(xml.matches("w14:paraId=").count(), 1, "paragraph ids stay unique");
    assert_eq!(xml.matches("<w:type w:val=\"nextPage\"/>").count(), 1, "one section break");

    // The section break ends the section after the second half; the image follows its text
    let second_start = xml.find("<w:p>").expect("second paragraph lost its ids");
    assert!(xml[second_start..].contains("<w:sectPr>"));
    assert!(xml[second_start..].contains("<w:drawing>"));
    assert_eq!(texts(&m), vec!["an", "tesdespués"]);
}

#[test]
fn test_backspace_and_delete_join_paragraphs() {
    let body = r#"<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>uno</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>dos</w:t></w:r></w:p>"#;

    // Backspace at the start of "dos": the first paragraph's properties win
    let mut m = docx(body);
    assert_eq!(m.replace_range(0, 3, 1, 0, "").unwrap(), (0, 3));
    let ps = m.extract_paragraphs().unwrap();
    assert_eq!((ps.len(), ps[0].text.as_str(), ps[0].align.as_str()), (1, "unodos", "right"));
    // As in Word, joined text takes the surviving paragraph's style (no more heading bold)
    // but keeps its direct formatting (italic)
    let fmt: Vec<(bool, bool)> = ps[0].runs.iter().map(|r| (r.bold, r.italic)).collect();
    assert_eq!(fmt, vec![(false, false), (false, true)]);

    // Deleting an empty line in front of a heading leaves the heading
    let mut m = docx(r#"<w:p><w:pPr><w:jc w:val="right"/></w:pPr></w:p><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>dos</w:t></w:r></w:p>"#);
    m.merge_with_next(0).unwrap();
    let ps = m.extract_paragraphs().unwrap();
    assert_eq!((ps.len(), ps[0].style.as_str()), (1, "Heading1"));
}

#[test]
fn test_multi_paragraph_selection_delete_removes_what_is_between() {
    let mut m = docx(r#"<w:p><w:r><w:t>primero</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>celda</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>medio</w:t></w:r></w:p><w:p><w:r><w:t>último</w:t></w:r></w:p>"#);
    assert_eq!(m.replace_range(0, 4, 2, 2, "X").unwrap(), (0, 5));
    assert_eq!(texts(&m), vec!["primXtimo"]);
    assert!(m.extract_elements().unwrap().iter().all(|e| matches!(e, DocumentElement::Paragraph(_))), "the table inside the selection is gone");

    // Joining across a table is refused (Backspace must not eat the table)
    let mut m = docx(r#"<w:p><w:r><w:t>a</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:p><w:r><w:t>b</w:t></w:r></w:p>"#);
    assert!(m.merge_with_next(0).is_err());
}

#[test]
fn test_paste_with_newlines_creates_paragraphs() {
    let mut m = docx(r#"<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>[]</w:t></w:r></w:p>"#);
    let end = m.replace_range(0, 1, 0, 1, "uno\ndos\u{000B}bis\ntres").unwrap();
    assert_eq!(end, (2, 4));
    assert_eq!(texts(&m), vec!["[uno", "dos\nbis", "tres]"]);
    assert!(m.extract_paragraphs().unwrap().iter().all(|p| p.runs.iter().all(|r| r.italic)), "pasted text takes the caret's formatting");
}

#[test]
fn test_undo_redo_restore_document_and_selection() {
    let mut m = docx(r#"<w:p><w:r><w:t>hola</w:t></w:r></w:p>"#);
    let original = document(&m);

    // Two keystrokes recorded as one step (typing coalesces into the open step)
    m.checkpoint(Some("sel-0".into()));
    m.replace_range(0, 4, 0, 4, "!").unwrap();
    m.replace_range(0, 5, 0, 5, "!").unwrap();
    m.set_selection_after(Some("sel-2".into()));
    let typed = document(&m);

    m.checkpoint(Some("sel-2".into()));
    m.replace_range(0, 6, 0, 6, "\n").unwrap();
    m.set_selection_after(Some("sel-enter".into()));
    assert_eq!(texts(&m).len(), 2);

    assert_eq!(m.undo(), Some(Some("sel-2".into())));
    assert_eq!(document(&m), typed);
    assert_eq!(m.undo(), Some(Some("sel-0".into())));
    assert_eq!(document(&m), original);
    assert_eq!(m.undo(), None, "nothing left to undo");

    assert_eq!(m.redo(), Some(Some("sel-2".into())));
    assert_eq!(texts(&m), vec!["hola!!"]);

    // A new change drops the redo history
    m.checkpoint(None);
    m.replace_range(0, 0, 0, 0, "¡").unwrap();
    assert!(!m.can_redo());
    assert!(m.can_undo());
}

#[test]
fn test_structural_edits_export_valid_documents() {
    let bytes = rust_web_docx::sample_generator::generate_sample_docx().unwrap();
    let mut m = DocxModifier::from_bytes(&bytes).unwrap();
    let before = m.extract_paragraphs().unwrap().len();
    m.replace_range(1, 3, 1, 3, "\nnuevo\n").unwrap();
    m.replace_range(0, 0, 0, 0, "inicio\n").unwrap();
    let last = m.extract_paragraphs().unwrap().len() - 1;
    m.replace_range(1, 2, last, 1, "").unwrap();

    let reloaded = DocxModifier::from_bytes(&m.to_bytes().unwrap()).unwrap();
    assert_eq!(texts(&reloaded), texts(&m));
    assert!(texts(&m).len() < before + 3);
}
