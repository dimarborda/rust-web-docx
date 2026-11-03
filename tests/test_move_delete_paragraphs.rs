use rust_web_docx::docx_parser::DocxModifier;
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

fn docx(body: &str) -> DocxModifier {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#,
        body
    );
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        for (name, content) in [
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#),
            ("word/document.xml", document.as_str()),
        ] {
            zip.start_file(name, SimpleFileOptions::default()).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
    }
    DocxModifier::from_bytes(&buf.into_inner()).unwrap()
}

fn para(text: &str) -> String {
    format!("<w:p><w:r><w:t>{}</w:t></w:r></w:p>", text)
}

fn texts(m: &DocxModifier) -> Vec<String> {
    m.extract_paragraphs().unwrap().into_iter().map(|p| p.text).collect()
}

fn xml(m: &DocxModifier) -> String {
    m.get_file_string("word/document.xml").unwrap()
}

const SOURCES: &str = r#"<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Fuentes: </w:t></w:r><w:hyperlink r:id="rId9"><w:r><w:t>tienda-a.com</w:t></w:r></w:hyperlink></w:p>"#;

#[test]
fn test_move_keeps_the_paragraph_xml_and_goes_to_the_end() {
    let mut m = docx(&format!("{}{}{}{}", para("Título"), SOURCES, para("Conclusiones"), para("Recomendación")));
    let first = m.move_paragraphs(1, 1, 4).unwrap();
    assert_eq!(first, 3);
    assert_eq!(texts(&m), ["Título", "Conclusiones", "Recomendación", "Fuentes: tienda-a.com"]);
    // Tal cual: enlace, negrita y alineación
    assert!(xml(&m).contains(SOURCES));
}

#[test]
fn test_move_forward_and_backward() {
    let body: String = ["A", "B", "C", "D", "E"].iter().map(|t| para(t)).collect();
    let mut m = docx(&body);
    assert_eq!(m.move_paragraphs(3, 4, 1).unwrap(), 1);
    assert_eq!(texts(&m), ["A", "D", "E", "B", "C"]);
    let mut m = docx(&body);
    assert_eq!(m.move_paragraphs(0, 1, 4).unwrap(), 2);
    assert_eq!(texts(&m), ["C", "D", "A", "B", "E"]);
    // Al mismo lugar: no cambia nada
    let mut m = docx(&body);
    assert_eq!(m.move_paragraphs(1, 2, 3).unwrap(), 1);
    assert_eq!(texts(&m), ["A", "B", "C", "D", "E"]);
}

#[test]
fn test_move_refuses_a_destination_inside_the_range_or_a_table() {
    let table = format!("<w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>", para("celda"));
    let mut m = docx(&format!("{}{}{}{}", para("A"), para("B"), table, para("C")));
    assert!(m.move_paragraphs(0, 1, 1).unwrap_err().contains("dentro del rango"));
    assert!(m.move_paragraphs(0, 0, 2).unwrap_err().contains("dentro de una tabla"));
    assert!(m.move_paragraphs(0, 0, 9).is_err());
    // (extract_paragraphs no lista las celdas)
    assert_eq!(texts(&m), ["A", "B", "C"]);
}

#[test]
fn test_a_whole_table_moves_with_its_range() {
    let table = format!("<w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>", para("celda"));
    let mut m = docx(&format!("{}{}{}{}", para("A"), table, para("B"), para("C")));
    // El rango A + tabla (párrafos 0 y 1, el último en la celda) va al final con la tabla entera
    assert_eq!(m.move_paragraphs(0, 1, 4).unwrap(), 2);
    assert!(xml(&m).contains(&format!("{}<w:tbl>", para("A"))));
    assert_eq!(texts(&m), ["B", "C", "A", ""]);
    // Word necesita un párrafo después de una tabla final
    let body = xml(&m);
    let end = body.rfind("<w:sectPr").unwrap();
    assert!(body[..end].ends_with("<w:p/>"), "{}", &body[..end]);
}

#[test]
fn test_delete_removes_the_range_in_one_go() {
    let body: String = ["A", "B", "C", "D"].iter().map(|t| para(t)).collect();
    let mut m = docx(&body);
    assert_eq!(m.delete_paragraphs(1, 2).unwrap(), 2);
    assert_eq!(texts(&m), ["A", "D"]);
}

#[test]
fn test_delete_keeps_one_paragraph_and_refuses_bad_ranges() {
    let mut m = docx(&format!("{}{}", para("A"), para("B")));
    m.delete_paragraphs(0, 1).unwrap();
    assert_eq!(texts(&m), [""]);

    let table = format!("<w:tbl><w:tr><w:tc>{}{}</w:tc></w:tr></w:tbl>", para("c1"), para("c2"));
    let mut m = docx(&format!("{}{}{}", para("A"), table, para("B")));
    assert!(m.delete_paragraphs(0, 1).unwrap_err().contains("corta una tabla"));
    assert!(m.delete_paragraphs(3, 9).is_err());
    let with_break = format!("{}<w:p><w:pPr><w:sectPr/></w:pPr></w:p>{}", para("A"), para("B"));
    let mut m = docx(&with_break);
    assert!(m.delete_paragraphs(0, 1).unwrap_err().contains("salto de sección"));
    assert_eq!(texts(&m).len(), 3);
}
