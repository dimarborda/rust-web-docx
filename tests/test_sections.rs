//! Logical sections (block-level content controls tagged `section:<key>`): listing, defining,
//! replacing and unwrapping them, and that the saved file keeps them.

use rust_web_docx::docx_parser::{DocumentElement, DocxModifier, NewParagraph, SectionInfo};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

const HEADING_STYLE: &str = r#"<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>"#;

fn docx(body: &str) -> DocxModifier {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#,
        body
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>{}</w:styles>"#,
        HEADING_STYLE
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

fn para(text: &str) -> String {
    format!("<w:p><w:r><w:t>{}</w:t></w:r></w:p>", text)
}

fn heading(text: &str) -> String {
    format!(r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>{}</w:t></w:r></w:p>"#, text)
}

fn table(cells: &[&str]) -> String {
    let tcs: String = cells.iter().map(|c| format!("<w:tc>{}</w:tc>", para(c))).collect();
    format!(r#"<w:tbl><w:tblGrid>{}</w:tblGrid><w:tr>{}</w:tr></w:tbl>"#, "<w:gridCol/>".repeat(cells.len()), tcs)
}

fn section(key: &str, alias: &str, content: &str) -> String {
    format!(
        r#"<w:sdt><w:sdtPr><w:alias w:val="{}"/><w:tag w:val="section:{}"/><w:id w:val="1"/></w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>"#,
        alias, key, content
    )
}

/// Text of every body paragraph by index, table cells included
fn texts(m: &DocxModifier) -> Vec<String> {
    let mut out: Vec<(usize, String)> = Vec::new();
    for el in m.extract_elements().unwrap() {
        match el {
            DocumentElement::Paragraph(p) => out.push((p.index, p.text)),
            DocumentElement::Table(t) => {
                for cell in t.rich_rows.iter().flat_map(|r| r.cells.iter()) {
                    out.extend(cell.paragraphs.iter().map(|p| (p.index, p.text.clone())));
                }
            }
        }
    }
    out.sort_by_key(|(i, _)| *i);
    out.into_iter().map(|(_, t)| t).collect()
}

fn reopen(m: &DocxModifier) -> DocxModifier {
    DocxModifier::from_bytes(&m.to_bytes().unwrap()).unwrap()
}

fn np(text: &str) -> NewParagraph {
    NewParagraph { text: text.into(), ..Default::default() }
}

/// Technical report: Introducción (with a nested Alcance), Caracterización with a table,
/// a paragraph outside any section and an untagged content control
fn report() -> DocxModifier {
    let body = [
        para("Informe técnico"),
        section(
            "introduccion",
            "Introducción",
            &[heading("1. Introducción"), para("Objeto del informe."), section("alcance", "Alcance", &para("Solo el edificio A."))].concat(),
        ),
        section("caracterizacion", "Caracterización &amp; datos", &[heading("2. Caracterización"), table(&["Muestra", "pH"]), para("Notas.")].concat()),
        para("Fuera de toda sección"),
        r#"<w:sdt><w:sdtPr><w:tag w:val="cliente"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>ACME</w:t></w:r></w:p></w:sdtContent></w:sdt>"#.to_string(),
    ]
    .concat();
    docx(&body)
}

#[test]
fn test_sections_lists_tagged_controls_with_ranges_and_nesting() {
    let m = report();
    assert_eq!(
        texts(&m),
        ["Informe técnico", "1. Introducción", "Objeto del informe.", "Solo el edificio A.", "2. Caracterización", "Muestra", "pH", "Notas.", "Fuera de toda sección", "ACME"]
    );
    let sections = m.sections().unwrap();
    let summary: Vec<(&str, &str, Option<usize>, Option<usize>, usize, Option<&str>)> = sections
        .iter()
        .map(|s| (s.key.as_str(), s.title.as_str(), s.first, s.last, s.depth, s.parent.as_deref()))
        .collect();
    assert_eq!(
        summary,
        [
            ("introduccion", "Introducción", Some(1), Some(3), 0, None),
            ("alcance", "Alcance", Some(3), Some(3), 1, Some("introduccion")),
            ("caracterizacion", "Caracterización & datos", Some(4), Some(7), 0, None),
        ],
        "untagged content controls are not sections"
    );
}

#[test]
fn test_define_section_wraps_blocks_and_survives_saving() {
    let mut m = docx(&[heading("Conclusiones"), para("Primera."), table(&["A", "B"]), para("Última."), para("Anexo")].concat());
    let info = m.define_section("conclusiones", 0, 4, Some("Conclusiones"), true).unwrap();
    assert_eq!(
        info,
        SectionInfo { key: "conclusiones".into(), title: "Conclusiones".into(), first: Some(0), last: Some(4), depth: 0, parent: None, locked: true }
    );
    let xml = m.get_file_string("word/document.xml").unwrap();
    assert!(xml.contains(r#"<w:tag w:val="section:conclusiones"/>"#));
    assert!(xml.contains(r#"<w:lock w:val="sdtLocked"/>"#));
    assert!(xml.contains("<w:sectPr><w:pgSz"), "the body's final sectPr stays outside");
    // Text and numbering do not change; the saved file reads the section back
    assert_eq!(texts(&m), ["Conclusiones", "Primera.", "A", "B", "Última.", "Anexo"]);
    assert_eq!(reopen(&m).sections().unwrap(), vec![info]);
}

#[test]
fn test_define_section_takes_whole_tables_and_nests() {
    let mut m = report();
    // From the table's first cell to its last one: the table is wrapped whole
    let datos = m.define_section("datos", 5, 6, None, false).unwrap();
    assert_eq!((datos.first, datos.last, datos.title.as_str()), (Some(5), Some(6), "datos"));
    assert_eq!((datos.depth, datos.parent.as_deref()), (1, Some("caracterizacion")));
    assert!(m.get_file_string("word/document.xml").unwrap().contains("<w:sdtContent><w:tbl>"));
}

#[test]
fn test_define_section_refusals_leave_the_document_untouched() {
    let mut m = report();
    let before = m.get_file_string("word/document.xml").unwrap();
    let err = |r: Result<SectionInfo, String>| r.unwrap_err();
    assert!(err(m.define_section("introduccion", 8, 8, None, false)).contains("Ya existe"));
    assert!(err(m.define_section("mal clave", 8, 8, None, false)).contains("no válida"));
    assert!(err(m.define_section("", 8, 8, None, false)).contains("no válida"));
    // Cuts the Introducción section (starts outside, ends inside)
    assert!(err(m.define_section("x", 0, 2, None, false)).contains("corta"));
    // Starts in the middle of a table
    assert!(err(m.define_section("x", 6, 7, None, false)).contains("tabla"));
    assert!(err(m.define_section("x", 50, 51, None, false)).contains("no existe"));
    assert_eq!(m.get_file_string("word/document.xml").unwrap(), before);
}

#[test]
fn test_define_section_refuses_a_word_section_break() {
    let mut m = docx(&[para("A"), r#"<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>B</w:t></w:r></w:p>"#.to_string(), para("C")].concat());
    assert!(m.define_section("x", 0, 2, None, false).unwrap_err().contains("salto de sección"));
    assert!(m.define_section("x", 2, 2, None, false).is_ok());
}

#[test]
fn test_replace_section_keeps_the_marker_and_optionally_the_heading() {
    let mut m = report();
    m.replace_section("introduccion", &[np("Nuevo objeto."), np("Segundo párrafo.")], true).unwrap();
    assert_eq!(
        texts(&m)[..5],
        ["Informe técnico", "1. Introducción", "Nuevo objeto.", "Segundo párrafo.", "2. Caracterización"],
        "the heading stays and the nested section goes away with the old content"
    );
    let keys: Vec<String> = m.sections().unwrap().into_iter().map(|s| s.key).collect();
    assert_eq!(keys, ["introduccion", "caracterizacion"]);
    let intro = &m.sections().unwrap()[0];
    assert_eq!((intro.first, intro.last, intro.title.as_str()), (Some(1), Some(3), "Introducción"));
    // The new paragraphs are not headings
    let paragraphs = m.extract_paragraphs().unwrap();
    assert!(paragraphs[1].outline_level.is_some() && paragraphs[2].outline_level.is_none());

    // Without keepHeading, the whole content goes, table included
    let (first, _) = m.replace_section("caracterizacion", &[np("Sin datos.")], false).unwrap();
    assert_eq!(first, 4);
    assert_eq!(texts(&m)[4..], ["Sin datos.", "Fuera de toda sección", "ACME"]);
    let car = &m.sections().unwrap()[1];
    assert_eq!((car.first, car.last), (Some(4), Some(4)));
    assert_eq!(reopen(&m).sections().unwrap(), m.sections().unwrap());
}

#[test]
fn test_replace_section_without_a_heading_keeps_everything_inside() {
    let mut m = report();
    // "alcance" opens with a regular paragraph: keepHeading has nothing to keep
    m.replace_section("alcance", &[np("Edificios A y B.")], true).unwrap();
    let alcance = m.sections().unwrap().into_iter().find(|s| s.key == "alcance").unwrap();
    assert_eq!((alcance.first, alcance.last, alcance.depth), (Some(3), Some(3), 1));
    assert_eq!(texts(&m)[3], "Edificios A y B.");
}

#[test]
fn test_unwrap_section_keeps_content_and_nested_sections() {
    let mut m = report();
    let before = texts(&m);
    m.unwrap_section("introduccion").unwrap();
    assert_eq!(texts(&m), before);
    let sections = m.sections().unwrap();
    let alcance = sections.iter().find(|s| s.key == "alcance").unwrap();
    assert_eq!((alcance.depth, alcance.parent.as_deref()), (0, None));
    assert!(!sections.iter().any(|s| s.key == "introduccion"));
    assert!(m.unwrap_section("introduccion").unwrap_err().contains("No existe"));
}

#[test]
fn test_duplicated_keys_are_listed_but_refused_by_key() {
    let mut m = docx(&[section("a", "A", &para("1")), section("a", "A", &para("2"))].concat());
    assert_eq!(m.sections().unwrap().len(), 2);
    assert!(m.replace_section("a", &[np("x")], false).unwrap_err().contains("Hay 2 secciones"));
    assert!(m.define_section("a", 0, 0, None, false).unwrap_err().contains("Ya existe"));
}

#[test]
fn test_section_edits_are_one_undo_step_each() {
    let mut m = report();
    let original = m.get_file_string("word/document.xml").unwrap();
    m.checkpoint(None);
    m.define_section("final", 8, 8, Some("Final"), false).unwrap();
    m.checkpoint(None);
    m.replace_section("final", &[np("Otro texto.")], false).unwrap();
    m.checkpoint(None);
    m.unwrap_section("final").unwrap();
    for _ in 0..3 {
        m.undo();
    }
    assert_eq!(m.get_file_string("word/document.xml").unwrap(), original);
}
