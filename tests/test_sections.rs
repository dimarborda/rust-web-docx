//! Logical sections (block-level content controls tagged `section:<key>`): listing, defining,
//! replacing and unwrapping them, and that the saved file keeps them.

use rust_web_docx::docx_parser::{
    section_key_from_title, DocumentElement, DocxModifier, NewParagraph, SchemaSection, SectionInfo, SectionPlace, SectionSchema,
};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

const HEADING_STYLE: &str = r#"<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>"#;

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

fn heading2(text: &str) -> String {
    format!(r#"<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>{}</w:t></w:r></w:p>"#, text)
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
    let info = m.define_section("conclusiones", 0, 4, Some("Conclusiones"), true, false).unwrap();
    assert_eq!(
        info,
        SectionInfo { key: "conclusiones".into(), title: "Conclusiones".into(), first: Some(0), last: Some(4), depth: 0, parent: None, locked: true, content_locked: false }
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
    let datos = m.define_section("datos", 5, 6, None, false, false).unwrap();
    assert_eq!((datos.first, datos.last, datos.title.as_str()), (Some(5), Some(6), "datos"));
    assert_eq!((datos.depth, datos.parent.as_deref()), (1, Some("caracterizacion")));
    assert!(m.get_file_string("word/document.xml").unwrap().contains("<w:sdtContent><w:tbl>"));
}

#[test]
fn test_define_section_refusals_leave_the_document_untouched() {
    let mut m = report();
    let before = m.get_file_string("word/document.xml").unwrap();
    let err = |r: Result<SectionInfo, String>| r.unwrap_err();
    assert!(err(m.define_section("introduccion", 8, 8, None, false, false)).contains("Ya existe"));
    assert!(err(m.define_section("mal clave", 8, 8, None, false, false)).contains("no válida"));
    assert!(err(m.define_section("", 8, 8, None, false, false)).contains("no válida"));
    // Cuts the Introducción section (starts outside, ends inside)
    assert!(err(m.define_section("x", 0, 2, None, false, false)).contains("corta"));
    // Starts in the middle of a table
    assert!(err(m.define_section("x", 6, 7, None, false, false)).contains("tabla"));
    assert!(err(m.define_section("x", 50, 51, None, false, false)).contains("no existe"));
    assert_eq!(m.get_file_string("word/document.xml").unwrap(), before);
}

#[test]
fn test_define_section_refuses_a_word_section_break() {
    let mut m = docx(&[para("A"), r#"<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>B</w:t></w:r></w:p>"#.to_string(), para("C")].concat());
    assert!(m.define_section("x", 0, 2, None, false, false).unwrap_err().contains("salto de sección"));
    assert!(m.define_section("x", 2, 2, None, false, false).is_ok());
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
    assert!(m.define_section("a", 0, 0, None, false, false).unwrap_err().contains("Ya existe"));
}

#[test]
fn test_section_edits_are_one_undo_step_each() {
    let mut m = report();
    let original = m.get_file_string("word/document.xml").unwrap();
    m.checkpoint(None);
    m.define_section("final", 8, 8, Some("Final"), false, false).unwrap();
    m.checkpoint(None);
    m.replace_section("final", &[np("Otro texto.")], false).unwrap();
    m.checkpoint(None);
    m.unwrap_section("final").unwrap();
    for _ in 0..3 {
        m.undo();
    }
    assert_eq!(m.get_file_string("word/document.xml").unwrap(), original);
}

// ---------- Locks ----------

fn document(m: &DocxModifier) -> String {
    m.get_file_string("word/document.xml").unwrap()
}

#[test]
fn test_set_section_lock_writes_word_lock_values_in_place() {
    let mut m = report();
    let locks = |m: &DocxModifier| {
        let s = m.sections().unwrap().into_iter().find(|s| s.key == "alcance").unwrap();
        (s.locked, s.content_locked)
    };
    for (locked, content, val) in [(true, true, Some("sdtContentLocked")), (false, true, Some("contentLocked")), (true, false, Some("sdtLocked")), (false, false, None)] {
        m.set_section_lock("alcance", locked, content).unwrap();
        assert_eq!(locks(&m), (locked, content));
        let xml = document(&m);
        let pr = &xml[xml.find(r#"<w:alias w:val="Alcance"/>"#).unwrap()..];
        let pr = &pr[..pr.find("</w:sdtPr>").unwrap()];
        assert_eq!(pr.matches("<w:lock").count(), val.is_some() as usize, "one lock at most");
        if let Some(val) = val {
            assert!(pr.ends_with(&format!(r#"<w:id w:val="1"/><w:lock w:val="{}"/>"#, val)), "lock goes after w:id: {}", pr);
        }
    }
    let info = m.define_section("cierre", 8, 8, None, false, true).unwrap();
    assert_eq!((info.locked, info.content_locked), (false, true));
    assert!(document(&m).contains(r#"<w:lock w:val="contentLocked"/>"#));
    assert_eq!(reopen(&m).sections().unwrap(), m.sections().unwrap());
}

// ---------- Sections from headings ----------

fn technical_report() -> DocxModifier {
    docx(&[
        para("Informe técnico de suelos"),
        heading("1. Introducción"),
        para("Objeto."),
        heading2("1.1 Alcance"),
        para("Edificio A."),
        heading("2. Caracterización del sitio"),
        table(&["Muestra", "pH"]),
        para("Notas."),
        heading("3. Conclusiones"),
        para("Apto."),
        heading("Anexo"),
        heading("Anexo"),
    ]
    .concat())
}

#[test]
fn test_section_keys_from_titles() {
    assert_eq!(section_key_from_title("2.1 Caracterización del sitio"), "caracterizacion-del-sitio");
    assert_eq!(section_key_from_title("3) Conclusiones y recomendaciones"), "conclusiones-y-recomendaciones");
    assert_eq!(section_key_from_title("Año 2026: Señalización"), "ano-2026-senalizacion");
    assert_eq!(section_key_from_title("  "), "seccion");
    assert_eq!(section_key_from_title(&"palabra ".repeat(20)).chars().count(), 63);
}

#[test]
fn test_sections_from_headings_nest_lower_levels() {
    let mut m = technical_report();
    let result = m.sections_from_headings(1, 2).unwrap();
    let summary: Vec<(String, String, Option<usize>, Option<usize>, Option<String>)> =
        result.created.iter().map(|s| (s.key.clone(), s.title.clone(), s.first, s.last, s.parent.clone())).collect();
    let row = |k: &str, t: &str, f: usize, l: usize, p: Option<&str>| (k.to_string(), t.to_string(), Some(f), Some(l), p.map(String::from));
    assert_eq!(
        summary,
        [
            row("introduccion", "1. Introducción", 1, 4, None),
            row("alcance", "1.1 Alcance", 3, 4, Some("introduccion")),
            row("caracterizacion-del-sitio", "2. Caracterización del sitio", 5, 8, None),
            row("conclusiones", "3. Conclusiones", 9, 10, None),
            row("anexo", "Anexo", 11, 11, None),
            row("anexo-2", "Anexo", 12, 12, None),
        ]
    );
    assert!(result.skipped.is_empty());
    assert_eq!(texts(&m)[0], "Informe técnico de suelos", "text before the first heading stays outside");
    assert_eq!(reopen(&m).sections().unwrap(), m.sections().unwrap());

    // Again: every heading already opens a section
    let again = m.sections_from_headings(1, 2).unwrap();
    assert!(again.created.is_empty());
    assert_eq!(again.skipped.len(), 6);
    assert!(again.skipped[0].reason.contains("ya abre la sección «introduccion»"));
}

#[test]
fn test_sections_from_headings_only_the_given_levels_in_one_undo_step() {
    let mut m = technical_report();
    let original = document(&m);
    m.checkpoint(None);
    let created = m.sections_from_headings(1, 1).unwrap().created;
    assert_eq!(created.iter().map(|s| s.key.as_str()).collect::<Vec<_>>(), ["introduccion", "caracterizacion-del-sitio", "conclusiones", "anexo", "anexo-2"]);
    assert_eq!(created[0].last, Some(4), "level 2 headings stay inside");
    m.undo();
    assert_eq!(document(&m), original);
    assert!(m.sections_from_headings(0, 1).is_err());
    assert!(m.sections_from_headings(2, 1).is_err());
}

#[test]
fn test_sections_from_headings_stay_inside_existing_sections_and_skip_breaks() {
    let mut m = docx(&[
        section("cuerpo", "Cuerpo", &[heading("A"), para("a"), heading("B"), para("b")].concat()),
        heading("Fuera"),
        r#"<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>con salto</w:t></w:r></w:p>"#.to_string(),
        heading("Final"),
    ]
    .concat());
    let result = m.sections_from_headings(1, 1).unwrap();
    let created: Vec<(&str, Option<usize>, Option<usize>, Option<&str>)> =
        result.created.iter().map(|s| (s.key.as_str(), s.first, s.last, s.parent.as_deref())).collect();
    assert_eq!(created, [("a", Some(0), Some(1), Some("cuerpo")), ("b", Some(2), Some(3), Some("cuerpo")), ("final", Some(6), Some(6), None)]);
    assert_eq!(result.skipped.len(), 1);
    assert_eq!((result.skipped[0].paragraph, result.skipped[0].text.as_str()), (4, "Fuera"));
    assert!(result.skipped[0].reason.contains("salto de sección"));
}

// ---------- Inserting sections ----------

fn heading_item(text: &str) -> NewParagraph {
    NewParagraph { text: text.into(), style: Some("Heading1".into()), ..Default::default() }
}

#[test]
fn test_insert_section_at_each_place() {
    let mut m = report();
    let (end, first) = m.insert_section("anexos", Some("Anexos"), &SectionPlace::End, &[heading_item("Anexos"), np("")], false, false).unwrap();
    assert_eq!((first, end.first, end.last, end.depth), (10, Some(10), Some(11), 0));
    assert!(document(&m).contains("</w:sdt><w:sectPr>"), "before the body's sectPr");

    m.insert_section("objetivos", Some("Objetivos"), &SectionPlace::After("alcance".into()), &[np("Objetivo 1.")], false, false).unwrap();
    m.insert_section("resumen", None, &SectionPlace::Before("introduccion".into()), &[np("Resumen.")], true, false).unwrap();
    m.insert_section("metodo", Some("Método"), &SectionPlace::Inside("caracterizacion".into()), &[np("Método.")], false, true).unwrap();

    let order: Vec<(String, Option<String>, Option<usize>)> = m.sections().unwrap().into_iter().map(|s| (s.key, s.parent, s.first)).collect();
    let row = |k: &str, p: Option<&str>, f: usize| (k.to_string(), p.map(String::from), Some(f));
    assert_eq!(
        order,
        [
            row("resumen", None, 1),
            row("introduccion", None, 2),
            row("alcance", Some("introduccion"), 4),
            row("objetivos", Some("introduccion"), 5),
            row("caracterizacion", None, 6),
            row("metodo", Some("caracterizacion"), 10),
            row("anexos", None, 13),
        ]
    );
    let texts = texts(&m);
    assert_eq!((texts[1].as_str(), texts[5].as_str(), texts[10].as_str(), texts[13].as_str()), ("Resumen.", "Objetivo 1.", "Método.", "Anexos"));
    let heading = m.extract_paragraphs().unwrap().into_iter().find(|p| p.index == 13).unwrap();
    assert_eq!(heading.outline_level, Some(0), "the heading item gets Word's Heading 1");
    assert!(m.insert_section("anexos", None, &SectionPlace::End, &[np("x")], false, false).unwrap_err().contains("Ya existe"));
    assert!(m.insert_section("x", None, &SectionPlace::After("nada".into()), &[np("x")], false, false).unwrap_err().contains("No existe"));
}

// ---------- Schema ----------

fn schema() -> SectionSchema {
    let s = |key: &str, title: &str, required: bool, parent: Option<&str>| SchemaSection {
        key: key.into(),
        title: Some(title.into()),
        required,
        parent: parent.map(String::from),
    };
    SectionSchema {
        kind: "informe-tecnico".into(),
        title: Some("Informe técnico".into()),
        sections: vec![
            s("introduccion", "Introducción", true, None),
            s("alcance", "Alcance", false, Some("introduccion")),
            s("caracterizacion", "Caracterización & datos", true, None),
            s("conclusiones", "Conclusiones", true, None),
        ],
    }
}

#[test]
fn test_section_schema_is_stored_as_custom_xml_and_survives_saving() {
    let mut m = report();
    assert_eq!(m.section_schema().unwrap(), None);
    m.set_section_schema(Some(&schema())).unwrap();
    assert_eq!(m.section_schema().unwrap(), Some(schema()));
    let item = m.get_file_string("customXml/item1.xml").unwrap();
    assert!(item.contains(r#"<sections xmlns="urn:rust-web-docx:sections:1" type="informe-tecnico""#));
    assert!(m.get_file_string("customXml/itemProps1.xml").unwrap().contains("ds:schemaRef ds:uri=\"urn:rust-web-docx:sections:1\""));
    assert!(m.get_file_string("customXml/_rels/item1.xml.rels").unwrap().contains(r#"Target="itemProps1.xml""#));
    assert!(m.get_file_string("word/_rels/document.xml.rels").unwrap().contains(r#"Target="../customXml/item1.xml""#));
    let types = m.get_file_string("[Content_Types].xml").unwrap();
    assert!(types.contains(r#"PartName="/customXml/itemProps1.xml""#) && types.contains(r#"Extension="xml""#));
    assert_eq!(reopen(&m).section_schema().unwrap(), Some(schema()));

    // Replacing rewrites the same part
    let mut other = schema();
    other.sections.truncate(1);
    m.set_section_schema(Some(&other)).unwrap();
    assert_eq!(m.section_schema().unwrap(), Some(other));
    assert!(m.get_file_string("customXml/item2.xml").is_err());

    // Removing leaves no trace
    m.set_section_schema(None).unwrap();
    assert_eq!(m.section_schema().unwrap(), None);
    for part in ["customXml/item1.xml", "customXml/itemProps1.xml", "customXml/_rels/item1.xml.rels"] {
        assert!(m.get_file_string(part).is_err(), "{} removed", part);
    }
    assert!(!m.get_file_string("word/_rels/document.xml.rels").unwrap().contains("customXml"));
    assert!(!m.get_file_string("[Content_Types].xml").unwrap().contains("itemProps1"));
    assert_eq!(reopen(&m).section_schema().unwrap(), None);
}

#[test]
fn test_section_schema_refuses_invalid_lists_and_undoes_in_one_step() {
    let mut m = report();
    let mut bad = schema();
    bad.sections[1].parent = Some("conclusiones".into());
    assert!(m.set_section_schema(Some(&bad)).unwrap_err().contains("debe aparecer antes"));
    let mut bad = schema();
    bad.sections.push(bad.sections[0].clone());
    assert!(m.set_section_schema(Some(&bad)).unwrap_err().contains("repetida"));
    let mut bad = schema();
    bad.kind = " ".into();
    assert!(m.set_section_schema(Some(&bad)).is_err());

    m.checkpoint(None);
    m.set_section_schema(Some(&schema())).unwrap();
    m.undo();
    assert_eq!(m.section_schema().unwrap(), None);
    assert!(m.get_file_string("customXml/item1.xml").is_err());
}
