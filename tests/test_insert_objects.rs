use rust_web_docx::blank_generator::{generate_blank_docx, PageSize};
use rust_web_docx::docx_parser::{DocumentElement, DocxModifier, NewImage, NewTable};
use std::io::{Cursor, Read, Write};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

/// Minimal .docx around the given body XML (no relationships part, self-closing content types)
fn docx(body: &str) -> DocxModifier {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#,
        body
    );
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let opts = SimpleFileOptions::default();
        for (name, content) in [
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#),
            ("word/document.xml", document.as_str()),
        ] {
            zip.start_file(name, opts).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
    }
    DocxModifier::from_bytes(&buf.into_inner()).unwrap()
}

fn document(m: &DocxModifier) -> String {
    m.get_file_string("word/document.xml").unwrap()
}

fn texts(m: &DocxModifier) -> Vec<String> {
    m.extract_paragraphs().unwrap().into_iter().map(|p| p.text).collect()
}

fn table(rows: &[&[&str]]) -> NewTable {
    NewTable {
        rows: rows.iter().map(|r| r.iter().map(|c| c.to_string()).collect()).collect(),
        header: true,
        widths: None,
        align: None,
    }
}

fn png(w: u32, h: u32) -> Vec<u8> {
    let mut b = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR".to_vec();
    b.extend_from_slice(&w.to_be_bytes());
    b.extend_from_slice(&h.to_be_bytes());
    b.extend_from_slice(&[8, 6, 0, 0, 0, 0, 0, 0, 0]);
    b
}

/// Re-reads the exported file, as Word or another program would
fn reopen(m: &DocxModifier) -> DocxModifier {
    DocxModifier::from_bytes(&m.to_bytes().unwrap()).unwrap()
}

#[test]
fn test_table_in_the_middle_of_a_paragraph() {
    let mut m = docx(r#"<w:p><w:r><w:t>antes después</w:t></w:r></w:p>"#);
    let t = table(&[&["Ítem", "Valor"], &["Diseño", "$1.200.000"], &["Soporte & mantenimiento", "$300.000"]]);
    let (first, caret) = m.insert_table(0, 6, &t).unwrap();

    // antes(0) | 6 cells (1..=6) | después(7)
    assert_eq!(first, 1);
    assert_eq!(caret, (7, 0));
    let elements = m.extract_elements().unwrap();
    assert_eq!(elements.len(), 3, "paragraph, table, paragraph");
    match &elements[1] {
        DocumentElement::Table(info) => {
            assert_eq!(info.rows, vec![
                vec!["Ítem".to_string(), "Valor".to_string()],
                vec!["Diseño".to_string(), "$1.200.000".to_string()],
                vec!["Soporte & mantenimiento".to_string(), "$300.000".to_string()],
            ]);
            assert!(info.rich_rows[0].is_header && !info.rich_rows[1].is_header);
            assert!(info.rich_rows[0].cells[0].bold, "header in bold");
            assert_eq!(info.rich_rows[0].cells[0].bg_color.as_deref(), Some("F2F2F2"));
            // 8.5in − 2in of margins = 6.5in = 9360 twips, split in two columns
            assert_eq!(info.grid_cols, vec![4680.0, 4680.0]);
        }
        other => panic!("expected a table, got {:?}", other),
    }
    let xml = document(&m);
    assert!(xml.contains("antes </w:t></w:r></w:p><w:tbl>"), "{}", xml);
    assert!(xml.contains("</w:tbl><w:p>"), "text after the position goes below the table");
    assert!(xml.contains(r#"<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>"#));
    assert!(xml.contains("Soporte &amp; mantenimiento"));
    assert!(!xml.contains("1E3A8A"), "gray look, no blue");
}

#[test]
fn test_table_at_start_end_and_in_a_blank_document() {
    let mut m = docx(r#"<w:p><w:r><w:t>hola</w:t></w:r></w:p>"#);
    // At the end: a new empty paragraph follows the table and gets the caret
    let (first, caret) = m.insert_table(0, 4, &table(&[&["A"]])).unwrap();
    assert_eq!((first, caret), (1, (2, 0)));
    assert_eq!(texts(&m), vec!["hola", ""]);

    // At the start: the table goes before the paragraph
    let (first, caret) = m.insert_table(0, 0, &table(&[&["B", "C"]])).unwrap();
    assert_eq!((first, caret), (0, (2, 0)));
    assert!(document(&m).starts_with(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl>"#));

    let mut blank = DocxModifier::from_bytes(&generate_blank_docx(PageSize::Letter).unwrap()).unwrap();
    let (_, caret) = blank.insert_table(0, 0, &table(&[&["x", "y"], &["1", "2"]])).unwrap();
    assert_eq!(caret, (4, 0));
    let reopened = reopen(&blank);
    assert_eq!(reopened.extract_tables().unwrap().len(), 1);
}

#[test]
fn test_table_options_and_errors() {
    let mut m = docx(r#"<w:p/>"#);
    let t = NewTable {
        rows: vec![vec!["Concepto".into(), "Total".into()], vec!["Línea 1\nLínea 2".into()]],
        header: false,
        widths: Some(vec![3.0, 1.0]),
        align: Some(vec!["left".into(), "right".into()]),
    };
    m.insert_table(0, 0, &t).unwrap();
    let xml = document(&m);
    assert!(!xml.contains("<w:tblHeader/>") && !xml.contains("F2F2F2"), "no header row");
    assert!(xml.contains(r#"<w:gridCol w:w="7020"/><w:gridCol w:w="2340"/>"#), "{}", xml);
    assert!(xml.contains(r#"<w:jc w:val="right"/>"#));
    assert!(xml.contains("Línea 1</w:t><w:br/><w:t xml:space=\"preserve\">Línea 2"), "line breaks stay in the cell");
    // The short row is padded so every row has two cells
    assert_eq!(xml.matches("<w:tc>").count(), 4);

    assert!(m.insert_table(0, 0, &table(&[])).is_err());
    // The caret inside a cell: no nested tables
    let err = m.insert_table(0, 0, &table(&[&["x"]])).unwrap_err();
    assert!(err.contains("dentro de otra tabla"), "{}", err);
    assert!(m.insert_table(99, 0, &table(&[&["x"]])).is_err());
}

#[test]
fn test_table_is_one_undo_step() {
    let mut m = docx(r#"<w:p><w:r><w:t>antes después</w:t></w:r></w:p>"#);
    let original = document(&m);
    m.checkpoint(Some("sel".into()));
    m.insert_table(0, 6, &table(&[&["a", "b"]])).unwrap();
    m.set_selection_after(Some("after".into()));
    assert_eq!(m.undo(), Some(Some("sel".into())));
    assert_eq!(document(&m), original);
    assert_eq!(m.redo(), Some(Some("after".into())));
    assert_eq!(m.extract_tables().unwrap().len(), 1);
}

#[test]
fn test_legacy_add_table_goes_before_the_section_properties() {
    let mut m = docx(r#"<w:p><w:r><w:t>hola</w:t></w:r></w:p>"#);
    m.add_table(2, 3, &["A".into(), "B".into()]).unwrap();
    let xml = document(&m);
    let tbl = xml.find("<w:tbl>").unwrap();
    assert!(tbl < xml.rfind("<w:sectPr").unwrap(), "{}", xml);
    assert!(!xml.contains(">Dato<"), "cells start empty");
    let info = &m.extract_tables().unwrap()[0];
    assert_eq!(info.rows[0], vec!["A", "B", ""]);
}

#[test]
fn test_image_in_its_own_paragraph() {
    let mut m = docx(r#"<w:p><w:r><w:t>antes después</w:t></w:r></w:p>"#);
    let options = NewImage { align: Some("center".into()), alt: Some("Logo \"ACME\"".into()), ..Default::default() };
    let (index, caret) = m.insert_image(0, 6, &png(320, 160), &options).unwrap();

    assert_eq!((index, caret), (1, (2, 0)));
    assert_eq!(texts(&m), vec!["antes ", "", "después"]);
    let picture = &m.extract_paragraphs().unwrap()[1];
    assert_eq!(picture.align, "center");
    assert_eq!(picture.images.len(), 1);
    assert_eq!((picture.images[0].width, picture.images[0].height), (320.0, 160.0));
    assert!(!picture.images[0].anchored);

    let xml = document(&m);
    assert!(xml.contains(r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""#));
    assert!(xml.contains(r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#));
    assert!(xml.contains(r#"descr="Logo &quot;ACME&quot;""#));

    let rels = m.get_file_string("word/_rels/document.xml.rels").unwrap();
    assert!(rels.contains(r#"Id="rIdImg1""#) && rels.contains(r#"Target="media/image1.png""#), "{}", rels);
    let types = m.get_file_string("[Content_Types].xml").unwrap();
    assert!(types.contains(r#"<Default Extension="png" ContentType="image/png"/></Types>"#), "{}", types);

    // The exported file holds the picture and opens again with it
    let bytes = m.to_bytes().unwrap();
    let mut zip = ZipArchive::new(Cursor::new(bytes.clone())).unwrap();
    let mut media = Vec::new();
    zip.by_name("word/media/image1.png").unwrap().read_to_end(&mut media).unwrap();
    assert_eq!(media, png(320, 160));
    let reopened = DocxModifier::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.extract_paragraphs().unwrap()[1].images.len(), 1);
}

#[test]
fn test_images_get_unique_names_ids_and_fit_the_page() {
    let mut m = DocxModifier::from_bytes(&generate_blank_docx(PageSize::Letter).unwrap()).unwrap();
    // Blank document: image + empty paragraph for the caret
    let (index, caret) = m.insert_image(0, 0, &png(4000, 1000), &NewImage::default()).unwrap();
    assert_eq!((index, caret), (0, (1, 0)));
    let (index, _) = m.insert_image(1, 0, &png(10, 10), &NewImage { width: Some(50.0), ..Default::default() }).unwrap();
    assert_eq!(index, 1);

    let paragraphs = m.extract_paragraphs().unwrap();
    // Letter (816 px) minus 1in margins (96 px each side) = 624 px of text width
    assert_eq!((paragraphs[0].images[0].width, paragraphs[0].images[0].height), (624.0, 156.0));
    assert_eq!((paragraphs[1].images[0].width, paragraphs[1].images[0].height), (50.0, 50.0));

    let xml = document(&m);
    assert!(xml.contains(r#"<wp:docPr id="1""#) && xml.contains(r#"<wp:docPr id="2""#));
    let rels = m.get_file_string("word/_rels/document.xml.rels").unwrap();
    assert!(rels.contains("media/image1.png") && rels.contains("media/image2.png"));
    let types = m.get_file_string("[Content_Types].xml").unwrap();
    assert_eq!(types.matches(r#"Extension="png""#).count(), 1, "one content type per extension");
}

#[test]
fn test_image_undo_removes_every_part() {
    let mut m = docx(r#"<w:p><w:r><w:t>hola</w:t></w:r></w:p>"#);
    let original = document(&m);
    m.checkpoint(None);
    m.insert_image(0, 4, &png(20, 20), &NewImage::default()).unwrap();
    assert!(m.get_file_string("word/_rels/document.xml.rels").is_ok());
    m.undo();
    assert_eq!(document(&m), original);
    assert!(m.get_file_string("word/_rels/document.xml.rels").is_err(), "relationships part removed");
    let mut zip = ZipArchive::new(Cursor::new(m.to_bytes().unwrap())).unwrap();
    assert!(zip.by_name("word/media/image1.png").is_err());
}

#[test]
fn test_image_errors() {
    let mut m = docx(r#"<w:p/>"#);
    let original = document(&m);
    assert!(m.insert_image(0, 0, b"no es una imagen", &NewImage::default()).unwrap_err().contains("PNG, JPEG o GIF"));
    let bad = NewImage { align: Some("justify".into()), ..Default::default() };
    assert!(m.insert_image(0, 0, &png(1, 1), &bad).is_err());
    assert_eq!(document(&m), original, "nothing changes on error");
}
