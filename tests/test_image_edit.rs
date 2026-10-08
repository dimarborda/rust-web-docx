use rust_web_docx::docx_parser::{DocxModifier, ImageRef, ImageUpdate, NewImage};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

/// Minimal .docx around the given body XML
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

fn png(w: u32, h: u32) -> Vec<u8> {
    let mut b = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR".to_vec();
    b.extend_from_slice(&w.to_be_bytes());
    b.extend_from_slice(&h.to_be_bytes());
    b.extend_from_slice(&[8, 6, 0, 0, 0, 0, 0, 0, 0]);
    b
}

fn images(m: &DocxModifier, paragraph: usize) -> Vec<ImageRef> {
    m.extract_paragraphs().unwrap().into_iter().find(|p| p.index == paragraph).unwrap().images
}

fn texts(m: &DocxModifier) -> Vec<String> {
    m.extract_paragraphs().unwrap().into_iter().map(|p| p.text).collect()
}

fn reopen(m: &DocxModifier) -> DocxModifier {
    DocxModifier::from_bytes(&m.to_bytes().unwrap()).unwrap()
}

#[test]
fn test_resize_move_and_wrap_an_inserted_picture() {
    let mut m = docx(r#"<w:p><w:r><w:t>Texto</w:t></w:r></w:p>"#);
    m.checkpoint(None);
    let (p, _) = m.insert_image(0, 5, &png(400, 200), &NewImage::default()).unwrap();
    assert_eq!(images(&m, p)[0].wrap, "inline");

    m.checkpoint(None);
    let img = m
        .update_image(p, 0, &ImageUpdate { width: Some(100.0), wrap: Some("square".into()), ..Default::default() })
        .unwrap();
    assert_eq!((img.width, img.height), (100.0, 50.0), "aspect ratio kept");
    assert!(img.anchored);
    assert_eq!(img.wrap, "square");

    m.checkpoint(None);
    let moved = ImageUpdate {
        h_relative: Some("margin".into()),
        h_align: Some("right".into()),
        v_relative: Some("paragraph".into()),
        v_offset: Some(12.0),
        wrap_side: Some("left".into()),
        alt: Some("Logo & sello".into()),
        ..Default::default()
    };
    m.update_image(p, 0, &moved).unwrap();

    // Survives export and a fresh read, as Word would see it
    let r = reopen(&m);
    let img = &images(&r, p)[0];
    assert_eq!(img.wrap, "square");
    assert_eq!(img.wrap_side, "left");
    assert_eq!((img.h_relative.as_str(), img.h_align.as_deref()), ("margin", Some("right")));
    assert_eq!((img.v_relative.as_str(), img.v_offset), ("paragraph", 12.0));
    assert_eq!(img.alt, "Logo & sello");
    let doc = r.get_file_string("word/document.xml").unwrap();
    assert!(doc.contains("xmlns:wp="), "the wp namespace is declared");
    assert_eq!(doc.matches("<wp:docPr").count(), 1);

    // Each change is its own undo step
    assert!(m.undo().is_some());
    assert_eq!(images(&m, p)[0].h_align.as_deref(), None);
    assert!(m.undo().is_some());
    assert_eq!(images(&m, p)[0].wrap, "inline");
    assert_eq!(images(&m, p)[0].width, 400.0);
}

#[test]
fn test_floating_insert_anchors_to_the_current_paragraph() {
    let mut m = docx(r#"<w:p><w:r><w:t>Uno</w:t></w:r></w:p><w:p><w:r><w:t>Dos</w:t></w:r></w:p>"#);
    let opts = NewImage {
        width: Some(120.0),
        wrap: Some("tight".into()),
        h_relative: Some("page".into()),
        h_offset: Some(40.0),
        v_relative: Some("page".into()),
        v_offset: Some(60.0),
        ..Default::default()
    };
    let (p, caret) = m.insert_image(1, 2, &png(300, 150), &opts).unwrap();
    assert_eq!(p, 1);
    assert_eq!(caret, (1, 2), "the caret stays in the text");
    assert_eq!(texts(&m), vec!["Uno", "Dos"], "no paragraph was added");
    let img = &images(&m, 1)[0];
    assert_eq!(img.wrap, "tight");
    assert_eq!((img.width, img.height), (120.0, 60.0));
    assert_eq!((img.h_relative.as_str(), img.h_offset), ("page", 40.0));
    assert!(m.get_file_string("word/document.xml").unwrap().contains("<wp:wrapPolygon"));

    // A position needs a floating wrap
    let bad = NewImage { h_offset: Some(10.0), ..Default::default() };
    assert!(m.insert_image(0, 0, &png(10, 10), &bad).is_err());
}

#[test]
fn test_inline_offset_follows_the_text() {
    let m = docx(concat!(
        r#"<w:p xmlns:wp="wp" xmlns:a="a" xmlns:r="r"><w:r><w:t xml:space="preserve">Hola </w:t></w:r>"#,
        r#"<w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="952500"/><wp:docPr id="1" name="x"/>"#,
        r#"<a:graphic><a:graphicData><a:blip r:embed="rId1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#,
        r#"<w:r><w:t>mundo</w:t></w:r></w:p>"#
    ));
    let img = &images(&m, 0)[0];
    assert_eq!(img.offset, 5);
    assert_eq!(img.width, 100.0);
}

#[test]
fn test_delete_picture_and_its_empty_run() {
    let mut m = docx(r#"<w:p><w:r><w:t>Texto</w:t></w:r></w:p>"#);
    let (p, _) = m.insert_image(0, 5, &png(40, 40), &NewImage::default()).unwrap();
    m.checkpoint(None);
    m.delete_image(p, 0).unwrap();
    assert!(images(&m, p).is_empty());
    let doc = m.get_file_string("word/document.xml").unwrap();
    assert!(!doc.contains("<w:drawing>"));
    assert!(!doc.contains("<w:r></w:r>"), "no empty run left behind");
    assert!(m.undo().is_some());
    assert_eq!(images(&m, p).len(), 1);
    assert!(m.delete_image(p, 3).is_err());
}

#[test]
fn test_text_edits_keep_pictures_in_place() {
    let mut m = docx(r#"<w:p><w:r><w:t>Texto</w:t></w:r></w:p>"#);
    let opts = NewImage { wrap: Some("square".into()), ..Default::default() };
    m.insert_image(0, 0, &png(40, 40), &opts).unwrap();
    m.replace_paragraph_range(0, 0, 0, "Nuevo ").unwrap();
    assert_eq!(texts(&m), vec!["Nuevo Texto"]);
    assert_eq!(images(&m, 0)[0].wrap, "square");
}

