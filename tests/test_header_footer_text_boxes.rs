//! A floating text box anchored inside a header or footer (institutional letterhead with a
//! logo in a text box): its paragraphs must get an editable index and be fillable, same as a
//! text box in the body (`tests/test_text_boxes.rs`), without disturbing the body or the
//! header/footer's own top-level paragraphs (`tests/test_header_footer.rs`).

use rust_web_docx::docx_parser::{
    DocumentElement, DocxModifier, ImageUpdate, NewImage, HEADER_FOOTER_BASE, HEADER_FOOTER_TEXTBOX_BASE, PART_SLOT, TextBoxContent,
    TextBoxImage,
};
use rust_web_docx::layout_engine::{LayoutEngine, RenderCommand};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

fn png(w: u32, h: u32) -> Vec<u8> {
    let mut b = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR".to_vec();
    b.extend_from_slice(&w.to_be_bytes());
    b.extend_from_slice(&h.to_be_bytes());
    b.extend_from_slice(&[8, 6, 0, 0, 0, 0, 0, 0, 0]);
    b
}

fn b64(bytes: &[u8]) -> String {
    use base64::Engine;
    base64::engine::general_purpose::STANDARD.encode(bytes)
}

const W: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
    r#"xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#
);

/// A floating text box (DrawingML, no VML fallback needed for these tests) with `text` as its
/// single paragraph
fn text_box(text: &str, doc_pr_id: u32) -> String {
    format!(
        concat!(
            r#"<w:r><w:drawing>"#,
            r#"<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">"#,
            r#"<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>"#,
            r#"<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>"#,
            r#"<wp:extent cx="952500" cy="476250"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>"#,
            r#"<wp:docPr id="{id}" name="Cuadro {id}"/><wp:cNvGraphicFramePr/>"#,
            r#"<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>"#,
            r#"<wps:cNvSpPr txBox="1"/><wps:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm>"#,
            r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></wps:spPr>"#,
            r#"<wps:txbx><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>"#,
            r#"<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:noAutofit/></wps:bodyPr>"#,
            r#"</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#,
        ),
        id = doc_pr_id,
        text = text,
    )
}

/// A .docx whose last section uses the given header/footer parts. `parts`: (file name, rel id,
/// reference type, kind, xml body) — same shape as `tests/test_header_footer.rs`
fn docx(parts: &[(&str, &str, &str, &str, &str)]) -> DocxModifier {
    let refs: String = parts
        .iter()
        .map(|(_, id, ty, kind, _)| format!(r#"<w:{}Reference w:type="{}" r:id="{}"/>"#, kind, ty, id))
        .collect();
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body><w:p><w:r><w:t>Cuerpo</w:t></w:r></w:p><w:sectPr>{refs}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>"#,
    );
    let rels: String = parts
        .iter()
        .map(|(file, id, _, kind, _)| {
            format!(r#"<Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{}" Target="{}"/>"#, id, kind, file)
        })
        .collect();
    let rels = format!(r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>"#, rels);
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let opts = SimpleFileOptions::default();
        let mut files = vec![
            ("[Content_Types].xml".to_string(), r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#.to_string()),
            ("word/document.xml".to_string(), document),
            ("word/_rels/document.xml.rels".to_string(), rels),
        ];
        for (file, _, _, kind, xml) in parts {
            let root = if *kind == "header" { "hdr" } else { "ftr" };
            files.push((format!("word/{}", file), format!(r#"<?xml version="1.0"?><w:{root} {W}>{xml}</w:{root}>"#)));
        }
        for (name, content) in files {
            zip.start_file(name, opts).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
    }
    DocxModifier::from_bytes(&buf.into_inner()).unwrap()
}

/// (paragraph index, text) of every text-box paragraph found inside the header/footer elements
/// of `part` ("header" picks `hf.header`, "footer" picks `hf.footer`)
fn box_texts(m: &DocxModifier, part: &str) -> Vec<(usize, String)> {
    let inputs = m.layout_inputs().unwrap();
    let hf = &inputs.header_footer;
    let elements = if part == "header" { &hf.header } else { &hf.footer };
    elements
        .iter()
        .filter_map(|e| match e {
            DocumentElement::Paragraph(p) => Some(p),
            _ => None,
        })
        .flat_map(|p| p.images.iter())
        .filter_map(|img| img.text_box.as_ref())
        .flat_map(|tb| tb.paragraphs.iter().map(|p| (p.index, p.text.clone())))
        .collect()
}

fn part_text(m: &DocxModifier, part: &str) -> String {
    m.get_file_string(part).unwrap()
}

// ---------- 1. A plain-text marker inside a header's text box ----------

#[test]
fn test_marker_inside_a_header_text_box_is_numbered_and_fillable() {
    let header = format!(r#"<w:p>{}</w:p>"#, text_box("{{empresa}}", 1));
    let mut m = docx(&[("header1.xml", "rIdH", "default", "header", &header)]);

    let texts = box_texts(&m, "header");
    assert_eq!(texts.len(), 1, "the box's paragraph is numbered: {:?}", texts);
    let (index, text) = texts[0].clone();
    assert_eq!(text, "{{empresa}}");
    assert!(index >= HEADER_FOOTER_TEXTBOX_BASE, "numbered from HEADER_FOOTER_TEXTBOX_BASE: {}", index);

    // Fillable like any other paragraph (replace_paragraph_range, as the header's own text is
    // already tested with in tests/test_header_footer.rs)
    m.checkpoint(None);
    m.replace_paragraph_range(index, 0, text.chars().count(), "Operadox S.A.S.").unwrap();
    assert_eq!(box_texts(&m, "header")[0].1, "Operadox S.A.S.");
    assert!(part_text(&m, "word/header1.xml").contains("Operadox S.A.S."));

    // The body is untouched
    assert_eq!(m.extract_paragraphs().unwrap()[0].text, "Cuerpo");
}

// ---------- 2. fill_text_box on a marker inside a header's text box ----------

#[test]
fn test_fill_text_box_on_a_header_text_box_replaces_only_the_header_part() {
    let header = format!(r#"<w:p>{}</w:p>"#, text_box("{{html:membrete}}", 1));
    let mut m = docx(&[("header1.xml", "rIdH", "default", "header", &header)]);
    let document_before = part_text(&m, "word/document.xml");

    let index = box_texts(&m, "header")[0].0;
    m.checkpoint(None);
    let content = rust_web_docx::docx_parser::TextBoxContent {
        ooxml: r#"<w:p><w:r><w:t>Operadox S.A.S.</w:t></w:r></w:p>"#.into(),
        images: vec![],
    };
    m.fill_text_box(index, &content).unwrap();

    assert!(part_text(&m, "word/header1.xml").contains("Operadox S.A.S."), "header1.xml's txbxContent was replaced");
    assert!(!part_text(&m, "word/header1.xml").contains("{{html:membrete}}"));
    assert_eq!(part_text(&m, "word/document.xml"), document_before, "document.xml is untouched");

    assert!(m.undo().is_some());
    assert!(part_text(&m, "word/header1.xml").contains("{{html:membrete}}"), "undo restores the header");
}

// ---------- 2b. fill_text_box with a logo image inside a header's text box: the image's
// relationship must live in header1.xml.rels, not document.xml.rels ----------

#[test]
fn test_fill_text_box_with_an_image_relates_it_to_the_header_part() {
    let header = format!(r#"<w:p>{}</w:p>"#, text_box("{{html:logo}}", 1));
    let mut m = docx(&[("header1.xml", "rIdH", "default", "header", &header)]);
    let index = box_texts(&m, "header")[0].0;

    let content = TextBoxContent {
        ooxml: "<w:p>{{img:logo}}</w:p>".into(),
        images: vec![TextBoxImage { placeholder: "logo".into(), data: format!("data:image/png;base64,{}", b64(&png(80, 40))) }],
    };
    m.fill_text_box(index, &content).unwrap();

    let header_xml = part_text(&m, "word/header1.xml");
    assert!(header_xml.contains("<a:blip r:embed="), "the picture is referenced from header1.xml");
    let rel_id = header_xml
        .split("r:embed=\"")
        .nth(1)
        .and_then(|s| s.split('"').next())
        .expect("r:embed present")
        .to_string();

    // The rId must resolve inside header1.xml's OWN relationships, not document.xml's
    let header_rels = part_text(&m, "word/_rels/header1.xml.rels");
    assert!(header_rels.contains(&format!(r#"Id="{}""#, rel_id)), "rId {} is declared in header1.xml.rels: {}", rel_id, header_rels);
    let doc_rels = m.get_file_string("word/_rels/document.xml.rels");
    assert!(
        doc_rels.is_err() || !doc_rels.unwrap().contains(&format!(r#"Id="{}""#, rel_id)),
        "the header's image relationship must not leak into document.xml.rels"
    );

    // The media part exists and the package reopens without error
    let bytes = m.to_bytes().unwrap();
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes.clone())).unwrap();
    assert!(zip.by_name("word/media/image1.png").is_ok(), "word/media/image1.png was added");
    let reopened = DocxModifier::from_bytes(&bytes).unwrap();
    assert!(part_text(&reopened, "word/header1.xml").contains("<a:blip r:embed="));
}

// ---------- 2c. The header's own root is missing xmlns:wp/xmlns:r (a hand-built template):
// ensure_namespace must declare them on <w:hdr>, not assume <w:document> ----------

#[test]
fn test_fill_text_box_with_an_image_declares_missing_namespaces_on_the_header_root() {
    // Deliberately without xmlns:wp / xmlns:r, unlike the `W` constant every other test uses
    let header_root_ns = concat!(
        r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
        r#"xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#
    );
    let header_body = format!(r#"<w:p>{}</w:p>"#, text_box("{{html:logo}}", 1));
    let header_xml = format!(r#"<?xml version="1.0"?><w:hdr {}>{}</w:hdr>"#, header_root_ns, header_body);

    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body><w:p><w:r><w:t>Cuerpo</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>"#,
    );
    let rels = r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#;

    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let opts = SimpleFileOptions::default();
        for (name, content) in [
            ("[Content_Types].xml".to_string(), r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#.to_string()),
            ("word/document.xml".to_string(), document),
            ("word/_rels/document.xml.rels".to_string(), rels.to_string()),
            ("word/header1.xml".to_string(), header_xml),
        ] {
            zip.start_file(name, opts).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
    }
    let mut m = DocxModifier::from_bytes(&buf.into_inner()).unwrap();

    let index = box_texts(&m, "header")[0].0;
    let content = TextBoxContent {
        ooxml: "<w:p>{{img:logo}}</w:p>".into(),
        images: vec![TextBoxImage { placeholder: "logo".into(), data: format!("data:image/png;base64,{}", b64(&png(80, 40))) }],
    };
    m.fill_text_box(index, &content).unwrap();

    let header_after = part_text(&m, "word/header1.xml");
    let hdr_start = header_after.find("<w:hdr").unwrap();
    let root = &header_after[hdr_start..hdr_start + header_after[hdr_start..].find('>').unwrap()];
    assert!(root.contains(r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing""#), "xmlns:wp declared on <w:hdr>: {}", root);
    assert!(root.contains(r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#), "xmlns:r declared on <w:hdr>: {}", root);

    // document.xml did not gain a second declaration of these namespaces (it already has them
    // via `W`, and the fix must not touch document.xml at all for a header's own fill)
    let doc_after = part_text(&m, "word/document.xml");
    assert_eq!(doc_after.matches("xmlns:wp=").count(), 1);
    assert_eq!(doc_after.matches("xmlns:r=").count(), 1);

    // The document reopens without error
    let reopened = DocxModifier::from_bytes(&m.to_bytes().unwrap()).unwrap();
    assert!(part_text(&reopened, "word/header1.xml").contains("<a:blip r:embed="));
}

// ---------- 3. Two text boxes in the same header (logo izquierdo y logo derecho) ----------

#[test]
fn test_two_text_boxes_in_the_same_header_are_numbered_and_filled_separately() {
    let header = format!(r#"<w:p>{}{}</w:p>"#, text_box("{{logo_izquierdo}}", 1), text_box("{{logo_derecho}}", 2));
    let mut m = docx(&[("header1.xml", "rIdH", "default", "header", &header)]);

    let texts = box_texts(&m, "header");
    assert_eq!(texts.len(), 2, "both boxes are numbered: {:?}", texts);
    assert_ne!(texts[0].0, texts[1].0, "each box gets its own index");
    assert_eq!(texts[0].1, "{{logo_izquierdo}}");
    assert_eq!(texts[1].1, "{{logo_derecho}}");

    let (left, right) = (texts[0].0, texts[1].0);
    m.checkpoint(None);
    m.replace_paragraph_range(left, 0, texts[0].1.chars().count(), "Izquierda").unwrap();
    m.replace_paragraph_range(right, 0, texts[1].1.chars().count(), "Derecha").unwrap();

    let after = box_texts(&m, "header");
    assert_eq!(after[0].1, "Izquierda", "left box only");
    assert_eq!(after[1].1, "Derecha", "right box only, not mixed with the left one");
}

// ---------- 4. Header and footer together: slots resolved correctly ----------

#[test]
fn test_header_and_footer_text_boxes_resolve_to_their_own_part() {
    let header = format!(r#"<w:p>{}</w:p>"#, text_box("{{header_box}}", 1));
    let footer = format!(r#"<w:p>{}</w:p>"#, text_box("{{footer_box}}", 1));
    let mut m = docx(&[
        ("header1.xml", "rIdH", "default", "header", &header),
        ("footer1.xml", "rIdF", "default", "footer", &footer),
    ]);

    let header_box = box_texts(&m, "header")[0].clone();
    let footer_box = box_texts(&m, "footer")[0].clone();
    assert_eq!(header_box.1, "{{header_box}}");
    assert_eq!(footer_box.1, "{{footer_box}}");
    // Slots follow header_footer_parts' sort order ("footer1.xml" < "header1.xml")
    let footer_slot = (footer_box.0 - HEADER_FOOTER_TEXTBOX_BASE) / PART_SLOT;
    let header_slot = (header_box.0 - HEADER_FOOTER_TEXTBOX_BASE) / PART_SLOT;
    assert_eq!(footer_slot, 0);
    assert_eq!(header_slot, 1);

    m.checkpoint(None);
    m.replace_paragraph_range(header_box.0, 0, header_box.1.chars().count(), "Encabezado").unwrap();
    m.replace_paragraph_range(footer_box.0, 0, footer_box.1.chars().count(), "Pie").unwrap();

    assert!(part_text(&m, "word/header1.xml").contains("Encabezado"));
    assert!(!part_text(&m, "word/header1.xml").contains("{{footer_box}}"));
    assert!(part_text(&m, "word/footer1.xml").contains("Pie"));
    assert!(!part_text(&m, "word/footer1.xml").contains("{{header_box}}"));

    // Top-level header/footer paragraphs still resolve to their own base/slot (unaffected)
    let hf_slot_base = HEADER_FOOTER_BASE;
    assert!(header_box.0 >= HEADER_FOOTER_TEXTBOX_BASE);
    assert!(hf_slot_base < HEADER_FOOTER_TEXTBOX_BASE);
}

// ---------- 5. Mixing a header text box with other parts is refused, like the body's ----------

#[test]
fn test_header_text_box_cannot_mix_with_the_body_or_a_table() {
    let header = format!(r#"<w:p>{}</w:p>"#, text_box("{{box}}", 1));
    let mut m = docx(&[("header1.xml", "rIdH", "default", "header", &header)]);
    let index = box_texts(&m, "header")[0].0;

    assert!(m.replace_range(0, 0, index, 0, "x").is_err(), "body -> header text box selection refused");
    let table = rust_web_docx::docx_parser::NewTable { rows: vec![vec!["a".into()]], header: false, widths: None, align: None };
    assert!(m.insert_table(index, 0, &table).is_err(), "cannot insert a table inside a header's text box");
    assert_eq!(box_texts(&m, "header")[0].1, "{{box}}", "nothing changed");
}

// ---------- 6. A text box anchored in a header/footer can be selected and dragged, like one in
// the body (bug: render_part used to null paragraph_index/image_index for every header/footer
// Image/Shape, and update_image/delete_image only ever touched document.xml) ----------

/// (paragraph index, image index) of the header's own drawing-bearing paragraph (the outer
/// paragraph holding the `w:drawing`, numbered like any other header paragraph — not the text
/// box's inner paragraph, which `box_texts` exposes)
fn header_drawing(m: &DocxModifier) -> (usize, usize) {
    let inputs = m.layout_inputs().unwrap();
    inputs
        .header_footer
        .header
        .iter()
        .find_map(|e| match e {
            DocumentElement::Paragraph(p) if !p.images.is_empty() => Some((p.index, 0)),
            _ => None,
        })
        .expect("the header has a paragraph with a drawing")
}

#[test]
fn test_anchored_text_box_in_a_header_gets_a_real_index_in_the_layout() {
    let header = format!(r#"<w:p>{}</w:p>"#, text_box("Logo", 1));
    let m = docx(&[("header1.xml", "rIdH", "default", "header", &header)]);

    let elements: Vec<DocumentElement> = m.extract_elements().unwrap();
    let inputs = m.layout_inputs().unwrap();
    let layout = LayoutEngine::new().with_images(inputs.body_images.clone()).compute_layout(
        &elements,
        "FFFFFF",
        &inputs.page_setup,
        &inputs.header_footer,
        None,
        None,
        0.0,
    );

    let shapes: Vec<_> = layout.pages[0]
        .items
        .iter()
        .filter_map(|it| match it {
            RenderCommand::Shape { paragraph_index, image_index, .. } => Some((*paragraph_index, *image_index)),
            _ => None,
        })
        .collect();
    assert!(!shapes.is_empty(), "the header's text box is drawn");
    assert!(
        shapes.iter().any(|(p, i)| p.is_some() && i.is_some()),
        "the header's floating text box must keep a real paragraph_index/image_index, like a body picture, so the canvas \
         editor's imageAt() can find it under a click: {:?}",
        shapes
    );
}

#[test]
fn test_update_image_on_a_header_text_box_moves_it_in_the_header_part() {
    let header = format!(r#"<w:p>{}</w:p>"#, text_box("Logo", 1));
    let mut m = docx(&[("header1.xml", "rIdH", "default", "header", &header)]);
    let document_before = part_text(&m, "word/document.xml");

    let (paragraph, image) = header_drawing(&m);
    m.checkpoint(None);
    let update = ImageUpdate { h_relative: Some("page".into()), h_offset: Some(500.0), ..Default::default() };
    let updated = m.update_image(paragraph, image, &update).unwrap();
    assert!(updated.anchored);
    assert_eq!((updated.h_relative.as_str(), updated.h_offset), ("page", 500.0));

    let header_xml = part_text(&m, "word/header1.xml");
    assert!(header_xml.contains(r#"relativeFrom="page""#), "the new position landed in header1.xml: {}", header_xml);
    assert!(header_xml.contains("<wp:posOffset>4762500</wp:posOffset>"), "500px in EMU: {}", header_xml);
    assert_eq!(part_text(&m, "word/document.xml"), document_before, "document.xml must stay untouched");

    // The document reopens without error and the new position survives the round trip
    let reopened = DocxModifier::from_bytes(&m.to_bytes().unwrap()).unwrap();
    assert!(part_text(&reopened, "word/header1.xml").contains(r#"relativeFrom="page""#));

    assert!(m.undo().is_some(), "the move is one undo step");
    assert!(!part_text(&m, "word/header1.xml").contains(r#"relativeFrom="page""#), "undo restores the header");
}

// ---------- 7. No regression: moving a BODY picture still writes to document.xml ----------

#[test]
fn test_update_image_on_a_body_picture_still_writes_to_document_xml() {
    let mut m = docx(&[]);
    m.checkpoint(None);
    let (p, _) = m.insert_image(0, 0, &png(40, 20), &NewImage::default()).unwrap();

    m.checkpoint(None);
    let updated = m
        .update_image(p, 0, &ImageUpdate { wrap: Some("square".into()), h_offset: Some(50.0), ..Default::default() })
        .unwrap();
    assert!(updated.anchored);
    assert_eq!(updated.h_offset, 50.0);
    assert!(part_text(&m, "word/document.xml").contains("<wp:posOffset>476250</wp:posOffset>"), "50px in EMU, in document.xml");
}
