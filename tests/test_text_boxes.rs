use rust_web_docx::docx_parser::{DocumentElement, DocxModifier, ImageRef, ImageUpdate};
use rust_web_docx::layout_engine::{DocumentLayout, LayoutEngine, RenderCommand};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
    r#"xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" "#,
    r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
    r#"xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word""#
);

const THEME: &str = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="1F6F6B"/></a:accent1></a:clrScheme></a:themeElements></a:theme>"#;

/// Minimal .docx (Letter, 1" margins) around the given body XML, with a theme
fn docx(body: &str) -> DocxModifier {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {}><w:body>{}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#,
        NS, body
    );
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        let opts = SimpleFileOptions::default();
        for (name, content) in [
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#),
            ("word/document.xml", document.as_str()),
            ("word/theme/theme1.xml", THEME),
        ] {
            zip.start_file(name, opts).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
    }
    DocxModifier::from_bytes(&buf.into_inner()).unwrap()
}

/// A text box as Word 2016+ writes it: DrawingML in mc:Choice, the same box in VML as fallback
fn word_text_box(fill: &str, text: &str) -> String {
    format!(
        concat!(
            r#"<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>"#,
            r#"<wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">"#,
            r#"<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>"#,
            r#"<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>"#,
            r#"<wp:extent cx="1905000" cy="952500"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>"#,
            r#"<wp:docPr id="217" name="Cuadro de texto 2"/><wp:cNvGraphicFramePr/>"#,
            r#"<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>"#,
            r#"<wps:cNvSpPr txBox="1"/><wps:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm>"#,
            r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill}"#,
            r#"<a:ln w="19050"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></wps:spPr>"#,
            r#"<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>{text}</w:t></w:r></w:p>"#,
            r#"<w:p><w:r><w:t>Segunda línea</w:t></w:r></w:p></w:txbxContent></wps:txbx>"#,
            r#"<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="ctr"><a:noAutofit/></wps:bodyPr>"#,
            r#"</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>"#,
            r#"<mc:Fallback><w:pict><v:shape style="position:absolute;margin-left:0;margin-top:0;width:150pt;height:75pt;z-index:251659264" stroked="t">"#,
            r#"<v:textbox><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></v:textbox><w10:wrap type="square"/></v:shape></w:pict></mc:Fallback>"#,
            r#"</mc:AlternateContent></w:r>"#
        ),
        fill = fill,
        text = text
    )
}

const LONG: &str = "Este párrafo es largo para que el texto tenga que rodear el cuadro de texto que está anclado en él, igual que en Word, línea tras línea hasta pasar el cuadro.";

fn paragraphs(m: &DocxModifier) -> Vec<rust_web_docx::docx_parser::ParagraphInfo> {
    m.extract_paragraphs().unwrap()
}

fn layout(m: &DocxModifier) -> DocumentLayout {
    let elements: Vec<DocumentElement> = m.extract_elements().unwrap();
    let inputs = m.layout_inputs().unwrap();
    LayoutEngine::new().with_images(inputs.body_images.clone()).compute_layout(
        &elements,
        "FFFFFF",
        &inputs.page_setup,
        &inputs.header_footer,
        None,
        None,
        0.0,
    )
}

/// (x, y, width, height, fill, stroke, paragraph index) of each shape drawn on page 1
type DrawnShape = (f64, f64, f64, f64, Option<String>, Option<String>, Option<usize>);

fn shapes(l: &DocumentLayout) -> Vec<DrawnShape> {
    l.pages[0]
        .items
        .iter()
        .filter_map(|it| match it {
            RenderCommand::Shape { x, y, width, height, fill, stroke, paragraph_index, .. } => {
                Some((*x, *y, *width, *height, fill.clone(), stroke.clone(), *paragraph_index))
            }
            _ => None,
        })
        .collect()
}

/// Text items: (text, x, baseline y, belongs to the body). Text box text is not the body's,
/// whether its caret lines are kept (editable box) or not.
fn texts(l: &DocumentLayout) -> Vec<(String, f64, f64, bool)> {
    l.pages[0]
        .items
        .iter()
        .filter_map(|it| match it {
            RenderCommand::Text { text, x, y, line, .. } => {
                Some((text.clone(), *x, *y, line.as_ref().is_some_and(|r| r.text_box.is_none())))
            }
            _ => None,
        })
        .collect()
}

#[test]
fn test_word_text_box_is_read_once_with_its_text() {
    let fill = r#"<a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill>"#;
    let m = docx(&format!(r#"<w:p>{}<w:r><w:t>{}</w:t></w:r></w:p>"#, word_text_box(fill, "Nota importante"), LONG));
    let ps = paragraphs(&m);
    assert_eq!(ps.len(), 1, "the box's paragraphs are not body paragraphs");
    assert_eq!(ps[0].text, LONG, "the box's text is not part of the paragraph");
    assert_eq!(ps[0].images.len(), 1, "the VML fallback is not read again");
    let img: &ImageRef = &ps[0].images[0];
    assert!(!img.vml);
    assert_eq!(img.wrap, "square");
    assert_eq!((img.width, img.height), (200.0, 100.0));
    let shape = img.shape.as_ref().unwrap();
    assert_eq!(shape.fill.as_deref(), Some("FFF2CC"));
    assert_eq!(shape.stroke.as_deref(), Some("C00000"));
    assert_eq!(shape.stroke_width, 2.0);
    let tb = img.text_box.as_ref().unwrap();
    assert_eq!(tb.text(), "Nota importante\nSegunda línea");
    assert_eq!(tb.v_anchor, "ctr");
    assert_eq!((tb.inset_left, tb.inset_top), (9.6, 4.8));
    assert!(tb.paragraphs[0].runs[0].bold);
    assert_eq!(tb.paragraphs[0].align, "center");
}

#[test]
fn test_text_box_is_drawn_and_text_wraps_around_it() {
    let fill = r#"<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></a:solidFill>"#;
    let m = docx(&format!(r#"<w:p>{}<w:r><w:t>{}</w:t></w:r></w:p>"#, word_text_box(fill, "Nota"), LONG));
    let l = layout(&m);
    let s = shapes(&l);
    assert_eq!(s.len(), 1);
    let (x, y, w, h, fill, stroke, para) = s[0].clone();
    // Right-aligned within the margins: 8.5in page, 1in margins
    assert!((x + w - (816.0 - 96.0)).abs() < 0.5, "right edge on the right margin: {}", x + w);
    assert_eq!((w, h), (200.0, 100.0));
    assert!(fill.is_some_and(|c| c.starts_with('#')), "theme color resolved");
    assert_eq!(stroke.as_deref(), Some("#C00000"));
    assert_eq!(para, Some(0), "selectable like a picture");

    let all = texts(&l);
    let inside: Vec<_> = all.iter().filter(|t| !t.3).collect();
    assert!(inside.iter().any(|t| t.0 == "Nota"), "box text is drawn: {:?}", inside);
    for t in &inside {
        assert!(t.1 >= x && t.1 <= x + w, "box text inside the box horizontally: {:?}", t);
        assert!(t.2 > y && t.2 < y + h, "box text inside the box vertically: {:?}", t);
    }
    // Centered vertically (anchor="ctr"): the text block sits around the box's middle
    let mid = inside.iter().map(|t| t.2).sum::<f64>() / inside.len() as f64;
    assert!((mid - (y + h / 2.0)).abs() < 14.0, "vertical center {} vs {}", mid, y + h / 2.0);

    // Body lines beside the box end before it (plus its 0.125" distance)
    let body: Vec<_> = l.pages[0]
        .items
        .iter()
        .filter_map(|it| match it {
            RenderCommand::Text { line: Some(r), .. } if r.top < y + h && r.text_box.is_none() => Some(r.right),
            _ => None,
        })
        .collect();
    assert!(!body.is_empty());
    assert!(body.iter().all(|r| *r <= x - 11.9), "body text keeps clear of the box: {:?} vs {}", body, x);
}

#[test]
fn test_legacy_vml_text_box() {
    let vml = concat!(
        r#"<w:p><w:r><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600"/>"#,
        r##"<v:shape id="Caja" type="#_x0000_t202" style="position:absolute;margin-left:36pt;margin-top:6pt;width:150pt;height:60pt;z-index:251659264;mso-position-horizontal-relative:margin" fillcolor="#DDEEFF" strokecolor="red" strokeweight="1.5pt">"##,
        r#"<v:textbox inset="3.6pt,3.6pt,3.6pt,3.6pt" style="v-text-anchor:bottom"><w:txbxContent><w:p><w:r><w:t>Caja antigua</w:t></w:r></w:p></w:txbxContent></v:textbox>"#,
        r#"<w10:wrap type="topAndBottom"/></v:shape></w:pict></w:r><w:r><w:t>Texto</w:t></w:r></w:p>"#
    );
    let m = docx(vml);
    let ps = paragraphs(&m);
    assert_eq!(ps[0].text, "Texto");
    let img = &ps[0].images[0];
    assert!(img.vml && img.anchored);
    assert_eq!(img.wrap, "topAndBottom");
    assert_eq!((img.width, img.height), (200.0, 80.0));
    assert_eq!((img.h_relative.as_str(), img.h_offset, img.v_offset), ("margin", 48.0, 8.0));
    let shape = img.shape.as_ref().unwrap();
    assert_eq!((shape.fill.as_deref(), shape.stroke.as_deref(), shape.stroke_width), (Some("DDEEFF"), Some("FF0000"), 2.0));
    let tb = img.text_box.as_ref().unwrap();
    assert_eq!((tb.text().as_str(), tb.v_anchor.as_str()), ("Caja antigua", "b"));
    assert!((tb.inset_left - 4.8).abs() < 1e-9);

    let l = layout(&m);
    let s = shapes(&l);
    assert_eq!(s.len(), 1);
    assert_eq!(s[0].6, None, "VML shapes cannot be selected");
    // Top-and-bottom: the paragraph's text goes below the box
    let body_y = texts(&l).into_iter().find(|t| t.3 && t.0 == "Texto").unwrap().2;
    assert!(body_y > s[0].1 + s[0].3, "text below the box: {} vs {}", body_y, s[0].1 + s[0].3);
}

#[test]
fn test_text_box_moves_and_resizes_like_a_picture() {
    let fill = r#"<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>"#;
    let mut m = docx(&format!(r#"<w:p>{}<w:r><w:t>Hola</w:t></w:r></w:p>"#, word_text_box(fill, "Caja")));
    m.checkpoint(None);
    let img = m
        .update_image(0, 0, &ImageUpdate { width: Some(300.0), h_relative: Some("page".into()), h_offset: Some(50.0), wrap: Some("topAndBottom".into()), ..Default::default() })
        .unwrap();
    assert_eq!((img.width, img.height), (300.0, 150.0));
    assert_eq!(img.wrap, "topAndBottom");
    assert_eq!(img.text_box.as_ref().unwrap().text(), "Caja\nSegunda línea", "the text survives");
    let doc = m.get_file_string("word/document.xml").unwrap();
    assert!(doc.contains(r#"<a:ext cx="2857500" cy="1428750"/>"#), "the shape's own size follows");
    assert!(doc.contains("<mc:Fallback>"), "the rest of the run is kept");
    // Editing the paragraph's text keeps the box
    m.replace_paragraph_range(0, 0, 0, "¡").unwrap();
    assert_eq!(paragraphs(&m)[0].text, "¡Hola");
    assert_eq!(paragraphs(&m)[0].images.len(), 1);
}

#[test]
fn test_shapes_without_text_and_groups() {
    let rect = concat!(
        r#"<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="476250"/><wp:docPr id="5" name="Elipse"/>"#,
        r#"<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr>"#,
        r#"<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>"#,
        r#"<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#,
        r#"<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="6" name="Grupo"/>"#,
        r#"<a:graphic><a:graphicData uri="g"><wpg:wgp xmlns:wpg="g"><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>en grupo</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></wpg:wgp>"#,
        r#"</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#
    );
    let m = docx(rect);
    let ps = paragraphs(&m);
    let ellipse = &ps[0].images[0];
    assert_eq!(ellipse.wrap, "inline");
    let s = ellipse.shape.as_ref().unwrap();
    assert_eq!((s.geometry.as_str(), s.fill.as_deref(), s.stroke.as_deref()), ("ellipse", Some("1F6F6B"), None));
    assert!(ellipse.text_box.is_none());
    assert!(ps[1].images.is_empty(), "groups are not drawn yet (nor mistaken for a single box)");
    let l = layout(&m);
    assert_eq!(shapes(&l).len(), 1);
}

// ---------- Editing the text of a text box ----------

use rust_web_docx::caret::{self, TextPosition};
use rust_web_docx::docx_parser::{NewTable, RunInfo, TEXT_BOX_BASE};
use rust_web_docx::layout_engine::EstimateMeasurer;

fn box_texts(m: &DocxModifier) -> Vec<(usize, String)> {
    paragraphs(m)
        .into_iter()
        .flat_map(|p| p.images.into_iter())
        .filter_map(|img| img.text_box)
        .flat_map(|tb| tb.paragraphs.into_iter().map(|p| (p.index, p.text)))
        .collect()
}

fn editable_doc() -> DocxModifier {
    let fill = r#"<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>"#;
    docx(&format!(
        r#"<w:p>{}<w:r><w:t>Cuerpo uno</w:t></w:r></w:p><w:p>{}<w:r><w:t>Cuerpo dos</w:t></w:r></w:p>"#,
        word_text_box(fill, "Primera caja"),
        word_text_box(fill, "Otra caja"),
    ))
}

#[test]
fn test_text_box_paragraphs_have_their_own_numbers() {
    let m = editable_doc();
    let ps = paragraphs(&m);
    assert_eq!(ps.iter().map(|p| p.index).collect::<Vec<_>>(), vec![0, 1], "body numbering is unchanged");
    assert_eq!(
        box_texts(&m),
        vec![
            (TEXT_BOX_BASE, "Primera caja".to_string()),
            (TEXT_BOX_BASE + 1, "Segunda línea".to_string()),
            (TEXT_BOX_BASE + 2, "Otra caja".to_string()),
            (TEXT_BOX_BASE + 3, "Segunda línea".to_string()),
        ]
    );
}

#[test]
fn test_typing_enter_and_backspace_inside_a_text_box() {
    let mut m = editable_doc();
    m.checkpoint(None);
    let caret = m.replace_range(TEXT_BOX_BASE, 0, TEXT_BOX_BASE, 0, "¡").unwrap();
    assert_eq!(caret, (TEXT_BOX_BASE, 1));
    assert_eq!(box_texts(&m)[0].1, "¡Primera caja");
    assert_eq!(paragraphs(&m)[0].text, "Cuerpo uno", "the anchor paragraph keeps its text");

    // Enter in the middle of the box's first paragraph: the box grows, the next box shifts
    let caret = m.replace_range(TEXT_BOX_BASE, 8, TEXT_BOX_BASE, 8, "\n").unwrap();
    assert_eq!(caret, (TEXT_BOX_BASE + 1, 0));
    let texts: Vec<String> = box_texts(&m).into_iter().map(|t| t.1).collect();
    assert_eq!(texts, vec!["¡Primera", " caja", "Segunda línea", "Otra caja", "Segunda línea"]);
    assert_eq!(paragraphs(&m).len(), 2, "no body paragraph was added");

    // Backspace at the start of the new paragraph joins it back
    m.replace_range(TEXT_BOX_BASE, 8, TEXT_BOX_BASE + 1, 0, "").unwrap();
    assert_eq!(box_texts(&m)[0].1, "¡Primera caja");
    assert_eq!(box_texts(&m).len(), 4);

    // The second box is edited independently
    m.replace_paragraph_range(TEXT_BOX_BASE + 2, 0, 4, "Esta").unwrap();
    assert_eq!(box_texts(&m)[2].1, "Esta caja");

    // Survives saving and reopening
    let reopened = DocxModifier::from_bytes(&m.to_bytes().unwrap()).unwrap();
    assert_eq!(box_texts(&reopened)[2].1, "Esta caja");
    assert_eq!(paragraphs(&reopened)[1].text, "Cuerpo dos");
}

#[test]
fn test_formatting_and_undo_inside_a_text_box() {
    let mut m = editable_doc();
    m.checkpoint(None);
    let runs = vec![RunInfo { text: "Primera".into(), bold: true, ..Default::default() }, RunInfo { text: " caja".into(), ..Default::default() }];
    m.update_paragraph_runs(TEXT_BOX_BASE, &runs, Some("right")).unwrap();
    let tb = paragraphs(&m)[0].images[0].text_box.clone().unwrap();
    assert!(tb.paragraphs[0].runs[0].bold);
    assert_eq!(tb.paragraphs[0].align, "right");
    assert!(m.undo().is_some());
    let tb = paragraphs(&m)[0].images[0].text_box.clone().unwrap();
    assert_eq!(tb.paragraphs[0].align, "center", "undo restores the box");
}

#[test]
fn test_mixed_selections_and_tables_are_refused() {
    let mut m = editable_doc();
    assert!(m.replace_range(0, 2, TEXT_BOX_BASE, 3, "x").is_err(), "body → text box selection");
    let table = NewTable { rows: vec![vec!["a".into()]], header: false, widths: None, align: None };
    assert!(m.insert_table(TEXT_BOX_BASE, 0, &table).is_err());
    assert_eq!(box_texts(&m)[0].1, "Primera caja", "nothing changed");
}

#[test]
fn test_clicks_inside_a_text_box_reach_its_text() {
    let m = editable_doc();
    let l = layout(&m);
    let (x, y, w, h, ..) = shapes(&l)[0].clone();
    let inside = caret::hit_test(&l, 1, x + w / 2.0, y + h / 2.0, &mut EstimateMeasurer).unwrap();
    // Both boxes overlap here (right-aligned on one-line paragraphs): the topmost one wins
    assert!(inside.paragraph >= TEXT_BOX_BASE, "{:?}", inside);
    let outside = caret::hit_test(&l, 1, 100.0, y + 5.0, &mut EstimateMeasurer).unwrap();
    assert!(outside.paragraph < TEXT_BOX_BASE, "{:?}", outside);
    // The caret can be shown in the box and moving down stays in it
    let pos = TextPosition { paragraph: TEXT_BOX_BASE, offset: 2 };
    assert!(caret::caret_box(&l, pos, &mut EstimateMeasurer).is_some());
    let down = caret::move_vertical(&l, pos, 1, x + 20.0, &mut EstimateMeasurer).unwrap();
    assert_eq!(down.paragraph, TEXT_BOX_BASE + 1);
}
