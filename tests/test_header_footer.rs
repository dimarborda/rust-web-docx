use rust_web_docx::docx_parser::{DocumentElement, DocxModifier};
use rust_web_docx::layout_engine::{DocumentLayout, LayoutEngine, RenderCommand};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

/// A .docx whose last section uses the given header/footer parts. `parts`: (file name, rel id,
/// reference type, kind, xml body)
fn docx(body: &str, title_page: bool, parts: &[(&str, &str, &str, &str, &str)]) -> DocxModifier {
    let refs: String = parts
        .iter()
        .map(|(_, id, ty, kind, _)| format!(r#"<w:{}Reference w:type="{}" r:id="{}"/>"#, kind, ty, id))
        .collect();
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}<w:sectPr>{refs}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>{tp}</w:sectPr></w:body></w:document>"#,
        tp = if title_page { "<w:titlePg/>" } else { "" }
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

fn layout(m: &DocxModifier) -> DocumentLayout {
    let elements: Vec<DocumentElement> = m.extract_elements().unwrap();
    let inputs = m.layout_inputs().unwrap();
    LayoutEngine::new().with_images(inputs.body_images.clone()).compute_layout(
        &elements, "FFFFFF", &inputs.page_setup, &inputs.header_footer, None, None, 0.0,
    )
}

/// (text, x, baseline y, belongs to the body) of page `n` (0-based)
fn texts(l: &DocumentLayout, n: usize) -> Vec<(String, f64, f64, bool)> {
    l.pages[n]
        .items
        .iter()
        .filter_map(|it| match it {
            RenderCommand::Text { text, x, y, line, .. } => Some((text.clone(), *x, *y, line.is_some())),
            _ => None,
        })
        .collect()
}

fn find(l: &DocumentLayout, n: usize, needle: &str) -> Option<(String, f64, f64, bool)> {
    texts(l, n).into_iter().find(|t| t.0.contains(needle))
}

const FOOTER_PAGES: &str = concat!(
    r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Página </w:t></w:r>"#,
    r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>"#,
    r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
    r#"<w:r><w:t xml:space="preserve"> de </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>"#
);

fn pages_of_text(n: usize) -> String {
    (0..n).map(|i| format!(r#"<w:p><w:r><w:t>Párrafo {}</w:t></w:r></w:p>"#, i)).collect()
}

#[test]
fn test_header_and_footer_text_with_page_numbers() {
    let header = r#"<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>{{html:signer2_unhtml}}</w:t></w:r></w:p>"#;
    let m = docx(&pages_of_text(90), false, &[("header1.xml", "rIdH", "default", "header", header), ("footer1.xml", "rIdF", "default", "footer", FOOTER_PAGES)]);
    let fields = &m.layout_inputs().unwrap().header_footer.footer[0];
    match fields {
        DocumentElement::Paragraph(p) => {
            assert_eq!(p.text, "Página 1 de 9");
            assert_eq!(p.fields.len(), 2, "PAGE and NUMPAGES found: {:?}", p.fields);
        }
        _ => panic!(),
    }
    let l = layout(&m);
    assert!(l.total_pages >= 2);
    for n in 0..l.total_pages {
        let h = find(&l, n, "signer2_unhtml").expect("header on every page");
        assert!(!h.3, "header text is not part of the body (no caret there)");
        assert!(h.2 < 96.0, "header sits in the top margin: {}", h.2);
        assert!(h.1 > 400.0, "right aligned: {}", h.1);
        let f = find(&l, n, "Página").expect("footer on every page");
        assert_eq!(f.0, format!("Página {} de {}", n + 1, l.total_pages));
        assert!(f.2 > l.page_height - 96.0, "footer in the bottom margin: {}", f.2);
    }
}

#[test]
fn test_tall_header_pushes_the_body_down() {
    let tall: String = (0..8).map(|i| format!(r#"<w:p><w:r><w:t>Línea de encabezado {}</w:t></w:r></w:p>"#, i)).collect();
    let short = docx(&pages_of_text(3), false, &[]);
    let with = docx(&pages_of_text(3), false, &[("header1.xml", "rIdH", "default", "header", &tall)]);
    let body_top = |m: &DocxModifier| find(&layout(m), 0, "Párrafo 0").unwrap().2;
    let last_header = find(&layout(&with), 0, "encabezado 7").unwrap().2;
    assert!(body_top(&with) > last_header, "the body starts below the header");
    assert!(body_top(&with) > body_top(&short) + 20.0);
}

#[test]
fn test_header_table_and_first_page() {
    let table = concat!(
        r#"<w:tbl><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr>"#,
        r#"<w:tc><w:p><w:r><w:t>Empresa S.A.S.</w:t></w:r></w:p></w:tc><w:tc><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>NIT 900.123</w:t></w:r></w:p></w:tc>"#,
        r#"</w:tr></w:tbl><w:p/>"#
    );
    let first = r#"<w:p><w:r><w:t>Portada</w:t></w:r></w:p>"#;
    let m = docx(&pages_of_text(90), true, &[
        ("header1.xml", "rIdH", "default", "header", table),
        ("header2.xml", "rIdH1", "first", "header", first),
    ]);
    let l = layout(&m);
    assert!(find(&l, 0, "Portada").is_some(), "first page uses the first-page header");
    assert!(find(&l, 0, "Empresa").is_none());
    let (company, nit) = (find(&l, 1, "Empresa").unwrap(), find(&l, 1, "NIT").unwrap());
    assert!((company.2 - nit.2).abs() < 0.5 && nit.1 > company.1 + 200.0, "two cells side by side");
    assert!(find(&l, 1, "Portada").is_none());
}
