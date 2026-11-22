//! Headers and footers that change by section, on each section's first page and on even
//! pages, as Word shows them.

use rust_web_docx::docx_parser::{DocumentElement, DocxModifier, HEADER_FOOTER_BASE, PART_SLOT};
use rust_web_docx::layout_engine::{DocumentLayout, LayoutEngine, RenderCommand};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;

/// One section of a test document
struct Section<'a> {
    /// (kind "header"/"footer", type "default"/"first"/"even", relationship id)
    refs: &'a [(&'a str, &'a str, &'a str)],
    /// How the section starts: "nextPage" or "continuous"
    start: &'a str,
    title_page: bool,
    body: String,
}

fn section<'a>(refs: &'a [(&'a str, &'a str, &'a str)], body: String) -> Section<'a> {
    Section {
        refs,
        start: "nextPage",
        title_page: false,
        body,
    }
}

fn sect_pr(s: &Section) -> String {
    let refs: String = s
        .refs
        .iter()
        .map(|(kind, ty, id)| format!(r#"<w:{kind}Reference w:type="{ty}" r:id="{id}"/>"#))
        .collect();
    let start = if s.start == "nextPage" {
        String::new()
    } else {
        format!(r#"<w:type w:val="{}"/>"#, s.start)
    };
    format!(
        r#"<w:sectPr>{refs}{start}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>{tp}</w:sectPr>"#,
        tp = if s.title_page { "<w:titlePg/>" } else { "" }
    )
}

/// A .docx with these sections. `parts`: (file name, relationship id, kind, text of its one
/// paragraph per line); `even_and_odd` writes `w:evenAndOddHeaders` in settings.xml
fn docx(
    sections: &[Section],
    parts: &[(&str, &str, &str, &str)],
    even_and_odd: bool,
) -> DocxModifier {
    let mut body = String::new();
    for (i, s) in sections.iter().enumerate() {
        body.push_str(&s.body);
        if i + 1 < sections.len() {
            body.push_str(&format!("<w:p><w:pPr>{}</w:pPr></w:p>", sect_pr(s)));
        } else {
            body.push_str(&sect_pr(s));
        }
    }
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}</w:body></w:document>"#
    );
    let rels: String = parts
        .iter()
        .map(|(file, id, kind, _)| {
            format!(r#"<Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{file}"/>"#)
        })
        .collect();
    let mut files = vec![
        ("[Content_Types].xml".to_string(), r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#.to_string()),
        ("word/document.xml".to_string(), document),
        (
            "word/_rels/document.xml.rels".to_string(),
            format!(r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#),
        ),
    ];
    if even_and_odd {
        files.push((
            "word/settings.xml".to_string(),
            format!(r#"<?xml version="1.0"?><w:settings {W}><w:evenAndOddHeaders/></w:settings>"#),
        ));
    }
    for (file, _, kind, text) in parts {
        let root = if *kind == "header" { "hdr" } else { "ftr" };
        let paragraphs: String = text
            .split('\n')
            .map(|line| format!("<w:p><w:r><w:t>{line}</w:t></w:r></w:p>"))
            .collect();
        files.push((
            format!("word/{file}"),
            format!(r#"<?xml version="1.0"?><w:{root} {W}>{paragraphs}</w:{root}>"#),
        ));
    }
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        for (name, content) in files {
            zip.start_file(name, SimpleFileOptions::default()).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
    }
    DocxModifier::from_bytes(&buf.into_inner()).unwrap()
}

fn layout(m: &DocxModifier) -> DocumentLayout {
    let elements: Vec<DocumentElement> = m.extract_elements().unwrap();
    let inputs = m.layout_inputs().unwrap();
    LayoutEngine::new()
        .with_images(inputs.body_images.clone())
        .compute_layout(
            &elements,
            "FFFFFF",
            &inputs.page_setup,
            &inputs.header_footer,
            None,
            None,
            0.0,
        )
}

/// Body paragraphs, the page break between pages: `pages(&["Uno", "Dos"])` fills two pages
fn pages(texts: &[&str]) -> String {
    texts
        .iter()
        .enumerate()
        .map(|(i, t)| {
            let br = if i == 0 {
                ""
            } else {
                r#"<w:r><w:br w:type="page"/></w:r>"#
            };
            format!("<w:p>{br}<w:r><w:t>{t}</w:t></w:r></w:p>")
        })
        .collect()
}

/// Header (top half) or footer (bottom half) text of page `n`
fn part_on_page(l: &DocumentLayout, n: usize, header: bool) -> String {
    let half = l.page_height / 2.0;
    l.pages[n]
        .items
        .iter()
        .filter_map(|it| match it {
            RenderCommand::Text {
                text,
                y,
                paragraph_index,
                ..
            } if *paragraph_index >= HEADER_FOOTER_BASE && (*y < half) == header => {
                Some(text.trim().to_string())
            }
            _ => None,
        })
        .collect::<Vec<_>>()
        .join(" | ")
}

fn headers(l: &DocumentLayout) -> Vec<String> {
    (0..l.pages.len())
        .map(|n| part_on_page(l, n, true))
        .collect()
}

fn footers(l: &DocumentLayout) -> Vec<String> {
    (0..l.pages.len())
        .map(|n| part_on_page(l, n, false))
        .collect()
}

/// Top of the first body line on page `n`
fn body_top(l: &DocumentLayout, n: usize) -> f64 {
    l.pages[n]
        .items
        .iter()
        .filter_map(|it| match it {
            RenderCommand::Text {
                y,
                paragraph_index,
                line: Some(_),
                ..
            } if *paragraph_index < HEADER_FOOTER_BASE => Some(*y),
            _ => None,
        })
        .fold(f64::MAX, f64::min)
}

#[test]
fn test_each_section_shows_its_own_header_and_footer() {
    let m = docx(
        &[
            section(
                &[("header", "default", "rH1"), ("footer", "default", "rF1")],
                pages(&["Uno"]),
            ),
            section(
                &[("header", "default", "rH2"), ("footer", "default", "rF2")],
                pages(&["Dos"]),
            ),
        ],
        &[
            ("header1.xml", "rH1", "header", "Cabecera A"),
            ("footer1.xml", "rF1", "footer", "Pie A"),
            ("header2.xml", "rH2", "header", "Cabecera B"),
            ("footer2.xml", "rF2", "footer", "Pie B"),
        ],
        false,
    );
    let l = layout(&m);
    assert_eq!(l.pages.len(), 2);
    assert_eq!(headers(&l), ["Cabecera A", "Cabecera B"]);
    assert_eq!(footers(&l), ["Pie A", "Pie B"]);
}

#[test]
fn test_a_section_without_references_keeps_the_previous_ones() {
    let m = docx(
        &[
            section(
                &[("header", "default", "rH1"), ("footer", "default", "rF1")],
                pages(&["Uno"]),
            ),
            // "Link to previous" for both
            section(&[], pages(&["Dos"])),
            // A new header, the footer still linked
            section(&[("header", "default", "rH3")], pages(&["Tres"])),
        ],
        &[
            ("header1.xml", "rH1", "header", "Cabecera A"),
            ("footer1.xml", "rF1", "footer", "Pie A"),
            ("header3.xml", "rH3", "header", "Cabecera C"),
        ],
        false,
    );
    let l = layout(&m);
    assert_eq!(headers(&l), ["Cabecera A", "Cabecera A", "Cabecera C"]);
    assert_eq!(footers(&l), ["Pie A", "Pie A", "Pie A"]);
}

#[test]
fn test_different_first_page_applies_to_each_section_on_its_own() {
    let mut second = section(
        &[("header", "default", "rH2"), ("header", "first", "rP2")],
        pages(&["Dos", "Dos bis"]),
    );
    second.title_page = true;
    let m = docx(
        &[
            section(&[("header", "default", "rH1")], pages(&["Uno", "Uno bis"])),
            second,
        ],
        &[
            ("header1.xml", "rH1", "header", "Cabecera A"),
            ("header2.xml", "rH2", "header", "Cabecera B"),
            ("headerp.xml", "rP2", "header", "Portada B"),
        ],
        false,
    );
    let l = layout(&m);
    assert_eq!(
        headers(&l),
        ["Cabecera A", "Cabecera A", "Portada B", "Cabecera B"]
    );
}

#[test]
fn test_even_pages_use_the_even_header_when_settings_ask_for_it() {
    let parts = [
        ("header1.xml", "rH", "header", "Impar"),
        ("header2.xml", "rE", "header", "Par"),
        ("footer1.xml", "rF", "footer", "Pie impar"),
        ("footer2.xml", "rG", "footer", "Pie par"),
    ];
    let refs = [
        ("header", "default", "rH"),
        ("header", "even", "rE"),
        ("footer", "default", "rF"),
        ("footer", "even", "rG"),
    ];
    let body = pages(&["Uno", "Dos", "Tres", "Cuatro"]);
    let l = layout(&docx(&[section(&refs, body.clone())], &parts, true));
    assert_eq!(headers(&l), ["Impar", "Par", "Impar", "Par"]);
    assert_eq!(
        footers(&l),
        ["Pie impar", "Pie par", "Pie impar", "Pie par"]
    );

    // Without w:evenAndOddHeaders the even header is ignored, as in Word
    let l = layout(&docx(&[section(&refs, body)], &parts, false));
    assert_eq!(headers(&l), ["Impar", "Impar", "Impar", "Impar"]);
}

#[test]
fn test_first_page_wins_over_even_and_missing_even_header_is_blank() {
    let mut s = section(
        &[("header", "default", "rH"), ("header", "first", "rP")],
        pages(&["Uno", "Dos", "Tres"]),
    );
    s.title_page = true;
    let l = layout(&docx(
        &[s],
        &[
            ("header1.xml", "rH", "header", "Impar"),
            ("header2.xml", "rP", "header", "Portada"),
        ],
        true,
    ));
    // Page 2 is even and the section has no even header: Word leaves it blank
    assert_eq!(headers(&l), ["Portada", "", "Impar"]);
}

#[test]
fn test_continuous_section_changes_the_header_from_the_next_page() {
    let mut second = section(
        &[("header", "default", "rH2"), ("header", "first", "rP2")],
        pages(&["Dos", "Tres"]),
    );
    second.start = "continuous";
    second.title_page = true;
    let m = docx(
        &[
            section(&[("header", "default", "rH1")], pages(&["Uno"])),
            second,
        ],
        &[
            ("header1.xml", "rH1", "header", "Cabecera A"),
            ("header2.xml", "rH2", "header", "Cabecera B"),
            ("headerp.xml", "rP2", "header", "Portada B"),
        ],
        false,
    );
    let l = layout(&m);
    assert_eq!(l.pages.len(), 2);
    // Page 1 starts in the first section; the second starts there too, so it never shows
    // its first-page header
    assert_eq!(headers(&l), ["Cabecera A", "Cabecera B"]);
    let body_page_1: Vec<String> = l.pages[0]
        .items
        .iter()
        .filter_map(|it| match it {
            RenderCommand::Text {
                text,
                paragraph_index,
                ..
            } if *paragraph_index < HEADER_FOOTER_BASE => Some(text.clone()),
            _ => None,
        })
        .collect();
    assert!(body_page_1.iter().any(|t| t.contains("Uno")));
    assert!(body_page_1.iter().any(|t| t.contains("Dos")));
}

#[test]
fn test_a_tall_header_pushes_the_body_down_only_in_its_section() {
    let tall = "Línea 1\nLínea 2\nLínea 3\nLínea 4\nLínea 5\nLínea 6\nLínea 7\nLínea 8";
    let m = docx(
        &[
            section(&[("header", "default", "rH1")], pages(&["Uno", "Uno bis"])),
            section(&[("header", "default", "rH2")], pages(&["Dos", "Dos bis"])),
        ],
        &[
            ("header1.xml", "rH1", "header", "Corta"),
            ("header2.xml", "rH2", "header", tall),
        ],
        false,
    );
    let l = layout(&m);
    assert_eq!(l.pages.len(), 4);
    let tops: Vec<f64> = (0..4).map(|n| body_top(&l, n)).collect();
    assert!((tops[0] - tops[1]).abs() < 0.01, "{:?}", tops);
    assert!((tops[2] - tops[3]).abs() < 0.01, "{:?}", tops);
    assert!(
        tops[2] > tops[0] + 50.0,
        "the second section's body starts below its tall header: {:?}",
        tops
    );
}

#[test]
fn test_typing_in_the_second_sections_header_edits_its_part() {
    let mut m = docx(
        &[
            section(&[("header", "default", "rH1")], pages(&["Uno"])),
            section(&[("header", "default", "rH2")], pages(&["Dos"])),
        ],
        &[
            ("header1.xml", "rH1", "header", "Cabecera A"),
            ("header2.xml", "rH2", "header", "Cabecera B"),
        ],
        false,
    );
    // Parts are numbered in name order: header1.xml is slot 0, header2.xml slot 1
    let second = HEADER_FOOTER_BASE + PART_SLOT;
    let all: Vec<usize> = m
        .layout_inputs()
        .unwrap()
        .header_footer
        .all_parts()
        .into_iter()
        .flatten()
        .filter_map(|e| match e {
            DocumentElement::Paragraph(p) => Some(p.index),
            _ => None,
        })
        .collect();
    assert!(
        all.contains(&HEADER_FOOTER_BASE),
        "the first section's header is listed"
    );
    assert!(all.contains(&second));

    m.checkpoint(None);
    m.replace_range(second, 10, second, 10, " editada").unwrap();
    assert!(m
        .get_file_string("word/header2.xml")
        .unwrap()
        .contains("Cabecera B editada"));
    assert!(m
        .get_file_string("word/header1.xml")
        .unwrap()
        .contains(">Cabecera A<"));
    let l = layout(&m);
    assert_eq!(headers(&l), ["Cabecera A", "Cabecera B editada"]);
}
