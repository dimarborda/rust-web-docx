use rust_web_docx::docx_parser::{DocumentElement, DocxModifier};
use rust_web_docx::layout_engine::{DocumentLayout, LayoutEngine, RenderCommand};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

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

fn cell(props: &str, text: &str) -> String {
    format!(r#"<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>{}</w:tcPr><w:p><w:r><w:t>{}</w:t></w:r></w:p></w:tc>"#, props, text)
}

const BORDERS: &str = r#"<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>"#;

/// 3 grid columns of 2000 twips (133.3 px):
///   row 0: [A spans 2 columns] [M restarts]
///   row 1: [B] [C]             [M continues]
///   row 2: [D] [E]             [M continues]
fn merged_table(m_text: &str) -> String {
    format!(
        r#"<w:tbl>{}<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>{}{}</w:tr><w:tr>{}{}{}</w:tr><w:tr>{}{}{}</w:tr></w:tbl><w:p/>"#,
        BORDERS,
        cell(r#"<w:gridSpan w:val="2"/>"#, "A"),
        cell(r#"<w:vMerge w:val="restart"/>"#, m_text),
        cell("", "B"),
        cell("", "C"),
        cell("<w:vMerge/>", ""),
        cell("", "D"),
        cell("", "E"),
        cell("<w:vMerge/>", ""),
    )
}

fn layout(m: &DocxModifier) -> DocumentLayout {
    let elements: Vec<DocumentElement> = m.extract_elements().unwrap();
    let inputs = m.layout_inputs().unwrap();
    LayoutEngine::new().compute_layout(&elements, "FFFFFF", &inputs.page_setup, &inputs.header_footer, None, None, 0.0)
}

/// (row, col, x, y, width, height) of every drawn cell
fn cells(l: &DocumentLayout) -> Vec<(usize, usize, f64, f64, f64, f64)> {
    l.pages
        .iter()
        .flat_map(|p| p.items.iter())
        .filter_map(|it| match it {
            RenderCommand::TableCell { row, col, x, y, width, height, .. } => Some((*row, *col, *x, *y, *width, *height)),
            _ => None,
        })
        .collect()
}

fn text_y(l: &DocumentLayout, text: &str) -> (f64, f64) {
    l.pages[0]
        .items
        .iter()
        .find_map(|it| match it {
            RenderCommand::Text { text: t, x, y, line: Some(_), .. } if t.starts_with(text) => Some((*x, *y)),
            _ => None,
        })
        .unwrap_or_else(|| panic!("no text {}", text))
}

#[test]
fn test_parser_reads_spans_and_merges() {
    let m = docx(&merged_table("M"));
    let tables: Vec<_> = m.extract_elements().unwrap().into_iter().filter_map(|e| match e {
        DocumentElement::Table(t) => Some(t),
        _ => None,
    }).collect();
    let rows = &tables[0].rich_rows;
    assert_eq!(rows[0].cells[0].grid_span, 2);
    assert_eq!(rows[0].cells[1].v_merge.as_deref(), Some("restart"));
    assert_eq!(rows[1].cells[2].v_merge.as_deref(), Some("continue"));
    assert_eq!(rows[1].cells[0].grid_span, 1);
}

#[test]
fn test_horizontal_and_vertical_merges_are_drawn_as_one_cell() {
    let m = docx(&merged_table("M"));
    let l = layout(&m);
    let drawn = cells(&l);
    let col_w = 2000.0 / 15.0;
    // A spans two columns
    let a = drawn.iter().find(|c| c.0 == 0 && c.1 == 0).unwrap();
    assert!((a.4 - 2.0 * col_w).abs() < 0.01, "A is two columns wide: {}", a.4);
    // M is drawn once, in the third column, as tall as the three rows
    let m_cells: Vec<_> = drawn.iter().filter(|c| (c.2 - (a.2 + 2.0 * col_w)).abs() < 0.01).collect();
    assert_eq!(m_cells.len(), 1, "the covered cells are not drawn: {:?}", m_cells);
    let row_tops: Vec<f64> = (0..3).map(|r| drawn.iter().find(|c| c.0 == r && c.1 == 0).unwrap().3).collect();
    let last = drawn.iter().find(|c| c.0 == 2 && c.1 == 0).unwrap();
    assert!((m_cells[0].3 - row_tops[0]).abs() < 0.01);
    assert!((m_cells[0].3 + m_cells[0].5 - (last.3 + last.5)).abs() < 0.01, "M reaches the bottom of row 2");
    // B, C, D, E sit in the first two columns
    let (bx, _) = text_y(&l, "B");
    let (cx, _) = text_y(&l, "C");
    assert!(bx < a.2 + col_w && cx > a.2 + col_w);
}

#[test]
fn test_tall_merged_cell_grows_the_last_row() {
    let long = "palabra ".repeat(60);
    let m = docx(&merged_table(&long));
    let l = layout(&m);
    let drawn = cells(&l);
    let heights: Vec<f64> = (0..3).map(|r| drawn.iter().find(|c| c.0 == r && c.1 == 0).unwrap().5).collect();
    assert!(heights[0] < 30.0 && heights[1] < 30.0, "rows 0 and 1 keep their own height: {:?}", heights);
    assert!(heights[2] > 60.0, "the last covered row grows to fit the merged text: {:?}", heights);
    // The merged text starts in row 0, not pushed down
    let (_, my) = text_y(&l, "palabra");
    let (_, ay) = text_y(&l, "A");
    assert!((my - ay).abs() < 0.5);
}

#[test]
fn test_grid_before_shifts_the_row() {
    let body = format!(
        r#"<w:tbl>{}<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>{}{}</w:tr><w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>{}</w:tr></w:tbl><w:p/>"#,
        BORDERS,
        cell("", "X"),
        cell("", "Y"),
        cell("", "Z"),
    );
    let l = layout(&docx(&body));
    let (xx, _) = text_y(&l, "X");
    let (yx, _) = text_y(&l, "Y");
    let (zx, _) = text_y(&l, "Z");
    assert!((zx - yx).abs() < 0.01 && zx > xx, "Z starts in the second column");
}
