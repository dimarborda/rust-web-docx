use rust_web_docx::docx_parser::{DocumentElement, DocxModifier, ParagraphUpdate};
use std::fs;

fn examples() -> Vec<(String, Vec<u8>)> {
    let mut out: Vec<(String, Vec<u8>)> = fs::read_dir("examples")
        .expect("examples dir")
        .flatten()
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|s| s.to_str()) == Some("docx"))
        .map(|p| (p.file_name().unwrap().to_string_lossy().to_string(), fs::read(&p).unwrap()))
        .collect();
    out.sort();
    out
}

/// Re-submitting every paragraph exactly as the editor sees it must not change a single byte
#[test]
fn test_resubmitting_unchanged_paragraphs_is_lossless() {
    for (name, bytes) in examples() {
        let mut modifier = DocxModifier::from_bytes(&bytes).unwrap();
        let original = modifier.get_file_string("word/document.xml").unwrap();
        for p in modifier.extract_paragraphs().unwrap() {
            modifier
                .update_paragraph_runs(p.index, &p.runs, Some(&p.align))
                .unwrap_or_else(|e| panic!("{}: paragraph {} failed: {}", name, p.index, e));
        }
        assert!(
            modifier.get_file_string("word/document.xml").unwrap() == original,
            "{}: document.xml changed after no-op edits",
            name
        );
    }
}

/// Appending text to every paragraph only touches the last run of each one
#[test]
fn test_appending_text_keeps_all_markup() {
    for (name, bytes) in examples() {
        let mut modifier = DocxModifier::from_bytes(&bytes).unwrap();
        let before = modifier.get_file_string("word/document.xml").unwrap();
        let paragraphs = modifier.extract_paragraphs().unwrap();
        let updates: Vec<ParagraphUpdate> = paragraphs
            .iter()
            .filter(|p| !p.text.is_empty())
            .map(|p| ParagraphUpdate { index: p.index, text: format!("{} ✓", p.text) })
            .collect();
        modifier.update_paragraphs(&updates).unwrap();
        let after = modifier.get_file_string("word/document.xml").unwrap();

        for tag in ["<w:hyperlink", "<w:highlight", "<w:numPr", "<w:drawing", "<w:rStyle", "<w:sz ", "<w:sectPr", "<w:bookmarkStart"] {
            assert_eq!(before.matches(tag).count(), after.matches(tag).count(), "{}: lost {} markup", name, tag);
        }
        let texts: Vec<String> = modifier.extract_paragraphs().unwrap().into_iter().map(|p| p.text).collect();
        for u in &updates {
            assert_eq!(texts[u.index], u.text, "{}: paragraph {}", name, u.index);
        }

        // Still a valid package after export
        let reloaded = DocxModifier::from_bytes(&modifier.to_bytes().unwrap()).unwrap();
        assert_eq!(reloaded.extract_paragraphs().unwrap().len(), paragraphs.len());
    }
}

#[test]
fn test_table_cell_edit_keeps_cell_formatting() {
    let bytes = rust_web_docx::sample_generator::generate_sample_docx().unwrap();
    let mut modifier = DocxModifier::from_bytes(&bytes).unwrap();
    let tables = |m: &DocxModifier| -> Vec<_> {
        m.extract_elements()
            .unwrap()
            .into_iter()
            .filter_map(|e| if let DocumentElement::Table(t) = e { Some(t) } else { None })
            .collect()
    };
    let before = tables(&modifier)[0].rich_rows[1].cells[1].clone();

    let new_text = format!("{} (editado)", before.text);
    modifier.update_table_cell(0, 1, 1, &new_text).unwrap();

    let after = tables(&modifier)[0].rich_rows[1].cells[1].clone();
    assert_eq!(after.text, new_text);
    assert_eq!((after.bold, after.italic, &after.color, &after.align), (before.bold, before.italic, &before.color, &before.align));
    assert_eq!(after.font_size, before.font_size);
}

#[test]
fn test_table_cell_multiline_edit_keeps_paragraph_boundaries() {
    let bytes = rust_web_docx::sample_generator::generate_sample_docx().unwrap();
    let mut modifier = DocxModifier::from_bytes(&bytes).unwrap();
    modifier.update_table_cell(0, 1, 0, "Línea A\nLínea B").unwrap();
    modifier.update_table_cell(0, 1, 0, "Línea A\nLínea B2").unwrap();
    let xml = modifier.get_file_string("word/document.xml").unwrap();
    assert!(xml.contains("Línea A"));
    let table = modifier.extract_tables().unwrap().remove(0);
    assert_eq!(table.rows[1][0], "Línea A\nLínea B2");
}
