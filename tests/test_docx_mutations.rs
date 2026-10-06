mod common;

use rust_web_docx::docx_parser::{DocxModifier, ParagraphUpdate};

#[test]
fn test_preserve_sectpr_and_drawings_in_plantilla() {
    let Some(bytes) = common::example("plantilla-ejemplo-vead.docx") else { return };
    let mut modifier = DocxModifier::from_bytes(&bytes).expect("Should parse docx");

    let stats_before = modifier.get_statistics().expect("Should get stats before");
    assert_eq!(stats_before.page_setup.orientation, "landscape");
    assert!(common::has_picture_behind_text(&modifier), "Should have background image initially");

    // 1. Update paragraph 1 ({name})
    let count1 = modifier.update_paragraphs(&[
        ParagraphUpdate {
            index: 1,
            text: "ALEXANDER VON HUMBOLDT".to_string(),
        }
    ]).expect("Should update paragraph 1");
    assert_eq!(count1, 1);
    let new_xml = modifier.get_file_string("word/document.xml").unwrap();
    println!("=== NEW XML ===\n{}\n===============", new_xml);



    let stats_after_p1 = modifier.get_statistics().expect("Should get stats after p1");
    assert_eq!(stats_after_p1.page_setup.orientation, "landscape", "Orientation MUST remain landscape after editing p1");
    assert!(common::has_picture_behind_text(&modifier), "Background image MUST remain present after editing p1");

    // 2. Update paragraph 2 ({asset_name})
    let count2 = modifier.update_paragraphs(&[
        ParagraphUpdate {
            index: 2,
            text: "CERTIFICADO DE HONOR".to_string(),
        }
    ]).expect("Should update paragraph 2");
    assert_eq!(count2, 1);

    let stats_after_p2 = modifier.get_statistics().expect("Should get stats after p2");
    assert_eq!(stats_after_p2.page_setup.orientation, "landscape", "Orientation MUST remain landscape after editing p2");
    assert!(common::has_picture_behind_text(&modifier), "Background image MUST remain present after editing p2");

    // 3. Update paragraph 4 ({documento_de_identidad}) - paragraph right before sectPr
    let count4 = modifier.update_paragraphs(&[
        ParagraphUpdate {
            index: 4,
            text: "ID: 1020304050".to_string(),
        }
    ]).expect("Should update paragraph 4");
    assert_eq!(count4, 1);

    let stats_after_p4 = modifier.get_statistics().expect("Should get stats after p4");
    assert_eq!(stats_after_p4.page_setup.orientation, "landscape", "Orientation MUST remain landscape after editing p4");
    assert!(common::has_picture_behind_text(&modifier), "Background image MUST remain present after editing p4");

    // 4. Update paragraph 0 (the paragraph that contains the drawing / background anchor)
    // If paragraph 0 is updated, the drawing MUST NOT be wiped out!
    let count0 = modifier.update_paragraphs(&[
        ParagraphUpdate {
            index: 0,
            text: "\n\n\n".to_string(),
        }
    ]).expect("Should update paragraph 0");
    assert_eq!(count0, 1);

    let stats_after_p0 = modifier.get_statistics().expect("Should get stats after p0");
    assert_eq!(stats_after_p0.page_setup.orientation, "landscape", "Orientation MUST remain landscape after editing p0");
    assert!(common::has_picture_behind_text(&modifier), "Background image drawing MUST NOT be lost when updating paragraph 0");

    // Verify raw text
    let raw = modifier.extract_raw_text().expect("Should get raw text");
    assert!(raw.contains("ALEXANDER VON HUMBOLDT"));
    assert!(raw.contains("CERTIFICADO DE HONOR"));
    assert!(raw.contains("ID: 1020304050"));
}

#[test]
fn test_rich_paragraph_formatting_preserves_sectpr() {
    let Some(bytes) = common::example("plantilla-ejemplo-vead.docx") else { return };
    let mut modifier = DocxModifier::from_bytes(&bytes).expect("Should parse docx");

    // Apply rich formatting to paragraph 1
    modifier.update_paragraph_rich(1, "JUAN PEREZ (BOLD RED)", "center", "FF0000", true, false)
        .expect("Should update rich paragraph 1");

    let stats = modifier.get_statistics().expect("Should get stats");
    assert_eq!(stats.page_setup.orientation, "landscape");
    assert!(common::has_picture_behind_text(&modifier));

    // Apply rich formatting to paragraph 4 (before sectPr)
    modifier.update_paragraph_rich(4, "ID: 999888777", "center", "008800", false, true)
        .expect("Should update rich paragraph 4");
}

#[test]
fn test_update_paragraph_runs_per_word() {
    let Some(bytes) = common::example("plantilla-ejemplo-vead.docx") else { return };
    let mut modifier = DocxModifier::from_bytes(&bytes).expect("Should parse docx");

    // Paragraph 1: Set 3 words with different formatting:
    // "CERTIFICA" (bold red), "QUE" (regular), "JUAN PEREZ" (bold italic blue)
    let runs = vec![
        rust_web_docx::docx_parser::RunInfo {
            text: "CERTIFICA ".to_string(),
            bold: true,
            italic: false,
            underline: false,
            color: "DC2626".to_string(),
            font_size: Some(18.0),
            font_family: Some("Calibri".to_string()),
        },
        rust_web_docx::docx_parser::RunInfo {
            text: "QUE ".to_string(),
            bold: false,
            italic: false,
            underline: false,
            color: "".to_string(),
            font_size: Some(18.0),
            font_family: Some("Calibri".to_string()),
        },
        rust_web_docx::docx_parser::RunInfo {
            text: "JUAN PEREZ".to_string(),
            bold: true,
            italic: true,
            underline: true,
            color: "2563EB".to_string(),
            font_size: Some(20.0),
            font_family: Some("Calibri".to_string()),
        },
    ];

    modifier.update_paragraph_runs(1, &runs, Some("center")).expect("Should update paragraph runs");

    let paragraphs = modifier.extract_paragraphs().expect("Should extract paragraphs");
    let p1 = &paragraphs[1];
    assert_eq!(p1.text, "CERTIFICA QUE JUAN PEREZ");
    assert_eq!(p1.runs.len(), 3);
    assert_eq!(p1.runs[0].text, "CERTIFICA ");
    assert!(p1.runs[0].bold);
    assert!(!p1.runs[0].italic);
    assert_eq!(p1.runs[0].color, "DC2626");

    assert_eq!(p1.runs[1].text, "QUE ");
    assert!(!p1.runs[1].bold);

    assert_eq!(p1.runs[2].text, "JUAN PEREZ");
    assert!(p1.runs[2].bold);
    assert!(p1.runs[2].italic);
    assert!(p1.runs[2].underline);
    assert_eq!(p1.runs[2].color, "2563EB");

    // Verify sectPr preserved
    let stats = modifier.get_statistics().expect("Should get stats");
    assert_eq!(stats.page_setup.orientation, "landscape");
    assert!(common::has_picture_behind_text(&modifier));
}
