mod common;

use rust_web_docx::docx_parser::{DocxModifier, KeyValuePair};

fn pair(key: &str, value: &str) -> KeyValuePair {
    KeyValuePair { key: key.to_string(), value: value.to_string() }
}

#[test]
fn test_vead_template_variables_split_across_runs() {
    let Some(bytes) = common::example("plantilla-ejemplo-vead.docx") else { return };
    let mut modifier = DocxModifier::from_bytes(&bytes).expect("Should parse docx");

    let text_before = modifier.extract_raw_text().unwrap();
    for var in ["{name}", "{asset_name}", "{email}", "{documento_de_identidad}"] {
        assert!(text_before.contains(var), "template should contain {}", var);
    }

    let result = modifier
        .batch_replace(&[
            pair("{name}", "Alexander von Humboldt"),
            pair("{asset_name}", "Certificado de Honor"),
            pair("{email}", "alex@example.com"),
            pair("{documento_de_identidad}", "CC 1.234.567"),
        ])
        .expect("Batch replace should succeed");
    assert_eq!(result.occurrences_replaced, 4);
    assert_eq!(result.affected_files, vec!["word/document.xml".to_string()]);

    let text_after = modifier.extract_raw_text().unwrap();
    for value in ["Alexander von Humboldt", "Certificado de Honor", "alex@example.com", "CC 1.234.567"] {
        assert!(text_after.contains(value), "missing '{}' in:\n{}", value, text_after);
    }
    assert!(!text_after.contains('{') && !text_after.contains('}'), "leftover braces:\n{}", text_after);

    // The exported file must round-trip with layout and media intact
    let exported = modifier.to_bytes().expect("Should export");
    let reloaded = DocxModifier::from_bytes(&exported).expect("Exported docx should parse");
    let stats = reloaded.get_statistics().unwrap();
    assert_eq!(stats.page_setup.orientation, "landscape");
    assert!(stats.bg_image_data_url.is_some(), "background image must survive");
    assert!(reloaded.extract_raw_text().unwrap().contains("Alexander von Humboldt"));
}

#[test]
fn test_vead_template_find_and_replace_single_variable() {
    let Some(bytes) = common::example("plantilla-ejemplo-vead.docx") else { return };
    let mut modifier = DocxModifier::from_bytes(&bytes).expect("Should parse docx");

    let result = modifier.find_and_replace("{EMAIL}", "x@y.co", false, false).unwrap();
    assert_eq!(result.occurrences_replaced, 1);
    assert!(modifier.extract_raw_text().unwrap().contains("x@y.co"));
}
