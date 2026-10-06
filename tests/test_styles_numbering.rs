mod common;

use rust_web_docx::docx_parser::DocxModifier;
use rust_web_docx::layout_engine::{LayoutEngine, RenderCommand};

const CONTRATO: &str = "CONTRATO DE PROMESA DE COMPRAVENTA DE BIEN INMUEBLE_firma_64d246aa720735d06b3d89c5_69e8ff615174a2ecfbd6ce80_1777056097303.docx";

fn load(name: &str) -> Option<DocxModifier> {
    common::example(name).map(|bytes| DocxModifier::from_bytes(&bytes).expect("docx"))
}

#[test]
fn test_contract_list_labels_follow_numbering_xml() {
    let Some(m) = load(CONTRATO) else { return };
    let labels: Vec<(String, String)> = m
        .extract_paragraphs()
        .unwrap()
        .into_iter()
        .filter_map(|p| p.list_label.map(|l| (l.text, p.text.chars().take(20).collect())))
        .collect();
    let texts: Vec<&str> = labels.iter().map(|(l, _)| l.as_str()).collect();
    assert!(texts.starts_with(&["1.", "2."]), "{:?}", texts);
    assert!(texts.windows(5).any(|w| w == ["a)", "b)", "c)", "d)", "e)"]), "{:?}", texts);
    assert!(texts.windows(2).any(|w| w == ["A.", "B."]), "{:?}", texts);
}

#[test]
fn test_style_cascade_gives_real_typography() {
    let Some(m) = load("costo-eficiencia-modelos-llm-rag.docx") else { return };
    let paragraphs = m.extract_paragraphs().unwrap();
    let bullets: Vec<_> = paragraphs.iter().filter(|p| p.style == "723").collect();
    // 11 "List Bullet" paragraphs; 4 of them remove the bullet with a direct numId="0"
    assert_eq!(bullets.len(), 11);
    let labeled: Vec<_> = bullets.iter().filter(|p| p.list_label.is_some()).collect();
    assert_eq!(labeled.len(), 7);
    for p in labeled {
        assert_eq!(p.list_label.as_ref().unwrap().text, "•", "paragraph {}", p.index);
        assert!(p.indent_first_line < 0.0, "hanging indent from numbering.xml");
    }
    // Every text run now has a concrete size and font from docDefaults/styles
    for p in &paragraphs {
        for r in &p.runs {
            assert!(r.font_size.is_some() && r.font_family.is_some(), "paragraph {} run {:?}", p.index, r.text);
        }
    }
}

#[test]
fn test_layout_draws_label_in_hanging_indent() {
    let Some(m) = load(CONTRATO) else { return };
    let elements = m.extract_elements().unwrap();
    let stats = m.get_statistics().unwrap();
    let layout = LayoutEngine::new().compute_layout(
        &elements,
        &stats.background_color,
        &stats.page_setup,
        &stats.header_footer,
        None,
        None,
        0.0,
    );
    let items: Vec<&RenderCommand> = layout.pages.iter().flat_map(|p| p.items.iter()).collect();
    let label = items
        .iter()
        .find_map(|it| match it {
            RenderCommand::Text { text, x, paragraph_index, .. } if text == "a)" => Some((*x, *paragraph_index)),
            _ => None,
        })
        .expect("label command");
    let body_x = items
        .iter()
        .find_map(|it| match it {
            RenderCommand::Text { text, x, paragraph_index, line_index: 0, .. }
                if *paragraph_index == label.1 && text != "a)" => Some(*x),
            _ => None,
        })
        .expect("first line of the list paragraph");
    assert!(label.0 < body_x, "label at {} must sit left of the text at {}", label.0, body_x);
}

#[test]
fn test_cto_029_numbering_layout() {
    let path = std::env::var("CTO_DOCX_PATH").unwrap_or_default();
    let Ok(bytes) = std::fs::read(&path) else { return };
    let m = DocxModifier::from_bytes(&bytes).expect("docx");
    let elements = m.extract_elements().unwrap();
    let stats = m.get_statistics().unwrap();
    let layout = LayoutEngine::new().compute_layout(
        &elements,
        &stats.background_color,
        &stats.page_setup,
        &stats.header_footer,
        None,
        None,
        0.0,
    );
    let p29 = elements
        .iter()
        .filter_map(|e| match e { rust_web_docx::docx_parser::DocumentElement::Paragraph(p) => Some(p), _ => None })
        .find(|p| p.text.contains("Entregar al líder"))
        .expect("paragraph 29");
    assert!((p29.indent_left - 47.53).abs() < 0.1, "indent_left must come from numbering level (713 dxa): {}", p29.indent_left);
    assert!((p29.indent_first_line - (-23.67)).abs() < 0.1, "indent_first_line must come from hanging 355 dxa: {}", p29.indent_first_line);

    let items: Vec<&RenderCommand> = layout.pages.iter().flat_map(|p| p.items.iter()).collect();
    let l0 = items.iter().find(|it| match it {
        RenderCommand::Text { paragraph_index, line_index: 0, text, .. } if *paragraph_index == p29.index && text.contains("Entregar") => true,
        _ => false,
    }).expect("line 0 text");
    let l1 = items.iter().find(|it| match it {
        RenderCommand::Text { paragraph_index, line_index: 1, text, .. } if *paragraph_index == p29.index && text.contains("cumplimiento") => true,
        _ => false,
    }).expect("line 1 text");
    let label = items.iter().find(|it| match it {
        RenderCommand::Text { paragraph_index, line_index: 0, text, .. } if *paragraph_index == p29.index && text == "1." => true,
        _ => false,
    }).expect("label text");

    let (x0, x1, xlabel) = match (l0, l1, label) {
        (RenderCommand::Text { x: x0, .. }, RenderCommand::Text { x: x1, .. }, RenderCommand::Text { x: xl, .. }) => (*x0, *x1, *xl),
        _ => unreachable!(),
    };
    assert!((x0 - x1).abs() < 0.01, "Line 0 ({}) and line 1 ({}) must align flush at paragraph left indent", x0, x1);
    assert!(xlabel < x0, "Label ({}) must sit left of line 0 text ({})", xlabel, x0);
    assert!((x0 - xlabel - 23.67).abs() < 1.0, "Label to text gap must match hanging indent (23.67px): {}", x0 - xlabel);

    let item16 = items.iter().filter(|it| match it {
        RenderCommand::Text { paragraph_index, .. } if *paragraph_index == 57 => true,
        _ => false,
    }).count();
    // Item 16 has label + 2 wrapped lines of text = 3 text items (matches Word)
    assert_eq!(item16, 3, "Item 16 must have 2 lines of text (3 text items total including label)");

    // Item 16 lands at the bottom of page 3; item 17 lands at the top of page 4
    let page3_has_item16 = layout.pages[2].items.iter().any(|it| match it {
        RenderCommand::Text { paragraph_index, .. } if *paragraph_index == 57 => true,
        _ => false,
    });
    let page4_has_item17 = layout.pages[3].items.iter().any(|it| match it {
        RenderCommand::Text { paragraph_index, .. } if *paragraph_index == 58 => true,
        _ => false,
    });
    assert!(page3_has_item16, "Item 16 must be on page 3");
    assert!(page4_has_item17, "Item 17 must be on page 4");
}
