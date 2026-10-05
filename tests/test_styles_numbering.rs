use rust_web_docx::docx_parser::DocxModifier;
use rust_web_docx::layout_engine::{LayoutEngine, RenderCommand};
use std::fs;

const CONTRATO: &str = "examples/CONTRATO DE PROMESA DE COMPRAVENTA DE BIEN INMUEBLE_firma_64d246aa720735d06b3d89c5_69e8ff615174a2ecfbd6ce80_1777056097303.docx";

fn load(path: &str) -> DocxModifier {
    DocxModifier::from_bytes(&fs::read(path).expect("example")).expect("docx")
}

#[test]
fn test_contract_list_labels_follow_numbering_xml() {
    let m = load(CONTRATO);
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
    let m = load("examples/costo-eficiencia-modelos-llm-rag.docx");
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
    let m = load(CONTRATO);
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
