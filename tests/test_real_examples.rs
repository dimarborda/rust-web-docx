use rust_web_docx::docx_parser::{DocxModifier, DocumentElement};
use rust_web_docx::layout_engine::LayoutEngine;
use std::fs;
use std::path::Path;

#[test]
fn test_analyze_real_examples() {
    let example_dir = Path::new("examples");
    if !example_dir.exists() {
        println!("examples directory not found, skipping");
        return;
    }

    let files = fs::read_dir(example_dir).expect("Failed to read examples dir");
    for entry in files {
        let entry = entry.expect("Valid entry");
        let path = entry.path();
        if path.extension().and_then(|s| s.to_str()) == Some("docx") {
            let filename = path.file_name().unwrap().to_string_lossy().to_string();
            println!("\n=======================================================");
            println!("🔍 ANALYZING REAL DOCX: {}", filename);
            println!("=======================================================");

            let bytes = fs::read(&path).expect("Failed to read docx file");
            let modifier = match DocxModifier::from_bytes(&bytes) {
                Ok(m) => m,
                Err(e) => {
                    println!("❌ Failed to parse docx {}: {}", filename, e);
                    panic!("Failed to parse: {}", e);
                }
            };

            let stats = modifier.get_statistics().expect("Should get stats");
            println!("📊 Statistics:");
            println!("   - Paragraphs: {}", stats.paragraph_count);
            println!("   - Tables: {}", stats.table_count);
            println!("   - Words: {}", stats.word_count);
            println!("   - Characters: {}", stats.char_count);
            println!("   - Zip Files: {}", stats.files_in_zip.len());
            println!("   - Background Color: {}", stats.background_color);

            let elements = modifier.extract_elements().expect("Should extract elements");
            println!("📑 Document Elements: Total {}", elements.len());

            let mut heading_count = 0;
            let mut runs_total = 0;
            let mut bold_runs_count = 0;
            let mut italic_runs_count = 0;
            let mut colored_runs_count = 0;
            let mut tabs_count = 0;
            let mut indented_p_count = 0;

            for (idx, el) in elements.iter().enumerate() {
                match el {
                    DocumentElement::Paragraph(p) => {
                        if p.is_heading {
                            heading_count += 1;
                        }
                        if p.indent_left > 0.0 || p.indent_first_line != 0.0 {
                            indented_p_count += 1;
                        }
                        runs_total += p.runs.len();
                        for r in &p.runs {
                            if r.bold { bold_runs_count += 1; }
                            if r.italic { italic_runs_count += 1; }
                            if !r.color.is_empty() { colored_runs_count += 1; }
                            if r.text.contains('\t') { tabs_count += 1; }
                        }

                        if idx < 5 || p.is_heading || p.indent_left > 0.0 {
                            let preview: String = p.text.chars().take(80).collect();
                            println!(
                                "   [P#{}] Style: '{}', Align: '{}', Bold: {}, Indent(L:{:.1}, FL:{:.1}, R:{:.1}), Runs: {}, Text: '{}'",
                                p.index, p.style, p.align, p.bold, p.indent_left, p.indent_first_line, p.indent_right, p.runs.len(), preview
                            );
                        }
                    }
                    DocumentElement::Table(t) => {
                        let num_cols = t.rows.first().map(|r| r.len()).unwrap_or(0);
                        println!(
                            "   [TBL#{}] Rows: {}, Cols: {}, Header: {}",
                            t.index, t.rows.len(), num_cols, t.header_row
                        );
                        for (r_idx, r) in t.rows.iter().enumerate().take(3) {
                            let row_preview: Vec<String> = r.iter().map(|c| c.replace('\n', "\\n").chars().take(30).collect()).collect();
                            println!("       Row {}: {:?}", r_idx, row_preview);
                        }
                    }
                }
            }

            println!("\n🔎 Structure Summary for {}:", filename);
            println!("   - Headings: {}", heading_count);
            println!("   - Total Runs: {}", runs_total);
            println!("   - Bold Runs: {}", bold_runs_count);
            println!("   - Italic Runs: {}", italic_runs_count);
            println!("   - Colored Runs: {}", colored_runs_count);
            println!("   - Tab Stop Occurrences: {}", tabs_count);
            println!("   - Indented Paragraphs: {}", indented_p_count);

            // Test layout calculation
            let engine = LayoutEngine::new();
            let layout = engine.compute_layout(
                &elements,
                &stats.background_color,
                &stats.page_setup,
                &stats.header_footer,
                stats.bg_image_data_url.as_deref(),
                None,
                0.0,
            );
            println!("📄 Multi-Page Canvas Layout computed: {} pages ({:?}, {}x{} px)", layout.total_pages, stats.page_setup.orientation, layout.page_width, layout.page_height);
            for p in &layout.pages {
                println!("   Page {}: {} render items", p.page_number, p.items.len());
            }

            // Test round-trip export
            let exported_bytes = modifier.to_bytes().expect("Should export bytes");
            println!("💾 Re-export verified: {} bytes (original: {} bytes)", exported_bytes.len(), bytes.len());
            assert!(!exported_bytes.is_empty());
        }
    }
}
