mod common;

use std::fs;
use std::io::Read;
use std::path::Path;
use zip::ZipArchive;

fn inspect_single_file(filename: &str) {
    let path = Path::new("examples").join(filename);
    if !path.exists() {
        println!("File not found: {:?}", path);
        return;
    }

    println!("\n=======================================================");
    println!("🔍 DEEP ANALYSIS: {}", filename);
    println!("=======================================================");

    let file = fs::File::open(&path).unwrap();
    let mut archive = ZipArchive::new(file).unwrap();

    let mut media_files = Vec::new();
    let mut headers = Vec::new();
    let mut footers = Vec::new();

    for i in 0..archive.len() {
        let f = archive.by_index(i).unwrap();
        let name = f.name().to_string();
        let size = f.size();
        if name.starts_with("word/media/") {
            media_files.push((name, size));
        } else if name.starts_with("word/header") {
            headers.push((name, size));
        } else if name.starts_with("word/footer") {
            footers.push((name, size));
        }
    }

    println!("🖼️ Media files ({}):", media_files.len());
    for (m, s) in &media_files {
        println!("   - {} ({} KB)", m, s / 1024);
    }
    println!("📄 Headers ({}):", headers.len());
    for (h, s) in &headers {
        println!("   - {} ({} bytes)", h, s);
    }
    println!("📄 Footers ({}):", footers.len());
    for (ft, s) in &footers {
        println!("   - {} ({} bytes)", ft, s);
    }

    let mut xml = String::new();
    {
        if let Ok(mut doc_file) = archive.by_name("word/document.xml") {
            doc_file.read_to_string(&mut xml).unwrap();
        }
    }

    if !xml.is_empty() {
        // Check page orientation and size
        if let Some(sect_idx) = xml.rfind("<w:sectPr") {
            let sect = &xml[sect_idx..];
            let end = sect.find("</w:sectPr>").unwrap_or(sect.len().min(500));
            println!("\n📐 Page Setup (<w:sectPr>):\n{}", &sect[..end]);
        }

        // Check paragraph count, table count, runs
        let p_count = xml.matches("<w:p>").count() + xml.matches("<w:p ").count();
        let tbl_count = xml.matches("<w:tbl>").count() + xml.matches("<w:tbl ").count();
        let r_count = xml.matches("<w:r>").count() + xml.matches("<w:r ").count();
        let drawing_count = xml.matches("<w:drawing").count();
        let num_count = xml.matches("<w:numPr").count();
        let tab_count = xml.matches("<w:tab/>").count() + xml.matches("<w:tab ").count();
        let jc_center = xml.matches("<w:jc w:val=\"center\"").count();
        let jc_both = xml.matches("<w:jc w:val=\"both\"").count();
        let jc_right = xml.matches("<w:jc w:val=\"right\"").count();
        let ind_count = xml.matches("<w:ind ").count();

        println!("\n📊 XML Tags Count:");
        println!("   - Paragraphs (<w:p>): {}", p_count);
        println!("   - Tables (<w:tbl>): {}", tbl_count);
        println!("   - Runs (<w:r>): {}", r_count);
        println!("   - Drawings (<w:drawing>): {}", drawing_count);
        println!("   - Numbering / Lists (<w:numPr>): {}", num_count);
        println!("   - Tab characters (<w:tab/>): {}", tab_count);
        println!("   - Alignments: Justified (both)={}, Center={}, Right={}", jc_both, jc_center, jc_right);
        println!("   - Indentations (<w:ind>): {}", ind_count);
    // Check relationships
    for rel_name in &["word/_rels/document.xml.rels", "word/_rels/header1.xml.rels"] {
        if let Ok(mut rel_file) = archive.by_name(rel_name) {
            let mut rel_xml = String::new();
            rel_file.read_to_string(&mut rel_xml).unwrap();
            println!("   🔗 {}: {}", rel_name, rel_xml);
        }
    }

    // Check fontTable.xml
    if let Ok(mut font_file) = archive.by_name("word/fontTable.xml") {
        let mut font_xml = String::new();
        font_file.read_to_string(&mut font_xml).unwrap();
        println!("   🔤 fontTable.xml: {}", &font_xml[..font_xml.len().min(800)]);
    }

    // Check theme1.xml
    if let Ok(mut theme_file) = archive.by_name("word/theme/theme1.xml") {
        let mut theme_xml = String::new();
        theme_file.read_to_string(&mut theme_xml).unwrap();
        if let Some(pos) = theme_xml.find("<a:fontScheme") {
            let chunk = &theme_xml[pos..theme_xml.len().min(pos + 600)];
            println!("   🎨 theme1.xml fontScheme:\n{}", chunk);
        }
    }

    // Check styles.xml docDefaults
    if let Ok(mut styles_file) = archive.by_name("word/styles.xml") {
        let mut styles_xml = String::new();
        styles_file.read_to_string(&mut styles_xml).unwrap();
        if let Some(pos) = styles_xml.find("<w:docDefaults") {
            let chunk = &styles_xml[pos..styles_xml.len().min(pos + 600)];
            println!("   📝 styles.xml docDefaults:\n{}", chunk);
        }
    }

    }
}

#[test]
fn test_inspect_contrato() {
    inspect_single_file("CONTRATO DE PROMESA DE COMPRAVENTA DE BIEN INMUEBLE_firma_64d246aa720735d06b3d89c5_69e8ff615174a2ecfbd6ce80_1777056097303.docx");
    let Some(bytes) = common::example("CONTRATO DE PROMESA DE COMPRAVENTA DE BIEN INMUEBLE_firma_64d246aa720735d06b3d89c5_69e8ff615174a2ecfbd6ce80_1777056097303.docx") else { return };
    let modifier = rust_web_docx::docx_parser::DocxModifier::from_bytes(&bytes).unwrap();
    let stats = modifier.get_statistics().unwrap();
    assert!(stats.bg_image_data_url.is_none(), "Contrato should not have background image");
    assert!(stats.header_footer.header_image_data_url.is_none(), "Contrato should not have header image");

    let paragraphs = modifier.extract_paragraphs().unwrap();
    let bdr_left_count = paragraphs.iter().filter(|p| p.borders.left.is_some()).count();
    let bdr_bottom_count = paragraphs.iter().filter(|p| p.borders.bottom.is_some()).count();
    assert_eq!(bdr_left_count, 0, "Contrato must have 0 left borders (nil borders must be ignored)");
    assert_eq!(bdr_bottom_count, 0, "Contrato must have 0 bottom borders (nil borders must be ignored)");
}

#[test]
fn test_inspect_costo_eficiencia() {
    inspect_single_file("costo-eficiencia-modelos-llm-rag.docx");
    let Some(bytes) = common::example("costo-eficiencia-modelos-llm-rag.docx") else { return };
    let modifier = rust_web_docx::docx_parser::DocxModifier::from_bytes(&bytes).unwrap();
    let stats = modifier.get_statistics().unwrap();
    assert!(stats.bg_image_data_url.is_none(), "Costo-eficiencia should NOT have background image!");
    assert!(stats.header_footer.header_image_data_url.is_some(), "Costo-eficiencia MUST have header logo image!");
    assert_eq!(stats.page_setup.orientation, "portrait");
    let paragraphs = modifier.extract_paragraphs().unwrap();
    let bdr_left_count = paragraphs.iter().filter(|p| p.borders.left.is_some()).count();
    assert!(bdr_left_count > 0, "Costo-eficiencia must have left borders on sections");

    let mut fonts = std::collections::HashSet::new();
    for p in &paragraphs {
        if let Some(ref f) = p.font_family {
            fonts.insert(f.clone());
        }
        for r in &p.runs {
            if let Some(ref f) = r.font_family {
                fonts.insert(f.clone());
            }
        }
    }
    println!("Unique fonts found in costo-eficiencia: {:?}", fonts);

    let tables = modifier.extract_tables().unwrap();
    println!("Found {} tables in costo-eficiencia:", tables.len());
    for (t_idx, tbl) in tables.iter().enumerate() {
        println!("--- TABLE {} (grid_cols: {:?}) ---", t_idx, tbl.grid_cols);
        for (r_idx, r) in tbl.rich_rows.iter().enumerate() {
            println!("  Row {} (is_header: {}):", r_idx, r.is_header);
            for (c_idx, c) in r.cells.iter().enumerate() {
                println!("    Cell {}: text={:?}, bg={:?}, color={}, align={}, bold={}, border={}",
                    c_idx, c.text, c.bg_color, c.color, c.align, c.bold, c.border_color);
            }
        }
    }
}

#[test]
fn test_inspect_plantilla() {
    inspect_single_file("plantilla-ejemplo-vead.docx");
    let Some(bytes) = common::example("plantilla-ejemplo-vead.docx") else { return };
    let modifier = rust_web_docx::docx_parser::DocxModifier::from_bytes(&bytes).unwrap();
    let stats = modifier.get_statistics().unwrap();
    assert!(stats.bg_image_data_url.is_some(), "Plantilla diploma MUST have full background image!");
    let bg_url = stats.bg_image_data_url.unwrap();
    assert!(bg_url.starts_with("data:image/png;base64,iVBOR"), "Background image must be a valid PNG base64 data url");
    assert!(stats.header_footer.header_image_data_url.is_none(), "Plantilla diploma should NOT have header image!");
    assert_eq!(stats.page_setup.orientation, "landscape");
}

