use rust_web_docx::docx_parser::{DocumentElement, DocxModifier};
use std::io::{Cursor, Write};
use zip::write::SimpleFileOptions;
use zip::ZipWriter;

const BORDERS: &str = r#"<w:tblBorders><w:top w:val="single" w:sz="8" w:color="1F4E79"/><w:left w:val="single" w:sz="8" w:color="1F4E79"/><w:bottom w:val="single" w:sz="8" w:color="1F4E79"/><w:right w:val="single" w:sz="8" w:color="1F4E79"/><w:insideH w:val="single" w:sz="4" w:color="1F4E79"/><w:insideV w:val="single" w:sz="4" w:color="1F4E79"/></w:tblBorders>"#;

fn header_cell(text: &str) -> String {
    format!(
        r#"<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="1F4E79"/></w:tcPr><w:p w14:paraId="1A2B3C4D"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="22"/></w:rPr><w:t>{}</w:t></w:r></w:p></w:tc>"#,
        text
    )
}

fn data_cell(text: &str) -> String {
    format!(
        r#"<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="DEEAF6"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:color w:val="C00000"/><w:sz w:val="20"/></w:rPr><w:t>{}</w:t></w:r></w:p></w:tc>"#,
        text
    )
}

fn row(cells: &[String], header: bool) -> String {
    let pr = if header { r#"<w:trPr><w:tblHeader/></w:trPr>"# } else { r#"<w:trPr><w:trHeight w:val="400"/></w:trPr>"# };
    format!(r#"<w:tr w14:paraId="0000AAAA">{}{}</w:tr>"#, pr, cells.concat())
}

/// Table 0: risk matrix with a shaded, repeated header and two data rows.
/// Table 1: merged cells (a gridSpan in row 0, a vertical merge in rows 1–2) and a nested table.
fn body() -> String {
    let matrix = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>{}</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>{}{}{}</w:tbl>"#,
        BORDERS,
        row(&[header_cell("Riesgo"), header_cell("Probabilidad"), header_cell("Impacto")], true),
        row(&[data_cell("Caída del servidor"), data_cell("Media"), data_cell("Alto")], false),
        row(&[data_cell("Fuga de datos"), data_cell("Baja"), data_cell("Crítico")], false),
    );
    let merged = r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Dos columnas</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Alto</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b1</w:t></w:r></w:p></w:tc><w:tc><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>anidada</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>b2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>c2</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>y</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>z</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
    format!(r#"<w:p><w:r><w:t>Matriz de riesgos</w:t></w:r></w:p>{}<w:p><w:r><w:t>Entre tablas</w:t></w:r></w:p>{}<w:p/>"#, matrix, merged)
}

fn doc() -> DocxModifier {
    doc_with(&body())
}

fn doc_with(body: &str) -> DocxModifier {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>{}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#,
        body
    );
    let mut buf = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut buf);
        for (name, content) in [
            ("[Content_Types].xml", r#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#),
            ("word/document.xml", document.as_str()),
        ] {
            zip.start_file(name, SimpleFileOptions::default()).unwrap();
            zip.write_all(content.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
    }
    DocxModifier::from_bytes(&buf.into_inner()).unwrap()
}

fn document_xml(m: &DocxModifier) -> String {
    m.get_file_string("word/document.xml").unwrap()
}

/// Cell texts of body table `index`
fn rows_of(m: &DocxModifier, index: usize) -> Vec<Vec<String>> {
    m.extract_tables().unwrap()[index].rows.clone()
}

/// XML of row `row` of the risk matrix (table 0)
fn matrix_row_xml(m: &DocxModifier, row: usize) -> String {
    let xml = document_xml(m);
    let table = &xml[xml.find("<w:tbl>").unwrap()..xml.find("</w:tbl>").unwrap()];
    let rows = &table[table.find("<w:tr").unwrap()..];
    rows.split("</w:tr>").nth(row).unwrap().to_string()
}

#[test]
fn test_tables_lists_body_tables_with_header_and_paragraph_range() {
    let m = doc();
    let tables = m.table_summaries().unwrap();
    assert_eq!(tables.len(), 2, "nested tables are not listed");
    assert_eq!((tables[0].index, tables[0].rows, tables[0].cols), (0, 3, 3));
    assert_eq!(tables[0].header, vec!["Riesgo", "Probabilidad", "Impacto"]);
    // Paragraph 0 is the title; the matrix holds paragraphs 1–9 and "Entre tablas" is 10
    assert_eq!((tables[0].first, tables[0].last), (1, 9));
    assert_eq!(tables[1].first, 11);
    assert_eq!(tables[1].header, vec!["Dos columnas", "", "C"]);
    let paragraphs = m.extract_paragraphs().unwrap();
    assert!(paragraphs.iter().any(|p| p.index == 10 && p.text == "Entre tablas"));
}

#[test]
fn test_set_table_cell_keeps_formatting_and_breaks_lines() {
    let mut m = doc();
    m.set_table_cell(0, 1, 1, "Alta\nrevisar").unwrap();
    assert_eq!(rows_of(&m, 0)[1][1], "Alta\nrevisar");
    let row = matrix_row_xml(&m, 1);
    assert!(row.contains("<w:br/>"), "{}", row);
    assert!(row.contains(r#"<w:color w:val="C00000"/>"#) && row.contains(r#"<w:sz w:val="20"/>"#));
    assert!(row.contains(r#"w:fill="DEEAF6""#));
    // Columns are grid columns: the merged cell is written through its first column
    m.set_table_cell(1, 0, 0, "Fusionada").unwrap();
    m.set_table_cell(1, 0, 2, "Tercera").unwrap();
    assert_eq!(rows_of(&m, 1)[0], vec!["Fusionada", "Tercera"]);
}

#[test]
fn test_insert_rows_at_start_middle_and_end_copy_the_formatting() {
    let mut m = doc();
    // Middle: copies the row before (data row)
    assert_eq!(m.insert_table_rows(0, 2, &[vec!["Phishing".into(), "Alta".into(), "Medio".into()]], None).unwrap(), 2);
    // End, several rows, the second one short
    m.insert_table_rows(0, 4, &[vec!["Ransomware".into(), "Media".into(), "Crítico".into()], vec!["Error humano".into()]], None).unwrap();
    // Start: copies the next row (the shaded header), still repeated on every page
    m.insert_table_rows(0, 0, &[vec!["Código".into(), "".into(), "".into()]], None).unwrap();

    let rows = rows_of(&m, 0);
    let texts: Vec<&str> = rows.iter().map(|r| r[0].as_str()).collect();
    assert_eq!(texts, ["Código", "Riesgo", "Caída del servidor", "Phishing", "Fuga de datos", "Ransomware", "Error humano"]);
    assert_eq!(rows[6], vec!["Error humano", "", ""], "short rows get empty cells");

    let top = matrix_row_xml(&m, 0);
    assert!(top.contains("<w:tblHeader/>") && top.contains(r#"w:fill="1F4E79""#) && top.contains("<w:b/>"), "{}", top);
    assert!(top.contains(r#"<w:jc w:val="center"/>"#));
    for r in [3, 5, 6] {
        let xml = matrix_row_xml(&m, r);
        assert!(!xml.contains("tblHeader"), "data rows are not headers: {}", xml);
        assert!(xml.contains(r#"<w:trHeight w:val="400"/>"#) && xml.contains(r#"w:fill="DEEAF6""#), "{}", xml);
        assert!(xml.contains(r#"<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>"#) && xml.contains(r#"<w:tcW w:w="3000" w:type="dxa"/>"#));
        assert!(!xml.contains("w14:paraId"), "copied rows must not repeat paragraph ids");
    }
    // Empty cells keep the formatting for text typed later
    assert!(matrix_row_xml(&m, 6).contains(r#"<w:pPr><w:spacing w:after="0"/><w:rPr><w:rFonts"#));

    // copyFormatFrom: a data row made from the header row's formatting, and back to the file
    m.insert_table_rows(0, 7, &[vec!["Total".into()]], Some(1)).unwrap();
    let last = matrix_row_xml(&m, 7);
    assert!(last.contains(r#"w:fill="1F4E79""#) && !last.contains("tblHeader"));

    let reopened = DocxModifier::from_bytes(&m.to_bytes().unwrap()).unwrap();
    assert_eq!(rows_of(&reopened, 0).len(), 8);
    assert_eq!(reopened.table_summaries().unwrap()[0].rows, 8);
}

#[test]
fn test_delete_rows() {
    let mut m = doc();
    m.delete_table_rows(0, 1, 1).unwrap();
    assert_eq!(rows_of(&m, 0).iter().map(|r| r[0].as_str()).collect::<Vec<_>>(), ["Riesgo", "Fuga de datos"]);
    let before = document_xml(&m);
    assert!(m.delete_table_rows(0, 0, 2).unwrap_err().contains("sin filas"));
    // A whole vertical merge can go; cutting it cannot
    assert!(m.delete_table_rows(1, 1, 1).unwrap_err().contains("combinadas"));
    assert!(m.delete_table_rows(1, 2, 1).unwrap_err().contains("combinadas"));
    assert_eq!(document_xml(&m), before);
    m.delete_table_rows(1, 1, 2).unwrap();
    assert_eq!(rows_of(&m, 1).len(), 2);
}

#[test]
fn test_invalid_indices_and_merged_cells_are_refused_without_changes() {
    let mut m = doc();
    let before = document_xml(&m);
    let errors = [
        m.set_table_cell(5, 0, 0, "x").unwrap_err(),
        m.set_table_cell(0, 3, 0, "x").unwrap_err(),
        m.set_table_cell(0, 0, 3, "x").unwrap_err(),
        m.set_table_cell(1, 0, 1, "x").unwrap_err(),
        m.set_table_cell(1, 2, 0, "x").unwrap_err(),
        m.set_table_cell(1, 1, 2, "x").unwrap_err(),
        m.insert_table_rows(0, 4, &[vec!["a".into()]], None).unwrap_err(),
        m.insert_table_rows(0, 1, &[vec!["a".into(), "b".into(), "c".into(), "d".into()]], None).unwrap_err(),
        m.insert_table_rows(0, 1, &[], None).unwrap_err(),
        m.insert_table_rows(0, 1, &[vec!["a".into()]], Some(9)).unwrap_err(),
        m.insert_table_rows(1, 2, &[vec!["a".into()]], None).unwrap_err(),
        m.delete_table_rows(0, 2, 2).unwrap_err(),
        m.delete_table_rows(0, 0, 0).unwrap_err(),
    ];
    assert_eq!(document_xml(&m), before, "failed calls must not touch the document");
    assert!(errors[0].contains("2 tablas"), "{}", errors[0]);
    assert!(errors[1].contains("fila 3") && errors[2].contains("columna 3"));
    assert!(errors[3].contains("escribe en la columna 0"), "{}", errors[3]);
    assert!(errors[4].contains("escribe en la fila 1"), "{}", errors[4]);
    assert!(errors[5].contains("anidada"), "{}", errors[5]);
    assert!(errors[7].contains("4 valores"), "{}", errors[7]);
    assert!(errors[10].contains("combinadas"), "{}", errors[10]);
}

#[test]
fn test_each_operation_is_one_undo_step() {
    let mut m = doc();
    let original = document_xml(&m);
    let ops: Vec<Box<dyn Fn(&mut DocxModifier)>> = vec![
        Box::new(|m| m.set_table_cell(0, 2, 0, "Otro\nriesgo").unwrap()),
        Box::new(|m| {
            m.insert_table_rows(0, 1, &[vec!["a".into()], vec!["b".into()], vec!["c".into()]], None).unwrap();
        }),
        Box::new(|m| m.delete_table_rows(0, 1, 2).unwrap()),
    ];
    for op in ops {
        m.checkpoint(None);
        op(&mut m);
        assert_ne!(document_xml(&m), original);
        let changed = document_xml(&m);
        assert!(m.undo().is_some());
        assert_eq!(document_xml(&m), original, "one undo reverts the whole operation");
        m.redo();
        assert_eq!(document_xml(&m), changed);
        m.undo();
    }
    assert!(matches!(m.extract_elements().unwrap()[1], DocumentElement::Table(_)));
}

#[test]
fn test_writing_an_empty_inserted_cell_uses_the_copied_formatting() {
    let mut m = doc();
    m.insert_table_rows(0, 3, &[vec!["Nuevo".into()]], None).unwrap();
    m.set_table_cell(0, 3, 2, "Moderado").unwrap();
    let row = matrix_row_xml(&m, 3);
    let cell = row.split("<w:tc>").nth(3).unwrap();
    assert!(cell.contains("Moderado"), "{}", cell);
    let run = &cell[cell.find("<w:r>").unwrap()..];
    assert!(run.contains(r#"<w:color w:val="C00000"/>"#) && run.contains("Arial"), "{}", run);
}

/// XML of body table `index` (top-level tables only)
fn table_xml(m: &DocxModifier, index: usize) -> String {
    let xml = document_xml(m);
    let mut starts = Vec::new();
    let mut depth = 0;
    let mut pos = 0;
    while let Some(i) = xml[pos..].find("<w:tbl") {
        let at = pos + i;
        let close = xml[pos..at].matches("</w:tbl>").count();
        depth -= close as i32;
        if xml[at..].starts_with("<w:tbl>") || xml[at..].starts_with("<w:tbl ") {
            if depth == 0 {
                starts.push(at);
            }
            depth += 1;
        }
        pos = at + 6;
    }
    let start = starts[index];
    // The matching close tag
    let mut depth = 0;
    let mut i = start;
    loop {
        let open = xml[i..].find("<w:tbl>").map(|k| i + k);
        let close = i + xml[i..].find("</w:tbl>").unwrap();
        match open {
            Some(o) if o < close => {
                depth += 1;
                i = o + 7;
            }
            _ => {
                depth -= 1;
                i = close + 8;
                if depth == 0 {
                    return xml[start..i].to_string();
                }
            }
        }
    }
}

fn texts(m: &DocxModifier) -> Vec<(usize, String)> {
    let mut out: Vec<(usize, String)> = m.extract_paragraphs().unwrap().into_iter().map(|p| (p.index, p.text)).collect();
    out.sort();
    out
}

#[test]
fn test_duplicate_table_copies_the_formatting_and_fills_values() {
    let mut m = doc();
    let original = table_xml(&m, 0);
    let summaries = m.table_summaries().unwrap();
    let after_tables = summaries[1].last + 1; // the empty paragraph at the end
    let values = vec![vec![None, Some("Probabilidad (%)".to_string())], vec![Some("Corte eléctrico".to_string())]];
    let (table, first, caret) = m.duplicate_table(0, after_tables, 0, Some(&values)).unwrap();
    assert_eq!((table, first), (2, after_tables));
    assert_eq!(caret, (after_tables + 9, 0), "the caret is in the paragraph after the copy");
    assert_eq!(table_xml(&m, 0), original, "the source table does not change");
    assert_eq!(
        rows_of(&m, 2),
        vec![
            vec!["Riesgo", "Probabilidad (%)", "Impacto"],
            vec!["Corte eléctrico", "Media", "Alto"],
            vec!["Fuga de datos", "Baja", "Crítico"],
        ]
    );
    let copy = table_xml(&m, 2);
    for kept in [BORDERS, r#"<w:tblHeader/>"#, r#"w:fill="1F4E79""#, r#"w:fill="DEEAF6""#, r#"<w:color w:val="C00000"/>"#, r#"<w:gridCol w:w="3000"/>"#] {
        assert!(copy.contains(kept), "{} missing in {}", kept, copy);
    }
    assert!(!copy.contains("w14:paraId"), "paragraph ids stay unique");
    // Written cells keep the cell's formatting
    let written = &copy[copy.find("Corte eléctrico").unwrap() - 200..];
    assert!(written.contains("Arial"), "{}", written);
    assert_eq!(m.table_summaries().unwrap().len(), 3);
}

#[test]
fn test_duplicate_table_before_the_source_and_inside_a_paragraph() {
    let mut m = doc();
    // In the middle of "Matriz de riesgos": the text after the position goes below the copy
    let (table, first, caret) = m.duplicate_table(1, 0, 7, None).unwrap();
    assert_eq!((table, first), (0, 1));
    let tables = m.table_summaries().unwrap();
    assert_eq!(tables.len(), 3);
    assert_eq!(tables[0].header, vec!["Dos columnas", "", "C"]);
    assert_eq!(tables[1].header, vec!["Riesgo", "Probabilidad", "Impacto"], "the source moves down");
    let all = texts(&m);
    assert_eq!(all[0], (0, "Matriz ".to_string()));
    assert_eq!(all.iter().find(|(i, _)| *i == caret.0).map(|(_, t)| t.as_str()), Some("de riesgos"));
    // Merged cells and the nested table come along
    let copy = table_xml(&m, 0);
    assert!(copy.contains(r#"<w:gridSpan w:val="2"/>"#) && copy.contains(r#"<w:vMerge w:val="restart"/>"#) && copy.contains("anidada"));
    assert_eq!(rows_of(&m, 0), rows_of(&m, 2));
}

#[test]
fn test_duplicate_table_refusals_leave_the_document_untouched() {
    let mut m = doc();
    let before = document_xml(&m);
    let err = |r: Result<(usize, usize, (usize, usize)), String>| r.unwrap_err();
    assert!(err(m.duplicate_table(5, 0, 0, None)).contains("No existe la tabla 5"));
    // Inside a table
    assert!(err(m.duplicate_table(0, 2, 0, None)).contains("dentro de otra tabla"));
    assert!(err(m.duplicate_table(0, 0, 0, Some(&[vec![None, None, None, Some("x".into())]]))).contains("4 valores"));
    assert!(err(m.duplicate_table(0, 0, 0, Some(&vec![vec![]; 4]))).contains("4 filas"));
    // Fails while filling (a cell continuing a vertical merge) after splitting the paragraph:
    // everything goes back
    let continued = vec![vec![], vec![], vec![Some("x".to_string())]];
    assert!(err(m.duplicate_table(1, 0, 7, Some(&continued))).contains("combinada"));
    assert_eq!(document_xml(&m), before);
}

#[test]
fn test_duplicate_table_leaves_out_what_must_stay_unique() {
    let cell = |inner: &str| format!("<w:tc><w:tcPr><w:shd w:val=\"clear\" w:fill=\"EEEEEE\"/></w:tcPr>{}</w:tc>", inner);
    let picture = r#"<w:r><w:drawing><wp:inline><wp:docPr id="7" name="Logo"/></wp:inline></w:drawing></w:r>"#;
    let table = format!(
        r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid><w:tr>{}{}</w:tr></w:tbl>"#,
        cell(&format!(
            r#"<w:p><w:bookmarkStart w:id="0" w:name="ficha"/><w:commentRangeStart w:id="3"/><w:r><w:t>Objetivo</w:t></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r><w:bookmarkEnd w:id="0"/><w:r><w:footnoteReference w:id="2"/></w:r>{}</w:p>"#,
            picture
        )),
        cell(r#"<w:sdt><w:sdtPr><w:id w:val="42"/><w:tag w:val="campo"/></w:sdtPr><w:sdtContent><w:p><w:ins w:id="5" w:author="A"><w:r><w:t>nuevo</w:t></w:r></w:ins><w:del w:id="6" w:author="A"><w:r><w:delText>viejo</w:delText></w:r></w:del></w:p></w:sdtContent></w:sdt>"#),
    );
    let mut m = doc_with(&format!("<w:p/>{}<w:p/>", table));
    let original = table_xml(&m, 0);
    let end = m.table_summaries().unwrap()[0].last + 1;
    m.duplicate_table(0, end, 0, None).unwrap();
    assert_eq!(table_xml(&m, 0), original);
    let copy = table_xml(&m, 1);
    for gone in ["bookmark", "comment", "footnoteReference", "<w:id ", "<w:ins", "<w:del", "viejo"] {
        assert!(!copy.contains(gone), "{} left in {}", gone, copy);
    }
    for kept in ["Objetivo", "nuevo", r#"<w:tag w:val="campo"/>"#, r#"w:fill="EEEEEE""#, r#"name="Logo""#] {
        assert!(copy.contains(kept), "{} missing in {}", kept, copy);
    }
    assert!(copy.contains(r#"<wp:docPr id="8""#), "the picture gets a new id: {}", copy);
    // Saved and opened again, the document reads both tables
    let reopened = DocxModifier::from_bytes(&m.to_bytes().unwrap()).unwrap();
    assert_eq!(rows_of(&reopened, 0), rows_of(&reopened, 1));

    // A table holding a logical section is refused
    let sectioned = r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:sdt><w:sdtPr><w:tag w:val="section:datos"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt></w:tc></w:tr></w:tbl>"#;
    let mut m = doc_with(&format!("<w:p/>{}<w:p/>", sectioned));
    assert!(m.duplicate_table(0, 0, 0, None).unwrap_err().contains("sección"));
}

#[test]
fn test_duplicate_table_is_one_undo_step() {
    let mut m = doc();
    let original = document_xml(&m);
    m.checkpoint(None);
    m.duplicate_table(0, 10, 3, Some(&[vec![Some("Uno".into())], vec![Some("Dos".into())]])).unwrap();
    let changed = document_xml(&m);
    assert!(m.undo().is_some());
    assert_eq!(document_xml(&m), original);
    m.redo();
    assert_eq!(document_xml(&m), changed);
}
