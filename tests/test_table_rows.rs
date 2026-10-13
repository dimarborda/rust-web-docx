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
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>{}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#,
        body()
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
