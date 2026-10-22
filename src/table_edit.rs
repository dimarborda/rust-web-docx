//! Filling tables that already exist in the document from code: listing the body tables,
//! writing a cell, and inserting rows (copying an existing row's formatting) or deleting them,
//! each as one undo step.
//!
//! Tables are numbered like `table_cell_paragraph_ranges`: top-level tables of the body in
//! order (tables nested in cells and in text boxes are not counted). Columns are grid columns,
//! so a cell merged across columns (`w:gridSpan`) is addressed by its first column.
//!
//! Child module of `docx_parser` so it can work on the package parts directly.

use super::*;
use crate::paragraph_edit::{element_end, element_name, is_element, prefix_of, tokenize, Token};
use std::ops::Range;

/// A body table as `DocxEditor.tables()` lists it
#[derive(Serialize, Debug, Clone, PartialEq)]
pub struct TableSummary {
    pub index: usize,
    pub rows: usize,
    pub cols: usize,
    /// Text of the first row by grid column ("" under the continuation of a merged cell)
    pub header: Vec<String>,
    /// First and last paragraph of the table's cells (body paragraph numbering)
    pub first: usize,
    pub last: usize,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum VMerge {
    None,
    Restart,
    Continue,
}

#[derive(Debug)]
struct CellLoc {
    /// Token range of the `w:tc` element
    tok: (usize, usize),
    grid: usize,
    span: usize,
    vmerge: VMerge,
    hmerge: bool,
}

#[derive(Debug)]
struct RowLoc {
    tok: (usize, usize),
    cells: Vec<CellLoc>,
    /// Inside a content control or custom XML element instead of directly in the table
    wrapped: bool,
}

impl RowLoc {
    fn has_vmerge(&self, kind: VMerge) -> bool {
        self.cells.iter().any(|c| c.vmerge == kind)
    }
}

#[derive(Debug)]
struct TableLoc {
    tok: (usize, usize),
    rows: Vec<RowLoc>,
    cols: usize,
}

fn val_attr(tok: &Token) -> Option<String> {
    match &tok.ev {
        Event::Start(e) | Event::Empty(e) => get_attr_value(e, "val"),
        _ => None,
    }
}

fn byte_range(tokens: &[Token], (start, end): (usize, usize)) -> Range<usize> {
    tokens[start].span.start..tokens[end].span.end
}

/// Top-level body tables, in order
fn locate_tables(tokens: &[Token]) -> Vec<TableLoc> {
    let mut tables = Vec::new();
    let mut in_body = false;
    let mut i = 0;
    while i < tokens.len() {
        match &tokens[i].ev {
            Event::Start(e) if tag_is(e.name().as_ref(), "body") => in_body = true,
            Event::Start(e) if in_body && tag_is(e.name().as_ref(), "p") => {
                i = element_end(tokens, i) + 1;
                continue;
            }
            Event::Start(e) if in_body && tag_is(e.name().as_ref(), "tbl") => {
                let end = element_end(tokens, i);
                tables.push(read_table(tokens, i, end));
                i = end + 1;
                continue;
            }
            _ => {}
        }
        i += 1;
    }
    tables
}

fn read_table(tokens: &[Token], start: usize, end: usize) -> TableLoc {
    let mut rows = Vec::new();
    let mut grid_cols = 0;
    let mut depth = 0usize;
    let mut i = start + 1;
    while i < end {
        let tok = &tokens[i];
        if is_element(tok, "tr") {
            let tr_end = element_end(tokens, i);
            rows.push(read_row(tokens, i, tr_end, depth > 0));
            i = tr_end + 1;
            continue;
        }
        if is_element(tok, "tblGrid") {
            let grid_end = element_end(tokens, i);
            grid_cols = (i..grid_end).filter(|&j| is_element(&tokens[j], "gridCol")).count();
            i = grid_end + 1;
            continue;
        }
        if is_element(tok, "tblPr") || is_element(tok, "sdtPr") || is_element(tok, "sdtEndPr") {
            i = element_end(tokens, i) + 1;
            continue;
        }
        match tok.ev {
            Event::Start(_) => depth += 1,
            Event::End(_) => depth = depth.saturating_sub(1),
            _ => {}
        }
        i += 1;
    }
    let used = rows
        .iter()
        .map(|r| r.cells.last().map(|c| c.grid + c.span).unwrap_or(0))
        .max()
        .unwrap_or(0);
    TableLoc { tok: (start, end), rows, cols: grid_cols.max(used) }
}

fn read_row(tokens: &[Token], start: usize, end: usize, wrapped: bool) -> RowLoc {
    let mut cells = Vec::new();
    let mut grid = 0;
    let mut i = start + 1;
    while i < end {
        let tok = &tokens[i];
        if is_element(tok, "trPr") {
            let pr_end = element_end(tokens, i);
            if let Some(j) = (i..pr_end).find(|&j| is_element(&tokens[j], "gridBefore")) {
                grid = val_attr(&tokens[j]).and_then(|v| v.parse().ok()).unwrap_or(0);
            }
            i = pr_end + 1;
            continue;
        }
        if is_element(tok, "tc") {
            let tc_end = element_end(tokens, i);
            let mut span = 1;
            let mut vmerge = VMerge::None;
            let mut hmerge = false;
            if let Some(pr) = (i + 1..tc_end).find(|&j| is_element(&tokens[j], "tcPr")) {
                let pr_end = element_end(tokens, pr);
                for tok in &tokens[pr..pr_end] {
                    if is_element(tok, "gridSpan") {
                        span = val_attr(tok).and_then(|v| v.parse().ok()).unwrap_or(1).max(1);
                    } else if is_element(tok, "vMerge") {
                        vmerge = if val_attr(tok).as_deref() == Some("restart") { VMerge::Restart } else { VMerge::Continue };
                    } else if is_element(tok, "hMerge") {
                        hmerge = true;
                    }
                }
            }
            cells.push(CellLoc { tok: (i, tc_end), grid, span, vmerge, hmerge });
            grid += span;
            i = tc_end + 1;
            continue;
        }
        if is_element(tok, "tblPrEx") || is_element(tok, "sdtPr") {
            i = element_end(tokens, i) + 1;
            continue;
        }
        i += 1;
    }
    RowLoc { tok: (start, end), cells, wrapped }
}

/// Byte ranges of the paragraphs of a cell; refuses cells holding a nested table
fn cell_paragraph_ranges(tokens: &[Token], cell: &CellLoc) -> Result<Vec<Range<usize>>, String> {
    let (start, end) = cell.tok;
    let mut ranges = Vec::new();
    let mut j = start + 1;
    while j < end {
        if is_element(&tokens[j], "tbl") {
            return Err("La celda contiene una tabla anidada; las tablas anidadas no se pueden editar desde código.".to_string());
        }
        if is_element(&tokens[j], "p") {
            let p_end = element_end(tokens, j);
            ranges.push(byte_range(tokens, (j, p_end)));
            j = p_end + 1;
            continue;
        }
        j += 1;
    }
    if ranges.is_empty() {
        return Err("La celda no tiene párrafos.".to_string());
    }
    Ok(ranges)
}

/// Removes the direct children named `names` (local names) from an element's XML
fn remove_children(fragment: &str, names: &[&str]) -> Result<String, String> {
    let tokens = tokenize(fragment)?;
    let mut cuts = Vec::new();
    let mut i = 1;
    while i + 1 < tokens.len() {
        if let Some(name) = element_name(&tokens[i]) {
            let end = element_end(&tokens, i);
            if names.iter().any(|n| tag_is(name, n)) {
                cuts.push(byte_range(&tokens, (i, end)));
            }
            i = end + 1;
            continue;
        }
        i += 1;
    }
    let mut out = fragment.to_string();
    for range in cuts.into_iter().rev() {
        out.replace_range(range, "");
    }
    Ok(out)
}

/// XML of the first element named `name` (local name) between tokens `from` and `to`
fn first_element<'x>(xml: &'x str, tokens: &[Token], from: usize, to: usize, name: &str) -> Option<&'x str> {
    let j = (from..to).find(|&j| is_element(&tokens[j], name))?;
    Some(&xml[byte_range(tokens, (j, element_end(tokens, j)))])
}

/// Run content for `text`: "\n" becomes a line break and "\t" a tab
fn run_content(p: &str, text: &str) -> String {
    let mut out = String::new();
    for (k, line) in text.split('\n').enumerate() {
        if k > 0 {
            out.push_str(&format!("<{p}br/>"));
        }
        for (m, piece) in line.split('\t').enumerate() {
            if m > 0 {
                out.push_str(&format!("<{p}tab/>"));
            }
            if !piece.is_empty() {
                out.push_str(&format!(r#"<{p}t xml:space="preserve">{}</{p}t>"#, escape_xml(piece)));
            }
        }
    }
    out
}

/// A new cell like `cell`: its properties, its first paragraph's properties and its first
/// run's formatting, holding `text`
fn cell_like(xml: &str, tokens: &[Token], cell: &CellLoc, p: &str, text: &str) -> Result<String, String> {
    let (start, end) = cell.tok;
    let tc_pr = match (start + 1..end).find(|&j| is_element(&tokens[j], "tcPr")) {
        Some(j) => remove_children(&xml[byte_range(tokens, (j, element_end(tokens, j)))], &["cnfStyle", "tcPrChange", "cellIns", "cellDel", "cellMerge"])?,
        None => String::new(),
    };
    let (mut p_pr, mut r_pr) = (String::new(), String::new());
    if let Some(pi) = (start + 1..end).find(|&j| is_element(&tokens[j], "p")) {
        let p_end = element_end(tokens, pi);
        if let Some(pr) = first_element(xml, tokens, pi + 1, p_end, "pPr") {
            p_pr = remove_children(pr, &["cnfStyle", "pPrChange", "sectPr"])?;
        }
        // The first run's formatting, or the paragraph mark's when the paragraph is empty
        let run_pr = (pi + 1..p_end)
            .find(|&j| is_element(&tokens[j], "r"))
            .and_then(|r| first_element(xml, tokens, r + 1, element_end(tokens, r), "rPr"));
        let mark_pr = (pi + 1..p_end)
            .find(|&j| is_element(&tokens[j], "pPr"))
            .and_then(|pr| first_element(xml, tokens, pr + 1, element_end(tokens, pr), "rPr"));
        if let Some(pr) = run_pr.or(mark_pr) {
            r_pr = remove_children(pr, &["rPrChange", "ins", "del", "moveFrom", "moveTo"])?;
        }
    }
    // Text typed later into an empty cell takes the paragraph mark's formatting
    if !r_pr.is_empty() && !p_pr.contains(&format!("<{p}rPr")) {
        let mark = r_pr.clone();
        if p_pr.is_empty() {
            p_pr = format!("<{p}pPr>{mark}</{p}pPr>");
        } else if let Some(at) = p_pr.rfind(&format!("</{p}pPr>")) {
            p_pr.insert_str(at, &mark);
        } else {
            // <w:pPr/>
            p_pr = format!("<{p}pPr>{mark}</{p}pPr>");
        }
    }
    let run = if text.is_empty() { String::new() } else { format!("<{p}r>{r_pr}{}</{p}r>", run_content(p, text)) };
    Ok(format!("<{p}tc>{tc_pr}<{p}p>{p_pr}{run}</{p}p></{p}tc>"))
}

impl DocxModifier {
    fn table_at(tables: Vec<TableLoc>, table: usize) -> Result<TableLoc, String> {
        let count = tables.len();
        tables.into_iter().nth(table).ok_or_else(|| match count {
            0 => "El documento no tiene tablas.".to_string(),
            1 => format!("No existe la tabla {}: el documento tiene una sola tabla (índice 0).", table),
            n => format!("No existe la tabla {}: el documento tiene {} tablas (índices 0 a {}).", table, n, n - 1),
        })
    }

    fn row_error(table: usize, rows: usize, row: usize) -> String {
        format!("No existe la fila {} en la tabla {}: tiene {} filas (índices 0 a {}).", row, table, rows, rows.saturating_sub(1))
    }

    /// The body tables with their size, first row text and paragraph range
    pub fn table_summaries(&self) -> Result<Vec<TableSummary>, String> {
        let xml = self.get_file_string("word/document.xml")?;
        let tokens = tokenize(&xml)?;
        let paragraphs = body_paragraph_ranges(&xml)?;
        let mut out = Vec::new();
        for (index, table) in locate_tables(&tokens).iter().enumerate() {
            let bytes = byte_range(&tokens, table.tok);
            let inside: Vec<usize> = paragraphs
                .iter()
                .enumerate()
                .filter(|(_, r)| r.start >= bytes.start && r.end <= bytes.end)
                .map(|(i, _)| i)
                .collect();
            let (Some(&first), Some(&last)) = (inside.first(), inside.last()) else { continue };
            let mut header = vec![String::new(); table.cols];
            if let Some(row) = table.rows.first() {
                for cell in &row.cells {
                    let text = match cell_paragraph_ranges(&tokens, cell) {
                        Ok(ranges) => ranges
                            .iter()
                            .map(|r| parse_paragraph_fragment(&xml[r.clone()], &self.styles).text)
                            .collect::<Vec<_>>()
                            .join("\n"),
                        Err(_) => String::new(),
                    };
                    if let Some(slot) = header.get_mut(cell.grid) {
                        *slot = text;
                    }
                }
            }
            out.push(TableSummary { index, rows: table.rows.len(), cols: table.cols, header, first, last });
        }
        Ok(out)
    }

    /// Replaces the text of cell (`row`, `col`) of body table `table`, keeping the cell's
    /// paragraph properties and run formatting; "\n" is a line break inside the cell. `col` is
    /// a grid column: a horizontally merged cell is written through its first column, and a
    /// cell continuing a vertical merge is refused (write the first row of the merge).
    pub fn set_table_cell(&mut self, table: usize, row: usize, col: usize, text: &str) -> Result<(), String> {
        let mut xml = self.get_file_string("word/document.xml")?;
        let ranges = {
            let tokens = tokenize(&xml)?;
            let loc = Self::table_at(locate_tables(&tokens), table)?;
            let r = loc.rows.get(row).ok_or_else(|| Self::row_error(table, loc.rows.len(), row))?;
            if col >= loc.cols {
                return Err(format!(
                    "No existe la columna {} en la tabla {}: tiene {} columnas (índices 0 a {}).",
                    col, table, loc.cols, loc.cols.saturating_sub(1)
                ));
            }
            let cell = r
                .cells
                .iter()
                .find(|c| col >= c.grid && col < c.grid + c.span)
                .ok_or_else(|| format!("La fila {} de la tabla {} no tiene celda en la columna {}.", row, table, col))?;
            if cell.grid != col {
                return Err(format!(
                    "La celda ({}, {}) de la tabla {} está combinada con la columna {}; escribe en la columna {}.",
                    row, col, table, cell.grid, cell.grid
                ));
            }
            if cell.hmerge {
                return Err(format!("La celda ({}, {}) de la tabla {} está combinada horizontalmente; no se puede escribir en ella.", row, col, table));
            }
            if cell.vmerge == VMerge::Continue {
                let first = (0..row)
                    .rev()
                    .find(|&k| loc.rows[k].cells.iter().any(|c| c.grid == col && c.vmerge != VMerge::Continue))
                    .unwrap_or(0);
                return Err(format!(
                    "La celda ({}, {}) de la tabla {} está combinada verticalmente con la de la fila {}; escribe en la fila {}.",
                    row, col, table, first, first
                ));
            }
            cell_paragraph_ranges(&tokens, cell)?
        };
        self.replace_cell_text(&mut xml, &ranges, text)?;
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        Ok(())
    }

    /// Inserts `rows` (cell texts) before row `at` of body table `table` (`at` = row count
    /// appends them). Each new row copies row `copy_from`'s properties, cell properties (borders,
    /// shading, widths), first paragraph's properties and first run's formatting; by default
    /// the row before `at`, or the next one when inserting at the top. Short rows get empty
    /// cells. Returns the index of the first new row.
    pub fn insert_table_rows(&mut self, table: usize, at: usize, rows: &[Vec<String>], copy_from: Option<usize>) -> Result<usize, String> {
        if rows.is_empty() {
            return Err("No hay filas para insertar.".to_string());
        }
        let mut xml = self.get_file_string("word/document.xml")?;
        let (position, fragment) = {
            let tokens = tokenize(&xml)?;
            let loc = Self::table_at(locate_tables(&tokens), table)?;
            let count = loc.rows.len();
            if count == 0 {
                return Err(format!("La tabla {} no tiene filas.", table));
            }
            if at > count {
                return Err(format!(
                    "No se puede insertar en la posición {} de la tabla {}: tiene {} filas (posiciones 0 a {}).",
                    at, table, count, count
                ));
            }
            if loc.rows.iter().any(|r| r.wrapped) {
                return Err(format!("La tabla {} tiene filas dentro de controles de contenido; no se pueden insertar filas en ella.", table));
            }
            let template_index = copy_from.unwrap_or(if at > 0 { at - 1 } else { 0 });
            let template = loc.rows.get(template_index).ok_or_else(|| Self::row_error(table, count, template_index))?;
            if template.has_vmerge(VMerge::Restart) || template.has_vmerge(VMerge::Continue) || template.cells.iter().any(|c| c.hmerge) {
                return Err(format!(
                    "La fila {} de la tabla {} tiene celdas combinadas verticalmente; elige otra fila con copyFormatFrom.",
                    template_index, table
                ));
            }
            if at > 0 && at < count && loc.rows[at].has_vmerge(VMerge::Continue) {
                return Err(format!(
                    "La posición {} de la tabla {} queda dentro de celdas combinadas verticalmente; inserta antes o después de ellas.",
                    at, table
                ));
            }
            let cells = template.cells.len();
            if cells == 0 {
                return Err(format!("La fila {} de la tabla {} no tiene celdas.", template_index, table));
            }
            if let Some((k, row)) = rows.iter().enumerate().find(|(_, r)| r.len() > cells) {
                return Err(format!(
                    "La fila {} tiene {} valores, pero las filas de la tabla {} tienen {} celdas.",
                    k, row.len(), table, cells
                ));
            }

            let (tr_start, tr_end) = template.tok;
            let p = match element_name(&tokens[tr_start]).map(prefix_of) {
                Some(prefix) if !prefix.is_empty() => format!("{}:", prefix),
                _ => String::new(),
            };
            let ids = Regex::new(r#"\s+w14:(?:paraId|textId)="[^"]*""#).unwrap();
            let open = ids.replace_all(&xml[tokens[tr_start].span.clone()], "").into_owned();
            let ex = first_element(&xml, &tokens, tr_start + 1, tr_end, "tblPrEx").unwrap_or("");
            let mut drop = vec!["cnfStyle", "ins", "del", "trPrChange"];
            if at > 0 {
                // Repeated header rows must stay at the top
                drop.push("tblHeader");
            }
            let tr_pr = match first_element(&xml, &tokens, tr_start + 1, tr_end, "trPr") {
                Some(pr) => remove_children(pr, &drop)?,
                None => String::new(),
            };
            let mut fragment = String::new();
            for values in rows {
                fragment.push_str(&open);
                fragment.push_str(ex);
                fragment.push_str(&tr_pr);
                for (k, cell) in template.cells.iter().enumerate() {
                    let text = values.get(k).map(|s| s.replace("\r\n", "\n").replace('\r', "\n")).unwrap_or_default();
                    fragment.push_str(&cell_like(&xml, &tokens, cell, &p, &text)?);
                }
                fragment.push_str(&format!("</{p}tr>"));
            }
            let position = if at < count {
                tokens[loc.rows[at].tok.0].span.start
            } else {
                tokens[loc.rows[count - 1].tok.1].span.end
            };
            (position, fragment)
        };
        xml.insert_str(position, &fragment);
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        Ok(at)
    }

    /// Deletes `count` rows of body table `table` from row `from`. The table keeps at least one
    /// row, and vertical merges crossing the deleted range are refused.
    pub fn delete_table_rows(&mut self, table: usize, from: usize, count: usize) -> Result<(), String> {
        if count == 0 {
            return Err("No hay filas para eliminar (count = 0).".to_string());
        }
        let mut xml = self.get_file_string("word/document.xml")?;
        let cuts = {
            let tokens = tokenize(&xml)?;
            let loc = Self::table_at(locate_tables(&tokens), table)?;
            let rows = loc.rows.len();
            if from >= rows || from + count > rows {
                return Err(format!(
                    "No se pueden eliminar las filas {} a {} de la tabla {}: tiene {} filas (índices 0 a {}).",
                    from, from + count - 1, table, rows, rows.saturating_sub(1)
                ));
            }
            if count >= rows {
                return Err(format!("La tabla {} no puede quedar sin filas; elimina la tabla completa con replaceParagraphs.", table));
            }
            if loc.rows.iter().any(|r| r.wrapped) {
                return Err(format!("La tabla {} tiene filas dentro de controles de contenido; no se pueden eliminar filas de ella.", table));
            }
            let crosses_above = loc.rows[from].has_vmerge(VMerge::Continue);
            let crosses_below = loc.rows.get(from + count).is_some_and(|r| r.has_vmerge(VMerge::Continue));
            if crosses_above || crosses_below {
                return Err(format!(
                    "Las filas {} a {} de la tabla {} cortan celdas combinadas verticalmente; amplía el rango para incluirlas completas.",
                    from, from + count - 1, table
                ));
            }
            loc.rows[from..from + count].iter().map(|r| byte_range(&tokens, r.tok)).collect::<Vec<_>>()
        };
        for range in cuts.into_iter().rev() {
            xml.replace_range(range, "");
        }
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        Ok(())
    }

    /// Inserts a copy of body table `table` at (`index`, `offset`), as `insert_table` places a
    /// new one: same table, row and cell properties, merged cells, widths and text. `values`
    /// rewrites cells of the copy by row and grid column, as `set_table_cell` does (`None`
    /// keeps the copied text). What must stay unique in a document is left out of the copy:
    /// bookmarks, comment marks, footnote and endnote references, editing permissions and
    /// content control ids; tracked changes are accepted in it and pictures get new ids. A
    /// table holding a logical section is refused. Returns the copy's table index, its first
    /// cell's paragraph and the caret at the start of the paragraph after it.
    pub fn duplicate_table(
        &mut self,
        table: usize,
        index: usize,
        offset: usize,
        values: Option<&[Vec<Option<String>>]>,
    ) -> Result<(usize, usize, (usize, usize)), String> {
        let original = self.get_file_string("word/document.xml")?;
        let copy = {
            let tokens = tokenize(&original)?;
            let loc = Self::table_at(locate_tables(&tokens), table)?;
            if let Some(values) = values {
                if values.len() > loc.rows.len() {
                    return Err(format!("Hay valores para {} filas, pero la tabla {} tiene {}.", values.len(), table, loc.rows.len()));
                }
                if let Some((r, row)) = values.iter().enumerate().find(|(_, row)| row.len() > loc.cols) {
                    return Err(format!("La fila {} tiene {} valores, pero la tabla {} tiene {} columnas.", r, row.len(), table, loc.cols));
                }
            }
            clean_table_copy(&original[byte_range(&tokens, loc.tok)])?
        };
        let target = self.table_insertion_point(index, offset)?;
        let result = self.place_table_copy(&copy, target, values);
        if result.is_err() {
            // The split paragraph and any cell already written go back as they were
            self.put_file("word/document.xml".to_string(), original.into_bytes());
        }
        result
    }

    fn place_table_copy(
        &mut self,
        copy: &str,
        target: usize,
        values: Option<&[Vec<Option<String>>]>,
    ) -> Result<(usize, usize, (usize, usize)), String> {
        let mut xml = self.get_file_string("word/document.xml")?;
        let mut next_id = super::insert_objects::next_doc_pr_id(&xml);
        let doc_pr = Regex::new(r#"(<wp:docPr\b[^>]*\bid=")\d+""#).unwrap();
        let copy = doc_pr.replace_all(copy, |c: &regex::Captures| {
            let out = format!(r#"{}{}""#, &c[1], next_id);
            next_id += 1;
            out
        });
        let at = self.body_paragraph(&xml, target)?.start;
        xml.insert_str(at, &copy);
        let new_table = {
            let tokens = tokenize(&xml)?;
            locate_tables(&tokens)
                .iter()
                .position(|t| tokens[t.tok.0].span.start == at)
                .ok_or("No se pudo ubicar la copia de la tabla.")?
        };
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        for (row, cells) in values.unwrap_or_default().iter().enumerate() {
            for (col, text) in cells.iter().enumerate() {
                if let Some(text) = text {
                    self.set_table_cell(new_table, row, col, text)?;
                }
            }
        }
        let xml = self.get_file_string("word/document.xml")?;
        let end = {
            let tokens = tokenize(&xml)?;
            let loc = Self::table_at(locate_tables(&tokens), new_table)?;
            tokens[loc.tok.1].span.end
        };
        let after = body_paragraph_ranges(&xml)?.partition_point(|r| r.start < end);
        Ok((new_table, target, (after, 0)))
    }
}

/// Elements left out of a copied table: marks whose ids must be unique in the document,
/// references to notes and comments, and tracked deletions and property changes (accepted)
const COPY_DROPS: &[&str] = &[
    "bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "commentReference",
    "footnoteReference", "endnoteReference", "permStart", "permEnd", "del", "moveFrom",
    "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd", "rPrChange",
    "pPrChange", "tcPrChange", "trPrChange", "tblPrChange", "tblPrExChange", "tblGridChange",
    "sectPrChange", "numberingChange", "cellIns", "cellDel", "cellMerge",
];
/// Elements whose content stays in a copied table without them (tracked insertions, accepted)
const COPY_UNWRAPS: &[&str] = &["ins", "moveTo"];

/// The XML of a table ready to be inserted again in the same document (see
/// `DocxModifier::duplicate_table`)
fn clean_table_copy(fragment: &str) -> Result<String, String> {
    let tokens = tokenize(fragment)?;
    let mut cuts: Vec<Range<usize>> = Vec::new();
    let mut in_sdt_pr = 0usize;
    let mut i = 0;
    while i < tokens.len() {
        let tok = &tokens[i];
        let Some(name) = element_name(tok) else {
            if matches!(&tok.ev, Event::End(e) if tag_is(e.name().as_ref(), "sdtPr")) {
                in_sdt_pr = in_sdt_pr.saturating_sub(1);
            }
            i += 1;
            continue;
        };
        let end = element_end(&tokens, i);
        if tag_is(name, "tag") && val_attr(tok).is_some_and(|v| v.starts_with(super::section_edit::SECTION_TAG_PREFIX)) {
            return Err("La tabla contiene una sección lógica; quítala con unwrapSection antes de duplicar la tabla.".to_string());
        }
        // A content control's id: Word gives the copy a new one
        let sdt_id = in_sdt_pr > 0 && tag_is(name, "id");
        if sdt_id || COPY_DROPS.iter().any(|n| tag_is(name, n)) {
            cuts.push(byte_range(&tokens, (i, end)));
            i = end + 1;
            continue;
        }
        if COPY_UNWRAPS.iter().any(|n| tag_is(name, n)) {
            if end == i {
                // An empty mark (e.g. an inserted paragraph mark in w:rPr)
                cuts.push(tokens[i].span.clone());
            } else {
                cuts.push(tokens[i].span.clone());
                cuts.push(tokens[end].span.clone());
            }
        }
        if tag_is(name, "sdtPr") && end > i {
            in_sdt_pr += 1;
        }
        i += 1;
    }
    cuts.sort_by_key(|r| r.start);
    let mut out = String::with_capacity(fragment.len());
    let mut pos = 0;
    for cut in cuts {
        out.push_str(&fragment[pos..cut.start]);
        pos = cut.end;
    }
    out.push_str(&fragment[pos..]);
    let ids = Regex::new(r#"\s+w14:(?:paraId|textId)="[^"]*""#).unwrap();
    Ok(ids.replace_all(&out, "").into_owned())
}
