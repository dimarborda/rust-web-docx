//! Areas of a workbook an agent can point at and read as a unit. The file itself marks most of
//! them, and they are listed in this order:
//!
//! - **Tables** (`xl/tables/tableN.xml`, "Format as Table" in Excel): name, range, header and
//!   totals rows and column names
//! - **Defined names** that refer to a range (`Ventas = Datos!$A$3:$H$40`), and the print area
//! - **Filters** (`<autoFilter>` of a sheet)
//! - **Regions**: blocks of adjacent filled cells, as Ctrl+A selects them in Excel, for sheets
//!   that mark nothing. Their leading single-cell rows (or merged ones) are taken as titles and a
//!   row of texts above the data as the header.
//!
//! The context of an area (`area_context`) is what an agent needs to work on it without reading
//! the whole sheet: title, header, column types and totals, and the first rows as shown.

use std::collections::{HashMap, HashSet, VecDeque};

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use serde_json::{json, Value};

use ooxml_core::rels::{parse_relationships, rels_part_for, resolve_target, type_is};
use ooxml_core::xml::{element_ranges, get_attr_value, tag_is};

use crate::cell_ref::{column_name, CellRange, CellRef};
use crate::formula::{self, Expr};
use crate::numfmt::Locale;
use crate::recalc::defined_names;
use crate::sheet::SheetXml;
use crate::view::SheetLayout;
use crate::workbook::{CellData, CellValue, SheetKind, Workbook};

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum AreaKind {
    Table,
    Name,
    PrintArea,
    Filter,
    Region,
}

impl AreaKind {
    pub fn as_str(self) -> &'static str {
        match self {
            AreaKind::Table => "table",
            AreaKind::Name => "name",
            AreaKind::PrintArea => "printArea",
            AreaKind::Filter => "filter",
            AreaKind::Region => "region",
        }
    }
}

/// A range of one sheet that holds something as a unit
#[derive(Debug, Clone, PartialEq)]
pub struct Area {
    pub kind: AreaKind,
    /// Table or defined name; regions and filters have none
    pub name: Option<String>,
    pub sheet: usize,
    pub range: CellRange,
    /// Title of a region: texts of its leading single-cell rows, or of a lone cell just above it
    pub title: Vec<String>,
    /// Leading rows of the range taken by the title
    pub title_rows: u32,
    /// Row of the header, if any
    pub header_row: Option<u32>,
    /// Totals row of a table
    pub totals_row: Option<u32>,
    /// Column names (header texts, or a table's column names)
    pub columns: Vec<String>,
}

impl Area {
    /// Rows of data: below the titles and header, above the totals
    pub fn body(&self) -> Option<CellRange> {
        let first = self.header_row.map(|r| r + 1).unwrap_or(self.range.first.row + self.title_rows);
        let last = self.totals_row.map(|r| r - 1).unwrap_or(self.range.last.row);
        (first <= last).then_some(CellRange { first: CellRef { row: first, col: self.range.first.col }, last: CellRef { row: last, col: self.range.last.col } })
    }

    pub fn to_json(&self, sheet_name: &str) -> Value {
        json!({
            "kind": self.kind.as_str(),
            "name": self.name,
            "sheet": self.sheet,
            "sheetName": sheet_name,
            "ref": self.range.to_string(),
            "address": address(sheet_name, &self.range),
            "title": if self.title.is_empty() { Value::Null } else { json!(self.title.join(" · ")) },
            "headerRow": self.header_row,
            "totalsRow": self.totals_row,
            "body": self.body().map(|b| b.to_string()),
            "columns": self.columns,
        })
    }
}

/// `'Mis datos'!A1:C3`
pub fn address(sheet: &str, range: &CellRange) -> String {
    let plain = sheet.chars().all(|c| c.is_alphanumeric() || c == '_') && !sheet.chars().next().is_some_and(|c| c.is_ascii_digit());
    if plain {
        format!("{}!{}", sheet, range)
    } else {
        format!("'{}'!{}", sheet.replace('\'', "''"), range)
    }
}

/// A table part: `<table name displayName ref headerRowCount totalsRowCount><tableColumns>…`
pub fn parse_table(xml: &str) -> Option<(String, CellRange, u32, u32, Vec<String>)> {
    let mut reader = Reader::from_str(xml);
    let mut table = None;
    let mut columns = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let name = e.name();
                if tag_is(name.as_ref(), "table") {
                    let range = get_attr_value(&e, "ref").and_then(|r| CellRange::parse(&r).ok())?;
                    let title = get_attr_value(&e, "displayName").or_else(|| get_attr_value(&e, "name")).unwrap_or_default();
                    let count = |a: &str, default: u32| get_attr_value(&e, a).and_then(|v| v.trim().parse().ok()).unwrap_or(default);
                    table = Some((title, range, count("headerRowCount", 1), count("totalsRowCount", 0)));
                } else if tag_is(name.as_ref(), "tableColumn") {
                    columns.push(get_attr_value(&e, "name").unwrap_or_default());
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    table.map(|(name, range, header, totals)| (name, range, header, totals, columns))
}

/// Range of the `<autoFilter>` of a worksheet
fn sheet_filter(xml: &str) -> Option<CellRange> {
    let range = element_ranges(xml, "autoFilter", |_| true).ok()?.into_iter().next()?;
    let mut reader = Reader::from_str(&xml[range]);
    match reader.read_event() {
        Ok(Event::Start(e)) | Ok(Event::Empty(e)) => get_attr_value(&e, "ref").and_then(|r| CellRange::parse(&r).ok()),
        _ => None,
    }
}

/// The single range a defined name refers to: (sheet name, range)
fn name_range(text: &str) -> Option<(String, CellRange)> {
    match formula::parse(text).ok()? {
        Expr::Ref(area) if !area.whole_cols => {
            let range = CellRange::around(CellRef { row: area.first.row, col: area.first.col }, CellRef { row: area.last.row, col: area.last.col });
            Some((area.sheet?, range))
        }
        _ => None,
    }
}

/// Blocks of adjacent filled cells (diagonals count, as in Excel's current region), each as the
/// smallest rectangle around it, grown until no other filled cell touches it
pub fn regions(filled: &HashSet<CellRef>) -> Vec<CellRange> {
    let mut seen: HashSet<CellRef> = HashSet::new();
    let mut boxes: Vec<CellRange> = Vec::new();
    let mut sorted: Vec<&CellRef> = filled.iter().collect();
    sorted.sort();
    for start in sorted {
        if !seen.insert(*start) {
            continue;
        }
        let mut range = CellRange { first: *start, last: *start };
        let mut queue = VecDeque::from([*start]);
        while let Some(at) = queue.pop_front() {
            range = range.union(CellRange { first: at, last: at });
            for dr in -1i64..=1 {
                for dc in -1i64..=1 {
                    let (r, c) = (at.row as i64 + dr, at.col as i64 + dc);
                    if r < 1 || c < 1 {
                        continue;
                    }
                    let next = CellRef { row: r as u32, col: c as u32 };
                    if filled.contains(&next) && seen.insert(next) {
                        queue.push_back(next);
                    }
                }
            }
        }
        boxes.push(range);
    }
    // Rectangles that touch or overlap grow into one
    loop {
        let mut merged = false;
        'outer: for i in 0..boxes.len() {
            for j in i + 1..boxes.len() {
                let (a, b) = (boxes[i], boxes[j]);
                let touch = a.first.row <= b.last.row + 1 && b.first.row <= a.last.row + 1 && a.first.col <= b.last.col + 1 && b.first.col <= a.last.col + 1;
                if touch {
                    boxes[i] = a.union(b);
                    boxes.remove(j);
                    merged = true;
                    break 'outer;
                }
            }
        }
        if !merged {
            break;
        }
    }
    boxes.sort_by_key(|b| (b.first.row, b.first.col));
    boxes
}

pub(crate) fn overlaps_range(a: &CellRange, b: &CellRange) -> bool {
    a.first.row <= b.last.row && b.first.row <= a.last.row && a.first.col <= b.last.col && b.first.col <= a.last.col
}

pub(crate) fn range_contains(a: &CellRange, at: CellRef) -> bool {
    at.row >= a.first.row && at.row <= a.last.row && at.col >= a.first.col && at.col <= a.last.col
}

fn is_text(v: &CellValue) -> bool {
    matches!(v, CellValue::Text(t) if !t.trim().is_empty())
}

fn is_empty(v: &CellValue) -> bool {
    matches!(v, CellValue::Empty) || matches!(v, CellValue::Text(t) if t.trim().is_empty())
}

/// Title rows and header row of a block of cells read by row: leading rows with a single filled
/// cell (or one merged range) are titles; then a row of at least two texts, with no numbers,
/// followed by a row that is not all text (or shown in bold), is the header
fn structure(rows: &[Vec<CellData>], merges: &[CellRange], bold: &dyn Fn(CellRef) -> bool) -> (Vec<String>, Option<usize>) {
    let mut titles = Vec::new();
    let mut i = 0;
    while i + 1 < rows.len() {
        let filled: Vec<&CellData> = rows[i].iter().filter(|c| !is_empty(&c.value)).collect();
        let one_merge = filled.len() == 1 && merges.iter().any(|m| m.first == filled[0].at && m.first.row == m.last.row);
        let lone = filled.len() == 1 && (one_merge || rows[i].len() > 1);
        if !lone || !is_text(&filled[0].value) {
            break;
        }
        if let CellValue::Text(t) = &filled[0].value {
            titles.push(t.trim().to_string());
        }
        i += 1;
    }
    let header = rows.get(i).and_then(|row| {
        let filled: Vec<&CellData> = row.iter().filter(|c| !is_empty(&c.value)).collect();
        let texts = filled.iter().filter(|c| is_text(&c.value)).count();
        if texts < 2.min(row.len()) || texts != filled.len() {
            return None;
        }
        let next_has_data = rows.get(i + 1).is_some_and(|next| next.iter().any(|c| !is_empty(&c.value) && !is_text(&c.value)));
        let shown_bold = filled.iter().all(|c| bold(c.at));
        (next_has_data || shown_bold && rows.len() > i + 1).then_some(i)
    });
    (titles, header)
}

impl Workbook {
    /// Tables of a worksheet: (name, range, header rows, totals rows, column names)
    fn sheet_tables(&self, part: &str) -> Vec<(String, CellRange, u32, u32, Vec<String>)> {
        let Ok(rels) = self.package.get_string(&rels_part_for(part)) else { return Vec::new() };
        parse_relationships(&rels)
            .into_iter()
            .filter(|r| type_is(&r.rel_type, "table") && !r.external)
            .filter_map(|r| self.package.get_string(&resolve_target(part, &r.target)).ok())
            .filter_map(|xml| parse_table(&xml))
            .collect()
    }

    /// Every area of the workbook (see the module documentation), sheet by sheet
    pub fn areas(&self) -> Result<Vec<Area>, String> {
        let sheets = self.sheets()?;
        let names = defined_names(&self.workbook_xml()?);
        let mut out = Vec::new();
        for (index, sheet) in sheets.iter().enumerate() {
            if sheet.kind != SheetKind::Worksheet || !self.package.contains(&sheet.part) {
                continue;
            }
            let mut found: Vec<Area> = Vec::new();
            for (name, range, header, totals, columns) in self.sheet_tables(&sheet.part) {
                found.push(Area {
                    kind: AreaKind::Table,
                    name: Some(name),
                    sheet: index,
                    range,
                    title: Vec::new(),
                    title_rows: 0,
                    header_row: (header > 0).then_some(range.first.row),
                    totals_row: (totals > 0).then_some(range.last.row),
                    columns,
                });
            }
            for name in &names {
                let Some((sheet_name, range)) = name_range(&name.text) else { continue };
                if !sheet_name.eq_ignore_ascii_case(&sheet.name) {
                    continue;
                }
                let kind = match name.name.as_str() {
                    "_xlnm.Print_Area" => AreaKind::PrintArea,
                    n if n.starts_with("_xlnm.") => continue,
                    _ => AreaKind::Name,
                };
                let shown = if kind == AreaKind::PrintArea { None } else { Some(name.name.clone()) };
                found.push(self.described(index, range, kind, shown)?);
            }
            let xml = self.package.get_string(&sheet.part)?;
            if let Some(range) = sheet_filter(&xml) {
                if !found.iter().any(|a| a.kind == AreaKind::Table && a.range == range) {
                    found.push(self.described(index, range, AreaKind::Filter, None)?);
                }
            }
            // Regions only where nothing above already covers the cells
            let parsed = SheetXml::parse(&xml)?;
            let layout = SheetLayout::parse(&xml)?;
            let mut filled: HashSet<CellRef> = HashSet::new();
            for cell in parsed.rows.iter().flat_map(|r| &r.cells) {
                if cell.value.as_deref().is_some_and(|v| !v.trim().is_empty()) || cell.inline.as_deref().is_some_and(|v| !v.trim().is_empty()) || cell.formula.is_some() {
                    filled.insert(cell.at);
                }
            }
            for merge in &layout.merges {
                if filled.contains(&merge.first) {
                    for row in merge.first.row..=merge.last.row {
                        for col in merge.first.col..=merge.last.col {
                            filled.insert(CellRef { row, col });
                        }
                    }
                }
            }
            let blocks = regions(&filled);
            for range in &blocks {
                let range = *range;
                if range.cell_count() < 2 || found.iter().any(|a| a.range == range || matches!(a.kind, AreaKind::Table | AreaKind::Filter) && overlaps_range(&a.range, &range)) {
                    continue;
                }
                let mut area = self.described(index, range, AreaKind::Region, None)?;
                // A lone text one or two rows above, over the block's columns, is its title
                if area.title.is_empty() {
                    let above = blocks.iter().find(|b| {
                        b.first.row == b.last.row
                            && b.last.row < range.first.row
                            && b.last.row + 2 >= range.first.row
                            && b.first.col >= range.first.col
                            && b.first.col <= range.last.col
                            && (b.cell_count() == 1 || layout.merges.contains(*b))
                    });
                    if let Some(b) = above {
                        if let CellValue::Text(t) = self.cell(index, &b.first.to_string())?.value {
                            area.title.push(t.trim().to_string());
                        }
                    }
                }
                found.push(area);
            }
            out.extend(found);
        }
        Ok(out)
    }

    /// An area with its titles, header and columns worked out from the cells
    fn described(&self, sheet: usize, range: CellRange, kind: AreaKind, name: Option<String>) -> Result<Area, String> {
        let range = self.clip_to_used(sheet, range)?;
        let rows = self.range(sheet, &range.to_string())?;
        let xml = self.package.get_string(&self.worksheet_part(sheet)?)?;
        let merges = SheetLayout::parse(&xml)?.merges;
        let bold = self.bold_cells(sheet)?;
        let (title, header) = structure(&rows, &merges, &|at| bold.contains(&at));
        let header_row = header.map(|i| range.first.row + i as u32);
        let columns = match header {
            Some(i) => rows[i].iter().map(|c| match &c.value {
                CellValue::Text(t) => t.trim().to_string(),
                _ => String::new(),
            }).collect(),
            None => Vec::new(),
        };
        let title_rows = title.len() as u32;
        Ok(Area { kind, name, sheet, range, title, title_rows, header_row, totals_row: None, columns })
    }

    /// Cells shown in bold (to tell headers apart)
    fn bold_cells(&self, sheet: usize) -> Result<HashSet<CellRef>, String> {
        let styles = self.style_view()?;
        let xml = self.package.get_string(&self.worksheet_part(sheet)?)?;
        let parsed = SheetXml::parse(&xml)?;
        Ok(parsed
            .rows
            .iter()
            .flat_map(|r| &r.cells)
            .filter(|c| styles.xfs.get(c.style as usize).and_then(|x| styles.fonts.get(x.font)).is_some_and(|f| f.bold))
            .map(|c| c.at)
            .collect())
    }

    /// The range cut to the cells the sheet uses (so `A1:Z100000` reads only what is there)
    fn clip_to_used(&self, sheet: usize, range: CellRange) -> Result<CellRange, String> {
        let xml = self.package.get_string(&self.worksheet_part(sheet)?)?;
        let parsed = SheetXml::parse(&xml)?;
        let (mut max_row, mut max_col) = (0, 0);
        for cell in parsed.rows.iter().flat_map(|r| &r.cells) {
            max_row = max_row.max(cell.at.row);
            max_col = max_col.max(cell.at.col);
        }
        let last = CellRef { row: range.last.row.min(max_row.max(range.first.row)), col: range.last.col.min(max_col.max(range.first.col)) };
        Ok(CellRange { first: range.first, last })
    }

    /// The area `target` points at: the name of a table or defined name, an address with its
    /// sheet (`Ventas!A1:C9`, `'Mis datos'!B2`), or a range of the sheet `sheet`. A range that
    /// matches an area found in the file takes its description.
    pub fn resolve_area(&self, target: &str, sheet: Option<usize>) -> Result<Area, String> {
        let target = target.trim();
        let areas = self.areas()?;
        if let Some(area) = areas.iter().find(|a| a.name.as_deref().is_some_and(|n| n.eq_ignore_ascii_case(target))) {
            return Ok(area.clone());
        }
        let (sheet, range_text) = match target.rsplit_once('!') {
            Some((name, range)) => {
                let name = name.trim();
                let name = name.strip_prefix('\'').and_then(|n| n.strip_suffix('\'')).map(|n| n.replace("''", "'")).unwrap_or_else(|| name.to_string());
                (self.sheet_index(&name)?, range)
            }
            None => (
                sheet.ok_or_else(|| format!("'{}' no es el nombre de una tabla ni de un rango con nombre; indica la hoja (Hoja1!A1:C9).", target))?,
                target,
            ),
        };
        let range = CellRange::parse(range_text)?;
        if let Some(area) = areas.iter().find(|a| a.sheet == sheet && a.range == range) {
            return Ok(area.clone());
        }
        self.described(sheet, range, AreaKind::Region, None)
    }

    /// The region around a cell (Excel's current region), or the cell alone
    pub fn region_at(&self, sheet: usize, reference: &str) -> Result<Area, String> {
        let at = CellRef::parse(reference)?;
        let areas = self.areas()?;
        let best = areas
            .iter()
            .filter(|a| a.sheet == sheet && range_contains(&a.range, at))
            .min_by_key(|a| (a.kind != AreaKind::Table, a.range.cell_count()));
        match best {
            Some(area) => Ok(area.clone()),
            None => self.described(sheet, CellRange { first: at, last: at }, AreaKind::Region, None),
        }
    }

    /// What an agent needs to work on an area: its description, each column with its type,
    /// formulas and number totals, and the first `max_rows` rows of data as the cells show them
    /// in `locale`, also as a Markdown table
    pub fn area_context(&self, area: &Area, max_rows: usize, locale: &Locale) -> Result<Value, String> {
        let sheets = self.sheets()?;
        let sheet_name = sheets.get(area.sheet).map(|s| s.name.clone()).unwrap_or_default();
        let rows = self.range(area.sheet, &area.range.to_string())?;
        let shown: HashMap<(u32, u32), String> = self
            .cell_views(area.sheet, locale)?
            .into_iter()
            .filter(|v| range_contains(&area.range, CellRef { row: v.row, col: v.col }))
            .map(|v| ((v.row, v.col), v.text))
            .collect();
        let text_at = |at: CellRef| shown.get(&(at.row, at.col)).cloned().unwrap_or_default();
        let body = area.body();
        let body_rows: Vec<&Vec<CellData>> = rows
            .iter()
            .filter(|r| body.is_some_and(|b| r.first().is_some_and(|c| c.at.row >= b.first.row && c.at.row <= b.last.row)))
            .collect();

        let width = (area.range.last.col - area.range.first.col + 1) as usize;
        let mut columns = Vec::with_capacity(width);
        for i in 0..width {
            let col = area.range.first.col + i as u32;
            let mut kinds: Vec<&str> = Vec::new();
            let (mut count, mut sum, mut min, mut max) = (0usize, 0.0, f64::INFINITY, f64::NEG_INFINITY);
            let mut formulas = 0;
            let mut example_formula: Option<String> = None;
            let mut filled = 0;
            for row in &body_rows {
                let cell = &row[i];
                if cell.formula.is_some() {
                    formulas += 1;
                    if example_formula.is_none() {
                        example_formula = cell.formula.clone().map(|f| format!("{}: ={}", cell.at, f));
                    }
                }
                let kind = match &cell.value {
                    CellValue::Empty => continue,
                    CellValue::Text(t) if t.trim().is_empty() => continue,
                    CellValue::Number(n) => {
                        count += 1;
                        sum += n;
                        min = min.min(*n);
                        max = max.max(*n);
                        "number"
                    }
                    CellValue::Date(_) => "date",
                    CellValue::Text(_) => "text",
                    CellValue::Bool(_) => "boolean",
                    CellValue::Error(_) => "error",
                };
                filled += 1;
                if !kinds.contains(&kind) {
                    kinds.push(kind);
                }
            }
            let kind = match kinds.as_slice() {
                [] => "empty",
                [one] => one,
                _ if kinds.iter().all(|k| *k == "number" || *k == "error") => "number",
                _ => "mixed",
            };
            let mut column = json!({
                "col": column_name(col),
                "name": area.columns.get(i).cloned().unwrap_or_default(),
                "type": kind,
                "filled": filled,
            });
            if count > 0 {
                column["numbers"] = json!({ "count": count, "sum": sum, "min": min, "max": max, "average": sum / count as f64 });
            }
            if formulas > 0 {
                column["formulas"] = json!(formulas);
                column["exampleFormula"] = json!(example_formula);
            }
            columns.push(column);
        }

        let shown_rows: Vec<Vec<String>> = body_rows.iter().take(max_rows).map(|r| r.iter().map(|c| text_at(c.at)).collect()).collect();
        let totals: Option<Vec<String>> = area.totals_row.map(|row| (0..width).map(|i| text_at(CellRef { row, col: area.range.first.col + i as u32 })).collect());
        let merges: Vec<String> = SheetLayout::parse(&self.package.get_string(&self.worksheet_part(area.sheet)?)?)?
            .merges
            .iter()
            .filter(|m| overlaps_range(m, &area.range))
            .map(|m| m.to_string())
            .collect();

        // Markdown table: header (or column letters), the rows shown, the totals
        let escape = |t: &str| t.replace('|', "\\|").replace('\n', " ");
        let mut md = String::new();
        let header: Vec<String> = (0..width)
            .map(|i| {
                let letter = column_name(area.range.first.col + i as u32);
                match area.columns.get(i).filter(|n| !n.is_empty()) {
                    Some(name) => format!("{} ({})", escape(name), letter),
                    None => letter,
                }
            })
            .collect();
        md.push_str(&format!("| fila | {} |\n", header.join(" | ")));
        md.push_str(&format!("|---|{}\n", "---|".repeat(width)));
        for (row, cells) in body_rows.iter().take(max_rows).zip(&shown_rows) {
            let number = row.first().map(|c| c.at.row).unwrap_or(0);
            md.push_str(&format!("| {} | {} |\n", number, cells.iter().map(|t| escape(t)).collect::<Vec<_>>().join(" | ")));
        }
        if let (Some(totals), Some(row)) = (&totals, area.totals_row) {
            md.push_str(&format!("| {} (totales) | {} |\n", row, totals.iter().map(|t| escape(t)).collect::<Vec<_>>().join(" | ")));
        }

        let mut out = area.to_json(&sheet_name);
        out["rowCount"] = json!(body_rows.len());
        out["rowsShown"] = json!(shown_rows.len());
        out["truncated"] = json!(body_rows.len() > shown_rows.len());
        out["columnsInfo"] = Value::Array(columns);
        out["rows"] = json!(shown_rows);
        out["firstRow"] = json!(body.map(|b| b.first.row));
        out["totals"] = json!(totals);
        out["merges"] = json!(merges);
        out["markdown"] = json!(md);
        Ok(out)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::workbook::CellInput;

    fn set(filled: &[&str]) -> HashSet<CellRef> {
        filled.iter().map(|r| CellRef::parse(r).unwrap()).collect()
    }

    #[test]
    fn finds_regions_like_excel_current_region() {
        let found = regions(&set(&["A1", "B2", "C2", "A5", "B6", "E1", "E2"]));
        let refs: Vec<String> = found.iter().map(|r| r.to_string()).collect();
        assert_eq!(refs, ["A1:C2", "E1:E2", "A5:B6"]);
        // Diagonals connect, and touching rectangles grow into one
        assert_eq!(regions(&set(&["A1", "B2", "C3", "D2"]))[0].to_string(), "A1:D3");
    }

    #[test]
    fn reads_tables_and_defined_names() {
        let table = r#"<table xmlns="x" id="1" name="Tabla1" displayName="Ventas" ref="B3:D10" totalsRowCount="1"><autoFilter ref="B3:D9"/><tableColumns count="3"><tableColumn id="1" name="Producto"/><tableColumn id="2" name="Unidades"/><tableColumn id="3" name="Importe"/></tableColumns></table>"#;
        let (name, range, header, totals, columns) = parse_table(table).unwrap();
        assert_eq!((name.as_str(), range.to_string(), header, totals), ("Ventas", "B3:D10".to_string(), 1, 1));
        assert_eq!(columns, ["Producto", "Unidades", "Importe"]);
        assert_eq!(name_range("'Mis datos'!$A$1:$C$4"), Some(("Mis datos".to_string(), CellRange::parse("A1:C4").unwrap())));
        assert_eq!(name_range("0.21"), None);
    }

    fn report() -> Workbook {
        let mut wb = Workbook::from_bytes(&crate::blank::blank_xlsx().unwrap()).unwrap();
        let t = |s: &str| Some(CellInput::Text(s.to_string()));
        let n = |x: f64| Some(CellInput::Number(x));
        wb.set_values(0, "A1", vec![
            vec![t("Ventas del trimestre")],
            vec![t("Región"), t("Pedidos"), t("Importe")],
            vec![t("Norte"), n(3.0), n(120.5)],
            vec![t("Sur"), n(5.0), n(80.0)],
        ]).unwrap();
        wb.set_value(0, "C5", CellInput::Formula("SUM(C3:C4)".into())).unwrap();
        wb.set_values(0, "F1", vec![vec![t("Notas")], vec![t("revisar")]]).unwrap();
        wb
    }

    #[test]
    fn describes_regions_with_title_and_header() {
        let wb = report();
        let areas = wb.areas().unwrap();
        assert_eq!(areas.len(), 2);
        let sales = &areas[0];
        assert_eq!((sales.kind, sales.range.to_string()), (AreaKind::Region, "A1:C5".to_string()));
        assert_eq!(sales.title, ["Ventas del trimestre"]);
        assert_eq!(sales.header_row, Some(2));
        assert_eq!(sales.columns, ["Región", "Pedidos", "Importe"]);
        assert_eq!(sales.body().unwrap().to_string(), "A3:C5");
        assert_eq!(wb.region_at(0, "B4").unwrap().range.to_string(), "A1:C5");
        assert_eq!(wb.region_at(0, "F2").unwrap().range.to_string(), "F1:F2");
        assert_eq!(wb.resolve_area("Hoja1!A1:C5", None).unwrap(), sales.clone());
        assert!(wb.resolve_area("A1:C5", None).is_err());
    }

    #[test]
    fn lists_tables_names_print_areas_and_filters_first() {
        let mut wb = report();
        wb.package.put(
            "xl/worksheets/_rels/sheet1.xml.rels".into(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#.to_vec(),
        );
        wb.package.put(
            "xl/tables/table1.xml".into(),
            br#"<table xmlns="x" id="1" name="Tabla1" displayName="Ventas" ref="A2:C5" totalsRowCount="1"><tableColumns count="3"><tableColumn id="1" name="Region"/><tableColumn id="2" name="Pedidos"/><tableColumn id="3" name="Importe"/></tableColumns></table>"#.to_vec(),
        );
        let xml = wb.package.get_string("xl/workbook.xml").unwrap().replace(
            "</sheets>",
            r#"</sheets><definedNames><definedName name="Notas">Hoja1!$F$1:$F$2</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">Hoja1!$A$1:$F$5</definedName><definedName name="Tasa">0.21</definedName></definedNames>"#,
        );
        wb.package.put("xl/workbook.xml".into(), xml.into_bytes());

        let areas = wb.areas().unwrap();
        let summary: Vec<(&str, Option<&str>, String)> = areas.iter().map(|a| (a.kind.as_str(), a.name.as_deref(), a.range.to_string())).collect();
        assert_eq!(summary, [
            ("table", Some("Ventas"), "A2:C5".to_string()),
            ("name", Some("Notas"), "F1:F2".to_string()),
            ("printArea", None, "A1:F5".to_string()),
        ]);
        let table = wb.resolve_area("ventas", None).unwrap();
        assert_eq!((table.header_row, table.totals_row, table.body().unwrap().to_string()), (Some(2), Some(5), "A3:C4".to_string()));
        let context = wb.area_context(&table, 10, &Locale::from_tag("en")).unwrap();
        assert_eq!(context["totals"], json!(["", "", "200.5"]));
        assert_eq!(wb.region_at(0, "B3").unwrap().kind, AreaKind::Table);
    }

    #[test]
    fn gives_the_context_of_an_area() {
        let wb = report();
        let area = wb.resolve_area("A1:C5", Some(0)).unwrap();
        let context = wb.area_context(&area, 2, &Locale::from_tag("en")).unwrap();
        assert_eq!(context["address"], "Hoja1!A1:C5");
        assert_eq!(context["title"], "Ventas del trimestre");
        assert_eq!(context["rowCount"], 3);
        assert_eq!(context["truncated"], true);
        assert_eq!(context["rows"], json!([["Norte", "3", "120.5"], ["Sur", "5", "80"]]));
        let importe = &context["columnsInfo"][2];
        assert_eq!((importe["name"].as_str(), importe["type"].as_str()), (Some("Importe"), Some("number")));
        assert_eq!(importe["numbers"]["sum"], 401.0);
        assert_eq!(importe["formulas"], 1);
        assert_eq!(importe["exampleFormula"], "C5: =SUM(C3:C4)");
        let md = context["markdown"].as_str().unwrap();
        assert!(md.starts_with("| fila | Región (A) | Pedidos (B) | Importe (C) |\n|---|---|---|---|\n| 3 | Norte | 3 | 120.5 |"), "{}", md);
    }
}
