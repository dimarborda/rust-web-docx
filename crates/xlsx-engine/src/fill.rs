//! Filling a range from the cells at its start, as Excel's fill handle does. Along each row or
//! column of the fill the source cells repeat, with their format, and:
//!
//! - formulas are written again for each cell (relative references move);
//! - with `series`, numbers continue their trend (one date goes on day by day; one plain number
//!   is copied), a text ending in a number counts on (`Lote 9`, `Lote 10`) and names of months and
//!   days follow their order, in Spanish or English.
//!
//! A cell that does not fit any of these is copied as it is.

use std::collections::{BTreeMap, HashMap};

use ooxml_core::xml::{element_ranges, set_attr};

use crate::areas::{overlaps_range, range_contains};
use crate::cell_ref::{CellRange, CellRef};
use crate::moves::{shift_formula, without_shared_sources};
use crate::shared_strings::{SharedStrings, EMPTY_SST};
use crate::sheet::{CellNode, SheetXml};
use crate::view::SheetLayout;
use crate::workbook::{cell_element, format_number, formula_element, Workbook, MAX_CELLS_PER_CALL, SHARED_STRINGS_TYPE};

/// Lists Excel continues; the short forms come after the long ones
const LISTS: &[&[&str]] = &[
    &["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"],
    &["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"],
    &["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"],
    &["lun", "mar", "mié", "jue", "vie", "sáb", "dom"],
    &["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"],
    &["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"],
    &["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"],
    &["mon", "tue", "wed", "thu", "fri", "sat", "sun"],
];

/// What fills one row or column
#[derive(Debug, PartialEq)]
enum Plan {
    /// The source cells repeat
    Copy,
    /// `start + step * position`
    Number { start: f64, step: f64 },
    /// A text ending in a counter: `prefix`, then `start + step * position` with at least `width` digits
    Counter { prefix: String, start: i64, step: i64, width: usize },
    /// Names of a list, `start + step * position`, modulo its length
    Names { list: &'static [&'static str], start: i64, step: i64, case: Case },
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum Case {
    Lower,
    Upper,
    Capitalized,
}

/// Least squares over positions 0, 1, 2…: (value at 0, step)
fn trend(values: &[f64]) -> (f64, f64) {
    let n = values.len() as f64;
    if values.len() < 2 {
        return (values.first().copied().unwrap_or(0.0), 0.0);
    }
    let mean_x = (n - 1.0) / 2.0;
    let mean_y = values.iter().sum::<f64>() / n;
    let mut top = 0.0;
    let mut bottom = 0.0;
    for (i, y) in values.iter().enumerate() {
        top += (i as f64 - mean_x) * (y - mean_y);
        bottom += (i as f64 - mean_x).powi(2);
    }
    let step = top / bottom;
    (mean_y - step * mean_x, step)
}

/// `Lote 09` → (`Lote `, 9, 2)
fn split_counter(text: &str) -> Option<(&str, i64, usize)> {
    let start = text.trim_end_matches(|c: char| c.is_ascii_digit()).len();
    let digits = &text[start..];
    if digits.is_empty() || digits.len() > 15 {
        return None;
    }
    Some((&text[..start], digits.parse().ok()?, if digits.starts_with('0') { digits.len() } else { 0 }))
}

fn names_plan(texts: &[&str]) -> Option<Plan> {
    let first = *texts.first()?;
    let case = if first.chars().count() > 1 && first == first.to_uppercase() {
        Case::Upper
    } else if first.chars().next().is_some_and(char::is_uppercase) {
        Case::Capitalized
    } else {
        Case::Lower
    };
    'lists: for list in LISTS {
        let mut places = Vec::with_capacity(texts.len());
        for text in texts {
            let lower = text.to_lowercase();
            match list.iter().position(|name| *name == lower) {
                Some(place) => places.push(place as i64),
                None => continue 'lists,
            }
        }
        let length = list.len() as i64;
        let step = if places.len() > 1 { (places[1] - places[0]).rem_euclid(length) } else { 1 };
        if places.windows(2).all(|w| (w[1] - w[0]).rem_euclid(length) == step) {
            return Some(Plan::Names { list, start: places[0], step, case });
        }
    }
    None
}

fn text_plan(texts: &[&str]) -> Plan {
    if let Some(plan) = names_plan(texts) {
        return plan;
    }
    let parts: Option<Vec<_>> = texts.iter().map(|t| split_counter(t)).collect();
    let Some(parts) = parts else { return Plan::Copy };
    let Some((prefix, _, width)) = parts.first().copied() else { return Plan::Copy };
    if parts.iter().any(|p| p.0 != prefix) {
        return Plan::Copy;
    }
    if parts.len() == 1 {
        return Plan::Counter { prefix: prefix.to_string(), start: parts[0].1, step: 1, width };
    }
    let numbers: Vec<f64> = parts.iter().map(|p| p.1 as f64).collect();
    let (start, step) = trend(&numbers);
    Plan::Counter { prefix: prefix.to_string(), start: start.round() as i64, step: step.round() as i64, width }
}

fn styled_name(name: &str, case: Case) -> String {
    match case {
        Case::Lower => name.to_string(),
        Case::Upper => name.to_uppercase(),
        Case::Capitalized => {
            let mut chars = name.chars();
            chars.next().map(|c| c.to_uppercase().chain(chars).collect()).unwrap_or_default()
        }
    }
}

/// The text of a cell that holds one, not as the result of a formula
fn cell_text<'a>(node: &'a CellNode, strings: &'a SharedStrings) -> Option<&'a str> {
    match node.kind.as_deref() {
        Some("s") => strings.get(node.value.as_deref()?.trim().parse().ok()?),
        Some("inlineStr") => node.inline.as_deref(),
        Some("str") => node.value.as_deref(),
        _ => None,
    }
}

fn cell_number(node: &CellNode) -> Option<f64> {
    if matches!(node.kind.as_deref(), None | Some("n")) {
        node.value.as_deref()?.trim().parse().ok()
    } else {
        None
    }
}

/// How to fill a row or column whose source cells are `cells`
fn lane_plan(cells: &[Option<&CellNode>], strings: &SharedStrings, is_date: &dyn Fn(u32) -> bool) -> Plan {
    let Some(present) = cells.iter().copied().collect::<Option<Vec<&CellNode>>>() else { return Plan::Copy };
    if present.is_empty() || present.iter().any(|c| c.formula.is_some()) {
        return Plan::Copy;
    }
    let numbers: Option<Vec<f64>> = present.iter().map(|c| cell_number(c)).collect();
    if let Some(numbers) = numbers {
        if numbers.len() > 1 {
            let (start, step) = trend(&numbers);
            return Plan::Number { start, step };
        }
        return if is_date(present[0].style) { Plan::Number { start: numbers[0], step: 1.0 } } else { Plan::Copy };
    }
    let texts: Option<Vec<&str>> = present.iter().map(|c| cell_text(c, strings)).collect();
    texts.map(|t| text_plan(&t)).unwrap_or(Plan::Copy)
}

/// A float without the noise of adding steps (0.1 + 0.2)
fn tidy(value: f64) -> f64 {
    let rounded = (value * 1e10).round() / 1e10;
    if rounded.is_finite() {
        rounded
    } else {
        value
    }
}

impl Workbook {
    /// Fills `target` from `source`, the cells at one of its ends along one axis: see the module
    /// documentation. Without `series` the source cells only repeat (formulas still move their
    /// references). Nothing changes when it is refused.
    pub fn fill(&mut self, sheet: usize, source: &str, target: &str, series: bool) -> Result<(), String> {
        let source = CellRange::parse(source)?;
        let target = CellRange::parse(target)?;
        if target.cell_count() > MAX_CELLS_PER_CALL {
            return Err(format!("El rango {} tiene demasiadas celdas para rellenarlo de una vez.", target));
        }
        let same_cols = source.first.col == target.first.col && source.last.col == target.last.col;
        let same_rows = source.first.row == target.first.row && source.last.row == target.last.row;
        let at_row_end = source.first.row == target.first.row || source.last.row == target.last.row;
        let at_col_end = source.first.col == target.first.col || source.last.col == target.last.col;
        if !(same_cols && at_row_end || same_rows && at_col_end) {
            return Err(format!("Para rellenar, {} debe ser un extremo de {}.", source, target));
        }
        if source == target {
            return Ok(());
        }
        let vertical = same_cols;

        let part = self.worksheet_part(sheet)?;
        let xml = self.package.get_string(&part)?;
        // The cells to fill: what they hold now may be the first cell of a shared formula
        let xml = without_shared_sources(xml, |at| range_contains(&target, at) && !range_contains(&source, at))?;
        if let Some(merge) = SheetLayout::parse(&xml)?.merges.iter().find(|m| overlaps_range(m, &target)) {
            return Err(format!("No se puede rellenar sobre las celdas combinadas {}.", merge));
        }
        let parsed = SheetXml::parse(&xml)?;
        let prefix = parsed.prefix().to_string();
        let nodes: HashMap<CellRef, &CellNode> = parsed.rows.iter().flat_map(|r| &r.cells).map(|c| (c.at, c)).collect();
        let masters: HashMap<u32, (CellRef, &str)> = nodes
            .values()
            .filter_map(|c| {
                let f = c.formula.as_ref()?;
                (f.kind.as_deref() == Some("shared") && !f.text.trim().is_empty()).then_some((f.shared_index?, (c.at, f.text.as_str())))
            })
            .collect();
        let strings_part = self.related_part("sharedStrings");
        let strings_xml = match &strings_part {
            Some(part) => self.package.get_string(part)?,
            None => EMPTY_SST.to_string(),
        };
        let mut strings = SharedStrings::parse(&strings_xml)?;
        let styles = self.styles()?;
        let is_date = |style: u32| styles.is_date_style(style);

        let (lanes, lane_first, length, offset) = if vertical {
            (source.first.col..=source.last.col, source.first.row, source.last.row - source.first.row + 1, target.first.row)
        } else {
            (source.first.row..=source.last.row, source.first.col, source.last.col - source.first.col + 1, target.first.col)
        };
        let target_length = if vertical { target.last.row - target.first.row + 1 } else { target.last.col - target.first.col + 1 };
        let at_of = |lane: u32, line: u32| if vertical { CellRef { row: line, col: lane } } else { CellRef { row: lane, col: line } };

        let mut edits: BTreeMap<CellRef, Option<String>> = BTreeMap::new();
        for lane in lanes {
            let cells: Vec<Option<&CellNode>> = (0..length).map(|i| nodes.get(&at_of(lane, lane_first + i)).copied()).collect();
            let plan = if series { lane_plan(&cells, &strings, &is_date) } else { Plan::Copy };
            for line in offset..offset + target_length {
                if line >= lane_first && line < lane_first + length {
                    continue;
                }
                let at = at_of(lane, line);
                let position = line as i64 - lane_first as i64;
                let from = cells[position.rem_euclid(length as i64) as usize];
                let style = from.and_then(|c| c.style_attr.as_deref());
                let text = match &plan {
                    Plan::Copy | Plan::Number { .. } => None,
                    Plan::Counter { prefix, start, step, width } => {
                        Some(format!("{}{:0width$}", prefix, (start + step * position).max(0), width = *width))
                    }
                    Plan::Names { list, start, step, case } => {
                        Some(styled_name(list[(start + step * position).rem_euclid(list.len() as i64) as usize], *case))
                    }
                };
                let element = if let Plan::Number { start, step } = plan {
                    Some(cell_element(&prefix, at, style, None, None, Some(&format_number(tidy(start + step * position as f64)))))
                } else if let Some(text) = text {
                    Some(cell_element(&prefix, at, style, None, Some("s"), Some(&strings.intern(&text).to_string())))
                } else {
                    match from {
                        None => None,
                        Some(node) => Some(copied_cell(&xml, node, at, &prefix, &masters)?),
                    }
                };
                if element.is_some() || nodes.contains_key(&at) {
                    edits.insert(at, element);
                }
            }
        }
        if edits.is_empty() {
            return Ok(());
        }
        let written = parsed.splice(&edits);
        drop(parsed);
        self.package.put(part, written.into_bytes());
        if strings.changed() {
            let written = strings.write(&strings_xml)?;
            match strings_part {
                Some(part) => self.package.put(part, written.into_bytes()),
                None => {
                    self.add_workbook_part("sharedStrings", "sharedStrings.xml", SHARED_STRINGS_TYPE, written)?;
                }
            }
        }
        self.recalculate_after_edit();
        self.prepare_recalculation()
    }
}

/// `node` as the cell at `at`: formulas written again for it, the rest as they are. Arrays and data
/// tables stand for many cells: only their value is copied.
fn copied_cell(xml: &str, node: &CellNode, at: CellRef, prefix: &str, masters: &HashMap<u32, (CellRef, &str)>) -> Result<String, String> {
    let element = set_attr(&xml[node.range.clone()], "r", &at.to_string());
    let Some(formula) = &node.formula else { return Ok(element) };
    let delta = |to: CellRef, from: CellRef| (to.row as i64 - from.row as i64, to.col as i64 - from.col as i64);
    let (text, (rows, cols)) = match formula.kind.as_deref() {
        Some("array") | Some("dataTable") => {
            let Some(range) = element_ranges(&element, "f", |_| true)?.into_iter().next() else { return Ok(element) };
            return Ok(format!("{}{}", &element[..range.start], &element[range.end..]));
        }
        Some("shared") => match formula.shared_index.and_then(|si| masters.get(&si)) {
            Some((anchor, text)) => (*text, delta(at, *anchor)),
            None => return Ok(element),
        },
        _ => (formula.text.as_str(), delta(at, node.at)),
    };
    Ok(formula_element(prefix, at, node.style_attr.as_deref(), node.phonetic.as_deref(), &shift_formula(text, rows, cols)))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::format::Format;
    use crate::workbook::{CellInput, CellValue};

    fn blank() -> Workbook {
        Workbook::from_bytes(&crate::blank::blank_xlsx().unwrap()).unwrap()
    }

    fn text(s: &str) -> Option<CellInput> {
        Some(CellInput::Text(s.into()))
    }

    fn number(n: f64) -> Option<CellInput> {
        Some(CellInput::Number(n))
    }

    fn column(wb: &Workbook, range: &str) -> Vec<CellValue> {
        wb.range(0, range).unwrap().into_iter().flatten().map(|c| c.value).collect()
    }

    fn numbers(values: &[f64]) -> Vec<CellValue> {
        values.iter().map(|n| CellValue::Number(*n)).collect()
    }

    fn texts(values: &[&str]) -> Vec<CellValue> {
        values.iter().map(|t| CellValue::Text(t.to_string())).collect()
    }

    #[test]
    fn continues_trends() {
        let mut wb = blank();
        wb.set_values(0, "A1", vec![vec![number(1.0)], vec![number(3.0)]]).unwrap();
        wb.fill(0, "A1:A2", "A1:A5", true).unwrap();
        assert_eq!(column(&wb, "A1:A5"), numbers(&[1.0, 3.0, 5.0, 7.0, 9.0]));

        // Upwards and leftwards too
        wb.set_values(0, "C5", vec![vec![number(10.0)], vec![number(20.0)]]).unwrap();
        wb.fill(0, "C5:C6", "C3:C6", true).unwrap();
        assert_eq!(column(&wb, "C3:C4"), numbers(&[-10.0, 0.0]));

        // Not exact: the line that fits best, without float noise
        wb.set_values(0, "E1", vec![vec![number(0.1), number(0.2), number(0.3)]]).unwrap();
        wb.fill(0, "E1:G1", "E1:J1", true).unwrap();
        assert_eq!(column(&wb, "H1:J1"), numbers(&[0.4, 0.5, 0.6]));

        // One plain number repeats; without series so do two
        wb.set_value(0, "L1", CellInput::Number(7.0)).unwrap();
        wb.fill(0, "L1", "L1:L3", true).unwrap();
        assert_eq!(column(&wb, "L2:L3"), numbers(&[7.0, 7.0]));
        wb.set_values(0, "M1", vec![vec![number(1.0)], vec![number(2.0)]]).unwrap();
        wb.fill(0, "M1:M2", "M1:M5", false).unwrap();
        assert_eq!(column(&wb, "M3:M5"), numbers(&[1.0, 2.0, 1.0]));
    }

    #[test]
    fn dates_go_on_day_by_day_keeping_their_format() {
        let mut wb = blank();
        wb.set_value(0, "A1", CellInput::Date(1_705_276_800_000.0)).unwrap();
        wb.fill(0, "A1", "A1:A3", true).unwrap();
        assert_eq!(column(&wb, "A3:A3"), vec![CellValue::Date(1_705_276_800_000.0 + 2.0 * 86_400_000.0)]);
    }

    #[test]
    fn counts_on_in_text_and_follows_lists() {
        let mut wb = blank();
        wb.set_values(0, "A1", vec![vec![text("Lote 9"), text("L08"), text("Enero"), text("LUNES"), text("Q1"), text("hola"), text("mar")]]).unwrap();
        wb.set_value(0, "E2", CellInput::Text("Q3".into())).unwrap();
        wb.fill(0, "A1", "A1:A3", true).unwrap();
        assert_eq!(column(&wb, "A2:A3"), texts(&["Lote 10", "Lote 11"]));
        wb.fill(0, "B1", "B1:B3", true).unwrap();
        assert_eq!(column(&wb, "B2:B3"), texts(&["L09", "L10"]));
        wb.fill(0, "C1", "C1:C13", true).unwrap();
        assert_eq!(column(&wb, "C2:C2"), texts(&["Febrero"]));
        assert_eq!(column(&wb, "C12:C13"), texts(&["Diciembre", "Enero"]));
        wb.fill(0, "D1", "D1:D3", true).unwrap();
        assert_eq!(column(&wb, "D2:D3"), texts(&["MARTES", "MIÉRCOLES"]));
        wb.fill(0, "E1:E2", "E1:E4", true).unwrap();
        assert_eq!(column(&wb, "E3:E4"), texts(&["Q5", "Q7"]));
        wb.fill(0, "F1", "F1:F2", true).unwrap();
        assert_eq!(column(&wb, "F2:F2"), texts(&["hola"]));
        wb.fill(0, "G1", "G1:G2", true).unwrap();
        assert_eq!(column(&wb, "G2:G2"), texts(&["abr"]));
    }

    #[test]
    fn formulas_move_their_references_and_formats_repeat() {
        let mut wb = blank();
        wb.set_values(0, "A1", vec![vec![number(1.0), Some(CellInput::Formula("A1*$A$1+B$1".into()))], vec![number(2.0)], vec![number(3.0)]]).unwrap();
        wb.set_format(0, "B1", &Format { bold: Some(true), ..Default::default() }).unwrap();
        wb.fill(0, "B1", "B1:B3", true).unwrap();
        assert_eq!(wb.cell(0, "B3").unwrap().formula.as_deref(), Some("A3*$A$1+B$1"));
        let style = |at: &str| SheetXml::parse(&wb.package.get_string("xl/worksheets/sheet1.xml").unwrap()).unwrap().cell(CellRef::parse(at).unwrap()).unwrap().style;
        assert_eq!(style("B3"), style("B1"));
        assert_ne!(style("B3"), 0);

        // Empty source cells empty the target, in one undo step
        let mut wb = blank();
        wb.set_values(0, "A1", vec![vec![text("a")], vec![], vec![text("borrar")]]).unwrap();
        wb.checkpoint(None);
        wb.fill(0, "A2", "A2:A3", true).unwrap();
        assert_eq!(column(&wb, "A3:A3"), vec![CellValue::Empty]);
        wb.undo().unwrap();
        assert_eq!(column(&wb, "A3:A3"), texts(&["borrar"]));
    }

    #[test]
    fn shared_formulas_are_written_out_for_the_new_cells() {
        let mut wb = blank();
        let xml = wb.package.get_string("xl/worksheets/sheet1.xml").unwrap();
        let start = xml.find("<sheetData").unwrap();
        let end = xml.find("</worksheet>").unwrap();
        let data = r#"<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row><row r="3"><c r="A3"><v>3</v></c></row></sheetData>"#;
        wb.package.put("xl/worksheets/sheet1.xml".into(), format!("{}{}{}", &xml[..start], data, &xml[end..]).into_bytes());
        wb.fill(0, "B2", "B2:B3", true).unwrap();
        assert_eq!(wb.cell(0, "B3").unwrap().formula.as_deref(), Some("A3*2"));
        assert_eq!(wb.cell(0, "B3").unwrap().value, CellValue::Number(6.0));
        // Over the first cell of a shared formula: B2, which read it from there, keeps its own
        let mut wb2 = blank();
        let xml = wb2.package.get_string("xl/worksheets/sheet1.xml").unwrap();
        let start = xml.find("<sheetData").unwrap();
        let end = xml.find("</worksheet>").unwrap();
        let data = r#"<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c><c r="C1"><v>9</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row></sheetData>"#;
        wb2.package.put("xl/worksheets/sheet1.xml".into(), format!("{}{}{}", &xml[..start], data, &xml[end..]).into_bytes());
        wb2.fill(0, "C1", "B1:C1", false).unwrap();
        assert_eq!(wb2.cell(0, "B1").unwrap().value, CellValue::Number(9.0));
        assert_eq!(wb2.cell(0, "B2").unwrap().formula.as_deref(), Some("A2*2"));
        assert_eq!(wb2.cell(0, "B2").unwrap().value, CellValue::Number(4.0));
    }

    #[test]
    fn refuses_what_does_not_fit() {
        let mut wb = blank();
        wb.set_value(0, "A1", CellInput::Number(1.0)).unwrap();
        assert!(wb.fill(0, "A1", "B2", true).is_err());
        assert!(wb.fill(0, "A1", "A1:B2", true).is_err(), "two axes at once");
        wb.merge_cells(0, "A3:B3").unwrap();
        assert!(wb.fill(0, "A1", "A1:A4", true).unwrap_err().contains("A3:B3"));
    }
}
