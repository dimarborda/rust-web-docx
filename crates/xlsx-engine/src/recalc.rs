//! Recalculation of a workbook after an edit: every formula is evaluated in dependency order
//! (each one after the formulas it reads) and the cells whose result changed get the new value
//! in `<v>`, with the rest of the sheet left as written.
//!
//! A formula that uses something `formula` or `calc` do not support keeps the value Excel
//! cached, and so do circular references. Excel still recalculates everything when it opens
//! the file (`fullCalcOnLoad`), so those cells are right again there.

use std::collections::{BTreeMap, HashMap};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use ooxml_core::xml::{element_ranges, escape_xml, get_attr_value, remove_attr, set_attr, tag_is};

use crate::calc::{self, Context, Val};
use crate::cell_ref::CellRef;
use crate::dates::{iso_to_unix_ms, unix_ms_to_serial};
use crate::formula::{self, Expr};
use crate::shared_strings::SharedStrings;
use crate::sheet::{CellNode, SheetXml};
use crate::text::decode;

/// Most references from formulas to formula cells followed in one recalculation; past it the
/// workbook is left for Excel to recalculate
const MAX_EDGES: usize = 5_000_000;

/// A defined name of the workbook (`<definedName>`)
#[derive(Debug, Clone)]
pub struct DefinedName {
    pub name: String,
    /// Sheet the name belongs to (`localSheetId`), or global
    pub local_sheet: Option<usize>,
    pub text: String,
}

/// The defined names of a workbook part
pub fn defined_names(workbook_xml: &str) -> Vec<DefinedName> {
    let mut names = Vec::new();
    let mut reader = Reader::from_str(workbook_xml);
    let mut current: Option<DefinedName> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) if tag_is(e.name().as_ref(), "definedName") => {
                current = Some(DefinedName {
                    name: get_attr_value(&e, "name").unwrap_or_default(),
                    local_sheet: get_attr_value(&e, "localSheetId").and_then(|v| v.trim().parse().ok()),
                    text: String::new(),
                });
            }
            Ok(Event::Text(t)) => {
                if let Some(name) = current.as_mut() {
                    name.text.push_str(&decode(&t.into_inner()));
                }
            }
            Ok(Event::GeneralRef(r)) => {
                if let Some(name) = current.as_mut() {
                    name.text.push_str(&decode(&format!("&{};", r.into_inner())));
                }
            }
            Ok(Event::End(e)) if tag_is(e.name().as_ref(), "definedName") => {
                if let Some(name) = current.take() {
                    names.push(name);
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    names
}

/// The value of a cell as formulas see it (dates stay serial numbers)
pub fn node_value(node: &CellNode, strings: &SharedStrings, date1904: bool) -> Val {
    let raw = node.value.as_deref().unwrap_or("");
    match node.kind.as_deref() {
        Some("s") => raw.trim().parse::<usize>().ok().and_then(|i| strings.get(i)).map(|s| Val::Text(s.to_string())).unwrap_or(Val::Empty),
        Some("str") => Val::Text(raw.to_string()),
        Some("inlineStr") => Val::Text(node.inline.clone().unwrap_or_else(|| raw.to_string())),
        Some("b") => Val::Bool(matches!(raw.trim(), "1" | "true")),
        Some("e") => Val::Error(raw.trim().to_string()),
        Some("d") => iso_to_unix_ms(raw)
            .and_then(|ms| unix_ms_to_serial(ms, date1904).ok())
            .map(Val::Number)
            .unwrap_or_else(|| Val::Text(raw.to_string())),
        _ if node.value.is_none() => Val::Empty,
        _ => raw.trim().parse::<f64>().map(Val::Number).unwrap_or_else(|_| Val::Text(raw.to_string())),
    }
}

/// The formula of every formula cell of a sheet, with shared formulas expanded: `Err` when it
/// is not supported (array and data table formulas, or what the parser does not handle)
pub fn sheet_formulas(parsed: &SheetXml) -> Vec<(CellRef, Result<Expr, String>)> {
    let mut masters: HashMap<u32, (CellRef, Result<Expr, String>)> = HashMap::new();
    for cell in parsed.rows.iter().flat_map(|r| &r.cells) {
        if let Some(f) = &cell.formula {
            if f.kind.as_deref() == Some("shared") && !f.text.trim().is_empty() {
                if let Some(si) = f.shared_index {
                    masters.insert(si, (cell.at, formula::parse(&f.text).map_err(|e| e.message().to_string())));
                }
            }
        }
    }
    let mut out = Vec::new();
    for cell in parsed.rows.iter().flat_map(|r| &r.cells) {
        let Some(f) = &cell.formula else { continue };
        let expr = match f.kind.as_deref() {
            Some("array") => Err("fórmulas matriciales".to_string()),
            Some("dataTable") => Err("tablas de datos".to_string()),
            Some("shared") if f.text.trim().is_empty() => match f.shared_index.and_then(|si| masters.get(&si)) {
                Some((anchor, Ok(master))) => {
                    Ok(master.shifted(cell.at.row as i64 - anchor.row as i64, cell.at.col as i64 - anchor.col as i64))
                }
                Some((_, Err(e))) => Err(e.clone()),
                None => Err("una fórmula compartida sin celda principal".to_string()),
            },
            _ => formula::parse(&f.text).map_err(|e| e.message().to_string()),
        };
        out.push((cell.at, expr));
    }
    out
}

struct FormulaCell {
    sheet: usize,
    at: CellRef,
    expr: Option<Expr>,
    cached: Val,
}

#[derive(Default)]
struct SheetModel {
    values: HashMap<CellRef, Val>,
    /// Formula cells by (column, row), for range queries column by column
    formulas: BTreeMap<(u32, u32), usize>,
    max_row: u32,
    max_col: u32,
}

struct Model<'n> {
    sheets: Vec<SheetModel>,
    sheet_names: Vec<String>,
    names: Vec<(String, Option<usize>, Option<Expr>)>,
    cells: Vec<FormulaCell>,
    results: Vec<Option<Val>>,
    date1904: bool,
    now: f64,
    _defined: std::marker::PhantomData<&'n ()>,
}

impl Context for Model<'_> {
    fn value(&self, sheet: usize, at: CellRef) -> Val {
        let Some(model) = self.sheets.get(sheet) else { return Val::Error("#REF!".into()) };
        if let Some(&i) = model.formulas.get(&(at.col, at.row)) {
            return self.results[i].clone().unwrap_or_else(|| self.cells[i].cached.clone());
        }
        model.values.get(&at).cloned().unwrap_or(Val::Empty)
    }

    fn sheet_index(&self, name: &str) -> Option<usize> {
        self.sheet_names.iter().position(|n| n.to_lowercase() == name.to_lowercase())
    }

    fn name(&self, name: &str, sheet: usize) -> Option<(&Expr, usize)> {
        let upper = name.to_uppercase();
        let found = self
            .names
            .iter()
            .find(|(n, local, _)| *n == upper && *local == Some(sheet))
            .or_else(|| self.names.iter().find(|(n, local, _)| *n == upper && local.is_none()))?;
        found.2.as_ref().map(|e| (e, found.1.unwrap_or(sheet)))
    }

    fn bounds(&self, sheet: usize) -> (u32, u32) {
        self.sheets.get(sheet).map(|s| (s.max_row, s.max_col)).unwrap_or((0, 0))
    }

    fn date1904(&self) -> bool {
        self.date1904
    }

    fn now(&self) -> f64 {
        self.now
    }
}

impl Model<'_> {
    /// Formula cells a formula reads, directly or through defined names
    fn precedents(&self, sheet: usize, expr: &Expr, out: &mut Vec<usize>, depth: usize) {
        expr.visit_refs(&mut |e| match e {
            Expr::Ref(area) => {
                let target = match &area.sheet {
                    Some(name) => match self.sheet_index(name) {
                        Some(i) => i,
                        None => return,
                    },
                    None => sheet,
                };
                let Some(model) = self.sheets.get(target) else { return };
                let (r1, r2) = (area.first.row.min(area.last.row), area.first.row.max(area.last.row).min(model.max_row));
                let (c1, c2) = (area.first.col.min(area.last.col), area.first.col.max(area.last.col).min(model.max_col));
                for col in c1..=c2 {
                    out.extend(model.formulas.range((col, r1)..=(col, r2.max(r1))).map(|(_, i)| *i));
                }
            }
            Expr::Name(name) if depth < 16 => {
                if let Some((definition, name_sheet)) = self.name(name, sheet) {
                    self.precedents(name_sheet, definition, out, depth + 1);
                }
            }
            _ => {}
        });
    }
}

/// What a recalculation changed
pub struct Outcome {
    /// New text of each sheet whose cached values changed (by position in the input)
    pub sheets: Vec<Option<String>>,
    /// Cells whose value was rewritten
    pub updated: usize,
}

/// Recalculates the formulas of the worksheets in `xmls` (one entry per sheet in tab order,
/// `None` for sheets without cells); `now` is the local date and time as a serial number
pub fn recalculate(
    sheet_names: &[String],
    xmls: &[Option<String>],
    strings: &SharedStrings,
    names: &[DefinedName],
    date1904: bool,
    now: f64,
) -> Result<Outcome, String> {
    let mut parsed = Vec::with_capacity(xmls.len());
    for xml in xmls {
        parsed.push(match xml {
            Some(xml) => Some(SheetXml::parse(xml)?),
            None => None,
        });
    }

    let mut model = Model {
        sheets: Vec::new(),
        sheet_names: sheet_names.to_vec(),
        names: names
            .iter()
            .filter(|n| !n.name.starts_with("_xlnm."))
            .map(|n| (n.name.to_uppercase(), n.local_sheet, formula::parse(&n.text).ok()))
            .collect(),
        cells: Vec::new(),
        results: Vec::new(),
        date1904,
        now,
        _defined: std::marker::PhantomData,
    };
    for (index, sheet) in parsed.iter().enumerate() {
        let mut sheet_model = SheetModel::default();
        if let Some(sheet) = sheet {
            for cell in sheet.rows.iter().flat_map(|r| &r.cells) {
                sheet_model.max_row = sheet_model.max_row.max(cell.at.row);
                sheet_model.max_col = sheet_model.max_col.max(cell.at.col);
                let value = node_value(cell, strings, date1904);
                if cell.formula.is_none() && value != Val::Empty {
                    sheet_model.values.insert(cell.at, value);
                }
            }
            for (at, expr) in sheet_formulas(sheet) {
                let cached = sheet.cell(at).map(|n| node_value(n, strings, date1904)).unwrap_or(Val::Empty);
                sheet_model.formulas.insert((at.col, at.row), model.cells.len());
                model.cells.push(FormulaCell { sheet: index, at, expr: expr.ok(), cached });
            }
        }
        model.sheets.push(sheet_model);
    }
    model.results = vec![None; model.cells.len()];
    if model.cells.is_empty() {
        return Ok(Outcome { sheets: vec![None; xmls.len()], updated: 0 });
    }

    // Order: depth-first over precedents, each formula after the ones it reads
    let mut precedents: Vec<Vec<usize>> = Vec::with_capacity(model.cells.len());
    let mut edges = 0;
    for cell in &model.cells {
        let mut list = Vec::new();
        if let Some(expr) = &cell.expr {
            model.precedents(cell.sheet, expr, &mut list, 0);
        }
        edges += list.len();
        if edges > MAX_EDGES {
            return Ok(Outcome { sheets: vec![None; xmls.len()], updated: 0 });
        }
        precedents.push(list);
    }
    #[derive(Clone, Copy, PartialEq)]
    enum State {
        New,
        Open,
        Done,
    }
    let mut state = vec![State::New; model.cells.len()];
    // Formulas in a cycle keep their cached value
    let mut cyclic = vec![false; model.cells.len()];
    let mut order = Vec::with_capacity(model.cells.len());
    for start in 0..model.cells.len() {
        if state[start] != State::New {
            continue;
        }
        let mut stack = vec![(start, false)];
        while let Some((i, expanded)) = stack.pop() {
            if expanded {
                state[i] = State::Done;
                order.push(i);
                continue;
            }
            if state[i] != State::New {
                continue;
            }
            state[i] = State::Open;
            stack.push((i, true));
            for &p in &precedents[i] {
                match state[p] {
                    State::New => stack.push((p, false)),
                    State::Open => {
                        cyclic[p] = true;
                        cyclic[i] = true;
                    }
                    State::Done => {}
                }
            }
        }
    }

    for i in order {
        if cyclic[i] {
            continue;
        }
        let cell = &model.cells[i];
        let Some(expr) = &cell.expr else { continue };
        if let Ok(value) = calc::evaluate(&model, cell.sheet, cell.at, expr) {
            model.results[i] = Some(value);
        }
    }

    // Write the values that changed
    let mut edits: Vec<BTreeMap<CellRef, Option<String>>> = vec![BTreeMap::new(); xmls.len()];
    let mut updated = 0;
    for (i, cell) in model.cells.iter().enumerate() {
        let Some(value) = &model.results[i] else { continue };
        if same_value(&cell.cached, value) {
            continue;
        }
        let sheet = parsed[cell.sheet].as_ref().expect("formula cells come from parsed sheets");
        let Some(node) = sheet.cell(cell.at) else { continue };
        let element = &sheet.xml[node.range.clone()];
        edits[cell.sheet].insert(cell.at, Some(with_cached_value(element, sheet.prefix(), value)?));
        updated += 1;
    }
    let sheets = parsed
        .iter()
        .zip(&edits)
        .map(|(sheet, edits)| match sheet {
            Some(sheet) if !edits.is_empty() => Some(sheet.splice(edits)),
            _ => None,
        })
        .collect();
    Ok(Outcome { sheets, updated })
}

/// Whether a new result equals the cached one (numbers compared at Excel's 15 digits)
fn same_value(cached: &Val, new: &Val) -> bool {
    match (cached, new) {
        (Val::Number(a), Val::Number(b)) => a == b || format!("{:.14e}", a) == format!("{:.14e}", b),
        (a, b) => a == b,
    }
}

/// A formula cell element with `value` as its cached result: the type attribute and `<v>`
/// change, the formula and everything else stay as written
fn with_cached_value(element: &str, prefix: &str, value: &Val) -> Result<String, String> {
    let (kind, text) = match value {
        Val::Number(n) => (None, crate::workbook::format_number(*n)),
        Val::Text(t) => (Some("str"), t.clone()),
        Val::Bool(b) => (Some("b"), if *b { "1" } else { "0" }.to_string()),
        Val::Error(e) => (Some("e"), e.clone()),
        Val::Empty => (None, "0".to_string()),
    };
    let mut out = match kind {
        Some(t) => set_attr(element, "t", t),
        None => remove_attr(element, "t"),
    };
    let v = format!("<{p}v>{}</{p}v>", escape_xml(&text), p = prefix);
    if let Some(range) = element_ranges(&out, "v", |_| true)?.into_iter().next() {
        out.replace_range(range, &v);
    } else if let Some(range) = element_ranges(&out, "f", |_| true)?.into_iter().next() {
        out.insert_str(range.end, &v);
    } else {
        return Err("Una celda con fórmula sin elemento <f>.".to_string());
    }
    Ok(out)
}

/// The local date and time as a serial number
pub fn now_serial(date1904: bool) -> f64 {
    unix_ms_to_serial(local_now_ms(), date1904).unwrap_or(0.0)
}

#[cfg(target_arch = "wasm32")]
mod clock {
    use wasm_bindgen::prelude::*;

    #[wasm_bindgen]
    extern "C" {
        #[wasm_bindgen(js_name = Date)]
        pub type JsDate;
        #[wasm_bindgen(constructor, js_class = "Date")]
        pub fn new() -> JsDate;
        #[wasm_bindgen(method, js_name = getTime)]
        pub fn get_time(this: &JsDate) -> f64;
        #[wasm_bindgen(method, js_name = getTimezoneOffset)]
        pub fn get_timezone_offset(this: &JsDate) -> f64;
    }
}

#[cfg(target_arch = "wasm32")]
fn local_now_ms() -> f64 {
    let now = clock::JsDate::new();
    now.get_time() - now.get_timezone_offset() * 60_000.0
}

#[cfg(not(target_arch = "wasm32"))]
fn local_now_ms() -> f64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|d| d.as_millis() as f64)
        .unwrap_or(0.0)
}

#[cfg(test)]
mod tests {
    use super::*;

    const NS: &str = r#"xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#;

    fn sheet(cells: &str) -> String {
        format!(r#"<worksheet {}><sheetData>{}</sheetData></worksheet>"#, NS, cells)
    }

    fn run(xmls: Vec<Option<String>>, names: &[DefinedName]) -> Outcome {
        let sheet_names: Vec<String> = (1..=xmls.len()).map(|i| format!("Hoja{}", i)).collect();
        recalculate(&sheet_names, &xmls, &SharedStrings::default(), names, false, 45_000.5).unwrap()
    }

    #[test]
    fn updates_only_changed_results_in_dependency_order() {
        // C1 = B1*2 reads B1 = A1+1, written after it; D1 is already right
        let xml = sheet(
            r#"<row r="1"><c r="A1"><v>5</v></c><c r="C1" s="3"><f>B1*2</f><v>0</v></c><c r="B1"><f>A1+1</f><v>0</v></c><c r="D1"><f>A1</f><v>5</v></c></row>"#,
        );
        let out = run(vec![Some(xml.clone())], &[]);
        assert_eq!(out.updated, 2);
        let new = out.sheets[0].as_ref().unwrap();
        assert!(new.contains(r#"<c r="C1" s="3"><f>B1*2</f><v>12</v></c>"#), "{}", new);
        assert!(new.contains(r#"<c r="B1"><f>A1+1</f><v>6</v></c>"#));
        assert!(new.contains(r#"<c r="D1"><f>A1</f><v>5</v></c>"#));
    }

    #[test]
    fn writes_text_booleans_and_errors_with_their_type() {
        let xml = sheet(
            r#"<row r="1"><c r="A1"><v>0</v></c><c r="B1" t="str"><f>"n="&amp;A1</f><v>x</v></c><c r="C1"><f>A1=0</f><v>0</v></c><c r="D1"><f>1/A1</f></c><c r="E1" t="str"><f>A1+1</f><v>x</v></c></row>"#,
        );
        let new = run(vec![Some(xml)], &[]).sheets.remove(0).unwrap();
        assert!(new.contains(r#"<c r="B1" t="str"><f>"n="&amp;A1</f><v>n=0</v></c>"#), "{}", new);
        assert!(new.contains(r#"<c r="C1" t="b"><f>A1=0</f><v>1</v></c>"#));
        assert!(new.contains(r#"<c r="D1" t="e"><f>1/A1</f><v>#DIV/0!</v></c>"#));
        assert!(new.contains(r#"<c r="E1"><f>A1+1</f><v>1</v></c>"#));
    }

    #[test]
    fn expands_shared_formulas_and_reads_other_sheets_and_names() {
        let data = sheet(
            r#"<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*Tasa</f><v>0</v></c></row><row r="2"><c r="A2"><v>3</v></c><c r="B2"><f t="shared" si="0"/><v>0</v></c></row>"#,
        );
        let summary = sheet(r#"<row r="1"><c r="A1"><f>SUM(Hoja1!B:B)</f><v>0</v></c></row>"#);
        let names = [DefinedName { name: "Tasa".into(), local_sheet: None, text: "Hoja2!$B$5".into() }];
        let summary = summary.replace("</row>", r#"</row><row r="5"><c r="B5"><v>10</v></c></row>"#);
        let out = run(vec![Some(data), Some(summary)], &names);
        assert!(out.sheets[0].as_ref().unwrap().contains(r#"<c r="B2"><f t="shared" si="0"/><v>30</v></c>"#));
        assert!(out.sheets[1].as_ref().unwrap().contains(r#"<c r="A1"><f>SUM(Hoja1!B:B)</f><v>50</v></c>"#));
    }

    #[test]
    fn keeps_cached_values_of_unsupported_and_circular_formulas() {
        let xml = sheet(
            r#"<row r="1"><c r="A1"><f>OFFSET(Z1,0,1)</f><v>7</v></c><c r="B1"><f>A1+1</f><v>0</v></c><c r="C1"><f>D1</f><v>1</v></c><c r="D1"><f>C1</f><v>1</v></c><c r="E1"><f t="array" ref="E1">SUM(B1)</f><v>9</v></c></row>"#,
        );
        let out = run(vec![Some(xml)], &[]);
        assert_eq!(out.updated, 1);
        let new = out.sheets[0].as_ref().unwrap();
        assert!(new.contains(r#"<c r="A1"><f>OFFSET(Z1,0,1)</f><v>7</v></c>"#));
        assert!(new.contains(r#"<c r="B1"><f>A1+1</f><v>8</v></c>"#));
        assert!(new.contains(r#"<c r="E1"><f t="array" ref="E1">SUM(B1)</f><v>9</v></c>"#));
    }

    #[test]
    fn nothing_changes_when_cached_values_are_right() {
        let xml = sheet(r#"<row r="1"><c r="A1"><v>0.1</v></c><c r="B1"><f>A1+0.2</f><v>0.3</v></c></row>"#);
        let out = run(vec![Some(xml), None], &[]);
        assert_eq!(out.updated, 0);
        assert!(out.sheets.iter().all(Option::is_none));
    }

    #[test]
    fn reads_defined_names() {
        let xml = r#"<workbook><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Hoja1!$A$1:$B$2</definedName><definedName name="Tasa">Hoja2!$B$5</definedName><definedName name="T&amp;C" localSheetId="1">"a&amp;b"</definedName></definedNames></workbook>"#;
        let names = defined_names(xml);
        assert_eq!(names.len(), 3);
        assert_eq!((names[1].name.as_str(), names[1].local_sheet, names[1].text.as_str()), ("Tasa", None, "Hoja2!$B$5"));
        assert_eq!((names[2].name.as_str(), names[2].local_sheet, names[2].text.as_str()), ("T&C", Some(1), "\"a&b\""));
    }
}
