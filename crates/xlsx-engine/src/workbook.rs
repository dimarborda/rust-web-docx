//! A workbook: the package of an .xlsx read with `ooxml-core`, its sheets found through their
//! relationships, and cell edits written back losslessly: only the parts an edit touches are
//! rewritten, and in them only the elements it changes.

use std::collections::{BTreeMap, HashMap};

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use ooxml_core::rels::{parse_relationships, rels_part_for, resolve_target, type_is, with_relationship, without_relationship, Relationship};
use ooxml_core::xml::{element_ranges, escape_xml, get_attr_value, prefix_of, root_element, set_attr, tag_is};
use ooxml_core::{content_types, Package, Undone, ZipLimits};

use crate::cell_ref::{CellRange, CellRef};
use crate::dates::{has_time, iso_to_unix_ms, serial_to_unix_ms, unix_ms_to_serial};
use crate::formula::{self, ParseError};
use crate::recalc::{self, defined_names, sheet_formulas};
use crate::shared_strings::{SharedStrings, EMPTY_SST};
use crate::sheet::{CellNode, SheetXml};
use crate::styles::{date_style, Styles};
use crate::text::fill_xml;
use crate::numfmt::Locale;
use crate::theme::Theme;
use crate::view::{self, CellView, SheetLayout, StyleSheetView};

/// Main part of the workbook when the package does not say otherwise
pub const WORKBOOK_PART: &str = "xl/workbook.xml";

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SHARED_STRINGS_TYPE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const STYLES_TYPE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

/// Most cells one read or write may cover, so a typo like `A1:XFD1048576` fails instead of
/// exhausting memory
pub const MAX_CELLS_PER_CALL: u64 = 5_000_000;

/// What a sheet tab holds
#[derive(Debug, Clone, PartialEq)]
pub enum SheetKind {
    Worksheet,
    /// A sheet that is a single chart, with no cells
    Chartsheet,
    /// Dialog or macro sheets and anything else
    Other,
}

#[derive(Debug, Clone)]
pub struct SheetInfo {
    pub name: String,
    /// Part name, resolved through the workbook relationships
    pub part: String,
    pub kind: SheetKind,
    /// `hidden` or `veryHidden`, when the tab is not visible
    pub state: Option<String>,
}

/// The value of a cell as Excel shows it (for formulas, the value cached when last calculated)
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Empty,
    Number(f64),
    Text(String),
    Bool(bool),
    /// Error code such as `#DIV/0!`
    Error(String),
    /// Milliseconds since 1970-01-01 UTC (see `dates`)
    Date(f64),
}

#[derive(Debug, Clone, PartialEq)]
pub struct CellData {
    pub at: CellRef,
    pub value: CellValue,
    /// Formula without `=`; empty for cells that reuse a shared formula of another cell
    pub formula: Option<String>,
}

/// A value to write in a cell
#[derive(Debug, Clone, PartialEq)]
pub enum CellInput {
    Text(String),
    Number(f64),
    Bool(bool),
    /// Milliseconds since 1970-01-01 UTC
    Date(f64),
    /// A formula (with or without the leading `=`), in Excel's English syntax
    Formula(String),
    /// Removes the value (and formula), keeping the cell format
    Clear,
}

pub struct Workbook {
    package: Package,
    workbook_part: String,
    /// The recalculation set-up has been done (see `prepare_recalculation`)
    recalc_prepared: bool,
}

impl Workbook {
    /// Opens an .xlsx from the bytes of its ZIP
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, String> {
        Self::from_bytes_with_limits(bytes, &ZipLimits::default())
    }

    pub fn from_bytes_with_limits(bytes: &[u8], limits: &ZipLimits) -> Result<Self, String> {
        let package = Package::from_bytes(bytes, limits, "XLSX")?;
        let workbook_part = package
            .get_string("_rels/.rels")
            .ok()
            .and_then(|xml| parse_relationships(&xml).into_iter().find(|r| type_is(&r.rel_type, "officeDocument") && !r.external))
            .map(|r| resolve_target("", &r.target))
            .filter(|part| package.contains(part))
            .unwrap_or_else(|| WORKBOOK_PART.to_string());
        if !package.contains(&workbook_part) {
            return Err(format!("El archivo no es un libro de Excel: falta '{}'.", WORKBOOK_PART));
        }
        Ok(Workbook { package, workbook_part, recalc_prepared: false })
    }

    fn workbook_xml(&self) -> Result<String, String> {
        self.package.get_string(&self.workbook_part)
    }

    fn workbook_rels(&self) -> Vec<Relationship> {
        self.package
            .get_string(&rels_part_for(&self.workbook_part))
            .map(|xml| parse_relationships(&xml))
            .unwrap_or_default()
    }

    /// Part related to the workbook by a relationship of type `name` (`sharedStrings`, `styles`)
    fn related_part(&self, name: &str) -> Option<String> {
        self.workbook_rels()
            .into_iter()
            .find(|r| type_is(&r.rel_type, name) && !r.external)
            .map(|r| resolve_target(&self.workbook_part, &r.target))
            .filter(|part| self.package.contains(part))
    }

    /// The sheets in tab order, each with the part its relationship points to
    pub fn sheets(&self) -> Result<Vec<SheetInfo>, String> {
        let xml = self.workbook_xml()?;
        let rels: HashMap<String, Relationship> = self.workbook_rels().into_iter().map(|r| (r.id.clone(), r)).collect();
        let mut reader = Reader::from_str(&xml);
        let mut sheets = Vec::new();
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) if tag_is(e.name().as_ref(), "sheet") => {
                    let name = get_attr_value(&e, "name").unwrap_or_default();
                    let rel = get_attr_value(&e, "id").and_then(|id| rels.get(&id));
                    let (part, kind) = match rel {
                        Some(rel) => (
                            resolve_target(&self.workbook_part, &rel.target),
                            if type_is(&rel.rel_type, "worksheet") {
                                SheetKind::Worksheet
                            } else if type_is(&rel.rel_type, "chartsheet") {
                                SheetKind::Chartsheet
                            } else {
                                SheetKind::Other
                            },
                        ),
                        None => (String::new(), SheetKind::Other),
                    };
                    sheets.push(SheetInfo { name, part, kind, state: get_attr_value(&e, "state") });
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(format!("'{}' no es XML válido: {}", self.workbook_part, e)),
                _ => {}
            }
        }
        Ok(sheets)
    }

    /// Names of the sheets, in the order Excel shows their tabs
    pub fn sheet_names(&self) -> Result<Vec<String>, String> {
        Ok(self.sheets()?.into_iter().map(|s| s.name).collect())
    }

    /// Index of the sheet called `name` (Excel compares sheet names ignoring case)
    pub fn sheet_index(&self, name: &str) -> Result<usize, String> {
        let sheets = self.sheets()?;
        sheets
            .iter()
            .position(|s| s.name == name)
            .or_else(|| sheets.iter().position(|s| s.name.to_lowercase() == name.to_lowercase()))
            .ok_or_else(|| format!("El libro no tiene ninguna hoja llamada '{}'.", name))
    }

    /// The worksheet at `index` (its part name), refusing chart sheets and missing parts
    fn worksheet_part(&self, index: usize) -> Result<String, String> {
        let sheets = self.sheets()?;
        let sheet = sheets
            .get(index)
            .ok_or_else(|| format!("El libro tiene {} hoja(s); no existe la hoja {}.", sheets.len(), index))?;
        if sheet.kind != SheetKind::Worksheet {
            return Err(format!("La hoja '{}' no tiene celdas (es una hoja de gráfico u otro tipo).", sheet.name));
        }
        if !self.package.contains(&sheet.part) {
            return Err(format!("Falta la parte '{}' de la hoja '{}'.", sheet.part, sheet.name));
        }
        Ok(sheet.part.clone())
    }

    /// Whether dates count from 1904 (`<workbookPr date1904="1"/>`) instead of 1900
    pub fn date1904(&self) -> bool {
        let Ok(xml) = self.workbook_xml() else { return false };
        let mut reader = Reader::from_str(&xml);
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) if tag_is(e.name().as_ref(), "workbookPr") => {
                    return get_attr_value(&e, "date1904").is_some_and(|v| v == "1" || v.eq_ignore_ascii_case("true"));
                }
                Ok(Event::Start(e)) if tag_is(e.name().as_ref(), "sheets") => return false,
                Ok(Event::Eof) | Err(_) => return false,
                _ => {}
            }
        }
    }

    fn shared_strings(&self) -> Result<SharedStrings, String> {
        match self.related_part("sharedStrings") {
            Some(part) => SharedStrings::parse(&self.package.get_string(&part)?),
            None => Ok(SharedStrings::default()),
        }
    }

    fn styles(&self) -> Result<Styles, String> {
        match self.related_part("styles") {
            Some(part) => Styles::parse(&self.package.get_string(&part)?),
            None => Ok(Styles::default()),
        }
    }

    /// The workbook theme (Office's default when it has none)
    fn theme(&self) -> Theme {
        self.related_part("theme")
            .and_then(|part| self.package.get_string(&part).ok())
            .map(|xml| Theme::parse(&xml))
            .unwrap_or_default()
    }

    /// The cell formats resolved for drawing
    pub fn style_view(&self) -> Result<StyleSheetView, String> {
        match self.related_part("styles") {
            Some(part) => StyleSheetView::parse(&self.package.get_string(&part)?, &self.theme()),
            None => StyleSheetView::parse("<styleSheet/>", &self.theme()),
        }
    }

    /// Column widths, row heights, merged cells and frozen panes of the sheet at `sheet`
    pub fn sheet_layout(&self, sheet: usize) -> Result<SheetLayout, String> {
        SheetLayout::parse(&self.package.get_string(&self.worksheet_part(sheet)?)?)
    }

    /// The cells of the sheet at `sheet` that show something, with the text Excel shows for
    /// them in the language `locale`
    pub fn cell_views(&self, sheet: usize, locale: &Locale) -> Result<Vec<CellView>, String> {
        let xml = self.package.get_string(&self.worksheet_part(sheet)?)?;
        view::cell_views(&xml, &self.shared_strings()?, &self.style_view()?, locale, self.date1904())
    }

    /// The cell at `reference` of the sheet at `sheet`
    pub fn cell(&self, sheet: usize, reference: &str) -> Result<CellData, String> {
        let at = CellRef::parse(reference)?;
        let range = CellRange { first: at, last: at };
        Ok(self.range(sheet, &range.to_string())?.remove(0).remove(0))
    }

    /// The cells of a range (`A1:C3`), by row
    pub fn range(&self, sheet: usize, range: &str) -> Result<Vec<Vec<CellData>>, String> {
        let range = CellRange::parse(range)?;
        if range.cell_count() > MAX_CELLS_PER_CALL {
            return Err(format!("El rango {} tiene demasiadas celdas para leerlo de una vez.", range));
        }
        let xml = self.package.get_string(&self.worksheet_part(sheet)?)?;
        let parsed = SheetXml::parse(&xml)?;
        let reader = CellReader { strings: self.shared_strings()?, styles: self.styles()?, date1904: self.date1904() };
        // Cells of a shared formula store only a link to its first cell: show their own formula
        let mut shared: HashMap<CellRef, String> = HashMap::new();
        let in_range = |at: CellRef| at.row >= range.first.row && at.row <= range.last.row && at.col >= range.first.col && at.col <= range.last.col;
        let has_shared = parsed.rows.iter().flat_map(|r| &r.cells).any(|c| {
            in_range(c.at) && c.formula.as_ref().is_some_and(|f| f.kind.as_deref() == Some("shared") && f.text.trim().is_empty())
        });
        if has_shared {
            for (at, expr) in sheet_formulas(&parsed) {
                if let (true, Ok(expr)) = (in_range(at), expr) {
                    shared.insert(at, expr.to_display());
                }
            }
        }
        let mut nodes: HashMap<CellRef, &CellNode> = HashMap::new();
        for row in parsed.rows.iter().filter(|r| r.number >= range.first.row && r.number <= range.last.row) {
            for cell in row.cells.iter().filter(|c| c.at.col >= range.first.col && c.at.col <= range.last.col) {
                nodes.insert(cell.at, cell);
            }
        }
        Ok((range.first.row..=range.last.row)
            .map(|row| {
                (range.first.col..=range.last.col)
                    .map(|col| {
                        let at = CellRef { row, col };
                        match nodes.get(&at) {
                            Some(node) => {
                                let mut data = reader.read(node);
                                if let Some(f) = data.formula.as_mut() {
                                    match shared.get(&at) {
                                        Some(text) if f.is_empty() => *f = text.clone(),
                                        _ => *f = display_formula(f),
                                    }
                                }
                                data
                            }
                            None => CellData { at, value: CellValue::Empty, formula: None },
                        }
                    })
                    .collect()
            })
            .collect())
    }

    /// Writes one value
    pub fn set_value(&mut self, sheet: usize, reference: &str, value: CellInput) -> Result<(), String> {
        self.set_values(sheet, reference, vec![vec![Some(value)]])
    }

    /// Writes rows of values from the cell `start` rightwards and down; `None` leaves a cell as
    /// it is. Text goes to the shared strings table (reusing an equal entry), dates get a date
    /// format based on the cell's current one, and a value replaces any formula. Nothing is
    /// written when any value is invalid.
    pub fn set_values(&mut self, sheet: usize, start: &str, rows: Vec<Vec<Option<CellInput>>>) -> Result<(), String> {
        let start = CellRef::parse(start)?;
        let total: u64 = rows.iter().map(|r| r.len() as u64).sum();
        if total > MAX_CELLS_PER_CALL {
            return Err("Demasiados valores para escribirlos de una vez.".to_string());
        }
        let mut inputs: BTreeMap<CellRef, CellInput> = BTreeMap::new();
        for (r, row) in rows.into_iter().enumerate() {
            for (c, value) in row.into_iter().enumerate() {
                if let Some(value) = value {
                    let at = start.offset(r, c)?;
                    if let CellInput::Number(n) | CellInput::Date(n) = value {
                        if !n.is_finite() {
                            return Err(format!("El valor de {} no es un número finito.", at));
                        }
                    }
                    inputs.insert(at, value);
                }
            }
        }
        if inputs.is_empty() {
            return Ok(());
        }

        let part = self.worksheet_part(sheet)?;
        let xml = self.package.get_string(&part)?;
        let parsed = SheetXml::parse(&xml)?;
        let prefix = parsed.prefix().to_string();
        let date1904 = self.date1904();

        let needs_strings = inputs.values().any(|v| matches!(v, CellInput::Text(_)));
        let strings_part = self.related_part("sharedStrings");
        let strings_xml = match (&strings_part, needs_strings) {
            (Some(part), true) => Some(self.package.get_string(part)?),
            (None, true) => Some(EMPTY_SST.to_string()),
            _ => None,
        };
        let mut strings = match &strings_xml {
            Some(xml) => SharedStrings::parse(xml)?,
            None => SharedStrings::default(),
        };
        let styles_part = self.related_part("styles");
        let mut styles_xml = match (&styles_part, inputs.values().any(|v| matches!(v, CellInput::Date(_)))) {
            (Some(part), true) => Some(self.package.get_string(part)?),
            (None, true) => Some(crate::blank::STYLES_XML.to_string()),
            _ => None,
        };
        let mut styles_changed = styles_part.is_none() && styles_xml.is_some();

        let mut edits: BTreeMap<CellRef, Option<String>> = BTreeMap::new();
        let mut formulas: HashMap<CellRef, String> = HashMap::new();
        for (at, input) in &inputs {
            let existing = parsed.cell(*at);
            if existing.and_then(|c| c.formula.as_ref()).is_some_and(|f| f.shared_master) {
                return Err(format!(
                    "La celda {} tiene una fórmula compartida que usan otras celdas; no se puede sobrescribir sin romperlas.",
                    at
                ));
            }
            let mut style_attr = existing.and_then(|c| c.style_attr.clone());
            let (kind, content) = match input {
                CellInput::Text(text) => (Some("s"), Some(strings.intern(text).to_string())),
                CellInput::Number(n) => (None, Some(format_number(*n))),
                CellInput::Bool(b) => (Some("b"), Some(if *b { "1" } else { "0" }.to_string())),
                CellInput::Date(ms) => {
                    let serial = unix_ms_to_serial(*ms, date1904).map_err(|e| format!("{} ({}).", e.trim_end_matches('.'), at))?;
                    let base = existing.map(|c| c.style).unwrap_or(0);
                    let stylesheet = styles_xml.as_mut().expect("loaded for dates");
                    let (index, changed) = date_style(stylesheet, base, has_time(*ms))?;
                    if let Some(changed) = changed {
                        *stylesheet = changed;
                        styles_changed = true;
                    }
                    if index != base || style_attr.is_none() && index != 0 {
                        style_attr = Some(index.to_string());
                    }
                    (None, Some(format_number(serial)))
                }
                CellInput::Formula(text) => {
                    let text = text.trim();
                    let text = text.strip_prefix('=').unwrap_or(text).trim();
                    let stored = match formula::parse(text) {
                        Ok(expr) => expr.to_formula(),
                        // Excel may know it: written as is and calculated when the file opens
                        Err(ParseError::Unsupported(_)) => text.to_string(),
                        Err(ParseError::Invalid(why)) => return Err(format!("La fórmula de {} no es válida: {}", at, why)),
                    };
                    formulas.insert(*at, stored);
                    (None, None)
                }
                CellInput::Clear => (None, None),
            };
            let phonetic = existing.and_then(|c| c.phonetic.as_deref());
            let element = match (formulas.get(at), &content, &style_attr) {
                (Some(f), _, _) => Some(formula_element(&prefix, *at, style_attr.as_deref(), phonetic, f)),
                (None, None, None) => None,
                _ => Some(cell_element(&prefix, *at, style_attr.as_deref(), phonetic, kind, content.as_deref())),
            };
            if existing.is_none() && element.is_none() {
                continue;
            }
            edits.insert(*at, element);
        }
        if edits.is_empty() {
            return Ok(());
        }

        // Every new text is ready: now write, so a failure above leaves the workbook untouched
        let new_sheet = parsed.splice(&edits);
        drop(parsed);
        self.package.put(part, new_sheet.into_bytes());
        if let (Some(xml), true) = (strings_xml, strings.changed()) {
            let written = strings.write(&xml)?;
            match strings_part {
                Some(part) => self.package.put(part, written.into_bytes()),
                None => self.add_workbook_part("sharedStrings", "sharedStrings.xml", SHARED_STRINGS_TYPE, written)?,
            }
        }
        if let (Some(xml), true) = (styles_xml, styles_changed) {
            match styles_part {
                Some(part) => self.package.put(part, xml.into_bytes()),
                None => self.add_workbook_part("styles", "styles.xml", STYLES_TYPE, xml)?,
            }
        }
        self.recalculate_after_edit();
        self.prepare_recalculation()
    }

    /// Replaces `{{FIELD}}` placeholders in the shared strings and in the inline strings of every
    /// worksheet with `values`; returns how many were replaced. Placeholders split across
    /// differently formatted runs of a rich text are not found.
    pub fn fill_template(&mut self, values: &HashMap<String, String>) -> Result<usize, String> {
        let mut changes = Vec::new();
        let mut total = 0;
        if let Some(part) = self.related_part("sharedStrings") {
            if let Some((xml, count)) = fill_xml(&self.package.get_string(&part)?, "si", values)? {
                changes.push((part, xml));
                total += count;
            }
        }
        for sheet in self.sheets()?.into_iter().filter(|s| s.kind == SheetKind::Worksheet && self.package.contains(&s.part)) {
            if let Some((xml, count)) = fill_xml(&self.package.get_string(&sheet.part)?, "is", values)? {
                changes.push((sheet.part, xml));
                total += count;
            }
        }
        if total > 0 {
            for (part, xml) in changes {
                self.package.put(part, xml.into_bytes());
            }
            self.recalculate_after_edit();
            self.prepare_recalculation()?;
        }
        Ok(total)
    }

    /// Recalculates every formula of the workbook and stores the results that changed in the
    /// cells; returns how many cells got a new value. Formulas this engine cannot evaluate keep
    /// the value Excel cached (Excel recalculates them when it opens the file).
    pub fn recalculate(&mut self) -> Result<usize, String> {
        let sheets = self.sheets()?;
        let mut xmls = Vec::with_capacity(sheets.len());
        for sheet in &sheets {
            xmls.push(match sheet.kind == SheetKind::Worksheet && self.package.contains(&sheet.part) {
                true => Some(self.package.get_string(&sheet.part)?),
                false => None,
            });
        }
        if !xmls.iter().flatten().any(|xml| xml.contains("<f") || xml.contains(":f")) {
            return Ok(0);
        }
        let names: Vec<String> = sheets.iter().map(|s| s.name.clone()).collect();
        let date1904 = self.date1904();
        let outcome = recalc::recalculate(
            &names,
            &xmls,
            &self.shared_strings()?,
            &defined_names(&self.workbook_xml()?),
            date1904,
            recalc::now_serial(date1904),
        )?;
        for (sheet, xml) in sheets.iter().zip(outcome.sheets) {
            if let Some(xml) = xml {
                self.package.put(sheet.part.clone(), xml.into_bytes());
            }
        }
        Ok(outcome.updated)
    }

    /// After an edit: a recalculation that fails leaves the cached values for Excel to refresh
    fn recalculate_after_edit(&mut self) {
        let _ = self.recalculate();
    }

    /// Adds a part related to the workbook: the part, its relationship and its content type
    fn add_workbook_part(&mut self, rel_name: &str, target: &str, content_type: &str, xml: String) -> Result<(), String> {
        let part = resolve_target(&self.workbook_part, target);
        let rels_part = rels_part_for(&self.workbook_part);
        let rels = self.package.get_string(&rels_part).unwrap_or_else(|_| {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#
                .to_string()
        });
        let (rels, _) = with_relationship(&rels, &format!("{}/{}", REL_NS, rel_name), target)?;
        let types = self.package.get_string(content_types::PART)?;
        let types = content_types::with_override(&types, &part, content_type)?.unwrap_or(types);
        self.package.put(part.clone(), xml.into_bytes());
        self.package.keep_in_order(&part);
        self.package.put(rels_part.clone(), rels.into_bytes());
        self.package.keep_in_order(&rels_part);
        self.package.put(content_types::PART.to_string(), types.into_bytes());
        Ok(())
    }

    /// After the first edit: drops the calculation chain (Excel rebuilds it, and a stale one can
    /// make it repair the file) and asks Excel to recalculate every formula on open, since cached
    /// values may depend on the cells written
    fn prepare_recalculation(&mut self) -> Result<(), String> {
        if self.recalc_prepared {
            return Ok(());
        }
        let rels_part = rels_part_for(&self.workbook_part);
        let chains: Vec<Relationship> = self.workbook_rels().into_iter().filter(|r| type_is(&r.rel_type, "calcChain")).collect();
        if !chains.is_empty() {
            let mut rels = self.package.get_string(&rels_part)?;
            let mut types = self.package.get_string(content_types::PART)?;
            for chain in &chains {
                let part = resolve_target(&self.workbook_part, &chain.target);
                self.package.remove(&part);
                rels = without_relationship(&rels, &chain.id)?;
                types = content_types::without_override(&types, &part)?;
            }
            self.package.put(rels_part, rels.into_bytes());
            self.package.put(content_types::PART.to_string(), types.into_bytes());
        }

        let xml = self.workbook_xml()?;
        let updated = with_full_calc_on_load(&xml)?;
        if updated != xml {
            self.package.put(self.workbook_part.clone(), updated.into_bytes());
        }
        self.recalc_prepared = true;
        Ok(())
    }

    /// Starts an undo step: the edits made until the next checkpoint undo together.
    /// `selection` is whatever the caller needs to restore its selection on undo (any text).
    pub fn checkpoint(&mut self, selection: Option<String>) {
        self.package.checkpoint(selection);
    }

    /// Where the selection ended up after the current step's edits (restored on redo)
    pub fn set_selection_after(&mut self, selection: Option<String>) {
        self.package.set_selection_after(selection);
    }

    /// Undoes the last step; `None` when there is nothing to undo
    pub fn undo(&mut self) -> Option<Undone> {
        let done = self.package.undo();
        self.after_history_move(&done);
        done
    }

    /// Redoes the last undone step
    pub fn redo(&mut self) -> Option<Undone> {
        let done = self.package.redo();
        self.after_history_move(&done);
        done
    }

    pub fn can_undo(&self) -> bool {
        self.package.can_undo()
    }

    pub fn can_redo(&self) -> bool {
        self.package.can_redo()
    }

    /// Undo may bring back the calc chain or drop `fullCalcOnLoad`: check again on the next edit
    /// (`prepare_recalculation` changes nothing when it is already in place)
    fn after_history_move(&mut self, done: &Option<Undone>) {
        if done.is_some() {
            self.recalc_prepared = false;
        }
    }

    /// The workbook as .xlsx bytes: untouched parts are written back as they were read
    pub fn to_bytes(&self) -> Result<Vec<u8>, String> {
        self.package.to_bytes()
    }

    #[cfg(test)]
    pub(crate) fn package(&self) -> &Package {
        &self.package
    }
}

/// Reads cell values with the tables they refer to
struct CellReader {
    strings: SharedStrings,
    styles: Styles,
    date1904: bool,
}

impl CellReader {
    fn read(&self, node: &CellNode) -> CellData {
        let raw = node.value.as_deref().unwrap_or("");
        let value = match node.kind.as_deref() {
            Some("s") => raw
                .trim()
                .parse::<usize>()
                .ok()
                .and_then(|i| self.strings.get(i))
                .map(|s| CellValue::Text(s.to_string()))
                .unwrap_or(CellValue::Empty),
            Some("str") => CellValue::Text(raw.to_string()),
            Some("inlineStr") => CellValue::Text(node.inline.clone().unwrap_or_else(|| raw.to_string())),
            Some("b") => CellValue::Bool(matches!(raw.trim(), "1" | "true")),
            Some("e") => CellValue::Error(raw.to_string()),
            Some("d") => iso_to_unix_ms(raw).map(CellValue::Date).unwrap_or(CellValue::Text(raw.to_string())),
            _ if raw.trim().is_empty() => CellValue::Empty,
            _ => match raw.trim().parse::<f64>() {
                Ok(n) if self.styles.is_date_style(node.style) => CellValue::Date(serial_to_unix_ms(n, self.date1904)),
                Ok(n) => CellValue::Number(n),
                Err(_) => CellValue::Text(raw.to_string()),
            },
        };
        CellData { at: node.at, value, formula: node.formula.as_ref().map(|f| f.text.clone()) }
    }
}

/// A number as written in `<v>`: the shortest text that reads back as the same number
pub(crate) fn format_number(n: f64) -> String {
    if n == 0.0 {
        "0".to_string()
    } else if n.abs() >= 1e16 || n.abs() < 1e-6 {
        format!("{:e}", n)
    } else {
        format!("{}", n)
    }
}

/// A formula as people read it: without the `_xlfn.` prefixes Excel stores on newer functions
fn display_formula(stored: &str) -> String {
    if stored.contains("_xlfn.") || stored.contains("_xlws.") {
        formula::parse(stored).map(|e| e.to_display()).unwrap_or_else(|_| stored.to_string())
    } else {
        stored.to_string()
    }
}

/// A `<c>` element with a formula and no cached value (the recalculation fills it in)
fn formula_element(prefix: &str, at: CellRef, style: Option<&str>, phonetic: Option<&str>, formula: &str) -> String {
    let mut out = format!(r#"<{}c r="{}""#, prefix, at);
    if let Some(s) = style {
        out.push_str(&format!(r#" s="{}""#, escape_xml(s)));
    }
    if let Some(ph) = phonetic {
        out.push_str(&format!(r#" ph="{}""#, escape_xml(ph)));
    }
    out.push_str(&format!("><{p}f>{}</{p}f></{p}c>", escape_xml(formula), p = prefix));
    out
}

/// A `<c>` element with only the attributes that survive a new value: reference, format and
/// phonetic flag
fn cell_element(prefix: &str, at: CellRef, style: Option<&str>, phonetic: Option<&str>, kind: Option<&str>, value: Option<&str>) -> String {
    let mut out = format!(r#"<{}c r="{}""#, prefix, at);
    if let Some(s) = style {
        out.push_str(&format!(r#" s="{}""#, escape_xml(s)));
    }
    if let Some(ph) = phonetic {
        out.push_str(&format!(r#" ph="{}""#, escape_xml(ph)));
    }
    if let Some(t) = kind {
        out.push_str(&format!(r#" t="{}""#, t));
    }
    match value {
        Some(v) => out.push_str(&format!("><{p}v>{}</{p}v></{p}c>", escape_xml(v), p = prefix)),
        None => out.push_str("/>"),
    }
    out
}

/// Elements that follow `<calcPr>` in a workbook, per the schema order
const AFTER_CALC_PR: &[&str] = &[
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "fileRecoveryPr",
    "webPublishObjects",
    "extLst",
];

/// Offset of the first child of the root element named one of `names`
fn first_child_named(xml: &str, names: &[&str]) -> Result<Option<usize>, String> {
    let mut reader = Reader::from_str(xml);
    let mut depth = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event().map_err(|e| format!("El libro no es XML válido: {}", e))? {
            Event::Start(e) | Event::Empty(e) if depth == 1 && names.iter().any(|n| tag_is(e.name().as_ref(), n)) => {
                return Ok(Some(before));
            }
            Event::Start(_) => depth += 1,
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => return Ok(None),
            _ => {}
        }
    }
}

/// The workbook with `fullCalcOnLoad="1"` in its `<calcPr>` (added in schema order if missing)
fn with_full_calc_on_load(xml: &str) -> Result<String, String> {
    if let Some(range) = element_ranges(xml, "calcPr", |_| true)?.into_iter().next() {
        let updated = set_attr(&xml[range.clone()], "fullCalcOnLoad", "1");
        return Ok(format!("{}{}{}", &xml[..range.start], updated, &xml[range.end..]));
    }
    let (_, root) = root_element(xml).ok_or("El libro no tiene elemento raíz.")?;
    let prefix = prefix_of(&root);
    let element = format!(r#"<{}calcPr fullCalcOnLoad="1"/>"#, prefix);
    let insert_at = match first_child_named(xml, AFTER_CALC_PR)? {
        Some(at) => at,
        None => xml.rfind(&format!("</{}>", root)).ok_or("El libro no tiene cierre de <workbook>.")?,
    };
    Ok(format!("{}{}{}", &xml[..insert_at], element, &xml[insert_at..]))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::blank::{blank_xlsx, zip_parts};

    const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/libro.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/hojas/datos.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/hojas/resumen.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#;
    const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/libro.xml"/></Relationships>"#;
    const WORKBOOK: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="1"/><bookViews><workbookView activeTab="0"><extLst/></workbookView></bookViews><sheets><sheet name="Resumen" sheetId="3" r:id="rId9"/><sheet name="Datos" sheetId="1" r:id="rId2"/><sheet name="Gráfico" sheetId="2" r:id="rId4"/></sheets><definedNames><definedName name="Total">Resumen!$B$2</definedName></definedNames><calcPr calcId="191029"/><extLst><ext uri="x"/></extLst></workbook>"#;
    const WORKBOOK_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="hojas/datos.xml"/><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/hojas/resumen.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>"#;
    const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts><fonts count="2"><font><sz val="11"/></font><font><b/><sz val="11"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="4" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>"#;
    const SHARED: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4"><si><t>Cliente</t></si><si><t>Informe para {{CLIENTE}}</t></si><si><r><rPr><b/></rPr><t>Total</t></r><r><t xml:space="preserve"> {{MONEDA}}</t></r></si><si><t>Fecha</t></si></sst>"#;
    const DATOS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:D3"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r="1" spans="1:4"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>3</v></c><c r="C1" t="inlineStr"><is><t>Nota {{NOTA}}</t></is></c></row><row r="2" spans="1:4"><c r="A2" s="2"><v>1234.5</v></c><c r="B2" s="1"><v>43844</v></c><c r="C2" t="b"><v>1</v></c><c r="D2" t="e"><v>#DIV/0!</v></c></row><row r="3" spans="1:4"><c r="A3" s="2"><f>SUM(A2:A2)</f><v>1234.5</v></c><c r="B3" t="str"><f>A1&amp;"!"</f><v>Cliente!</v></c><c r="C3"><f t="shared" ref="C3:C4" si="0">A3*2</f><v>2469</v></c><c r="D3" t="d"><v>2024-01-15T12:00:00Z</v></c></row><row r="4"><c r="C4"><f t="shared" si="0"/><v>0</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A6:B6"/></mergeCells><drawing r:id="rId1"/></worksheet>"#;
    const RESUMEN: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row><row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><f>Datos!A2</f><v>1234.5</v></c></row></sheetData></worksheet>"#;
    const CALC_CHAIN: &str = r#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A3" i="2"/></calcChain>"#;

    /// A template like the ones Excel saves: parts in non-default places, a chart, a calc chain
    fn template() -> Vec<u8> {
        zip_parts(&[
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELS),
            ("xl/libro.xml", WORKBOOK),
            ("xl/_rels/libro.xml.rels", WORKBOOK_RELS),
            ("xl/hojas/datos.xml", DATOS),
            ("xl/hojas/_rels/datos.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#),
            ("xl/hojas/resumen.xml", RESUMEN),
            ("xl/chartsheets/sheet1.xml", r#"<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#),
            ("xl/drawings/drawing1.xml", "<xdr:wsDr xmlns:xdr=\"d\">  <!-- dibujo -->  </xdr:wsDr>"),
            ("xl/charts/chart1.xml", "<c:chartSpace xmlns:c=\"c\"><c:chart/></c:chartSpace>"),
            ("xl/styles.xml", STYLES),
            ("xl/sharedStrings.xml", SHARED),
            ("xl/calcChain.xml", CALC_CHAIN),
            ("docProps/app.xml", "<Properties>\r\n  <Application>Microsoft Excel</Application>\r\n</Properties>"),
        ])
        .unwrap()
    }

    fn open(bytes: &[u8]) -> Workbook {
        Workbook::from_bytes(bytes).unwrap()
    }

    fn value(wb: &Workbook, sheet: usize, r: &str) -> CellValue {
        wb.cell(sheet, r).unwrap().value
    }

    fn part(wb: &Workbook, name: &str) -> String {
        wb.package().get_string(name).unwrap()
    }

    const JAN_15_2024_NOON: f64 = 1_705_320_000_000.0;

    #[test]
    fn reads_the_sheets_of_a_blank_workbook() {
        let wb = open(&blank_xlsx().unwrap());
        assert_eq!(wb.sheet_names().unwrap(), vec!["Hoja1".to_string()]);
        assert_eq!(value(&wb, 0, "A1"), CellValue::Empty);
    }

    #[test]
    fn refuses_a_zip_that_is_not_a_workbook() {
        let docx_like = zip_parts(&[("word/document.xml", "<w:document/>")]).unwrap();
        let err = Workbook::from_bytes(&docx_like).err().unwrap();
        assert!(err.contains(WORKBOOK_PART), "{}", err);
    }

    #[test]
    fn finds_sheets_through_their_relationships() {
        let wb = open(&template());
        let sheets = wb.sheets().unwrap();
        assert_eq!(sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(), ["Resumen", "Datos", "Gráfico"]);
        assert_eq!(sheets[0].part, "xl/hojas/resumen.xml");
        assert_eq!(sheets[1].part, "xl/hojas/datos.xml");
        assert_eq!(sheets[2].kind, SheetKind::Chartsheet);
        assert_eq!(wb.sheet_index("datos").unwrap(), 1);
        assert!(wb.sheet_index("Otra").is_err());
        assert!(wb.cell(2, "A1").unwrap_err().contains("gráfico"));
        assert!(wb.cell(7, "A1").is_err());
    }

    #[test]
    fn reads_every_kind_of_value() {
        let wb = open(&template());
        assert!(wb.date1904());
        assert_eq!(value(&wb, 1, "A1"), CellValue::Text("Cliente".into()));
        assert_eq!(value(&wb, 1, "C1"), CellValue::Text("Nota {{NOTA}}".into()));
        assert_eq!(value(&wb, 1, "A2"), CellValue::Number(1234.5));
        // 43844 in the 1904 system is 2024-01-15
        assert_eq!(value(&wb, 1, "B2"), CellValue::Date(JAN_15_2024_NOON - 43_200_000.0));
        assert_eq!(value(&wb, 1, "C2"), CellValue::Bool(true));
        assert_eq!(value(&wb, 1, "D2"), CellValue::Error("#DIV/0!".into()));
        assert_eq!(value(&wb, 1, "D3"), CellValue::Date(JAN_15_2024_NOON));
        assert_eq!(value(&wb, 0, "A2"), CellValue::Text("Total {{MONEDA}}".into()));
        assert_eq!(value(&wb, 1, "Z99"), CellValue::Empty);

        let a3 = wb.cell(1, "A3").unwrap();
        assert_eq!((a3.value, a3.formula.as_deref()), (CellValue::Number(1234.5), Some("SUM(A2:A2)")));
        let b3 = wb.cell(1, "B3").unwrap();
        assert_eq!((b3.value, b3.formula.as_deref()), (CellValue::Text("Cliente!".into()), Some("A1&\"!\"")));
        assert_eq!(wb.cell(1, "C4").unwrap().formula.as_deref(), Some("A4*2"));

        let range = wb.range(1, "A1:B2").unwrap();
        assert_eq!(range.len(), 2);
        assert_eq!(range[1][0].at.to_string(), "A2");
        assert_eq!(range[0][1].value, CellValue::Text("Fecha".into()));
    }

    #[test]
    fn dates_in_the_1900_system() {
        let mut wb = open(&blank_xlsx().unwrap());
        wb.set_value(0, "A1", CellInput::Date(JAN_15_2024_NOON)).unwrap();
        wb.set_value(0, "A2", CellInput::Date(JAN_15_2024_NOON - 43_200_000.0)).unwrap();
        assert_eq!(value(&wb, 0, "A1"), CellValue::Date(JAN_15_2024_NOON));
        let sheet = part(&wb, "xl/worksheets/sheet1.xml");
        assert!(sheet.contains(r#"<c r="A1" s="1"><v>45306.5</v></c>"#), "{}", sheet);
        assert!(sheet.contains(r#"<c r="A2" s="2"><v>45306</v></c>"#), "{}", sheet);
        let styles = part(&wb, "xl/styles.xml");
        assert!(styles.contains(r#"<cellXfs count="3">"#), "{}", styles);
        assert!(styles.contains(r#"numFmtId="22""#) && styles.contains(r#"numFmtId="14""#));
    }

    #[test]
    fn writes_values_in_place_and_in_new_rows() {
        let mut wb = open(&template());
        wb.set_values(
            1,
            "B2",
            vec![
                vec![Some(CellInput::Text("Cliente".into())), Some(CellInput::Number(7.0)), None, Some(CellInput::Bool(false))],
                vec![Some(CellInput::Number(-0.25)), None, Some(CellInput::Clear)],
                vec![],
                vec![None, None, None, Some(CellInput::Text("Nuevo".into()))],
            ],
        )
        .unwrap();
        // B2 had a date format: text replaces the value but keeps the format
        assert_eq!(value(&wb, 1, "B2"), CellValue::Text("Cliente".into()));
        assert_eq!(value(&wb, 1, "C2"), CellValue::Number(7.0));
        assert_eq!(value(&wb, 1, "D2"), CellValue::Error("#DIV/0!".into()));
        assert_eq!(value(&wb, 1, "E2"), CellValue::Bool(false));
        assert_eq!(wb.cell(1, "B3").unwrap(), CellData { at: CellRef::parse("B3").unwrap(), value: CellValue::Number(-0.25), formula: None });
        assert_eq!(value(&wb, 1, "D3"), CellValue::Empty);
        assert_eq!(value(&wb, 1, "E5"), CellValue::Text("Nuevo".into()));

        let sheet = part(&wb, "xl/hojas/datos.xml");
        assert!(sheet.contains(r#"<row r="2"><c r="A2" s="2"><v>1234.5</v></c><c r="B2" s="1" t="s"><v>0</v></c><c r="C2"><v>7</v></c><c r="D2" t="e"><v>#DIV/0!</v></c><c r="E2" t="b"><v>0</v></c></row>"#), "{}", sheet);
        assert!(sheet.contains(r#"<c r="B3"><v>-0.25</v></c><c r="C3"><f t="shared" ref="C3:C4" si="0">A3*2</f><v>2469</v></c></row>"#), "{}", sheet);
        assert!(sheet.contains(r#"<row r="4"><c r="C4"><f t="shared" si="0"/><v>0</v></c></row><row r="5"><c r="E5" t="s"><v>4</v></c></row></sheetData>"#), "{}", sheet);
        assert!(sheet.contains(r#"<dimension ref="A1:E5"/>"#));
        assert!(sheet.ends_with(r#"<mergeCells count="1"><mergeCell ref="A6:B6"/></mergeCells><drawing r:id="rId1"/></worksheet>"#));

        // "Cliente" reused, "Nuevo" appended
        let shared = part(&wb, "xl/sharedStrings.xml");
        assert!(shared.contains(r#"count="6" uniqueCount="5""#), "{}", shared);
        assert!(shared.ends_with("<si><t>Nuevo</t></si></sst>"));
    }

    #[test]
    fn writes_dates_with_a_date_format_based_on_the_cell() {
        let mut wb = open(&template());
        // A2 has a bold number format; B2 already shows dates
        wb.set_values(1, "A2", vec![vec![Some(CellInput::Date(JAN_15_2024_NOON)), Some(CellInput::Date(JAN_15_2024_NOON - 43_200_000.0))]]).unwrap();
        assert_eq!(value(&wb, 1, "A2"), CellValue::Date(JAN_15_2024_NOON));
        assert_eq!(value(&wb, 1, "B2"), CellValue::Date(JAN_15_2024_NOON - 43_200_000.0));
        let sheet = part(&wb, "xl/hojas/datos.xml");
        assert!(sheet.contains(r#"<c r="A2" s="3"><v>43844.5</v></c><c r="B2" s="1"><v>43844</v></c>"#), "{}", sheet);
        let styles = part(&wb, "xl/styles.xml");
        assert!(styles.contains(r#"<xf numFmtId="22" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/></cellXfs>"#), "{}", styles);

        assert!(wb.set_value(1, "A9", CellInput::Date(-3_000_000_000_000.0)).is_err(), "before 1904");
    }

    #[test]
    fn refuses_to_break_shared_formulas_and_keeps_the_workbook_unchanged() {
        let mut wb = open(&template());
        // A valid value next to the shared formula: the whole batch is refused
        let err = wb.set_values(1, "B3", vec![vec![Some(CellInput::Text("nuevo".into())), Some(CellInput::Number(1.0))]]);
        assert_eq!(part(&wb, "xl/hojas/datos.xml"), DATOS);
        assert_eq!(part(&wb, "xl/sharedStrings.xml"), SHARED);
        assert!(wb.package().contains("xl/calcChain.xml"));
        assert!(err.unwrap_err().contains("compartida"));
        assert!(wb.set_value(1, "C4", CellInput::Number(1.0)).is_ok(), "a cell that only reuses it can be overwritten");
        assert!(wb.set_value(1, "A1", CellInput::Number(f64::NAN)).is_err());
        assert!(wb.set_value(1, "A0", CellInput::Number(1.0)).is_err());
    }

    #[test]
    fn a_change_drops_the_calc_chain_and_asks_for_recalculation() {
        let mut wb = open(&template());
        wb.set_value(0, "C1", CellInput::Number(1.0)).unwrap();
        let out = open(&wb.to_bytes().unwrap());
        assert!(!out.package().contains("xl/calcChain.xml"));
        let rels = part(&out, "xl/_rels/libro.xml.rels");
        assert!(!rels.contains("calcChain") && rels.contains(r#"Id="rId6""#));
        let types = part(&out, "[Content_Types].xml");
        assert!(!types.contains("calcChain") && types.contains("/xl/charts/chart1.xml"));
        let workbook = part(&out, "xl/libro.xml");
        assert_eq!(workbook, WORKBOOK.replace(r#"<calcPr calcId="191029"/>"#, r#"<calcPr calcId="191029" fullCalcOnLoad="1"/>"#));

        // Without <calcPr>, it is added before the elements that follow it, not in nested ones
        let no_calc = WORKBOOK.replace(r#"<calcPr calcId="191029"/>"#, "");
        assert_eq!(with_full_calc_on_load(&no_calc).unwrap(), WORKBOOK.replace(r#"<calcPr calcId="191029"/>"#, r#"<calcPr fullCalcOnLoad="1"/>"#));
        let bare = r#"<x:workbook xmlns:x="m"><x:sheets/></x:workbook>"#;
        assert_eq!(with_full_calc_on_load(bare).unwrap(), r#"<x:workbook xmlns:x="m"><x:sheets/><x:calcPr fullCalcOnLoad="1"/></x:workbook>"#);
    }

    #[test]
    fn undoes_and_redoes_edits_step_by_step() {
        let original = template();
        let mut wb = open(&original);
        assert!(!wb.can_undo());
        wb.checkpoint(Some("A1".into()));
        wb.set_value(0, "A1", CellInput::Text("Nuevo".into())).unwrap();
        wb.set_selection_after(Some("A2".into()));
        wb.checkpoint(Some("C2".into()));
        wb.set_value(0, "C2", CellInput::Number(7.0)).unwrap();
        assert_eq!(value(&wb, 0, "C2"), CellValue::Number(7.0));

        let done = wb.undo().unwrap();
        assert_eq!(done.selection.as_deref(), Some("C2"));
        assert_eq!(value(&wb, 0, "C2"), CellValue::Empty);
        assert_eq!(value(&wb, 0, "A1"), CellValue::Text("Nuevo".into()));
        let done = wb.undo().unwrap();
        assert_eq!(done.selection.as_deref(), Some("A1"));
        assert!(!wb.can_undo() && wb.can_redo());
        // Back to the file as it was read, calc chain included
        for name in ["xl/hojas/resumen.xml", "xl/libro.xml", "xl/sharedStrings.xml", "[Content_Types].xml"] {
            assert_eq!(part(&wb, name), part(&open(&original), name), "{}", name);
        }
        assert!(wb.package().contains("xl/calcChain.xml"));

        assert_eq!(wb.redo().unwrap().selection.as_deref(), Some("A2"));
        assert_eq!(value(&wb, 0, "A1"), CellValue::Text("Nuevo".into()));
        assert!(!wb.package().contains("xl/calcChain.xml"));

        // A new edit after undoing everything asks for recalculation again
        wb.undo().unwrap();
        wb.checkpoint(None);
        wb.set_value(0, "C1", CellInput::Bool(true)).unwrap();
        assert!(!wb.can_redo());
        assert!(part(&wb, "xl/libro.xml").contains(r#"fullCalcOnLoad="1""#));
        assert!(!wb.package().contains("xl/calcChain.xml"));
    }

    #[test]
    fn writes_formulas_and_recalculates_what_depends_on_an_edit() {
        let mut wb = open(&blank_xlsx().unwrap());
        wb.set_values(0, "A1", vec![vec![Some(CellInput::Number(2.0)), Some(CellInput::Number(3.0))]]).unwrap();
        wb.set_value(0, "C1", CellInput::Formula("=SUM(A1:B1)*10%".into())).unwrap();
        wb.set_value(0, "D1", CellInput::Formula("xlookup(2,A1:B1,A1:B1)".into())).unwrap();
        assert_eq!(value(&wb, 0, "C1"), CellValue::Number(0.5));
        assert_eq!(wb.cell(0, "D1").unwrap().formula.as_deref(), Some("XLOOKUP(2,A1:B1,A1:B1)"));
        assert!(part(&wb, "xl/worksheets/sheet1.xml").contains("<f>_xlfn.XLOOKUP(2,A1:B1,A1:B1)</f><v>2</v>"));

        wb.checkpoint(None);
        wb.set_value(0, "A1", CellInput::Number(7.0)).unwrap();
        assert_eq!(value(&wb, 0, "C1"), CellValue::Number(1.0));
        wb.undo();
        assert_eq!(value(&wb, 0, "C1"), CellValue::Number(0.5));

        // Unknown to this engine: written, left for Excel to calculate
        wb.set_value(0, "E1", CellInput::Formula("OFFSET(A1,0,1)".into())).unwrap();
        assert_eq!(wb.cell(0, "E1").unwrap(), CellData { at: CellRef::parse("E1").unwrap(), value: CellValue::Empty, formula: Some("OFFSET(A1,0,1)".into()) });
        let err = wb.set_value(0, "F1", CellInput::Formula("SUM(A1".into())).unwrap_err();
        assert!(err.contains("F1"), "{}", err);
    }

    #[test]
    fn untouched_parts_round_trip_byte_for_byte() {
        let original = open(&template());
        // Opening and saving without changes keeps every part
        let same = open(&original.to_bytes().unwrap());
        assert_eq!(original.package().order(), same.package().order());
        for name in original.package().order() {
            assert_eq!(original.package().get(name), same.package().get(name), "part {}", name);
        }

        let mut wb = open(&template());
        wb.set_value(1, "A2", CellInput::Number(99.0)).unwrap();
        let edited = open(&wb.to_bytes().unwrap());
        // The summary sheet reads the edited cell, so its cached result changes too
        let touched = ["xl/hojas/datos.xml", "xl/hojas/resumen.xml", "xl/libro.xml", "xl/_rels/libro.xml.rels", "[Content_Types].xml"];
        for name in original.package().order() {
            if name == "xl/calcChain.xml" {
                assert!(!edited.package().contains(name));
            } else if !touched.contains(&name.as_str()) {
                assert_eq!(original.package().get(name), edited.package().get(name), "part {}", name);
            }
        }
        // In the edited sheet only the cell and the cached results of the formulas that read it
        // changed
        let expected = DATOS
            .replace(r#"<c r="A2" s="2"><v>1234.5</v></c>"#, r#"<c r="A2" s="2"><v>99</v></c>"#)
            .replace(r#"<f>SUM(A2:A2)</f><v>1234.5</v>"#, r#"<f>SUM(A2:A2)</f><v>99</v>"#)
            .replace(r#"si="0">A3*2</f><v>2469</v>"#, r#"si="0">A3*2</f><v>198</v>"#);
        assert_eq!(part(&edited, "xl/hojas/datos.xml"), expected);
        assert_eq!(part(&edited, "xl/hojas/resumen.xml"), RESUMEN.replace("<v>1234.5</v>", "<v>99</v>"));
        let names: Vec<_> = edited.package().order().to_vec();
        assert_eq!(names, original.package().order().iter().filter(|n| *n != "xl/calcChain.xml").cloned().collect::<Vec<_>>());
    }

    #[test]
    fn fills_placeholders_in_shared_and_inline_strings() {
        let mut wb = open(&template());
        let values: HashMap<String, String> =
            [("CLIENTE", "Acme & Cía"), ("MONEDA", "EUR"), ("NOTA", "<b>")].iter().map(|(k, v)| (k.to_string(), v.to_string())).collect();
        assert_eq!(wb.fill_template(&values).unwrap(), 3);
        assert_eq!(value(&wb, 0, "A1"), CellValue::Text("Informe para Acme & Cía".into()));
        assert_eq!(value(&wb, 0, "A2"), CellValue::Text("Total EUR".into()));
        assert_eq!(value(&wb, 1, "C1"), CellValue::Text("Nota <b>".into()));
        assert!(part(&wb, "xl/sharedStrings.xml").contains(r#"<r><t xml:space="preserve"> EUR</t></r>"#));
        assert!(!wb.package().contains("xl/calcChain.xml"));
        assert_eq!(wb.fill_template(&values).unwrap(), 0, "nothing left to fill");

        let mut untouched = open(&template());
        let unknown: HashMap<String, String> = [("OTRO".to_string(), "x".to_string())].into();
        assert_eq!(untouched.fill_template(&unknown).unwrap(), 0);
        assert!(untouched.package().contains("xl/calcChain.xml"), "no change, no recalculation set-up");
    }

    #[test]
    fn creates_the_shared_strings_part_when_missing() {
        let mut wb = open(&blank_xlsx().unwrap());
        assert!(!wb.package().contains("xl/sharedStrings.xml"));
        wb.set_values(0, "A1", vec![vec![Some(CellInput::Text("Hola".into())), Some(CellInput::Text("Hola".into()))]]).unwrap();
        let out = open(&wb.to_bytes().unwrap());
        assert_eq!(value(&out, 0, "B1"), CellValue::Text("Hola".into()));
        assert!(part(&out, "xl/_rels/workbook.xml.rels").contains(r#"<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>"#));
        assert!(part(&out, "[Content_Types].xml").contains(r#"<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"#));
        assert!(part(&out, "xl/sharedStrings.xml").contains(r#"count="2" uniqueCount="1"><si><t>Hola</t></si>"#));
        assert!(part(&out, "xl/workbook.xml").contains(r#"<calcPr fullCalcOnLoad="1"/></workbook>"#));
    }

    #[test]
    fn formats_numbers_so_they_read_back_the_same() {
        for n in [0.0, -0.0, 1.0, 0.1 + 0.2, 1e20, 1.5e-9, -123456.789, f64::MAX] {
            let text = format_number(n);
            assert_eq!(text.parse::<f64>().unwrap(), if n == 0.0 { 0.0 } else { n }, "{}", text);
        }
        assert_eq!(format_number(42.0), "42");
    }
}
