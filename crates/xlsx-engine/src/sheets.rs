//! Adding and renaming sheets. A new sheet is an empty worksheet part with its relationship,
//! content type and `<sheet>` entry; renaming also rewrites the references to the old name in
//! formulas, defined names, data validations, conditional formats and charts.

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use ooxml_core::rels::resolve_target;
use ooxml_core::xml::{element_ranges, escape_xml, get_attr_value, prefix_of, root_element, set_attr, tag_is, unescape_text};

use crate::blank::SHEET_XML;
use crate::formula::quote_sheet;
use crate::workbook::{SheetInfo, Workbook, REL_NS};

const WORKSHEET_TYPE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

/// Longest sheet name Excel accepts
pub const MAX_SHEET_NAME: usize = 31;

/// Elements whose text is a formula, in worksheets (`f` also in charts, as `c:f`)
pub(crate) const FORMULA_ELEMENTS: &[&str] = &["f", "formula", "formula1", "formula2"];

/// Checks a sheet name as Excel does: 1–31 characters, none of `[ ] : * ? / \`, not starting or
/// ending with an apostrophe, not `History`, and not the name of another sheet (ignoring case).
/// `except` is the sheet being renamed, which may keep its own name.
pub fn validate_sheet_name(name: &str, sheets: &[SheetInfo], except: Option<usize>) -> Result<(), String> {
    let length = name.chars().count();
    if length == 0 || name.trim().is_empty() {
        return Err("El nombre de la hoja no puede estar vacío.".to_string());
    }
    if length > MAX_SHEET_NAME {
        return Err(format!("El nombre de la hoja tiene {} caracteres; el máximo es {}.", length, MAX_SHEET_NAME));
    }
    if let Some(c) = name.chars().find(|c| "[]:*?/\\".contains(*c) || c.is_control()) {
        return Err(format!("El nombre de la hoja no puede llevar '{}'.", c.escape_default()));
    }
    if name.starts_with('\'') || name.ends_with('\'') {
        return Err("El nombre de la hoja no puede empezar ni terminar con un apóstrofo.".to_string());
    }
    if name.eq_ignore_ascii_case("History") {
        return Err("Excel reserva el nombre 'History'.".to_string());
    }
    let lower = name.to_lowercase();
    if sheets.iter().enumerate().any(|(i, s)| Some(i) != except && s.name.to_lowercase() == lower) {
        return Err(format!("Ya hay una hoja llamada '{}'.", name));
    }
    Ok(())
}

impl Workbook {
    /// Adds an empty worksheet called `name` after the tab at `after` (at the end when `None`).
    /// Returns its tab position.
    pub fn add_sheet(&mut self, name: &str, after: Option<usize>) -> Result<usize, String> {
        let sheets = self.sheets()?;
        validate_sheet_name(name, &sheets, None)?;
        let position = match after {
            Some(i) if i < sheets.len() => i + 1,
            Some(i) => return Err(format!("El libro tiene {} hoja(s); no existe la hoja {}.", sheets.len(), i)),
            None => sheets.len(),
        };

        let mut n = sheets.len() + 1;
        let target = loop {
            let target = format!("worksheets/sheet{}.xml", n);
            if !self.package.contains(&resolve_target(&self.workbook_part, &target)) {
                break target;
            }
            n += 1;
        };
        let rel_id = self.add_workbook_part("worksheet", &target, WORKSHEET_TYPE, SHEET_XML.to_string())?;

        let xml = self.workbook_xml()?;
        let ranges = element_ranges(&xml, "sheet", |_| true)?;
        let first = ranges.first().ok_or("El libro no tiene la lista de hojas.")?;
        let prefix = start_tag_name(&xml[first.clone()]).map(|q| prefix_of(&q).to_string()).unwrap_or_default();
        let sheet_id = ranges
            .iter()
            .filter_map(|r| start_tag_attr(&xml[r.clone()], "sheetId").and_then(|v| v.parse::<u32>().ok()))
            .max()
            .unwrap_or(0)
            + 1;
        let id_attr = match relationships_prefix(&xml) {
            Some(r) => format!(r#"{}:id="{}""#, r, rel_id),
            None => format!(r#"xmlns:r="{}" r:id="{}""#, REL_NS, rel_id),
        };
        let entry = format!(r#"<{}sheet name="{}" sheetId="{}" {}/>"#, prefix, escape_xml(name), sheet_id, id_attr);
        let at = if position == 0 { first.start } else { ranges[position - 1].end };
        let mut updated = format!("{}{}{}", &xml[..at], entry, &xml[at..]);
        // Tab positions after the new sheet move one place
        updated = shift_index_attr(&updated, "definedName", "localSheetId", position)?;
        updated = shift_index_attr(&updated, "workbookView", "activeTab", position)?;
        self.package.put(self.workbook_part.clone(), updated.into_bytes());
        Ok(position)
    }

    /// Renames the tab at `index` and the references to it in formulas, defined names, data
    /// validations, conditional formats and charts. Returns how many of those texts changed.
    /// References inside text (`INDIRECT("Hoja1!A1")`) and from other workbooks are left as
    /// they are, as Excel does.
    pub fn rename_sheet(&mut self, index: usize, name: &str) -> Result<usize, String> {
        let sheets = self.sheets()?;
        let old = sheets
            .get(index)
            .ok_or_else(|| format!("El libro tiene {} hoja(s); no existe la hoja {}.", sheets.len(), index))?
            .name
            .clone();
        validate_sheet_name(name, &sheets, Some(index))?;
        if old == name {
            return Ok(0);
        }
        let rename = |formula: &str| rename_in_formula(formula, &old, name);

        let xml = self.workbook_xml()?;
        let range = element_ranges(&xml, "sheet", |_| true)?
            .into_iter()
            .nth(index)
            .ok_or("La lista de hojas no coincide con el libro.")?;
        let xml = format!("{}{}{}", &xml[..range.start], set_attr(&xml[range.clone()], "name", name), &xml[range.end..]);
        let (xml, mut changed) = rewrite_texts(&xml, &["definedName"], &rename)?;
        self.package.put(self.workbook_part.clone(), xml.into_bytes());

        let mut parts: Vec<String> = sheets.iter().filter(|s| self.package.contains(&s.part)).map(|s| s.part.clone()).collect();
        let mut charts: Vec<String> = self
            .package
            .parts()
            .keys()
            .filter(|p| p.starts_with("xl/charts/") && p.ends_with(".xml") && !p.contains("/_rels/"))
            .cloned()
            .collect();
        charts.sort();
        parts.extend(charts);
        for part in parts {
            let xml = self.package.get_string(&part)?;
            let (written, n) = rewrite_texts(&xml, FORMULA_ELEMENTS, &rename)?;
            if n > 0 {
                self.package.put(part, written.into_bytes());
                changed += n;
            }
        }
        Ok(changed)
    }
}

/// Qualified name of the element that starts `element`
fn start_tag_name(element: &str) -> Option<String> {
    let mut reader = Reader::from_str(element);
    match reader.read_event() {
        Ok(Event::Start(e)) | Ok(Event::Empty(e)) => Some(ooxml_core::xml::utf8(e.name().as_ref()).to_string()),
        _ => None,
    }
}

/// An attribute of the start tag of `element`
fn start_tag_attr(element: &str, name: &str) -> Option<String> {
    let mut reader = Reader::from_str(element);
    match reader.read_event() {
        Ok(Event::Start(e)) | Ok(Event::Empty(e)) => get_attr_value(&e, name),
        _ => None,
    }
}

/// Prefix the workbook's root element declares for the relationships namespace (`r`)
fn relationships_prefix(xml: &str) -> Option<String> {
    let (start, _) = root_element(xml)?;
    let tag = &xml[start..start + xml[start..].find('>')?];
    let at = tag.find(&format!("\"{}\"", REL_NS))?;
    let before = tag[..at].strip_suffix('=')?;
    let declared = &before[before.rfind("xmlns:")? + "xmlns:".len()..];
    Some(declared.to_string())
}

/// Adds one to the attribute `attr` of the elements `local_name` whose value is at least `from`
fn shift_index_attr(xml: &str, local_name: &str, attr: &str, from: usize) -> Result<String, String> {
    let ranges = element_ranges(xml, local_name, |e| {
        get_attr_value(e, attr).and_then(|v| v.parse::<usize>().ok()).is_some_and(|v| v >= from)
    })?;
    let mut out = xml.to_string();
    for range in ranges.into_iter().rev() {
        let element = &xml[range.clone()];
        let value = start_tag_attr(element, attr).and_then(|v| v.parse::<usize>().ok()).unwrap_or(from);
        out.replace_range(range, &set_attr(element, attr, &(value + 1).to_string()));
    }
    Ok(out)
}

/// `xml` with the text of the elements named in `names` passed through `transform`, and how many
/// texts changed. Those elements hold only text.
pub(crate) fn rewrite_texts(xml: &str, names: &[&str], transform: &dyn Fn(&str) -> String) -> Result<(String, usize), String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut contents = Vec::new();
    let mut open: Option<usize> = None;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|e| format!("XML no válido: {}", e))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(e) if names.iter().any(|n| tag_is(e.name().as_ref(), n)) => open = Some(after),
            Event::End(e) if names.iter().any(|n| tag_is(e.name().as_ref(), n)) => {
                if let Some(start) = open.take() {
                    contents.push(start..before);
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    let mut out = xml.to_string();
    let mut changed = 0;
    for range in contents.into_iter().rev() {
        let text = unescape_text(&xml[range.clone()]);
        let renamed = transform(&text);
        if renamed != text {
            out.replace_range(range, &escape_xml(&renamed));
            changed += 1;
        }
    }
    Ok((out, changed))
}

/// `formula` with the references to the sheet `old` (`Hoja1!A1`, `'Hoja 1'!A1`, the ends of
/// `Hoja1:Hoja3!A1`) pointing to `new`, compared ignoring case. Text between double quotes and
/// references to other workbooks (`[1]Hoja1!A1`) are not touched.
pub fn rename_in_formula(formula: &str, old: &str, new: &str) -> String {
    let chars: Vec<char> = formula.chars().collect();
    let old = old.to_lowercase();
    let replacement = quote_sheet(new);
    let mut out = String::with_capacity(formula.len());
    let mut i = 0;
    // A sheet name may stand at position `i`: after an operator, a parenthesis, a comma, a space…
    let at_boundary = |i: usize| i == 0 || !(chars[i - 1].is_alphanumeric() || "_.]'\\$".contains(chars[i - 1]));
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            // A text: copied as is, with its doubled quotes
            out.push(c);
            i += 1;
            while i < chars.len() {
                out.push(chars[i]);
                if chars[i] == '"' {
                    if chars.get(i + 1) == Some(&'"') {
                        out.push('"');
                        i += 2;
                        continue;
                    }
                    i += 1;
                    break;
                }
                i += 1;
            }
            continue;
        }
        if c == '\'' && at_boundary(i) {
            // A quoted sheet name: 'Hoja 1'!A1 (or the ends of a quoted 3D reference, left alone)
            let mut j = i + 1;
            let mut name = String::new();
            while j < chars.len() {
                if chars[j] == '\'' {
                    if chars.get(j + 1) == Some(&'\'') {
                        name.push('\'');
                        j += 2;
                        continue;
                    }
                    break;
                }
                name.push(chars[j]);
                j += 1;
            }
            let raw: String = chars[i..(j + 1).min(chars.len())].iter().collect();
            if chars.get(j + 1) == Some(&'!') && name.to_lowercase() == old {
                out.push_str(&replacement);
            } else {
                out.push_str(&raw);
            }
            i = j + 1;
            continue;
        }
        if (c.is_alphabetic() || c == '_') && at_boundary(i) {
            let mut j = i;
            while j < chars.len() && (chars[j].is_alphanumeric() || chars[j] == '_' || chars[j] == '.') {
                j += 1;
            }
            let word: String = chars[i..j].iter().collect();
            let next = chars.get(j);
            // Hoja1!A1, or Hoja1:Hoja3!A1 (the first end is followed by ':' and then the second end)
            let is_sheet = next == Some(&'!') || (next == Some(&':') && second_end_ahead(&chars[j + 1..]));
            if is_sheet && word.to_lowercase() == old {
                out.push_str(&replacement);
            } else {
                out.push_str(&word);
            }
            i = j;
            continue;
        }
        out.push(c);
        i += 1;
    }
    out
}

/// Whether `rest` (after a `:`) is a sheet name followed by `!`, the second end of a 3D reference
fn second_end_ahead(rest: &[char]) -> bool {
    let len = rest.iter().take_while(|c| c.is_alphanumeric() || **c == '_' || **c == '.').count();
    len > 0 && rest.get(len) == Some(&'!')
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::workbook::CellInput;

    fn blank() -> Workbook {
        Workbook::from_bytes(&crate::blank::blank_xlsx().unwrap()).unwrap()
    }

    #[test]
    fn validates_names_like_excel() {
        let wb = blank();
        let sheets = wb.sheets().unwrap();
        for bad in ["", "  ", "a/b", "[x]", "dos:puntos", "'Ventas", "Ventas'", "history", &"x".repeat(32), "hoja1"] {
            assert!(validate_sheet_name(bad, &sheets, None).is_err(), "{:?}", bad);
        }
        for good in ["Ventas", "Resumen 2026", "Año fiscal", &"x".repeat(31), "Ven'tas"] {
            assert!(validate_sheet_name(good, &sheets, None).is_ok(), "{:?}", good);
        }
        // Its own name (other case) when renaming
        assert!(validate_sheet_name("HOJA1", &sheets, Some(0)).is_ok());
    }

    #[test]
    fn adds_sheets_at_the_end_or_after_one() {
        let mut wb = blank();
        assert_eq!(wb.add_sheet("Resumen", None).unwrap(), 1);
        assert_eq!(wb.add_sheet("Datos", Some(0)).unwrap(), 1);
        assert_eq!(wb.sheet_names().unwrap(), ["Hoja1", "Datos", "Resumen"]);
        assert!(wb.add_sheet("resumen", None).is_err());
        assert!(wb.add_sheet("Otra", Some(7)).is_err());

        // The new sheets take cells and survive a round trip
        wb.set_values(1, "A1", vec![vec![Some(CellInput::Number(5.0))]]).unwrap();
        wb.set_values(0, "A1", vec![vec![Some(CellInput::Formula("Datos!A1*2".into()))]]).unwrap();
        let reopened = Workbook::from_bytes(&wb.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.sheet_names().unwrap(), ["Hoja1", "Datos", "Resumen"]);
        assert_eq!(reopened.cell(0, "A1").unwrap().value, crate::workbook::CellValue::Number(10.0));
        let types = reopened.package.get_string("[Content_Types].xml").unwrap();
        assert!(types.contains("/xl/worksheets/sheet2.xml") && types.contains("/xl/worksheets/sheet3.xml"));
        let xml = reopened.workbook_xml().unwrap();
        assert!(xml.contains(r#"sheetId="2""#) && xml.contains(r#"sheetId="3""#));
    }

    #[test]
    fn adding_in_the_middle_shifts_sheet_positions() {
        let mut wb = blank();
        wb.add_sheet("B", None).unwrap();
        let xml = wb.workbook_xml().unwrap().replace(
            "</sheets>",
            r#"</sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="1">B!$A$1:$C$3</definedName><definedName name="X" localSheetId="0">Hoja1!$A$1</definedName></definedNames>"#,
        );
        let xml = xml.replacen("<sheets>", r#"<bookViews><workbookView activeTab="1"/></bookViews><sheets>"#, 1);
        wb.package.put(wb.workbook_part.clone(), xml.into_bytes());

        wb.add_sheet("Nueva", Some(0)).unwrap();
        let xml = wb.workbook_xml().unwrap();
        assert!(xml.contains(r#"localSheetId="2">B!"#), "{}", xml);
        assert!(xml.contains(r#"localSheetId="0">Hoja1!"#));
        assert!(xml.contains(r#"activeTab="2""#));
        assert_eq!(wb.sheet_names().unwrap(), ["Hoja1", "Nueva", "B"]);
    }

    #[test]
    fn add_sheet_undoes_in_one_step() {
        let mut wb = blank();
        wb.checkpoint(None);
        wb.add_sheet("Resumen", None).unwrap();
        wb.undo().unwrap();
        assert_eq!(wb.sheet_names().unwrap(), ["Hoja1"]);
        assert!(!wb.package.contains("xl/worksheets/sheet2.xml"));
        wb.redo().unwrap();
        assert_eq!(wb.sheet_names().unwrap(), ["Hoja1", "Resumen"]);
    }

    #[test]
    fn renames_references_in_formulas() {
        let cases = [
            ("Hoja1!A1+1", "Ventas!A1+1"),
            ("SUM('hoja1'!A1:B2)", "SUM(Ventas!A1:B2)"),
            ("Hoja1:Hoja3!A1", "Ventas:Hoja3!A1"),
            ("Hoja3:Hoja1!A1", "Hoja3:Ventas!A1"),
            (r#""Hoja1!A1"&Hoja1!B1"#, r#""Hoja1!A1"&Ventas!B1"#),
            ("[1]Hoja1!A1", "[1]Hoja1!A1"),
            ("Hoja10!A1+MiHoja1!A1", "Hoja10!A1+MiHoja1!A1"),
            ("Hoja1", "Hoja1"),
        ];
        for (formula, expected) in cases {
            assert_eq!(rename_in_formula(formula, "Hoja1", "Ventas"), expected, "{}", formula);
        }
        assert_eq!(rename_in_formula("Hoja1!A1", "Hoja1", "Ventas 2026"), "'Ventas 2026'!A1");
        assert_eq!(rename_in_formula("'Ventas 2026'!A1", "Ventas 2026", "Año"), "Año!A1");
        assert_eq!(rename_in_formula("Hoja1!A1", "Hoja1", "O'Brien"), "'O''Brien'!A1");
    }

    #[test]
    fn rename_updates_formulas_names_and_charts() {
        let mut wb = blank();
        wb.add_sheet("Datos", None).unwrap();
        wb.set_values(1, "A1", vec![vec![Some(CellInput::Number(4.0))]]).unwrap();
        wb.set_values(0, "A1", vec![vec![Some(CellInput::Formula("Datos!A1*2".into()))]]).unwrap();
        let xml = wb.workbook_xml().unwrap().replace("</sheets>", r#"</sheets><definedNames><definedName name="Tasa">Datos!$A$1</definedName></definedNames>"#);
        wb.package.put(wb.workbook_part.clone(), xml.into_bytes());
        wb.package.put("xl/charts/chart1.xml".into(), br#"<c:chartSpace xmlns:c="c"><c:f>Datos!$A$1:$A$3</c:f><c:f>Otra!$B$1</c:f></c:chartSpace>"#.to_vec());

        wb.checkpoint(None);
        assert_eq!(wb.rename_sheet(1, "Precios & más").unwrap(), 3);
        assert_eq!(wb.sheet_names().unwrap(), ["Hoja1", "Precios & más"]);
        assert_eq!(wb.cell(0, "A1").unwrap().formula.as_deref(), Some("'Precios & más'!A1*2"));
        assert!(wb.workbook_xml().unwrap().contains("<definedName name=\"Tasa\">&apos;Precios &amp; más&apos;!$A$1</definedName>"));
        let chart = wb.package.get_string("xl/charts/chart1.xml").unwrap();
        assert!(chart.contains("<c:f>&apos;Precios &amp; más&apos;!$A$1:$A$3</c:f><c:f>Otra!$B$1</c:f>"), "{}", chart);
        // The formula still calculates with the new name
        wb.set_values(1, "A1", vec![vec![Some(CellInput::Number(6.0))]]).unwrap();
        assert_eq!(wb.cell(0, "A1").unwrap().value, crate::workbook::CellValue::Number(12.0));

        assert!(wb.rename_sheet(1, "hoja1").is_err());
        assert_eq!(wb.rename_sheet(1, "Precios & más").unwrap(), 0);
        let reopened = Workbook::from_bytes(&wb.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.sheet_names().unwrap(), ["Hoja1", "Precios & más"]);
    }
}
