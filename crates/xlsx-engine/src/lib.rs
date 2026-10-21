//! Excel (.xlsx) engine compiled to WebAssembly for @dimarborda/xlsx. The package layer (ZIP
//! with limits, lossless writing) comes from `ooxml-core`, shared with the .docx engine.
//! Cells cross the boundary as JSON; `packages/xlsx/src/index.js` turns them into JS values.

pub mod areas;
pub mod blank;
pub mod calc;
pub mod format;
pub mod formula;
pub mod cell_ref;
pub mod dates;
pub mod shared_strings;
pub mod sheet;
pub mod styles;
pub mod numfmt;
pub mod recalc;
pub mod text;
pub mod theme;
pub mod view;
pub mod workbook;

use std::collections::HashMap;

use serde_json::{json, Value};
use wasm_bindgen::prelude::*;
use workbook::{CellData, CellInput, CellValue, SheetKind, Workbook};

#[wasm_bindgen(start)]
pub fn main_js() {
    console_error_panic_hook::set_once();
}

fn js_err(e: String) -> JsValue {
    JsValue::from_str(&e)
}

/// `{ref, type, value, formula?}`; dates as milliseconds since 1970-01-01 UTC
fn cell_json(cell: &CellData) -> Value {
    let (kind, value) = match &cell.value {
        CellValue::Empty => ("empty", Value::Null),
        CellValue::Number(n) => ("number", json!(n)),
        CellValue::Text(s) => ("string", json!(s)),
        CellValue::Bool(b) => ("boolean", json!(b)),
        CellValue::Error(e) => ("error", json!(e)),
        CellValue::Date(ms) => ("date", json!(ms)),
    };
    let mut out = json!({ "ref": cell.at.to_string(), "type": kind, "value": value });
    if let Some(formula) = &cell.formula {
        out["formula"] = json!(formula);
    }
    out
}

/// One value to write: a string, number or boolean, `null` (clear), `{"date": ms}`,
/// `{"formula": "SUM(A1:A3)"}` or
/// `{"skip": true}` (leave the cell as it is)
fn cell_input(value: &Value) -> Result<Option<CellInput>, String> {
    Ok(Some(match value {
        Value::Null => CellInput::Clear,
        Value::String(s) => CellInput::Text(s.clone()),
        Value::Bool(b) => CellInput::Bool(*b),
        Value::Number(n) => CellInput::Number(n.as_f64().ok_or("Número no válido.")?),
        Value::Object(o) if o.get("skip") == Some(&Value::Bool(true)) => return Ok(None),
        Value::Object(o) => match (o.get("date").and_then(Value::as_f64), o.get("formula").and_then(Value::as_str)) {
            (Some(ms), _) => CellInput::Date(ms),
            (None, Some(f)) => CellInput::Formula(f.to_string()),
            _ => return Err("Valor de celda no admitido.".to_string()),
        },
        Value::Array(_) => return Err("Valor de celda no admitido: una lista dentro de una celda.".to_string()),
    }))
}

/// A format from JSON: `{bold, italic, underline, strike, color, fontSize, fontName, fill,
/// numberFormat, align, valign, wrap, border}`; `border` is a style (`"thin"`), `null` to remove
/// borders, or `{style, color, edges}`
fn format_input(value: &Value) -> Result<format::Format, String> {
    let obj = value.as_object().ok_or("El formato debe ser un objeto.")?;
    let known = ["bold", "italic", "underline", "strike", "color", "fontSize", "fontName", "fill", "numberFormat", "align", "valign", "wrap", "border"];
    if let Some(unknown) = obj.keys().find(|k| !known.contains(&k.as_str())) {
        return Err(format!("Propiedad de formato desconocida: '{}'. Usa {}.", unknown, known.join(", ")));
    }
    let flag = |k: &str| -> Result<Option<bool>, String> {
        match obj.get(k) {
            None | Some(Value::Null) => Ok(None),
            Some(Value::Bool(b)) => Ok(Some(*b)),
            _ => Err(format!("'{}' debe ser true o false.", k)),
        }
    };
    let text = |k: &str| -> Result<Option<String>, String> {
        match obj.get(k) {
            None | Some(Value::Null) => Ok(None),
            Some(Value::String(s)) => Ok(Some(s.clone())),
            _ => Err(format!("'{}' debe ser un texto.", k)),
        }
    };
    let fill = match obj.get("fill") {
        None => None,
        Some(Value::Null) => Some(None),
        Some(Value::String(s)) => Some(Some(s.clone())),
        _ => return Err("'fill' debe ser un color (#RRGGBB) o null.".to_string()),
    };
    let border = match obj.get("border") {
        None => None,
        Some(Value::Null) => Some(format::BorderSpec { style: None, color: None, edges: format::Edges::All }),
        Some(Value::String(s)) => Some(format::BorderSpec { style: Some(s.clone()), color: None, edges: format::Edges::All }),
        Some(Value::Object(b)) => Some(format::BorderSpec {
            style: b.get("style").and_then(Value::as_str).map(str::to_string),
            color: b.get("color").and_then(Value::as_str).map(str::to_string),
            edges: format::Edges::parse(b.get("edges").and_then(Value::as_str).unwrap_or("all"))?,
        }),
        _ => return Err("'border' debe ser un estilo (\"thin\"), null o {style, color, edges}.".to_string()),
    };
    Ok(format::Format {
        bold: flag("bold")?,
        italic: flag("italic")?,
        underline: flag("underline")?,
        strike: flag("strike")?,
        font_color: text("color")?,
        font_size: match obj.get("fontSize") {
            None | Some(Value::Null) => None,
            Some(v) => Some(v.as_f64().ok_or("'fontSize' debe ser un número.")?),
        },
        font_name: text("fontName")?,
        fill,
        number_format: text("numberFormat")?,
        horizontal: text("align")?,
        vertical: text("valign")?,
        wrap: flag("wrap")?,
        border,
    })
}

fn undone_json(done: Option<ooxml_core::Undone>) -> String {
    match done {
        Some(d) => json!({ "selection": d.selection, "parts": d.parts }).to_string(),
        None => "null".to_string(),
    }
}

/// An open workbook
#[wasm_bindgen]
pub struct XlsxWorkbook {
    workbook: Workbook,
}

#[wasm_bindgen]
impl XlsxWorkbook {
    /// Opens an .xlsx from its bytes
    #[wasm_bindgen(constructor)]
    pub fn new(bytes: &[u8]) -> Result<XlsxWorkbook, JsValue> {
        let workbook = Workbook::from_bytes(bytes).map_err(js_err)?;
        Ok(XlsxWorkbook { workbook })
    }

    /// An empty workbook with one sheet
    pub fn new_blank() -> Result<XlsxWorkbook, JsValue> {
        let bytes = blank::blank_xlsx().map_err(js_err)?;
        XlsxWorkbook::new(&bytes)
    }

    /// Sheet names, in tab order
    pub fn sheet_names(&self) -> Result<Vec<String>, JsValue> {
        self.workbook.sheet_names().map_err(js_err)
    }

    /// The sheets as JSON: `[{name, kind, state}]`, kind `worksheet`, `chartsheet` or `other`
    pub fn sheets_json(&self) -> Result<String, JsValue> {
        let sheets = self.workbook.sheets().map_err(js_err)?;
        let list: Vec<Value> = sheets
            .iter()
            .map(|s| {
                let kind = match s.kind {
                    SheetKind::Worksheet => "worksheet",
                    SheetKind::Chartsheet => "chartsheet",
                    SheetKind::Other => "other",
                };
                json!({ "name": s.name, "kind": kind, "state": s.state.as_deref().unwrap_or("visible") })
            })
            .collect();
        Ok(Value::Array(list).to_string())
    }

    /// One cell of the sheet at `sheet`, as JSON (see `cell_json`)
    pub fn cell_json(&self, sheet: usize, reference: &str) -> Result<String, JsValue> {
        Ok(cell_json(&self.workbook.cell(sheet, reference).map_err(js_err)?).to_string())
    }

    /// The cells of a range (`A1:C3`) as JSON rows
    pub fn range_json(&self, sheet: usize, range: &str) -> Result<String, JsValue> {
        let rows = self.workbook.range(sheet, range).map_err(js_err)?;
        let rows: Vec<Value> = rows.iter().map(|row| Value::Array(row.iter().map(cell_json).collect())).collect();
        Ok(Value::Array(rows).to_string())
    }

    /// Writes JSON rows of values (see `cell_input`) from the cell `start`
    pub fn set_values_json(&mut self, sheet: usize, start: &str, rows: &str) -> Result<(), JsValue> {
        let parsed: Value = serde_json::from_str(rows).map_err(|e| js_err(format!("Valores no válidos: {}", e)))?;
        let rows = parsed.as_array().ok_or_else(|| js_err("Los valores deben ser una lista de filas.".to_string()))?;
        let mut inputs = Vec::with_capacity(rows.len());
        for row in rows {
            let row = row.as_array().ok_or_else(|| js_err("Cada fila debe ser una lista de valores.".to_string()))?;
            inputs.push(row.iter().map(cell_input).collect::<Result<Vec<_>, _>>().map_err(js_err)?);
        }
        self.workbook.set_values(sheet, start, inputs).map_err(js_err)
    }

    /// Areas of the workbook (tables, defined names, print areas, filters, regions) as JSON;
    /// only those of the sheet at `sheet` when given
    pub fn areas_json(&self, sheet: Option<usize>) -> Result<String, JsValue> {
        let names = self.workbook.sheet_names().map_err(js_err)?;
        let areas = self.workbook.areas().map_err(js_err)?;
        let list: Vec<Value> = areas
            .iter()
            .filter(|a| sheet.is_none_or(|s| a.sheet == s))
            .map(|a| a.to_json(names.get(a.sheet).map(String::as_str).unwrap_or("")))
            .collect();
        Ok(Value::Array(list).to_string())
    }

    /// Context of an area for an agent (see `Workbook::area_context`): `target` is a table or
    /// defined name, an address with its sheet, or a range of the sheet at `sheet`
    pub fn area_context_json(&self, target: &str, sheet: Option<usize>, max_rows: usize, locale: &str) -> Result<String, JsValue> {
        let area = self.workbook.resolve_area(target, sheet).map_err(js_err)?;
        let context = self.workbook.area_context(&area, max_rows, &numfmt::Locale::from_tag(locale)).map_err(js_err)?;
        Ok(context.to_string())
    }

    /// The region around a cell (Excel's current region) as JSON
    pub fn region_at_json(&self, sheet: usize, reference: &str) -> Result<String, JsValue> {
        let names = self.workbook.sheet_names().map_err(js_err)?;
        let area = self.workbook.region_at(sheet, reference).map_err(js_err)?;
        Ok(area.to_json(names.get(sheet).map(String::as_str).unwrap_or("")).to_string())
    }

    /// Formats the cells of a range; `format` is JSON (see `format_input`)
    pub fn set_format_json(&mut self, sheet: usize, range: &str, format: &str) -> Result<(), JsValue> {
        let value: Value = serde_json::from_str(format).map_err(|e| js_err(format!("Formato no válido: {}", e)))?;
        let format = format_input(&value).map_err(js_err)?;
        self.workbook.set_format(sheet, range, &format).map_err(js_err)
    }

    /// Back to the default format for the cells of a range
    pub fn clear_format(&mut self, sheet: usize, range: &str) -> Result<(), JsValue> {
        self.workbook.clear_format(sheet, range).map_err(js_err)
    }

    /// Width in characters of the columns `first..=last` (1 = A)
    pub fn set_column_width(&mut self, sheet: usize, first: u32, last: u32, width: f64) -> Result<(), JsValue> {
        self.workbook.set_column_width(sheet, first, last, width).map_err(js_err)
    }

    pub fn merge_cells(&mut self, sheet: usize, range: &str) -> Result<(), JsValue> {
        self.workbook.merge_cells(sheet, range).map_err(js_err)
    }

    /// Splits merged cells overlapping a range; returns how many
    pub fn unmerge_cells(&mut self, sheet: usize, range: &str) -> Result<usize, JsValue> {
        self.workbook.unmerge_cells(sheet, range).map_err(js_err)
    }

    /// Recalculates the formulas of the workbook; returns how many cells got a new value
    pub fn recalculate(&mut self) -> Result<usize, JsValue> {
        self.workbook.recalculate().map_err(js_err)
    }

    /// Replaces `{{FIELD}}` placeholders with the strings of a JSON object; returns how many
    pub fn fill_template_json(&mut self, values: &str) -> Result<usize, JsValue> {
        let values: HashMap<String, String> =
            serde_json::from_str(values).map_err(|e| js_err(format!("Los valores deben ser un objeto de textos: {}", e)))?;
        self.workbook.fill_template(&values).map_err(js_err)
    }

    /// Markers of the shared lexicon (`{{KEY}}`, `{key}`, `[ … ]`, see `ooxml_core::markers`)
    /// in the text cells of every worksheet, as JSON:
    /// `[{sheet, sheetName, ref, row, col, kind, syntax, raw, key, hint, offset, length}]`
    pub fn markers_json(&self) -> Result<String, JsValue> {
        let sheets = self.workbook.sheets().map_err(js_err)?;
        let locale = numfmt::Locale::from_tag("en");
        let mut out = Vec::new();
        for (index, info) in sheets.iter().enumerate() {
            if info.kind != SheetKind::Worksheet {
                continue;
            }
            for cell in self.workbook.cell_views(index, &locale).map_err(js_err)? {
                if cell.kind != "s" {
                    continue;
                }
                for m in ooxml_core::markers::scan(&cell.text) {
                    let reference = format!("{}{}", cell_ref::column_name(cell.col), cell.row);
                    out.push(json!({
                        "sheet": index, "sheetName": info.name, "ref": reference, "row": cell.row, "col": cell.col,
                        "kind": m.kind.as_str(), "syntax": m.syntax.as_str(), "raw": m.raw,
                        "key": m.key, "hint": m.hint, "offset": m.start, "length": m.len,
                    }));
                }
            }
        }
        Ok(Value::Array(out).to_string())
    }

    /// Cell formats for drawing, as JSON: `{defaultFont, xfs: [{font, fill, border, h, v, …}]}`
    pub fn styles_view_json(&self) -> Result<String, JsValue> {
        Ok(self.workbook.style_view().map_err(js_err)?.to_json().to_string())
    }

    /// Grid of the sheet at `sheet` as JSON: column widths, row heights, merged cells, frozen panes
    pub fn sheet_layout_json(&self, sheet: usize) -> Result<String, JsValue> {
        Ok(self.workbook.sheet_layout(sheet).map_err(js_err)?.to_json().to_string())
    }

    /// Cells of the sheet at `sheet` that show something, formatted in the language `locale`
    /// (`es`, `en-US`…), as JSON rows `[row, col, style, text, kind, color?]`
    pub fn sheet_cells_json(&self, sheet: usize, locale: &str) -> Result<String, JsValue> {
        let cells = self.workbook.cell_views(sheet, &numfmt::Locale::from_tag(locale)).map_err(js_err)?;
        Ok(view::cells_json(&cells).to_string())
    }

    /// Starts an undo step; `selection` comes back from `undo` (any text, e.g. JSON)
    pub fn checkpoint(&mut self, selection: Option<String>) {
        self.workbook.checkpoint(selection);
    }

    /// Where the selection ended up after the current step (comes back from `redo`)
    pub fn set_selection_after(&mut self, selection: Option<String>) {
        self.workbook.set_selection_after(selection);
    }

    /// Undoes the last step: JSON `{selection, parts}`, or `null` when there is nothing to undo
    pub fn undo_json(&mut self) -> String {
        undone_json(self.workbook.undo())
    }

    /// Redoes the last undone step (see `undo_json`)
    pub fn redo_json(&mut self) -> String {
        undone_json(self.workbook.redo())
    }

    pub fn can_undo(&self) -> bool {
        self.workbook.can_undo()
    }

    pub fn can_redo(&self) -> bool {
        self.workbook.can_redo()
    }

    /// The workbook as .xlsx bytes
    pub fn export_bytes(&self) -> Result<Vec<u8>, JsValue> {
        self.workbook.to_bytes().map_err(js_err)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn converts_values_to_and_from_json() {
        let mut wb = XlsxWorkbook::new_blank().unwrap();
        wb.set_values_json(0, "A1", r#"[["a", 1.5, true, {"date": 0}, {"skip": true}, null]]"#).unwrap();
        let rows: Value = serde_json::from_str(&wb.range_json(0, "A1:E1").unwrap()).unwrap();
        assert_eq!(rows[0][0], json!({"ref": "A1", "type": "string", "value": "a"}));
        assert_eq!(rows[0][1]["value"], json!(1.5));
        assert_eq!(rows[0][2]["type"], "boolean");
        assert_eq!(rows[0][3], json!({"ref": "D1", "type": "date", "value": 0.0}));
        assert_eq!(rows[0][4], json!({"ref": "E1", "type": "empty", "value": null}));
        assert!(cell_input(&json!({"otra": 1})).is_err());
        assert!(cell_input(&json!([1])).is_err());
        let sheets: Value = serde_json::from_str(&wb.sheets_json().unwrap()).unwrap();
        assert_eq!(sheets, json!([{"name": "Hoja1", "kind": "worksheet", "state": "visible"}]));

        wb.checkpoint(Some("A1".into()));
        wb.set_values_json(0, "A1", r#"[["b"]]"#).unwrap();
        let undone: Value = serde_json::from_str(&wb.undo_json()).unwrap();
        assert_eq!(undone["selection"], "A1");
        assert!(wb.can_redo());
        wb.redo_json();
        assert_eq!(wb.redo_json(), "null");
    }
}
