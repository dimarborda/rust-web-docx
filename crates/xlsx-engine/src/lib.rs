//! Excel (.xlsx) engine compiled to WebAssembly for @dimarborda/xlsx. The package layer (ZIP
//! with limits, lossless writing) comes from `ooxml-core`, shared with the .docx engine.
//! Cells cross the boundary as JSON; `packages/xlsx/src/index.js` turns them into JS values.

pub mod blank;
pub mod cell_ref;
pub mod dates;
pub mod shared_strings;
pub mod sheet;
pub mod styles;
pub mod numfmt;
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

/// One value to write: a string, number or boolean, `null` (clear), `{"date": ms}` or
/// `{"skip": true}` (leave the cell as it is)
fn cell_input(value: &Value) -> Result<Option<CellInput>, String> {
    Ok(Some(match value {
        Value::Null => CellInput::Clear,
        Value::String(s) => CellInput::Text(s.clone()),
        Value::Bool(b) => CellInput::Bool(*b),
        Value::Number(n) => CellInput::Number(n.as_f64().ok_or("Número no válido.")?),
        Value::Object(o) if o.get("skip") == Some(&Value::Bool(true)) => return Ok(None),
        Value::Object(o) => match o.get("date").and_then(Value::as_f64) {
            Some(ms) => CellInput::Date(ms),
            None => return Err("Valor de celda no admitido.".to_string()),
        },
        Value::Array(_) => return Err("Valor de celda no admitido: una lista dentro de una celda.".to_string()),
    }))
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

    /// Replaces `{{FIELD}}` placeholders with the strings of a JSON object; returns how many
    pub fn fill_template_json(&mut self, values: &str) -> Result<usize, JsValue> {
        let values: HashMap<String, String> =
            serde_json::from_str(values).map_err(|e| js_err(format!("Los valores deben ser un objeto de textos: {}", e)))?;
        self.workbook.fill_template(&values).map_err(js_err)
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
    }
}
