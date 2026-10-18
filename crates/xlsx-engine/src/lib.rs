//! Excel (.xlsx) engine compiled to WebAssembly for @dimarborda/xlsx. The package layer (ZIP
//! with limits, lossless writing) comes from `ooxml-core`, shared with the .docx engine.

pub mod blank;
pub mod workbook;

use wasm_bindgen::prelude::*;
use workbook::Workbook;

#[wasm_bindgen(start)]
pub fn main_js() {
    console_error_panic_hook::set_once();
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
        let workbook = Workbook::from_bytes(bytes).map_err(|e| JsValue::from_str(&e))?;
        Ok(XlsxWorkbook { workbook })
    }

    /// An empty workbook with one sheet
    pub fn new_blank() -> Result<XlsxWorkbook, JsValue> {
        let bytes = blank::blank_xlsx().map_err(|e| JsValue::from_str(&e))?;
        XlsxWorkbook::new(&bytes)
    }

    /// Sheet names, in tab order
    pub fn sheet_names(&self) -> Result<Vec<String>, JsValue> {
        self.workbook.sheet_names().map_err(|e| JsValue::from_str(&e))
    }

    /// The workbook as .xlsx bytes
    pub fn export_bytes(&self) -> Result<Vec<u8>, JsValue> {
        self.workbook.to_bytes().map_err(|e| JsValue::from_str(&e))
    }
}
