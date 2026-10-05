pub mod docx_parser;
pub mod layout_engine;
pub mod paragraph_edit;
pub mod sample_generator;
pub mod styles;

use docx_parser::{DocxModifier, KeyValuePair, ParagraphUpdate};
use layout_engine::LayoutEngine;
use wasm_bindgen::prelude::*;

#[wasm_bindgen(start)]
pub fn main_js() -> Result<(), JsValue> {
    console_error_panic_hook::set_once();
    Ok(())
}

/// Generates a sample contract DOCX file in memory with tables, alignments and styles
#[wasm_bindgen]
pub fn generate_sample_docx_wasm() -> Result<js_sys::Uint8Array, JsValue> {
    let bytes = sample_generator::generate_sample_docx()
        .map_err(|e| JsValue::from_str(&e))?;
    Ok(js_sys::Uint8Array::from(&bytes[..]))
}

/// In-memory DOCX Session for interactive web editing
#[wasm_bindgen]
pub struct DocxSession {
    modifier: DocxModifier,
}

#[wasm_bindgen]
impl DocxSession {
    /// Creates a new DocxSession from byte slice
    #[wasm_bindgen(constructor)]
    pub fn new(bytes: &[u8]) -> Result<DocxSession, JsValue> {
        let modifier = DocxModifier::from_bytes(bytes).map_err(|e| JsValue::from_str(&e))?;
        Ok(DocxSession { modifier })
    }

    /// Creates a session with sample contract docx
    #[wasm_bindgen]
    pub fn new_sample() -> Result<DocxSession, JsValue> {
        let bytes = sample_generator::generate_sample_docx().map_err(|e| JsValue::from_str(&e))?;
        let modifier = DocxModifier::from_bytes(&bytes).map_err(|e| JsValue::from_str(&e))?;
        Ok(DocxSession { modifier })
    }

    /// Computes multi-page layout and render commands for high-performance Canvas rendering
    #[wasm_bindgen]
    pub fn compute_canvas_layout_json(
        &self,
        watermark_text: Option<String>,
        watermark_opacity: f64,
    ) -> Result<String, JsValue> {
        let elements = self.modifier.extract_elements().map_err(|e| JsValue::from_str(&e))?;
        let stats = self.modifier.get_statistics().map_err(|e| JsValue::from_str(&e))?;
        let engine = LayoutEngine::new();
        let layout = engine.compute_layout(
            &elements,
            &stats.background_color,
            &stats.page_setup,
            &stats.header_footer,
            stats.bg_image_data_url.as_deref(),
            watermark_text.as_deref(),
            watermark_opacity,
        );
        serde_json::to_string(&layout).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Gets all document elements (paragraphs and tables in order) as JSON
    #[wasm_bindgen]
    pub fn get_document_elements_json(&self) -> Result<String, JsValue> {
        let elements = self.modifier.extract_elements().map_err(|e| JsValue::from_str(&e))?;
        serde_json::to_string(&elements).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Gets paragraphs list in JSON format
    #[wasm_bindgen]
    pub fn get_paragraphs_json(&self) -> Result<String, JsValue> {
        let paragraphs = self.modifier.extract_paragraphs().map_err(|e| JsValue::from_str(&e))?;
        serde_json::to_string(&paragraphs).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Gets tables list in JSON format
    #[wasm_bindgen]
    pub fn get_tables_json(&self) -> Result<String, JsValue> {
        let tables = self.modifier.extract_tables().map_err(|e| JsValue::from_str(&e))?;
        serde_json::to_string(&tables).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Gets full document plain text
    #[wasm_bindgen]
    pub fn get_raw_text(&self) -> Result<String, JsValue> {
        self.modifier.extract_raw_text().map_err(|e| JsValue::from_str(&e))
    }

    /// Gets statistics in JSON format
    #[wasm_bindgen]
    pub fn get_stats_json(&self) -> Result<String, JsValue> {
        let stats = self.modifier.get_statistics().map_err(|e| JsValue::from_str(&e))?;
        serde_json::to_string(&stats).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Performs search and replace
    #[wasm_bindgen]
    pub fn find_and_replace(
        &mut self,
        search: &str,
        replacement: &str,
        match_case: bool,
        use_regex: bool,
    ) -> Result<String, JsValue> {
        let res = self
            .modifier
            .find_and_replace(search, replacement, match_case, use_regex)
            .map_err(|e| JsValue::from_str(&e))?;
        serde_json::to_string(&res).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Performs batch replacement with a JSON array of [{ "key": "...", "value": "..." }]
    #[wasm_bindgen]
    pub fn batch_replace(&mut self, pairs_json: &str) -> Result<String, JsValue> {
        let pairs: Vec<KeyValuePair> = serde_json::from_str(pairs_json)
            .map_err(|e| JsValue::from_str(&format!("JSON inválido para batch replace: {}", e)))?;
        let res = self.modifier.batch_replace(&pairs).map_err(|e| JsValue::from_str(&e))?;
        serde_json::to_string(&res).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Updates a single paragraph text by index
    #[wasm_bindgen]
    pub fn update_paragraph(&mut self, index: usize, new_text: &str) -> Result<bool, JsValue> {
        let update = vec![ParagraphUpdate {
            index,
            text: new_text.to_string(),
        }];
        let count = self.modifier.update_paragraphs(&update).map_err(|e| JsValue::from_str(&e))?;
        Ok(count > 0)
    }

    /// Updates paragraph with individual styled text runs (bold, italic, color, underline per word/segment)
    #[wasm_bindgen]
    pub fn update_paragraph_runs(
        &mut self,
        index: usize,
        runs_json: &str,
        align: Option<String>,
    ) -> Result<bool, JsValue> {
        let runs: Vec<crate::docx_parser::RunInfo> = serde_json::from_str(runs_json)
            .map_err(|e| JsValue::from_str(&format!("JSON deserialization error: {}", e)))?;
        self.modifier
            .update_paragraph_runs(index, &runs, align.as_deref())
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Updates paragraph with full rich formatting (alignment, color, bold, italic, text)
    #[wasm_bindgen]
    pub fn update_paragraph_rich(
        &mut self,
        index: usize,
        text: &str,
        align: &str,
        color: &str,
        bold: bool,
        italic: bool,
    ) -> Result<bool, JsValue> {
        self.modifier
            .update_paragraph_rich(index, text, align, color, bold, italic)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Updates a table cell by table_index, row, col
    #[wasm_bindgen]
    pub fn update_table_cell(
        &mut self,
        table_index: usize,
        row: usize,
        col: usize,
        new_text: &str,
    ) -> Result<bool, JsValue> {
        self.modifier
            .update_table_cell(table_index, row, col, new_text)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Inserts a new table into document
    #[wasm_bindgen]
    pub fn add_table(&mut self, rows: usize, cols: usize, headers_json: &str) -> Result<bool, JsValue> {
        let headers: Vec<String> = serde_json::from_str(headers_json)
            .unwrap_or_else(|_| (1..=cols).map(|i| format!("Columna {}", i)).collect());
        self.modifier
            .add_table(rows, cols, &headers)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Sets page background color (HEX without '#')
    #[wasm_bindgen]
    pub fn set_background_color(&mut self, hex_color: &str) -> Result<(), JsValue> {
        self.modifier
            .set_background_color(hex_color)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Sets background / watermark image from bytes
    #[wasm_bindgen]
    pub fn set_background_image(&mut self, bytes: &[u8], ext: &str) -> Result<(), JsValue> {
        self.modifier
            .set_background_image(bytes.to_vec(), ext)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Exports the modified docx file as Uint8Array
    #[wasm_bindgen]
    pub fn export_bytes(&self) -> Result<js_sys::Uint8Array, JsValue> {
        let bytes = self.modifier.to_bytes().map_err(|e| JsValue::from_str(&e))?;
        Ok(js_sys::Uint8Array::from(&bytes[..]))
    }
}
