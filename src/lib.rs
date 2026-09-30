pub mod docx_parser;
pub mod sample_generator;

use docx_parser::{DocxModifier, KeyValuePair, ParagraphUpdate};
use wasm_bindgen::prelude::*;

#[wasm_bindgen(start)]
pub fn main_js() -> Result<(), JsValue> {
    console_error_panic_hook::set_once();
    Ok(())
}

/// Generates a sample contract DOCX file in memory
#[wasm_bindgen]
pub fn generate_sample_docx_wasm() -> Result<js_sys::Uint8Array, JsValue> {
    let bytes = sample_generator::generate_sample_docx()
        .map_err(|e| JsValue::from_str(&e))?;
    Ok(js_sys::Uint8Array::from(&bytes[..]))
}

/// Extracts paragraphs from a docx file as a JSON string
#[wasm_bindgen]
pub fn extract_docx_paragraphs_json(bytes: &[u8]) -> Result<String, JsValue> {
    let modifier = DocxModifier::from_bytes(bytes).map_err(|e| JsValue::from_str(&e))?;
    let paragraphs = modifier.extract_paragraphs().map_err(|e| JsValue::from_str(&e))?;
    serde_json::to_string(&paragraphs).map_err(|e| JsValue::from_str(&e.to_string()))
}

/// Extracts full document plain text
#[wasm_bindgen]
pub fn extract_docx_raw_text(bytes: &[u8]) -> Result<String, JsValue> {
    let modifier = DocxModifier::from_bytes(bytes).map_err(|e| JsValue::from_str(&e))?;
    modifier.extract_raw_text().map_err(|e| JsValue::from_str(&e))
}

/// Returns statistics about the DOCX
#[wasm_bindgen]
pub fn get_docx_statistics_json(bytes: &[u8]) -> Result<String, JsValue> {
    let modifier = DocxModifier::from_bytes(bytes).map_err(|e| JsValue::from_str(&e))?;
    let stats = modifier.get_statistics().map_err(|e| JsValue::from_str(&e))?;
    serde_json::to_string(&stats).map_err(|e| JsValue::from_str(&e.to_string()))
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

    /// Gets paragraphs list in JSON format
    #[wasm_bindgen]
    pub fn get_paragraphs_json(&self) -> Result<String, JsValue> {
        let paragraphs = self.modifier.extract_paragraphs().map_err(|e| JsValue::from_str(&e))?;
        serde_json::to_string(&paragraphs).map_err(|e| JsValue::from_str(&e.to_string()))
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

    /// Updates a single paragraph by index
    #[wasm_bindgen]
    pub fn update_paragraph(&mut self, index: usize, new_text: &str) -> Result<bool, JsValue> {
        let update = vec![ParagraphUpdate {
            index,
            text: new_text.to_string(),
        }];
        let count = self.modifier.update_paragraphs(&update).map_err(|e| JsValue::from_str(&e))?;
        Ok(count > 0)
    }

    /// Updates multiple paragraphs with a JSON array of [{ "index": 0, "text": "..." }]
    #[wasm_bindgen]
    pub fn update_paragraphs_batch(&mut self, updates_json: &str) -> Result<usize, JsValue> {
        let updates: Vec<ParagraphUpdate> = serde_json::from_str(updates_json)
            .map_err(|e| JsValue::from_str(&format!("JSON inválido para updates: {}", e)))?;
        let count = self.modifier.update_paragraphs(&updates).map_err(|e| JsValue::from_str(&e))?;
        Ok(count)
    }

    /// Exports the modified docx file as Uint8Array
    #[wasm_bindgen]
    pub fn export_bytes(&self) -> Result<js_sys::Uint8Array, JsValue> {
        let bytes = self.modifier.to_bytes().map_err(|e| JsValue::from_str(&e))?;
        Ok(js_sys::Uint8Array::from(&bytes[..]))
    }
}
