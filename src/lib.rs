pub mod blank_generator;
pub mod caret;
pub mod docx_parser;
pub mod layout_engine;
pub mod paragraph_edit;
pub mod sample_generator;
pub mod styles;

use docx_parser::{DocxModifier, KeyValuePair, ParagraphUpdate};
use caret::TextPosition;
use std::collections::HashMap;
use std::hash::{Hash, Hasher};
use layout_engine::{CachedMeasurer, DocumentLayout, EstimateMeasurer, FontSpec, LayoutEngine, TextMeasurer};
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

/// Measures text through a JS callback `(text, family, sizePx, bold, italic) => width`
struct JsMeasurer {
    func: js_sys::Function,
}

impl TextMeasurer for JsMeasurer {
    fn measure(&mut self, text: &str, font: &FontSpec) -> f64 {
        let args = js_sys::Array::new();
        args.push(&JsValue::from_str(text));
        args.push(&JsValue::from_str(font.family));
        args.push(&JsValue::from_f64(font.size));
        args.push(&JsValue::from_bool(font.bold));
        args.push(&JsValue::from_bool(font.italic));
        self.func
            .apply(&JsValue::NULL, &args)
            .ok()
            .and_then(|w| w.as_f64())
            .unwrap_or_else(|| EstimateMeasurer.measure(text, font))
    }
}

/// In-memory DOCX Session for interactive web editing
#[wasm_bindgen]
pub struct DocxSession {
    modifier: DocxModifier,
    /// Last computed layout, used for caret geometry
    layout: Option<DocumentLayout>,
    /// Text widths measured by the frontend, kept across layouts (cleared when fonts change)
    measure_cache: HashMap<String, f64>,
    /// Hash of each page's render commands from the last layout sent to the frontend
    page_hashes: Vec<u64>,
}

/// Runs `f` with a measurer: the JS callback behind the session's measurement cache, or
/// the built-in estimate when no callback is given
fn with_measurer<R>(
    cache: &mut HashMap<String, f64>,
    measure: Option<js_sys::Function>,
    f: impl FnOnce(&mut dyn TextMeasurer) -> R,
) -> R {
    match measure {
        Some(func) => {
            let mut measurer = CachedMeasurer::with_cache(JsMeasurer { func }, std::mem::take(cache));
            let result = f(&mut measurer);
            *cache = measurer.into_cache();
            result
        }
        None => f(&mut CachedMeasurer::new(EstimateMeasurer)),
    }
}

/// A picture or shape for the frontend: its kind and text instead of the parsed paragraphs
fn image_json(img: &docx_parser::ImageRef) -> serde_json::Value {
    let mut value = serde_json::to_value(img).unwrap_or_default();
    if let Some(obj) = value.as_object_mut() {
        obj.remove("rel_id");
        obj.remove("text_box");
        obj.remove("vml");
        let kind = if img.text_box.is_some() { "textbox" } else if img.shape.is_some() { "shape" } else { "picture" };
        obj.insert("kind".into(), kind.into());
        obj.insert("text".into(), img.text_box.as_ref().map(|t| t.text()).unwrap_or_default().into());
    }
    value
}

impl DocxSession {
    /// Where the last layout drew a body picture: `{ page, x, y, width, height }` or null
    fn image_bounds(&self, paragraph: usize, index: usize) -> serde_json::Value {
        let Some(layout) = &self.layout else { return serde_json::Value::Null };
        for page in &layout.pages {
            for item in &page.items {
                let found = match item {
                    layout_engine::RenderCommand::Image { x, y, width, height, paragraph_index: Some(p), image_index: Some(i), .. }
                    | layout_engine::RenderCommand::Shape { x, y, width, height, paragraph_index: Some(p), image_index: Some(i), .. } => {
                        (*p == paragraph && *i == index).then_some((*x, *y, *width, *height))
                    }
                    _ => None,
                };
                if let Some((x, y, width, height)) = found {
                    return serde_json::json!({ "page": page.page_number, "x": x, "y": y, "width": width, "height": height });
                }
            }
        }
        serde_json::Value::Null
    }
}

fn history_result(result: Option<Option<String>>) -> String {
    serde_json::json!({ "done": result.is_some(), "selection": result.flatten() }).to_string()
}

fn to_json<T: serde::Serialize>(value: &T) -> Result<String, JsValue> {
    serde_json::to_string(value).map_err(|e| JsValue::from_str(&e.to_string()))
}

#[wasm_bindgen]
impl DocxSession {
    /// Creates a new DocxSession from byte slice
    #[wasm_bindgen(constructor)]
    pub fn new(bytes: &[u8]) -> Result<DocxSession, JsValue> {
        let modifier = DocxModifier::from_bytes(bytes).map_err(|e| JsValue::from_str(&e))?;
        Ok(DocxSession { modifier, layout: None, measure_cache: HashMap::new(), page_hashes: Vec::new() })
    }

    /// Creates a session with sample contract docx
    #[wasm_bindgen]
    pub fn new_sample() -> Result<DocxSession, JsValue> {
        let bytes = sample_generator::generate_sample_docx().map_err(|e| JsValue::from_str(&e))?;
        let modifier = DocxModifier::from_bytes(&bytes).map_err(|e| JsValue::from_str(&e))?;
        Ok(DocxSession { modifier, layout: None, measure_cache: HashMap::new(), page_hashes: Vec::new() })
    }

    /// Creates a session with an empty document: one empty paragraph and Word's default
    /// styles. `page_size` is "a4" (default), "letter" or "legal".
    #[wasm_bindgen]
    pub fn new_blank(page_size: Option<String>) -> Result<DocxSession, JsValue> {
        let page = blank_generator::PageSize::parse(page_size.as_deref()).map_err(|e| JsValue::from_str(&e))?;
        let bytes = blank_generator::generate_blank_docx(page).map_err(|e| JsValue::from_str(&e))?;
        let modifier = DocxModifier::from_bytes(&bytes).map_err(|e| JsValue::from_str(&e))?;
        Ok(DocxSession { modifier, layout: None, measure_cache: HashMap::new(), page_hashes: Vec::new() })
    }

    /// Computes multi-page layout and render commands for high-performance Canvas rendering.
    /// `measure(text, family, sizePx, bold, italic) => width` should measure with the same
    /// canvas font the renderer uses; without it, glyph widths are estimated.
    /// With `only_changed`, pages identical to the previous call are sent as
    /// `{"page_number": n, "unchanged": true}` so the frontend can keep their canvases.
    #[wasm_bindgen]
    pub fn compute_canvas_layout_json(
        &mut self,
        watermark_text: Option<String>,
        watermark_opacity: f64,
        measure: Option<js_sys::Function>,
        only_changed: Option<bool>,
    ) -> Result<String, JsValue> {
        let elements = self.modifier.elements_shared().map_err(|e| JsValue::from_str(&e))?;
        let inputs = self.modifier.layout_inputs().map_err(|e| JsValue::from_str(&e))?;
        let layout = with_measurer(&mut self.measure_cache, measure, |m| {
            LayoutEngine::new().with_images(inputs.body_images.clone()).compute_layout_with(
                &elements,
                &inputs.background_color,
                &inputs.page_setup,
                &inputs.header_footer,
                inputs.bg_image_data_url.as_deref(),
                watermark_text.as_deref(),
                watermark_opacity,
                m,
            )
        });

        let mut pages_json = Vec::with_capacity(layout.pages.len());
        let mut hashes = Vec::with_capacity(layout.pages.len());
        for (i, page) in layout.pages.iter().enumerate() {
            let json = to_json(page)?;
            let mut hasher = std::collections::hash_map::DefaultHasher::new();
            json.hash(&mut hasher);
            let hash = hasher.finish();
            if only_changed.unwrap_or(false) && self.page_hashes.get(i) == Some(&hash) {
                pages_json.push(format!(r#"{{"page_number":{},"unchanged":true}}"#, page.page_number));
            } else {
                pages_json.push(json);
            }
            hashes.push(hash);
        }
        self.page_hashes = hashes;
        let json = format!(
            r#"{{"total_pages":{},"page_width":{},"page_height":{},"pages":[{}]}}"#,
            layout.total_pages,
            layout.page_width,
            layout.page_height,
            pages_json.join(",")
        );
        self.layout = Some(layout);
        Ok(json)
    }

    /// Forgets cached text measurements (call after web fonts finish loading)
    #[wasm_bindgen]
    pub fn reset_measurements(&mut self) {
        self.measure_cache.clear();
    }

    /// Text position `{paragraph, offset}` under a point of a page (1-based), or `null`
    #[wasm_bindgen]
    pub fn hit_test(&mut self, page: usize, x: f64, y: f64, measure: Option<js_sys::Function>) -> Result<String, JsValue> {
        let Some(layout) = &self.layout else { return Ok("null".into()) };
        to_json(&with_measurer(&mut self.measure_cache, measure, |m| caret::hit_test(layout, page, x, y, m)))
    }

    /// Caret box `{page, x, y, height, line_start, line_end}` for a text position, or `null`
    #[wasm_bindgen]
    pub fn caret_box(&mut self, paragraph: usize, offset: usize, measure: Option<js_sys::Function>) -> Result<String, JsValue> {
        let Some(layout) = &self.layout else { return Ok("null".into()) };
        let pos = TextPosition { paragraph, offset };
        to_json(&with_measurer(&mut self.measure_cache, measure, |m| caret::caret_box(layout, pos, m)))
    }

    /// Position one line up (`direction` < 0) or down at horizontal position `goal_x`, or `null`
    #[wasm_bindgen]
    pub fn move_vertical(
        &mut self,
        paragraph: usize,
        offset: usize,
        direction: i32,
        goal_x: f64,
        measure: Option<js_sys::Function>,
    ) -> Result<String, JsValue> {
        let Some(layout) = &self.layout else { return Ok("null".into()) };
        let pos = TextPosition { paragraph, offset };
        to_json(&with_measurer(&mut self.measure_cache, measure, |m| caret::move_vertical(layout, pos, direction, goal_x, m)))
    }

    /// Highlight rectangles `[{page, x, y, width, height}]` for characters `start..end`
    #[wasm_bindgen]
    pub fn selection_rects(
        &mut self,
        paragraph: usize,
        start: usize,
        end: usize,
        measure: Option<js_sys::Function>,
    ) -> Result<String, JsValue> {
        let Some(layout) = &self.layout else { return Ok("[]".into()) };
        to_json(&with_measurer(&mut self.measure_cache, measure, |m| caret::selection_rects(layout, paragraph, start, end, m)))
    }

    /// Replaces characters `start..end` of a body paragraph with `text` (offsets in Unicode
    /// code points). Typed text takes the formatting of the character before the caret.
    #[wasm_bindgen]
    pub fn replace_text(&mut self, paragraph: usize, start: usize, end: usize, text: &str) -> Result<bool, JsValue> {
        self.modifier.checkpoint(None);
        self.modifier
            .replace_paragraph_range(paragraph, start, end, text)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// The editor's edit primitive: replaces the selection from (`p1`,`o1`) to (`p2`,`o2`) with
    /// `text` (`\n` = new paragraph, `\u{000B}` = line break) and returns the caret
    /// `{paragraph, offset}`. With `coalesce`, the change joins the open undo step (typing);
    /// otherwise it starts a new one. `selection_before` / the returned caret are stored with
    /// the step so undo and redo can put the caret back.
    #[wasm_bindgen]
    #[allow(clippy::too_many_arguments)]
    pub fn edit(
        &mut self,
        p1: usize,
        o1: usize,
        p2: usize,
        o2: usize,
        text: &str,
        selection_before: Option<String>,
        coalesce: bool,
    ) -> Result<String, JsValue> {
        if !coalesce || !self.modifier.is_recording() {
            self.modifier.checkpoint(selection_before);
        }
        let (paragraph, offset) = self
            .modifier
            .replace_range(p1, o1, p2, o2, text)
            .map_err(|e| JsValue::from_str(&e))?;
        let caret = TextPosition { paragraph, offset };
        self.modifier.set_selection_after(Some(to_json(&caret)?));
        to_json(&caret)
    }

    /// Inserts paragraphs at a position as one undo step (see `DocxModifier::insert_paragraphs`).
    /// `paragraphs_json`: `[{ "text": "...", "bold": true, "font_size": 16, "align": "center" }, ...]`.
    /// Returns `{ "first": n, "count": n, "caret": { "paragraph": n, "offset": n } }`.
    #[wasm_bindgen]
    pub fn insert_paragraphs(
        &mut self,
        paragraph: usize,
        offset: usize,
        paragraphs_json: &str,
        selection_before: Option<String>,
    ) -> Result<String, JsValue> {
        let paragraphs: Vec<docx_parser::NewParagraph> = serde_json::from_str(paragraphs_json)
            .map_err(|e| JsValue::from_str(&format!("JSON deserialization error: {}", e)))?;
        self.modifier.checkpoint(selection_before);
        let (first, (p, o)) = self
            .modifier
            .insert_paragraphs(paragraph, offset, &paragraphs)
            .map_err(|e| JsValue::from_str(&e))?;
        let caret = TextPosition { paragraph: p, offset: o };
        self.modifier.set_selection_after(Some(to_json(&caret)?));
        to_json(&serde_json::json!({ "first": first, "count": paragraphs.len(), "caret": caret }))
    }

    /// Inserts a table at a position as one undo step (see `DocxModifier::insert_table`).
    /// `table_json`: `{ "rows": [["Ítem", "Valor"], ["A", "1"]], "header": true, "widths": [2, 1], "align": ["left", "right"] }`.
    /// Returns `{ "first": n, "caret": { "paragraph": n, "offset": 0 } }` (`first` = first cell's paragraph).
    #[wasm_bindgen]
    pub fn insert_table(
        &mut self,
        paragraph: usize,
        offset: usize,
        table_json: &str,
        selection_before: Option<String>,
    ) -> Result<String, JsValue> {
        let table: docx_parser::NewTable = serde_json::from_str(table_json)
            .map_err(|e| JsValue::from_str(&format!("JSON deserialization error: {}", e)))?;
        self.modifier.checkpoint(selection_before);
        let (first, (p, o)) = self
            .modifier
            .insert_table(paragraph, offset, &table)
            .map_err(|e| JsValue::from_str(&e))?;
        let caret = TextPosition { paragraph: p, offset: o };
        self.modifier.set_selection_after(Some(to_json(&caret)?));
        to_json(&serde_json::json!({ "first": first, "caret": caret }))
    }

    /// Inserts a PNG, JPEG or GIF picture in its own paragraph as one undo step (see
    /// `DocxModifier::insert_image`). `options_json`: `{ "width": px, "height": px, "align": "center", "alt": "..." }`.
    /// Returns `{ "paragraph": n, "caret": { "paragraph": n, "offset": 0 } }`.
    #[wasm_bindgen]
    pub fn insert_image(
        &mut self,
        paragraph: usize,
        offset: usize,
        bytes: &[u8],
        options_json: &str,
        selection_before: Option<String>,
    ) -> Result<String, JsValue> {
        let options: docx_parser::NewImage = serde_json::from_str(options_json)
            .map_err(|e| JsValue::from_str(&format!("JSON deserialization error: {}", e)))?;
        self.modifier.checkpoint(selection_before);
        let (index, (p, o)) = self
            .modifier
            .insert_image(paragraph, offset, bytes, &options)
            .map_err(|e| JsValue::from_str(&e))?;
        let caret = TextPosition { paragraph: p, offset: o };
        self.modifier.set_selection_after(Some(to_json(&caret)?));
        to_json(&serde_json::json!({ "paragraph": index, "caret": caret }))
    }

    /// Every picture, text box and shape of the body (table cells included) in document order:
    /// `[{ paragraph, index, kind: "picture" | "textbox" | "shape", width, height, wrap,
    /// wrap_side, anchored, behind_text, h_relative, h_offset, h_align, v_relative, v_offset,
    /// v_align, dist_top…, offset, alt, doc_pr_id, shape, text, bounds: { page, x, y, width,
    /// height } | null }]`. `bounds` is where the last layout drew it. Legacy VML shapes are
    /// drawn but left out (they cannot be edited).
    #[wasm_bindgen]
    pub fn list_images(&self) -> Result<String, JsValue> {
        let elements = self.modifier.elements_shared().map_err(|e| JsValue::from_str(&e))?;
        let mut paragraphs: Vec<&docx_parser::ParagraphInfo> = Vec::new();
        for el in elements.iter() {
            match el {
                docx_parser::DocumentElement::Paragraph(p) => paragraphs.push(p),
                docx_parser::DocumentElement::Table(t) => {
                    for row in &t.rich_rows {
                        for cell in &row.cells {
                            paragraphs.extend(cell.paragraphs.iter());
                        }
                    }
                }
            }
        }
        paragraphs.sort_by_key(|p| p.index);
        let images: Vec<serde_json::Value> = paragraphs
            .iter()
            .flat_map(|p| p.images.iter().enumerate().map(move |(i, img)| (p.index, i, img)))
            .filter(|(_, _, img)| !img.vml)
            .map(|(paragraph, index, img)| {
                let mut value = image_json(img);
                if let Some(obj) = value.as_object_mut() {
                    obj.insert("paragraph".into(), paragraph.into());
                    obj.insert("index".into(), index.into());
                    obj.insert("bounds".into(), self.image_bounds(paragraph, index));
                }
                value
            })
            .collect();
        to_json(&images)
    }

    /// Changes size, position, text wrapping, distances or alternative text of picture `index`
    /// of `paragraph` as one undo step (see `docx_parser::ImageUpdate`).
    /// `update_json`: `{ "width": 200, "wrap": "square", "h_relative": "margin", "h_align": "right" }`.
    /// Returns the picture as it is now (same shape as `list_images` items, without `bounds`).
    #[wasm_bindgen]
    pub fn update_image(
        &mut self,
        paragraph: usize,
        index: usize,
        update_json: &str,
        selection_before: Option<String>,
    ) -> Result<String, JsValue> {
        let update: docx_parser::ImageUpdate = serde_json::from_str(update_json)
            .map_err(|e| JsValue::from_str(&format!("JSON deserialization error: {}", e)))?;
        self.modifier.checkpoint(selection_before.clone());
        let image = self
            .modifier
            .update_image(paragraph, index, &update)
            .map_err(|e| JsValue::from_str(&e))?;
        self.modifier.set_selection_after(selection_before);
        let mut value = image_json(&image);
        if let Some(obj) = value.as_object_mut() {
            obj.insert("paragraph".into(), paragraph.into());
            obj.insert("index".into(), index.into());
        }
        to_json(&value)
    }

    /// Deletes picture `index` of `paragraph` as one undo step
    #[wasm_bindgen]
    pub fn delete_image(&mut self, paragraph: usize, index: usize, selection_before: Option<String>) -> Result<(), JsValue> {
        self.modifier.checkpoint(selection_before.clone());
        self.modifier.delete_image(paragraph, index).map_err(|e| JsValue::from_str(&e))?;
        self.modifier.set_selection_after(selection_before);
        Ok(())
    }

    /// Undoes the last step: `{done, selection}` (selection is the JSON given before the step)
    #[wasm_bindgen]
    pub fn undo(&mut self) -> String {
        history_result(self.modifier.undo())
    }

    /// Redoes the last undone step: `{done, selection}`
    #[wasm_bindgen]
    pub fn redo(&mut self) -> String {
        history_result(self.modifier.redo())
    }

    #[wasm_bindgen]
    pub fn can_undo(&self) -> bool {
        self.modifier.can_undo()
    }

    #[wasm_bindgen]
    pub fn can_redo(&self) -> bool {
        self.modifier.can_redo()
    }

    /// Highlight rectangles for a selection spanning paragraphs
    #[wasm_bindgen]
    pub fn selection_rects_range(
        &mut self,
        p1: usize,
        o1: usize,
        p2: usize,
        o2: usize,
        measure: Option<js_sys::Function>,
    ) -> Result<String, JsValue> {
        let Some(layout) = &self.layout else { return Ok("[]".into()) };
        let from = TextPosition { paragraph: p1, offset: o1 };
        let to = TextPosition { paragraph: p2, offset: o2 };
        to_json(&with_measurer(&mut self.measure_cache, measure, |m| caret::selection_rects_range(layout, from, to, m)))
    }

    /// Gets all document elements (paragraphs and tables in order) as JSON
    #[wasm_bindgen]
    pub fn get_document_elements_json(&self) -> Result<String, JsValue> {
        let elements = self.modifier.elements_shared().map_err(|e| JsValue::from_str(&e))?;
        to_json(&*elements)
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
        self.modifier.checkpoint(None);
        let res = self
            .modifier
            .find_and_replace(search, replacement, match_case, use_regex)
            .map_err(|e| JsValue::from_str(&e))?;
        serde_json::to_string(&res).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Performs batch replacement with a JSON array of [{ "key": "...", "value": "..." }]
    #[wasm_bindgen]
    pub fn batch_replace(&mut self, pairs_json: &str) -> Result<String, JsValue> {
        self.modifier.checkpoint(None);
        let pairs: Vec<KeyValuePair> = serde_json::from_str(pairs_json)
            .map_err(|e| JsValue::from_str(&format!("JSON inválido para batch replace: {}", e)))?;
        let res = self.modifier.batch_replace(&pairs).map_err(|e| JsValue::from_str(&e))?;
        serde_json::to_string(&res).map_err(|e| JsValue::from_str(&e.to_string()))
    }

    /// Updates a single paragraph text by index
    #[wasm_bindgen]
    pub fn update_paragraph(&mut self, index: usize, new_text: &str) -> Result<bool, JsValue> {
        self.modifier.checkpoint(None);
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
        self.modifier.checkpoint(None);
        let runs: Vec<crate::docx_parser::RunInfo> = serde_json::from_str(runs_json)
            .map_err(|e| JsValue::from_str(&format!("JSON deserialization error: {}", e)))?;
        self.modifier
            .update_paragraph_runs(index, &runs, align.as_deref())
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Applies run formatting (and alignment) to several paragraphs as one undo step:
    /// `[{ "index": 3, "runs": [...], "align": "left" }, ...]`
    #[wasm_bindgen]
    pub fn update_paragraphs_runs(&mut self, updates_json: &str) -> Result<bool, JsValue> {
        #[derive(serde::Deserialize)]
        struct RunsUpdate {
            index: usize,
            runs: Vec<crate::docx_parser::RunInfo>,
            align: Option<String>,
        }
        let updates: Vec<RunsUpdate> = serde_json::from_str(updates_json)
            .map_err(|e| JsValue::from_str(&format!("JSON deserialization error: {}", e)))?;
        self.modifier.checkpoint(None);
        for u in updates {
            self.modifier
                .update_paragraph_runs(u.index, &u.runs, u.align.as_deref())
                .map_err(|e| JsValue::from_str(&e))?;
        }
        Ok(true)
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
        self.modifier.checkpoint(None);
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
        self.modifier.checkpoint(None);
        self.modifier
            .update_table_cell(table_index, row, col, new_text)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Inserts a new table into document
    #[wasm_bindgen]
    pub fn add_table(&mut self, rows: usize, cols: usize, headers_json: &str) -> Result<bool, JsValue> {
        self.modifier.checkpoint(None);
        let headers: Vec<String> = serde_json::from_str(headers_json)
            .unwrap_or_else(|_| (1..=cols).map(|i| format!("Columna {}", i)).collect());
        self.modifier
            .add_table(rows, cols, &headers)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Sets page background color (HEX without '#')
    #[wasm_bindgen]
    pub fn set_background_color(&mut self, hex_color: &str) -> Result<(), JsValue> {
        self.modifier.checkpoint(None);
        self.modifier
            .set_background_color(hex_color)
            .map_err(|e| JsValue::from_str(&e))
    }

    /// Sets background / watermark image from bytes
    #[wasm_bindgen]
    pub fn set_background_image(&mut self, bytes: &[u8], ext: &str) -> Result<(), JsValue> {
        self.modifier.checkpoint(None);
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
