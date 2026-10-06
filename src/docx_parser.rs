use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine;
use quick_xml::events::{BytesStart, BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use regex::Regex;
use serde::{Deserialize, Serialize};
use std::cell::RefCell;
use std::collections::HashMap;
use std::hash::{Hash, Hasher};
use std::rc::Rc;
use std::io::{Cursor, Read, Write};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

use crate::styles::{
    apply_border_element, apply_ppr_element, apply_rpr_element, is_symbol_font, NumberingCounters, ParaProps,
    RunProps, StyleSheet,
};
use crate::paragraph_edit::{
    body_paragraph_ranges, edit_paragraph, edit_paragraph_range, is_balanced, is_blank, merge_paragraphs,
    parse_paragraph_fragment, split_paragraph, table_cell_paragraph_ranges,
    FormatTarget,
};

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
pub struct BorderInfo {
    pub val: String,   // "single", "dashed", etc.
    pub color: String, // hex without '#' e.g. "1F6F6B"
    pub sz_px: f64,    // border width converted to px
    pub space: f64,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
pub struct ParagraphBorders {
    pub top: Option<BorderInfo>,
    pub bottom: Option<BorderInfo>,
    pub left: Option<BorderInfo>,
    pub right: Option<BorderInfo>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
pub struct RunInfo {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub color: String,
    pub font_size: Option<f64>,
    pub font_family: Option<String>,
}

#[derive(Serialize, Deserialize, Debug, Clone, Default)]
pub struct ParagraphInfo {
    pub index: usize,
    pub text: String,
    pub style: String,
    pub is_heading: bool,
    pub run_count: usize,
    pub align: String, // "left", "center", "right", "both"
    pub color: String, // hex without '#' e.g. "1E3A8A" or empty
    pub bold: bool,
    pub italic: bool,
    pub font_size: Option<f64>,
    pub font_family: Option<String>,
    pub indent_left: f64,
    pub indent_first_line: f64,
    pub indent_right: f64,
    pub space_before: f64,
    pub space_after: f64,
    pub line_spacing: Option<f64>,
    pub borders: ParagraphBorders,
    #[serde(default)]
    pub runs: Vec<RunInfo>,
    /// Rendered list number/bullet ("1.", "a)", "•") with its formatting
    #[serde(default)]
    pub list_label: Option<RunInfo>,
    /// Fixed line height in points (lineRule exact)
    #[serde(default)]
    pub line_exact: Option<f64>,
    /// Minimum line height in points (lineRule atLeast)
    #[serde(default)]
    pub line_at_least: Option<f64>,
    #[serde(default)]
    pub keep_next: bool,
    #[serde(default)]
    pub keep_lines: bool,
    #[serde(default)]
    pub page_break_before: bool,
    #[serde(default)]
    pub widow_control: bool,
    #[serde(default)]
    pub contextual_spacing: bool,
    /// Spacing comes from the document (styles/direct) rather than layout heuristics
    #[serde(default)]
    pub spacing_resolved: bool,
    /// The next section starts on a new page after this paragraph
    #[serde(default)]
    pub section_break_after: bool,
    /// This paragraph ends a section (`w:sectPr` in its properties)
    #[serde(skip)]
    pub ends_section: bool,
    /// Pictures anchored in or inline with this paragraph
    #[serde(default)]
    pub images: Vec<ImageRef>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct TableCellInfo {
    pub row: usize,
    pub col: usize,
    pub text: String,
}

#[derive(Serialize, Deserialize, Debug, Clone, Default)]
pub struct TableCellData {
    pub text: String,
    pub bg_color: Option<String>,
    pub color: String,
    pub align: String,
    pub bold: bool,
    pub italic: bool,
    pub font_size: f64,
    pub font_family: String,
    pub border_color: String,
    /// The cell's paragraphs, numbered in the same document-wide sequence as body paragraphs
    #[serde(default)]
    pub paragraphs: Vec<ParagraphInfo>,
}

#[derive(Serialize, Deserialize, Debug, Clone, Default)]
pub struct TableRowData {
    pub cells: Vec<TableCellData>,
    pub is_header: bool,
}

#[derive(Serialize, Deserialize, Debug, Clone, Default)]
pub struct TableInfo {
    pub index: usize,
    pub rows: Vec<Vec<String>>,
    #[serde(default)]
    pub rich_rows: Vec<TableRowData>,
    #[serde(default)]
    pub grid_cols: Vec<f64>,
    pub header_row: bool,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct PageSetup {
    pub width: f64,
    pub height: f64,
    pub orientation: String, // "portrait" or "landscape"
    pub margin_top: f64,
    pub margin_right: f64,
    pub margin_bottom: f64,
    pub margin_left: f64,
    pub header_margin: f64,
    pub footer_margin: f64,
}

impl Default for PageSetup {
    fn default() -> Self {
        PageSetup {
            width: 800.0,
            height: 1130.0,
            orientation: "portrait".to_string(),
            margin_top: 70.0,
            margin_right: 65.0,
            margin_bottom: 70.0,
            margin_left: 65.0,
            header_margin: 36.0,
            footer_margin: 36.0,
        }
    }
}

#[derive(Serialize, Deserialize, Debug, Clone, Default)]
pub struct HeaderFooterInfo {
    pub header_text: String,
    pub footer_text: String,
    pub has_header: bool,
    pub has_footer: bool,
    pub header_image_data_url: Option<String>,
    /// Images of the default header and footer, positioned like Word does on every page
    #[serde(default)]
    pub header_images: Vec<PlacedImage>,
    #[serde(default)]
    pub footer_images: Vec<PlacedImage>,
}

/// A picture (`w:drawing`) as written in the document. Lengths in px.
#[derive(Serialize, Deserialize, Debug, Clone, Default, PartialEq)]
pub struct ImageRef {
    /// Relationship id of the picture (`r:embed`)
    pub rel_id: String,
    pub width: f64,
    pub height: f64,
    /// Floating (`wp:anchor`) rather than in line with the text (`wp:inline`)
    pub anchored: bool,
    /// Drawn behind the text (`behindDoc`)
    pub behind_text: bool,
    /// What the horizontal position is relative to: page, margin, column, character…
    pub h_relative: String,
    pub h_offset: f64,
    /// left / center / right instead of an offset
    pub h_align: Option<String>,
    /// What the vertical position is relative to: page, margin, paragraph, line…
    pub v_relative: String,
    pub v_offset: f64,
    /// top / center / bottom instead of an offset
    pub v_align: Option<String>,
}

/// An image together with its pixels as a data URL
#[derive(Serialize, Deserialize, Debug, Clone, Default)]
pub struct PlacedImage {
    pub image: ImageRef,
    pub data_url: String,
}

const EMU_PER_PX: f64 = 9525.0;

/// Reads the `w:drawing` whose start tag was just consumed
pub(crate) fn parse_drawing(reader: &mut Reader<&[u8]>) -> Option<ImageRef> {
    let mut img = ImageRef {
        h_relative: "column".into(),
        v_relative: "paragraph".into(),
        ..Default::default()
    };
    let mut axis = ' ';
    let mut text_kind: Option<&'static str> = None;
    let mut buf = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buf);
        let (e, is_start) = match event {
            Ok(Event::Start(e)) => (e, true),
            Ok(Event::Empty(e)) => (e, false),
            Ok(Event::Text(t)) => {
                if let (Some(kind), Ok(value)) = (text_kind, t.unescape()) {
                    let value = value.trim().to_string();
                    match (axis, kind) {
                        ('h', "offset") => img.h_offset = value.parse::<f64>().unwrap_or(0.0) / EMU_PER_PX,
                        ('v', "offset") => img.v_offset = value.parse::<f64>().unwrap_or(0.0) / EMU_PER_PX,
                        ('h', "align") => img.h_align = Some(value),
                        ('v', "align") => img.v_align = Some(value),
                        _ => {}
                    }
                }
                buf.clear();
                continue;
            }
            Ok(Event::End(e)) => {
                let name = e.name();
                let n = name.as_ref();
                if tag_is(n, "drawing") {
                    break;
                } else if tag_is(n, "positionH") || tag_is(n, "positionV") {
                    axis = ' ';
                } else if tag_is(n, "posOffset") || tag_is(n, "align") {
                    text_kind = None;
                }
                buf.clear();
                continue;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {
                buf.clear();
                continue;
            }
        };
        let name = e.name();
        let n = name.as_ref();
        if tag_is(n, "anchor") {
            img.anchored = true;
            img.behind_text = get_attr_value(&e, "behindDoc").is_some_and(|v| v == "1" || v == "true");
        } else if tag_is(n, "positionH") || tag_is(n, "positionV") {
            axis = if tag_is(n, "positionH") { 'h' } else { 'v' };
            if let Some(rel) = get_attr_value(&e, "relativeFrom") {
                if axis == 'h' { img.h_relative = rel } else { img.v_relative = rel }
            }
        } else if is_start && tag_is(n, "posOffset") {
            text_kind = Some("offset");
        } else if is_start && tag_is(n, "align") {
            text_kind = Some("align");
        } else if tag_is(n, "extent") && img.width == 0.0 {
            img.width = get_attr_i64(&e, "cx").unwrap_or(0) as f64 / EMU_PER_PX;
            img.height = get_attr_i64(&e, "cy").unwrap_or(0) as f64 / EMU_PER_PX;
        } else if tag_is(n, "blip") && img.rel_id.is_empty() {
            img.rel_id = get_attr_value(&e, "embed").unwrap_or_default();
        } else if is_start && tag_is(n, "txbxContent") {
            // Text boxes are not rendered yet; skip their content
            let end = n.to_vec();
            let _ = reader.read_to_end_into(quick_xml::name::QName(&end), &mut Vec::new());
        }
        buf.clear();
    }
    (!img.rel_id.is_empty() && img.width > 0.0).then_some(img)
}

/// Every picture of an XML part (e.g. a header), in document order
pub(crate) fn extract_drawings(xml: &str) -> Vec<ImageRef> {
    let mut reader = Reader::from_str(xml);
    let mut images = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) if tag_is(e.name().as_ref(), "drawing") => {
                buf.clear();
                if let Some(img) = parse_drawing(&mut reader) {
                    images.push(img);
                }
                continue;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    images
}

/// `data:` URL of a related image part (`target` as written in a .rels file)
pub(crate) fn image_data_url(files: &HashMap<String, Vec<u8>>, target: &str) -> Option<String> {
    let clean = target.trim_start_matches("../").trim_start_matches('/');
    let path = if clean.starts_with("word/") { clean.to_string() } else { format!("word/{}", clean) };
    let bytes = files.get(&path)?;
    let lower = path.to_lowercase();
    let mime = if lower.ends_with(".jpg") || lower.ends_with(".jpeg") {
        "image/jpeg"
    } else if lower.ends_with(".gif") {
        "image/gif"
    } else if lower.ends_with(".svg") {
        "image/svg+xml"
    } else {
        "image/png"
    };
    Some(format!("data:{};base64,{}", mime, BASE64.encode(bytes)))
}

/// Pictures of a part with their pixels, resolved through the part's relationships
fn placed_images(files: &HashMap<String, Vec<u8>>, part: &str) -> Vec<PlacedImage> {
    let Some(xml) = files.get(part).and_then(|b| std::str::from_utf8(b).ok()) else { return Vec::new() };
    let rels_path = match part.rsplit_once('/') {
        Some((dir, file)) => format!("{}/_rels/{}.rels", dir, file),
        None => return Vec::new(),
    };
    let rels = files
        .get(&rels_path)
        .and_then(|b| std::str::from_utf8(b).ok())
        .map(parse_relationships_map)
        .unwrap_or_default();
    extract_drawings(xml)
        .into_iter()
        .filter_map(|image| {
            let data_url = image_data_url(files, rels.get(&image.rel_id)?)?;
            Some(PlacedImage { image, data_url })
        })
        .collect()
}

/// Part name of the default header or footer of the document's last section
fn default_header_footer_part(files: &HashMap<String, Vec<u8>>, kind: &str) -> Option<String> {
    let doc = std::str::from_utf8(files.get("word/document.xml")?).ok()?;
    let rels = parse_relationships_map(std::str::from_utf8(files.get("word/_rels/document.xml.rels")?).ok()?);
    let tag = format!("<w:{}Reference ", kind);
    let mut chosen = None;
    let mut from = 0;
    while let Some(i) = doc[from..].find(&tag).map(|i| i + from) {
        let end = doc[i..].find('>').map(|j| i + j).unwrap_or(doc.len());
        let element = &doc[i..end];
        if element.contains("w:type=\"default\"") {
            chosen = element.split("r:id=\"").nth(1).and_then(|r| r.split('"').next()).map(str::to_string);
        }
        from = end;
    }
    let target = rels.get(&chosen?)?;
    Some(format!("word/{}", target.trim_start_matches('/').trim_start_matches("word/")))
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(tag = "type")]
pub enum DocumentElement {
    #[serde(rename = "paragraph")]
    Paragraph(ParagraphInfo),
    #[serde(rename = "table")]
    Table(TableInfo),
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ParagraphUpdate {
    pub index: usize,
    pub text: String,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct KeyValuePair {
    pub key: String,
    pub value: String,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct DocxStats {
    pub paragraph_count: usize,
    pub table_count: usize,
    pub word_count: usize,
    pub char_count: usize,
    pub files_in_zip: Vec<String>,
    pub original_size_bytes: usize,
    pub background_color: String,
    pub has_background_image: bool,
    pub page_setup: PageSetup,
    pub header_footer: HeaderFooterInfo,
    pub bg_image_data_url: Option<String>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ReplaceResult {
    pub occurrences_replaced: usize,
    pub affected_files: Vec<String>,
    pub message: String,
}

/// One paragraph edit: new text, optional per-character formatting and alignment
struct ParagraphEdit<'a> {
    index: usize,
    text: &'a str,
    formats: Option<Vec<FormatTarget>>,
    align: Option<&'a str>,
}

pub struct DocxModifier {
    files: HashMap<String, Vec<u8>>,
    original_order: Vec<String>,
    original_size: usize,
    background_color: String,
    background_image: Option<Vec<u8>>,
    background_image_ext: String,
    styles: StyleSheet,
    history: History,
    /// Bumped on every change of any part; `parts_revision` only for parts other than
    /// document.xml (headers, media, relationships), which typing never touches
    revision: u64,
    parts_revision: u64,
    cache: RefCell<ModelCache>,
}

/// Inputs of the page layout besides the body content
#[derive(Debug, Clone)]
pub struct LayoutInputs {
    pub background_color: String,
    pub page_setup: PageSetup,
    pub header_footer: HeaderFooterInfo,
    pub bg_image_data_url: Option<String>,
    /// Pixels of the body's pictures by relationship id
    pub body_images: HashMap<String, String>,
}

/// Derived data reused until the document changes (parsing and image encoding are the
/// expensive parts of every keystroke otherwise)
#[derive(Default)]
struct ModelCache {
    elements: Option<(u64, Rc<Vec<DocumentElement>>)>,
    /// Keyed by (signature of the layout-relevant bits of document.xml, parts revision)
    inputs: Option<(u64, u64, Rc<LayoutInputs>)>,
}

/// Hash of what page setup, background and images depend on in document.xml: section
/// properties, the page background and image references (not the text)
fn layout_inputs_signature(doc: &str) -> u64 {
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    let mut hash_spans = |open: &str, close: &str| {
        let mut from = 0;
        while let Some(start) = doc[from..].find(open).map(|i| i + from) {
            let end = doc[start..].find(close).map(|i| start + i + close.len()).unwrap_or(doc.len());
            doc[start..end].hash(&mut hasher);
            from = end;
        }
    };
    hash_spans("<w:sectPr", "</w:sectPr>");
    hash_spans("<w:background", ">");
    hash_spans("r:embed=\"", "\"");
    hasher.finish()
}

/// One undoable step: previous contents of every part it changed, plus the editor selection
/// before and after (opaque JSON owned by the frontend)
#[derive(Default)]
struct UndoStep {
    parts: Vec<(String, Option<Vec<u8>>)>,
    selection_before: Option<String>,
    selection_after: Option<String>,
}

#[derive(Default)]
struct History {
    undo: Vec<UndoStep>,
    redo: Vec<UndoStep>,
    /// Step being recorded since the last checkpoint
    recording: Option<UndoStep>,
}

const MAX_UNDO_STEPS: usize = 200;

impl DocxModifier {
    /// Loads a docx from byte buffer
    pub fn from_bytes(bytes: &[u8]) -> Result<Self, String> {
        if bytes.len() < 4 {
            return Err("El archivo proporcionado está vacío o es demasiado pequeño.".to_string());
        }

        let original_size = bytes.len();
        let reader = Cursor::new(bytes);
        let mut archive = ZipArchive::new(reader)
            .map_err(|e| format!("No se pudo leer el archivo como formato DOCX (ZIP): {}", e))?;

        let mut files = HashMap::new();
        let mut original_order = Vec::new();

        for i in 0..archive.len() {
            let mut file = archive
                .by_index(i)
                .map_err(|e| format!("Error al leer entrada ZIP {}: {}", i, e))?;
            let name = file.name().to_string();
            let mut content = Vec::new();
            file.read_to_end(&mut content)
                .map_err(|e| format!("Error al extraer '{}': {}", name, e))?;

            files.insert(name.clone(), content);
            original_order.push(name);
        }

        // Verify that this is indeed a docx
        if !files.contains_key("word/document.xml") {
            return Err("El archivo no es un documento DOCX válido (falta 'word/document.xml').".to_string());
        }

        // Extract background color if present
        let mut bg_color = "FFFFFF".to_string();
        if let Some(doc_bytes) = files.get("word/document.xml") {
            if let Ok(doc_str) = String::from_utf8(doc_bytes.clone()) {
                bg_color = extract_bg_color_quick_xml(&doc_str);
            }
        }

        // Custom background image set dynamically via API (None by default; DOCX embedded images are resolved via relationships)
        let bg_image = None;
        let bg_image_ext = "png".to_string();

        let part = |path: &str| files.get(path).and_then(|b| std::str::from_utf8(b).ok());
        let theme_path = files.keys().filter(|k| k.starts_with("word/theme/") && k.ends_with(".xml")).min().cloned();
        let styles = StyleSheet::load(
            part("word/styles.xml"),
            theme_path.as_deref().and_then(part),
            part("word/numbering.xml"),
        );

        Ok(DocxModifier {
            history: History::default(),
            revision: 0,
            parts_revision: 0,
            cache: RefCell::new(ModelCache::default()),
            styles,
            files,
            original_order,
            original_size,
            background_color: bg_color,
            background_image: bg_image,
            background_image_ext: bg_image_ext,
        })
    }

    /// Writes a package part, remembering its previous content for undo when a step is open
    fn put_file(&mut self, name: String, bytes: Vec<u8>) {
        self.revision += 1;
        if name != "word/document.xml" {
            self.parts_revision += 1;
        }
        if let Some(step) = self.history.recording.as_mut() {
            if !step.parts.iter().any(|(n, _)| *n == name) {
                step.parts.push((name.clone(), self.files.get(&name).cloned()));
                self.history.redo.clear();
            }
        }
        self.files.insert(name, bytes);
    }

    /// Starts a new undo step; changes made until the next checkpoint undo together
    pub fn checkpoint(&mut self, selection_before: Option<String>) {
        self.finish_step();
        self.history.recording = Some(UndoStep { selection_before, ..Default::default() });
    }

    /// Records where the selection ended up after the current step's changes
    pub fn set_selection_after(&mut self, selection: Option<String>) {
        if let Some(step) = self.history.recording.as_mut() {
            step.selection_after = selection;
        }
    }

    /// True while a step is being recorded (edits can be merged into it, e.g. while typing)
    pub fn is_recording(&self) -> bool {
        self.history.recording.is_some()
    }

    fn finish_step(&mut self) {
        if let Some(step) = self.history.recording.take() {
            if !step.parts.is_empty() {
                self.history.undo.push(step);
                if self.history.undo.len() > MAX_UNDO_STEPS {
                    self.history.undo.remove(0);
                }
            }
        }
    }

    /// Swaps the stored part contents with the current ones and returns the inverse step
    fn apply_step(&mut self, step: UndoStep) -> UndoStep {
        self.revision += 1;
        self.parts_revision += 1;
        let parts = step
            .parts
            .into_iter()
            .map(|(name, content)| {
                let current = match content {
                    Some(bytes) => self.files.insert(name.clone(), bytes),
                    None => self.files.remove(&name),
                };
                (name, current)
            })
            .collect();
        UndoStep { parts, selection_before: step.selection_before, selection_after: step.selection_after }
    }

    /// Undoes the last step; returns the selection to restore (`Some(None)` when unknown)
    pub fn undo(&mut self) -> Option<Option<String>> {
        self.finish_step();
        let step = self.history.undo.pop()?;
        let inverse = self.apply_step(step);
        let selection = inverse.selection_before.clone();
        self.history.redo.push(inverse);
        Some(selection)
    }

    /// Redoes the last undone step; returns the selection to restore
    pub fn redo(&mut self) -> Option<Option<String>> {
        self.finish_step();
        let step = self.history.redo.pop()?;
        let inverse = self.apply_step(step);
        let selection = inverse.selection_after.clone();
        self.history.undo.push(inverse);
        Some(selection)
    }

    pub fn can_undo(&self) -> bool {
        !self.history.undo.is_empty() || self.history.recording.as_ref().is_some_and(|s| !s.parts.is_empty())
    }

    pub fn can_redo(&self) -> bool {
        !self.history.redo.is_empty()
    }

    fn body_paragraph(&self, xml: &str, index: usize) -> Result<std::ops::Range<usize>, String> {
        body_paragraph_ranges(xml)?
            .get(index)
            .cloned()
            .ok_or_else(|| format!("No existe el párrafo {}.", index))
    }

    /// Splits body paragraph `index` at `offset` (Enter)
    pub fn split_paragraph(&mut self, index: usize, offset: usize) -> Result<(), String> {
        let mut xml = self.get_file_string("word/document.xml")?;
        let range = self.body_paragraph(&xml, index)?;
        let (first, second) = split_paragraph(&xml[range.clone()], &self.styles, offset)?;
        xml.replace_range(range, &(first + &second));
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        Ok(())
    }

    /// Deletes from (`p1`, `o1`) to (`p2`, `o2`) across paragraphs: whole paragraphs and tables
    /// in between are removed and the two ends are joined into one paragraph
    pub fn delete_range(&mut self, p1: usize, o1: usize, p2: usize, o2: usize) -> Result<(), String> {
        if p1 == p2 {
            return self.replace_paragraph_range(p1, o1, o2, "").map(|_| ());
        }
        if p2 < p1 {
            return self.delete_range(p2, o2, p1, o1);
        }
        let mut xml = self.get_file_string("word/document.xml")?;
        let ranges = body_paragraph_ranges(&xml)?;
        let (r1, r2) = match (ranges.get(p1), ranges.get(p2)) {
            (Some(a), Some(b)) => (a.clone(), b.clone()),
            _ => return Err("La selección apunta a párrafos que no existen.".to_string()),
        };
        // Cutting must not tear apart a content control or other wrapper element
        if !is_balanced(&xml[r1.end..r2.start]) {
            return Err("La selección cruza el borde de una tabla o de un control de contenido; no se puede borrar de una vez.".to_string());
        }
        let head = edit_paragraph_range(&xml[r1.clone()], &self.styles, o1, usize::MAX, "")?;
        let tail = edit_paragraph_range(&xml[r2.clone()], &self.styles, 0, o2, "")?;
        let merged = merge_paragraphs(&head, &tail, &self.styles)?;
        xml.replace_range(r1.start..r2.end, &merged);
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        Ok(())
    }

    /// Joins body paragraph `index` with the next one when nothing (e.g. a table) sits between
    pub fn merge_with_next(&mut self, index: usize) -> Result<(), String> {
        let xml = self.get_file_string("word/document.xml")?;
        let ranges = body_paragraph_ranges(&xml)?;
        match (ranges.get(index), ranges.get(index + 1)) {
            (Some(a), Some(b)) if is_blank(&xml[a.end..b.start]) => {}
            _ => return Err("Solo se pueden unir párrafos contiguos.".to_string()),
        }
        let len = parse_paragraph_fragment(&xml[ranges[index].clone()], &self.styles).text.chars().count();
        self.delete_range(index, len, index + 1, 0)
    }

    /// The editor's single edit primitive: replaces the selection (`p1`,`o1`)–(`p2`,`o2`) with
    /// `text`. `\n` in `text` starts a new paragraph and `\u{000B}` is a line break inside
    /// the paragraph (like Word's ^l). Returns the caret position after the inserted text.
    pub fn replace_range(
        &mut self,
        p1: usize,
        o1: usize,
        p2: usize,
        o2: usize,
        text: &str,
    ) -> Result<(usize, usize), String> {
        let ((p1, o1), (p2, o2)) = if (p2, o2) < (p1, o1) { ((p2, o2), (p1, o1)) } else { ((p1, o1), (p2, o2)) };
        if (p1, o1) != (p2, o2) {
            self.delete_range(p1, o1, p2, o2)?;
        }
        let (mut p, mut o) = (p1, o1);
        for (i, line) in text.split('\n').enumerate() {
            if i > 0 {
                self.split_paragraph(p, o)?;
                p += 1;
                o = 0;
            }
            if !line.is_empty() {
                let line = line.replace('\u{000B}', "\n");
                self.replace_paragraph_range(p, o, o, &line)?;
                o += line.chars().count();
            }
        }
        Ok((p, o))
    }

    /// Extracts list of document elements (paragraphs and tables) in sequential order
    pub fn extract_elements(&self) -> Result<Vec<DocumentElement>, String> {
        Ok((*self.elements_shared()?).clone())
    }

    /// Parsed document elements, parsed once per revision of the document
    pub fn elements_shared(&self) -> Result<Rc<Vec<DocumentElement>>, String> {
        if let Some((rev, elements)) = &self.cache.borrow().elements {
            if *rev == self.revision {
                return Ok(elements.clone());
            }
        }
        let doc_xml = self.get_file_string("word/document.xml")?;
        let elements = Rc::new(parse_document_elements_with(&doc_xml, &self.styles));
        self.cache.borrow_mut().elements = Some((self.revision, elements.clone()));
        Ok(elements)
    }

    /// Page setup, headers/footers and background, recomputed only when they can change
    pub fn layout_inputs(&self) -> Result<Rc<LayoutInputs>, String> {
        let doc = self
            .files
            .get("word/document.xml")
            .and_then(|b| std::str::from_utf8(b).ok())
            .ok_or("No se encontró 'word/document.xml' en el archivo DOCX.")?;
        let signature = layout_inputs_signature(doc);
        if let Some((sig, parts, inputs)) = &self.cache.borrow().inputs {
            if *sig == signature && *parts == self.parts_revision {
                return Ok(inputs.clone());
            }
        }
        let inputs = Rc::new(LayoutInputs {
            background_color: self.background_color.clone(),
            page_setup: extract_page_setup_quick_xml(doc),
            header_footer: self.get_header_footer(),
            bg_image_data_url: self.get_bg_image_data_url(),
            body_images: self.body_image_data_urls(doc),
        });
        self.cache.borrow_mut().inputs = Some((signature, self.parts_revision, inputs.clone()));
        Ok(inputs)
    }

    /// Extracts list of paragraphs from word/document.xml
    pub fn extract_paragraphs(&self) -> Result<Vec<ParagraphInfo>, String> {
        let elements = self.extract_elements()?;
        let mut paragraphs = Vec::new();
        for el in elements {
            if let DocumentElement::Paragraph(p) = el {
                paragraphs.push(p);
            }
        }
        Ok(paragraphs)
    }

    /// Extracts tables from word/document.xml
    pub fn extract_tables(&self) -> Result<Vec<TableInfo>, String> {
        let elements = self.extract_elements()?;
        let mut tables = Vec::new();
        for el in elements {
            if let DocumentElement::Table(t) = el {
                tables.push(t);
            }
        }
        Ok(tables)
    }

    /// Extracts full consolidated plain text from the document (including tables)
    pub fn extract_raw_text(&self) -> Result<String, String> {
        let elements = self.extract_elements()?;
        let mut text_parts = Vec::new();

        for el in elements {
            match el {
                DocumentElement::Paragraph(p) => {
                    if !p.text.trim().is_empty() {
                        text_parts.push(p.text);
                    }
                }
                DocumentElement::Table(t) => {
                    for row in t.rows {
                        let row_str = row.join(" | ");
                        if !row_str.trim().is_empty() {
                            text_parts.push(row_str);
                        }
                    }
                }
            }
        }

        Ok(text_parts.join("\n\n"))
    }

    /// Finds and replaces text across word/document.xml, headers, footers and notes
    pub fn find_and_replace(
        &mut self,
        search: &str,
        replacement: &str,
        match_case: bool,
        use_regex: bool,
    ) -> Result<ReplaceResult, String> {
        if search.is_empty() {
            return Err("El término de búsqueda no puede estar vacío.".to_string());
        }

        let rule = ReplaceRule::new(search, replacement, match_case, use_regex)?;
        let (total_replacements, affected_files) = self.apply_replace_rules(&[rule])?;

        Ok(ReplaceResult {
            occurrences_replaced: total_replacements,
            affected_files,
            message: format!(
                "Se reemplazaron exitosamente {} coincidencias.",
                total_replacements
            ),
        })
    }

    /// Replaces multiple key-value pairs (template variables) in a single pass per file
    pub fn batch_replace(&mut self, pairs: &[KeyValuePair]) -> Result<ReplaceResult, String> {
        let rules = pairs
            .iter()
            .filter(|pair| !pair.key.trim().is_empty())
            .map(|pair| ReplaceRule::new(&pair.key, &pair.value, true, false))
            .collect::<Result<Vec<_>, _>>()?;
        let (total_replacements, affected) = self.apply_replace_rules(&rules)?;

        Ok(ReplaceResult {
            occurrences_replaced: total_replacements,
            affected_files: affected,
            message: format!(
                "Reemplazo por lotes completado: {} variables reemplazadas.",
                total_replacements
            ),
        })
    }

    /// Applies replace rules to every XML part under word/, returning (count, affected files)
    fn apply_replace_rules(&mut self, rules: &[ReplaceRule]) -> Result<(usize, Vec<String>), String> {
        let mut target_files: Vec<String> = self
            .files
            .keys()
            .filter(|name| name.starts_with("word/") && name.ends_with(".xml"))
            .cloned()
            .collect();
        target_files.sort();

        let mut total_replacements = 0;
        let mut affected_files = Vec::new();

        for filename in target_files {
            let Ok(xml_str) = self.get_file_string(&filename) else { continue };
            let (new_xml, count) = replace_in_docx_xml(&xml_str, rules)?;
            if count > 0 {
                total_replacements += count;
                affected_files.push(filename.clone());
                self.put_file(filename, new_xml.into_bytes());
            }
        }

        Ok((total_replacements, affected_files))
    }

    /// Updates paragraph texts in word/document.xml by index, keeping their formatting
    pub fn update_paragraphs(&mut self, updates: &[ParagraphUpdate]) -> Result<usize, String> {
        let edits: Vec<ParagraphEdit> = updates
            .iter()
            .map(|u| ParagraphEdit { index: u.index, text: &u.text, formats: None, align: None })
            .collect();
        self.edit_body_paragraphs(&edits)
    }

    /// Replaces characters `start..end` of a body paragraph with `text` (caret editing)
    pub fn replace_paragraph_range(
        &mut self,
        index: usize,
        start: usize,
        end: usize,
        text: &str,
    ) -> Result<bool, String> {
        let mut xml = self.get_file_string("word/document.xml")?;
        let range = body_paragraph_ranges(&xml)?
            .get(index)
            .cloned()
            .ok_or_else(|| format!("No existe el párrafo {}.", index))?;
        let edited = edit_paragraph_range(&xml[range.clone()], &self.styles, start, end, text)?;
        xml.replace_range(range, &edited);
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        Ok(true)
    }

    /// Updates paragraph with individual styled text runs (bold, italic, color, underline per word/segment)
    pub fn update_paragraph_runs(
        &mut self,
        index: usize,
        runs: &[RunInfo],
        align: Option<&str>,
    ) -> Result<bool, String> {
        let text: String = runs.iter().map(|r| r.text.as_str()).collect();
        let formats: Vec<FormatTarget> = runs
            .iter()
            .flat_map(|r| std::iter::repeat(FormatTarget::from_run(r)).take(r.text.chars().count()))
            .collect();
        self.edit_body_paragraphs(&[ParagraphEdit { index, text: &text, formats: Some(formats), align }])?;
        Ok(true)
    }

    /// Updates paragraph formatting (alignment, color, bold, italic) and text
    pub fn update_paragraph_rich(
        &mut self,
        index: usize,
        text: &str,
        align: &str,
        color: &str,
        bold: bool,
        italic: bool,
    ) -> Result<bool, String> {
        let format = FormatTarget {
            bold: Some(bold),
            italic: Some(italic),
            color: Some(color.to_string()),
            ..Default::default()
        };
        let formats = vec![format; text.chars().count()];
        self.edit_body_paragraphs(&[ParagraphEdit { index, text, formats: Some(formats), align: Some(align) }])?;
        Ok(true)
    }

    /// Applies minimal edits to body paragraphs, preserving everything the editor doesn't model
    fn edit_body_paragraphs(&mut self, edits: &[ParagraphEdit]) -> Result<usize, String> {
        let mut xml = self.get_file_string("word/document.xml")?;
        let ranges = body_paragraph_ranges(&xml)?;

        // Splice from the end so earlier byte ranges stay valid
        let mut ordered: Vec<&ParagraphEdit> = edits.iter().collect();
        ordered.sort_by(|a, b| b.index.cmp(&a.index));
        for edit in &ordered {
            let range = ranges
                .get(edit.index)
                .cloned()
                .ok_or_else(|| format!("No existe el párrafo {}.", edit.index))?;
            let edited = edit_paragraph(&xml[range.clone()], &self.styles, edit.text, edit.formats.as_deref(), edit.align)?;
            xml.replace_range(range, &edited);
        }

        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        Ok(ordered.len())
    }

    /// Updates a table cell's text, keeping its paragraphs' properties and run formatting.
    /// Paragraph boundaries outside the edited region survive; new line breaks become `<w:br/>`.
    pub fn update_table_cell(
        &mut self,
        table_index: usize,
        row: usize,
        col: usize,
        new_text: &str,
    ) -> Result<bool, String> {
        let mut xml = self.get_file_string("word/document.xml")?;
        let ranges = table_cell_paragraph_ranges(&xml, table_index, row, col)?;
        let old_lens: Vec<usize> = ranges
            .iter()
            .map(|r| parse_paragraph_fragment(&xml[r.clone()], &self.styles).text.chars().count())
            .collect();
        let old: Vec<char> = ranges
            .iter()
            .map(|r| parse_paragraph_fragment(&xml[r.clone()], &self.styles).text)
            .collect::<Vec<_>>()
            .join("\n")
            .chars()
            .collect();
        let new: Vec<char> = new_text.chars().collect();

        let mut prefix = 0;
        while prefix < old.len() && prefix < new.len() && old[prefix] == new[prefix] {
            prefix += 1;
        }
        let mut suffix = 0;
        while suffix < old.len() - prefix
            && suffix < new.len() - prefix
            && old[old.len() - 1 - suffix] == new[new.len() - 1 - suffix]
        {
            suffix += 1;
        }

        // (paragraph, start of its text in the new cell text)
        let mut kept: Vec<(usize, usize)> = vec![(0, 0)];
        let mut boundary = 0;
        for p in 1..ranges.len() {
            boundary += old_lens[p - 1];
            if boundary < prefix {
                kept.push((p, boundary + 1));
            } else if boundary >= old.len() - suffix {
                kept.push((p, boundary + 1 + new.len() - old.len()));
            }
            boundary += 1;
        }

        let mut edits: Vec<(std::ops::Range<usize>, String)> = Vec::new();
        for (k, &(p, start)) in kept.iter().enumerate() {
            let end = kept.get(k + 1).map(|&(_, s)| s - 1).unwrap_or(new.len());
            let slice: String = new[start..end].iter().collect();
            edits.push((ranges[p].clone(), edit_paragraph(&xml[ranges[p].clone()], &self.styles, &slice, None, None)?));
        }
        // Paragraphs whose boundary fell inside the edited region are merged into the previous one
        for (p, range) in ranges.iter().enumerate() {
            if !kept.iter().any(|&(kp, _)| kp == p) {
                edits.push((range.clone(), String::new()));
            }
        }

        edits.sort_by(|a, b| b.0.start.cmp(&a.0.start));
        for (range, replacement) in edits {
            xml.replace_range(range, &replacement);
        }
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        Ok(true)
    }

    /// Appends a new table to the document
    pub fn add_table(&mut self, rows: usize, cols: usize, headers: &[String]) -> Result<bool, String> {
        let doc_xml = self.get_file_string("word/document.xml")?;
        let new_xml = insert_table_into_xml(&doc_xml, rows, cols, headers)?;
        self.put_file("word/document.xml".to_string(), new_xml.into_bytes());
        Ok(true)
    }

    /// Sets page background color (HEX)
    pub fn set_background_color(&mut self, hex_color: &str) -> Result<(), String> {
        self.parts_revision += 1;
        let clean_hex = hex_color.trim_start_matches('#').to_uppercase();
        self.background_color = clean_hex.clone();

        let doc_xml = self.get_file_string("word/document.xml")?;
        let new_xml = set_bg_color_in_xml(&doc_xml, &clean_hex);
        self.put_file("word/document.xml".to_string(), new_xml.into_bytes());
        Ok(())
    }

    /// Sets background/watermark image bytes
    pub fn set_background_image(&mut self, image_bytes: Vec<u8>, ext: &str) -> Result<(), String> {
        self.parts_revision += 1;
        let clean_ext = if ext.contains("jpg") || ext.contains("jpeg") { "jpeg" } else { "png" };
        let image_filename = format!("word/media/background.{}", clean_ext);
        
        self.put_file(image_filename.clone(), image_bytes);
        if !self.original_order.contains(&image_filename) {
            self.original_order.push(image_filename.clone());
        }

        self.background_image_ext = clean_ext.to_string();

        // Update [Content_Types].xml if needed
        if let Ok(mut types_xml) = self.get_file_string("[Content_Types].xml") {
            let ext_tag = format!("Extension=\"{}\"", clean_ext);
            if !types_xml.contains(&ext_tag) {
                let mime = if clean_ext == "png" { "image/png" } else { "image/jpeg" };
                let default_tag = format!("<Default Extension=\"{}\" ContentType=\"{}\"/>", clean_ext, mime);
                if let Some(pos) = types_xml.find("</Types>") {
                    types_xml.insert_str(pos, &default_tag);
                    self.put_file("[Content_Types].xml".to_string(), types_xml.into_bytes());
                }
            }
        }

        // Update word/_rels/document.xml.rels
        let rels_path = "word/_rels/document.xml.rels";
        let mut rels_xml = self.get_file_string(rels_path).unwrap_or_else(|_| {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.to_string()
        });

        if !rels_xml.contains("media/background") {
            let rel_id = "rIdBgImage";
            let rel_entry = format!(
                r#"<Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/background.{}"/>"#,
                rel_id, clean_ext
            );
            if let Some(pos) = rels_xml.find("</Relationships>") {
                rels_xml.insert_str(pos, &rel_entry);
                self.put_file(rels_path.to_string(), rels_xml.into_bytes());
                if !self.original_order.contains(&rels_path.to_string()) {
                    self.original_order.push(rels_path.to_string());
                }
            }
        }

        Ok(())
    }

    /// Retrieves document statistics
    pub fn get_statistics(&self) -> Result<DocxStats, String> {
        let elements = self.elements_shared()?;
        let mut paragraph_count = 0;
        let mut table_count = 0;
        let mut word_count = 0;
        let mut char_count = 0;

        for el in elements.iter() {
            match el {
                DocumentElement::Paragraph(p) => {
                    paragraph_count += 1;
                    char_count += p.text.chars().count();
                    word_count += p.text.split_whitespace().count();
                }
                DocumentElement::Table(t) => {
                    table_count += 1;
                    for row in &t.rows {
                        for cell in row {
                            char_count += cell.chars().count();
                            word_count += cell.split_whitespace().count();
                        }
                    }
                }
            }
        }

        let inputs = self.layout_inputs()?;
        Ok(DocxStats {
            paragraph_count,
            table_count,
            word_count,
            char_count,
            files_in_zip: self.original_order.clone(),
            original_size_bytes: self.original_size,
            background_color: inputs.background_color.clone(),
            has_background_image: self.background_image.is_some() || inputs.bg_image_data_url.is_some(),
            page_setup: inputs.page_setup.clone(),
            header_footer: inputs.header_footer.clone(),
            bg_image_data_url: inputs.bg_image_data_url.clone(),
        })
    }

    /// Gets parsed page setup (dimensions, margins, orientation) from word/document.xml
    pub fn get_page_setup(&self) -> PageSetup {
        if let Some(bytes) = self.files.get("word/document.xml") {
            if let Ok(xml) = String::from_utf8(bytes.clone()) {
                return extract_page_setup_quick_xml(&xml);
            }
        }
        PageSetup::default()
    }

    /// Gets header and footer info (text, images) from header/footer XML files
    pub fn get_header_footer(&self) -> HeaderFooterInfo {
        extract_header_footer_quick_xml(&self.files)
    }

    /// Pixels of every picture referenced by document.xml, by relationship id
    fn body_image_data_urls(&self, doc: &str) -> HashMap<String, String> {
        let rels = self
            .files
            .get("word/_rels/document.xml.rels")
            .and_then(|b| std::str::from_utf8(b).ok())
            .map(parse_relationships_map)
            .unwrap_or_default();
        extract_drawing_embed_ids(doc)
            .into_iter()
            .filter_map(|id| {
                let url = image_data_url(&self.files, rels.get(&id)?)?;
                Some((id, url))
            })
            .collect()
    }

    /// Data URL of the page background chosen in the editor ("Fondo / Marca de agua")
    pub fn get_bg_image_data_url(&self) -> Option<String> {
        // 1. If explicitly set via UI/API
        if let Some(bytes) = &self.background_image {
            let mime = if self.background_image_ext == "jpeg" || self.background_image_ext == "jpg" {
                "image/jpeg"
            } else {
                "image/png"
            };
            return Some(format!("data:{};base64,{}", mime, BASE64.encode(bytes)));
        }

        None
    }

    /// Exports modified DOCX as bytes
    pub fn to_bytes(&self) -> Result<Vec<u8>, String> {
        let mut buffer = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(&mut buffer);

        let options = SimpleFileOptions::default()
            .compression_method(CompressionMethod::Deflated);

        for name in &self.original_order {
            if let Some(content) = self.files.get(name) {
                zip.start_file(name, options)
                    .map_err(|e| format!("Error al escribir entrada ZIP '{}': {}", name, e))?;
                zip.write_all(content)
                    .map_err(|e| format!("Error al escribir contenido de '{}': {}", name, e))?;
            }
        }

        // Include any new files added
        for (name, content) in &self.files {
            if !self.original_order.contains(name) {
                zip.start_file(name, options)
                    .map_err(|e| format!("Error al escribir nueva entrada ZIP '{}': {}", name, e))?;
                zip.write_all(content)
                    .map_err(|e| format!("Error al escribir contenido de '{}': {}", name, e))?;
            }
        }

        zip.finish()
            .map_err(|e| format!("Error al generar archivo final DOCX: {}", e))?;

        Ok(buffer.into_inner())
    }

    pub fn get_file_string(&self, path: &str) -> Result<String, String> {
        match self.files.get(path) {
            Some(bytes) => String::from_utf8(bytes.clone())
                .map_err(|e| format!("El archivo '{}' no contiene UTF-8 válido: {}", path, e)),
            None => Err(format!("No se encontró '{}' en el archivo DOCX.", path)),
        }
    }
}

// ---------------- Helper XML Functions using quick-xml ----------------

pub fn parse_relationships_map(xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let name = e.name();
                if tag_is(name.as_ref(), "Relationship") {
                    let id_opt = get_attr_value(e, "Id");
                    let target_opt = get_attr_value(e, "Target");
                    if let (Some(id), Some(target)) = (id_opt, target_opt) {
                        map.insert(id, target);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    map
}

pub fn extract_drawing_embed_ids(xml: &str) -> Vec<String> {
    let mut ids = Vec::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let name = e.name();
                if tag_is(name.as_ref(), "blip") || tag_is(name.as_ref(), "imagedata") {
                    if let Some(embed_id) = get_attr_value(e, "embed").or_else(|| get_attr_value(e, "id")) {
                        ids.push(embed_id);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    ids
}

pub fn extract_page_setup_quick_xml(xml: &str) -> PageSetup {
    let mut setup = PageSetup::default();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let name = e.name();
                if tag_is(name.as_ref(), "pgSz") {
                    let w_opt = get_attr_i64(e, "w");
                    let h_opt = get_attr_i64(e, "h");
                    let orient_opt = get_attr_value(e, "orient");

                    if let Some(orient) = orient_opt {
                        setup.orientation = orient;
                    }

                    if let (Some(w), Some(h)) = (w_opt, h_opt) {
                        let w_px = (w as f64) / 15.0;
                        let h_px = (h as f64) / 15.0;

                        if setup.orientation == "landscape" {
                            setup.width = w_px.max(h_px);
                            setup.height = w_px.min(h_px);
                        } else {
                            setup.width = w_px.min(h_px);
                            setup.height = w_px.max(h_px);
                        }
                    } else if let Some(w) = w_opt {
                        setup.width = (w as f64) / 15.0;
                    } else if let Some(h) = h_opt {
                        setup.height = (h as f64) / 15.0;
                    }
                } else if tag_is(name.as_ref(), "pgMar") {
                    if let Some(top) = get_attr_i64(e, "top") {
                        setup.margin_top = ((top as f64) / 15.0).max(20.0);
                    }
                    if let Some(right) = get_attr_i64(e, "right") {
                        setup.margin_right = ((right as f64) / 15.0).max(20.0);
                    }
                    if let Some(bottom) = get_attr_i64(e, "bottom") {
                        setup.margin_bottom = ((bottom as f64) / 15.0).max(20.0);
                    }
                    if let Some(left) = get_attr_i64(e, "left") {
                        setup.margin_left = ((left as f64) / 15.0).max(20.0);
                    }
                    if let Some(header) = get_attr_i64(e, "header") {
                        setup.header_margin = ((header as f64) / 15.0).max(10.0);
                    }
                    if let Some(footer) = get_attr_i64(e, "footer") {
                        setup.footer_margin = ((footer as f64) / 15.0).max(10.0);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    setup
}

pub fn extract_header_footer_quick_xml(files: &HashMap<String, Vec<u8>>) -> HeaderFooterInfo {
    let mut info = HeaderFooterInfo::default();

    // Check headers
    for h_name in &["word/header1.xml", "word/header2.xml", "word/header3.xml"] {
        if let Some(bytes) = files.get(*h_name) {
            if let Ok(xml) = String::from_utf8(bytes.clone()) {
                info.has_header = true;
                let text = extract_text_runs_quick_xml(&xml);
                if !text.trim().is_empty() && info.header_text.is_empty() {
                    info.header_text = text;
                }

                // Check header image via header rels
                if info.header_image_data_url.is_none() {
                    let rel_path = format!("word/_rels/{}.rels", h_name.trim_start_matches("word/"));
                    if let Some(rel_bytes) = files.get(&rel_path) {
                        if let Ok(rel_xml) = String::from_utf8(rel_bytes.clone()) {
                            let rels = parse_relationships_map(&rel_xml);
                            let embed_ids = extract_drawing_embed_ids(&xml);
                            for eid in embed_ids {
                                if let Some(target) = rels.get(&eid) {
                                    let clean_target = target.trim_start_matches("../").trim_start_matches('/');
                                    let img_path = if clean_target.starts_with("media/") {
                                        format!("word/{}", clean_target)
                                    } else if clean_target.starts_with("word/") {
                                        clean_target.to_string()
                                    } else {
                                        format!("word/media/{}", clean_target)
                                    };

                                    if let Some(img_bytes) = files.get(&img_path) {
                                        let mime = if img_path.ends_with(".jpg") || img_path.ends_with(".jpeg") {
                                            "image/jpeg"
                                        } else if img_path.ends_with(".png") {
                                            "image/png"
                                        } else if img_path.ends_with(".svg") {
                                            "image/svg+xml"
                                        } else {
                                            "image/png"
                                        };
                                        let b64 = BASE64.encode(img_bytes);
                                        info.header_image_data_url = Some(format!("data:{};base64,{}", mime, b64));
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Pictures of the default header/footer (the ones shown on regular pages)
    if let Some(part) = default_header_footer_part(files, "header").filter(|p| files.contains_key(p)) {
        info.header_images = placed_images(files, &part);
    }
    if let Some(part) = default_header_footer_part(files, "footer").filter(|p| files.contains_key(p)) {
        info.footer_images = placed_images(files, &part);
    }

    // Check footers
    for f_name in &["word/footer1.xml", "word/footer2.xml", "word/footer3.xml"] {
        if let Some(bytes) = files.get(*f_name) {
            if let Ok(xml) = String::from_utf8(bytes.clone()) {
                info.has_footer = true;
                let text = extract_text_runs_quick_xml(&xml);
                if !text.trim().is_empty() && info.footer_text.is_empty() {
                    info.footer_text = text;
                }
            }
        }
    }

    info
}

fn extract_text_runs_quick_xml(xml: &str) -> String {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut text_parts = Vec::new();
    let mut in_t = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                if tag_is(e.name().as_ref(), "t") {
                    in_t = true;
                }
            }
            Ok(Event::Text(ref e)) => {
                if in_t {
                    if let Ok(unescaped) = e.unescape() {
                        text_parts.push(unescaped.to_string());
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                if tag_is(e.name().as_ref(), "t") {
                    in_t = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    text_parts.join(" ")
}

pub(crate) fn tag_is(tag_bytes: &[u8], name: &str) -> bool {
    if let Ok(s) = std::str::from_utf8(tag_bytes) {
        s == name || s.ends_with(&format!(":{}", name))
    } else {
        false
    }
}

pub(crate) fn get_attr_value(e: &BytesStart, local_name: &str) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        let key_str = std::str::from_utf8(key).unwrap_or("");
        if key_str == local_name || key_str.ends_with(&format!(":{}", local_name)) {
            return attr.unescape_value().ok().map(|s| s.to_string());
        }
    }
    None
}

fn get_attr_i64(e: &BytesStart, local_name: &str) -> Option<i64> {
    get_attr_value(e, local_name).and_then(|v| v.parse::<i64>().ok())
}

pub(crate) fn is_bool_element_true(e: &BytesStart) -> bool {
    if let Some(val) = get_attr_value(e, "val") {
        let v = val.to_lowercase();
        !(v == "0" || v == "false" || v == "off" || v == "none")
    } else {
        true
    }
}

fn extract_bg_color_quick_xml(xml: &str) -> String {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                if tag_is(e.name().as_ref(), "background") {
                    if let Some(col) = get_attr_value(e, "color") {
                        return col;
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    "FFFFFF".to_string()
}

/// Parses `<w:body>` child elements (`<w:p>` and `<w:tbl>`) in sequential document order with quick-xml
pub fn parse_document_elements(xml: &str) -> Vec<DocumentElement> {
    parse_document_elements_with(xml, &StyleSheet::default())
}

/// Like `parse_document_elements`, resolving styles and list numbering
pub fn parse_document_elements_with(xml: &str, styles: &StyleSheet) -> Vec<DocumentElement> {
    let mut counters = NumberingCounters::default();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut elements = Vec::new();
    let mut buf = Vec::new();

    let mut p_index = 0;
    let mut tbl_index = 0;
    let mut in_body = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                if tag_is(name.as_ref(), "body") {
                    in_body = true;
                } else if in_body && tag_is(name.as_ref(), "p") {
                    let p = parse_paragraph_with(&mut reader, p_index, styles, Some(&mut counters));
                    elements.push(DocumentElement::Paragraph(p));
                    p_index += 1;
                } else if in_body && tag_is(name.as_ref(), "tbl") {
                    let tbl = parse_table_with(&mut reader, tbl_index, styles, Some(&mut counters), &mut p_index);
                    elements.push(DocumentElement::Table(tbl));
                    tbl_index += 1;
                }
            }
            Ok(Event::Empty(ref e)) => {
                if in_body && tag_is(e.name().as_ref(), "p") {
                    let p = resolve_paragraph(
                        p_index,
                        styles,
                        None,
                        None,
                        &ParaProps::default(),
                        &RunProps::default(),
                        Vec::new(),
                    );
                    elements.push(DocumentElement::Paragraph(p));
                    p_index += 1;
                }
            }
            Ok(Event::End(ref e)) => {
                if tag_is(e.name().as_ref(), "body") {
                    in_body = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    mark_section_breaks(xml, &mut elements);
    elements
}

/// A section's `w:type` says how *that* section starts, so the break after a section-ending
/// paragraph is decided by the type of the following section (default: next page)
fn mark_section_breaks(xml: &str, elements: &mut [DocumentElement]) {
    let types = section_types(xml);
    let mut section = 0;
    for el in elements.iter_mut() {
        if let DocumentElement::Paragraph(p) = el {
            if p.ends_section {
                section += 1;
                let next = types.get(section).map(String::as_str).unwrap_or("nextPage");
                p.section_break_after = next != "continuous" && next != "nextColumn";
            }
        }
    }
}

/// `w:type` of every `w:sectPr` in document order (paragraph-level ones, then the body one)
fn section_types(xml: &str) -> Vec<String> {
    let mut reader = Reader::from_str(xml);
    let mut types = Vec::new();
    let mut in_sect = false;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if tag_is(e.name().as_ref(), "sectPr") => {
                in_sect = true;
                types.push("nextPage".to_string());
            }
            Ok(Event::Empty(ref e)) if tag_is(e.name().as_ref(), "sectPr") => {
                types.push("nextPage".to_string());
            }
            Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) if in_sect && tag_is(e.name().as_ref(), "type") => {
                if let (Some(last), Some(v)) = (types.last_mut(), get_attr_value(e, "val")) {
                    *last = v;
                }
            }
            Ok(Event::End(ref e)) if tag_is(e.name().as_ref(), "sectPr") => in_sect = false,
            // Tracked section changes hold the old properties
            Ok(Event::Start(ref e)) if tag_is(e.name().as_ref(), "sectPrChange") => {
                let end = e.name().as_ref().to_vec();
                let _ = reader.read_to_end_into(quick_xml::name::QName(&end), &mut Vec::new());
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    types
}

/// Parses a paragraph without a style sheet (direct formatting only)
pub fn parse_paragraph_from_reader(reader: &mut Reader<&[u8]>, index: usize) -> ParagraphInfo {
    parse_paragraph_with(reader, index, &StyleSheet::default(), None)
}

/// Parses the paragraph whose `<w:p>` start tag was just read, resolving the style cascade.
/// `counters` advances list numbering; pass `None` for isolated fragments.
pub fn parse_paragraph_with(
    reader: &mut Reader<&[u8]>,
    index: usize,
    styles: &StyleSheet,
    counters: Option<&mut NumberingCounters>,
) -> ParagraphInfo {
    let mut style_id: Option<String> = None;
    let mut direct_ppr = ParaProps::default();
    let mut mark_rpr = RunProps::default();
    let mut raw_runs: Vec<RawRun> = Vec::new();
    let mut images: Vec<ImageRef> = Vec::new();
    let mut ends_section = false;

    let mut in_ppr = false;
    let mut in_ppr_rpr = false;
    let mut in_pbdr = false;
    let mut in_r = false;
    let mut in_rpr = false;
    let mut in_t = false;
    let mut run = RawRun::default();
    let mut buf = Vec::new();

    loop {
        let (e, is_start) = match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => (e, true),
            Ok(Event::Empty(e)) => (e, false),
            Ok(Event::Text(ref t)) => {
                if in_t {
                    if let Ok(s) = t.unescape() {
                        run.text.push_str(&s);
                    }
                }
                buf.clear();
                continue;
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                let n = name.as_ref();
                if tag_is(n, "t") {
                    in_t = false;
                } else if tag_is(n, "rPr") {
                    if in_r {
                        in_rpr = false;
                    } else {
                        in_ppr_rpr = false;
                    }
                } else if tag_is(n, "r") && in_r {
                    in_r = false;
                    if !run.text.is_empty() {
                        raw_runs.push(std::mem::take(&mut run));
                    }
                } else if tag_is(n, "pBdr") {
                    in_pbdr = false;
                } else if tag_is(n, "pPr") {
                    in_ppr = false;
                } else if tag_is(n, "p") {
                    break;
                }
                buf.clear();
                continue;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {
                buf.clear();
                continue;
            }
        };

        let name = e.name();
        let n = name.as_ref();
        if is_start && tag_is(n, "drawing") {
            buf.clear();
            if let Some(img) = parse_drawing(reader) {
                images.push(img);
            }
            continue;
        }
        // VML/objects may hold text boxes with their own paragraphs, and tracked property
        // changes hold the *old* properties: none of them belong to this paragraph
        if is_start
            && ["pict", "object", "txbxContent", "pPrChange", "rPrChange"]
                .iter()
                .any(|t| tag_is(n, t))
        {
            let end_name = n.to_vec();
            let mut skip_buf = Vec::new();
            let _ = reader.read_to_end_into(quick_xml::name::QName(&end_name), &mut skip_buf);
            buf.clear();
            continue;
        }

        if in_r {
            if in_rpr {
                if tag_is(n, "rStyle") {
                    run.style = get_attr_value(&e, "val");
                } else {
                    apply_rpr_element(&mut run.props, &e, &styles.theme);
                }
            } else if tag_is(n, "rPr") {
                in_rpr = is_start;
            } else if tag_is(n, "t") {
                in_t = is_start;
            } else if tag_is(n, "tab") {
                run.text.push('\t');
            } else if tag_is(n, "br") {
                run.text.push(break_char(&e));
            } else if tag_is(n, "cr") {
                run.text.push('\n');
            }
        } else if in_ppr {
            if tag_is(n, "sectPr") {
                ends_section = true;
                if is_start {
                    let end_name = n.to_vec();
                    let mut skip_buf = Vec::new();
                    let _ = reader.read_to_end_into(quick_xml::name::QName(&end_name), &mut skip_buf);
                }
            } else if in_ppr_rpr {
                apply_rpr_element(&mut mark_rpr, &e, &styles.theme);
            } else if in_pbdr {
                apply_border_element(&mut direct_ppr.borders, &e);
            } else if tag_is(n, "rPr") {
                in_ppr_rpr = is_start;
            } else if tag_is(n, "pBdr") {
                in_pbdr = is_start;
            } else if tag_is(n, "pStyle") {
                style_id = get_attr_value(&e, "val");
            } else {
                apply_ppr_element(&mut direct_ppr, &e);
            }
        } else if tag_is(n, "pPr") {
            in_ppr = is_start;
        } else if tag_is(n, "r") && is_start {
            in_r = true;
            run = RawRun::default();
        }
        buf.clear();
    }

    let mut info = resolve_paragraph(index, styles, counters, style_id, &direct_ppr, &mark_rpr, raw_runs);
    info.ends_section = ends_section;
    info.images = images;
    info
}

/// Character used in paragraph text for a `w:br`: form feed for page breaks, newline otherwise
pub(crate) fn break_char(e: &BytesStart) -> char {
    if get_attr_value(e, "type").as_deref() == Some("page") {
        PAGE_BREAK
    } else {
        '\n'
    }
}

/// Paragraph text marker for a manual page break (`<w:br w:type="page"/>`)
pub const PAGE_BREAK: char = '\u{000C}';

/// A run as written in the XML, before the style cascade is applied
#[derive(Default)]
struct RawRun {
    text: String,
    props: RunProps,
    style: Option<String>,
}

fn to_run_info(text: String, p: &RunProps) -> RunInfo {
    RunInfo {
        text,
        bold: p.bold.unwrap_or(false),
        italic: p.italic.unwrap_or(false),
        underline: p.underline.unwrap_or(false),
        color: p.color.clone().unwrap_or_default(),
        font_size: p.font_size,
        font_family: p.font_family.clone(),
    }
}

/// Applies document defaults, numbering, paragraph/character styles and direct formatting
fn resolve_paragraph(
    index: usize,
    styles: &StyleSheet,
    counters: Option<&mut NumberingCounters>,
    style_id: Option<String>,
    direct_ppr: &ParaProps,
    mark_rpr: &RunProps,
    raw_runs: Vec<RawRun>,
) -> ParagraphInfo {
    let effective_style = styles.effective_paragraph_style(style_id.as_deref()).map(str::to_string);
    let (style_ppr, style_rpr) = styles.style_props(effective_style.as_deref());

    let num_id = direct_ppr
        .num_id
        .clone()
        .or_else(|| style_ppr.num_id.clone())
        .filter(|id| id != "0");
    let ilvl = direct_ppr.ilvl.or(style_ppr.ilvl).unwrap_or(0);

    let mut ppr = styles.doc_ppr.clone();
    if let Some(level) = num_id.as_deref().and_then(|id| styles.numbering.level(id, ilvl)) {
        ppr.merge(&level.ppr);
    }
    ppr.merge(&style_ppr);
    ppr.merge(direct_ppr);

    let base_rpr = styles.doc_rpr.merged(&style_rpr);
    let mark = base_rpr.merged(mark_rpr);
    let runs: Vec<RunInfo> = raw_runs
        .into_iter()
        .map(|r| {
            let eff = base_rpr.merged(&styles.character_props(r.style.as_deref())).merged(&r.props);
            to_run_info(r.text, &eff)
        })
        .collect();

    let list_label = match (num_id.as_deref(), counters) {
        (Some(id), Some(counters)) => counters
            .next_label(&styles.numbering, id, ilvl)
            .filter(|(label, _)| !label.is_empty())
            .map(|(label, level)| {
                let text = if level.suffix == "space" { format!("{} ", label) } else { label };
                let mut run = to_run_info(text, &mark.merged(&level.rpr));
                if is_symbol_font(level.rpr.font_family.as_deref()) {
                    run.font_family = mark.font_family.clone();
                }
                run
            }),
        _ => None,
    };

    let style_name = effective_style
        .as_deref()
        .and_then(|id| styles.style_name(id))
        .unwrap_or("")
        .to_lowercase();
    let style_key = style_id.clone().or(effective_style).unwrap_or_else(|| "Normal".to_string());
    let is_heading = ppr.outline_level.is_some_and(|l| l < 9)
        || [style_name.as_str(), style_key.to_lowercase().as_str()].iter().any(|s| {
            s.contains("heading") || s.contains("title") || s.contains("encabezado") || s.contains("título")
        });

    let full_text: String = runs.iter().map(|r| r.text.as_str()).collect();
    let non_empty_runs: Vec<&RunInfo> = runs.iter().filter(|r| !r.text.trim().is_empty()).collect();
    let bold = if !styles.loaded && is_heading {
        true
    } else if !non_empty_runs.is_empty() {
        non_empty_runs.iter().all(|r| r.bold)
    } else {
        mark.bold.unwrap_or(false)
    };
    let italic = if !non_empty_runs.is_empty() {
        non_empty_runs.iter().all(|r| r.italic)
    } else {
        mark.italic.unwrap_or(false)
    };

    ParagraphInfo {
        index,
        text: full_text,
        style: style_key,
        is_heading,
        run_count: runs.len(),
        align: ppr.align.clone().unwrap_or_else(|| "left".to_string()),
        color: mark.color.clone().unwrap_or_default(),
        bold,
        italic,
        font_size: mark.font_size,
        font_family: mark.font_family.clone(),
        indent_left: ppr.indent_left.unwrap_or(0.0),
        indent_first_line: ppr.indent_first_line.unwrap_or(0.0),
        indent_right: ppr.indent_right.unwrap_or(0.0),
        space_before: ppr.space_before.unwrap_or(0.0),
        space_after: ppr.space_after.unwrap_or(0.0),
        line_spacing: ppr.line_spacing,
        borders: ppr.borders.clone(),
        runs,
        list_label,
        line_exact: ppr.line_pt.filter(|_| ppr.line_rule.as_deref() == Some("exact")),
        line_at_least: ppr.line_pt.filter(|_| ppr.line_rule.as_deref() == Some("atLeast")),
        keep_next: ppr.keep_next.unwrap_or(false),
        keep_lines: ppr.keep_lines.unwrap_or(false),
        page_break_before: ppr.page_break_before.unwrap_or(false),
        // Omitted widowControl means "on" (Word writes w:val="0" to disable it)
        widow_control: ppr.widow_control.unwrap_or(true),
        contextual_spacing: ppr.contextual_spacing.unwrap_or(false),
        spacing_resolved: styles.loaded || ppr.space_before.is_some() || ppr.space_after.is_some(),
        section_break_after: false,
        ends_section: false,
        images: Vec::new(),
    }
}

pub(crate) fn parse_border_element(e: &BytesStart) -> Option<BorderInfo> {
    let val = get_attr_value(e, "val").unwrap_or_default().to_lowercase();
    if val.is_empty() || val == "none" || val == "nil" || val == "off" || val == "0" {
        return None;
    }
    let color = get_attr_value(e, "color").unwrap_or_default();
    let sz_raw = get_attr_i64(e, "sz").unwrap_or(4) as f64;
    let sz_px = ((sz_raw / 8.0) * 1.3333).clamp(0.75, 8.0);
    let space = get_attr_i64(e, "space").unwrap_or(0) as f64;
    Some(BorderInfo {
        val,
        color,
        sz_px,
        space,
    })
}

pub fn is_dark_hex_str(hex: &str) -> bool {
    let clean = hex.trim_start_matches('#');
    if clean.len() == 6 {
        if let (Ok(r), Ok(g), Ok(b)) = (
            u8::from_str_radix(&clean[0..2], 16),
            u8::from_str_radix(&clean[2..4], 16),
            u8::from_str_radix(&clean[4..6], 16),
        ) {
            let lum = 0.299 * (r as f64) + 0.587 * (g as f64) + 0.114 * (b as f64);
            return lum < 140.0;
        }
    }
    false
}

pub fn parse_table_from_reader(reader: &mut Reader<&[u8]>, index: usize) -> TableInfo {
    parse_table_with(reader, index, &StyleSheet::default(), None, &mut 0)
}

/// Parses the table whose `<w:tbl>` start tag was just read, resolving cell paragraph styles
pub fn parse_table_with(
    reader: &mut Reader<&[u8]>,
    index: usize,
    styles: &StyleSheet,
    mut counters: Option<&mut NumberingCounters>,
    p_index: &mut usize,
) -> TableInfo {
    let mut rows: Vec<Vec<String>> = Vec::new();
    let mut rich_rows: Vec<TableRowData> = Vec::new();
    let mut grid_cols: Vec<f64> = Vec::new();

    let mut current_row_strings: Vec<String> = Vec::new();
    let mut current_rich_cells: Vec<TableCellData> = Vec::new();
    let mut current_cell_paragraphs: Vec<ParagraphInfo> = Vec::new();

    let mut table_border_color: Option<String> = None;
    let mut current_cell_bg: Option<String> = None;
    let mut current_cell_border: Option<String> = None;
    let mut current_row_is_header = false;

    let mut in_tbl_pr = false;
    let mut in_tbl_borders = false;
    let mut in_tbl_grid = false;
    let mut in_tr = false;
    let mut in_tr_pr = false;
    let mut in_tc = false;
    let mut in_tc_pr = false;
    let mut in_tc_borders = false;
    let mut row_idx = 0;

    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                if tag_is(name.as_ref(), "tblPr") {
                    in_tbl_pr = true;
                } else if in_tbl_pr && tag_is(name.as_ref(), "tblBorders") {
                    in_tbl_borders = true;
                } else if in_tbl_borders {
                    if let Some(col) = get_attr_value(e, "color") {
                        if !col.is_empty() && col.to_lowercase() != "auto" && col.to_lowercase() != "none" {
                            table_border_color = Some(col);
                        }
                    }
                } else if tag_is(name.as_ref(), "tblGrid") {
                    in_tbl_grid = true;
                } else if tag_is(name.as_ref(), "tr") {
                    in_tr = true;
                    current_row_strings = Vec::new();
                    current_rich_cells = Vec::new();
                    current_row_is_header = row_idx == 0;
                } else if in_tr && tag_is(name.as_ref(), "trPr") {
                    in_tr_pr = true;
                } else if in_tr && tag_is(name.as_ref(), "tc") {
                    in_tc = true;
                    current_cell_paragraphs = Vec::new();
                    current_cell_bg = None;
                    current_cell_border = None;
                } else if in_tc && tag_is(name.as_ref(), "tcPr") {
                    in_tc_pr = true;
                } else if in_tc_pr && tag_is(name.as_ref(), "tcBorders") {
                    in_tc_borders = true;
                } else if in_tc_borders {
                    if let Some(col) = get_attr_value(e, "color") {
                        if !col.is_empty() && col.to_lowercase() != "auto" && col.to_lowercase() != "none" {
                            current_cell_border = Some(col);
                        }
                    }
                } else if in_tc && tag_is(name.as_ref(), "p") {
                    let p = parse_paragraph_with(reader, *p_index, styles, counters.as_deref_mut());
                    *p_index += 1;
                    current_cell_paragraphs.push(p);
                }
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                if in_tc && tag_is(name.as_ref(), "p") {
                    current_cell_paragraphs.push(resolve_paragraph(
                        *p_index,
                        styles,
                        None,
                        None,
                        &ParaProps::default(),
                        &RunProps::default(),
                        Vec::new(),
                    ));
                    *p_index += 1;
                }
                if in_tbl_borders {
                    if let Some(col) = get_attr_value(e, "color") {
                        if !col.is_empty() && col.to_lowercase() != "auto" && col.to_lowercase() != "none" {
                            table_border_color = Some(col);
                        }
                    }
                } else if in_tbl_grid && tag_is(name.as_ref(), "gridCol") {
                    if let Some(w) = get_attr_i64(e, "w") {
                        grid_cols.push(w as f64);
                    }
                } else if in_tr_pr && tag_is(name.as_ref(), "tblHeader") {
                    current_row_is_header = true;
                } else if in_tc_pr && tag_is(name.as_ref(), "shd") {
                    if let Some(fill) = get_attr_value(e, "fill") {
                        let f_low = fill.to_lowercase();
                        if !f_low.is_empty() && f_low != "auto" && f_low != "clear" && f_low != "none" && f_low != "ffffff" {
                            current_cell_bg = Some(fill);
                        }
                    }
                } else if in_tc_borders {
                    if let Some(col) = get_attr_value(e, "color") {
                        if !col.is_empty() && col.to_lowercase() != "auto" && col.to_lowercase() != "none" {
                            current_cell_border = Some(col);
                        }
                    }
                }
            }
            Ok(Event::End(ref e)) => {
                let name = e.name();
                if tag_is(name.as_ref(), "tblBorders") {
                    in_tbl_borders = false;
                } else if tag_is(name.as_ref(), "tblPr") {
                    in_tbl_pr = false;
                } else if tag_is(name.as_ref(), "tblGrid") {
                    in_tbl_grid = false;
                } else if tag_is(name.as_ref(), "trPr") {
                    in_tr_pr = false;
                } else if tag_is(name.as_ref(), "tcBorders") {
                    in_tc_borders = false;
                } else if tag_is(name.as_ref(), "tcPr") {
                    in_tc_pr = false;
                } else if in_tc && tag_is(name.as_ref(), "tc") {
                    in_tc = false;
                    let full_text = current_cell_paragraphs
                        .iter()
                        .map(|p| p.text.as_str())
                        .collect::<Vec<_>>()
                        .join("\n");

                    let first_p = current_cell_paragraphs.first();
                    let align = first_p.map(|p| p.align.clone()).unwrap_or_else(|| "left".to_string());

                    // Check explicit text color from runs or paragraph
                    let explicit_color = first_p
                        .and_then(|p| p.runs.iter().find(|r| !r.color.is_empty()).map(|r| r.color.clone()))
                        .or_else(|| first_p.map(|p| p.color.clone()).filter(|c| !c.is_empty()));

                    let is_dark_bg = current_cell_bg.as_deref().map(is_dark_hex_str).unwrap_or(false);
                    let color = if let Some(c) = explicit_color {
                        c
                    } else if is_dark_bg {
                        "FAF7F0".to_string()
                    } else if current_row_is_header && current_cell_bg.is_some() {
                        "FAF7F0".to_string()
                    } else {
                        "1B1F1E".to_string()
                    };

                    let bold = first_p
                        .map(|p| p.bold || p.runs.iter().any(|r| r.bold))
                        .unwrap_or(current_row_is_header);
                    let italic = first_p
                        .map(|p| p.italic || p.runs.iter().any(|r| r.italic))
                        .unwrap_or(false);
                    let font_size = first_p
                        .and_then(|p| p.font_size.or_else(|| p.runs.first().and_then(|r| r.font_size)))
                        .unwrap_or(if current_row_is_header { 10.0 } else { 9.5 });
                    let font_family = first_p
                        .and_then(|p| p.font_family.clone().or_else(|| p.runs.first().and_then(|r| r.font_family.clone())))
                        .unwrap_or_else(|| "Calibri, Inter, sans-serif".to_string());

                    let border_color = current_cell_border
                        .clone()
                        .or_else(|| table_border_color.clone())
                        .unwrap_or_else(|| "DDD5C2".to_string());

                    current_row_strings.push(full_text.clone());
                    current_rich_cells.push(TableCellData {
                        text: full_text,
                        bg_color: current_cell_bg.clone(),
                        color,
                        align,
                        bold,
                        italic,
                        font_size,
                        font_family,
                        border_color,
                        paragraphs: current_cell_paragraphs.clone(),
                    });
                } else if in_tr && tag_is(name.as_ref(), "tr") {
                    in_tr = false;
                    rows.push(current_row_strings.clone());
                    rich_rows.push(TableRowData {
                        cells: current_rich_cells.clone(),
                        is_header: current_row_is_header,
                    });
                    row_idx += 1;
                } else if tag_is(name.as_ref(), "tbl") {
                    break;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    TableInfo {
        index,
        rows,
        rich_rows,
        grid_cols,
        header_row: true,
    }
}

/// A search pattern plus its replacement, applied to the logical text of a paragraph.
pub struct ReplaceRule {
    regex: Regex,
    replacement: String,
    /// Regex mode: expand `$1`/`${name}` capture references in the replacement
    expand: bool,
}

impl ReplaceRule {
    pub fn new(search: &str, replacement: &str, match_case: bool, use_regex: bool) -> Result<Self, String> {
        let body = if use_regex { search.to_string() } else { regex::escape(search) };
        let pattern = if match_case { body } else { format!("(?i){}", body) };
        let regex = Regex::new(&pattern).map_err(|e| format!("Expresión regular inválida: {}", e))?;
        Ok(ReplaceRule {
            regex,
            replacement: replacement.to_string(),
            expand: use_regex,
        })
    }

    /// Returns (start, end, replacement) for every non-overlapping match in `text`
    fn find_matches(&self, text: &str) -> Vec<(usize, usize, String)> {
        if self.expand {
            self.regex
                .captures_iter(text)
                .map(|caps| {
                    let m = caps.get(0).unwrap();
                    let mut out = String::new();
                    caps.expand(&self.replacement, &mut out);
                    (m.start(), m.end(), out)
                })
                .collect()
        } else {
            self.regex
                .find_iter(text)
                .map(|m| (m.start(), m.end(), self.replacement.clone()))
                .collect()
        }
    }
}

/// A `<w:t>` element: event indices of its Start/End tags and its current text
struct TextNode {
    start: usize,
    end: usize,
    text: String,
    modified: bool,
}

/// Elements that interrupt the visible text flow: a match may never span across them
const TEXT_FLOW_BREAKS: &[&str] = &[
    "p", "tab", "ptab", "br", "cr", "drawing", "pict", "object", "sym", "fldChar",
    "instrText", "footnoteReference", "endnoteReference", "commentReference",
    "noBreakHyphen", "softHyphen", "delText", "txbxContent", "tc",
];

fn is_text_flow_break(tag: &[u8]) -> bool {
    TEXT_FLOW_BREAKS.iter().any(|name| tag_is(tag, name))
}

/// Replaces text across run boundaries. Word frequently splits a placeholder like
/// `{name}` into several runs (`{` | `name` | `}`), so matching is done on the joined
/// text of each contiguous `<w:t>` sequence. The replacement goes into the run where
/// the match starts (keeping its formatting) and the matched remainder is removed
/// from the following runs.
fn replace_in_docx_xml(xml: &str, rules: &[ReplaceRule]) -> Result<(String, usize), String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut events: Vec<Event> = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(ev) => events.push(ev),
            Err(e) => return Err(format!("XML error: {:?}", e)),
        }
    }

    // 1. Collect <w:t> nodes, grouped into contiguous text sequences
    let mut nodes: Vec<TextNode> = Vec::new();
    let mut groups: Vec<Vec<usize>> = Vec::new();
    let mut current_group: Vec<usize> = Vec::new();
    let mut open_t: Option<(usize, String)> = None;

    for (i, ev) in events.iter().enumerate() {
        match ev {
            Event::Start(e) if open_t.is_none() && tag_is(e.name().as_ref(), "t") => {
                open_t = Some((i, String::new()));
            }
            Event::Text(t) => {
                if let Some((_, ref mut text)) = open_t {
                    text.push_str(&t.unescape().map_err(|e| e.to_string())?);
                }
            }
            Event::CData(c) => {
                if let Some((_, ref mut text)) = open_t {
                    text.push_str(&String::from_utf8_lossy(c));
                }
            }
            Event::End(e) if open_t.is_some() && tag_is(e.name().as_ref(), "t") => {
                let (start, text) = open_t.take().unwrap();
                current_group.push(nodes.len());
                nodes.push(TextNode { start, end: i, text, modified: false });
            }
            Event::Start(e) | Event::Empty(e) if is_text_flow_break(e.name().as_ref()) => {
                if !current_group.is_empty() {
                    groups.push(std::mem::take(&mut current_group));
                }
            }
            Event::End(e) if is_text_flow_break(e.name().as_ref()) => {
                if !current_group.is_empty() {
                    groups.push(std::mem::take(&mut current_group));
                }
            }
            _ => {}
        }
    }
    if !current_group.is_empty() {
        groups.push(current_group);
    }

    // 2. Apply every rule to the joined text of each group and redistribute
    let mut total_replacements = 0;
    for group in &groups {
        for rule in rules {
            let mut full = String::new();
            let mut bounds = Vec::with_capacity(group.len());
            for &n in group {
                let a = full.len();
                full.push_str(&nodes[n].text);
                bounds.push((a, full.len()));
            }

            let matches = rule.find_matches(&full);
            if matches.is_empty() {
                continue;
            }
            total_replacements += matches.len();

            let mut new_texts = vec![String::new(); group.len()];
            let copy_range = |from: usize, to: usize, out: &mut Vec<String>| {
                for (k, &(a, b)) in bounds.iter().enumerate() {
                    let (lo, hi) = (a.max(from), b.min(to));
                    if lo < hi {
                        out[k].push_str(&full[lo..hi]);
                    }
                }
            };

            let mut pos = 0;
            for (start, end, replacement) in &matches {
                copy_range(pos, *start, &mut new_texts);
                let owner = bounds
                    .iter()
                    .position(|&(a, b)| *start >= a && *start < b)
                    .unwrap_or(bounds.len() - 1);
                new_texts[owner].push_str(replacement);
                pos = *end;
            }
            copy_range(pos, full.len(), &mut new_texts);

            for (k, &n) in group.iter().enumerate() {
                if nodes[n].text != new_texts[k] {
                    nodes[n].text = std::mem::take(&mut new_texts[k]);
                    nodes[n].modified = true;
                }
            }
        }
    }

    if total_replacements == 0 {
        return Ok((xml.to_string(), 0));
    }

    // 3. Write events back, rewriting modified <w:t> nodes
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let modified: HashMap<usize, &TextNode> =
        nodes.iter().filter(|n| n.modified).map(|n| (n.start, n)).collect();

    let mut i = 0;
    while i < events.len() {
        if let (Some(node), Event::Start(e)) = (modified.get(&i), &events[i]) {
            let mut t_start = BytesStart::new(String::from_utf8_lossy(e.name().as_ref()).into_owned());
            for attr in e.attributes().flatten() {
                if attr.key.as_ref() != b"xml:space" {
                    t_start.push_attribute(attr);
                }
            }
            t_start.push_attribute(("xml:space", "preserve"));
            writer.write_event(Event::Start(t_start)).map_err(|e| e.to_string())?;
            if !node.text.is_empty() {
                writer.write_event(Event::Text(BytesText::new(&node.text))).map_err(|e| e.to_string())?;
            }
            writer.write_event(events[node.end].clone()).map_err(|e| e.to_string())?;
            i = node.end + 1;
            continue;
        }
        writer.write_event(events[i].clone()).map_err(|e| e.to_string())?;
        i += 1;
    }

    let bytes = writer.into_inner().into_inner();
    let result_str = String::from_utf8(bytes).map_err(|e| e.to_string())?;
    Ok((result_str, total_replacements))
}

/// Sets `<w:background>` color in document XML
fn set_bg_color_in_xml(xml: &str, hex: &str) -> String {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                let name = e.name();
                if tag_is(name.as_ref(), "background") {
                    let mut new_bg = BytesStart::new("w:background");
                    new_bg.push_attribute(("w:color", hex));
                    let _ = writer.write_event(Event::Empty(new_bg));
                    buf.clear();
                    continue;
                } else if tag_is(name.as_ref(), "document") {
                    let _ = writer.write_event(Event::Start(e.clone()));
                    if !xml.contains("<w:background") {
                        let mut new_bg = BytesStart::new("w:background");
                        new_bg.push_attribute(("w:color", hex));
                        let _ = writer.write_event(Event::Empty(new_bg));
                    }
                    buf.clear();
                    continue;
                }
                let _ = writer.write_event(Event::Start(e.clone()));
            }
            Ok(Event::Empty(ref e)) => {
                let name = e.name();
                if tag_is(name.as_ref(), "background") {
                    let mut new_bg = BytesStart::new("w:background");
                    new_bg.push_attribute(("w:color", hex));
                    let _ = writer.write_event(Event::Empty(new_bg));
                    buf.clear();
                    continue;
                }
                let _ = writer.write_event(Event::Empty(e.clone()));
            }
            Ok(Event::Eof) => break,
            Ok(event) => {
                let _ = writer.write_event(event);
            }
            Err(_) => break,
        }
        buf.clear();
    }

    let bytes = writer.into_inner().into_inner();
    String::from_utf8(bytes).unwrap_or_else(|_| xml.to_string())
}

/// Inserts a new table into document XML
fn insert_table_into_xml(
    xml: &str,
    rows_count: usize,
    cols_count: usize,
    headers: &[String],
) -> Result<String, String> {
    let mut table_xml = String::new();
    table_xml.push_str("<w:tbl>");
    table_xml.push_str(r#"<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/><w:left w:val="none"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="94A3B8"/><w:right w:val="none"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr>"#);

    for r in 0..rows_count {
        table_xml.push_str("<w:tr>");
        for c in 0..cols_count {
            let is_header = r == 0;
            let val = if is_header && c < headers.len() {
                &headers[c]
            } else if is_header {
                "Columna"
            } else {
                "Dato"
            };

            table_xml.push_str("<w:tc><w:p>");
            if is_header {
                table_xml.push_str("<w:r><w:rPr><w:b/><w:color w:val=\"1E3A8A\"/></w:rPr>");
            } else {
                table_xml.push_str("<w:r>");
            }
            table_xml.push_str(&format!("<w:t xml:space=\"preserve\">{}</w:t></w:r></w:p></w:tc>", escape_xml(val)));
        }
        table_xml.push_str("</w:tr>");
    }

    table_xml.push_str("</w:tbl>");

    if let Some(pos) = xml.rfind("</w:body>") {
        let (before, after) = xml.split_at(pos);
        Ok(format!("{}{}{}", before, table_xml, after))
    } else {
        Ok(format!("{}{}", xml, table_xml))
    }
}

fn escape_xml(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::sample_generator::generate_sample_docx;

    #[test]
    fn test_sample_rich_docx_parsing_and_elements() {
        let sample_bytes = generate_sample_docx().expect("Should generate sample docx");
        let mut modifier = DocxModifier::from_bytes(&sample_bytes).expect("Should parse docx");

        let elements = modifier.extract_elements().expect("Should extract elements");
        assert!(elements.len() >= 5);

        // Verify tables exist in sample
        let tables = modifier.extract_tables().expect("Should extract tables");
        assert!(!tables.is_empty(), "Sample docx should have tables");
        assert_eq!(tables[0].rows[0][0], "Hito / Fase");

        // Test Table cell update with multiline text
        modifier.update_table_cell(0, 1, 0, "Fase 1: Nueva Arquitectura\n(Detalle Técnico)").expect("Should update cell");

        // Test Paragraph rich update (alignment and color)
        modifier.update_paragraph_rich(0, "CONTRATO MODIFICADO", "center", "DC2626", true, false)
            .expect("Should update paragraph formatting");

        // Test Page background
        modifier.set_background_color("F0F9FF").expect("Should set background");

        // Export bytes and re-parse
        let exported_bytes = modifier.to_bytes().expect("Should export bytes");
        let reloaded = DocxModifier::from_bytes(&exported_bytes).expect("Should reload docx");
        let raw_text = reloaded.extract_raw_text().expect("Should get raw text");
        assert!(raw_text.contains("CONTRATO MODIFICADO"));
        assert!(raw_text.contains("Fase 1: Nueva Arquitectura"));

        let reloaded_tables = reloaded.extract_tables().expect("Should reload tables");
        assert!(reloaded_tables[0].rows[1][0].contains("\n(Detalle Técnico)"));
    }

    #[test]
    fn test_bold_bleeding_fix() {
        let xml = r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>1. OBJETO:</w:t></w:r><w:r><w:t> El presente contrato regula los servicios.</w:t></w:r></w:p>"#;
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(false);
        let p = parse_paragraph_from_reader(&mut reader, 0);

        assert!(!p.bold, "Paragraph should NOT be marked bold if only first run is bold");
        assert_eq!(p.runs.len(), 2);
        assert!(p.runs[0].bold, "First run should be bold");
        assert!(!p.runs[1].bold, "Second run should NOT be bold");
        assert_eq!(p.text, "1. OBJETO: El presente contrato regula los servicios.");
    }

    #[test]
    fn test_paragraph_mark_formatting_does_not_leak_into_runs() {
        let xml = r#"<w:p><w:pPr><w:rPr><w:b/><w:sz w:val="40"/></w:rPr></w:pPr><w:r><w:t>normal</w:t></w:r></w:p>"#;
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(false);
        let p = parse_paragraph_from_reader(&mut reader, 0);
        assert!(!p.runs[0].bold, "pPr/rPr only formats the paragraph mark");
        assert_eq!(p.runs[0].font_size, None);
        assert_eq!(p.font_size, Some(20.0), "the mark size is still reported for empty-line height");
    }

    #[test]
    fn test_indents_and_tabs_parsing() {
        let xml = r#"<w:p><w:pPr><w:ind w:left="1440" w:firstLine="720"/></w:pPr><w:r><w:t>Primer run</w:t><w:tab/><w:t>Segundo run</w:t></w:r></w:p>"#;
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(false);
        let p = parse_paragraph_from_reader(&mut reader, 0);

        assert!((p.indent_left - 96.0).abs() < 0.01);
        assert!((p.indent_first_line - 48.0).abs() < 0.01);
        assert!(p.text.contains('\t'), "Text should contain tab character");
        assert_eq!(p.text, "Primer run\tSegundo run");
    }

    #[test]
    fn test_multiline_table_cell_parsing() {
        let tbl_xml = r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Línea 1</w:t></w:r></w:p><w:p><w:r><w:t>Línea 2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#;
        let mut reader = Reader::from_str(tbl_xml);
        reader.config_mut().trim_text(false);
        let tbl = parse_table_from_reader(&mut reader, 0);
        assert_eq!(tbl.rows.len(), 1);
        assert_eq!(tbl.rows[0].len(), 1);
        assert_eq!(tbl.rows[0][0], "Línea 1\nLínea 2");
    }

    fn wrap_body(inner: &str) -> String {
        format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{}</w:body></w:document>"#,
            inner
        )
    }

    fn literal(search: &str, replacement: &str) -> ReplaceRule {
        ReplaceRule::new(search, replacement, true, false).unwrap()
    }

    #[test]
    fn test_replace_placeholder_split_across_runs() {
        let xml = wrap_body(
            r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hola {</w:t></w:r><w:r><w:t>name</w:t></w:r><w:r><w:t>}!</w:t></w:r></w:p>"#,
        );
        let (out, count) = replace_in_docx_xml(&xml, &[literal("{name}", "Ana")]).unwrap();
        assert_eq!(count, 1);
        // Replacement lands in the run where the match starts, keeping its formatting
        assert!(out.contains(r#"<w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Hola Ana</w:t>"#), "{}", out);
        assert!(out.contains(r#"<w:t xml:space="preserve"></w:t>"#), "{}", out);
        assert!(out.contains(r#"<w:t xml:space="preserve">!</w:t>"#), "{}", out);
        let p = &parse_document_elements(&out)[0];
        match p {
            DocumentElement::Paragraph(p) => assert_eq!(p.text, "Hola Ana!"),
            _ => panic!("expected paragraph"),
        }
    }

    #[test]
    fn test_replace_multiple_placeholders_in_one_paragraph() {
        let xml = wrap_body(
            r#"<w:p><w:r><w:t>{</w:t></w:r><w:r><w:t>a</w:t></w:r><w:r><w:t>}-{b</w:t></w:r><w:r><w:t>}</w:t></w:r></w:p>"#,
        );
        let (out, count) =
            replace_in_docx_xml(&xml, &[literal("{a}", "1"), literal("{b}", "2")]).unwrap();
        assert_eq!(count, 2);
        match &parse_document_elements(&out)[0] {
            DocumentElement::Paragraph(p) => assert_eq!(p.text, "1-2"),
            _ => panic!("expected paragraph"),
        }
    }

    #[test]
    fn test_replace_does_not_cross_paragraphs_or_tabs() {
        let xml = wrap_body(
            r#"<w:p><w:r><w:t>{na</w:t></w:r></w:p><w:p><w:r><w:t>me}</w:t></w:r></w:p><w:p><w:r><w:t>{na</w:t><w:tab/><w:t>me}</w:t></w:r></w:p>"#,
        );
        let (out, count) = replace_in_docx_xml(&xml, &[literal("{name}", "X")]).unwrap();
        assert_eq!(count, 0);
        assert_eq!(out, xml);
    }

    #[test]
    fn test_replace_case_insensitive_and_unicode_safe() {
        let xml = wrap_body(r#"<w:p><w:r><w:t>İstanbul NOMBRE nombre</w:t></w:r></w:p>"#);
        let rule = ReplaceRule::new("nombre", "Ñandú", false, false).unwrap();
        let (out, count) = replace_in_docx_xml(&xml, &[rule]).unwrap();
        assert_eq!(count, 2);
        assert!(out.contains("İstanbul Ñandú Ñandú"), "{}", out);
    }

    #[test]
    fn test_replace_literal_dollar_and_regex_captures() {
        let xml = wrap_body(r#"<w:p><w:r><w:t>{precio} 2026-10-05</w:t></w:r></w:p>"#);
        let (out, _) = replace_in_docx_xml(&xml, &[literal("{precio}", "$1 USD")]).unwrap();
        assert!(out.contains("$1 USD"), "{}", out);

        let rule = ReplaceRule::new(r"(\d{4})-(\d{2})-(\d{2})", "$3/$2/$1", true, true).unwrap();
        let (out, count) = replace_in_docx_xml(&xml, &[rule]).unwrap();
        assert_eq!(count, 1);
        assert!(out.contains("05/10/2026"), "{}", out);
    }

    #[test]
    fn test_replace_escapes_xml_special_chars() {
        let xml = wrap_body(r#"<w:p><w:r><w:t>{empresa}</w:t></w:r></w:p>"#);
        let (out, _) = replace_in_docx_xml(&xml, &[literal("{empresa}", "A&B <S.A.>")]).unwrap();
        assert!(out.contains("A&amp;B &lt;S.A.&gt;"), "{}", out);
        match &parse_document_elements(&out)[0] {
            DocumentElement::Paragraph(p) => assert_eq!(p.text, "A&B <S.A.>"),
            _ => panic!("expected paragraph"),
        }
    }
}

