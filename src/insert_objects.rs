//! Tables and pictures inserted from code (`insertTable` / `insertImage`), each as one undo step.
//!
//! Child module of `docx_parser` so it can work on the package parts directly.

use super::*;

const MAX_TABLE_ROWS: usize = 500;
const MAX_TABLE_COLS: usize = 20;
const MAX_IMAGE_BYTES: usize = 15 * 1024 * 1024;

/// Gray, quiet look shared by every table the editor creates
const TABLE_BORDER_COLOR: &str = "BFBFBF";
const HEADER_FILL: &str = "F2F2F2";

pub(super) const NS_WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const REL_IMAGE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/// A table for `insert_table`: its cells as text (a "\n" is a line break inside the cell)
#[derive(Deserialize, Debug, Clone, PartialEq)]
pub struct NewTable {
    pub rows: Vec<Vec<String>>,
    /// The first row is a header: bold, shaded and repeated on every page
    #[serde(default = "default_true")]
    pub header: bool,
    /// Relative column widths (e.g. `[2, 1, 1]`); equal columns when unset
    #[serde(default)]
    pub widths: Option<Vec<f64>>,
    /// Per-column alignment: "left" | "center" | "right"
    #[serde(default)]
    pub align: Option<Vec<String>>,
}

fn default_true() -> bool {
    true
}

/// Options of `insert_image`. Lengths in px (96 per inch, like the layout).
#[derive(Deserialize, Debug, Clone, Default, PartialEq)]
pub struct NewImage {
    #[serde(default)]
    pub width: Option<f64>,
    #[serde(default)]
    pub height: Option<f64>,
    /// "left" | "center" | "right"
    #[serde(default)]
    pub align: Option<String>,
    /// Alternative text (accessibility; Word shows it as the picture description)
    #[serde(default)]
    pub alt: Option<String>,
    /// How text flows around it (default "inline"); see `ImageUpdate` for the values
    #[serde(default)]
    pub wrap: Option<String>,
    #[serde(default)]
    pub wrap_side: Option<String>,
    /// Position of a floating picture (see `ImageUpdate`)
    #[serde(default)]
    pub h_relative: Option<String>,
    #[serde(default)]
    pub h_offset: Option<f64>,
    #[serde(default)]
    pub h_align: Option<String>,
    #[serde(default)]
    pub v_relative: Option<String>,
    #[serde(default)]
    pub v_offset: Option<f64>,
    #[serde(default)]
    pub v_align: Option<String>,
    /// Space between a floating picture and the text, px
    #[serde(default)]
    pub distance: Option<f64>,
}

impl NewImage {
    /// The floating layout asked for, applied once the picture is in place
    fn layout(&self) -> Option<super::ImageUpdate> {
        let update = super::ImageUpdate {
            wrap: self.wrap.clone().filter(|w| w != "inline"),
            wrap_side: self.wrap_side.clone(),
            h_relative: self.h_relative.clone(),
            h_offset: self.h_offset,
            h_align: self.h_align.clone(),
            v_relative: self.v_relative.clone(),
            v_offset: self.v_offset,
            v_align: self.v_align.clone(),
            distance: self.distance,
            ..Default::default()
        };
        (update != super::ImageUpdate::default()).then_some(update)
    }
}

impl DocxModifier {
    /// Inserts `table` at (`index`, `offset`): text before the position stays above the table
    /// and text after it goes below. Returns the index of the first cell's paragraph and the
    /// caret at the start of the paragraph after the table.
    pub fn insert_table(&mut self, index: usize, offset: usize, table: &NewTable) -> Result<(usize, (usize, usize)), String> {
        if index >= HEADER_FOOTER_BASE {
            return Err("No se pueden insertar tablas en encabezados ni pies de página.".to_string());
        }
        if index >= TEXT_BOX_BASE {
            return Err("No se puede insertar una tabla dentro de un cuadro de texto.".to_string());
        }
        let cols = table.rows.iter().map(Vec::len).max().unwrap_or(0);
        if table.rows.is_empty() || cols == 0 {
            return Err("La tabla no tiene celdas.".to_string());
        }
        if table.rows.len() > MAX_TABLE_ROWS || cols > MAX_TABLE_COLS {
            return Err(format!("La tabla supera el máximo de {} filas y {} columnas.", MAX_TABLE_ROWS, MAX_TABLE_COLS));
        }

        let (len, in_table) = {
            let xml = self.get_file_string("word/document.xml")?;
            let range = self.body_paragraph(&xml, index)?;
            (parse_paragraph_fragment(&xml[range.clone()], &self.styles).text.chars().count(), inside_table(&xml, range.start))
        };
        if in_table {
            return Err("No se puede insertar una tabla dentro de otra tabla.".to_string());
        }
        // The table goes right before the paragraph that holds the text after the position
        let offset = offset.min(len);
        let target = if offset > 0 {
            self.split_paragraph(index, offset)?;
            index + 1
        } else {
            index
        };

        let mut xml = self.get_file_string("word/document.xml")?;
        let range = self.body_paragraph(&xml, target)?;
        let table_xml = build_table_xml(table, cols, text_width_twips(&xml));
        xml.insert_str(range.start, &table_xml);
        self.put_file("word/document.xml".to_string(), xml.into_bytes());

        let cells = table.rows.len() * cols;
        Ok((target, (target + cells, 0)))
    }

    /// Inserts a picture (PNG, JPEG or GIF) in a paragraph of its own at (`index`, `offset`).
    /// Without a size it keeps its pixel size; it never exceeds the text area of the page.
    /// Returns the index of the picture's paragraph and the caret at the start of the next one.
    pub fn insert_image(&mut self, index: usize, offset: usize, bytes: &[u8], image: &NewImage) -> Result<(usize, (usize, usize)), String> {
        if bytes.len() > MAX_IMAGE_BYTES {
            return Err(format!("La imagen supera el máximo de {} MB.", MAX_IMAGE_BYTES / 1024 / 1024));
        }
        let (format, px_w, px_h) = image_info(bytes).ok_or("Formato de imagen no soportado: usa PNG, JPEG o GIF.")?;
        if index >= HEADER_FOOTER_BASE {
            return Err("No se pueden insertar imágenes en encabezados ni pies de página.".to_string());
        }
        let layout = image.layout();
        if layout.is_some() && index >= TEXT_BOX_BASE {
            return Err("Dentro de un cuadro de texto solo se pueden insertar imágenes en línea con el texto.".to_string());
        }
        if let Some(update) = &layout {
            if update.wrap.is_none() {
                return Err("Para posicionar la imagen indica también su ajuste (wrap), por ejemplo \"square\".".to_string());
            }
            update.validate()?;
        }
        if let Some(align) = &image.align {
            if !["left", "center", "right"].contains(&align.as_str()) {
                return Err(format!("Alineación no válida: {}", align));
            }
        }

        let (len, page) = {
            let xml = self.get_file_string("word/document.xml")?;
            let range = self.body_paragraph(&xml, index)?;
            (parse_paragraph_fragment(&xml[range], &self.styles).text.chars().count(), extract_page_setup_quick_xml(&xml))
        };
        let (width, height) = fit_image(px_w, px_h, image.width, image.height, &page);
        let offset = offset.min(len);

        // A floating picture is anchored to the paragraph at the position, as Word does; the
        // text stays where it was and flows around it
        if let Some(update) = layout {
            let rel_id = self.add_image_part(bytes, format)?;
            let mut xml = self.get_file_string("word/document.xml")?;
            xml = ensure_namespace(&xml, "wp", NS_WP);
            xml = ensure_namespace(&xml, "r", NS_R);
            let doc_pr_id = next_doc_pr_id(&xml);
            let run = drawing_run(&rel_id, doc_pr_id, width, height, image.alt.as_deref().unwrap_or(""));
            let range = self.body_paragraph(&xml, index)?;
            let paragraph = with_first_run(&xml[range.clone()], &run)?;
            xml.replace_range(range, &paragraph);
            self.put_file("word/document.xml".to_string(), xml.into_bytes());
            self.update_image(index, 0, &update)?;
            return Ok((index, (index, offset)));
        }

        // An empty paragraph for the picture, and always one after it for the caret
        let picture = NewParagraph { align: image.align.clone(), ..Default::default() };
        let mut paragraphs = vec![picture];
        if offset == len {
            paragraphs.push(NewParagraph::default());
        }
        let (first, _) = self.insert_paragraphs(index, offset, &paragraphs)?;

        let rel_id = self.add_image_part(bytes, format)?;
        let mut xml = self.get_file_string("word/document.xml")?;
        xml = ensure_namespace(&xml, "wp", NS_WP);
        xml = ensure_namespace(&xml, "r", NS_R);
        let doc_pr_id = next_doc_pr_id(&xml);
        let run = drawing_run(&rel_id, doc_pr_id, width, height, image.alt.as_deref().unwrap_or(""));
        let range = self.body_paragraph(&xml, first)?;
        let paragraph = with_run(&xml[range.clone()], &run);
        xml.replace_range(range, &paragraph);
        self.put_file("word/document.xml".to_string(), xml.into_bytes());

        Ok((first, (first + 1, 0)))
    }

    /// Stores the picture under word/media and relates it to document.xml; returns the rId
    fn add_image_part(&mut self, bytes: &[u8], format: ImageFormat) -> Result<String, String> {
        let ext = format.extension();
        let mut n = 1;
        let media = loop {
            let name = format!("word/media/image{}.{}", n, ext);
            if !self.files.contains_key(&name) {
                break name;
            }
            n += 1;
        };
        self.put_file(media.clone(), bytes.to_vec());

        let types = self.get_file_string("[Content_Types].xml")?;
        if !has_content_type(&types, ext) {
            let entry = format!(r#"<Default Extension="{}" ContentType="{}"/>"#, ext, format.mime());
            self.put_file("[Content_Types].xml".to_string(), append_child(&types, "Types", &entry)?.into_bytes());
        }

        let rels_path = "word/_rels/document.xml.rels";
        let rels = self.get_file_string(rels_path).unwrap_or_else(|_| {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.to_string()
        });
        let existing = parse_relationships_map(&rels);
        let mut i = 1;
        let rel_id = loop {
            let id = format!("rIdImg{}", i);
            if !existing.contains_key(&id) {
                break id;
            }
            i += 1;
        };
        let target = media.trim_start_matches("word/");
        let entry = format!(r#"<Relationship Id="{}" Type="{}" Target="{}"/>"#, rel_id, REL_IMAGE, target);
        self.put_file(rels_path.to_string(), append_child(&rels, "Relationships", &entry)?.into_bytes());
        Ok(rel_id)
    }
}

/// True when the byte position sits inside a `w:tbl` element
fn inside_table(xml: &str, pos: usize) -> bool {
    let before = &xml[..pos];
    let opened = before.matches("<w:tbl>").count() + before.matches("<w:tbl ").count();
    opened > before.matches("</w:tbl>").count()
}

/// Width between the page margins in twips (1 px of the layout = 15 twips)
fn text_width_twips(doc: &str) -> u32 {
    let page = extract_page_setup_quick_xml(doc);
    (((page.width - page.margin_left - page.margin_right) * 15.0).round() as u32).max(1440)
}

/// Column widths in twips that add up to `total`
fn column_widths(weights: Option<&Vec<f64>>, cols: usize, total: u32) -> Vec<u32> {
    let weights: Vec<f64> = match weights {
        Some(w) if w.len() == cols && w.iter().all(|v| v.is_finite() && *v > 0.0) => w.clone(),
        _ => vec![1.0; cols],
    };
    let sum: f64 = weights.iter().sum();
    let mut widths: Vec<u32> = weights.iter().map(|w| ((w / sum) * total as f64).floor() as u32).collect();
    let rest = total.saturating_sub(widths.iter().sum());
    if let Some(last) = widths.last_mut() {
        *last += rest;
    }
    widths
}

/// The table as WordprocessingML with one paragraph per cell
fn build_table_xml(table: &NewTable, cols: usize, total_width: u32) -> String {
    let widths = column_widths(table.widths.as_ref(), cols, total_width);
    let border = |side: &str| format!(r#"<w:{} w:val="single" w:sz="4" w:space="0" w:color="{}"/>"#, side, TABLE_BORDER_COLOR);
    let mut xml = String::from("<w:tbl><w:tblPr>");
    xml.push_str(&format!(r#"<w:tblW w:w="{}" w:type="dxa"/>"#, total_width));
    xml.push_str("<w:tblBorders>");
    for side in ["top", "left", "bottom", "right", "insideH", "insideV"] {
        xml.push_str(&border(side));
    }
    xml.push_str("</w:tblBorders>");
    xml.push_str(r#"<w:tblLayout w:type="fixed"/>"#);
    xml.push_str(r#"<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>"#);
    xml.push_str(&format!(r#"<w:tblLook w:val="{}" w:firstRow="{}" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>"#,
        if table.header { "0420" } else { "0400" }, u8::from(table.header)));
    xml.push_str("</w:tblPr><w:tblGrid>");
    for w in &widths {
        xml.push_str(&format!(r#"<w:gridCol w:w="{}"/>"#, w));
    }
    xml.push_str("</w:tblGrid>");

    for (r, row) in table.rows.iter().enumerate() {
        let is_header = table.header && r == 0;
        xml.push_str("<w:tr>");
        if is_header {
            xml.push_str("<w:trPr><w:tblHeader/></w:trPr>");
        }
        for (c, width) in widths.iter().enumerate() {
            let text = row.get(c).map(String::as_str).unwrap_or("");
            xml.push_str(&format!(r#"<w:tc><w:tcPr><w:tcW w:w="{}" w:type="dxa"/>"#, width));
            if is_header {
                xml.push_str(&format!(r#"<w:shd w:val="clear" w:color="auto" w:fill="{}"/>"#, HEADER_FILL));
            }
            xml.push_str(r#"</w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0"/>"#);
            let align = table.align.as_ref().and_then(|a| a.get(c)).map(String::as_str);
            if let Some(jc @ ("center" | "right")) = align {
                xml.push_str(&format!(r#"<w:jc w:val="{}"/>"#, jc));
            }
            xml.push_str("</w:pPr>");
            if !text.is_empty() {
                xml.push_str("<w:r>");
                if is_header {
                    xml.push_str("<w:rPr><w:b/></w:rPr>");
                }
                let normalized = text.replace("\r\n", "\n");
                for (i, line) in normalized.split('\n').enumerate() {
                    if i > 0 {
                        xml.push_str("<w:br/>");
                    }
                    xml.push_str(&format!(r#"<w:t xml:space="preserve">{}</w:t>"#, escape_xml(line)));
                }
                xml.push_str("</w:r>");
            }
            xml.push_str("</w:p></w:tc>");
        }
        xml.push_str("</w:tr>");
    }
    xml.push_str("</w:tbl>");
    xml
}

/// `rows`×`cols` table (header row from `headers`) placed at the end of the body, before the
/// section properties (`insertTable(rows, cols, headers)` of earlier versions)
pub(super) fn append_table_xml(xml: &str, rows: usize, cols: usize, headers: &[String]) -> Result<String, String> {
    let rows = rows.clamp(1, MAX_TABLE_ROWS);
    let cols = cols.clamp(1, MAX_TABLE_COLS);
    let mut cells = vec![vec![String::new(); cols]; rows];
    for (c, h) in headers.iter().take(cols).enumerate() {
        cells[0][c] = h.clone();
    }
    let table = NewTable { rows: cells, header: true, widths: None, align: None };
    let table_xml = build_table_xml(&table, cols, text_width_twips(xml));
    // Word needs a paragraph between a table and the end of the body
    let block = format!("{}<w:p/>", table_xml);
    let end = xml.rfind("</w:body>").ok_or("El documento no tiene cuerpo (w:body).")?;
    let at = xml[..end].rfind("<w:sectPr").filter(|&i| !inside_paragraph(xml, i)).unwrap_or(end);
    let mut out = xml.to_string();
    out.insert_str(at, &block);
    Ok(out)
}

/// True when `pos` is inside a `w:p` (a section break of a paragraph, not of the body)
fn inside_paragraph(xml: &str, pos: usize) -> bool {
    let before = &xml[..pos];
    let opened = before.matches("<w:p>").count() + before.matches("<w:p ").count();
    opened > before.matches("</w:p>").count()
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub(crate) enum ImageFormat {
    Png,
    Jpeg,
    Gif,
}

impl ImageFormat {
    fn extension(self) -> &'static str {
        match self {
            ImageFormat::Png => "png",
            ImageFormat::Jpeg => "jpeg",
            ImageFormat::Gif => "gif",
        }
    }

    fn mime(self) -> &'static str {
        match self {
            ImageFormat::Png => "image/png",
            ImageFormat::Jpeg => "image/jpeg",
            ImageFormat::Gif => "image/gif",
        }
    }
}

/// Format and pixel size read from the file header
pub(crate) fn image_info(bytes: &[u8]) -> Option<(ImageFormat, u32, u32)> {
    let be16 = |i: usize| bytes.get(i..i + 2).map(|b| u16::from_be_bytes([b[0], b[1]]) as u32);
    let be32 = |i: usize| bytes.get(i..i + 4).map(|b| u32::from_be_bytes([b[0], b[1], b[2], b[3]]));
    let (format, w, h) = if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        (ImageFormat::Png, be32(16)?, be32(20)?)
    } else if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        let le16 = |i: usize| bytes.get(i..i + 2).map(|b| u16::from_le_bytes([b[0], b[1]]) as u32);
        (ImageFormat::Gif, le16(6)?, le16(8)?)
    } else if bytes.starts_with(&[0xFF, 0xD8]) {
        // Walk the JPEG segments up to the frame header (SOFn)
        let mut i = 2;
        loop {
            while bytes.get(i)? != &0xFF {
                i += 1;
            }
            while bytes.get(i)? == &0xFF {
                i += 1;
            }
            let marker = *bytes.get(i)?;
            i += 1;
            if matches!(marker, 0xD8 | 0x01 | 0xD0..=0xD7) {
                continue;
            }
            let len = be16(i)? as usize;
            let is_sof = matches!(marker, 0xC0..=0xCF) && !matches!(marker, 0xC4 | 0xC8 | 0xCC);
            if is_sof {
                break (ImageFormat::Jpeg, be16(i + 5)?, be16(i + 3)?);
            }
            if marker == 0xD9 || len < 2 {
                return None;
            }
            i += len;
        }
    } else {
        return None;
    };
    (w > 0 && h > 0).then_some((format, w, h))
}

/// Requested size (or the pixel size), keeping the aspect ratio and within the text area
fn fit_image(px_w: u32, px_h: u32, width: Option<f64>, height: Option<f64>, page: &PageSetup) -> (f64, f64) {
    let ratio = px_h as f64 / px_w as f64;
    let valid = |v: Option<f64>| v.filter(|v| v.is_finite() && *v > 0.0);
    let (mut w, mut h) = match (valid(width), valid(height)) {
        (Some(w), Some(h)) => (w, h),
        (Some(w), None) => (w, w * ratio),
        (None, Some(h)) => (h / ratio, h),
        (None, None) => (px_w as f64, px_h as f64),
    };
    let max_w = (page.width - page.margin_left - page.margin_right).max(48.0);
    let max_h = (page.height - page.margin_top - page.margin_bottom).max(48.0);
    let scale = (max_w / w).min(max_h / h).min(1.0);
    w *= scale;
    h *= scale;
    (w.round().max(1.0), h.round().max(1.0))
}

/// The run with an inline picture, as Word writes it
fn drawing_run(rel_id: &str, id: u64, width: f64, height: f64, alt: &str) -> String {
    let cx = (width * EMU_PER_PX).round() as i64;
    let cy = (height * EMU_PER_PX).round() as i64;
    let alt = escape_xml(alt);
    format!(
        concat!(
            r#"<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">"#,
            r#"<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>"#,
            r#"<wp:docPr id="{id}" name="Imagen {id}" descr="{alt}"/>"#,
            r#"<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="{a}" noChangeAspect="1"/></wp:cNvGraphicFramePr>"#,
            r#"<a:graphic xmlns:a="{a}"><a:graphicData uri="{pic}"><pic:pic xmlns:pic="{pic}">"#,
            r#"<pic:nvPicPr><pic:cNvPr id="0" name="Imagen {id}" descr="{alt}"/><pic:cNvPicPr/></pic:nvPicPr>"#,
            r#"<pic:blipFill><a:blip r:embed="{rel}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>"#,
            r#"<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>"#,
            r#"</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
        ),
        cx = cx, cy = cy, id = id, alt = alt, a = NS_A, pic = NS_PIC, rel = rel_id
    )
}

/// The paragraph with `run` appended
fn with_run(paragraph: &str, run: &str) -> String {
    if let Some(pos) = paragraph.rfind("</w:p>") {
        let mut out = paragraph.to_string();
        out.insert_str(pos, run);
        out
    } else {
        // Self-closing <w:p/> or <w:p .../>
        let open = paragraph.trim_end_matches("/>").trim_end();
        format!("{}>{}</w:p>", open, run)
    }
}

/// The paragraph with `run` as its first run (right after its properties)
fn with_first_run(paragraph: &str, run: &str) -> Result<String, String> {
    let tokens = crate::paragraph_edit::tokenize(paragraph)?;
    if !matches!(tokens.first().map(|t| &t.ev), Some(Event::Start(_))) {
        return Ok(with_run(paragraph, run));
    }
    let at = match tokens.get(1).map(|t| &t.ev) {
        Some(Event::Start(e) | Event::Empty(e)) if tag_is(e.name().as_ref(), "pPr") => {
            tokens[crate::paragraph_edit::element_end(&tokens, 1)].span.end
        }
        _ => tokens[0].span.end,
    };
    let mut out = paragraph.to_string();
    out.insert_str(at, run);
    Ok(out)
}

/// Unique id for a new `wp:docPr` (Word refuses duplicates)
fn next_doc_pr_id(doc: &str) -> u64 {
    let re = regex::Regex::new(r#"<wp:docPr\b[^>]*\bid="(\d+)""#).expect("regex");
    re.captures_iter(doc).filter_map(|c| c[1].parse::<u64>().ok()).max().unwrap_or(0) + 1
}

/// Declares `xmlns:prefix` on the root element when it is missing
pub(super) fn ensure_namespace(xml: &str, prefix: &str, uri: &str) -> String {
    let Some(start) = xml.find("<w:document") else { return xml.to_string() };
    let Some(end) = xml[start..].find('>').map(|i| start + i) else { return xml.to_string() };
    if xml[start..end].contains(&format!("xmlns:{}=", prefix)) {
        return xml.to_string();
    }
    let mut out = xml.to_string();
    out.insert_str(start + "<w:document".len(), &format!(r#" xmlns:{}="{}""#, prefix, uri));
    out
}

fn has_content_type(types: &str, ext: &str) -> bool {
    types.to_lowercase().contains(&format!("extension=\"{}\"", ext))
}

/// Appends `child` as the last child of the root element `root` (which may be self-closing)
fn append_child(xml: &str, root: &str, child: &str) -> Result<String, String> {
    let close = format!("</{}>", root);
    if let Some(pos) = xml.rfind(&close) {
        let mut out = xml.to_string();
        out.insert_str(pos, child);
        return Ok(out);
    }
    let start = xml.find(&format!("<{}", root)).ok_or_else(|| format!("No se encontró <{}>.", root))?;
    let end = xml[start..].find("/>").map(|i| start + i).ok_or_else(|| format!("<{}> mal formado.", root))?;
    Ok(format!("{}>{}{}{}", &xml[..end], child, close, &xml[end + 2..]))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn png(w: u32, h: u32) -> Vec<u8> {
        let mut b = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR".to_vec();
        b.extend_from_slice(&w.to_be_bytes());
        b.extend_from_slice(&h.to_be_bytes());
        b.extend_from_slice(&[8, 6, 0, 0, 0]);
        b
    }

    #[test]
    fn reads_image_sizes() {
        assert_eq!(image_info(&png(640, 480)), Some((ImageFormat::Png, 640, 480)));
        let gif = [b"GIF89a".as_slice(), &[0x40, 0x01, 0xF0, 0x00]].concat();
        assert_eq!(image_info(&gif), Some((ImageFormat::Gif, 320, 240)));
        // SOI, APP0 (len 16), SOF0 with height 300 and width 400
        let mut jpeg = vec![0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10];
        jpeg.extend_from_slice(&[0; 14]);
        jpeg.extend_from_slice(&[0xFF, 0xC0, 0x00, 0x11, 0x08, 0x01, 0x2C, 0x01, 0x90, 0x03]);
        assert_eq!(image_info(&jpeg), Some((ImageFormat::Jpeg, 400, 300)));
        assert_eq!(image_info(b"RIFF....WEBP"), None);
        assert_eq!(image_info(&[0xFF, 0xD8, 0xFF]), None);
    }

    #[test]
    fn images_fit_the_text_area() {
        let page = PageSetup::default(); // 800 × 1130 px with 65 px side margins
        assert_eq!(fit_image(200, 100, None, None, &page), (200.0, 100.0));
        assert_eq!(fit_image(2000, 1000, None, None, &page), (670.0, 335.0));
        assert_eq!(fit_image(200, 100, Some(100.0), None, &page), (100.0, 50.0));
        assert_eq!(fit_image(200, 100, None, Some(25.0), &page), (50.0, 25.0));
    }

    #[test]
    fn column_widths_add_up() {
        assert_eq!(column_widths(None, 3, 9000), vec![3000, 3000, 3000]);
        assert_eq!(column_widths(Some(&vec![2.0, 1.0, 1.0]), 3, 9001), vec![4500, 2250, 2251]);
        // Wrong number of weights falls back to equal columns
        assert_eq!(column_widths(Some(&vec![1.0]), 2, 100), vec![50, 50]);
    }

    #[test]
    fn appends_to_self_closing_roots() {
        let types = r#"<?xml version="1.0"?><Types xmlns="x"/>"#;
        assert_eq!(append_child(types, "Types", "<Default/>").unwrap(), r#"<?xml version="1.0"?><Types xmlns="x"><Default/></Types>"#);
        assert_eq!(with_run("<w:p/>", "<w:r/>"), "<w:p><w:r/></w:p>");
        assert_eq!(with_run(r#"<w:p w:rsidR="1"/>"#, "<w:r/>"), r#"<w:p w:rsidR="1"><w:r/></w:p>"#);
        assert_eq!(with_run("<w:p><w:pPr/></w:p>", "<w:r/>"), "<w:p><w:pPr/><w:r/></w:p>");
        assert_eq!(with_first_run("<w:p><w:pPr><w:jc/></w:pPr><w:r><w:t>a</w:t></w:r></w:p>", "<w:r/>").unwrap(), "<w:p><w:pPr><w:jc/></w:pPr><w:r/><w:r><w:t>a</w:t></w:r></w:p>");
        assert_eq!(with_first_run("<w:p><w:r><w:t>a</w:t></w:r></w:p>", "<w:r/>").unwrap(), "<w:p><w:r/><w:r><w:t>a</w:t></w:r></w:p>");
        assert_eq!(with_first_run("<w:p/>", "<w:r/>").unwrap(), "<w:p><w:r/></w:p>");
    }
}
