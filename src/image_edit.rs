//! Editing pictures already in the document: size, position, how text flows around them,
//! alternative text and deletion (`updateImage` / `deleteImage`), each as one undo step.
//!
//! A picture is addressed by its paragraph and its index among that paragraph's pictures,
//! the same numbering `ParagraphInfo::images` uses. Only the `wp:inline` / `wp:anchor`
//! element is rewritten: the picture itself (`a:graphic`), its id and its description are
//! carried over verbatim.

use super::insert_objects::{ensure_namespace, NS_WP};
use super::*;
use crate::paragraph_edit::{element_end, tokenize};
use std::ops::Range;

const MIN_SIZE_PX: f64 = 4.0;
const MAX_SIZE_PX: f64 = 4000.0;
/// Word's default space between a floating picture and the text: 0.125" left and right
const DEFAULT_DIST_X: f64 = 12.0;

const WRAPS: [&str; 7] = ["inline", "square", "tight", "through", "topAndBottom", "behind", "inFront"];
const WRAP_SIDES: [&str; 4] = ["bothSides", "left", "right", "largest"];
const H_RELATIVE: [&str; 8] = ["margin", "page", "column", "character", "leftMargin", "rightMargin", "insideMargin", "outsideMargin"];
const V_RELATIVE: [&str; 8] = ["margin", "page", "paragraph", "line", "topMargin", "bottomMargin", "insideMargin", "outsideMargin"];
const H_ALIGN: [&str; 5] = ["left", "center", "right", "inside", "outside"];
const V_ALIGN: [&str; 5] = ["top", "center", "bottom", "inside", "outside"];

/// Changes for `update_image`; unset fields keep their current value. Lengths in px.
#[derive(Deserialize, Debug, Clone, Default, PartialEq)]
pub struct ImageUpdate {
    #[serde(default)]
    pub width: Option<f64>,
    #[serde(default)]
    pub height: Option<f64>,
    /// With only one of width/height, scale the other to keep the aspect ratio (default true)
    #[serde(default)]
    pub keep_ratio: Option<bool>,
    /// "inline" | "square" | "tight" | "through" | "topAndBottom" | "behind" | "inFront"
    #[serde(default)]
    pub wrap: Option<String>,
    /// "bothSides" | "left" | "right" | "largest"
    #[serde(default)]
    pub wrap_side: Option<String>,
    /// Horizontal frame: "margin" | "page" | "column" | "character" | "leftMargin" | …
    #[serde(default)]
    pub h_relative: Option<String>,
    /// Distance from the frame's left edge (clears `h_align`)
    #[serde(default)]
    pub h_offset: Option<f64>,
    /// "left" | "center" | "right" | "inside" | "outside" within the frame (clears `h_offset`)
    #[serde(default)]
    pub h_align: Option<String>,
    /// Vertical frame: "margin" | "page" | "paragraph" | "line" | "topMargin" | …
    #[serde(default)]
    pub v_relative: Option<String>,
    #[serde(default)]
    pub v_offset: Option<f64>,
    /// "top" | "center" | "bottom" | "inside" | "outside"
    #[serde(default)]
    pub v_align: Option<String>,
    /// Space between the picture and the text on every side…
    #[serde(default)]
    pub distance: Option<f64>,
    /// …or per side
    #[serde(default)]
    pub dist_top: Option<f64>,
    #[serde(default)]
    pub dist_bottom: Option<f64>,
    #[serde(default)]
    pub dist_left: Option<f64>,
    #[serde(default)]
    pub dist_right: Option<f64>,
    /// Alternative text
    #[serde(default)]
    pub alt: Option<String>,
}

impl ImageUpdate {
    pub(super) fn validate(&self) -> Result<(), String> {
        let check = |value: &Option<String>, allowed: &[&str], what: &str| match value {
            Some(v) if !allowed.contains(&v.as_str()) => {
                Err(format!("{} no válido: {} (usa {})", what, v, allowed.join(", ")))
            }
            _ => Ok(()),
        };
        check(&self.wrap, &WRAPS, "Ajuste de texto")?;
        check(&self.wrap_side, &WRAP_SIDES, "Lado del ajuste")?;
        check(&self.h_relative, &H_RELATIVE, "Referencia horizontal")?;
        check(&self.v_relative, &V_RELATIVE, "Referencia vertical")?;
        check(&self.h_align, &H_ALIGN, "Alineación horizontal")?;
        check(&self.v_align, &V_ALIGN, "Alineación vertical")?;
        let lengths = [
            self.width, self.height, self.h_offset, self.v_offset, self.distance,
            self.dist_top, self.dist_bottom, self.dist_left, self.dist_right,
        ];
        if lengths.iter().flatten().any(|v| !v.is_finite()) {
            return Err("Las medidas de la imagen deben ser números finitos.".to_string());
        }
        Ok(())
    }

    /// Something that only floating pictures have was asked for
    fn positions(&self) -> bool {
        self.h_relative.is_some() || self.h_offset.is_some() || self.h_align.is_some()
            || self.v_relative.is_some() || self.v_offset.is_some() || self.v_align.is_some()
    }
}

/// A picture of a paragraph fragment: the `w:drawing` element and what encloses it
#[derive(Debug, Clone)]
pub(crate) struct DrawingSpan {
    /// The `w:drawing` element
    pub range: Range<usize>,
    /// The `w:r` holding it
    pub run: Option<Range<usize>>,
    /// An `mc:AlternateContent` around it inside the run (its fallback shows the same picture)
    pub alternate: Option<Range<usize>>,
    pub image: ImageRef,
}

/// Pictures of a `w:p` fragment in the order (and with the filtering) of `parse_paragraph_with`
pub(crate) fn drawing_spans(p_xml: &str) -> Result<Vec<DrawingSpan>, String> {
    let tokens = tokenize(p_xml)?;
    let mut spans = Vec::new();
    // Open elements: (local name, token index)
    let mut stack: Vec<(String, usize)> = Vec::new();
    let mut i = 0;
    while i < tokens.len() {
        match &tokens[i].ev {
            Event::Start(e) => {
                let name = e.name();
                let n = name.as_ref();
                // Same containers the paragraph parser skips: text boxes, VML, tracked changes
                if ["pict", "object", "txbxContent", "pPrChange", "rPrChange", "Fallback"].iter().any(|t| tag_is(n, t)) {
                    i = element_end(&tokens, i) + 1;
                    continue;
                }
                if tag_is(n, "drawing") {
                    let end = element_end(&tokens, i);
                    let range = tokens[i].span.start..tokens[end].span.end;
                    let mut reader = Reader::from_str(&p_xml[range.clone()]);
                    let _ = reader.read_event();
                    if let Some(image) = parse_drawing(&mut reader, &StyleSheet::default()) {
                        let outer = |local: &str| {
                            stack.iter().rev().find(|(name, _)| name == local).map(|&(_, j)| {
                                tokens[j].span.start..tokens[element_end(&tokens, j)].span.end
                            })
                        };
                        spans.push(DrawingSpan { run: outer("r"), alternate: outer("AlternateContent"), range, image });
                    }
                    i = end + 1;
                    continue;
                }
                let local = std::str::from_utf8(n).unwrap_or("").rsplit(':').next().unwrap_or("").to_string();
                stack.push((local, i));
            }
            Event::End(_) => {
                stack.pop();
            }
            _ => {}
        }
        i += 1;
    }
    Ok(spans)
}

impl DocxModifier {
    /// Pictures of body paragraph `index` (cells included) as the parser reads them
    fn paragraph_drawing(&self, xml: &str, paragraph: usize, image: usize) -> Result<(Range<usize>, DrawingSpan), String> {
        let range = self.body_paragraph(xml, paragraph)?;
        let span = drawing_spans(&xml[range.clone()])?
            .into_iter()
            .nth(image)
            .ok_or_else(|| format!("El párrafo {} no tiene la imagen {}.", paragraph, image))?;
        Ok((range, span))
    }

    /// Applies `update` to picture `image` of paragraph `paragraph` and returns it as read back
    pub fn update_image(&mut self, paragraph: usize, image: usize, update: &ImageUpdate) -> Result<ImageRef, String> {
        update.validate()?;
        let mut xml = self.get_file_string("word/document.xml")?;
        let (p_range, span) = self.paragraph_drawing(&xml, paragraph, image)?;
        let current = span.image.clone();
        let target = merge_update(&current, update, &paragraph_align(&xml[p_range.clone()], &self.styles), self.page_text_area(&xml))?;

        let start = p_range.start + span.range.start;
        let end = p_range.start + span.range.end;
        let rebuilt = rebuild_drawing(&xml[start..end], &current, &target, next_relative_height(&xml))?;
        xml.replace_range(start..end, &rebuilt);
        if target.anchored {
            xml = ensure_namespace(&xml, "wp", NS_WP);
        }
        self.put_file("word/document.xml".to_string(), xml.into_bytes());

        let xml = self.get_file_string("word/document.xml")?;
        let mut updated = self.paragraph_drawing(&xml, paragraph, image)?.1.image;
        updated.offset = current.offset;
        Ok(updated)
    }

    /// Removes picture `image` of paragraph `paragraph` (and its run when nothing else is left
    /// in it). The media part stays in the package, as Word does until it saves.
    pub fn delete_image(&mut self, paragraph: usize, image: usize) -> Result<(), String> {
        let mut xml = self.get_file_string("word/document.xml")?;
        let (p_range, span) = self.paragraph_drawing(&xml, paragraph, image)?;
        let local = span.alternate.clone().unwrap_or(span.range.clone());
        let remove = match &span.run {
            Some(run) if run_is_empty_without(&xml[p_range.start..p_range.end], run, &local) => run.clone(),
            _ => local,
        };
        xml.replace_range(p_range.start + remove.start..p_range.start + remove.end, "");
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        Ok(())
    }

    /// Width and height of the text area of the page, px
    fn page_text_area(&self, xml: &str) -> (f64, f64) {
        let page = extract_page_setup_quick_xml(xml);
        (
            (page.width - page.margin_left - page.margin_right).max(48.0),
            (page.height - page.margin_top - page.margin_bottom).max(48.0),
        )
    }
}

fn paragraph_align(p_xml: &str, styles: &StyleSheet) -> String {
    parse_paragraph_fragment(p_xml, styles).align
}

/// The picture as it should be after `update`
fn merge_update(current: &ImageRef, update: &ImageUpdate, paragraph_align: &str, text_area: (f64, f64)) -> Result<ImageRef, String> {
    let mut img = current.clone();

    // Size, keeping the aspect ratio when only one side is given
    let ratio = if current.width > 0.0 { current.height / current.width } else { 1.0 };
    let keep = update.keep_ratio.unwrap_or(true);
    let (w, h) = match (update.width, update.height) {
        (Some(w), Some(h)) => (w, h),
        (Some(w), None) => (w, if keep { w * ratio } else { current.height }),
        (None, Some(h)) => (if keep { h / ratio } else { current.width }, h),
        (None, None) => (current.width, current.height),
    };
    // Huge pictures shrink into the text area (floating ones may stick out of it, as in Word)
    let (max_w, max_h) = if img.anchored || update.wrap.as_deref().is_some_and(|w| w != "inline") {
        (MAX_SIZE_PX, MAX_SIZE_PX)
    } else {
        text_area
    };
    let resized = update.width.is_some() || update.height.is_some() || (img.anchored && update.wrap.as_deref() == Some("inline"));
    let scale = if resized { (max_w / w).min(max_h / h).min(1.0) } else { 1.0 };
    img.width = (w * scale).clamp(MIN_SIZE_PX, MAX_SIZE_PX);
    img.height = (h * scale).clamp(MIN_SIZE_PX, MAX_SIZE_PX);

    if let Some(wrap) = &update.wrap {
        let was_inline = !img.anchored;
        img.wrap = wrap.clone();
        img.anchored = wrap != "inline";
        img.behind_text = wrap == "behind";
        if was_inline && img.anchored {
            // Float where the picture was: at the start of its paragraph, aligned like it
            img.h_relative = "column".into();
            img.h_offset = 0.0;
            img.h_align = match paragraph_align {
                "center" => Some("center".into()),
                "right" => Some("right".into()),
                _ => None,
            };
            img.v_relative = "paragraph".into();
            img.v_offset = 0.0;
            img.v_align = None;
            if img.dist_left == 0.0 && img.dist_right == 0.0 {
                img.dist_left = DEFAULT_DIST_X;
                img.dist_right = DEFAULT_DIST_X;
            }
        }
        if matches!(wrap.as_str(), "square" | "tight" | "through") && img.wrap_side.is_empty() {
            img.wrap_side = "bothSides".into();
        }
    }
    if !img.anchored {
        if update.positions() {
            return Err("La imagen está en línea con el texto: cambia su ajuste (por ejemplo a \"square\") para moverla.".to_string());
        }
        img.wrap = "inline".into();
        img.wrap_side.clear();
        img.h_align = None;
        img.v_align = None;
        img.h_offset = 0.0;
        img.v_offset = 0.0;
    }

    if let Some(side) = &update.wrap_side {
        img.wrap_side = side.clone();
    }
    if let Some(rel) = &update.h_relative {
        img.h_relative = rel.clone();
    }
    if let Some(rel) = &update.v_relative {
        img.v_relative = rel.clone();
    }
    if let Some(offset) = update.h_offset {
        img.h_offset = offset;
        img.h_align = None;
    }
    if let Some(align) = &update.h_align {
        img.h_align = Some(align.clone());
    }
    if let Some(offset) = update.v_offset {
        img.v_offset = offset;
        img.v_align = None;
    }
    if let Some(align) = &update.v_align {
        img.v_align = Some(align.clone());
    }

    let dist = |side: Option<f64>, current: f64| side.or(update.distance).unwrap_or(current).clamp(0.0, 500.0);
    img.dist_top = dist(update.dist_top, img.dist_top);
    img.dist_bottom = dist(update.dist_bottom, img.dist_bottom);
    img.dist_left = dist(update.dist_left, img.dist_left);
    img.dist_right = dist(update.dist_right, img.dist_right);
    if let Some(alt) = &update.alt {
        img.alt = alt.clone();
    }
    Ok(img)
}

/// Children of the `wp:inline` / `wp:anchor` element kept verbatim
#[derive(Default)]
struct DrawingParts<'a> {
    /// Start tag of the container (its attributes)
    container: Option<BytesStart<'a>>,
    effect_extent: Option<&'a str>,
    /// The current wrap element when it is a polygon wrap (kept for the same wrap kind)
    wrap_polygon: Option<(&'a str, &'a str)>,
    doc_pr: Option<&'a str>,
    frame_pr: Option<&'a str>,
    graphic: Option<&'a str>,
    /// Anchor-only extensions after the graphic (e.g. `wp14:sizeRelH`)
    extensions: Vec<&'a str>,
}

fn drawing_parts(drawing: &str) -> Result<DrawingParts<'_>, String> {
    let tokens = tokenize(drawing)?;
    let mut parts = DrawingParts::default();
    let Some(c) = tokens.iter().position(|t| matches!(&t.ev, Event::Start(e) if tag_is(e.name().as_ref(), "inline") || tag_is(e.name().as_ref(), "anchor"))) else {
        return Err("La imagen no tiene un elemento wp:inline ni wp:anchor.".to_string());
    };
    if let Event::Start(e) = &tokens[c].ev {
        parts.container = Some(e.clone().into_owned());
    }
    let c_end = element_end(&tokens, c);
    let mut i = c + 1;
    while i < c_end {
        let (Event::Start(e) | Event::Empty(e)) = &tokens[i].ev else {
            i += 1;
            continue;
        };
        let end = element_end(&tokens, i);
        let raw = &drawing[tokens[i].span.start..tokens[end].span.end];
        let name = e.name();
        let n = name.as_ref();
        let local = std::str::from_utf8(n).unwrap_or("").rsplit(':').next().unwrap_or("").to_string();
        match local.as_str() {
            "effectExtent" => parts.effect_extent = Some(raw),
            "wrapTight" | "wrapThrough" => {
                let kind = if local == "wrapTight" { "tight" } else { "through" };
                parts.wrap_polygon = Some((kind, raw));
            }
            "docPr" => parts.doc_pr = Some(raw),
            "cNvGraphicFramePr" => parts.frame_pr = Some(raw),
            "graphic" => parts.graphic = Some(raw),
            "simplePos" | "positionH" | "positionV" | "extent" | "wrapNone" | "wrapSquare" | "wrapTopAndBottom" => {}
            _ => parts.extensions.push(raw),
        }
        i = end + 1;
    }
    Ok(parts)
}

fn emu(px: f64) -> i64 {
    (px * EMU_PER_PX).round() as i64
}

/// The `w:drawing` element for `img`, reusing the picture of `drawing`
fn rebuild_drawing(drawing: &str, current: &ImageRef, img: &ImageRef, relative_height: u64) -> Result<String, String> {
    let parts = drawing_parts(drawing)?;
    let graphic = parts.graphic.ok_or("La imagen no tiene a:graphic.")?;
    let graphic = if img.width != current.width || img.height != current.height {
        resize_graphic(graphic, emu(img.width), emu(img.height))
    } else {
        graphic.to_string()
    };
    let doc_pr = match parts.doc_pr {
        Some(raw) if img.alt != current.alt => set_attr(raw, "descr", &escape_xml(&img.alt)),
        Some(raw) => raw.to_string(),
        None => format!(r#"<wp:docPr id="{0}" name="Imagen {0}" descr="{1}"/>"#, img.doc_pr_id.max(1), escape_xml(&img.alt)),
    };
    let (cx, cy) = (emu(img.width), emu(img.height));
    let dist = format!(
        r#"distT="{}" distB="{}" distL="{}" distR="{}""#,
        emu(img.dist_top), emu(img.dist_bottom), emu(img.dist_left), emu(img.dist_right)
    );
    let extent = format!(r#"<wp:extent cx="{}" cy="{}"/>"#, cx, cy);
    let effect = parts.effect_extent.unwrap_or(r#"<wp:effectExtent l="0" t="0" r="0" b="0"/>"#);
    let frame_pr = parts.frame_pr.unwrap_or("");

    let mut out = String::from("<w:drawing>");
    if !img.anchored {
        out.push_str(&format!("<wp:inline {}>", dist));
        out.push_str(&extent);
        out.push_str(effect);
        out.push_str(&doc_pr);
        out.push_str(frame_pr);
        out.push_str(&graphic);
        out.push_str("</wp:inline>");
    } else {
        // Keep the anchor's other attributes (z-order, ids, lock…) when it already floated
        let kept = match &parts.container {
            Some(e) if current.anchored => e
                .attributes()
                .flatten()
                .filter_map(|a| {
                    let key = std::str::from_utf8(a.key.as_ref()).ok()?.to_string();
                    let local = key.rsplit(':').next().unwrap_or("");
                    let replaced = ["distT", "distB", "distL", "distR", "behindDoc", "simplePos"].contains(&local);
                    let value = a.unescape_value().ok()?;
                    (!replaced).then(|| format!(r#"{}="{}""#, key, escape_xml(&value)))
                })
                .collect::<Vec<_>>(),
            _ => Vec::new(),
        };
        let has = |name: &str| kept.iter().any(|a| a.starts_with(&format!("{}=", name)));
        let mut attrs = vec![dist, r#"simplePos="0""#.to_string()];
        if !has("relativeHeight") {
            attrs.push(format!(r#"relativeHeight="{}""#, relative_height));
        }
        attrs.push(format!(r#"behindDoc="{}""#, u8::from(img.behind_text)));
        for (name, value) in [("locked", "0"), ("layoutInCell", "1"), ("allowOverlap", "1")] {
            if !has(name) {
                attrs.push(format!(r#"{}="{}""#, name, value));
            }
        }
        attrs.extend(kept);
        out.push_str(&format!("<wp:anchor {}>", attrs.join(" ")));
        out.push_str(r#"<wp:simplePos x="0" y="0"/>"#);
        out.push_str(&position_xml("H", &img.h_relative, img.h_align.as_deref(), img.h_offset));
        out.push_str(&position_xml("V", &img.v_relative, img.v_align.as_deref(), img.v_offset));
        out.push_str(&extent);
        out.push_str(effect);
        out.push_str(&wrap_xml(img, parts.wrap_polygon));
        out.push_str(&doc_pr);
        out.push_str(frame_pr);
        out.push_str(&graphic);
        if current.anchored {
            parts.extensions.iter().for_each(|raw| out.push_str(raw));
        }
        out.push_str("</wp:anchor>");
    }
    out.push_str("</w:drawing>");
    Ok(out)
}

fn position_xml(axis: &str, relative: &str, align: Option<&str>, offset: f64) -> String {
    let inner = match align {
        Some(a) => format!("<wp:align>{}</wp:align>", a),
        None => format!("<wp:posOffset>{}</wp:posOffset>", emu(offset)),
    };
    format!(r#"<wp:position{0} relativeFrom="{1}">{2}</wp:position{0}>"#, axis, relative, inner)
}

fn wrap_xml(img: &ImageRef, polygon: Option<(&str, &str)>) -> String {
    let side = if img.wrap_side.is_empty() { "bothSides" } else { img.wrap_side.as_str() };
    match img.wrap.as_str() {
        "square" => format!(r#"<wp:wrapSquare wrapText="{}"/>"#, side),
        "topAndBottom" => "<wp:wrapTopAndBottom/>".to_string(),
        kind @ ("tight" | "through") => {
            let element = if kind == "tight" { "wrapTight" } else { "wrapThrough" };
            match polygon {
                // Same kind: keep its contour, only the side changes
                Some((k, raw)) if k == kind => set_attr(raw, "wrapText", side),
                _ => format!(
                    concat!(
                        r#"<wp:{0} wrapText="{1}"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/>"#,
                        r#"<wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/>"#,
                        r#"<wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:{0}>"#
                    ),
                    element, side
                ),
            }
        }
        _ => "<wp:wrapNone/>".to_string(),
    }
}

/// Sets (or adds) attribute `name` on the first start tag of `xml`. `value` is already escaped.
fn set_attr(xml: &str, name: &str, value: &str) -> String {
    let tag_end = xml.find('>').unwrap_or(xml.len());
    let tag = &xml[..tag_end];
    let re = Regex::new(&format!(r#"\s{}="[^"]*""#, regex::escape(name))).expect("regex");
    if let Some(m) = re.find(tag) {
        format!("{} {}=\"{}\"{}", &xml[..m.start()], name, value, &xml[m.end()..])
    } else {
        let at = if tag.ends_with('/') { tag_end - 1 } else { tag_end };
        format!("{} {}=\"{}\"{}", &xml[..at], name, value, &xml[at..])
    }
}

/// `a:graphic` with the picture's own size (`a:xfrm/a:ext`) set to `cx`×`cy`
fn resize_graphic(graphic: &str, cx: i64, cy: i64) -> String {
    let re = Regex::new(r#"(<a:xfrm\b[^>]*>(?:\s*<a:off\b[^>]*/>)?\s*)<a:ext\b[^>]*/>"#).expect("regex");
    re.replace(graphic, |c: &regex::Captures| format!(r#"{}<a:ext cx="{}" cy="{}"/>"#, &c[1], cx, cy)).into_owned()
}

/// A z-order above every floating object of the document
fn next_relative_height(doc: &str) -> u64 {
    let re = Regex::new(r#"relativeHeight="(\d+)""#).expect("regex");
    re.captures_iter(doc).filter_map(|c| c[1].parse::<u64>().ok()).max().map_or(251_658_240, |h| h + 1024)
}

/// True when run `run` holds nothing but its properties once `removed` is taken out of it
fn run_is_empty_without(p_xml: &str, run: &Range<usize>, removed: &Range<usize>) -> bool {
    let mut rest = String::new();
    rest.push_str(&p_xml[run.start..removed.start]);
    rest.push_str(&p_xml[removed.end..run.end]);
    let re = Regex::new(r#"(?s)^<w:r\b[^>]*>\s*(<w:rPr\b.*?</w:rPr>|<w:rPr\b[^>]*/>)?\s*(<w:lastRenderedPageBreak/>\s*)?</w:r>$"#).expect("regex");
    re.is_match(rest.trim())
}

#[cfg(test)]
mod tests {
    use super::*;

    const INLINE: &str = concat!(
        r#"<w:p><w:r><w:t>Antes </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">"#,
        r#"<wp:extent cx="952500" cy="476250"/><wp:effectExtent l="0" t="0" r="0" b="0"/>"#,
        r#"<wp:docPr id="3" name="Imagen 3" descr="logo"/>"#,
        r#"<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="pic"><pic:pic xmlns:pic="pic">"#,
        r#"<pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill>"#,
        r#"<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm></pic:spPr>"#,
        r#"</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t>después</w:t></w:r></w:p>"#
    );

    fn parse(p: &str) -> ImageRef {
        drawing_spans(p).unwrap().remove(0).image
    }

    #[test]
    fn finds_pictures_and_their_runs() {
        let spans = drawing_spans(INLINE).unwrap();
        assert_eq!(spans.len(), 1);
        let s = &spans[0];
        assert!(INLINE[s.range.clone()].starts_with("<w:drawing>"));
        assert!(INLINE[s.run.clone().unwrap()].starts_with("<w:r><w:rPr><w:b/>"));
        assert_eq!(s.image.width, 100.0);
        assert_eq!(s.image.wrap, "inline");
        assert_eq!(s.image.doc_pr_id, 3);
        assert_eq!(s.image.alt, "logo");
    }

    #[test]
    fn inline_to_square_and_back() {
        let s = drawing_spans(INLINE).unwrap().remove(0);
        let update = ImageUpdate { wrap: Some("square".into()), width: Some(200.0), ..Default::default() };
        let target = merge_update(&s.image, &update, "center", (600.0, 900.0)).unwrap();
        let xml = rebuild_drawing(&INLINE[s.range.clone()], &s.image, &target, 251658240).unwrap();
        let p = format!("<w:p><w:r>{}</w:r></w:p>", xml);
        let img = parse(&p);
        assert!(img.anchored);
        assert_eq!(img.wrap, "square");
        assert_eq!(img.wrap_side, "bothSides");
        assert_eq!((img.width, img.height), (200.0, 100.0));
        assert_eq!(img.h_align.as_deref(), Some("center"), "a centered paragraph floats centered");
        assert_eq!(img.v_relative, "paragraph");
        assert_eq!(img.dist_left, DEFAULT_DIST_X);
        assert_eq!(img.rel_id, "rId9");
        assert!(xml.contains(r#"<a:ext cx="1905000" cy="952500"/>"#), "picture size follows: {}", xml);

        // Back in line with the text
        let back = merge_update(&img, &ImageUpdate { wrap: Some("inline".into()), ..Default::default() }, "left", (600.0, 900.0)).unwrap();
        let xml2 = rebuild_drawing(&xml, &img, &back, 0).unwrap();
        let img2 = parse(&format!("<w:p><w:r>{}</w:r></w:p>", xml2));
        assert!(!img2.anchored);
        assert_eq!(img2.wrap, "inline");
        assert!(!xml2.contains("positionH"));
    }

    #[test]
    fn positions_and_behind_text() {
        let s = drawing_spans(INLINE).unwrap().remove(0);
        let to_float = merge_update(&s.image, &ImageUpdate { wrap: Some("behind".into()), ..Default::default() }, "left", (600.0, 900.0)).unwrap();
        let xml = rebuild_drawing(&INLINE[s.range.clone()], &s.image, &to_float, 7).unwrap();
        let img = parse(&format!("<w:p><w:r>{}</w:r></w:p>", xml));
        assert!(img.behind_text);
        assert_eq!(img.wrap, "behind");

        let moved = ImageUpdate {
            h_relative: Some("page".into()),
            h_offset: Some(150.0),
            v_relative: Some("margin".into()),
            v_align: Some("bottom".into()),
            wrap: Some("topAndBottom".into()),
            distance: Some(6.0),
            ..Default::default()
        };
        let target = merge_update(&img, &moved, "left", (600.0, 900.0)).unwrap();
        let xml2 = rebuild_drawing(&xml, &img, &target, 9).unwrap();
        let img2 = parse(&format!("<w:p><w:r>{}</w:r></w:p>", xml2));
        assert_eq!(img2.wrap, "topAndBottom");
        assert!(!img2.behind_text);
        assert_eq!((img2.h_relative.as_str(), img2.h_offset, img2.h_align.as_deref()), ("page", 150.0, None));
        assert_eq!((img2.v_relative.as_str(), img2.v_align.as_deref()), ("margin", Some("bottom")));
        assert_eq!(img2.dist_top, 6.0);
        assert!(xml2.contains(r#"relativeHeight="7""#), "z-order of the existing anchor is kept");
    }

    #[test]
    fn rejects_bad_values() {
        let s = drawing_spans(INLINE).unwrap().remove(0);
        assert!(ImageUpdate { wrap: Some("diagonal".into()), ..Default::default() }.validate().is_err());
        assert!(ImageUpdate { width: Some(f64::NAN), ..Default::default() }.validate().is_err());
        // Inline pictures have no position
        let err = merge_update(&s.image, &ImageUpdate { h_offset: Some(5.0), ..Default::default() }, "left", (600.0, 900.0));
        assert!(err.is_err());
    }

    #[test]
    fn inline_pictures_fit_the_text_area() {
        let s = drawing_spans(INLINE).unwrap().remove(0);
        let big = merge_update(&s.image, &ImageUpdate { width: Some(1200.0), ..Default::default() }, "left", (600.0, 900.0)).unwrap();
        assert_eq!((big.width, big.height), (600.0, 300.0));
        let free = merge_update(&s.image, &ImageUpdate { width: Some(50.0), height: Some(80.0), ..Default::default() }, "left", (600.0, 900.0)).unwrap();
        assert_eq!((free.width, free.height), (50.0, 80.0));
    }

    #[test]
    fn empty_runs_go_with_the_picture() {
        let s = drawing_spans(INLINE).unwrap().remove(0);
        assert!(run_is_empty_without(INLINE, s.run.as_ref().unwrap(), &s.range));
        let shared = r#"<w:p><w:r><w:t>a</w:t><w:drawing>x</w:drawing></w:r></w:p>"#;
        let run = 5..shared.len() - 6;
        let drawing = shared.find("<w:drawing>").unwrap()..shared.find("</w:r>").unwrap();
        assert!(!run_is_empty_without(shared, &run, &drawing));
    }

    #[test]
    fn sets_attributes() {
        assert_eq!(set_attr(r#"<wp:docPr id="1" descr="a"/>"#, "descr", "b"), r#"<wp:docPr id="1" descr="b"/>"#);
        assert_eq!(set_attr(r#"<wp:docPr id="1"/>"#, "descr", "b"), r#"<wp:docPr id="1" descr="b"/>"#);
        assert_eq!(set_attr(r#"<wp:wrapTight wrapText="left"><x/></wp:wrapTight>"#, "wrapText", "right"), r#"<wp:wrapTight wrapText="right"><x/></wp:wrapTight>"#);
    }
}
