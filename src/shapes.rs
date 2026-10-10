//! Text boxes and simple shapes: DrawingML (`wps:wsp`, what current Word writes) and legacy
//! VML (`w:pict` with `v:shape` / `v:rect` / `v:textbox`, from older documents).
//!
//! Child module of `docx_parser`: a shape is read into the same `ImageRef` as a picture, so it
//! is positioned, wrapped and edited the same way; it carries a `ShapeStyle` (geometry, fill,
//! outline) and, for text boxes, the paragraphs of its `w:txbxContent`.

use super::*;
use crate::styles::ThemeFonts;

/// Word's default text box insets: 0.1" left and right, 0.05" top and bottom
const INSET_X: f64 = 9.6;
const INSET_Y: f64 = 4.8;

/// How a shape is drawn. Lengths in px; colors as "RRGGBB".
#[derive(Serialize, Deserialize, Debug, Clone, Default, PartialEq)]
pub struct ShapeStyle {
    /// "rect", "roundRect", "ellipse" or "line"
    pub geometry: String,
    #[serde(default)]
    pub fill: Option<String>,
    #[serde(default)]
    pub stroke: Option<String>,
    #[serde(default)]
    pub stroke_width: f64,
}

/// The text of a text box and how it sits inside the shape
#[derive(Serialize, Deserialize, Debug, Clone, Default)]
pub struct TextBox {
    pub paragraphs: Vec<ParagraphInfo>,
    pub inset_left: f64,
    pub inset_top: f64,
    pub inset_right: f64,
    pub inset_bottom: f64,
    /// Vertical alignment of the text: "t", "ctr" or "b"
    pub v_anchor: String,
}

impl PartialEq for TextBox {
    fn eq(&self, other: &Self) -> bool {
        self.inset_left == other.inset_left
            && self.inset_top == other.inset_top
            && self.inset_right == other.inset_right
            && self.inset_bottom == other.inset_bottom
            && self.v_anchor == other.v_anchor
            && self.paragraphs.len() == other.paragraphs.len()
            && self.paragraphs.iter().zip(&other.paragraphs).all(|(a, b)| a.text == b.text)
    }
}

impl TextBox {
    fn new() -> Self {
        TextBox {
            inset_left: INSET_X,
            inset_top: INSET_Y,
            inset_right: INSET_X,
            inset_bottom: INSET_Y,
            v_anchor: "t".into(),
            ..Default::default()
        }
    }

    /// Plain text of the box, paragraphs separated by line breaks
    pub fn text(&self) -> String {
        self.paragraphs.iter().map(|p| p.text.as_str()).collect::<Vec<_>>().join("\n")
    }
}

fn local_name(n: impl AsRef<[u8]>) -> String {
    utf8(&n).rsplit(':').next().unwrap_or("").to_string()
}

/// Reads the paragraphs of the `w:txbxContent` whose start tag was just consumed
fn parse_txbx_content(reader: &mut Reader<&[u8]>, styles: &StyleSheet) -> Vec<ParagraphInfo> {
    let mut paragraphs = Vec::new();
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name().as_ref().to_string();
                buf.clear();
                if tag_is(&name, "p") {
                    paragraphs.push(parse_paragraph_with(reader, 0, styles, None));
                } else if tag_is(&name, "tbl") || tag_is(&name, "sdtPr") {
                    // Tables inside text boxes are not laid out yet
                    let _ = reader.read_to_end_into(quick_xml::name::QName(&name), &mut Vec::new());
                }
                continue;
            }
            Ok(Event::Empty(e)) if tag_is(e.name().as_ref(), "p") => {
                paragraphs.push(resolve_empty_paragraph(styles));
            }
            Ok(Event::End(e)) if tag_is(e.name().as_ref(), "txbxContent") => break,
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    paragraphs
}

/// An empty `<w:p/>` with the document's default formatting
fn resolve_empty_paragraph(styles: &StyleSheet) -> ParagraphInfo {
    let xml = "<w:p></w:p>";
    let mut reader = Reader::from_str(xml);
    let _ = reader.read_event();
    parse_paragraph_with(&mut reader, 0, styles, None)
}

// ---------- DrawingML ----------

/// Color of a DrawingML color element (`a:srgbClr`, `a:schemeClr`, `a:prstClr`, `a:sysClr`)
/// with its `lumMod` / `lumOff` / `shade` / `tint` children applied, as "RRGGBB"
struct ColorBuilder {
    base: Option<(u8, u8, u8)>,
    lum_mod: Option<f64>,
    lum_off: Option<f64>,
    shade: Option<f64>,
    tint: Option<f64>,
}

impl ColorBuilder {
    fn start(e: &BytesStart, theme: &ThemeFonts) -> Option<Self> {
        let name = e.name();
        let n = name.as_ref();
        let val = get_attr_value(e, "val");
        let hex = if tag_is(n, "srgbClr") {
            val
        } else if tag_is(n, "sysClr") {
            get_attr_value(e, "lastClr").or(val.map(|v| if v == "window" { "FFFFFF".into() } else { "000000".into() }))
        } else if tag_is(n, "schemeClr") {
            val.and_then(|v| theme.color(&v))
        } else if tag_is(n, "prstClr") {
            val.map(|v| preset_color(&v).to_string())
        } else {
            return None;
        };
        Some(ColorBuilder { base: hex.as_deref().and_then(parse_hex), lum_mod: None, lum_off: None, shade: None, tint: None })
    }

    fn modifier(&mut self, e: &BytesStart) {
        let value = get_attr_i64(e, "val").map(|v| v as f64 / 100_000.0);
        let name = e.name();
        let n = name.as_ref();
        if tag_is(n, "lumMod") {
            self.lum_mod = value;
        } else if tag_is(n, "lumOff") {
            self.lum_off = value;
        } else if tag_is(n, "shade") {
            self.shade = value;
        } else if tag_is(n, "tint") {
            self.tint = value;
        }
    }

    fn finish(&self) -> Option<String> {
        let (mut r, mut g, mut b) = self.base?;
        if let Some(s) = self.shade {
            let f = |c: u8| (c as f64 * s).round().clamp(0.0, 255.0) as u8;
            (r, g, b) = (f(r), f(g), f(b));
        }
        if let Some(t) = self.tint {
            let f = |c: u8| (255.0 - (255.0 - c as f64) * t).round().clamp(0.0, 255.0) as u8;
            (r, g, b) = (f(r), f(g), f(b));
        }
        if self.lum_mod.is_some() || self.lum_off.is_some() {
            let (h, s, l) = rgb_to_hsl(r, g, b);
            let l = (l * self.lum_mod.unwrap_or(1.0) + self.lum_off.unwrap_or(0.0)).clamp(0.0, 1.0);
            (r, g, b) = hsl_to_rgb(h, s, l);
        }
        Some(format!("{:02X}{:02X}{:02X}", r, g, b))
    }
}

fn parse_hex(s: &str) -> Option<(u8, u8, u8)> {
    let s = s.trim_start_matches('#');
    if s.len() != 6 {
        return None;
    }
    let v = u32::from_str_radix(s, 16).ok()?;
    Some(((v >> 16) as u8, (v >> 8) as u8, v as u8))
}

fn preset_color(name: &str) -> &'static str {
    match name.to_ascii_lowercase().as_str() {
        "white" => "FFFFFF",
        "red" => "FF0000",
        "green" => "008000",
        "blue" => "0000FF",
        "yellow" => "FFFF00",
        "gray" | "grey" => "808080",
        "silver" => "C0C0C0",
        "navy" => "000080",
        "maroon" => "800000",
        "purple" => "800080",
        "teal" => "008080",
        "orange" => "FFA500",
        "lime" => "00FF00",
        "aqua" | "cyan" => "00FFFF",
        "fuchsia" | "magenta" => "FF00FF",
        "olive" => "808000",
        _ => "000000",
    }
}

fn rgb_to_hsl(r: u8, g: u8, b: u8) -> (f64, f64, f64) {
    let (r, g, b) = (r as f64 / 255.0, g as f64 / 255.0, b as f64 / 255.0);
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    if (max - min).abs() < f64::EPSILON {
        return (0.0, 0.0, l);
    }
    let d = max - min;
    let s = if l > 0.5 { d / (2.0 - max - min) } else { d / (max + min) };
    let h = if max == r {
        (g - b) / d + if g < b { 6.0 } else { 0.0 }
    } else if max == g {
        (b - r) / d + 2.0
    } else {
        (r - g) / d + 4.0
    };
    (h / 6.0, s, l)
}

fn hsl_to_rgb(h: f64, s: f64, l: f64) -> (u8, u8, u8) {
    let to = |v: f64| (v * 255.0).round().clamp(0.0, 255.0) as u8;
    if s == 0.0 {
        return (to(l), to(l), to(l));
    }
    let q = if l < 0.5 { l * (1.0 + s) } else { l + s - l * s };
    let p = 2.0 * l - q;
    let hue = |mut t: f64| {
        if t < 0.0 {
            t += 1.0;
        }
        if t > 1.0 {
            t -= 1.0;
        }
        if t < 1.0 / 6.0 {
            p + (q - p) * 6.0 * t
        } else if t < 0.5 {
            q
        } else if t < 2.0 / 3.0 {
            p + (q - p) * (2.0 / 3.0 - t) * 6.0
        } else {
            p
        }
    };
    (to(hue(h + 1.0 / 3.0)), to(hue(h)), to(hue(h - 1.0 / 3.0)))
}

/// What a DrawingML shape (`wps:wsp`) says about its look and its text
pub(super) struct WspResult {
    pub style: ShapeStyle,
    pub text_box: Option<TextBox>,
}

/// Reads the `wps:wsp` whose start tag was just consumed
pub(super) fn parse_wsp(reader: &mut Reader<&[u8]>, styles: &StyleSheet) -> WspResult {
    let theme = &styles.theme;
    let mut path: Vec<String> = vec!["wsp".into()];
    let mut geometry = "rect".to_string();
    // None = not said; Some(None) = explicitly none; Some(Some(c)) = color
    let mut fill: Option<Option<String>> = None;
    let mut stroke: Option<Option<String>> = None;
    let mut stroke_width: Option<f64> = None;
    let mut style_fill: Option<String> = None;
    let mut style_line: Option<String> = None;
    // idx 0 means "no fill" / "no line" from the theme, whatever color follows
    let mut no_style_fill = false;
    let mut no_style_line = false;
    let mut color: Option<(ColorBuilder, Vec<String>)> = None;
    let mut text_box: Option<TextBox> = None;
    let mut body = TextBox::new();
    let mut buf = Vec::new();

    let target = |p: &[String]| -> Option<&'static str> {
        let has = |s: &str| p.iter().any(|x| x == s);
        if has("style") {
            if has("fillRef") {
                Some("style_fill")
            } else if has("lnRef") {
                Some("style_line")
            } else {
                None
            }
        } else if has("spPr") && has("solidFill") {
            if has("ln") {
                Some("stroke")
            } else if !has("effectLst") && !has("gradFill") {
                Some("fill")
            } else {
                None
            }
        } else {
            None
        }
    };

    loop {
        let event = reader.read_event_into(&mut buf);
        let (e, is_start) = match event {
            Ok(Event::Start(e)) => (e.into_owned(), true),
            Ok(Event::Empty(e)) => (e.into_owned(), false),
            Ok(Event::End(e)) => {
                let local = local_name(e.name().as_ref());
                if let Some((builder, at)) = &color {
                    if at.len() == path.len() && path.last() == Some(&local) {
                        let value = builder.finish();
                        match target(&path) {
                            Some("fill") => fill = Some(value),
                            Some("stroke") => stroke = Some(value),
                            Some("style_fill") => style_fill = value,
                            Some("style_line") => style_line = value,
                            _ => {}
                        }
                        color = None;
                    }
                }
                path.pop();
                if local == "wsp" {
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
        buf.clear();
        let name = e.name();
        let n = name.as_ref();
        let local = local_name(n);

        if is_start && local == "txbxContent" {
            let mut tb = std::mem::replace(&mut body, TextBox::new());
            tb.paragraphs = parse_txbx_content(reader, styles);
            text_box = Some(tb);
            continue;
        }
        if let Some((builder, _)) = color.as_mut() {
            builder.modifier(&e);
        } else if let Some(builder) = ColorBuilder::start(&e, theme) {
            if is_start {
                path.push(local.clone());
                color = Some((builder, path.clone()));
                continue;
            }
            let value = builder.finish();
            path.push(local.clone());
            match target(&path) {
                Some("fill") => fill = Some(value),
                Some("stroke") => stroke = Some(value),
                Some("style_fill") => style_fill = value,
                Some("style_line") => style_line = value,
                _ => {}
            }
            path.pop();
            continue;
        }

        let in_sppr = path.iter().any(|x| x == "spPr");
        match local.as_str() {
            "prstGeom" if in_sppr => {
                geometry = get_attr_value(&e, "prst").unwrap_or_else(|| "rect".into());
            }
            "noFill" if in_sppr => {
                if path.last().map(String::as_str) == Some("ln") {
                    stroke = Some(None);
                } else if path.last().map(String::as_str) == Some("spPr") {
                    fill = Some(None);
                }
            }
            "ln" if in_sppr || path.last().map(String::as_str) == Some("spPr") => {
                if let Some(w) = get_attr_i64(&e, "w") {
                    stroke_width = Some(w as f64 / EMU_PER_PX);
                }
            }
            "fillRef" | "lnRef" => {
                if get_attr_i64(&e, "idx") == Some(0) {
                    if local == "fillRef" {
                        no_style_fill = true;
                    } else {
                        no_style_line = true;
                    }
                }
            }
            "bodyPr" => {
                let inset = |name: &str, default: f64| get_attr_i64(&e, name).map_or(default, |v| v as f64 / EMU_PER_PX);
                body.inset_left = inset("lIns", INSET_X);
                body.inset_top = inset("tIns", INSET_Y);
                body.inset_right = inset("rIns", INSET_X);
                body.inset_bottom = inset("bIns", INSET_Y);
                body.v_anchor = get_attr_value(&e, "anchor").unwrap_or_else(|| "t".into());
                if let Some(tb) = text_box.as_mut() {
                    // bodyPr comes after the text in wps:wsp
                    tb.inset_left = body.inset_left;
                    tb.inset_top = body.inset_top;
                    tb.inset_right = body.inset_right;
                    tb.inset_bottom = body.inset_bottom;
                    tb.v_anchor = body.v_anchor.clone();
                }
            }
            _ => {}
        }
        if is_start {
            path.push(local);
        }
    }

    let style = ShapeStyle {
        geometry,
        fill: fill.unwrap_or(if no_style_fill { None } else { style_fill }),
        stroke: stroke.unwrap_or(if no_style_line { None } else { style_line }),
        stroke_width: stroke_width.unwrap_or(0.75).max(0.0),
    };
    WspResult { style, text_box }
}

// ---------- VML ----------

/// A VML length ("12pt", "1.5in", "2cm", "40px", "0") in px
fn vml_length(v: &str) -> Option<f64> {
    let v = v.trim();
    let split = v.find(|c: char| c.is_ascii_alphabetic() || c == '%').unwrap_or(v.len());
    let (num, unit) = v.split_at(split);
    let n: f64 = num.trim().parse().ok()?;
    Some(match unit.trim() {
        "pt" => n * 96.0 / 72.0,
        "in" => n * 96.0,
        "cm" => n * 96.0 / 2.54,
        "mm" => n * 96.0 / 25.4,
        "pc" => n * 16.0,
        "emu" => n / EMU_PER_PX,
        _ => n,
    })
}

/// A VML color ("#1F4E79", "red", "white [3212]") as "RRGGBB"
fn vml_color(v: &str) -> Option<String> {
    let v = v.split_whitespace().next()?.trim();
    if let Some(hex) = v.strip_prefix('#') {
        return match hex.len() {
            6 => Some(hex.to_ascii_uppercase()),
            3 => Some(hex.chars().flat_map(|c| [c, c]).collect::<String>().to_ascii_uppercase()),
            _ => None,
        };
    }
    Some(preset_color(v).to_string())
}

fn vml_on(v: Option<String>, default: bool) -> bool {
    match v.as_deref() {
        Some("f") | Some("false") | Some("0") | Some("off") => false,
        Some(_) => true,
        None => default,
    }
}

/// Reads the `w:pict` whose start tag was just consumed: the first shape with a text box or
/// a picture (`v:imagedata`), positioned from its CSS-like `style`
pub(super) fn parse_vml_pict(reader: &mut Reader<&[u8]>, styles: &StyleSheet) -> Option<ImageRef> {
    let mut img: Option<ImageRef> = None;
    let mut depth_in_shape = 0usize;
    let mut buf = Vec::new();
    loop {
        let event = reader.read_event_into(&mut buf);
        let (e, is_start) = match event {
            Ok(Event::Start(e)) => (e.into_owned(), true),
            Ok(Event::Empty(e)) => (e.into_owned(), false),
            Ok(Event::End(e)) => {
                if tag_is(e.name().as_ref(), "pict") {
                    break;
                }
                depth_in_shape = depth_in_shape.saturating_sub(1);
                buf.clear();
                continue;
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {
                buf.clear();
                continue;
            }
        };
        buf.clear();
        let name = e.name();
        let n = name.as_ref();
        let local = local_name(n);

        // Templates and groups are not drawn
        if is_start && (local == "shapetype" || local == "group") {
            let end = n.to_string();
            let _ = reader.read_to_end_into(quick_xml::name::QName(&end), &mut Vec::new());
            continue;
        }
        let is_shape = matches!(local.as_str(), "shape" | "rect" | "roundrect" | "oval");
        if is_shape && img.is_none() {
            img = Some(vml_shape(&e, &local));
            if is_start {
                depth_in_shape = 1;
            }
            continue;
        }
        let Some(current) = img.as_mut() else {
            if is_start {
                let end = n.to_string();
                let _ = reader.read_to_end_into(quick_xml::name::QName(&end), &mut Vec::new());
            }
            continue;
        };
        match local.as_str() {
            "txbxContent" if is_start => {
                let mut tb = current.text_box.take().unwrap_or_else(TextBox::new);
                tb.paragraphs = parse_txbx_content(reader, styles);
                current.text_box = Some(tb);
                continue;
            }
            "textbox" => {
                let mut tb = current.text_box.take().unwrap_or_else(TextBox::new);
                if let Some(inset) = get_attr_value(&e, "inset") {
                    let parts: Vec<Option<f64>> = inset.split(',').map(vml_length).collect();
                    let get = |i: usize, d: f64| parts.get(i).copied().flatten().unwrap_or(d);
                    tb.inset_left = get(0, INSET_X);
                    tb.inset_top = get(1, INSET_Y);
                    tb.inset_right = get(2, INSET_X);
                    tb.inset_bottom = get(3, INSET_Y);
                }
                let style = get_attr_value(&e, "style").unwrap_or_default();
                if style.contains("v-text-anchor:middle") {
                    tb.v_anchor = "ctr".into();
                } else if style.contains("v-text-anchor:bottom") {
                    tb.v_anchor = "b".into();
                }
                current.text_box = Some(tb);
            }
            "imagedata" => {
                if let Some(id) = get_attr_value(&e, "id").or_else(|| get_attr_value(&e, "relid")) {
                    current.rel_id = id;
                    current.shape = None;
                }
            }
            "wrap" => {
                current.wrap = match get_attr_value(&e, "type").as_deref() {
                    Some("square") => "square",
                    Some("tight") => "tight",
                    Some("through") => "through",
                    Some("topAndBottom") => "topAndBottom",
                    _ => if current.behind_text { "behind" } else { "inFront" },
                }
                .into();
                if matches!(current.wrap.as_str(), "square" | "tight" | "through") {
                    current.wrap_side = match get_attr_value(&e, "side").as_deref() {
                        Some("left") => "left",
                        Some("right") => "right",
                        Some("largest") => "largest",
                        _ => "bothSides",
                    }
                    .into();
                }
                if !current.anchored {
                    current.wrap = "inline".into();
                }
            }
            "fill" => {
                if let Some(shape) = current.shape.as_mut() {
                    if !vml_on(get_attr_value(&e, "on"), true) {
                        shape.fill = None;
                    } else if let Some(c) = get_attr_value(&e, "color").as_deref().and_then(vml_color) {
                        shape.fill = Some(c);
                    }
                }
            }
            "stroke" => {
                if let Some(shape) = current.shape.as_mut() {
                    if !vml_on(get_attr_value(&e, "on"), true) {
                        shape.stroke = None;
                    } else if let Some(c) = get_attr_value(&e, "color").as_deref().and_then(vml_color) {
                        shape.stroke = Some(c);
                    }
                }
            }
            _ => {}
        }
        if is_start {
            depth_in_shape += 1;
        }
    }
    let img = img?;
    let drawable = !img.rel_id.is_empty() || img.text_box.is_some() || img.shape.as_ref().is_some_and(|s| s.fill.is_some() || s.stroke.is_some());
    (drawable && img.width > 0.0 && img.height > 0.0).then_some(img)
}

/// Position, size and look of a VML shape element (`v:shape`, `v:rect`, …)
fn vml_shape(e: &BytesStart, local: &str) -> ImageRef {
    let style = get_attr_value(e, "style").unwrap_or_default();
    let mut css: HashMap<String, String> = HashMap::new();
    for decl in style.split(';') {
        if let Some((k, v)) = decl.split_once(':') {
            css.insert(k.trim().to_ascii_lowercase(), v.trim().to_string());
        }
    }
    let len = |k: &str| css.get(k).and_then(|v| vml_length(v));
    let anchored = css.get("position").map(String::as_str) == Some("absolute");
    let z_index = css.get("z-index").and_then(|z| z.trim().parse::<i64>().ok()).unwrap_or(0);
    let behind = z_index < 0;

    let h_relative = match css.get("mso-position-horizontal-relative").map(String::as_str) {
        Some("margin") => "margin",
        Some("page") => "page",
        Some("char") => "character",
        Some("left-margin-area") => "leftMargin",
        Some("right-margin-area") => "rightMargin",
        Some("inner-margin-area") => "insideMargin",
        Some("outer-margin-area") => "outsideMargin",
        _ => "column",
    };
    let v_relative = match css.get("mso-position-vertical-relative").map(String::as_str) {
        Some("margin") => "margin",
        Some("page") => "page",
        Some("line") => "line",
        Some("top-margin-area") => "topMargin",
        Some("bottom-margin-area") => "bottomMargin",
        _ => "paragraph",
    };
    let align = |k: &str| css.get(k).filter(|v| v.as_str() != "absolute").cloned();

    let filled = vml_on(get_attr_value(e, "filled"), true);
    let stroked = vml_on(get_attr_value(e, "stroked"), true);
    let shape = ShapeStyle {
        geometry: match local {
            "roundrect" => "roundRect",
            "oval" => "ellipse",
            _ => "rect",
        }
        .into(),
        fill: filled.then(|| get_attr_value(e, "fillcolor").as_deref().and_then(vml_color).unwrap_or_else(|| "FFFFFF".into())),
        stroke: stroked.then(|| get_attr_value(e, "strokecolor").as_deref().and_then(vml_color).unwrap_or_else(|| "000000".into())),
        stroke_width: get_attr_value(e, "strokeweight").as_deref().and_then(vml_length).unwrap_or(1.0),
    };

    ImageRef {
        width: len("width").unwrap_or(0.0),
        height: len("height").unwrap_or(0.0),
        anchored,
        behind_text: anchored && behind,
        wrap: if !anchored { "inline" } else if behind { "behind" } else { "inFront" }.into(),
        h_relative: h_relative.into(),
        h_offset: len("margin-left").or_else(|| len("left")).unwrap_or(0.0),
        h_align: align("mso-position-horizontal"),
        v_relative: v_relative.into(),
        v_offset: len("margin-top").or_else(|| len("top")).unwrap_or(0.0),
        v_align: align("mso-position-vertical"),
        dist_left: len("mso-wrap-distance-left").unwrap_or(12.0),
        dist_right: len("mso-wrap-distance-right").unwrap_or(12.0),
        dist_top: len("mso-wrap-distance-top").unwrap_or(0.0),
        dist_bottom: len("mso-wrap-distance-bottom").unwrap_or(0.0),
        alt: get_attr_value(e, "alt").unwrap_or_default(),
        shape: Some(shape),
        vml: true,
        z_order: z_index,
        ..Default::default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn colors_from_themes_and_modifiers() {
        let mut theme = ThemeFonts::default();
        theme.colors.insert("accent1".into(), "4472C4".into());
        theme.colors.insert("lt1".into(), "FFFFFF".into());
        let xml = r#"<a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr>"#;
        let mut reader = Reader::from_str(xml);
        let Ok(Event::Start(e)) = reader.read_event() else { panic!() };
        let mut b = ColorBuilder::start(&e, &theme).unwrap();
        let Ok(Event::Empty(m)) = reader.read_event() else { panic!() };
        b.modifier(&m);
        // accent1 at 50 % luminance: Word's "darker 50 %"
        assert_eq!(b.finish().as_deref(), Some("203864"));
        assert_eq!(vml_color("#f00"), Some("FF0000".into()));
        assert_eq!(vml_color("white [3212]"), Some("FFFFFF".into()));
        assert_eq!(vml_length("72pt"), Some(96.0));
        assert_eq!(vml_length("1in"), Some(96.0));
    }
}
