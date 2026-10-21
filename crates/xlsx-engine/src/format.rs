//! Basic cell formatting: font (bold, italic, underline, strike, color, size, name), fill,
//! borders, number format and alignment, column widths and merged cells.
//!
//! A cell's look is its cell format (`<xf>` in `xl/styles.xml`), which points to a font, a fill,
//! a border and a number format. Formatting a cell derives a new cell format from the one it has
//! (so whatever is not changed stays) and reuses an identical one when the stylesheet already
//! has it. New fonts, fills, borders and formats are appended; nothing existing is rewritten.

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use ooxml_core::xml::{element_ranges, escape_xml, get_attr_value, prefix_of, remove_attr, root_element, set_attr, utf8};

use crate::cell_ref::{CellRange, CellRef};
use crate::sheet::SheetXml;
use crate::view::SheetLayout;
use crate::workbook::{first_child_named, Workbook};

/// Most cells one formatting call may cover
pub const MAX_FORMAT_CELLS: u64 = 200_000;

const STYLES_TYPE: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

/// Which sides of the cells of a range get a border
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Edges {
    /// Every side of every cell
    All,
    /// The outside of the range
    Outline,
    Top,
    Bottom,
    Left,
    Right,
    /// The lines between cells, not the outside
    Inside,
}

impl Edges {
    pub fn parse(text: &str) -> Result<Edges, String> {
        Ok(match text {
            "all" => Edges::All,
            "outline" => Edges::Outline,
            "top" => Edges::Top,
            "bottom" => Edges::Bottom,
            "left" => Edges::Left,
            "right" => Edges::Right,
            "inside" => Edges::Inside,
            other => return Err(format!("Bordes '{}' no válidos: usa all, outline, top, bottom, left, right o inside.", other)),
        })
    }
}

/// A border to draw: line style (`thin`, `medium`, `thick`, `dashed`, `dotted`, `double`…,
/// `None` to remove the lines) and color
#[derive(Debug, Clone, PartialEq)]
pub struct BorderSpec {
    pub style: Option<String>,
    pub color: Option<String>,
    pub edges: Edges,
}

const BORDER_STYLES: &[&str] = &[
    "thin", "medium", "thick", "dashed", "dotted", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot",
    "mediumDashDotDot", "slantDashDot",
];

/// What to change in the look of cells; `None` leaves that part as it is
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Format {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strike: Option<bool>,
    /// `RRGGBB`
    pub font_color: Option<String>,
    /// Points
    pub font_size: Option<f64>,
    pub font_name: Option<String>,
    /// Background `RRGGBB`; `Some(None)` removes the fill
    pub fill: Option<Option<String>>,
    /// Excel format code: `#,##0.00`, `0%`, `dd/mm/yyyy`, `"$"#,##0`…
    pub number_format: Option<String>,
    /// `left`, `center`, `right`, `justify`, `general`…
    pub horizontal: Option<String>,
    /// `top`, `center`, `bottom`
    pub vertical: Option<String>,
    pub wrap: Option<bool>,
    pub border: Option<BorderSpec>,
}

/// `#RGB`, `#RRGGBB` or `RRGGBB` → `FFRRGGBB` (ARGB as Excel writes it)
pub fn argb(color: &str) -> Result<String, String> {
    let hex = color.trim().trim_start_matches('#');
    let full = match hex.len() {
        3 => hex.chars().flat_map(|c| [c, c]).collect(),
        6 => hex.to_string(),
        8 => return Ok(hex.to_uppercase()),
        _ => String::new(),
    };
    if full.len() != 6 || !full.chars().all(|c| c.is_ascii_hexdigit()) {
        return Err(format!("Color '{}' no válido: usa #RRGGBB.", color));
    }
    Ok(format!("FF{}", full.to_uppercase()))
}

impl Format {
    fn check(&self) -> Result<(), String> {
        if let Some(c) = &self.font_color {
            argb(c)?;
        }
        if let Some(Some(c)) = &self.fill {
            argb(c)?;
        }
        if let Some(size) = self.font_size {
            if !(1.0..=409.0).contains(&size) {
                return Err("El tamaño de letra debe estar entre 1 y 409 puntos.".to_string());
            }
        }
        if let Some(h) = &self.horizontal {
            if !["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"].contains(&h.as_str()) {
                return Err(format!("Alineación horizontal '{}' no válida.", h));
            }
        }
        if let Some(v) = &self.vertical {
            if !["top", "center", "bottom", "justify", "distributed"].contains(&v.as_str()) {
                return Err(format!("Alineación vertical '{}' no válida.", v));
            }
        }
        if let Some(BorderSpec { style: Some(s), color, .. }) = &self.border {
            if !BORDER_STYLES.contains(&s.as_str()) {
                return Err(format!("Estilo de borde '{}' no válido: usa thin, medium, thick, dashed, dotted o double.", s));
            }
            if let Some(c) = color {
                argb(c)?;
            }
        }
        if self.number_format.as_deref().is_some_and(|f| f.trim().is_empty()) {
            return Err("El formato de número está vacío.".to_string());
        }
        Ok(())
    }

    fn touches_font(&self) -> bool {
        self.bold.is_some()
            || self.italic.is_some()
            || self.underline.is_some()
            || self.strike.is_some()
            || self.font_color.is_some()
            || self.font_size.is_some()
            || self.font_name.is_some()
    }

    fn touches_alignment(&self) -> bool {
        self.horizontal.is_some() || self.vertical.is_some() || self.wrap.is_some()
    }
}

// ---------- Elements ----------

/// An element split into its start tag (without the closing `>` or `/>`), its name and its
/// children as (local name, text)
struct Element {
    start: String,
    name: String,
    children: Vec<(String, String)>,
}

impl Element {
    fn parse(xml: &str) -> Result<Element, String> {
        let mut reader = Reader::from_str(xml);
        let mut depth = 0usize;
        let mut start = None;
        let mut name = String::new();
        let mut children = Vec::new();
        let mut child_start = 0;
        loop {
            let before = reader.buffer_position() as usize;
            let event = reader.read_event().map_err(|e| format!("Estilo no válido: {}", e))?;
            let after = reader.buffer_position() as usize;
            match event {
                Event::Start(e) | Event::Empty(e) if depth == 0 => {
                    name = utf8(e.name().as_ref()).to_string();
                    let tag = &xml[before..after];
                    let self_closing = tag.ends_with("/>");
                    start = Some(tag.trim_end_matches("/>").trim_end_matches('>').to_string());
                    if self_closing {
                        break;
                    }
                    depth = 1;
                }
                Event::Start(_) if depth == 1 => {
                    child_start = before;
                    depth = 2;
                }
                Event::Empty(e) if depth == 1 => {
                    children.push((local(utf8(e.name().as_ref())), xml[before..after].to_string()));
                }
                Event::Start(_) => depth += 1,
                Event::End(_) if depth == 2 => {
                    let open = &xml[child_start..];
                    let child_name = open[1..].split(|c: char| c.is_whitespace() || c == '>' || c == '/').next().unwrap_or("");
                    children.push((local(child_name), xml[child_start..after].to_string()));
                    depth = 1;
                }
                Event::End(_) if depth == 1 => break,
                Event::End(_) => depth -= 1,
                Event::Eof => break,
                _ => {}
            }
        }
        let start = start.ok_or("Elemento de estilo vacío.")?;
        Ok(Element { start, name, children })
    }

    fn get(&self, local_name: &str) -> Option<&str> {
        self.children.iter().find(|(n, _)| n == local_name).map(|(_, t)| t.as_str())
    }

    fn set(&mut self, local_name: &str, value: Option<String>) {
        self.children.retain(|(n, _)| n != local_name);
        if let Some(v) = value {
            self.children.push((local_name.to_string(), v));
        }
    }

    /// The element with its children in `order` (others after them, as they were)
    fn write(&self, order: &[&str]) -> String {
        let mut kids: Vec<&(String, String)> = Vec::new();
        for name in order {
            kids.extend(self.children.iter().filter(|(n, _)| n == name));
        }
        kids.extend(self.children.iter().filter(|(n, _)| !order.contains(&n.as_str())));
        if kids.is_empty() {
            format!("{}/>", self.start)
        } else {
            format!("{}>{}</{}>", self.start, kids.iter().map(|(_, t)| t.as_str()).collect::<String>(), self.name)
        }
    }
}

fn local(name: &str) -> String {
    name.rsplit(':').next().unwrap_or(name).to_string()
}

const FONT_ORDER: &[&str] = &["b", "i", "strike", "condense", "extend", "outline", "shadow", "u", "vertAlign", "sz", "color", "name", "family", "charset", "scheme"];
const BORDER_ORDER: &[&str] = &["start", "left", "end", "right", "top", "bottom", "diagonal", "vertical", "horizontal"];
const XF_ORDER: &[&str] = &["alignment", "protection", "extLst"];

/// Built-in number formats by code
fn builtin_format_id(code: &str) -> Option<u32> {
    Some(match code {
        "General" => 0,
        "0" => 1,
        "0.00" => 2,
        "#,##0" => 3,
        "#,##0.00" => 4,
        "0%" => 9,
        "0.00%" => 10,
        "0.00E+00" => 11,
        "# ?/?" => 12,
        "d-mmm-yy" => 15,
        "d-mmm" => 16,
        "mmm-yy" => 17,
        "h:mm AM/PM" => 18,
        "h:mm:ss AM/PM" => 19,
        "h:mm" => 20,
        "h:mm:ss" => 21,
        "mm:ss" => 45,
        "@" => 49,
        _ => return None,
    })
}

// ---------- Stylesheet ----------

/// A section of the stylesheet (`<fonts>`, `<cellXfs>`…) with the elements appended to it
struct Section {
    items: Vec<String>,
    added: Vec<String>,
}

impl Section {
    fn id_of(&mut self, element: String) -> usize {
        if let Some(i) = self.items.iter().position(|e| *e == element) {
            return i;
        }
        if let Some(i) = self.added.iter().position(|e| *e == element) {
            return self.items.len() + i;
        }
        self.added.push(element);
        self.items.len() + self.added.len() - 1
    }

    fn get(&self, i: usize) -> Option<&str> {
        self.items.get(i).or_else(|| self.added.get(i - self.items.len().min(i))).map(String::as_str)
    }
}

struct StyleBook {
    xml: String,
    prefix: String,
    num_fmts: HashMap<u32, String>,
    new_formats: Vec<(u32, String)>,
    sections: HashMap<&'static str, Section>,
    /// Derived cell formats: (base xf, per-cell border sides) → new xf
    cache: HashMap<(u32, u8), u32>,
}

const SECTIONS: [(&str, &str); 4] = [("fonts", "font"), ("fills", "fill"), ("borders", "border"), ("cellXfs", "xf")];

impl StyleBook {
    fn parse(xml: &str) -> Result<StyleBook, String> {
        let (_, root) = root_element(xml).ok_or("La hoja de estilos está vacía.")?;
        let prefix = prefix_of(&root).to_string();
        let mut sections = HashMap::new();
        for (section, item) in SECTIONS {
            let range = element_ranges(xml, section, |_| true)?
                .into_iter()
                .next()
                .ok_or_else(|| format!("La hoja de estilos no tiene <{}>.", section))?;
            let inner = &xml[range];
            let items = element_ranges(inner, item, |_| true)?.into_iter().map(|r| inner[r].to_string()).collect();
            sections.insert(section, Section { items, added: Vec::new() });
        }
        let mut num_fmts = HashMap::new();
        let mut reader = Reader::from_str(xml);
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) if utf8(e.name().as_ref()).ends_with("numFmt") => {
                    if let (Some(id), Some(code)) = (get_attr_value(&e, "numFmtId").and_then(|v| v.parse().ok()), get_attr_value(&e, "formatCode")) {
                        num_fmts.insert(id, code);
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        Ok(StyleBook { xml: xml.to_string(), prefix, num_fmts, new_formats: Vec::new(), sections, cache: HashMap::new() })
    }

    fn section(&mut self, name: &'static str) -> &mut Section {
        self.sections.get_mut(name).expect("parsed sections")
    }

    fn format_id(&mut self, code: &str) -> u32 {
        if let Some(id) = builtin_format_id(code) {
            return id;
        }
        if let Some((id, _)) = self.num_fmts.iter().find(|(_, c)| c.as_str() == code) {
            return *id;
        }
        let id = self.num_fmts.keys().copied().max().unwrap_or(163).max(163) + 1;
        self.num_fmts.insert(id, code.to_string());
        self.new_formats.push((id, code.to_string()));
        id
    }

    /// The cell format `base` with `format` applied; `sides` says which border sides of this
    /// cell the format's border covers (bits: left, right, top, bottom)
    fn derive(&mut self, base: u32, format: &Format, sides: u8) -> Result<u32, String> {
        if let Some(id) = self.cache.get(&(base, sides)) {
            return Ok(*id);
        }
        let p = self.prefix.clone();
        let xfs = &self.sections["cellXfs"];
        let xf_text = xfs.get(base as usize).or_else(|| xfs.get(0)).map(str::to_string).ok_or("La hoja de estilos no tiene formatos de celda.")?;
        let mut xf = Element::parse(&xf_text)?;
        let attr = |el: &str, name: &str| -> usize {
            let mut reader = Reader::from_str(el);
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => get_attr_value(&e, name).and_then(|v| v.trim().parse().ok()).unwrap_or(0),
                _ => 0,
            }
        };

        if format.touches_font() {
            let font_text = self.section("fonts").get(attr(&xf_text, "fontId")).unwrap_or("<font/>").to_string();
            let mut font = Element::parse(&font_text)?;
            let flag = |on: bool, name: &str| on.then(|| format!("<{}{}/>", p, name));
            if let Some(b) = format.bold {
                font.set("b", flag(b, "b"));
            }
            if let Some(i) = format.italic {
                font.set("i", flag(i, "i"));
            }
            if let Some(s) = format.strike {
                font.set("strike", flag(s, "strike"));
            }
            if let Some(u) = format.underline {
                font.set("u", flag(u, "u"));
            }
            if let Some(size) = format.font_size {
                font.set("sz", Some(format!(r#"<{}sz val="{}"/>"#, p, size)));
            }
            if let Some(color) = &format.font_color {
                font.set("color", Some(format!(r#"<{}color rgb="{}"/>"#, p, argb(color)?)));
            }
            if let Some(name) = &format.font_name {
                font.set("name", Some(format!(r#"<{}name val="{}"/>"#, p, escape_xml(name))));
                // A theme font scheme would override the name
                font.set("scheme", None);
            }
            let id = self.section("fonts").id_of(font.write(FONT_ORDER));
            xf.start = with_attrs(&xf.start, &[("fontId", id.to_string()), ("applyFont", "1".to_string())]);
        }
        if let Some(fill) = &format.fill {
            let text = match fill {
                Some(color) => format!(
                    r#"<{p}fill><{p}patternFill patternType="solid"><{p}fgColor rgb="{}"/><{p}bgColor indexed="64"/></{p}patternFill></{p}fill>"#,
                    argb(color)?,
                    p = p
                ),
                None => format!(r#"<{p}fill><{p}patternFill patternType="none"/></{p}fill>"#, p = p),
            };
            let id = self.section("fills").id_of(text);
            xf.start = with_attrs(&xf.start, &[("fillId", id.to_string()), ("applyFill", "1".to_string())]);
        }
        if let (Some(border), true) = (&format.border, sides != 0) {
            let border_text = self.section("borders").get(attr(&xf_text, "borderId")).unwrap_or("<border/>").to_string();
            let mut el = Element::parse(&border_text)?;
            for (bit, side) in [(1, "left"), (2, "right"), (4, "top"), (8, "bottom")] {
                if sides & bit == 0 {
                    continue;
                }
                let text = match &border.style {
                    Some(style) => {
                        let color = match &border.color {
                            Some(c) => format!(r#"<{}color rgb="{}"/>"#, p, argb(c)?),
                            None => format!(r#"<{}color indexed="64"/>"#, p),
                        };
                        format!(r#"<{p}{side} style="{}">{}</{p}{side}>"#, style, color, p = p, side = side)
                    }
                    None => format!("<{}{}/>", p, side),
                };
                el.set(side, Some(text));
            }
            for side in ["left", "right", "top", "bottom", "diagonal"] {
                if el.get(side).is_none() {
                    el.set(side, Some(format!("<{}{}/>", p, side)));
                }
            }
            let id = self.section("borders").id_of(el.write(BORDER_ORDER));
            xf.start = with_attrs(&xf.start, &[("borderId", id.to_string()), ("applyBorder", "1".to_string())]);
        }
        if let Some(code) = &format.number_format {
            let id = self.format_id(code);
            xf.start = with_attrs(&xf.start, &[("numFmtId", id.to_string()), ("applyNumberFormat", "1".to_string())]);
        }
        if format.touches_alignment() {
            let mut alignment = xf.get("alignment").unwrap_or("").to_string();
            if alignment.is_empty() {
                alignment = format!("<{}alignment/>", p);
            }
            if let Some(h) = &format.horizontal {
                alignment = set_attr(&alignment, "horizontal", h);
            }
            if let Some(v) = &format.vertical {
                alignment = set_attr(&alignment, "vertical", v);
            }
            match format.wrap {
                Some(true) => alignment = set_attr(&alignment, "wrapText", "1"),
                Some(false) => alignment = remove_attr(&alignment, "wrapText"),
                None => {}
            }
            xf.set("alignment", Some(alignment));
            xf.start = with_attrs(&xf.start, &[("applyAlignment", "1".to_string())]);
        }
        let id = self.section("cellXfs").id_of(xf.write(XF_ORDER)) as u32;
        self.cache.insert((base, sides), id);
        Ok(id)
    }

    /// The stylesheet with the appended elements and updated counts
    fn write(self) -> Result<String, String> {
        let mut xml = self.xml;
        // From the end of the file backwards, so earlier offsets stay valid
        // (where, text to insert, section range, new count)
        let mut inserts: Vec<(usize, String, std::ops::Range<usize>, usize)> = Vec::new();
        for (section, _) in SECTIONS {
            let s = &self.sections[section];
            if s.added.is_empty() {
                continue;
            }
            let range = element_ranges(&xml, section, |_| true)?.into_iter().next().ok_or("Sección de estilos perdida.")?;
            let text = &xml[range.clone()];
            let close = text.rfind("</").map(|i| range.start + i);
            let total = s.items.len() + s.added.len();
            match close {
                Some(at) => inserts.push((at, s.added.concat(), range.clone(), total)),
                None => {
                    // Self-closing section (`<fonts count="0"/>`): later sections come after it,
                    // so the offsets kept above stay valid
                    let name = text[1..].split(|c: char| c.is_whitespace() || c == '/' || c == '>').next().unwrap_or(section).to_string();
                    let open = set_attr(text, "count", &total.to_string());
                    let open = open.trim_end_matches("/>").to_string();
                    xml.replace_range(range.clone(), &format!("{}>{}</{}>", open, s.added.concat(), name));
                }
            }
        }
        inserts.sort_by_key(|(at, _, _, _)| std::cmp::Reverse(*at));
        for (at, text, range, total) in inserts {
            xml.insert_str(at, &text);
            let start_tag_end = xml[range.start..].find('>').map(|i| range.start + i + 1).unwrap_or(range.start);
            let tag = xml[range.start..start_tag_end].to_string();
            let new_tag = set_attr(&tag, "count", &total.to_string());
            xml.replace_range(range.start..start_tag_end, &new_tag);
        }
        if !self.new_formats.is_empty() {
            let p = &self.prefix;
            let entries: String = self
                .new_formats
                .iter()
                .map(|(id, code)| format!(r#"<{}numFmt numFmtId="{}" formatCode="{}"/>"#, p, id, escape_xml(code)))
                .collect();
            match element_ranges(&xml, "numFmts", |_| true)?.into_iter().next() {
                Some(range) => {
                    let text = xml[range.clone()].to_string();
                    let count = self.num_fmts.len();
                    let updated = match text.rfind("</") {
                        Some(i) => format!("{}{}{}", &text[..i], entries, &text[i..]),
                        None => format!("{}>{}</{}numFmts>", text.trim_end_matches("/>"), entries, p),
                    };
                    let tag_end = updated.find('>').map(|i| i + 1).unwrap_or(0);
                    let tag = set_attr(&updated[..tag_end], "count", &count.to_string());
                    xml.replace_range(range, &format!("{}{}", tag, &updated[tag_end..]));
                }
                None => {
                    let (root_at, _) = root_element(&xml).ok_or("Hoja de estilos vacía.")?;
                    let after_root = xml[root_at..].find('>').map(|i| root_at + i + 1).ok_or("Hoja de estilos no válida.")?;
                    xml.insert_str(after_root, &format!(r#"<{p}numFmts count="{}">{}</{p}numFmts>"#, self.new_formats.len(), entries, p = p));
                }
            }
        }
        Ok(xml)
    }
}

/// A start tag (without `>`) with attributes set
fn with_attrs(start: &str, attrs: &[(&str, String)]) -> String {
    let mut tag = format!("{}/>", start);
    for (name, value) in attrs {
        tag = set_attr(&tag, name, value);
    }
    tag.truncate(tag.len() - 2);
    tag
}

/// Border sides of the cell `at` covered by `edges` over `range` (bits: left, right, top, bottom)
fn sides_for(edges: Edges, range: &CellRange, at: CellRef) -> u8 {
    let (top, bottom) = (at.row == range.first.row, at.row == range.last.row);
    let (left, right) = (at.col == range.first.col, at.col == range.last.col);
    let bit = |on: bool, b: u8| if on { b } else { 0 };
    match edges {
        Edges::All => 15,
        Edges::Outline => bit(left, 1) | bit(right, 2) | bit(top, 4) | bit(bottom, 8),
        Edges::Top => bit(top, 4),
        Edges::Bottom => bit(bottom, 8),
        Edges::Left => bit(left, 1),
        Edges::Right => bit(right, 2),
        Edges::Inside => bit(!left, 1) | bit(!right, 2) | bit(!top, 4) | bit(!bottom, 8),
    }
}

impl Workbook {
    /// The stylesheet part and its text (a default one when the workbook has none)
    fn stylesheet(&self) -> Result<(Option<String>, String), String> {
        match self.related_part("styles") {
            Some(part) => {
                let xml = self.package.get_string(&part)?;
                Ok((Some(part), xml))
            }
            None => Ok((None, crate::blank::STYLES_XML.to_string())),
        }
    }

    fn save_stylesheet(&mut self, part: Option<String>, xml: String) -> Result<(), String> {
        match part {
            Some(part) => {
                self.package.put(part, xml.into_bytes());
                Ok(())
            }
            None => self.add_workbook_part("styles", "styles.xml", STYLES_TYPE, xml),
        }
    }

    /// Applies `format` to every cell of `range` (cells without an element get one)
    pub fn set_format(&mut self, sheet: usize, range: &str, format: &Format) -> Result<(), String> {
        format.check()?;
        let range = CellRange::parse(range)?;
        if range.cell_count() > MAX_FORMAT_CELLS {
            return Err(format!("El rango {} tiene demasiadas celdas para darle formato de una vez.", range));
        }
        let part = self.worksheet_part(sheet)?;
        let xml = self.package.get_string(&part)?;
        let parsed = SheetXml::parse(&xml)?;
        let (styles_part, styles_xml) = self.stylesheet()?;
        let mut book = StyleBook::parse(&styles_xml)?;
        let prefix = parsed.prefix().to_string();
        let mut edits = std::collections::BTreeMap::new();
        for row in range.first.row..=range.last.row {
            for col in range.first.col..=range.last.col {
                let at = CellRef { row, col };
                let existing = parsed.cell(at);
                let base = existing.map(|c| c.style).unwrap_or(0);
                let sides = format.border.as_ref().map(|b| sides_for(b.edges, &range, at)).unwrap_or(0);
                let id = book.derive(base, format, sides)?;
                if existing.is_some_and(|c| c.style == id) {
                    continue;
                }
                let element = match existing {
                    Some(node) => set_attr(&xml[node.range.clone()], "s", &id.to_string()),
                    None => format!(r#"<{}c r="{}" s="{}"/>"#, prefix, at, id),
                };
                edits.insert(at, Some(element));
            }
        }
        if edits.is_empty() {
            return Ok(());
        }
        let new_sheet = parsed.splice(&edits);
        let new_styles = book.write()?;
        self.package.put(part, new_sheet.into_bytes());
        if new_styles != styles_xml || styles_part.is_none() {
            self.save_stylesheet(styles_part, new_styles)?;
        }
        Ok(())
    }

    /// Back to the default look (the workbook's first cell format) for the cells of `range`
    pub fn clear_format(&mut self, sheet: usize, range: &str) -> Result<(), String> {
        let range = CellRange::parse(range)?;
        let part = self.worksheet_part(sheet)?;
        let xml = self.package.get_string(&part)?;
        let parsed = SheetXml::parse(&xml)?;
        let mut edits = std::collections::BTreeMap::new();
        for cell in parsed.rows.iter().flat_map(|r| &r.cells) {
            let inside = cell.at.row >= range.first.row && cell.at.row <= range.last.row && cell.at.col >= range.first.col && cell.at.col <= range.last.col;
            if !inside || cell.style_attr.is_none() {
                continue;
            }
            let element = remove_attr(&xml[cell.range.clone()], "s");
            let empty = !element.contains("<v") && !element.contains(":v") && !element.contains("<f") && !element.contains("<is") && !element.contains(":is");
            edits.insert(cell.at, if empty && element.ends_with("/>") { None } else { Some(element) });
        }
        if edits.is_empty() {
            return Ok(());
        }
        let new_sheet = parsed.splice(&edits);
        self.package.put(part, new_sheet.into_bytes());
        Ok(())
    }

    /// Sets the width of columns `first..=last` (1 = A) in characters, as Excel measures it
    pub fn set_column_width(&mut self, sheet: usize, first: u32, last: u32, width: f64) -> Result<(), String> {
        if !(0.0..=255.0).contains(&width) {
            return Err("El ancho de columna debe estar entre 0 y 255 caracteres.".to_string());
        }
        let (first, last) = (first.min(last), first.max(last));
        if first == 0 || last > crate::cell_ref::MAX_COL {
            return Err("Columna fuera de la hoja.".to_string());
        }
        let part = self.worksheet_part(sheet)?;
        let xml = self.package.get_string(&part)?;
        let prefix = SheetXml::parse(&xml)?.prefix().to_string();
        let cols_range = element_ranges(&xml, "cols", |_| true)?.into_iter().next();
        let existing: Vec<String> = match &cols_range {
            Some(r) => {
                let inner = &xml[r.clone()];
                element_ranges(inner, "col", |_| true)?.into_iter().map(|c| inner[c].to_string()).collect()
            }
            None => Vec::new(),
        };
        let span = |el: &str| -> (u32, u32) {
            let mut reader = Reader::from_str(el);
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let n = |a: &str| get_attr_value(&e, a).and_then(|v| v.trim().parse().ok()).unwrap_or(0);
                    (n("min"), n("max"))
                }
                _ => (0, 0),
            }
        };
        let width_text = format!("{}", (width * 256.0).round() / 256.0);
        let sized = |el: &str, a: u32, b: u32| -> String {
            let el = set_attr(&set_attr(el, "min", &a.to_string()), "max", &b.to_string());
            set_attr(&set_attr(&el, "width", &width_text), "customWidth", "1")
        };
        let mut cols: Vec<(u32, String)> = Vec::new();
        let mut covered: Vec<(u32, u32)> = Vec::new();
        for el in &existing {
            let (a, b) = span(el);
            if b < first || a > last {
                cols.push((a, el.clone()));
                continue;
            }
            if a < first {
                cols.push((a, set_attr(&set_attr(el, "min", &a.to_string()), "max", &(first - 1).to_string())));
            }
            let (from, to) = (a.max(first), b.min(last));
            cols.push((from, sized(el, from, to)));
            covered.push((from, to));
            if b > last {
                cols.push((last + 1, set_attr(&set_attr(el, "min", &(last + 1).to_string()), "max", &b.to_string())));
            }
        }
        // Columns no <col> described yet
        covered.sort();
        let mut next = first;
        for (from, to) in covered.iter().copied().chain(std::iter::once((last + 1, last + 1))) {
            if from > next {
                cols.push((next, format!(r#"<{}col min="{}" max="{}" width="{}" customWidth="1"/>"#, prefix, next, from - 1, width_text)));
            }
            next = next.max(to + 1);
        }
        cols.sort_by_key(|(min, _)| *min);
        let body: String = cols.into_iter().map(|(_, el)| el).collect();
        let new_xml = match cols_range {
            Some(r) => {
                let text = &xml[r.clone()];
                let tag_end = text.find('>').map(|i| i + 1).unwrap_or(text.len());
                let open = text[..tag_end].trim_end_matches("/>").trim_end_matches('>');
                format!("{}{}>{}</{}cols>{}", &xml[..r.start], open, body, prefix, &xml[r.end..])
            }
            None => {
                let at = first_child_named(&xml, &["sheetData"])?.ok_or("La hoja no tiene <sheetData>.")?;
                format!("{}<{p}cols>{}</{p}cols>{}", &xml[..at], body, &xml[at..], p = prefix)
            }
        };
        self.package.put(part, new_xml.into_bytes());
        Ok(())
    }

    /// Merges the cells of `range` into one; only its first cell may hold a value
    pub fn merge_cells(&mut self, sheet: usize, range: &str) -> Result<(), String> {
        let range = CellRange::parse(range)?;
        if range.cell_count() < 2 {
            return Err("Para combinar hace falta un rango de más de una celda.".to_string());
        }
        let part = self.worksheet_part(sheet)?;
        let xml = self.package.get_string(&part)?;
        let layout = SheetLayout::parse(&xml)?;
        if layout.merges.contains(&range) {
            return Ok(());
        }
        if let Some(other) = layout.merges.iter().find(|m| crate::areas::overlaps_range(m, &range)) {
            return Err(format!("El rango {} se cruza con las celdas ya combinadas {}.", range, other));
        }
        let parsed = SheetXml::parse(&xml)?;
        let prefix = parsed.prefix().to_string();
        let filled: Vec<String> = parsed
            .rows
            .iter()
            .flat_map(|r| &r.cells)
            .filter(|c| c.at != range.first && crate::areas::range_contains(&range, c.at))
            .filter(|c| c.value.as_deref().is_some_and(|v| !v.is_empty()) || c.inline.is_some() || c.formula.is_some())
            .map(|c| c.at.to_string())
            .collect();
        if !filled.is_empty() {
            return Err(format!("Solo la primera celda de {} puede tener valor; vacía antes {}.", range, filled.join(", ")));
        }
        let entry = format!(r#"<{}mergeCell ref="{}"/>"#, prefix, range);
        let new_xml = match element_ranges(&xml, "mergeCells", |_| true)?.into_iter().next() {
            Some(r) => {
                let text = &xml[r.clone()];
                let count = layout.merges.len() + 1;
                let updated = match text.rfind("</") {
                    Some(i) => format!("{}{}{}", &text[..i], entry, &text[i..]),
                    None => format!("{}>{}</{}mergeCells>", text.trim_end_matches("/>"), entry, prefix),
                };
                let tag_end = updated.find('>').map(|i| i + 1).unwrap_or(0);
                let tag = set_attr(&updated[..tag_end], "count", &count.to_string());
                format!("{}{}{}{}", &xml[..r.start], tag, &updated[tag_end..], &xml[r.end..])
            }
            None => {
                // Schema order: mergeCells goes before these
                const AFTER: &[&str] = &[
                    "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup",
                    "headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
                    "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst",
                ];
                let at = match first_child_named(&xml, AFTER)? {
                    Some(at) => at,
                    None => xml.rfind("</").ok_or("La hoja no es XML válido.")?,
                };
                format!(r#"{}<{p}mergeCells count="1">{}</{p}mergeCells>{}"#, &xml[..at], entry, &xml[at..], p = prefix)
            }
        };
        self.package.put(part, new_xml.into_bytes());
        Ok(())
    }

    /// Splits the merged cells that overlap `range`; returns how many
    pub fn unmerge_cells(&mut self, sheet: usize, range: &str) -> Result<usize, String> {
        let range = CellRange::parse(range)?;
        let part = self.worksheet_part(sheet)?;
        let xml = self.package.get_string(&part)?;
        let Some(section) = element_ranges(&xml, "mergeCells", |_| true)?.into_iter().next() else { return Ok(0) };
        let text = xml[section.clone()].to_string();
        let mut removed = 0;
        let mut kept = 0;
        let mut out = String::new();
        let mut last = 0;
        for r in element_ranges(&text, "mergeCell", |_| true)? {
            let el = &text[r.clone()];
            let mut reader = Reader::from_str(el);
            let merge = match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => get_attr_value(&e, "ref").and_then(|v| CellRange::parse(&v).ok()),
                _ => None,
            };
            if merge.is_some_and(|m| crate::areas::overlaps_range(&m, &range)) {
                out.push_str(&text[last..r.start]);
                last = r.end;
                removed += 1;
            } else {
                kept += 1;
            }
        }
        if removed == 0 {
            return Ok(0);
        }
        out.push_str(&text[last..]);
        let replacement = if kept == 0 {
            String::new()
        } else {
            let tag_end = out.find('>').map(|i| i + 1).unwrap_or(0);
            format!("{}{}", set_attr(&out[..tag_end], "count", &kept.to_string()), &out[tag_end..])
        };
        let new_xml = format!("{}{}{}", &xml[..section.start], replacement, &xml[section.end..]);
        self.package.put(part, new_xml.into_bytes());
        Ok(removed)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::view::StyleSheetView;
    use crate::workbook::CellInput;

    fn blank() -> Workbook {
        Workbook::from_bytes(&crate::blank::blank_xlsx().unwrap()).unwrap()
    }

    fn sheet_xml(wb: &Workbook) -> String {
        wb.package.get_string("xl/worksheets/sheet1.xml").unwrap()
    }

    fn view(wb: &Workbook) -> StyleSheetView {
        wb.style_view().unwrap()
    }

    fn style_of(wb: &Workbook, r: &str) -> u32 {
        let xml = sheet_xml(wb);
        SheetXml::parse(&xml).unwrap().cell(CellRef::parse(r).unwrap()).map(|c| c.style).unwrap_or(0)
    }

    #[test]
    fn formats_fonts_fills_numbers_and_alignment_keeping_the_rest() {
        let mut wb = blank();
        wb.set_values(0, "A1", vec![vec![Some(CellInput::Text("Total".into())), Some(CellInput::Number(1234.5))]]).unwrap();
        let bold = Format { bold: Some(true), fill: Some(Some("#FFFF00".into())), horizontal: Some("center".into()), ..Default::default() };
        wb.set_format(0, "A1:C1", &bold).unwrap();
        let money = Format { number_format: Some("#,##0.00 \"€\"".into()), font_color: Some("c00000".into()), ..Default::default() };
        wb.set_format(0, "B1", &money).unwrap();

        let styles = view(&wb);
        let (a1, b1, c1) = (style_of(&wb, "A1"), style_of(&wb, "B1"), style_of(&wb, "C1"));
        assert_eq!(a1, c1, "same format reused");
        let xf_a = &styles.xfs[a1 as usize];
        assert!(styles.fonts[xf_a.font].bold);
        assert_eq!(styles.fills[xf_a.fill].as_deref(), Some("#FFFF00"));
        assert_eq!(xf_a.horizontal.as_deref(), Some("center"));
        // B1 keeps bold, fill and alignment and adds the color and the number format
        let xf_b = &styles.xfs[b1 as usize];
        assert!(styles.fonts[xf_b.font].bold);
        assert_eq!(styles.fonts[xf_b.font].color.as_deref(), Some("#C00000"));
        assert_eq!(styles.fills[xf_b.fill].as_deref(), Some("#FFFF00"));
        assert_eq!(styles.custom_formats.get(&xf_b.num_fmt).map(String::as_str), Some("#,##0.00 \"€\""));
        // The cell C1 did not exist: it was created with only the format
        assert!(sheet_xml(&wb).contains(&format!(r#"<c r="C1" s="{}"/>"#, c1)));
        assert_eq!(wb.cell(0, "B1").unwrap().value, crate::workbook::CellValue::Number(1234.5));

        wb.clear_format(0, "A1:C1").unwrap();
        assert_eq!((style_of(&wb, "A1"), style_of(&wb, "B1")), (0, 0));
        assert!(!sheet_xml(&wb).contains(r#"r="C1""#));
        assert!(wb.set_format(0, "A1", &Format { font_color: Some("rojo".into()), ..Default::default() }).is_err());
    }

    #[test]
    fn draws_outline_and_inside_borders() {
        let mut wb = blank();
        let outline = Format { border: Some(BorderSpec { style: Some("medium".into()), color: None, edges: Edges::Outline }), ..Default::default() };
        wb.set_format(0, "B2:C3", &outline).unwrap();
        let styles = view(&wb);
        let sides = |r: &str| -> Vec<bool> {
            let border = &styles.borders[styles.xfs[style_of(&wb, r) as usize].border];
            border.sides.iter().map(|s| s.is_some()).collect()
        };
        // left, right, top, bottom
        assert_eq!(sides("B2"), [true, false, true, false]);
        assert_eq!(sides("C3"), [false, true, false, true]);
        let all = Format { border: Some(BorderSpec { style: Some("thin".into()), color: Some("#000000".into()), edges: Edges::All }), ..Default::default() };
        wb.set_format(0, "E5", &all).unwrap();
        let styles = view(&wb);
        assert!(styles.borders[styles.xfs[style_of(&wb, "E5") as usize].border].sides.iter().all(|s| s.as_ref().is_some_and(|(st, _)| st == "thin")));
    }

    #[test]
    fn sets_column_widths_splitting_existing_spans() {
        let mut wb = blank();
        wb.set_column_width(0, 2, 4, 20.0).unwrap();
        assert!(sheet_xml(&wb).contains(r#"<cols><col min="2" max="4" width="20" customWidth="1"/></cols><sheetData"#), "{}", sheet_xml(&wb));
        wb.set_column_width(0, 3, 6, 8.5).unwrap();
        let layout = SheetLayout::parse(&sheet_xml(&wb)).unwrap();
        assert_eq!(layout.cols, vec![(2, 2, Some(20.0), false), (3, 4, Some(8.5), false), (5, 6, Some(8.5), false)]);
    }

    #[test]
    fn merges_and_unmerges_cells() {
        let mut wb = blank();
        wb.set_value(0, "A1", CellInput::Text("Título".into())).unwrap();
        wb.merge_cells(0, "A1:D1").unwrap();
        assert!(sheet_xml(&wb).contains(r#"</sheetData><mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells>"#), "{}", sheet_xml(&wb));
        assert!(wb.merge_cells(0, "C1:C2").unwrap_err().contains("A1:D1"));
        wb.set_value(0, "B3", CellInput::Number(1.0)).unwrap();
        assert!(wb.merge_cells(0, "A3:B3").unwrap_err().contains("B3"));
        wb.merge_cells(0, "A5:B6").unwrap();
        assert_eq!(SheetLayout::parse(&sheet_xml(&wb)).unwrap().merges.len(), 2);
        assert_eq!(wb.unmerge_cells(0, "B1").unwrap(), 1);
        assert_eq!(wb.unmerge_cells(0, "A5").unwrap(), 1);
        assert!(!sheet_xml(&wb).contains("mergeCell"));
    }
}
