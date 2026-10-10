//! Style cascade (`word/styles.xml` + theme fonts) and list numbering (`word/numbering.xml`).
//!
//! Effective formatting follows the OOXML order:
//! document defaults → numbering level (paragraph props) → paragraph style chain →
//! character style chain → direct formatting.

use crate::docx_parser::{
    get_attr_value, is_bool_element_true, parse_border_element, tag_is, ParagraphBorders, TableBorders,
};
use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;
use std::collections::HashMap;

/// Run properties as written in the XML: `None` means "not specified here, inherit"
#[derive(Debug, Clone, Default, PartialEq)]
pub struct RunProps {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    /// Hex color without '#'; `Some("")` is an explicit `auto`
    pub color: Option<String>,
    /// Points
    pub font_size: Option<f64>,
    pub font_family: Option<String>,
}

impl RunProps {
    pub fn merge(&mut self, over: &RunProps) {
        if over.bold.is_some() {
            self.bold = over.bold;
        }
        if over.italic.is_some() {
            self.italic = over.italic;
        }
        if over.underline.is_some() {
            self.underline = over.underline;
        }
        if over.color.is_some() {
            self.color = over.color.clone();
        }
        if over.font_size.is_some() {
            self.font_size = over.font_size;
        }
        if over.font_family.is_some() {
            self.font_family = over.font_family.clone();
        }
    }

    pub fn merged(&self, over: &RunProps) -> RunProps {
        let mut out = self.clone();
        out.merge(over);
        out
    }
}

/// Paragraph properties as written in the XML (lengths already converted like the parser:
/// indents in px, spacing in pt, line spacing as a multiplier)
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ParaProps {
    pub align: Option<String>,
    pub indent_left: Option<f64>,
    pub indent_right: Option<f64>,
    /// Positive = first line indent, negative = hanging indent
    pub indent_first_line: Option<f64>,
    pub space_before: Option<f64>,
    pub space_after: Option<f64>,
    /// "auto", "exact" or "atLeast"; set together with `line_spacing` / `line_pt`
    pub line_rule: Option<String>,
    /// Line spacing multiplier (lineRule auto)
    pub line_spacing: Option<f64>,
    /// Fixed or minimum line height in points (lineRule exact / atLeast)
    pub line_pt: Option<f64>,
    pub keep_next: Option<bool>,
    pub keep_lines: Option<bool>,
    pub page_break_before: Option<bool>,
    pub widow_control: Option<bool>,
    pub contextual_spacing: Option<bool>,
    /// "0" explicitly removes numbering
    pub num_id: Option<String>,
    pub ilvl: Option<u32>,
    pub outline_level: Option<u32>,
    pub borders: ParagraphBorders,
}

impl ParaProps {
    pub fn merge(&mut self, over: &ParaProps) {
        macro_rules! take {
            ($($f:ident),*) => { $( if over.$f.is_some() { self.$f = over.$f.clone(); } )* };
        }
        take!(align, indent_left, indent_right, indent_first_line, space_before, space_after,
              num_id, ilvl, outline_level, keep_next, keep_lines, page_break_before, widow_control,
              contextual_spacing);
        if over.line_rule.is_some() {
            self.line_rule = over.line_rule.clone();
            self.line_spacing = over.line_spacing;
            self.line_pt = over.line_pt;
        }
        let b = &over.borders;
        take_border(&mut self.borders.top, &b.top);
        take_border(&mut self.borders.bottom, &b.bottom);
        take_border(&mut self.borders.left, &b.left);
        take_border(&mut self.borders.right, &b.right);
    }
}

fn take_border<T: Clone>(dst: &mut Option<T>, src: &Option<T>) {
    if src.is_some() {
        *dst = src.clone();
    }
}

#[derive(Debug, Clone, Default)]
pub struct ThemeFonts {
    pub major: Option<String>,
    pub minor: Option<String>,
    /// Color scheme of the theme ("dk1", "lt1", "accent1", …) as "RRGGBB"
    pub colors: HashMap<String, String>,
}

impl ThemeFonts {
    /// A scheme color by the name shapes use ("tx1" is "dk1", "bg1" is "lt1", …)
    pub fn color(&self, name: &str) -> Option<String> {
        let key = match name {
            "tx1" => "dk1",
            "bg1" => "lt1",
            "tx2" => "dk2",
            "bg2" => "lt2",
            other => other,
        };
        self.colors.get(key).cloned().or_else(|| match key {
            "dk1" => Some("000000".into()),
            "lt1" => Some("FFFFFF".into()),
            "dk2" => Some("44546A".into()),
            "lt2" => Some("E7E6E6".into()),
            "accent1" => Some("4472C4".into()),
            "accent2" => Some("ED7D31".into()),
            "accent3" => Some("A5A5A5".into()),
            "accent4" => Some("FFC000".into()),
            "accent5" => Some("5B9BD5".into()),
            "accent6" => Some("70AD47".into()),
            _ => None,
        })
    }
}

#[derive(Debug, Clone, Default)]
pub struct Style {
    pub id: String,
    pub name: String,
    /// "paragraph", "character", "table" or "numbering"
    pub kind: String,
    pub based_on: Option<String>,
    /// Style for the paragraph created by pressing Enter at the end of this one
    pub next: Option<String>,
    pub ppr: ParaProps,
    pub rpr: RunProps,
    /// `w:tblPr/w:tblBorders` of a table style
    pub table_borders: TableBorders,
}

#[derive(Debug, Clone, Default)]
pub struct Level {
    pub start: i64,
    pub fmt: String,
    pub text: String,
    /// "tab" (default), "space" or "nothing"
    pub suffix: String,
    pub ppr: ParaProps,
    pub rpr: RunProps,
}

#[derive(Debug, Clone, Default)]
pub struct NumDef {
    pub abstract_id: String,
    pub start_overrides: HashMap<u32, i64>,
}

#[derive(Debug, Clone, Default)]
pub struct Numbering {
    pub abstracts: HashMap<String, HashMap<u32, Level>>,
    pub nums: HashMap<String, NumDef>,
}

impl Numbering {
    pub fn level(&self, num_id: &str, ilvl: u32) -> Option<&Level> {
        let num = self.nums.get(num_id)?;
        self.abstracts.get(&num.abstract_id)?.get(&ilvl)
    }
}

#[derive(Debug, Clone, Default)]
pub struct StyleSheet {
    /// False when the package has no styles.xml (layout then falls back to heuristics)
    pub loaded: bool,
    pub doc_ppr: ParaProps,
    pub doc_rpr: RunProps,
    pub styles: HashMap<String, Style>,
    pub default_paragraph_style: Option<String>,
    pub default_character_style: Option<String>,
    pub default_table_style: Option<String>,
    pub theme: ThemeFonts,
    pub numbering: Numbering,
    /// Style chains flattened at load time: id → (pPr, rPr) including all `basedOn` ancestors
    resolved: HashMap<String, (ParaProps, RunProps)>,
}

impl StyleSheet {
    pub fn load(styles_xml: Option<&str>, theme_xml: Option<&str>, numbering_xml: Option<&str>) -> Self {
        let mut sheet = StyleSheet {
            theme: theme_xml.map(parse_theme_fonts).unwrap_or_default(),
            ..Default::default()
        };
        if let Some(xml) = styles_xml {
            sheet.parse_styles(xml);
            sheet.loaded = true;
        }
        if let Some(xml) = numbering_xml {
            sheet.numbering = parse_numbering(xml, &sheet.theme);
        }
        sheet.flatten();
        sheet
    }

    /// Paragraph style id actually in effect (unknown or missing ids fall back to the default style)
    pub fn effective_paragraph_style<'a>(&'a self, id: Option<&'a str>) -> Option<&'a str> {
        match id {
            Some(id) if self.styles.contains_key(id) => Some(id),
            _ => self.default_paragraph_style.as_deref(),
        }
    }

    /// Style for a new paragraph after one in style `id` (Word's "style for following paragraph")
    pub fn next_style(&self, id: &str) -> Option<&str> {
        self.styles.get(id)?.next.as_deref()
    }

    pub fn style_name(&self, id: &str) -> Option<&str> {
        self.styles.get(id).map(|s| s.name.as_str())
    }

    /// Flattened (pPr, rPr) of a style chain, without document defaults
    pub fn style_props(&self, id: Option<&str>) -> (ParaProps, RunProps) {
        id.and_then(|id| self.resolved.get(id)).cloned().unwrap_or_default()
    }

    /// Borders of a table style chain (falls back to the default table style)
    pub fn table_borders(&self, id: Option<&str>) -> TableBorders {
        let id = match id {
            Some(id) if self.styles.contains_key(id) => Some(id.to_string()),
            _ => self.default_table_style.clone(),
        };
        let mut chain = Vec::new();
        let mut cur = id;
        while let Some(c) = cur {
            if chain.contains(&c) || chain.len() > 32 {
                break;
            }
            cur = self.styles.get(&c).and_then(|s| s.based_on.clone());
            chain.push(c);
        }
        let mut borders = TableBorders::default();
        for c in chain.iter().rev() {
            if let Some(style) = self.styles.get(c) {
                borders.merge(&style.table_borders);
            }
        }
        borders
    }

    /// Run properties of a character style chain (falls back to the default character style)
    pub fn character_props(&self, id: Option<&str>) -> RunProps {
        let id = match id {
            Some(id) if self.styles.contains_key(id) => Some(id),
            _ => self.default_character_style.as_deref(),
        };
        self.style_props(id).1
    }

    fn flatten(&mut self) {
        let ids: Vec<String> = self.styles.keys().cloned().collect();
        for id in ids {
            let mut chain = Vec::new();
            let mut cur = Some(id.clone());
            while let Some(c) = cur {
                if chain.contains(&c) || chain.len() > 32 {
                    break;
                }
                cur = self.styles.get(&c).and_then(|s| s.based_on.clone());
                chain.push(c);
            }
            let mut ppr = ParaProps::default();
            let mut rpr = RunProps::default();
            for c in chain.iter().rev() {
                if let Some(s) = self.styles.get(c) {
                    ppr.merge(&s.ppr);
                    rpr.merge(&s.rpr);
                }
            }
            self.resolved.insert(id, (ppr, rpr));
        }
    }

    fn parse_styles(&mut self, xml: &str) {
        let mut reader = Reader::from_str(xml);
        let mut path: Vec<String> = Vec::new();
        let mut current: Option<Style> = None;
        let mut buf = Vec::new();
        loop {
            let (e, is_start) = match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => (e.into_owned(), true),
                Ok(Event::Empty(e)) => (e.into_owned(), false),
                Ok(Event::End(_)) => {
                    if path.pop().as_deref() == Some("style") {
                        if let Some(style) = current.take() {
                            self.add_style(style);
                        }
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
            let name = local_name(&e);
            if name == "style" {
                current = Some(Style {
                    id: get_attr_value(&e, "styleId").unwrap_or_default(),
                    kind: get_attr_value(&e, "type").unwrap_or_else(|| "paragraph".to_string()),
                    ..Default::default()
                });
                let is_default = get_attr_value(&e, "default").is_some_and(|v| v == "1" || v == "true");
                if let (true, Some(s)) = (is_default, current.as_ref()) {
                    match s.kind.as_str() {
                        "paragraph" => self.default_paragraph_style = Some(s.id.clone()),
                        "character" => self.default_character_style = Some(s.id.clone()),
                        "table" => self.default_table_style = Some(s.id.clone()),
                        _ => {}
                    }
                }
            } else if let Some(style) = current.as_mut() {
                if path.last().map(String::as_str) == Some("style") {
                    match name.as_str() {
                        "name" => style.name = get_attr_value(&e, "val").unwrap_or_default(),
                        "basedOn" => style.based_on = get_attr_value(&e, "val"),
                        "next" => style.next = get_attr_value(&e, "val"),
                        _ => {}
                    }
                } else if path.iter().any(|p| p == "tblStylePr") {
                    // Conditional formatting (header row, banding) is not applied yet
                } else if path.ends_with(&["tblPr".to_string(), "tblBorders".to_string()]) {
                    style.table_borders.set_side(&e);
                } else {
                    apply_property(&path, &e, &mut style.ppr, &mut style.rpr, &self.theme);
                }
            } else if path.iter().any(|p| p == "docDefaults") {
                apply_property(&path, &e, &mut self.doc_ppr, &mut self.doc_rpr, &self.theme);
            }
            if is_start {
                path.push(name);
            }
            buf.clear();
        }
    }

    fn add_style(&mut self, style: Style) {
        if !style.id.is_empty() {
            self.styles.insert(style.id.clone(), style);
        }
    }
}

fn local_name(e: &BytesStart) -> String {
    let name = e.name();
    let s = crate::docx_parser::utf8(name.as_ref());
    s.rsplit(':').next().unwrap_or(s).to_string()
}

/// Applies an element found inside a property block; `path` holds the enclosing local names
fn apply_property(path: &[String], e: &BytesStart, ppr: &mut ParaProps, rpr: &mut RunProps, theme: &ThemeFonts) {
    if path.iter().any(|p| p == "pPrChange" || p == "rPrChange") {
        return;
    }
    let parent = path.last().map(String::as_str);
    let grandparent = path.len().checked_sub(2).map(|i| path[i].as_str());
    match (parent, grandparent) {
        (Some("rPr"), gp) if gp != Some("pPr") => apply_rpr_element(rpr, e, theme),
        (Some("pPr"), _) | (Some("numPr"), Some("pPr")) => apply_ppr_element(ppr, e),
        (Some("pBdr"), Some("pPr")) => apply_border_element(&mut ppr.borders, e),
        _ => {}
    }
}

pub fn apply_border_element(borders: &mut ParagraphBorders, e: &BytesStart) {
    let name = e.name();
    let n = name.as_ref();
    let side = if tag_is(n, "top") {
        &mut borders.top
    } else if tag_is(n, "bottom") {
        &mut borders.bottom
    } else if tag_is(n, "left") || tag_is(n, "start") {
        &mut borders.left
    } else if tag_is(n, "right") || tag_is(n, "end") {
        &mut borders.right
    } else {
        return;
    };
    *side = parse_border_element(e);
}

/// Applies one child of a `w:rPr` block
pub fn apply_rpr_element(rpr: &mut RunProps, e: &BytesStart, theme: &ThemeFonts) {
    let name = e.name();
    let n = name.as_ref();
    if tag_is(n, "b") || tag_is(n, "bCs") {
        rpr.bold = Some(is_bool_element_true(e));
    } else if tag_is(n, "i") || tag_is(n, "iCs") {
        rpr.italic = Some(is_bool_element_true(e));
    } else if tag_is(n, "u") {
        rpr.underline = Some(is_bool_element_true(e));
    } else if tag_is(n, "color") {
        if let Some(val) = get_attr_value(e, "val") {
            rpr.color = Some(if val.eq_ignore_ascii_case("auto") { String::new() } else { val.to_uppercase() });
        }
    } else if tag_is(n, "sz") || tag_is(n, "szCs") {
        if let Some(v) = get_attr_value(e, "val").and_then(|v| v.parse::<f64>().ok()) {
            rpr.font_size = Some(v / 2.0);
        }
    } else if tag_is(n, "rFonts") {
        let themed = get_attr_value(e, "asciiTheme")
            .or_else(|| get_attr_value(e, "hAnsiTheme"))
            .or_else(|| get_attr_value(e, "cstheme"))
            .and_then(|t| if t.starts_with("major") { theme.major.clone() } else { theme.minor.clone() });
        if let Some(fam) = themed
            .or_else(|| get_attr_value(e, "ascii"))
            .or_else(|| get_attr_value(e, "hAnsi"))
            .or_else(|| get_attr_value(e, "cs"))
        {
            rpr.font_family = Some(fam.trim().to_string());
        }
    }
}

/// Applies one child of a `w:pPr` block (pStyle and pBdr are handled by the caller)
pub fn apply_ppr_element(ppr: &mut ParaProps, e: &BytesStart) {
    let name = e.name();
    let n = name.as_ref();
    let num = |attr: &str| get_attr_value(e, attr).and_then(|v| v.parse::<f64>().ok());
    if tag_is(n, "jc") {
        if let Some(val) = get_attr_value(e, "val") {
            ppr.align = Some(
                match val.as_str() {
                    "start" => "left",
                    "end" => "right",
                    "distribute" | "justify" => "both",
                    other => other,
                }
                .to_string(),
            );
        }
    } else if tag_is(n, "ind") {
        if let Some(v) = num("left").or_else(|| num("start")) {
            ppr.indent_left = Some(v / 15.0);
        }
        if let Some(v) = num("right").or_else(|| num("end")) {
            ppr.indent_right = Some(v / 15.0);
        }
        if let Some(v) = num("hanging") {
            ppr.indent_first_line = Some(-v / 15.0);
        } else if let Some(v) = num("firstLine") {
            ppr.indent_first_line = Some(v / 15.0);
        }
    } else if tag_is(n, "spacing") {
        if let Some(v) = num("before") {
            ppr.space_before = Some(v / 20.0);
        }
        if let Some(v) = num("after") {
            ppr.space_after = Some(v / 20.0);
        }
        if let Some(v) = num("line") {
            let rule = get_attr_value(e, "lineRule").unwrap_or_else(|| "auto".to_string());
            if rule == "exact" || rule == "atLeast" {
                ppr.line_spacing = None;
                ppr.line_pt = Some(v / 20.0);
            } else {
                ppr.line_spacing = Some(v / 240.0);
                ppr.line_pt = None;
            }
            ppr.line_rule = Some(rule);
        }
    } else if tag_is(n, "keepNext") {
        ppr.keep_next = Some(is_bool_element_true(e));
    } else if tag_is(n, "keepLines") {
        ppr.keep_lines = Some(is_bool_element_true(e));
    } else if tag_is(n, "pageBreakBefore") {
        ppr.page_break_before = Some(is_bool_element_true(e));
    } else if tag_is(n, "widowControl") {
        ppr.widow_control = Some(is_bool_element_true(e));
    } else if tag_is(n, "contextualSpacing") {
        ppr.contextual_spacing = Some(is_bool_element_true(e));
    } else if tag_is(n, "numId") {
        ppr.num_id = get_attr_value(e, "val");
    } else if tag_is(n, "ilvl") {
        ppr.ilvl = get_attr_value(e, "val").and_then(|v| v.parse().ok());
    } else if tag_is(n, "outlineLvl") {
        ppr.outline_level = get_attr_value(e, "val").and_then(|v| v.parse().ok());
    }
}

fn parse_theme_fonts(xml: &str) -> ThemeFonts {
    let mut reader = Reader::from_str(xml);
    let mut fonts = ThemeFonts::default();
    let mut in_major = false;
    let mut in_minor = false;
    let mut in_colors = false;
    let mut scheme_slot: Option<String> = None;
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                let n = e.name();
                let local = crate::docx_parser::utf8(n.as_ref()).rsplit(':').next().unwrap_or("").to_string();
                if local == "clrScheme" {
                    in_colors = true;
                } else if in_colors && scheme_slot.is_none() && local != "srgbClr" && local != "sysClr" {
                    scheme_slot = Some(local);
                } else if in_colors && (local == "srgbClr" || local == "sysClr") {
                    let value = get_attr_value(&e, "lastClr").or_else(|| get_attr_value(&e, "val"));
                    if let (Some(slot), Some(v)) = (scheme_slot.take(), value) {
                        fonts.colors.insert(slot, v.to_ascii_uppercase());
                    }
                } else if tag_is(n.as_ref(), "majorFont") {
                    in_major = true;
                } else if tag_is(n.as_ref(), "minorFont") {
                    in_minor = true;
                } else if tag_is(n.as_ref(), "latin") {
                    let face = get_attr_value(&e, "typeface").filter(|f| !f.is_empty());
                    if in_major && fonts.major.is_none() {
                        fonts.major = face;
                    } else if in_minor && fonts.minor.is_none() {
                        fonts.minor = face;
                    }
                }
            }
            Ok(Event::End(e)) => {
                if tag_is(e.name().as_ref(), "clrScheme") {
                    in_colors = false;
                } else if in_colors {
                    scheme_slot = None;
                }
                if tag_is(e.name().as_ref(), "majorFont") {
                    in_major = false;
                } else if tag_is(e.name().as_ref(), "minorFont") {
                    in_minor = false;
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
    fonts
}

fn parse_numbering(xml: &str, theme: &ThemeFonts) -> Numbering {
    let mut reader = Reader::from_str(xml);
    let mut numbering = Numbering::default();
    let mut path: Vec<String> = Vec::new();
    let mut abstract_id: Option<String> = None;
    let mut level: Option<(u32, Level)> = None;
    let mut num: Option<(String, NumDef)> = None;
    let mut override_lvl: Option<u32> = None;
    let mut buf = Vec::new();
    loop {
        let (e, is_start) = match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => (e.into_owned(), true),
            Ok(Event::Empty(e)) => (e.into_owned(), false),
            Ok(Event::End(_)) => {
                match path.pop().as_deref() {
                    Some("lvl") if path.last().map(String::as_str) == Some("abstractNum") => {
                        if let (Some(id), Some((ilvl, lvl))) = (abstract_id.as_ref(), level.take()) {
                            numbering.abstracts.entry(id.clone()).or_default().insert(ilvl, lvl);
                        }
                    }
                    Some("abstractNum") => abstract_id = None,
                    Some("num") => {
                        if let Some((id, def)) = num.take() {
                            numbering.nums.insert(id, def);
                        }
                    }
                    Some("lvlOverride") => override_lvl = None,
                    _ => {}
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
        let name = local_name(&e);
        let parent = path.last().map(String::as_str);
        let val = || get_attr_value(&e, "val");
        match (name.as_str(), parent) {
            ("abstractNum", _) => {
                abstract_id = get_attr_value(&e, "abstractNumId");
                numbering.abstracts.entry(abstract_id.clone().unwrap_or_default()).or_default();
            }
            ("lvl", Some("abstractNum")) => {
                let ilvl = get_attr_value(&e, "ilvl").and_then(|v| v.parse().ok()).unwrap_or(0);
                level = Some((ilvl, Level { start: 1, fmt: "decimal".into(), suffix: "tab".into(), ..Default::default() }));
            }
            ("start", Some("lvl")) => {
                if let (Some((_, l)), Some(v)) = (level.as_mut(), val().and_then(|v| v.parse().ok())) {
                    l.start = v;
                }
            }
            ("numFmt", Some("lvl")) => {
                if let (Some((_, l)), Some(v)) = (level.as_mut(), val()) {
                    l.fmt = v;
                }
            }
            ("lvlText", Some("lvl")) => {
                if let (Some((_, l)), Some(v)) = (level.as_mut(), val()) {
                    l.text = v;
                }
            }
            ("suff", Some("lvl")) => {
                if let (Some((_, l)), Some(v)) = (level.as_mut(), val()) {
                    l.suffix = v;
                }
            }
            ("num", _) => {
                num = get_attr_value(&e, "numId").map(|id| (id, NumDef::default()));
            }
            ("abstractNumId", Some("num")) => {
                if let (Some((_, def)), Some(v)) = (num.as_mut(), val()) {
                    def.abstract_id = v;
                }
            }
            ("lvlOverride", _) => {
                override_lvl = get_attr_value(&e, "ilvl").and_then(|v| v.parse().ok());
            }
            ("startOverride", Some("lvlOverride")) => {
                if let (Some((_, def)), Some(ilvl), Some(v)) =
                    (num.as_mut(), override_lvl, val().and_then(|v| v.parse().ok()))
                {
                    def.start_overrides.insert(ilvl, v);
                }
            }
            _ => {
                if let Some((_, l)) = level.as_mut() {
                    if path.iter().any(|p| p == "lvl") {
                        apply_property(&path, &e, &mut l.ppr, &mut l.rpr, theme);
                    }
                }
            }
        }
        if is_start {
            path.push(name);
        }
        buf.clear();
    }
    numbering
}

/// Running list counters while walking the document in order
#[derive(Debug, Default)]
pub struct NumberingCounters {
    /// Keyed by abstractNum: Word continues numbering across `w:num`s sharing one
    lists: HashMap<String, [Option<i64>; 9]>,
    seen_nums: Vec<String>,
}

impl NumberingCounters {
    /// Advances the counter of (num_id, ilvl) and returns the rendered label
    pub fn next_label(&mut self, numbering: &Numbering, num_id: &str, ilvl: u32) -> Option<(String, Level)> {
        let ilvl = ilvl.min(8);
        let num = numbering.nums.get(num_id)?;
        let levels = numbering.abstracts.get(&num.abstract_id)?;
        let level = levels.get(&ilvl)?.clone();
        let values = self.lists.entry(num.abstract_id.clone()).or_insert([None; 9]);

        if !self.seen_nums.iter().any(|n| n == num_id) {
            self.seen_nums.push(num_id.to_string());
            for (&l, &start) in &num.start_overrides {
                if (l as usize) < 9 {
                    values[l as usize] = Some(start - 1);
                    for v in values.iter_mut().skip(l as usize + 1) {
                        *v = None;
                    }
                }
            }
        }

        let i = ilvl as usize;
        values[i] = Some(values[i].map_or(level.start, |v| v + 1));
        for v in values.iter_mut().skip(i + 1) {
            *v = None;
        }

        if level.fmt == "none" {
            return Some((String::new(), level));
        }
        if level.fmt == "bullet" {
            let label = map_bullet(&level.text, level.rpr.font_family.as_deref());
            return Some((label, level));
        }

        let mut label = level.text.clone();
        for n in (1..=9).rev() {
            let placeholder = format!("%{}", n);
            if label.contains(&placeholder) {
                let lvl = levels.get(&(n - 1));
                let value = values[(n - 1) as usize].unwrap_or_else(|| lvl.map_or(1, |l| l.start));
                let fmt = lvl.map_or("decimal", |l| l.fmt.as_str());
                label = label.replace(&placeholder, &format_number(value, fmt));
            }
        }
        Some((label, level))
    }
}

pub fn format_number(n: i64, fmt: &str) -> String {
    match fmt {
        "lowerLetter" => to_letters(n).to_lowercase(),
        "upperLetter" => to_letters(n),
        "lowerRoman" => to_roman(n).to_lowercase(),
        "upperRoman" => to_roman(n),
        "decimalZero" => format!("{:02}", n),
        "ordinal" => format!("{}º", n),
        _ => n.to_string(),
    }
}

/// Word-style letters: 1 → A, 26 → Z, 27 → AA, 28 → BB
fn to_letters(n: i64) -> String {
    if n <= 0 {
        return n.to_string();
    }
    let letter = (b'A' + ((n - 1) % 26) as u8) as char;
    std::iter::repeat(letter).take(((n - 1) / 26 + 1) as usize).collect()
}

fn to_roman(mut n: i64) -> String {
    if n <= 0 || n >= 4000 {
        return n.to_string();
    }
    const TABLE: [(i64, &str); 13] = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
        (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    ];
    let mut out = String::new();
    for (value, numeral) in TABLE {
        while n >= value {
            out.push_str(numeral);
            n -= value;
        }
    }
    out
}

/// Symbol/Wingdings bullets use private-use code points; map them to Unicode glyphs
pub fn map_bullet(text: &str, font: Option<&str>) -> String {
    let font = font.unwrap_or("").to_lowercase();
    text.chars()
        .map(|c| {
            let code = c as u32;
            let low = if (0xF000..=0xF0FF).contains(&code) { code - 0xF000 } else { code };
            if font.contains("wingdings") {
                match low {
                    0xA7 | 0x6E => '▪',
                    0xD8 => '➢',
                    0xFC => '✓',
                    0x76 => '❖',
                    0x71 => '❑',
                    0xA8 => '◻',
                    _ => '•',
                }
            } else if font.contains("symbol") || (0xF000..=0xF0FF).contains(&code) || code == 0xB7 {
                '•'
            } else if c == 'o' && font.contains("courier") {
                '◦'
            } else {
                c
            }
        })
        .collect()
}

/// True when a bullet label was remapped from a symbol font (so that font must not be used to draw it)
pub fn is_symbol_font(font: Option<&str>) -> bool {
    font.is_some_and(|f| {
        let f = f.to_lowercase();
        f.contains("symbol") || f.contains("wingdings")
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    const STYLES: &str = r#"<w:styles xmlns:w="w">
        <w:docDefaults>
            <w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
            <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
        </w:docDefaults>
        <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
        <w:style w:type="paragraph" w:styleId="Ttulo1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
            <w:pPr><w:keepNext/><w:spacing w:before="240"/><w:outlineLvl w:val="0"/></w:pPr>
            <w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:b/><w:color w:val="2f5496"/><w:sz w:val="32"/></w:rPr>
        </w:style>
        <w:style w:type="paragraph" w:styleId="Ttulo2"><w:name w:val="heading 2"/><w:basedOn w:val="Ttulo1"/>
            <w:rPr><w:sz w:val="26"/></w:rPr>
        </w:style>
        <w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>
    </w:styles>"#;

    const THEME: &str = r#"<a:theme><a:themeElements><a:fontScheme>
        <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
        <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
    </a:fontScheme></a:themeElements></a:theme>"#;

    #[test]
    fn test_style_chain_and_theme_fonts() {
        let sheet = StyleSheet::load(Some(STYLES), Some(THEME), None);
        assert_eq!(sheet.default_paragraph_style.as_deref(), Some("Normal"));
        assert_eq!(sheet.doc_rpr.font_family.as_deref(), Some("Calibri"));
        assert_eq!(sheet.doc_rpr.font_size, Some(11.0));
        assert_eq!(sheet.doc_ppr.space_after, Some(8.0));

        let (ppr, rpr) = sheet.style_props(Some("Ttulo2"));
        assert_eq!(rpr.font_size, Some(13.0), "own size overrides the base style");
        assert_eq!(rpr.bold, Some(true), "bold inherited from Ttulo1");
        assert_eq!(rpr.color.as_deref(), Some("2F5496"));
        assert_eq!(rpr.font_family.as_deref(), Some("Calibri Light"));
        assert_eq!(ppr.outline_level, Some(0));
        assert_eq!(sheet.character_props(Some("Strong")).bold, Some(true));
    }

    #[test]
    fn test_style_numbering_reference() {
        let xml = r#"<w:styles><w:style w:type="paragraph" w:styleId="723"><w:name w:val="List Bullet"/>
            <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="7"/></w:numPr></w:pPr></w:style></w:styles>"#;
        let sheet = StyleSheet::load(Some(xml), None, None);
        let (ppr, _) = sheet.style_props(Some("723"));
        assert_eq!((ppr.num_id.as_deref(), ppr.ilvl), (Some("7"), Some(1)));
    }

    const NUMBERING: &str = r#"<w:numbering>
        <w:abstractNum w:abstractNumId="0">
            <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
            <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/></w:lvl>
        </w:abstractNum>
        <w:abstractNum w:abstractNumId="1">
            <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl>
        </w:abstractNum>
        <w:abstractNum w:abstractNumId="2">
            <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl>
        </w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
        <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
        <w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>
        <w:num w:numId="4"><w:abstractNumId w:val="2"/></w:num>
    </w:numbering>"#;

    #[test]
    fn test_numbering_labels() {
        let sheet = StyleSheet::load(None, None, Some(NUMBERING));
        let mut c = NumberingCounters::default();
        let mut label = |num: &str, ilvl: u32| c.next_label(&sheet.numbering, num, ilvl).unwrap().0;
        assert_eq!(label("1", 0), "1.");
        assert_eq!(label("1", 1), "1.a)");
        assert_eq!(label("1", 1), "1.b)");
        assert_eq!(label("1", 0), "2.");
        assert_eq!(label("1", 1), "2.a)", "deeper levels restart");
        assert_eq!(label("2", 0), "•");
        assert_eq!(label("3", 0), "1.", "startOverride restarts the shared list");
        assert_eq!(label("1", 0), "2.", "and the list continues from there");
        assert_eq!(label("4", 0), "I.");
        assert_eq!(label("4", 0), "II.");

        let lvl = sheet.numbering.level("1", 0).unwrap();
        assert_eq!(lvl.ppr.indent_left, Some(48.0));
        assert_eq!(lvl.ppr.indent_first_line, Some(-24.0));
    }

    #[test]
    fn test_number_formats() {
        assert_eq!(format_number(27, "upperLetter"), "AA");
        assert_eq!(format_number(14, "lowerRoman"), "xiv");
        assert_eq!(format_number(3, "decimalZero"), "03");
    }
}
