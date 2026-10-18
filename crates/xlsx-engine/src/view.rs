//! What a viewer needs to draw a sheet: the cell formats resolved to fonts, fills, borders and
//! alignment (`StyleSheetView`), the grid (`SheetLayout`: column widths, row heights, merged
//! cells, frozen panes) and the text each cell shows (`CellView`), formatted as Excel does

use std::collections::HashMap;

use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;
use serde_json::{json, Value};

use ooxml_core::xml::{get_attr_value, tag_is};

use crate::cell_ref::CellRange;
use crate::dates::{iso_to_unix_ms, unix_ms_to_serial};
use crate::numfmt::{builtin_code, format, Locale, Value as FmtValue};
use crate::shared_strings::SharedStrings;
use crate::sheet::{CellNode, SheetXml};
use crate::theme::Theme;

#[derive(Debug, Clone, Default, PartialEq)]
pub struct FontView {
    pub name: String,
    /// Points
    pub size: f64,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    pub color: Option<String>,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct BorderView {
    /// (style, color) of each side: left, right, top, bottom
    pub sides: [Option<(String, Option<String>)>; 4],
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct XfView {
    pub num_fmt: u32,
    pub font: usize,
    pub fill: usize,
    pub border: usize,
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap: bool,
    pub shrink: bool,
    pub indent: u32,
    pub rotation: i32,
}

/// The stylesheet resolved for drawing
#[derive(Debug, Default)]
pub struct StyleSheetView {
    pub custom_formats: HashMap<u32, String>,
    pub fonts: Vec<FontView>,
    /// Background color of each fill
    pub fills: Vec<Option<String>>,
    pub borders: Vec<BorderView>,
    pub xfs: Vec<XfView>,
}

const SIDES: [&str; 4] = ["left", "right", "top", "bottom"];

fn attr_bool(e: &BytesStart, name: &str) -> Option<bool> {
    get_attr_value(e, name).map(|v| !matches!(v.trim(), "0" | "false" | "none"))
}

/// A boolean font property such as `<b/>` or `<b val="0"/>`
fn flag(e: &BytesStart) -> bool {
    attr_bool(e, "val").unwrap_or(true)
}

impl StyleSheetView {
    pub fn parse(xml: &str, theme: &Theme) -> Result<StyleSheetView, String> {
        let mut view = StyleSheetView::default();
        let mut reader = Reader::from_str(xml);
        #[derive(PartialEq)]
        enum In {
            None,
            Fonts,
            Fills,
            Borders,
            CellXfs,
            Other,
        }
        let mut section = In::None;
        let mut font: Option<FontView> = None;
        let mut fill: Option<(Option<String>, Option<String>, Option<String>)> = None; // pattern, fg, bg
        let mut gradient_color: Option<String> = None;
        let mut border: Option<BorderView> = None;
        let mut side: Option<usize> = None;
        loop {
            let event = reader.read_event().map_err(|e| format!("La hoja de estilos no es XML válido: {}", e))?;
            let (e, empty) = match &event {
                Event::Start(e) => (Some(e.clone()), false),
                Event::Empty(e) => (Some(e.clone()), true),
                _ => (None, false),
            };
            if let Some(e) = e {
                let name = e.name();
                let name = name.as_ref();
                if tag_is(name, "numFmt") {
                    let id = get_attr_value(&e, "numFmtId").and_then(|v| v.parse().ok());
                    if let (Some(id), Some(code)) = (id, get_attr_value(&e, "formatCode")) {
                        view.custom_formats.insert(id, code);
                    }
                } else if tag_is(name, "fonts") && !empty {
                    section = In::Fonts;
                } else if tag_is(name, "fills") && !empty {
                    section = In::Fills;
                } else if tag_is(name, "borders") && !empty {
                    section = In::Borders;
                } else if tag_is(name, "cellXfs") && !empty {
                    section = In::CellXfs;
                } else if (tag_is(name, "cellStyleXfs") || tag_is(name, "dxfs") || tag_is(name, "cellStyles")) && !empty {
                    section = In::Other;
                } else {
                    match section {
                        In::Fonts => {
                            if tag_is(name, "font") {
                                font = Some(FontView { name: "Calibri".to_string(), size: 11.0, ..Default::default() });
                                if empty {
                                    view.fonts.push(font.take().unwrap());
                                }
                            } else if let Some(f) = font.as_mut() {
                                if tag_is(name, "name") || tag_is(name, "rFont") {
                                    if let Some(v) = get_attr_value(&e, "val") {
                                        f.name = v;
                                    }
                                } else if tag_is(name, "sz") {
                                    f.size = get_attr_value(&e, "val").and_then(|v| v.parse().ok()).unwrap_or(11.0);
                                } else if tag_is(name, "b") {
                                    f.bold = flag(&e);
                                } else if tag_is(name, "i") {
                                    f.italic = flag(&e);
                                } else if tag_is(name, "strike") {
                                    f.strike = flag(&e);
                                } else if tag_is(name, "u") {
                                    f.underline = get_attr_value(&e, "val").map(|v| v != "none").unwrap_or(true);
                                } else if tag_is(name, "color") {
                                    f.color = theme.resolve(&e);
                                }
                            }
                        }
                        In::Fills => {
                            if tag_is(name, "fill") {
                                fill = Some((None, None, None));
                                gradient_color = None;
                                if empty {
                                    view.fills.push(None);
                                    fill = None;
                                }
                            } else if let Some(f) = fill.as_mut() {
                                if tag_is(name, "patternFill") {
                                    f.0 = Some(get_attr_value(&e, "patternType").unwrap_or_else(|| "none".to_string()));
                                } else if tag_is(name, "fgColor") {
                                    f.1 = theme.resolve(&e);
                                } else if tag_is(name, "bgColor") {
                                    f.2 = theme.resolve(&e);
                                } else if tag_is(name, "color") && gradient_color.is_none() {
                                    gradient_color = theme.resolve(&e);
                                }
                            }
                        }
                        In::Borders => {
                            if tag_is(name, "border") {
                                border = Some(BorderView::default());
                                if empty {
                                    view.borders.push(border.take().unwrap());
                                }
                            } else if let Some(b) = border.as_mut() {
                                let index = SIDES.iter().position(|s| tag_is(name, s)).or_else(|| {
                                    if tag_is(name, "start") {
                                        Some(0)
                                    } else if tag_is(name, "end") {
                                        Some(1)
                                    } else {
                                        None
                                    }
                                });
                                if let Some(i) = index {
                                    side = None;
                                    if let Some(style) = get_attr_value(&e, "style").filter(|s| s != "none") {
                                        b.sides[i] = Some((style, None));
                                        if !empty {
                                            side = Some(i);
                                        }
                                    }
                                } else if tag_is(name, "color") {
                                    if let Some(i) = side {
                                        if let Some(s) = b.sides[i].as_mut() {
                                            s.1 = theme.resolve(&e);
                                        }
                                    }
                                }
                            }
                        }
                        In::CellXfs => {
                            if tag_is(name, "xf") {
                                let id = |n: &str| get_attr_value(&e, n).and_then(|v| v.parse().ok()).unwrap_or(0);
                                view.xfs.push(XfView {
                                    num_fmt: id("numFmtId") as u32,
                                    font: id("fontId"),
                                    fill: id("fillId"),
                                    border: id("borderId"),
                                    ..Default::default()
                                });
                            } else if tag_is(name, "alignment") {
                                if let Some(xf) = view.xfs.last_mut() {
                                    xf.horizontal = get_attr_value(&e, "horizontal");
                                    xf.vertical = get_attr_value(&e, "vertical");
                                    xf.wrap = attr_bool(&e, "wrapText").unwrap_or(false);
                                    xf.shrink = attr_bool(&e, "shrinkToFit").unwrap_or(false);
                                    xf.indent = get_attr_value(&e, "indent").and_then(|v| v.parse().ok()).unwrap_or(0);
                                    xf.rotation = get_attr_value(&e, "textRotation").and_then(|v| v.parse().ok()).unwrap_or(0);
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            match &event {
                Event::End(e) => {
                    let name = e.name();
                    let name = name.as_ref();
                    if tag_is(name, "font") && section == In::Fonts {
                        if let Some(f) = font.take() {
                            view.fonts.push(f);
                        }
                    } else if tag_is(name, "fill") && section == In::Fills {
                        if let Some((pattern, fg, bg)) = fill.take() {
                            let color = match pattern.as_deref() {
                                Some("none") | Some("gray125") => None,
                                // Solid fills paint the foreground color; other patterns are
                                // drawn as their foreground too, the closest flat color
                                Some(_) => fg.or(bg),
                                None => gradient_color.take(),
                            };
                            view.fills.push(color);
                        }
                    } else if tag_is(name, "border") && section == In::Borders {
                        if let Some(b) = border.take() {
                            view.borders.push(b);
                        }
                        side = None;
                    } else if tag_is(name, "fonts") || tag_is(name, "fills") || tag_is(name, "borders") || tag_is(name, "cellXfs") || tag_is(name, "cellStyleXfs") || tag_is(name, "dxfs") || tag_is(name, "cellStyles") {
                        section = In::None;
                    }
                }
                Event::Eof => break,
                _ => {}
            }
        }
        if view.fonts.is_empty() {
            view.fonts.push(FontView { name: "Calibri".to_string(), size: 11.0, ..Default::default() });
        }
        if view.xfs.is_empty() {
            view.xfs.push(XfView::default());
        }
        Ok(view)
    }

    /// The format code of a number format id
    pub fn format_code<'a>(&'a self, id: u32, locale: &Locale) -> &'a str {
        match self.custom_formats.get(&id) {
            Some(code) => code,
            None => builtin_code(id, locale),
        }
    }

    /// Every cell format as JSON for the viewer: `{font, fill, border, h, v, wrap, …}`
    pub fn to_json(&self) -> Value {
        let default_font = &self.fonts[0];
        let xfs: Vec<Value> = self
            .xfs
            .iter()
            .map(|xf| {
                let font = self.fonts.get(xf.font).unwrap_or(default_font);
                let border = self.borders.get(xf.border).cloned().unwrap_or_default();
                let sides: Vec<Value> = border
                    .sides
                    .iter()
                    .map(|s| match s {
                        Some((style, color)) => json!([style, color]),
                        None => Value::Null,
                    })
                    .collect();
                json!({
                    "font": {
                        "name": font.name, "size": font.size, "bold": font.bold, "italic": font.italic,
                        "underline": font.underline, "strike": font.strike, "color": font.color,
                    },
                    "fill": self.fills.get(xf.fill).cloned().flatten(),
                    "border": sides,
                    "h": xf.horizontal, "v": xf.vertical, "wrap": xf.wrap, "shrink": xf.shrink,
                    "indent": xf.indent, "rotation": xf.rotation,
                })
            })
            .collect();
        json!({ "defaultFont": { "name": default_font.name, "size": default_font.size }, "xfs": xfs })
    }
}

/// Grid of a worksheet
#[derive(Debug, Default, PartialEq)]
pub struct SheetLayout {
    /// Points
    pub default_row_height: Option<f64>,
    /// Characters of the default font's widest digit, padding included
    pub default_col_width: Option<f64>,
    pub base_col_width: Option<f64>,
    /// (first column, last column, width or None, hidden)
    pub cols: Vec<(u32, u32, Option<f64>, bool)>,
    /// (row, height in points or None, hidden)
    pub rows: Vec<(u32, Option<f64>, bool)>,
    pub merges: Vec<CellRange>,
    pub gridlines: bool,
    pub headings: bool,
    pub right_to_left: bool,
    /// Frozen rows and columns at the top and left
    pub frozen: Option<(u32, u32)>,
}

impl SheetLayout {
    pub fn parse(xml: &str) -> Result<SheetLayout, String> {
        let mut layout = SheetLayout { gridlines: true, headings: true, ..Default::default() };
        let mut reader = Reader::from_str(xml);
        let mut views = 0;
        let mut in_cell = false;
        loop {
            let event = reader.read_event().map_err(|e| format!("La hoja no es XML válido: {}", e))?;
            match event {
                Event::Start(e) | Event::Empty(e) => {
                    let name = e.name();
                    let name = name.as_ref();
                    if in_cell {
                        continue;
                    }
                    if tag_is(name, "c") {
                        in_cell = !reader_is_empty(&e, xml, reader.buffer_position() as usize);
                    } else if tag_is(name, "row") {
                        let r = get_attr_value(&e, "r").and_then(|v| v.parse().ok());
                        let ht = get_attr_value(&e, "ht").and_then(|v| v.parse().ok());
                        let hidden = attr_bool(&e, "hidden").unwrap_or(false);
                        let r = r.unwrap_or_else(|| layout.rows.last().map(|x| x.0 + 1).unwrap_or(1));
                        if ht.is_some() || hidden {
                            layout.rows.push((r, ht, hidden));
                        } else {
                            // Keep the numbering of rows without `r` working
                            layout.rows.push((r, None, false));
                        }
                    } else if tag_is(name, "col") {
                        let n = |a: &str| get_attr_value(&e, a).and_then(|v| v.parse::<u32>().ok());
                        if let (Some(min), Some(max)) = (n("min"), n("max")) {
                            let width = get_attr_value(&e, "width").and_then(|v| v.parse().ok());
                            layout.cols.push((min, max.min(16_384), width, attr_bool(&e, "hidden").unwrap_or(false)));
                        }
                    } else if tag_is(name, "sheetFormatPr") {
                        let f = |a: &str| get_attr_value(&e, a).and_then(|v| v.parse::<f64>().ok());
                        layout.default_row_height = f("defaultRowHeight");
                        layout.default_col_width = f("defaultColWidth");
                        layout.base_col_width = f("baseColWidth");
                    } else if tag_is(name, "sheetView") {
                        views += 1;
                        if views == 1 {
                            layout.gridlines = attr_bool(&e, "showGridLines").unwrap_or(true);
                            layout.headings = attr_bool(&e, "showRowColHeaders").unwrap_or(true);
                            layout.right_to_left = attr_bool(&e, "rightToLeft").unwrap_or(false);
                        }
                    } else if tag_is(name, "pane") && views == 1 {
                        let state = get_attr_value(&e, "state").unwrap_or_default();
                        if state == "frozen" || state == "frozenSplit" {
                            let f = |a: &str| get_attr_value(&e, a).and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) as u32;
                            let (cols, rows) = (f("xSplit"), f("ySplit"));
                            if cols > 0 || rows > 0 {
                                layout.frozen = Some((rows, cols));
                            }
                        }
                    } else if tag_is(name, "mergeCell") {
                        if let Some(range) = get_attr_value(&e, "ref").and_then(|r| CellRange::parse(&r).ok()) {
                            layout.merges.push(range);
                        }
                    }
                }
                Event::End(e) if in_cell && tag_is(e.name().as_ref(), "c") => in_cell = false,
                Event::Eof => break,
                _ => {}
            }
        }
        layout.rows.retain(|r| r.1.is_some() || r.2);
        Ok(layout)
    }

    pub fn to_json(&self) -> Value {
        json!({
            "defaultRowHeight": self.default_row_height,
            "defaultColWidth": self.default_col_width,
            "baseColWidth": self.base_col_width,
            "cols": self.cols.iter().map(|c| json!([c.0, c.1, c.2, c.3])).collect::<Vec<_>>(),
            "rows": self.rows.iter().map(|r| json!([r.0, r.1, r.2])).collect::<Vec<_>>(),
            "merges": self.merges.iter().map(|m| m.to_string()).collect::<Vec<_>>(),
            "gridlines": self.gridlines,
            "headings": self.headings,
            "rightToLeft": self.right_to_left,
            "frozen": self.frozen.map(|(r, c)| json!({ "rows": r, "cols": c })),
        })
    }
}

/// Whether the start tag just read ends with `/>`
fn reader_is_empty(_e: &BytesStart, xml: &str, after: usize) -> bool {
    xml.as_bytes().get(after.wrapping_sub(2)) == Some(&b'/')
}

/// One cell as the viewer draws it
#[derive(Debug, Clone, PartialEq)]
pub struct CellView {
    pub row: u32,
    pub col: u32,
    pub style: u32,
    pub text: String,
    /// `n` number or date, `s` text, `b` boolean, `e` error, `` empty: decides the default
    /// alignment
    pub kind: &'static str,
    /// Color from the number format (`[Red]`)
    pub color: Option<String>,
}

/// The cells of a worksheet that show something: a value or a format other than the default
pub fn cell_views(xml: &str, strings: &SharedStrings, styles: &StyleSheetView, locale: &Locale, date1904: bool) -> Result<Vec<CellView>, String> {
    let sheet = SheetXml::parse(xml)?;
    let mut out = Vec::new();
    for row in &sheet.rows {
        for cell in &row.cells {
            if let Some(view) = cell_view(cell, strings, styles, locale, date1904) {
                out.push(view);
            }
        }
    }
    Ok(out)
}

fn cell_view(node: &CellNode, strings: &SharedStrings, styles: &StyleSheetView, locale: &Locale, date1904: bool) -> Option<CellView> {
    let raw = node.value.as_deref().unwrap_or("");
    let text_value: Option<String>;
    let (value, kind) = match node.kind.as_deref() {
        Some("s") => {
            text_value = raw.trim().parse::<usize>().ok().and_then(|i| strings.get(i)).map(str::to_string);
            (text_value.as_deref().map(FmtValue::Text), "s")
        }
        Some("str") => (Some(FmtValue::Text(raw)), "s"),
        Some("inlineStr") => (Some(FmtValue::Text(node.inline.as_deref().unwrap_or(raw))), "s"),
        Some("b") => (Some(FmtValue::Bool(matches!(raw.trim(), "1" | "true"))), "b"),
        Some("e") => (Some(FmtValue::Error(raw)), "e"),
        Some("d") => match iso_to_unix_ms(raw).and_then(|ms| unix_ms_to_serial(ms, date1904).ok()) {
            Some(serial) => (Some(FmtValue::Number(serial)), "n"),
            None => (Some(FmtValue::Text(raw)), "s"),
        },
        _ if raw.trim().is_empty() => (None, ""),
        _ => match raw.trim().parse::<f64>() {
            Ok(n) => (Some(FmtValue::Number(n)), "n"),
            Err(_) => (Some(FmtValue::Text(raw)), "s"),
        },
    };
    let xf = styles.xfs.get(node.style as usize);
    let Some(value) = value else {
        return (node.style != 0).then(|| CellView { row: node.at.row, col: node.at.col, style: node.style, text: String::new(), kind: "", color: None });
    };
    let code = styles.format_code(xf.map(|x| x.num_fmt).unwrap_or(0), locale);
    let formatted = format(value, code, locale, date1904);
    Some(CellView { row: node.at.row, col: node.at.col, style: node.style, text: formatted.text, kind, color: formatted.color })
}

/// Cells as compact JSON rows: `[row, col, style, text, kind, color?]`
pub fn cells_json(cells: &[CellView]) -> Value {
    Value::Array(
        cells
            .iter()
            .map(|c| match &c.color {
                Some(color) => json!([c.row, c.col, c.style, c.text, c.kind, color]),
                None => json!([c.row, c.col, c.style, c.text, c.kind]),
            })
            .collect(),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    const STYLES: &str = r##"<styleSheet xmlns="x"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00 &quot;€&quot;"/></numFmts>
<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Calibri"/></font><font><b/><i val="0"/><u/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color indexed="64"/></left><right/><top/><bottom style="medium"/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0"><alignment horizontal="center" vertical="top" wrapText="1"/></xf><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>"##;

    #[test]
    fn resolves_fonts_fills_borders_and_alignment() {
        let view = StyleSheetView::parse(STYLES, &Theme::default()).unwrap();
        assert_eq!(view.fonts[0].color.as_deref(), Some("#000000"));
        let f = &view.fonts[1];
        assert!(f.bold && !f.italic && f.underline && !f.strike);
        assert_eq!((f.name.as_str(), f.size, f.color.as_deref()), ("Arial", 16.0, Some("#FFFFFF")));
        assert_eq!(view.fills[0], None);
        assert_eq!(view.fills[1], None, "gray125 is the reserved pattern, not a background");
        assert!(view.fills[2].as_deref().unwrap().starts_with('#'));
        assert_eq!(view.borders[1].sides[0], Some(("thin".to_string(), Some("#000000".to_string()))));
        assert_eq!(view.borders[1].sides[1], None);
        assert_eq!(view.borders[1].sides[3], Some(("medium".to_string(), None)));
        let xf = &view.xfs[1];
        assert_eq!((xf.font, xf.fill, xf.border, xf.num_fmt), (1, 2, 1, 164));
        assert_eq!((xf.horizontal.as_deref(), xf.vertical.as_deref(), xf.wrap), (Some("center"), Some("top"), true));
        assert_eq!(view.xfs.len(), 3, "cellStyleXfs are not cell formats");
        let json = view.to_json();
        assert_eq!(json["xfs"][1]["font"]["name"], "Arial");
        assert_eq!(json["defaultFont"]["size"], 11.0);
    }

    #[test]
    fn reads_the_grid() {
        let xml = r#"<worksheet xmlns="x"><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" xSplit="2" topLeftCell="C2" state="frozen"/></sheetView></sheetViews>
<sheetFormatPr baseColWidth="10" defaultRowHeight="16"/><cols><col min="1" max="1" width="22" customWidth="1"/><col min="3" max="5" width="0" hidden="1"/></cols>
<sheetData><row r="1" ht="30" customHeight="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>row</t></is></c></row><row r="4" hidden="1"/></sheetData>
<mergeCells count="1"><mergeCell ref="A1:D1"/></mergeCells></worksheet>"#;
        let layout = SheetLayout::parse(xml).unwrap();
        assert!(!layout.gridlines);
        assert_eq!(layout.frozen, Some((1, 2)));
        assert_eq!((layout.default_row_height, layout.base_col_width), (Some(16.0), Some(10.0)));
        assert_eq!(layout.cols, vec![(1, 1, Some(22.0), false), (3, 5, Some(0.0), true)]);
        assert_eq!(layout.rows, vec![(1, Some(30.0), false), (4, None, true)]);
        assert_eq!(layout.merges[0].to_string(), "A1:D1");
    }

    #[test]
    fn formats_the_cells_it_shows() {
        let styles = StyleSheetView::parse(STYLES, &Theme::default()).unwrap();
        let strings = SharedStrings::parse(r#"<sst><si><t>Hola</t></si></sst>"#).unwrap();
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="1"><v>1234.5</v></c><c r="C1" s="2"><v>46037</v></c><c r="D1" t="b"><v>1</v></c><c r="E1" s="1"/><c r="F1"/><c r="G1" t="e"><f>1/0</f><v>#DIV/0!</v></c></row></sheetData></worksheet>"#;
        let cells = cell_views(xml, &strings, &styles, &Locale::from_tag("es"), false).unwrap();
        let texts: Vec<(&str, &str)> = cells.iter().map(|c| (c.text.as_str(), c.kind)).collect();
        assert_eq!(texts, vec![("Hola", "s"), ("1.234,50 €", "n"), ("15/01/2026", "n"), ("VERDADERO", "b"), ("", ""), ("#DIV/0!", "e")]);
        assert_eq!(cells[4].col, 5, "an empty cell with a format is drawn; one without is not");
        assert_eq!(cells_json(&cells[..1]), json!([[1, 1, 0, "Hola", "s"]]));
    }
}
