//! Colors of SpreadsheetML: `rgb` (ARGB), `indexed` (the legacy 64-color palette), `theme`
//! (a color of `xl/theme/theme1.xml`) and `tint` (lighter or darker), all resolved to `#RRGGBB`

use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;

use ooxml_core::xml::{get_attr_value, tag_is};

/// The legacy palette `indexed="n"` points to (64 is the system text color, 65 the background)
const INDEXED: [&str; 64] = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", //
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", //
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080", //
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", //
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF", //
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99", //
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696", //
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
];

/// A color of the legacy palette as RRGGBB
pub fn indexed_color(index: usize) -> Option<&'static str> {
    INDEXED.get(index).copied()
}

/// Office's default theme colors, used when the workbook has no theme part
const DEFAULT_THEME: [&str; 12] = [
    "FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72",
];

/// Theme colors in the order `theme="n"` counts them: lt1, dk1, lt2, dk2, accent1–6, hlink,
/// folHlink (the theme part lists dark before light; Excel swaps each pair)
#[derive(Debug, Clone)]
pub struct Theme {
    colors: Vec<String>,
}

impl Default for Theme {
    fn default() -> Self {
        Theme { colors: DEFAULT_THEME.iter().map(|c| c.to_string()).collect() }
    }
}

const SCHEME: [&str; 12] = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

impl Theme {
    /// Reads the color scheme of a theme part; missing colors keep Office's defaults
    pub fn parse(xml: &str) -> Theme {
        let mut scheme: Vec<Option<String>> = vec![None; SCHEME.len()];
        let mut reader = Reader::from_str(xml);
        let mut in_scheme = false;
        let mut current: Option<usize> = None;
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.name();
                    let name = name.as_ref();
                    if tag_is(name, "clrScheme") {
                        in_scheme = true;
                    } else if in_scheme {
                        if let Some(i) = SCHEME.iter().position(|s| tag_is(name, s)) {
                            current = Some(i);
                        } else if let Some(i) = current {
                            if tag_is(name, "srgbClr") {
                                scheme[i] = get_attr_value(&e, "val");
                            } else if tag_is(name, "sysClr") {
                                scheme[i] = get_attr_value(&e, "lastClr");
                            }
                        }
                    }
                }
                Ok(Event::End(e)) if tag_is(e.name().as_ref(), "clrScheme") => break,
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        let mut theme = Theme::default();
        // Scheme order (dk1, lt1, dk2, lt2, …) to theme index order (lt1, dk1, lt2, dk2, …)
        for (i, color) in scheme.into_iter().enumerate() {
            if let Some(color) = color.filter(|c| c.len() == 6) {
                let index = match i {
                    0 => 1,
                    1 => 0,
                    2 => 3,
                    3 => 2,
                    n => n,
                };
                theme.colors[index] = color.to_ascii_uppercase();
            }
        }
        theme
    }

    /// The color of a `<color>`, `<fgColor>`, `<bgColor>`… element; `None` for `auto` and for
    /// elements that name no color
    pub fn resolve(&self, e: &BytesStart) -> Option<String> {
        if get_attr_value(e, "auto").is_some_and(|v| v == "1" || v == "true") {
            return None;
        }
        let base = if let Some(rgb) = get_attr_value(e, "rgb") {
            let rgb = rgb.trim().to_ascii_uppercase();
            match rgb.len() {
                8 => rgb[2..].to_string(),
                6 => rgb,
                _ => return None,
            }
        } else if let Some(theme) = get_attr_value(e, "theme").and_then(|v| v.parse::<usize>().ok()) {
            self.colors.get(theme)?.clone()
        } else if let Some(indexed) = get_attr_value(e, "indexed").and_then(|v| v.parse::<usize>().ok()) {
            match indexed {
                64 => "000000".to_string(),
                65 => "FFFFFF".to_string(),
                n => INDEXED.get(n)?.to_string(),
            }
        } else {
            return None;
        };
        let tint = get_attr_value(e, "tint").and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0);
        Some(format!("#{}", apply_tint(&base, tint)))
    }
}

/// Lightens (tint > 0) or darkens (tint < 0) an RRGGBB color through its HSL lightness
pub fn apply_tint(rgb: &str, tint: f64) -> String {
    if tint == 0.0 {
        return rgb.to_string();
    }
    let channel = |i: usize| u8::from_str_radix(&rgb[i..i + 2], 16).unwrap_or(0) as f64 / 255.0;
    let (r, g, b) = (channel(0), channel(2), channel(4));
    let (h, s, mut l) = rgb_to_hsl(r, g, b);
    l = if tint < 0.0 { l * (1.0 + tint) } else { l * (1.0 - tint) + tint };
    let (r, g, b) = hsl_to_rgb(h, s, l.clamp(0.0, 1.0));
    let to_hex = |v: f64| format!("{:02X}", (v * 255.0).round().clamp(0.0, 255.0) as u8);
    format!("{}{}{}", to_hex(r), to_hex(g), to_hex(b))
}

fn rgb_to_hsl(r: f64, g: f64, b: f64) -> (f64, f64, f64) {
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    if max == min {
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

fn hsl_to_rgb(h: f64, s: f64, l: f64) -> (f64, f64, f64) {
    if s == 0.0 {
        return (l, l, l);
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
    (hue(h + 1.0 / 3.0), hue(h), hue(h - 1.0 / 3.0))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn color(xml: &str, theme: &Theme) -> Option<String> {
        let mut reader = Reader::from_str(xml);
        match reader.read_event().unwrap() {
            Event::Empty(e) | Event::Start(e) => theme.resolve(&e),
            _ => None,
        }
    }

    #[test]
    fn resolves_rgb_indexed_theme_and_auto() {
        let theme = Theme::default();
        assert_eq!(color(r#"<color rgb="FF1F4E78"/>"#, &theme).as_deref(), Some("#1F4E78"));
        assert_eq!(color(r#"<color indexed="10"/>"#, &theme).as_deref(), Some("#FF0000"));
        assert_eq!(color(r#"<color indexed="64"/>"#, &theme).as_deref(), Some("#000000"));
        assert_eq!(color(r#"<color theme="1"/>"#, &theme).as_deref(), Some("#000000"), "theme 1 is dk1");
        assert_eq!(color(r#"<color theme="4"/>"#, &theme).as_deref(), Some("#4472C4"));
        assert_eq!(color(r#"<color auto="1"/>"#, &theme), None);
        assert_eq!(color(r#"<color/>"#, &theme), None);
    }

    /// Same color within `tolerance` per channel (Office rounds through integer HLS)
    fn close(actual: &str, expected: &str, tolerance: i32) -> bool {
        (0..3).all(|i| {
            let a = i32::from_str_radix(&actual[i * 2..i * 2 + 2], 16).unwrap();
            let e = i32::from_str_radix(&expected[i * 2..i * 2 + 2], 16).unwrap();
            (a - e).abs() <= tolerance
        })
    }

    #[test]
    fn applies_tints_like_office() {
        // "Blue, Accent 1, Lighter 80%" and "Darker 25%" of the Office theme
        let lighter = apply_tint("4472C4", 0.7999816888943144);
        assert!(close(&lighter, "D9E2F3", 2), "{}", lighter);
        let darker = apply_tint("4472C4", -0.249977111117893);
        assert!(close(&darker, "2F5496", 2), "{}", darker);
        assert_eq!(apply_tint("FFFFFF", -0.1499984740745262), "D9D9D9");
        assert_eq!(apply_tint("1F4E78", 0.0), "1F4E78");
    }

    #[test]
    fn reads_the_theme_swapping_dark_and_light() {
        let xml = r#"<a:theme xmlns:a="x"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FEFEFE"/></a:lt1><a:dk2><a:srgbClr val="0E2841"/></a:dk2><a:lt2><a:srgbClr val="E8E8E8"/></a:lt2><a:accent1><a:srgbClr val="156082"/></a:accent1></a:clrScheme></a:themeElements></a:theme>"#;
        let theme = Theme::parse(xml);
        assert_eq!(color(r#"<color theme="0"/>"#, &theme).as_deref(), Some("#FEFEFE"));
        assert_eq!(color(r#"<color theme="1"/>"#, &theme).as_deref(), Some("#111111"));
        assert_eq!(color(r#"<color theme="3"/>"#, &theme).as_deref(), Some("#0E2841"));
        assert_eq!(color(r#"<color theme="4"/>"#, &theme).as_deref(), Some("#156082"));
        assert_eq!(color(r#"<color theme="5"/>"#, &theme).as_deref(), Some("#ED7D31"), "missing accents keep the default");
    }
}
