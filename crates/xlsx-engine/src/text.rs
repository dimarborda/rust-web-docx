//! Text in SpreadsheetML: the `<t>` elements of shared and inline strings, the `_xHHHH_` escapes
//! Excel uses for characters XML cannot hold, and `{{FIELD}}` placeholders

use std::collections::HashMap;
use std::ops::Range;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use ooxml_core::xml::{escape_xml, set_attr, tag_is, unescape_text};

/// A `<t>` element: its start tag and its raw content
#[derive(Debug, Clone)]
pub struct TextSpan {
    pub start_tag: Range<usize>,
    pub content: Range<usize>,
}

/// The `<t>` elements inside elements named `within` (`si`, `is`), skipping phonetic runs
/// (`<rPh>`), whose text Excel does not show in the cell
pub fn text_spans(xml: &str, within: &str) -> Result<Vec<TextSpan>, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    let (mut within_depth, mut phonetic_depth) = (0usize, 0usize);
    let mut open: Option<Range<usize>> = None;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|e| format!("XML no válido: {}", e))?;
        let after = reader.buffer_position() as usize;
        match event {
            Event::Start(e) => {
                let name = e.name();
                if tag_is(name.as_ref(), within) {
                    within_depth += 1;
                } else if tag_is(name.as_ref(), "rPh") {
                    phonetic_depth += 1;
                } else if tag_is(name.as_ref(), "t") && within_depth > 0 && phonetic_depth == 0 {
                    open = Some(before..after);
                }
            }
            Event::End(e) => {
                let name = e.name();
                if tag_is(name.as_ref(), within) {
                    within_depth = within_depth.saturating_sub(1);
                } else if tag_is(name.as_ref(), "rPh") {
                    phonetic_depth = phonetic_depth.saturating_sub(1);
                } else if tag_is(name.as_ref(), "t") {
                    if let Some(start_tag) = open.take() {
                        let content = start_tag.end..before;
                        spans.push(TextSpan { start_tag, content });
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(spans)
}

/// Text of raw `<t>` content: XML entities and `_xHHHH_` escapes resolved
pub fn decode(raw: &str) -> String {
    decode_escapes(&unescape_text(raw))
}

/// `_x000D_` → carriage return; `_x005F_` stands for a literal underscore
fn decode_escapes(text: &str) -> String {
    if !text.contains("_x") {
        return text.to_string();
    }
    let mut out = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(at) = rest.find("_x") {
        out.push_str(&rest[..at]);
        let candidate = &rest[at..];
        let code = candidate.get(2..6).filter(|h| h.chars().all(|c| c.is_ascii_hexdigit()));
        match (code, candidate.as_bytes().get(6)) {
            (Some(hex), Some(b'_')) => {
                if let Some(c) = u32::from_str_radix(hex, 16).ok().and_then(char::from_u32) {
                    out.push(c);
                }
                rest = &candidate[7..];
            }
            _ => {
                out.push_str("_x");
                rest = &candidate[2..];
            }
        }
    }
    out.push_str(rest);
    out
}

/// Whether `text` starts with `_xHHHH_`
fn starts_with_escape(text: &str) -> bool {
    let b = text.as_bytes();
    b.len() >= 7 && b.starts_with(b"_x") && b[2..6].iter().all(u8::is_ascii_hexdigit) && b[6] == b'_'
}

/// Raw `<t>` content for `text`: XML-escaped, with `_xHHHH_` for control characters and for
/// literal text that would otherwise read as an escape
pub fn encode(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for (i, c) in text.char_indices() {
        let control = (c as u32) < 0x20 && !matches!(c, '\t' | '\n' | '\r');
        let looks_escaped = c == '_' && starts_with_escape(&text[i..]);
        if control || looks_escaped {
            out.push_str(&format!("_x{:04X}_", c as u32));
        } else {
            out.push(c);
        }
    }
    escape_xml(&out)
}

/// A `<t>` element for `text`, with `xml:space="preserve"` when spaces at its ends would
/// otherwise be dropped
pub fn t_element(prefix: &str, text: &str) -> String {
    let space = if needs_preserve(text) { r#" xml:space="preserve""# } else { "" };
    format!("<{p}t{}>{}</{p}t>", space, encode(text), p = prefix)
}

fn needs_preserve(text: &str) -> bool {
    text.starts_with(char::is_whitespace) || text.ends_with(char::is_whitespace) || text.contains('\n')
}

/// `text` with every `{{FIELD}}` (spaces around the name allowed) that `values` knows replaced,
/// and how many were replaced; unknown placeholders are left as they are
pub fn fill_placeholders(text: &str, values: &HashMap<String, String>) -> (String, usize) {
    let mut out = String::with_capacity(text.len());
    let mut count = 0;
    let mut rest = text;
    while let Some(open) = rest.find("{{") {
        let Some(close) = rest[open + 2..].find("}}") else { break };
        let key = rest[open + 2..open + 2 + close].trim();
        out.push_str(&rest[..open]);
        match values.get(key) {
            Some(value) => {
                out.push_str(value);
                count += 1;
            }
            None => out.push_str(&rest[open..open + 2 + close + 2]),
        }
        rest = &rest[open + 2 + close + 2..];
    }
    out.push_str(rest);
    (out, count)
}

/// `xml` with the placeholders filled in every `<t>` inside `within` elements, and how many
/// were replaced (`None` when nothing changed, so the part can be left untouched)
pub fn fill_xml(xml: &str, within: &str, values: &HashMap<String, String>) -> Result<Option<(String, usize)>, String> {
    if !xml.contains("{{") {
        return Ok(None);
    }
    let mut out = String::with_capacity(xml.len());
    let mut last = 0;
    let mut total = 0;
    for span in text_spans(xml, within)? {
        let (filled, count) = fill_placeholders(&decode(&xml[span.content.clone()]), values);
        if count == 0 {
            continue;
        }
        total += count;
        let start_tag = &xml[span.start_tag.clone()];
        out.push_str(&xml[last..span.start_tag.start]);
        if needs_preserve(&filled) && !start_tag.contains("xml:space") {
            out.push_str(&set_attr(start_tag, "xml:space", "preserve"));
        } else {
            out.push_str(start_tag);
        }
        out.push_str(&encode(&filled));
        last = span.content.end;
    }
    if total == 0 {
        return Ok(None);
    }
    out.push_str(&xml[last..]);
    Ok(Some((out, total)))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn values(pairs: &[(&str, &str)]) -> HashMap<String, String> {
        pairs.iter().map(|(k, v)| (k.to_string(), v.to_string())).collect()
    }

    #[test]
    fn escapes_round_trip() {
        for text in ["a & <b>", "línea\r\nnueva", "bell\u{7}", "_x0041_ literal", "snake_xcase", "_x12"] {
            assert_eq!(decode(&encode(text)), text, "{}", text);
        }
        assert_eq!(decode("_x0041__x000D_"), "A\r");
        assert_eq!(encode("_x0041_"), "_x005F_x0041_");
    }

    #[test]
    fn finds_shown_text_only() {
        let xml = r#"<si><r><t>Ho</t></r><r><t xml:space="preserve">la </t></r><rPh><t>ホ</t></rPh></si><x><t>no</t></x>"#;
        let texts: Vec<_> = text_spans(xml, "si").unwrap().iter().map(|s| decode(&xml[s.content.clone()])).collect();
        assert_eq!(texts, ["Ho", "la "]);
    }

    #[test]
    fn fills_known_placeholders() {
        let v = values(&[("NOMBRE", "Ana & Luis"), ("X", " ")]);
        assert_eq!(fill_placeholders("Hola {{NOMBRE}}, {{ NOMBRE }} y {{OTRO}} {{", &v), ("Hola Ana & Luis, Ana & Luis y {{OTRO}} {{".to_string(), 2));

        let xml = r#"<sst><si><t>Para {{NOMBRE}}</t></si><si><t>{{X}}</t></si><si><t>{{OTRO}}</t></si></sst>"#;
        let (out, count) = fill_xml(xml, "si", &v).unwrap().unwrap();
        assert_eq!(count, 2);
        assert_eq!(out, r#"<sst><si><t>Para Ana &amp; Luis</t></si><si><t xml:space="preserve"> </t></si><si><t>{{OTRO}}</t></si></sst>"#);
        assert!(fill_xml(xml, "si", &values(&[("Z", "1")])).unwrap().is_none());
    }
}
