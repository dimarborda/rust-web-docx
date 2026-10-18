//! Small XML helpers that work on any OOXML part

use std::borrow::Cow;

use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;

/// Whether an element name (with or without a namespace prefix) is `name`
pub fn tag_is(tag: impl AsRef<[u8]>, name: &str) -> bool {
    if let Ok(s) = std::str::from_utf8(tag.as_ref()) {
        s == name || s.ends_with(&format!(":{}", name))
    } else {
        false
    }
}

/// An element or attribute name as text, whether given as text or bytes
pub fn utf8<T: AsRef<[u8]> + ?Sized>(name: &T) -> &str {
    std::str::from_utf8(name.as_ref()).unwrap_or("")
}

/// Value of the attribute `local_name` (with or without a namespace prefix)
pub fn get_attr_value(e: &BytesStart, local_name: &str) -> Option<String> {
    for attr in e.attributes().flatten() {
        let key = attr.key.as_ref();
        let key_str = utf8(key);
        if key_str == local_name || key_str.ends_with(&format!(":{}", local_name)) {
            return attr.normalized_value(quick_xml::XmlVersion::Implicit1_0).ok().map(|s| s.to_string());
        }
    }
    None
}

/// Escapes text for use in element content or attribute values
pub fn escape_xml(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Appends `child` as the last child of the root element `root` (which may be self-closing)
pub fn append_child(xml: &str, root: &str, child: &str) -> Result<String, String> {
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

/// Text of raw element content (`a &amp; b` → "a & b"); unknown entities are kept as written
pub fn unescape_text(raw: &str) -> Cow<'_, str> {
    quick_xml::escape::unescape(raw).unwrap_or(Cow::Borrowed(raw))
}

/// Index of the `>` that closes the start tag `element` begins with (quotes are skipped)
fn start_tag_end(element: &str) -> Option<usize> {
    let mut quote = None;
    for (i, c) in element.char_indices() {
        match (quote, c) {
            (Some(q), c) if c == q => quote = None,
            (Some(_), _) => {}
            (None, '"') | (None, '\'') => quote = Some(c),
            (None, '>') => return Some(i),
            _ => {}
        }
    }
    None
}

/// Byte range of the value of the attribute `name` in the start tag `tag`, quotes excluded
fn attr_value_range(tag: &str, name: &str) -> Option<std::ops::Range<usize>> {
    let needle = format!("{}=", name);
    let mut from = 0;
    while let Some(found) = tag[from..].find(&needle) {
        let at = from + found;
        let after = at + needle.len();
        let preceded_by_space = tag[..at].ends_with(|c: char| c.is_ascii_whitespace());
        if let (true, Some(q)) = (preceded_by_space, tag[after..].chars().next().filter(|c| *c == '"' || *c == '\'')) {
            let end = tag[after + 1..].find(q)? + after + 1;
            return Some(after + 1..end);
        }
        from = after;
    }
    None
}

/// `element` (which starts with its start tag) with the attribute `name` set to `value` in its
/// start tag; the rest of the element is left as written
pub fn set_attr(element: &str, name: &str, value: &str) -> String {
    let Some(end) = start_tag_end(element) else { return element.to_string() };
    let tag = &element[..end];
    let value = escape_xml(value);
    if let Some(range) = attr_value_range(tag, name) {
        return format!("{}{}{}", &element[..range.start], value, &element[range.end..]);
    }
    let insert_at = if tag.ends_with('/') { end - 1 } else { end };
    format!("{} {}=\"{}\"{}", &element[..insert_at], name, value, &element[insert_at..])
}

/// `element` without the attribute `name` in its start tag
pub fn remove_attr(element: &str, name: &str) -> String {
    let Some(end) = start_tag_end(element) else { return element.to_string() };
    match attr_value_range(&element[..end], name) {
        Some(range) => {
            let start = element[..range.start - 1 - name.len() - 1].trim_end().len();
            format!("{}{}", &element[..start], &element[range.end + 1..])
        }
        None => element.to_string(),
    }
}

/// Byte offset and qualified name (`sst`, `x:worksheet`) of the root element
pub fn root_element(xml: &str) -> Option<(usize, String)> {
    let mut reader = Reader::from_str(xml);
    loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => return Some((before, utf8(e.name().as_ref()).to_string())),
            Ok(Event::Eof) | Err(_) => return None,
            _ => {}
        }
    }
}

/// Namespace prefix of a qualified name, with its colon (`x:sheetData` → `x:`, `row` → ``)
pub fn prefix_of(qualified: &str) -> &str {
    qualified.rfind(':').map(|i| &qualified[..=i]).unwrap_or("")
}

/// Byte ranges of the elements named `local_name` (start tag to end tag, or the empty tag) for
/// which `keep` returns true, in document order. Elements nested in a match are not visited.
pub fn element_ranges(
    xml: &str,
    local_name: &str,
    mut keep: impl FnMut(&BytesStart) -> bool,
) -> Result<Vec<std::ops::Range<usize>>, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut ranges = Vec::new();
    // Start of the open match and how deep inside it the reader is
    let mut open: Option<(usize, usize)> = None;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|e| format!("XML no válido: {}", e))?;
        let after = reader.buffer_position() as usize;
        match (&event, open.as_mut()) {
            (Event::Eof, _) => break,
            (Event::Start(_), Some((_, depth))) => *depth += 1,
            (Event::End(_), Some((start, depth))) => {
                if *depth == 0 {
                    ranges.push(*start..after);
                    open = None;
                } else {
                    *depth -= 1;
                }
            }
            (Event::Start(e), None) if tag_is(e.name().as_ref(), local_name) && keep(e) => open = Some((before, 0)),
            (Event::Empty(e), None) if tag_is(e.name().as_ref(), local_name) && keep(e) => ranges.push(before..after),
            _ => {}
        }
    }
    Ok(ranges)
}

/// `xml` without the elements named `local_name` for which `remove` returns true
pub fn remove_elements(xml: &str, local_name: &str, remove: impl FnMut(&BytesStart) -> bool) -> Result<String, String> {
    let ranges = element_ranges(xml, local_name, remove)?;
    let mut out = String::with_capacity(xml.len());
    let mut last = 0;
    for range in ranges {
        out.push_str(&xml[last..range.start]);
        last = range.end;
    }
    out.push_str(&xml[last..]);
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn appends_to_open_and_self_closing_roots() {
        let open = r#"<Types xmlns="x"><A/></Types>"#;
        assert_eq!(append_child(open, "Types", "<B/>").unwrap(), r#"<Types xmlns="x"><A/><B/></Types>"#);
        let empty = r#"<?xml version="1.0"?><Types xmlns="x"/>"#;
        assert_eq!(append_child(empty, "Types", "<B/>").unwrap(), r#"<?xml version="1.0"?><Types xmlns="x"><B/></Types>"#);
        assert!(append_child("<Other/>", "Types", "<B/>").is_err());
    }

    #[test]
    fn escapes_markup() {
        assert_eq!(escape_xml(r#"a<b>&"c"'"#), "a&lt;b&gt;&amp;&quot;c&quot;&apos;");
    }

    #[test]
    fn matches_prefixed_names() {
        assert!(tag_is(b"w:p", "p"));
        assert!(tag_is("Relationship", "Relationship"));
        assert!(!tag_is(b"w:pPr", "p"));
    }

    #[test]
    fn removes_matching_elements_only() {
        let xml = r#"<R><A id="1"/><A id="2"><A id="3"/></A><B/></R>"#;
        let out = remove_elements(xml, "A", |e| get_attr_value(e, "id").as_deref() != Some("1")).unwrap();
        assert_eq!(out, r#"<R><A id="1"/><B/></R>"#);
        assert_eq!(element_ranges(xml, "B", |_| true).unwrap(), vec![39..43]);
    }

    #[test]
    fn finds_the_root_element() {
        assert_eq!(root_element(r#"<?xml version="1.0"?>
<!-- c --><x:sst a="1"/>"#), Some((32, "x:sst".to_string())));
        assert_eq!(prefix_of("x:sst"), "x:");
        assert_eq!(prefix_of("sst"), "");
    }

    #[test]
    fn unescapes_entities() {
        assert_eq!(unescape_text("a &amp; &#233; &lt;"), "a & é <");
        assert_eq!(unescape_text("a & b"), "a & b");
    }

    #[test]
    fn sets_and_removes_attributes_in_the_start_tag_only() {
        assert_eq!(set_attr(r#"<xf numFmtId="0" xfId="0"/>"#, "numFmtId", "14"), r#"<xf numFmtId="14" xfId="0"/>"#);
        assert_eq!(set_attr(r#"<xf fontId="0"/>"#, "Id", "1"), r#"<xf fontId="0" Id="1"/>"#);
        assert_eq!(set_attr(r#"<calcPr calcId='1'>x="2"</calcPr>"#, "x", "a&b"), r#"<calcPr calcId='1' x="a&amp;b">x="2"</calcPr>"#);
        assert_eq!(remove_attr(r#"<row r="1" spans="1:3" ht="2">"#, "spans"), r#"<row r="1" ht="2">"#);
        assert_eq!(remove_attr(r#"<row r="1"/>"#, "spans"), r#"<row r="1"/>"#);
    }
}
