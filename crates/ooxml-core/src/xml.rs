//! Small XML helpers that work on any OOXML part

use quick_xml::events::BytesStart;

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
}
