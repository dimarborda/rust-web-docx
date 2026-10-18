//! `[Content_Types].xml`: the media type of every part, by extension or by part name

use crate::xml::{append_child, escape_xml, get_attr_value, remove_elements};

/// Name of the content types part
pub const PART: &str = "[Content_Types].xml";

/// Whether the content types declare a default for the extension `ext`
pub fn has_default(types: &str, ext: &str) -> bool {
    types.to_lowercase().contains(&format!("extension=\"{}\"", ext))
}

/// The content types with a `<Default>` for `ext`, or `None` when one is already declared
pub fn with_default(types: &str, ext: &str, mime: &str) -> Result<Option<String>, String> {
    if has_default(types, ext) {
        return Ok(None);
    }
    let entry = format!(r#"<Default Extension="{}" ContentType="{}"/>"#, ext, mime);
    append_child(types, "Types", &entry).map(Some)
}

/// Whether the content types give the part `part` (e.g. `xl/sharedStrings.xml`) its own type
pub fn has_override(types: &str, part: &str) -> bool {
    types.contains(&format!("PartName=\"/{}\"", part))
}

/// The content types with an `<Override>` for `part`, or `None` when one is already declared
pub fn with_override(types: &str, part: &str, mime: &str) -> Result<Option<String>, String> {
    if has_override(types, part) {
        return Ok(None);
    }
    let entry = format!(r#"<Override PartName="/{}" ContentType="{}"/>"#, escape_xml(part), escape_xml(mime));
    append_child(types, "Types", &entry).map(Some)
}

/// The content types without the `<Override>` of `part`
pub fn without_override(types: &str, part: &str) -> Result<String, String> {
    let name = format!("/{}", part);
    remove_elements(types, "Override", |e| get_attr_value(e, "PartName").is_some_and(|p| p.eq_ignore_ascii_case(&name)))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn adds_a_default_once() {
        let types = r#"<?xml version="1.0"?><Types xmlns="x"/>"#;
        let added = with_default(types, "png", "image/png").unwrap().unwrap();
        assert_eq!(added, r#"<?xml version="1.0"?><Types xmlns="x"><Default Extension="png" ContentType="image/png"/></Types>"#);
        assert!(has_default(&added, "png"));
        assert_eq!(with_default(&added, "png", "image/png").unwrap(), None);
    }

    #[test]
    fn adds_and_removes_overrides() {
        let types = r#"<Types xmlns="x"><Override PartName="/xl/calcChain.xml" ContentType="c"/></Types>"#;
        assert!(has_override(types, "xl/calcChain.xml"));
        let added = with_override(types, "xl/sharedStrings.xml", "s").unwrap().unwrap();
        assert!(with_override(&added, "xl/sharedStrings.xml", "s").unwrap().is_none());
        let removed = without_override(&added, "xl/calcChain.xml").unwrap();
        assert_eq!(removed, r#"<Types xmlns="x"><Override PartName="/xl/sharedStrings.xml" ContentType="s"/></Types>"#);
    }
}
