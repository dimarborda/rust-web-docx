//! `[Content_Types].xml`: the media type of every part, by extension or by part name

use crate::xml::append_child;

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
}
