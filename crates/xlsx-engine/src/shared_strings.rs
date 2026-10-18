//! `xl/sharedStrings.xml`: the table of texts that cells with `t="s"` point to by index

use std::collections::HashMap;

use ooxml_core::xml::{append_child, element_ranges, get_attr_value, prefix_of, root_element, set_attr};
use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::text::{decode, t_element, text_spans};

/// Root of a new shared strings part
pub const EMPTY_SST: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>"#;

/// The texts of the table, by index, and the plain ones (a single run with no formatting) by
/// text, which can be reused for new cells
#[derive(Debug, Default)]
pub struct SharedStrings {
    pub items: Vec<String>,
    plain: HashMap<String, u32>,
    /// Strings appended by `intern` and not yet written
    added: Vec<String>,
    /// New cell references to the table, for its `count`
    new_refs: u64,
}

impl SharedStrings {
    pub fn parse(xml: &str) -> Result<SharedStrings, String> {
        let mut table = SharedStrings::default();
        for range in element_ranges(xml, "si", |_| true).map_err(|e| format!("Las cadenas compartidas no son {}", e))? {
            let item = &xml[range];
            let spans = text_spans(item, "si")?;
            let text: String = spans.iter().map(|s| decode(&item[s.content.clone()])).collect();
            if spans.len() == 1 && element_ranges(item, "r", |_| true)?.is_empty() {
                table.plain.entry(text.clone()).or_insert(table.items.len() as u32);
            }
            table.items.push(text);
        }
        Ok(table)
    }

    pub fn get(&self, index: usize) -> Option<&str> {
        self.items.get(index).or_else(|| self.added.get(index.wrapping_sub(self.items.len()))).map(String::as_str)
    }

    /// Index of `text` in the table, adding it when no plain entry holds it yet
    pub fn intern(&mut self, text: &str) -> u32 {
        self.new_refs += 1;
        if let Some(index) = self.plain.get(text) {
            return *index;
        }
        let index = (self.items.len() + self.added.len()) as u32;
        self.added.push(text.to_string());
        self.plain.insert(text.to_string(), index);
        index
    }

    pub fn changed(&self) -> bool {
        !self.added.is_empty() || self.new_refs > 0
    }

    /// The part with the added strings appended and its counters updated; the existing entries
    /// stay as written
    pub fn write(&self, xml: &str) -> Result<String, String> {
        let (root_at, root) = root_element(xml).ok_or("Las cadenas compartidas no tienen elemento raíz.")?;
        let prefix = prefix_of(&root).to_string();
        let mut out = xml.to_string();
        if !self.added.is_empty() {
            let items: String = self.added.iter().map(|t| format!("<{p}si>{}</{p}si>", t_element(&prefix, t), p = prefix)).collect();
            out = append_child(&out, &root, &items)?;
        }
        let tag_attr = |name: &str| -> Option<u64> {
            let mut reader = Reader::from_str(&xml[root_at..]);
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => get_attr_value(&e, name).and_then(|v| v.parse().ok()),
                _ => None,
            }
        };
        let mut root_part = out[root_at..].to_string();
        if let Some(count) = tag_attr("count") {
            root_part = set_attr(&root_part, "count", &(count + self.new_refs).to_string());
        }
        if tag_attr("uniqueCount").is_some() {
            root_part = set_attr(&root_part, "uniqueCount", &(self.items.len() + self.added.len()).to_string());
        }
        Ok(format!("{}{}", &out[..root_at], root_part))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const SST: &str = r#"<sst xmlns="m" count="3" uniqueCount="3"><si><t>Hola</t></si><si><r><rPr><b/></rPr><t>Ho</t></r><r><t>la</t></r></si><si><t xml:space="preserve"> a &amp; b</t><rPh sb="0" eb="1"><t>x</t></rPh></si></sst>"#;

    #[test]
    fn reads_plain_and_rich_strings() {
        let table = SharedStrings::parse(SST).unwrap();
        assert_eq!(table.items, ["Hola", "Hola", " a & b"]);
        assert_eq!(table.plain.get("Hola"), Some(&0));
        assert_eq!(table.plain.get(" a & b"), Some(&2));
    }

    #[test]
    fn reuses_plain_strings_and_appends_new_ones() {
        let mut table = SharedStrings::parse(SST).unwrap();
        assert_eq!(table.intern("Hola"), 0);
        assert_eq!(table.intern("Nuevo <1>"), 3);
        assert_eq!(table.intern("Nuevo <1>"), 3);
        assert_eq!(table.get(3), Some("Nuevo <1>"));
        let xml = table.write(SST).unwrap();
        assert!(xml.starts_with(r#"<sst xmlns="m" count="6" uniqueCount="4">"#), "{}", xml);
        assert!(xml.ends_with("<si><t>Nuevo &lt;1&gt;</t></si></sst>"), "{}", xml);
        assert_eq!(SharedStrings::parse(&xml).unwrap().items[3], "Nuevo <1>");

        let mut empty = SharedStrings::parse(EMPTY_SST).unwrap();
        assert_eq!(empty.intern(" x"), 0);
        let xml = empty.write(EMPTY_SST).unwrap();
        assert!(xml.contains(r#"count="1" uniqueCount="1"><si><t xml:space="preserve"> x</t></si></sst>"#), "{}", xml);
    }
}
