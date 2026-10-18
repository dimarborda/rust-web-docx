//! Relationship parts (`_rels/*.rels`)

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::xml::{append_child, escape_xml, get_attr_value, remove_elements, tag_is};

/// One relationship of a part
#[derive(Debug, Clone, PartialEq)]
pub struct Relationship {
    pub id: String,
    /// Relationship type URI (e.g. `.../relationships/worksheet`)
    pub rel_type: String,
    pub target: String,
    /// `TargetMode="External"`: the target is a URL, not a part of the package
    pub external: bool,
}

/// The relationships of a relationships part, in document order
pub fn parse_relationships(xml: &str) -> Vec<Relationship> {
    let mut rels = Vec::new();
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) if tag_is(e.name().as_ref(), "Relationship") => {
                if let (Some(id), Some(target)) = (get_attr_value(e, "Id"), get_attr_value(e, "Target")) {
                    rels.push(Relationship {
                        id,
                        rel_type: get_attr_value(e, "Type").unwrap_or_default(),
                        target,
                        external: get_attr_value(e, "TargetMode").is_some_and(|m| m.eq_ignore_ascii_case("External")),
                    });
                }
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    rels
}

/// Whether a relationship type URI ends with `/name` (`worksheet`, `sharedStrings`…), whatever
/// its namespace (transitional or strict)
pub fn type_is(rel_type: &str, name: &str) -> bool {
    rel_type.rsplit('/').next() == Some(name)
}

/// Name of the relationships part of `part`: `xl/workbook.xml` → `xl/_rels/workbook.xml.rels`
pub fn rels_part_for(part: &str) -> String {
    match part.rfind('/') {
        Some(i) => format!("{}/_rels/{}.rels", &part[..i], &part[i + 1..]),
        None => format!("_rels/{}.rels", part),
    }
}

/// Part name a relationship target points to, resolved against the part that holds the
/// relationship: (`xl/workbook.xml`, `worksheets/sheet1.xml`) → `xl/worksheets/sheet1.xml`;
/// absolute targets (`/xl/styles.xml`) are taken from the package root
pub fn resolve_target(source_part: &str, target: &str) -> String {
    let (base, rest) = match target.strip_prefix('/') {
        Some(absolute) => ("", absolute),
        None => (source_part.rfind('/').map(|i| &source_part[..i]).unwrap_or(""), target),
    };
    let mut segments: Vec<&str> = base.split('/').filter(|s| !s.is_empty()).collect();
    for segment in rest.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            s => segments.push(s),
        }
    }
    segments.join("/")
}

/// The relationships part with a new relationship appended, and the id it was given (`rIdN`
/// one past the highest numbered id)
pub fn with_relationship(xml: &str, rel_type: &str, target: &str) -> Result<(String, String), String> {
    let next = parse_relationships(xml)
        .iter()
        .filter_map(|r| r.id.strip_prefix("rId").and_then(|n| n.parse::<u32>().ok()))
        .max()
        .unwrap_or(0)
        + 1;
    let id = format!("rId{}", next);
    let entry = format!(
        r#"<Relationship Id="{}" Type="{}" Target="{}"/>"#,
        id,
        escape_xml(rel_type),
        escape_xml(target)
    );
    Ok((append_child(xml, "Relationships", &entry)?, id))
}

/// The relationships part without the relationship `id`
pub fn without_relationship(xml: &str, id: &str) -> Result<String, String> {
    remove_elements(xml, "Relationship", |e| get_attr_value(e, "Id").as_deref() == Some(id))
}

/// Targets of a relationships part by relationship id
pub fn parse_relationships_map(xml: &str) -> HashMap<String, String> {
    let mut map = HashMap::new();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                let name = e.name();
                if tag_is(name.as_ref(), "Relationship") {
                    let id_opt = get_attr_value(e, "Id");
                    let target_opt = get_attr_value(e, "Target");
                    if let (Some(id), Some(target)) = (id_opt, target_opt) {
                        map.insert(id, target);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    map
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn maps_ids_to_targets() {
        let xml = r#"<Relationships xmlns="x"><Relationship Id="rId1" Type="t" Target="media/a.png"/><Relationship Id="rId2" Type="t" Target="b.xml"></Relationship></Relationships>"#;
        let map = parse_relationships_map(xml);
        assert_eq!(map.len(), 2);
        assert_eq!(map["rId1"], "media/a.png");
        assert_eq!(map["rId2"], "b.xml");
    }

    #[test]
    fn reads_types_and_modes() {
        let xml = r#"<Relationships xmlns="x"><Relationship Id="rId1" Type="http://a/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId7" Type="http://a/relationships/hyperlink" Target="https://e.com" TargetMode="External"/></Relationships>"#;
        let rels = parse_relationships(xml);
        assert_eq!(rels.len(), 2);
        assert!(type_is(&rels[0].rel_type, "worksheet") && !rels[0].external);
        assert!(rels[1].external);

        let (added, id) = with_relationship(xml, "http://a/relationships/sharedStrings", "sharedStrings.xml").unwrap();
        assert_eq!(id, "rId8");
        assert_eq!(parse_relationships(&added)[2].target, "sharedStrings.xml");
        let removed = without_relationship(&added, "rId1").unwrap();
        assert_eq!(parse_relationships(&removed).iter().map(|r| r.id.as_str()).collect::<Vec<_>>(), ["rId7", "rId8"]);
    }

    #[test]
    fn resolves_targets_and_rels_parts() {
        assert_eq!(rels_part_for("xl/workbook.xml"), "xl/_rels/workbook.xml.rels");
        assert_eq!(rels_part_for("doc.xml"), "_rels/doc.xml.rels");
        assert_eq!(resolve_target("xl/workbook.xml", "worksheets/sheet1.xml"), "xl/worksheets/sheet1.xml");
        assert_eq!(resolve_target("xl/worksheets/sheet1.xml", "../drawings/d1.xml"), "xl/drawings/d1.xml");
        assert_eq!(resolve_target("xl/workbook.xml", "/xl/styles.xml"), "xl/styles.xml");
        assert_eq!(resolve_target("", "xl/workbook.xml"), "xl/workbook.xml");
    }
}
