//! Relationship parts (`_rels/*.rels`)

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use crate::xml::{get_attr_value, tag_is};

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
}
