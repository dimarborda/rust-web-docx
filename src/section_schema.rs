//! The section schema of a document type (e.g. a technical report must have Introducción,
//! Caracterización and Conclusiones, in that order), stored in the .docx as a custom XML part
//! so it travels with the file and other tools can read it:
//!
//! ```xml
//! <sections xmlns="urn:rust-web-docx:sections:1" type="informe-tecnico" title="Informe técnico">
//!   <section key="introduccion" title="Introducción" required="true"/>
//!   <section key="alcance" title="Alcance" required="false" parent="introduccion"/>
//! </sections>
//! ```
//!
//! The part is `customXml/itemN.xml`, with its `itemPropsN.xml`, related from
//! `word/document.xml` as Word does for its own custom XML. Word keeps it when saving.
//! Child module of `docx_parser` so it can work on the package parts directly.

use super::section_edit::check_section_key;
use super::*;
use ooxml_core::content_types;
use ooxml_core::rels::{parse_relationships, rels_part_for, resolve_target, with_relationship, without_relationship};

/// Namespace of the schema part
pub const SECTION_SCHEMA_NS: &str = "urn:rust-web-docx:sections:1";

const REL_CUSTOM_XML: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const REL_CUSTOM_XML_PROPS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps";
const MIME_CUSTOM_XML_PROPS: &str = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
const DOCUMENT: &str = "word/document.xml";

/// The sections a type of document must have, in order
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
pub struct SectionSchema {
    /// Document type, e.g. "informe-tecnico"
    #[serde(rename = "type")]
    pub kind: String,
    #[serde(default)]
    pub title: Option<String>,
    pub sections: Vec<SchemaSection>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
pub struct SchemaSection {
    pub key: String,
    #[serde(default)]
    pub title: Option<String>,
    /// Defaults to true
    #[serde(default = "required_default")]
    pub required: bool,
    /// Key of the section it must be nested in
    #[serde(default)]
    pub parent: Option<String>,
}

fn required_default() -> bool {
    true
}

impl SectionSchema {
    /// Refuses empty types, invalid or repeated keys and parents that are not listed before
    pub fn check(&self) -> Result<(), String> {
        if self.kind.trim().is_empty() {
            return Err("El esquema necesita un tipo de documento (type), p. ej. «informe-tecnico».".to_string());
        }
        let mut seen = std::collections::HashSet::new();
        for s in &self.sections {
            check_section_key(&s.key)?;
            if let Some(parent) = &s.parent {
                if !seen.contains(parent.as_str()) {
                    return Err(format!("La sección «{}» del esquema va dentro de «{}», que debe aparecer antes en la lista.", s.key, parent));
                }
            }
            if !seen.insert(s.key.as_str()) {
                return Err(format!("La clave «{}» está repetida en el esquema.", s.key));
            }
        }
        Ok(())
    }

    fn to_xml(&self) -> String {
        let attr = |name: &str, value: &Option<String>| match value {
            Some(v) => format!(r#" {}="{}""#, name, escape_xml(v)),
            None => String::new(),
        };
        let mut xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sections xmlns="{}" type="{}"{}>"#,
            SECTION_SCHEMA_NS,
            escape_xml(&self.kind),
            attr("title", &self.title)
        );
        for s in &self.sections {
            xml.push_str(&format!(
                r#"<section key="{}"{} required="{}"{}/>"#,
                escape_xml(&s.key),
                attr("title", &s.title),
                s.required,
                attr("parent", &s.parent)
            ));
        }
        xml.push_str("</sections>");
        xml
    }

    fn from_xml(xml: &str) -> Option<SectionSchema> {
        let mut reader = Reader::from_str(xml);
        let mut schema: Option<SectionSchema> = None;
        loop {
            match reader.read_event() {
                Ok(Event::Start(ref e)) | Ok(Event::Empty(ref e)) => {
                    let name = e.name();
                    if tag_is(name.as_ref(), "sections") {
                        schema = Some(SectionSchema {
                            kind: get_attr_value(e, "type").unwrap_or_default(),
                            title: get_attr_value(e, "title"),
                            sections: Vec::new(),
                        });
                    } else if tag_is(name.as_ref(), "section") {
                        if let (Some(schema), Some(key)) = (schema.as_mut(), get_attr_value(e, "key")) {
                            schema.sections.push(SchemaSection {
                                key,
                                title: get_attr_value(e, "title"),
                                required: get_attr_value(e, "required").as_deref() != Some("false"),
                                parent: get_attr_value(e, "parent"),
                            });
                        }
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                _ => {}
            }
        }
        schema
    }
}

/// A GUID-shaped id for `ds:itemID`, derived from the content (no randomness in WebAssembly)
fn item_guid(seed: &str) -> String {
    let hash = |salt: u64| {
        let mut h = std::collections::hash_map::DefaultHasher::new();
        (salt, seed).hash(&mut h);
        h.finish()
    };
    let (a, b) = (hash(1), hash(2));
    format!(
        "{{{:08X}-{:04X}-4{:03X}-{:04X}-{:012X}}}",
        a >> 32,
        (a >> 16) & 0xFFFF,
        a & 0x0FFF,
        0x8000 | (b >> 48) & 0x3FFF,
        b & 0xFFFF_FFFF_FFFF
    )
}

impl DocxModifier {
    /// Name of the custom XML part holding the schema, if any
    fn schema_part(&self) -> Option<String> {
        let mut names: Vec<&String> = self
            .package
            .parts()
            .keys()
            .filter(|n| n.starts_with("customXml/item") && n.ends_with(".xml") && !n.contains("Props"))
            .collect();
        names.sort();
        names
            .into_iter()
            .find(|n| self.get_file_string(n).is_ok_and(|xml| xml.contains(SECTION_SCHEMA_NS)))
            .cloned()
    }

    /// The document's section schema, or `None`
    pub fn section_schema(&self) -> Result<Option<SectionSchema>, String> {
        Ok(self.schema_part().and_then(|part| SectionSchema::from_xml(&self.get_file_string(&part).ok()?)))
    }

    /// Stores the section schema in the document (replacing the current one), or removes it
    /// with `None`
    pub fn set_section_schema(&mut self, schema: Option<&SectionSchema>) -> Result<(), String> {
        if let Some(schema) = schema {
            schema.check()?;
        }
        match (self.schema_part(), schema) {
            (Some(part), Some(schema)) => self.put_file(part, schema.to_xml().into_bytes()),
            (None, Some(schema)) => self.add_schema_part(schema)?,
            (Some(part), None) => self.remove_schema_part(&part)?,
            (None, None) => {}
        }
        Ok(())
    }

    fn add_schema_part(&mut self, schema: &SectionSchema) -> Result<(), String> {
        let n = (1..).find(|n| !self.package.contains(&format!("customXml/item{}.xml", n))).unwrap();
        let item = format!("customXml/item{}.xml", n);
        let props = format!("customXml/itemProps{}.xml", n);
        let item_xml = schema.to_xml();
        let props_xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="no"?><ds:datastoreItem ds:itemID="{}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="{}"/></ds:schemaRefs></ds:datastoreItem>"#,
            item_guid(&format!("{}{}", n, item_xml)),
            SECTION_SCHEMA_NS
        );
        let item_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{}" Target="itemProps{}.xml"/></Relationships>"#,
            REL_CUSTOM_XML_PROPS, n
        );
        let doc_rels_part = rels_part_for(DOCUMENT);
        let doc_rels = self.get_file_string(&doc_rels_part).unwrap_or_else(|_| {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.to_string()
        });
        let (doc_rels, _) = with_relationship(&doc_rels, REL_CUSTOM_XML, &format!("../{}", item))?;
        let mut types = self.get_file_string(content_types::PART)?;
        if let Some(t) = content_types::with_default(&types, "xml", "application/xml")? {
            types = t;
        }
        if let Some(t) = content_types::with_override(&types, &props, MIME_CUSTOM_XML_PROPS)? {
            types = t;
        }
        let item_rels_part = rels_part_for(&item);
        for (name, xml) in [(item.clone(), item_xml), (props, props_xml), (item_rels_part, item_rels), (doc_rels_part, doc_rels)] {
            self.package.keep_in_order(&name);
            self.put_file(name, xml.into_bytes());
        }
        self.put_file(content_types::PART.to_string(), types.into_bytes());
        Ok(())
    }

    fn remove_schema_part(&mut self, item: &str) -> Result<(), String> {
        let item_rels_part = rels_part_for(item);
        let props: Vec<String> = self
            .get_file_string(&item_rels_part)
            .map(|xml| parse_relationships(&xml).into_iter().map(|r| resolve_target(item, &r.target)).collect())
            .unwrap_or_default();
        let mut types = self.get_file_string(content_types::PART)?;
        for part in &props {
            types = content_types::without_override(&types, part)?;
            self.remove_file(part);
        }
        self.remove_file(&item_rels_part);
        self.remove_file(item);
        let doc_rels_part = rels_part_for(DOCUMENT);
        if let Ok(mut doc_rels) = self.get_file_string(&doc_rels_part) {
            for rel in parse_relationships(&doc_rels.clone()) {
                if resolve_target(DOCUMENT, &rel.target) == item {
                    doc_rels = without_relationship(&doc_rels, &rel.id)?;
                }
            }
            self.put_file(doc_rels_part, doc_rels.into_bytes());
        }
        self.put_file(content_types::PART.to_string(), types.into_bytes());
        Ok(())
    }
}
