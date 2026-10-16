//! Logical sections of a document (Introducción, Caracterización, Conclusiones...): blocks
//! wrapped in a block-level content control (`w:sdt`) tagged `section:<key>`, so the editor
//! and any other tool reading the .docx know where each one starts and ends. They are not
//! Word's page-layout sections (`w:sectPr`) and change nothing in how the document looks.
//!
//! ```xml
//! <w:sdt>
//!   <w:sdtPr><w:alias w:val="Introducción"/><w:tag w:val="section:introduccion"/><w:id w:val="…"/></w:sdtPr>
//!   <w:sdtContent> …paragraphs and tables… </w:sdtContent>
//! </w:sdt>
//! ```
//!
//! Sections can be nested. Child module of `docx_parser` so it can work on the package parts
//! directly.

use super::*;
use crate::paragraph_edit::{element_end, is_element, tokenize, Token};
use std::ops::Range;

/// Prefix of `w:tag` that marks a content control as a section
pub const SECTION_TAG_PREFIX: &str = "section:";

const MAX_KEY_CHARS: usize = 64;

/// A logical section, from `DocxModifier::sections`
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
pub struct SectionInfo {
    /// What the section API takes (the tag without `section:`)
    pub key: String,
    /// Human-readable name (`w:alias`; the key when the control has none)
    pub title: String,
    /// First and last paragraph inside, table cells included (same numbering as `structure()`);
    /// `None` for a section with no paragraphs
    pub first: Option<usize>,
    pub last: Option<usize>,
    /// 0 for a top-level section, 1 inside another one, …
    pub depth: usize,
    /// Key of the enclosing section
    pub parent: Option<String>,
    /// Word users cannot delete the control (`w:lock` sdtLocked or sdtContentLocked)
    pub locked: bool,
}

struct SectionLoc {
    key: String,
    title: Option<String>,
    locked: bool,
    /// The whole `w:sdt` element
    element: Range<usize>,
    /// What lies between `<w:sdtContent>` and `</w:sdtContent>` (`None`: no or empty content)
    content: Option<Range<usize>>,
}

/// Body-level content controls tagged as sections, in document order (outer before inner).
/// Controls inside paragraphs or tables are not sections.
fn locate_sections(tokens: &[Token]) -> Vec<SectionLoc> {
    let mut out = Vec::new();
    let mut in_body = false;
    let mut i = 0;
    while i < tokens.len() {
        let tok = &tokens[i];
        match &tok.ev {
            Event::Start(e) if tag_is(e.name().as_ref(), "body") => in_body = true,
            Event::End(e) if tag_is(e.name().as_ref(), "body") => in_body = false,
            Event::Start(_) if in_body && (is_element(tok, "p") || is_element(tok, "tbl")) => {
                i = element_end(tokens, i) + 1;
                continue;
            }
            Event::Start(_) if in_body && is_element(tok, "sdt") => {
                if let Some(loc) = read_section(tokens, i, element_end(tokens, i)) {
                    out.push(loc);
                }
                // Walk on into the content: sections can be nested
            }
            _ => {}
        }
        i += 1;
    }
    out
}

/// The section of the `w:sdt` spanning tokens `start..=end`, if its tag marks one
fn read_section(tokens: &[Token], start: usize, end: usize) -> Option<SectionLoc> {
    let mut key = None;
    let mut title = None;
    let mut locked = false;
    let mut content = None;
    let mut i = start + 1;
    while i < end {
        let tok = &tokens[i];
        if is_element(tok, "sdtPr") {
            let pr_end = element_end(tokens, i);
            for t in &tokens[i + 1..pr_end] {
                let (Event::Start(e) | Event::Empty(e)) = &t.ev else { continue };
                let name = e.name();
                if tag_is(name.as_ref(), "tag") {
                    key = get_attr_value(e, "val").and_then(|v| v.strip_prefix(SECTION_TAG_PREFIX).map(str::to_string));
                } else if tag_is(name.as_ref(), "alias") {
                    title = get_attr_value(e, "val").filter(|v| !v.trim().is_empty());
                } else if tag_is(name.as_ref(), "lock") {
                    locked = matches!(get_attr_value(e, "val").as_deref(), Some("sdtLocked" | "sdtContentLocked"));
                }
            }
            i = pr_end + 1;
            continue;
        }
        if is_element(tok, "sdtContent") {
            let c_end = element_end(tokens, i);
            if c_end > i {
                content = Some(tokens[i].span.end..tokens[c_end].span.start);
            }
            i = c_end + 1;
            continue;
        }
        if matches!(tok.ev, Event::Start(_)) {
            i = element_end(tokens, i) + 1;
            continue;
        }
        i += 1;
    }
    let key = key.filter(|k| !k.is_empty())?;
    Some(SectionLoc { key, title, locked, element: tokens[start].span.start..tokens[end].span.end, content })
}

/// Refuses keys other tools could not use as identifiers: letters, digits, `-`, `_` and `.`
fn check_key(key: &str) -> Result<(), String> {
    let valid = !key.is_empty()
        && key.chars().count() <= MAX_KEY_CHARS
        && key.chars().next().is_some_and(char::is_alphanumeric)
        && key.chars().all(|c| c.is_alphanumeric() || matches!(c, '-' | '_' | '.'));
    if valid {
        Ok(())
    } else {
        Err(format!(
            "Clave de sección no válida: «{}». Usa hasta {} letras, dígitos, «-», «_» o «.», empezando por letra o dígito (p. ej. «introduccion»).",
            key, MAX_KEY_CHARS
        ))
    }
}

/// Byte ranges of the top-level body tables (tables nested in cells are inside them)
fn body_table_spans(tokens: &[Token]) -> Vec<Range<usize>> {
    let mut out = Vec::new();
    let mut in_body = false;
    let mut i = 0;
    while i < tokens.len() {
        let tok = &tokens[i];
        match &tok.ev {
            Event::Start(e) if tag_is(e.name().as_ref(), "body") => in_body = true,
            Event::Start(_) if in_body && is_element(tok, "p") => {
                i = element_end(tokens, i) + 1;
                continue;
            }
            Event::Start(_) if in_body && is_element(tok, "tbl") => {
                let end = element_end(tokens, i);
                out.push(tokens[i].span.start..tokens[end].span.end);
                i = end + 1;
                continue;
            }
            _ => {}
        }
        i += 1;
    }
    out
}

/// Holds a paragraph-level section break (`w:sectPr` in `w:pPr`); Word does not allow one
/// inside a content control
fn has_section_break(xml: &str) -> bool {
    xml.contains("<w:sectPr")
}

impl DocxModifier {
    /// Logical sections of the body in document order (outer before inner)
    pub fn sections(&self) -> Result<Vec<SectionInfo>, String> {
        let xml = self.get_file_string("word/document.xml")?;
        let tokens = tokenize(&xml)?;
        let paragraphs = body_paragraph_ranges(&xml)?;
        let locs = locate_sections(&tokens);
        let mut out = Vec::with_capacity(locs.len());
        let mut open: Vec<(usize, Range<usize>)> = Vec::new();
        for (n, loc) in locs.iter().enumerate() {
            while open.last().is_some_and(|(_, r)| loc.element.start >= r.end) {
                open.pop();
            }
            let inside: Vec<usize> = match &loc.content {
                Some(c) => paragraphs
                    .iter()
                    .enumerate()
                    .filter(|(_, r)| r.start >= c.start && r.end <= c.end)
                    .map(|(i, _)| i)
                    .collect(),
                None => Vec::new(),
            };
            out.push(SectionInfo {
                key: loc.key.clone(),
                title: loc.title.clone().unwrap_or_else(|| loc.key.clone()),
                first: inside.first().copied(),
                last: inside.last().copied(),
                depth: open.len(),
                parent: open.last().map(|(p, _)| locs[*p].key.clone()),
                locked: loc.locked,
            });
            open.push((n, loc.element.clone()));
        }
        Ok(out)
    }

    /// The only section with `key` in `xml`
    fn find_section(xml: &str, key: &str) -> Result<SectionLoc, String> {
        let tokens = tokenize(xml)?;
        let mut locs = locate_sections(&tokens);
        let count = locs.iter().filter(|l| l.key == key).count();
        match count {
            1 => Ok(locs.swap_remove(locs.iter().position(|l| l.key == key).unwrap())),
            0 => {
                let keys: Vec<String> = locs.iter().map(|l| format!("«{}»", l.key)).collect();
                Err(if keys.is_empty() {
                    format!("No existe la sección «{}»: el documento no tiene secciones.", key)
                } else {
                    format!("No existe la sección «{}». Secciones del documento: {}.", key, keys.join(", "))
                })
            }
            n => Err(format!(
                "Hay {} secciones con la clave «{}» (p. ej. por copiar y pegar en Word); quita las sobrantes con unwrapSection o cámbiales la clave.",
                n, key
            )),
        }
    }

    /// Wraps the blocks from paragraph `from` to `to` (both included, with the tables between
    /// them) in a new section. A table is wrapped whole when `from` is its first paragraph or
    /// `to` its last one. `locked`: Word users cannot delete the section (its content stays
    /// editable). Refuses keys already in use and ranges that cut a table or another section,
    /// lie inside a table, or hold a section break.
    pub fn define_section(
        &mut self,
        key: &str,
        from: usize,
        to: usize,
        title: Option<&str>,
        locked: bool,
    ) -> Result<SectionInfo, String> {
        check_key(key)?;
        let mut xml = self.get_file_string("word/document.xml")?;
        let tokens = tokenize(&xml)?;
        if locate_sections(&tokens).iter().any(|l| l.key == key) {
            return Err(format!("Ya existe una sección con la clave «{}».", key));
        }
        let (from, to) = if to < from { (to, from) } else { (from, to) };
        let paragraphs = body_paragraph_ranges(&xml)?;
        let (mut start, mut end) = match (paragraphs.get(from), paragraphs.get(to)) {
            (Some(a), Some(b)) => (a.start, b.end),
            _ => return Err(format!("El rango {}–{} no existe en el documento.", from, to)),
        };
        // A range starting at a table's first paragraph (or ending at its last) takes it whole
        for table in body_table_spans(&tokens) {
            let cells: Vec<usize> = (0..paragraphs.len())
                .filter(|&i| paragraphs[i].start >= table.start && paragraphs[i].end <= table.end)
                .collect();
            let within = |pos: usize| pos > table.start && pos < table.end;
            if within(start) {
                if cells.first() != Some(&from) {
                    return Err(format!(
                        "El párrafo {} está dentro de una tabla: la sección debe empezar en el primer párrafo de la tabla ({}) o fuera de ella.",
                        from,
                        cells.first().copied().unwrap_or(from)
                    ));
                }
                start = table.start;
            }
            if within(end) {
                if cells.last() != Some(&to) {
                    return Err(format!(
                        "El párrafo {} está dentro de una tabla: la sección debe terminar en el último párrafo de la tabla ({}) o fuera de ella.",
                        to,
                        cells.last().copied().unwrap_or(to)
                    ));
                }
                end = table.end;
            }
        }
        if !is_balanced(&xml[start..end]) {
            return Err("El rango corta otra sección o un control de contenido; amplíalo para incluirlos completos o redúcelo para quedar dentro.".to_string());
        }
        if has_section_break(&xml[start..end]) {
            return Err("El rango incluye un salto de sección de Word; una sección lógica no puede contenerlo.".to_string());
        }
        let id = Self::new_sdt_id(&xml, key);
        let title = title.map(str::trim).filter(|t| !t.is_empty()).unwrap_or(key);
        let lock = if locked { r#"<w:lock w:val="sdtLocked"/>"# } else { "" };
        let wrapped = format!(
            r#"<w:sdt><w:sdtPr><w:alias w:val="{}"/><w:tag w:val="{}{}"/><w:id w:val="{}"/>{}</w:sdtPr><w:sdtContent>{}</w:sdtContent></w:sdt>"#,
            escape_xml(title),
            SECTION_TAG_PREFIX,
            escape_xml(key),
            id,
            lock,
            &xml[start..end]
        );
        xml.replace_range(start..end, &wrapped);
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        self.sections()?
            .into_iter()
            .find(|s| s.key == key)
            .ok_or_else(|| "No se pudo crear la sección.".to_string())
    }

    /// A content-control id (`w:id`, a signed 32-bit number) not used elsewhere in `xml`
    fn new_sdt_id(xml: &str, key: &str) -> i32 {
        let mut hasher = std::collections::hash_map::DefaultHasher::new();
        key.hash(&mut hasher);
        let mut id = (hasher.finish() % 0x3FFF_FFFF) as i32 + 1;
        while xml.contains(&format!("w:id w:val=\"{}\"", id)) {
            id = id % 0x3FFF_FFFF + 1;
        }
        id
    }

    /// Removes the section marker, leaving its content where it was (sections nested in it
    /// stay)
    pub fn unwrap_section(&mut self, key: &str) -> Result<(), String> {
        let mut xml = self.get_file_string("word/document.xml")?;
        let loc = Self::find_section(&xml, key)?;
        let content = loc.content.map(|c| xml[c].to_string()).unwrap_or_default();
        xml.replace_range(loc.element, &content);
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        Ok(())
    }

    /// Replaces the content of a section by `items`, keeping the section (its key, title and
    /// lock). `keep_heading`: a heading paragraph opening the section stays and the items go
    /// after it. Sections nested in the replaced content go away with it. Returns the first
    /// new paragraph and the caret after the last one.
    pub fn replace_section(&mut self, key: &str, items: &[NewParagraph], keep_heading: bool) -> Result<(usize, (usize, usize)), String> {
        if items.is_empty() {
            return Err("No hay párrafos para insertar.".to_string());
        }
        let mut xml = self.get_file_string("word/document.xml")?;
        let loc = Self::find_section(&xml, key)?;
        let content = loc
            .content
            .ok_or_else(|| format!("La sección «{}» no tiene contenido que reemplazar.", key))?;
        let paragraphs = body_paragraph_ranges(&xml)?;
        let mut start = content.start;
        if keep_heading {
            // Only a paragraph that is the section's first block (not in a table or a nested
            // section) counts as its heading
            let opening = paragraphs
                .iter()
                .enumerate()
                .find(|(_, r)| r.start >= content.start && r.end <= content.end)
                .filter(|(_, r)| is_blank(&xml[content.start..r.start]));
            if let Some((index, range)) = opening {
                let is_heading = self
                    .extract_paragraphs()?
                    .iter()
                    .any(|p| p.index == index && (p.is_heading || p.outline_level.is_some()));
                if is_heading {
                    start = range.end;
                }
            }
        }
        let end = content.end;
        if has_section_break(&xml[start..end]) {
            return Err(format!("La sección «{}» contiene un salto de sección de Word; no se puede reemplazar de una vez.", key));
        }
        let index = paragraphs.iter().filter(|r| r.end <= start).count();
        xml.replace_range(start..end, "<w:p/>");
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        self.insert_paragraphs(index, 0, items)
    }
}
