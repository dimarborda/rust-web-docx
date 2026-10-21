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
//! Sections can be nested, locked against deletion (`w:lock` sdtLocked) and against editing
//! their content (contentLocked), created from the document's headings and inserted new.
//! Child module of `docx_parser` so it can work on the package parts directly.

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
    /// The section cannot be deleted (`w:lock` sdtLocked or sdtContentLocked)
    pub locked: bool,
    /// Its content cannot be edited by hand (`w:lock` contentLocked or sdtContentLocked); code
    /// can still fill it
    #[serde(rename = "contentLocked")]
    pub content_locked: bool,
}

/// Where `DocxModifier::insert_section` puts a new section
#[derive(Debug, Clone, PartialEq)]
pub enum SectionPlace {
    /// Right after section `key` (inside the same parent)
    After(String),
    /// Right before section `key`
    Before(String),
    /// At the end of section `key`'s content (nested in it)
    Inside(String),
    /// At the end of the document
    End,
}

/// What `DocxModifier::sections_from_headings` did
#[derive(Serialize, Debug, Clone, PartialEq)]
pub struct HeadingSections {
    /// Sections created, in document order
    pub created: Vec<SectionInfo>,
    /// Headings left without a section, and why
    pub skipped: Vec<SkippedHeading>,
}

#[derive(Serialize, Debug, Clone, PartialEq)]
pub struct SkippedHeading {
    pub paragraph: usize,
    pub text: String,
    pub reason: String,
}

struct SectionLoc {
    key: String,
    title: Option<String>,
    locked: bool,
    content_locked: bool,
    /// The whole `w:sdt` element
    element: Range<usize>,
    /// The `w:sdtPr` element
    properties: Range<usize>,
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
    let mut lock = None;
    let mut properties = None;
    let mut content = None;
    let mut i = start + 1;
    while i < end {
        let tok = &tokens[i];
        if is_element(tok, "sdtPr") {
            let pr_end = element_end(tokens, i);
            properties = Some(tokens[i].span.start..tokens[pr_end].span.end);
            for t in &tokens[i + 1..pr_end] {
                let (Event::Start(e) | Event::Empty(e)) = &t.ev else { continue };
                let name = e.name();
                if tag_is(name.as_ref(), "tag") {
                    key = get_attr_value(e, "val").and_then(|v| v.strip_prefix(SECTION_TAG_PREFIX).map(str::to_string));
                } else if tag_is(name.as_ref(), "alias") {
                    title = get_attr_value(e, "val").filter(|v| !v.trim().is_empty());
                } else if tag_is(name.as_ref(), "lock") {
                    lock = get_attr_value(e, "val");
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
    let lock = lock.as_deref();
    Some(SectionLoc {
        key,
        title,
        locked: matches!(lock, Some("sdtLocked" | "sdtContentLocked")),
        content_locked: matches!(lock, Some("contentLocked" | "sdtContentLocked")),
        element: tokens[start].span.start..tokens[end].span.end,
        properties: properties?,
        content,
    })
}

/// Refuses keys other tools could not use as identifiers: letters, digits, `-`, `_` and `.`
pub(super) fn check_section_key(key: &str) -> Result<(), String> {
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

/// A section key from a heading: without its numbering ("2.1 "), in lowercase ASCII where
/// possible and words joined by `-` ("2.1 Caracterización del sitio" → "caracterizacion-del-sitio")
pub fn section_key_from_title(title: &str) -> String {
    let unnumbered = title
        .trim_start()
        .trim_start_matches(|c: char| c.is_ascii_digit() || c == '.' || c == ')')
        .trim();
    let mut key = String::new();
    for c in unnumbered.chars().flat_map(char::to_lowercase) {
        let c = match c {
            'á' | 'à' | 'â' | 'ä' | 'ã' | 'å' => 'a',
            'é' | 'è' | 'ê' | 'ë' => 'e',
            'í' | 'ì' | 'î' | 'ï' => 'i',
            'ó' | 'ò' | 'ô' | 'ö' | 'õ' => 'o',
            'ú' | 'ù' | 'û' | 'ü' => 'u',
            'ñ' => 'n',
            'ç' => 'c',
            c => c,
        };
        if c.is_alphanumeric() {
            key.push(c);
        } else if !key.is_empty() && !key.ends_with('-') {
            key.push('-');
        }
    }
    let key: String = key.trim_end_matches('-').chars().take(MAX_KEY_CHARS).collect();
    let key = key.trim_end_matches('-').to_string();
    if key.is_empty() {
        "seccion".to_string()
    } else {
        key
    }
}

/// `base`, or `base-2`, `base-3`… when taken
fn unique_key(base: &str, taken: &std::collections::HashSet<String>) -> String {
    if !taken.contains(base) {
        return base.to_string();
    }
    (2..)
        .map(|n| {
            let suffix = format!("-{}", n);
            let head: String = base.chars().take(MAX_KEY_CHARS - suffix.len()).collect();
            format!("{}{}", head, suffix)
        })
        .find(|k| !taken.contains(k))
        .unwrap()
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

/// Bytes of the blocks from paragraph `from` to `to` (both included), ready to be wrapped in a
/// section: a table is taken whole when `from` is its first paragraph or `to` its last one.
/// Refuses ranges that cut a table or a content control, or hold a section break.
fn block_span(xml: &str, paragraphs: &[Range<usize>], tables: &[Range<usize>], from: usize, to: usize) -> Result<Range<usize>, String> {
    let (mut start, mut end) = match (paragraphs.get(from), paragraphs.get(to)) {
        (Some(a), Some(b)) => (a.start, b.end),
        _ => return Err(format!("El rango {}–{} no existe en el documento.", from, to)),
    };
    for table in tables {
        let within = |pos: usize| pos > table.start && pos < table.end;
        if !within(start) && !within(end) {
            continue;
        }
        let cells: Vec<usize> = (0..paragraphs.len())
            .filter(|&i| paragraphs[i].start >= table.start && paragraphs[i].end <= table.end)
            .collect();
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
    Ok(start..end)
}

/// `w:lock` for the two locks (none when both are off)
fn lock_element(locked: bool, content_locked: bool) -> &'static str {
    match (locked, content_locked) {
        (true, true) => r#"<w:lock w:val="sdtContentLocked"/>"#,
        (true, false) => r#"<w:lock w:val="sdtLocked"/>"#,
        (false, true) => r#"<w:lock w:val="contentLocked"/>"#,
        (false, false) => "",
    }
}

/// Opening of a section's content control, up to `<w:sdtContent>`
fn section_open(key: &str, title: &str, id: i32, locked: bool, content_locked: bool) -> String {
    format!(
        r#"<w:sdt><w:sdtPr><w:alias w:val="{}"/><w:tag w:val="{}{}"/><w:id w:val="{}"/>{}</w:sdtPr><w:sdtContent>"#,
        escape_xml(title),
        SECTION_TAG_PREFIX,
        escape_xml(key),
        id,
        lock_element(locked, content_locked)
    )
}

const SECTION_CLOSE: &str = "</w:sdtContent></w:sdt>";

/// Content-control ids (`w:id`, signed 32-bit numbers) not used in `xml`, one per key
struct SdtIds {
    used: std::collections::HashSet<i32>,
}

impl SdtIds {
    fn new(xml: &str) -> Self {
        let mut used = std::collections::HashSet::new();
        let mut rest = xml;
        while let Some(at) = rest.find("w:id w:val=\"") {
            rest = &rest[at + 12..];
            if let Some(n) = rest.split('"').next().and_then(|v| v.parse::<i32>().ok()) {
                used.insert(n);
            }
        }
        SdtIds { used }
    }

    fn next(&mut self, key: &str) -> i32 {
        let mut hasher = std::collections::hash_map::DefaultHasher::new();
        key.hash(&mut hasher);
        let mut id = (hasher.finish() % 0x3FFF_FFFF) as i32 + 1;
        while self.used.contains(&id) {
            id = id % 0x3FFF_FFFF + 1;
        }
        self.used.insert(id);
        id
    }
}

/// Byte position at the end of the body, before its final `w:sectPr`
fn body_end(xml: &str) -> Result<usize, String> {
    let end = xml.rfind("</w:body>").ok_or("El documento no tiene cuerpo (w:body).")?;
    match xml[..end].rfind("<w:sectPr") {
        // The body's own sectPr is its last child: no paragraph closes after it
        Some(at) if !xml[at..end].contains("</w:p>") => Ok(at),
        _ => Ok(end),
    }
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
            let (first, last) = match &loc.content {
                Some(c) => {
                    // Paragraph ranges are sorted: binary search the first and last inside
                    let a = paragraphs.partition_point(|r| r.start < c.start);
                    let b = paragraphs.partition_point(|r| r.end <= c.end);
                    if a < b { (Some(a), Some(b - 1)) } else { (None, None) }
                }
                None => (None, None),
            };
            out.push(SectionInfo {
                key: loc.key.clone(),
                title: loc.title.clone().unwrap_or_else(|| loc.key.clone()),
                first,
                last,
                depth: open.len(),
                parent: open.last().map(|(p, _)| locs[*p].key.clone()),
                locked: loc.locked,
                content_locked: loc.content_locked,
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

    fn section_info(&self, key: &str) -> Result<SectionInfo, String> {
        self.sections()?
            .into_iter()
            .find(|s| s.key == key)
            .ok_or_else(|| format!("No se pudo crear la sección «{}».", key))
    }

    /// Wraps the blocks from paragraph `from` to `to` (both included, with the tables between
    /// them) in a new section. A table is wrapped whole when `from` is its first paragraph or
    /// `to` its last one. `locked`: Word users cannot delete the section; `content_locked`:
    /// nor edit its content (code still can). Refuses keys already in use and ranges that cut
    /// a table or another section, lie inside a table, or hold a section break.
    pub fn define_section(
        &mut self,
        key: &str,
        from: usize,
        to: usize,
        title: Option<&str>,
        locked: bool,
        content_locked: bool,
    ) -> Result<SectionInfo, String> {
        check_section_key(key)?;
        let mut xml = self.get_file_string("word/document.xml")?;
        let tokens = tokenize(&xml)?;
        if locate_sections(&tokens).iter().any(|l| l.key == key) {
            return Err(format!("Ya existe una sección con la clave «{}».", key));
        }
        let (from, to) = if to < from { (to, from) } else { (from, to) };
        let paragraphs = body_paragraph_ranges(&xml)?;
        let span = block_span(&xml, &paragraphs, &body_table_spans(&tokens), from, to)?;
        let id = SdtIds::new(&xml).next(key);
        let title = title.map(str::trim).filter(|t| !t.is_empty()).unwrap_or(key);
        let wrapped = format!("{}{}{}", section_open(key, title, id, locked, content_locked), &xml[span.clone()], SECTION_CLOSE);
        xml.replace_range(span, &wrapped);
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        self.section_info(key)
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

    /// Sets a section's locks: `locked`, it cannot be deleted; `content_locked`, its content
    /// cannot be edited by hand (Word and this editor; code still can fill it)
    pub fn set_section_lock(&mut self, key: &str, locked: bool, content_locked: bool) -> Result<SectionInfo, String> {
        let mut xml = self.get_file_string("word/document.xml")?;
        let loc = Self::find_section(&xml, key)?;
        let pr = &xml[loc.properties.clone()];
        let tokens = tokenize(pr)?;
        // Drop the current w:lock, then put the new one where CT_SdtPr wants it: after
        // rPr / alias / tag / id, before everything else
        let mut out = String::with_capacity(pr.len() + 40);
        let mut insert_at = None;
        let mut i = 1;
        out.push_str(&pr[tokens[0].span.clone()]);
        insert_at.get_or_insert(out.len());
        let last = tokens.len() - 1;
        while i < last {
            let tok = &tokens[i];
            let end = element_end(&tokens, i);
            let piece = &pr[tok.span.start..tokens[end].span.end];
            if is_element(tok, "lock") {
                i = end + 1;
                continue;
            }
            out.push_str(piece);
            if ["rPr", "alias", "tag", "id"].iter().any(|n| is_element(tok, n)) {
                insert_at = Some(out.len());
            }
            i = end + 1;
        }
        out.push_str(&pr[tokens[last].span.clone()]);
        out.insert_str(insert_at.unwrap_or(0), lock_element(locked, content_locked));
        xml.replace_range(loc.properties, &out);
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        self.section_info(key)
    }

    /// Inserts a new section with `items` as its content (e.g. a heading and an empty
    /// paragraph) at `place`. Returns the section and the first new paragraph.
    pub fn insert_section(
        &mut self,
        key: &str,
        title: Option<&str>,
        place: &SectionPlace,
        items: &[NewParagraph],
        locked: bool,
        content_locked: bool,
    ) -> Result<(SectionInfo, usize), String> {
        check_section_key(key)?;
        if items.is_empty() {
            return Err("Una sección nueva necesita al menos un párrafo.".to_string());
        }
        let mut xml = self.get_file_string("word/document.xml")?;
        if locate_sections(&tokenize(&xml)?).iter().any(|l| l.key == key) {
            return Err(format!("Ya existe una sección con la clave «{}».", key));
        }
        let at = match place {
            SectionPlace::After(k) => Self::find_section(&xml, k)?.element.end,
            SectionPlace::Before(k) => Self::find_section(&xml, k)?.element.start,
            SectionPlace::Inside(k) => {
                let loc = Self::find_section(&xml, k)?;
                loc.content.map(|c| c.end).ok_or_else(|| format!("La sección «{}» está vacía; usa after o before.", k))?
            }
            SectionPlace::End => body_end(&xml)?,
        };
        let index = body_paragraph_ranges(&xml)?.iter().filter(|r| r.end <= at).count();
        xml.insert_str(at, "<w:p/>");
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        let (first, _) = self.insert_paragraphs(index, 0, items)?;
        let info = self.define_section(key, first, first + items.len() - 1, title, locked, content_locked)?;
        Ok((info, first))
    }

    /// Wraps each heading of level `from_level` to `to_level` (1 = Heading 1) and what follows
    /// it, up to the next heading of the same or a higher level, in a section keyed after its
    /// text ("2. Caracterización" → "caracterizacion"). Lower levels nest in higher ones.
    /// Headings whose part is already a section, or whose range would cut a table, another
    /// content control or hold a section break, are skipped and reported.
    pub fn sections_from_headings(&mut self, from_level: u32, to_level: u32) -> Result<HeadingSections, String> {
        if from_level == 0 || to_level < from_level || to_level > 9 {
            return Err("Los niveles de título van de 1 a 9 (p. ej. de 1 a 2).".to_string());
        }
        let xml = self.get_file_string("word/document.xml")?;
        let tokens = tokenize(&xml)?;
        let paragraphs = body_paragraph_ranges(&xml)?;
        let tables = body_table_spans(&tokens);
        let existing = locate_sections(&tokens);
        let infos = self.sections()?;
        let mut taken: std::collections::HashSet<String> = existing.iter().map(|l| l.key.clone()).collect();
        // Headings of the body (not in tables or text boxes), with their 1-based level
        let headings: Vec<(usize, u32, String)> = self
            .extract_paragraphs()?
            .into_iter()
            .filter(|p| p.index < paragraphs.len())
            .filter_map(|p| p.outline_level.map(|l| (p.index, l + 1, p.text.trim().to_string())))
            .filter(|(_, l, _)| *l <= 9)
            .collect();
        let mut ids = SdtIds::new(&xml);
        let mut created: Vec<(String, Range<usize>, String)> = Vec::new(); // key, bytes, opening
        let mut skipped = Vec::new();
        for (n, (index, level, text)) in headings.iter().enumerate() {
            if *level < from_level || *level > to_level {
                continue;
            }
            let skip = |reason: String| SkippedHeading { paragraph: *index, text: text.clone(), reason };
            // Up to the next heading of the same or a higher level, within the section (or
            // table cell…) that holds the heading
            let mut last = headings[n + 1..]
                .iter()
                .find(|(_, l, _)| l <= level)
                .map(|(i, _, _)| i - 1)
                .unwrap_or(paragraphs.len() - 1);
            let start = paragraphs[*index].start;
            if let Some(container) = existing
                .iter()
                .filter_map(|l| l.content.clone())
                .filter(|c| c.start <= start && start < c.end)
                .min_by_key(|c| c.end - c.start)
            {
                last = last.min(paragraphs.partition_point(|r| r.end <= container.end).saturating_sub(1));
            }
            // A section that already spans exactly this heading's part is the same section
            if let Some(s) = infos.iter().find(|s| s.first == Some(*index) && s.last == Some(last)) {
                skipped.push(skip(format!("ya abre la sección «{}»", s.key)));
                continue;
            }
            let span = match block_span(&xml, &paragraphs, &tables, *index, last) {
                Ok(span) => span,
                Err(reason) => {
                    skipped.push(skip(reason));
                    continue;
                }
            };
            let key = unique_key(&section_key_from_title(text), &taken);
            taken.insert(key.clone());
            let title = if text.is_empty() { key.clone() } else { text.clone() };
            let opening = section_open(&key, &title, ids.next(&key), false, false);
            created.push((key, span, opening));
        }
        if !created.is_empty() {
            // The new sections nest properly (a lower heading ends no later than the one above
            // it): insert every opening and closing tag in one pass. At the same position,
            // closings go first, the inner section's before the outer one's.
            let mut inserts: Vec<(usize, u8, std::cmp::Reverse<usize>, &str)> = Vec::new();
            for (_, span, opening) in &created {
                inserts.push((span.start, 1, std::cmp::Reverse(0), opening.as_str()));
                inserts.push((span.end, 0, std::cmp::Reverse(span.start), SECTION_CLOSE));
            }
            inserts.sort_by(|a, b| (a.0, a.1, a.2).cmp(&(b.0, b.1, b.2)));
            let mut out = String::with_capacity(xml.len() + inserts.iter().map(|i| i.3.len()).sum::<usize>());
            let mut pos = 0;
            for (at, _, _, piece) in inserts {
                out.push_str(&xml[pos..at]);
                out.push_str(piece);
                pos = at;
            }
            out.push_str(&xml[pos..]);
            self.put_file("word/document.xml".to_string(), out.into_bytes());
        }
        let keys: std::collections::HashSet<&str> = created.iter().map(|(k, _, _)| k.as_str()).collect();
        let created = self.sections()?.into_iter().filter(|s| keys.contains(s.key.as_str())).collect();
        Ok(HeadingSections { created, skipped })
    }
}
