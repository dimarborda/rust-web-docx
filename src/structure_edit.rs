//! Document structure from code: Word heading styles and replacing whole ranges of blocks
//! (`replaceParagraphs`), each as one undo step.
//!
//! Child module of `docx_parser` so it can work on the package parts directly.

use super::*;
use crate::paragraph_edit::set_paragraph_property;

/// Built-in styles the editor can apply, by `NewParagraph.style`: (id, Word's name, definition).
/// Gray tones and Word's sizes; headings carry their outline level so they become sections.
const HEADING_STYLES: &[(&str, &str, &str)] = &[
    (
        "Title",
        "title",
        r#"<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="240" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:b/><w:color w:val="262626"/><w:kern w:val="28"/><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr></w:style>"#,
    ),
    (
        "Heading1",
        "heading 1",
        r#"<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:color w:val="262626"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>"#,
    ),
    (
        "Heading2",
        "heading 2",
        r#"<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:color w:val="404040"/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>"#,
    ),
    (
        "Heading3",
        "heading 3",
        r#"<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="60"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:color w:val="595959"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>"#,
    ),
];

impl DocxModifier {
    /// Re-reads styles, theme and numbering (after a style is added, or undo/redo touched them)
    pub(super) fn reload_styles(&mut self) {
        let files = &self.files;
        let part = |path: &str| files.get(path).and_then(|b| std::str::from_utf8(b).ok());
        let theme_path = files.keys().filter(|k| k.starts_with("word/theme/") && k.ends_with(".xml")).min().cloned();
        self.styles = StyleSheet::load(part("word/styles.xml"), theme_path.as_deref().and_then(part), part("word/numbering.xml"));
        self.revision += 1;
    }

    /// Id of the document's style for `kind` (Title, Heading1–3): the one Word already defines
    /// under its built-in name (e.g. "Ttulo1" in Spanish documents) or a new one added here
    fn ensure_heading_style(&mut self, kind: &str) -> Result<String, String> {
        let (id, name, definition) = HEADING_STYLES
            .iter()
            .find(|(id, _, _)| *id == kind)
            .ok_or_else(|| format!("Estilo no válido: {}. Usa Title, Heading1, Heading2 o Heading3.", kind))?;
        let xml = self
            .get_file_string("word/styles.xml")
            .map_err(|_| "El documento no tiene hoja de estilos; no se pueden crear títulos.".to_string())?;
        let defined = |style_id: &str| xml.contains(&format!("w:styleId=\"{}\"", style_id));
        let existing = self
            .styles
            .styles
            .values()
            .filter(|s| s.kind == "paragraph" && s.name.eq_ignore_ascii_case(name))
            .map(|s| s.id.clone())
            .find(|style_id| defined(style_id));
        if let Some(style_id) = existing {
            return Ok(style_id);
        }
        if defined(id) {
            return Ok(id.to_string());
        }
        let end = xml.rfind("</w:styles>").ok_or("La hoja de estilos del documento no es válida.")?;
        let mut updated = xml.clone();
        updated.insert_str(end, definition);
        self.put_file("word/styles.xml".to_string(), updated.into_bytes());
        self.reload_styles();
        Ok(id.to_string())
    }

    /// Applies `NewParagraph.style` to the paragraphs inserted from `first`. Paragraphs without
    /// a style must not inherit a heading from the paragraph they were split from (inserting
    /// before a title would otherwise turn the new text into titles): they go back to the
    /// default paragraph style.
    pub(super) fn apply_paragraph_styles(&mut self, first: usize, items: &[NewParagraph]) -> Result<(), String> {
        let headings: std::collections::HashSet<usize> = self
            .extract_paragraphs()?
            .into_iter()
            .filter(|p| p.is_heading || p.outline_level.is_some())
            .map(|p| p.index)
            .collect();
        let mut changes: Vec<(usize, Option<String>)> = Vec::new();
        for (i, item) in items.iter().enumerate() {
            let index = first + i;
            match item.style.as_deref() {
                Some(kind) => changes.push((index, Some(self.ensure_heading_style(kind)?))),
                None if headings.contains(&index) => changes.push((index, None)),
                None => {}
            }
        }
        if changes.is_empty() {
            return Ok(());
        }
        let mut xml = self.get_file_string("word/document.xml")?;
        let ranges = body_paragraph_ranges(&xml)?;
        for (index, style_id) in changes.iter().rev() {
            let range = ranges.get(*index).cloned().ok_or_else(|| format!("No existe el párrafo {}.", index))?;
            let element = style_id.as_ref().map(|id| format!(r#"<w:pStyle w:val="{}"/>"#, escape_xml(id)));
            let mut edited = set_paragraph_property(&xml[range.clone()], "pStyle", element.as_deref())?;
            if style_id.is_none() {
                // A direct outline level would keep it in the outline
                edited = set_paragraph_property(&edited, "outlineLvl", None)?;
            }
            xml.replace_range(range, &edited);
        }
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        Ok(())
    }

    /// Replaces every block from paragraph `from` to paragraph `to` (both included, with the
    /// tables between them) by `items`. Returns the first new paragraph and the caret after the
    /// last one. Refuses ranges that cut a table or a content control, or hold a section break.
    pub fn replace_paragraphs(&mut self, from: usize, to: usize, items: &[NewParagraph]) -> Result<(usize, (usize, usize)), String> {
        if items.is_empty() {
            return Err("No hay párrafos para insertar.".to_string());
        }
        let (from, to) = if to < from { (to, from) } else { (from, to) };
        let mut xml = self.get_file_string("word/document.xml")?;
        let ranges = body_paragraph_ranges(&xml)?;
        let (start, end) = match (ranges.get(from), ranges.get(to)) {
            (Some(a), Some(b)) => (a.start, b.end),
            _ => return Err(format!("El rango {}–{} no existe en el documento.", from, to)),
        };
        if !is_balanced(&xml[start..end]) {
            return Err("El rango corta una tabla o un control de contenido; amplíalo para incluirlos completos.".to_string());
        }
        if xml[start..end].contains("<w:sectPr") {
            return Err("El rango incluye un salto de sección; no se puede reemplazar de una vez.".to_string());
        }
        xml.replace_range(start..end, "<w:p/>");
        self.put_file("word/document.xml".to_string(), xml.into_bytes());
        self.insert_paragraphs(from, 0, items)
    }
}
