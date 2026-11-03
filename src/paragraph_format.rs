//! Paragraph formatting from the toolbar: paragraph styles, line spacing, indents and lists
//! (bullets and numbering with levels), each as a minimal edit of `w:pPr`.
//!
//! Child module of `docx_parser` so it can work on the package parts directly.

use super::*;
use crate::paragraph_edit::{edit_paragraph, parse_paragraph_fragment, set_paragraph_attrs, set_paragraph_property, FormatTarget};
use crate::styles::{NumDef, Style};
use ooxml_core::content_types;
use ooxml_core::rels::{rels_part_for, with_relationship};

const NUMBERING: &str = "word/numbering.xml";
const DOCUMENT: &str = "word/document.xml";
const REL_NUMBERING: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const MIME_NUMBERING: &str = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/// Built-in styles the gallery always offers (created on first use when the document lacks
/// them): (id, Word's built-in name)
const BUILT_IN: &[(&str, &str)] = &[("Title", "title"), ("Heading1", "heading 1"), ("Heading2", "heading 2"), ("Heading3", "heading 3")];

/// One entry of the paragraph style gallery
#[derive(Serialize, Debug, Clone, PartialEq)]
pub struct StyleEntry {
    pub id: String,
    pub name: String,
    /// Word's built-in style this one is ("normal", "title", "heading 1", ...) so the UI can
    /// show it in its own language; `None` for custom styles
    pub builtin: Option<String>,
}

/// Bullets of each list level, and the number formats Word cycles through
const BULLETS: [&str; 3] = ["•", "◦", "▪"];
const NUMBER_FORMATS: [&str; 3] = ["decimal", "lowerLetter", "lowerRoman"];

/// Left indent of each list level step (0.5") and the hanging indent of its label (0.25")
const LEVEL_INDENT: i64 = 720;
const LABEL_HANGING: i64 = 360;

impl DocxModifier {
    /// Paragraph styles for a style gallery, in Word's order: quick styles of the document
    /// plus Title and Heading 1–3, which are created when applied if the document lacks them
    pub fn paragraph_styles(&self) -> Vec<StyleEntry> {
        let mut styles: Vec<&Style> = self
            .styles
            .styles
            .values()
            .filter(|s| s.kind == "paragraph" && s.quick && !s.hidden)
            .collect();
        styles.sort_by(|a, b| {
            let pa = a.ui_priority.unwrap_or(99);
            let pb = b.ui_priority.unwrap_or(99);
            pa.cmp(&pb).then_with(|| a.name.to_lowercase().cmp(&b.name.to_lowercase()))
        });
        let builtin = |name: &str| {
            let n = name.to_lowercase();
            let known = n == "normal" || n == "title" || n == "subtitle" || n == "quote" || n == "intense quote"
                || n == "list paragraph" || n == "no spacing" || n.starts_with("heading ");
            known.then_some(n)
        };
        let default_id = self.styles.default_paragraph_style.clone();
        let mut out: Vec<StyleEntry> = Vec::new();
        // The default paragraph style first, even when it is not marked as a quick style
        if let Some(id) = &default_id {
            if let Some(s) = self.styles.styles.get(id) {
                out.push(StyleEntry { id: id.clone(), name: s.name.clone(), builtin: Some("normal".to_string()) });
            }
        } else {
            out.push(StyleEntry { id: "Normal".to_string(), name: "Normal".to_string(), builtin: Some("normal".to_string()) });
        }
        for s in styles {
            if Some(&s.id) != default_id.as_ref() {
                out.push(StyleEntry { id: s.id.clone(), name: s.name.clone(), builtin: builtin(&s.name) });
            }
        }
        for (id, name) in BUILT_IN {
            if !out.iter().any(|e| e.builtin.as_deref() == Some(name)) {
                out.push(StyleEntry { id: id.to_string(), name: name.to_string(), builtin: Some(name.to_string()) });
            }
        }
        out
    }

    /// Applies paragraph style `id` (an id from `paragraph_styles`) to the paragraphs.
    /// The default style removes `w:pStyle`; Title and Heading 1–3 are created if missing.
    pub fn set_paragraph_style(&mut self, indices: &[usize], id: &str) -> Result<(), String> {
        let is_default = id.is_empty() || self.styles.default_paragraph_style.as_deref() == Some(id)
            || (id == "Normal" && !self.styles.styles.contains_key(id));
        let style_id = if is_default {
            None
        } else if self.styles.styles.get(id).is_some_and(|s| s.kind == "paragraph") {
            Some(id.to_string())
        } else if BUILT_IN.iter().any(|(b, _)| *b == id) {
            Some(self.ensure_heading_style(id)?)
        } else {
            return Err(format!("El documento no tiene el estilo de párrafo «{}».", id));
        };
        let element = style_id.as_ref().map(|id| format!(r#"<w:pStyle w:val="{}"/>"#, escape_xml(id)));
        self.edit_paragraphs_each(indices, |_, _, p| {
            let mut edited = set_paragraph_property(p, "pStyle", element.as_deref())?;
            // A direct outline level would keep a paragraph that is no longer a heading in
            // the outline, and a heading style brings its own
            edited = set_paragraph_property(&edited, "outlineLvl", None)?;
            Ok(edited)
        })
    }

    /// Line spacing as a multiple of single spacing (1, 1.15, 1.5, 2, ...), keeping the space
    /// before and after
    pub fn set_line_spacing(&mut self, indices: &[usize], multiple: f64) -> Result<(), String> {
        if !(multiple > 0.0 && multiple <= 10.0) {
            return Err("El interlineado debe estar entre 0 y 10 líneas.".to_string());
        }
        let line = ((multiple * 240.0).round() as i64).to_string();
        self.edit_paragraphs_each(indices, |_, _, p| {
            set_paragraph_attrs(p, "spacing", &[("line", Some(line.clone())), ("lineRule", Some("auto".to_string()))])
        })
    }

    /// Space before and after the paragraphs, in points (`None` keeps the current value)
    pub fn set_paragraph_spacing(&mut self, indices: &[usize], before: Option<f64>, after: Option<f64>) -> Result<(), String> {
        let twips = |pt: f64| ((pt.max(0.0) * 20.0).round() as i64).to_string();
        let mut changes: Vec<(&str, Option<String>)> = Vec::new();
        if let Some(b) = before {
            changes.push(("before", Some(twips(b))));
            changes.push(("beforeAutospacing", None));
            changes.push(("beforeLines", None));
        }
        if let Some(a) = after {
            changes.push(("after", Some(twips(a))));
            changes.push(("afterAutospacing", None));
            changes.push(("afterLines", None));
        }
        if changes.is_empty() {
            return Ok(());
        }
        self.edit_paragraphs_each(indices, |_, _, p| set_paragraph_attrs(p, "spacing", &changes))
    }

    /// Increase (`steps` > 0) or decrease indent, as Word's buttons: list items move to the
    /// next or previous level; other paragraphs move their left indent by 0.5"
    pub fn change_indent(&mut self, indices: &[usize], steps: i32) -> Result<(), String> {
        if steps == 0 {
            return Ok(());
        }
        self.edit_paragraphs_each(indices, |styles, _, p| {
            let info = parse_paragraph_fragment(p, styles);
            match (info.num_id.as_deref(), info.list_level) {
                (Some(num_id), Some(level)) => {
                    let level = (level as i32 + steps).clamp(0, 8) as u32;
                    set_paragraph_property(p, "numPr", Some(&num_pr(num_id, level)))
                }
                _ => {
                    // The effective indent (style included) is in px: 1px = 15 twips
                    let current = (info.indent_left * 15.0).round() as i64;
                    let step = LEVEL_INDENT;
                    // Snap to the 0.5" grid, as Word does
                    let next = if steps > 0 {
                        (current / step + steps as i64) * step
                    } else {
                        ((current + step - 1) / step + steps as i64).max(0) * step
                    };
                    set_paragraph_attrs(p, "ind", &[("left", Some(next.max(0).to_string())), ("start", None)])
                }
            }
        })
    }

    /// Turns the paragraphs into a bulleted (`"bullet"`) or numbered (`"number"`) list, or
    /// takes them out of any list (`"none"`). A list right before the first paragraph of the
    /// same kind continues; otherwise a numbered list starts again at 1.
    pub fn set_list(&mut self, indices: &[usize], kind: &str) -> Result<(), String> {
        let mut sorted: Vec<usize> = indices.to_vec();
        sorted.sort_unstable();
        sorted.dedup();
        let Some(&first) = sorted.first() else { return Ok(()) };
        let list_style = self.list_paragraph_style();
        let default_style = self.styles.default_paragraph_style.clone();

        if kind == "none" {
            return self.edit_paragraphs_each(&sorted, |styles, _, p| {
                let mut edited = set_paragraph_property(p, "numPr", None)?;
                // Numbering that comes from the paragraph's style is switched off with numId 0
                if parse_paragraph_fragment(&edited, styles).num_id.is_some() {
                    edited = set_paragraph_property(&edited, "numPr", Some(&num_pr("0", 0)))?;
                }
                let info = parse_paragraph_fragment(&edited, styles);
                if list_style.as_deref().is_some_and(|ls| info.style == ls) {
                    edited = set_paragraph_property(&edited, "pStyle", None)?;
                }
                Ok(edited)
            });
        }
        if kind != "bullet" && kind != "number" {
            return Err(format!("Tipo de lista no válido: {}. Usa bullet, number o none.", kind));
        }

        // Continue the list right above (Word does the same when it has the same kind)
        let previous = first
            .checked_sub(1)
            .filter(|i| *i < TEXT_BOX_BASE && first < TEXT_BOX_BASE)
            .and_then(|i| self.extract_paragraphs().ok()?.into_iter().find(|p| p.index == i))
            .filter(|p| p.list_kind.as_deref() == Some(kind))
            .and_then(|p| p.num_id);
        let num_id = match previous {
            Some(id) => id,
            None => self.new_list(kind)?,
        };
        self.edit_paragraphs_each(&sorted, |styles, _, p| {
            let info = parse_paragraph_fragment(p, styles);
            // Switching between bullets and numbers keeps the level
            let level = info.list_level.unwrap_or(0);
            let mut edited = set_paragraph_property(p, "numPr", Some(&num_pr(&num_id, level)))?;
            let plain = info.style.is_empty() || Some(&info.style) == default_style.as_ref() || info.style == "Normal";
            if let (true, Some(ls)) = (plain, list_style.as_deref()) {
                edited = set_paragraph_property(&edited, "pStyle", Some(&format!(r#"<w:pStyle w:val="{}"/>"#, escape_xml(ls))))?;
            }
            Ok(edited)
        })
    }

    /// Removes the direct character formatting (font, size, bold, italic, underline,
    /// strikethrough, superscript/subscript, color, highlight) from (`p1`, `o1`) to (`p2`, `o2`)
    pub fn clear_formatting(&mut self, p1: usize, o1: usize, p2: usize, o2: usize) -> Result<(), String> {
        let ((p1, o1), (p2, o2)) = if (p2, o2) < (p1, o1) { ((p2, o2), (p1, o1)) } else { ((p1, o1), (p2, o2)) };
        let indices: Vec<usize> = (p1..=p2).collect();
        self.edit_paragraphs_each(&indices, |styles, i, p| {
            let info = parse_paragraph_fragment(p, styles);
            let len = info.text.chars().count();
            let start = if i == p1 { o1.min(len) } else { 0 };
            let end = if i == p2 { o2.min(len) } else { len };
            if start >= end {
                return Ok(p.to_string());
            }
            let mut formats = Vec::with_capacity(len);
            for run in &info.runs {
                for _ in run.text.chars() {
                    let at = formats.len();
                    formats.push(if at >= start && at < end {
                        FormatTarget { clear: true, ..Default::default() }
                    } else {
                        FormatTarget::from_run(run)
                    });
                }
            }
            edit_paragraph(p, styles, &info.text, Some(&formats), None)
        })
    }

    /// Word's "List Paragraph" style, when the document defines it
    fn list_paragraph_style(&self) -> Option<String> {
        self.styles
            .styles
            .values()
            .find(|s| s.kind == "paragraph" && s.name.eq_ignore_ascii_case("list paragraph"))
            .map(|s| s.id.clone())
    }

    /// A `w:num` for a new list of `kind`: bullets share one definition; each numbered list
    /// gets its own `w:num` restarting at 1. Creates numbering.xml when the document has none.
    fn new_list(&mut self, kind: &str) -> Result<String, String> {
        let numbering = &self.styles.numbering;
        let level0 = |num: &NumDef| numbering.abstracts.get(&num.abstract_id).and_then(|l| l.get(&0));
        let is_bullet = |num: &NumDef| level0(num).is_some_and(|l| l.fmt == "bullet");
        let is_decimal = |num: &NumDef| level0(num).is_some_and(|l| l.fmt == "decimal" && l.text == "%1.");

        if kind == "bullet" {
            // An existing bullet list definition keeps the document's own look
            let mut ids: Vec<(&String, &NumDef)> = numbering.nums.iter().filter(|(_, n)| is_bullet(n)).collect();
            ids.sort_by_key(|(id, _)| id.parse::<i64>().unwrap_or(i64::MAX));
            if let Some((id, _)) = ids.first() {
                return Ok((*id).clone());
            }
        }
        let existing_abstract = if kind == "number" {
            let mut found: Vec<&NumDef> = numbering.nums.values().filter(|n| is_decimal(n)).collect();
            found.sort_by_key(|n| n.abstract_id.parse::<i64>().unwrap_or(i64::MAX));
            found.first().map(|n| n.abstract_id.clone())
        } else {
            None
        };

        let next_id = |ids: Vec<&String>| ids.iter().filter_map(|id| id.parse::<i64>().ok()).max().map_or(1, |m| m + 1);
        let next_abstract = next_id(self.styles.numbering.abstracts.keys().collect()).to_string();
        let next_num = next_id(self.styles.numbering.nums.keys().collect()).to_string();

        let mut xml = match self.get_file_string(NUMBERING) {
            Ok(xml) => xml,
            Err(_) => {
                self.add_numbering_part()?;
                self.get_file_string(NUMBERING)?
            }
        };
        let abstract_id = match existing_abstract {
            Some(id) => id,
            None => {
                let id = next_abstract;
                let definition = abstract_num(&id, kind);
                let at = find_tag(&xml, "num").or_else(|| find_tag(&xml, "numIdMacAtCleanup")).or_else(|| xml.rfind("</w:numbering>"))
                    .ok_or("La numeración del documento no es válida.")?;
                xml.insert_str(at, &definition);
                id
            }
        };
        let num_id = next_num;
        let restart = if kind == "number" {
            r#"<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>"#
        } else {
            ""
        };
        let num = format!(r#"<w:num w:numId="{}"><w:abstractNumId w:val="{}"/>{}</w:num>"#, num_id, abstract_id, restart);
        let at = find_tag(&xml, "numIdMacAtCleanup").or_else(|| xml.rfind("</w:numbering>")).ok_or("La numeración del documento no es válida.")?;
        xml.insert_str(at, &num);
        self.put_file(NUMBERING.to_string(), xml.into_bytes());
        self.reload_styles();
        Ok(num_id)
    }

    /// Adds an empty numbering part with its relationship and content type
    fn add_numbering_part(&mut self) -> Result<(), String> {
        let part = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="{}"></w:numbering>"#,
            W_NS
        );
        let rels_part = rels_part_for(DOCUMENT);
        let rels = self.get_file_string(&rels_part).unwrap_or_else(|_| {
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#.to_string()
        });
        let (rels, _) = with_relationship(&rels, REL_NUMBERING, "numbering.xml")?;
        let types = self.get_file_string(content_types::PART)?;
        self.package.keep_in_order(NUMBERING);
        self.put_file(NUMBERING.to_string(), part.into_bytes());
        self.put_file(rels_part, rels.into_bytes());
        if let Some(types) = content_types::with_override(&types, NUMBERING, MIME_NUMBERING)? {
            self.put_file(content_types::PART.to_string(), types.into_bytes());
        }
        Ok(())
    }

    /// Applies `f` to each paragraph (body, table, text box, header or footer) as one change
    /// per part; the document is only written when something changed
    fn edit_paragraphs_each(
        &mut self,
        indices: &[usize],
        f: impl Fn(&StyleSheet, usize, &str) -> Result<String, String>,
    ) -> Result<(), String> {
        let mut sorted: Vec<usize> = indices.to_vec();
        sorted.sort_unstable();
        sorted.dedup();
        let (body, other): (Vec<usize>, Vec<usize>) = sorted.into_iter().partition(|i| *i < TEXT_BOX_BASE);
        if !body.is_empty() {
            let original = self.get_file_string(DOCUMENT)?;
            let mut xml = original.clone();
            let ranges = body_paragraph_ranges(&xml)?;
            for i in body.iter().rev() {
                let range = ranges.get(*i).cloned().ok_or_else(|| format!("No existe el párrafo {}.", i))?;
                let edited = f(&self.styles, *i, &xml[range.clone()])?;
                xml.replace_range(range, &edited);
            }
            if xml != original {
                self.put_file(DOCUMENT.to_string(), xml.into_bytes());
            }
        }
        for i in other {
            let (part, mut xml, range) = self.paragraph_in_part(i)?;
            let edited = f(&self.styles, i, &xml[range.clone()])?;
            if edited != xml[range.clone()] {
                xml.replace_range(range, &edited);
                self.put_file(part, xml.into_bytes());
            }
        }
        Ok(())
    }
}

fn num_pr(num_id: &str, level: u32) -> String {
    format!(r#"<w:numPr><w:ilvl w:val="{}"/><w:numId w:val="{}"/></w:numPr>"#, level, escape_xml(num_id))
}

/// Byte offset of the first `<w:tag ` / `<w:tag>` element
fn find_tag(xml: &str, tag: &str) -> Option<usize> {
    let open = format!("<w:{}", tag);
    xml.match_indices(&open)
        .find(|(i, _)| matches!(xml.as_bytes().get(i + open.len()), Some(b' ' | b'>' | b'/')))
        .map(|(i, _)| i)
}

/// A nine-level list definition: bullets • ◦ ▪ or numbers 1. a. i., each level 0.5" further in
fn abstract_num(id: &str, kind: &str) -> String {
    let mut levels = String::new();
    for l in 0..9usize {
        let (fmt, text) = if kind == "bullet" {
            ("bullet".to_string(), BULLETS[l % 3].to_string())
        } else {
            (NUMBER_FORMATS[l % 3].to_string(), format!("%{}.", l + 1))
        };
        levels.push_str(&format!(
            r#"<w:lvl w:ilvl="{l}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{left}" w:hanging="{hanging}"/></w:pPr></w:lvl>"#,
            left = LEVEL_INDENT * (l as i64 + 1),
            hanging = LABEL_HANGING,
        ));
    }
    format!(
        r#"<w:abstractNum w:abstractNumId="{}"><w:multiLevelType w:val="hybridMultilevel"/>{}</w:abstractNum>"#,
        id, levels
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::blank_generator::{generate_blank_docx, PageSize};

    /// A blank document (no numbering part) with `texts` as its paragraphs
    fn doc(texts: &[&str]) -> DocxModifier {
        let mut m = DocxModifier::from_bytes(&generate_blank_docx(PageSize::A4).unwrap()).unwrap();
        let items: Vec<NewParagraph> = texts.iter().map(|t| NewParagraph { text: t.to_string(), ..Default::default() }).collect();
        m.replace_paragraphs(0, 0, &items).unwrap();
        m
    }

    fn reopen(m: &DocxModifier) -> DocxModifier {
        DocxModifier::from_bytes(&m.to_bytes().unwrap()).unwrap()
    }

    #[test]
    fn test_bullets_create_numbering_part_and_survive_reopening() {
        let mut m = doc(&["uno", "dos", "tres"]);
        assert!(m.get_file_string(NUMBERING).is_err());
        m.set_list(&[0, 1], "bullet").unwrap();
        let m = reopen(&m);
        let ps = m.extract_paragraphs().unwrap();
        assert_eq!(ps[0].list_kind.as_deref(), Some("bullet"));
        assert_eq!(ps[1].list_kind.as_deref(), Some("bullet"));
        assert_eq!(ps[2].list_kind, None);
        assert_eq!(ps[0].list_label.as_ref().map(|l| l.text.as_str()), Some("•"));
        assert!(m.get_file_string("[Content_Types].xml").unwrap().contains("/word/numbering.xml"));
        assert!(m.get_file_string("word/_rels/document.xml.rels").unwrap().contains("numbering.xml"));
        // The level's indent applies: 0.5" = 48 px
        assert_eq!(ps[0].indent_left, 48.0);
    }

    #[test]
    fn test_numbered_lists_restart_and_continue() {
        let mut m = doc(&["a", "b", "texto", "c", "d"]);
        m.set_list(&[0, 1], "number").unwrap();
        m.set_list(&[3, 4], "number").unwrap();
        let labels: Vec<Option<String>> = m.extract_paragraphs().unwrap().into_iter().map(|p| p.list_label.map(|l| l.text)).collect();
        assert_eq!(labels[0].as_deref(), Some("1."));
        assert_eq!(labels[1].as_deref(), Some("2."));
        assert_eq!(labels[2], None);
        // A separate list starts again at 1
        assert_eq!(labels[3].as_deref(), Some("1."));
        assert_eq!(labels[4].as_deref(), Some("2."));

        // A paragraph right after a list joins it
        let mut m = doc(&["a", "b", "c"]);
        m.set_list(&[0, 1], "number").unwrap();
        m.set_list(&[2], "number").unwrap();
        let ps = m.extract_paragraphs().unwrap();
        assert_eq!(ps[2].list_label.as_ref().map(|l| l.text.as_str()), Some("3."));
        assert_eq!(ps[0].num_id, ps[2].num_id);
    }

    #[test]
    fn test_list_levels_and_removal() {
        let mut m = doc(&["a", "b"]);
        m.set_list(&[0, 1], "number").unwrap();
        m.change_indent(&[1], 1).unwrap();
        let ps = m.extract_paragraphs().unwrap();
        assert_eq!(ps[1].list_level, Some(1));
        assert_eq!(ps[1].list_label.as_ref().map(|l| l.text.as_str()), Some("a."));
        m.change_indent(&[1], -3).unwrap();
        assert_eq!(m.extract_paragraphs().unwrap()[1].list_level, Some(0));

        // Switching kind keeps the level
        m.change_indent(&[1], 1).unwrap();
        m.set_list(&[1], "bullet").unwrap();
        let ps = m.extract_paragraphs().unwrap();
        assert_eq!((ps[1].list_kind.as_deref(), ps[1].list_level), (Some("bullet"), Some(1)));

        m.set_list(&[0, 1], "none").unwrap();
        let ps = m.extract_paragraphs().unwrap();
        assert!(ps.iter().all(|p| p.num_id.is_none() && p.list_label.is_none()));
        assert!(!m.get_file_string(DOCUMENT).unwrap().contains("numPr"));
    }

    #[test]
    fn test_undo_restores_numbering() {
        let mut m = doc(&["a"]);
        m.checkpoint(None);
        m.set_list(&[0], "bullet").unwrap();
        assert!(m.undo().is_some());
        assert!(m.get_file_string(NUMBERING).is_err());
        assert_eq!(m.extract_paragraphs().unwrap()[0].list_kind, None);
        assert!(m.redo().is_some());
        assert_eq!(m.extract_paragraphs().unwrap()[0].list_kind.as_deref(), Some("bullet"));
    }

    #[test]
    fn test_clear_formatting_removes_direct_formatting_in_range() {
        let mut m = doc(&["uno dos tres"]);
        let mut run = m.extract_paragraphs().unwrap()[0].runs[0].clone();
        run.bold = true;
        run.strike = true;
        run.highlight = Some("FFFF00".into());
        run.font_size = Some(20.0);
        m.update_paragraph_runs(0, &[run], None).unwrap();
        m.clear_formatting(0, 4, 0, 7).unwrap();
        let p = &m.extract_paragraphs().unwrap()[0];
        let got: Vec<(&str, bool, bool, Option<&str>)> =
            p.runs.iter().map(|r| (r.text.as_str(), r.bold, r.strike, r.highlight.as_deref())).collect();
        assert_eq!(
            got,
            vec![("uno ", true, true, Some("FFFF00")), ("dos", false, false, None), (" tres", true, true, Some("FFFF00"))]
        );
        // The style's size shows again
        assert_ne!(p.runs[1].font_size, Some(20.0));
        assert!(!m.get_file_string(DOCUMENT).unwrap().contains(r#"<w:b w:val="0"/>"#));
    }

    #[test]
    fn test_indent_steps_on_half_inch_grid() {
        let mut m = doc(&["a"]);
        m.change_indent(&[0], 1).unwrap();
        assert_eq!(m.extract_paragraphs().unwrap()[0].indent_left, 48.0);
        m.change_indent(&[0], 2).unwrap();
        assert_eq!(m.extract_paragraphs().unwrap()[0].indent_left, 144.0);
        m.change_indent(&[0], -5).unwrap();
        assert_eq!(m.extract_paragraphs().unwrap()[0].indent_left, 0.0);
    }

    #[test]
    fn test_line_spacing_keeps_space_after() {
        let mut m = doc(&["a"]);
        m.set_paragraph_spacing(&[0], None, Some(6.0)).unwrap();
        m.set_line_spacing(&[0], 1.5).unwrap();
        let p = &m.extract_paragraphs().unwrap()[0];
        assert_eq!(p.line_spacing, Some(1.5));
        assert_eq!(p.space_after, 6.0);
        assert!(m.get_file_string(DOCUMENT).unwrap().contains(r#"<w:spacing w:after="120" w:line="360" w:lineRule="auto"/>"#));
    }

    #[test]
    fn test_paragraph_styles_gallery_and_apply() {
        let mut m = doc(&["Capítulo", "texto"]);
        let gallery = m.paragraph_styles();
        assert_eq!(gallery[0].builtin.as_deref(), Some("normal"));
        assert!(gallery.iter().any(|s| s.id == "Heading1"));

        m.set_paragraph_style(&[0], "Heading1").unwrap();
        let ps = m.extract_paragraphs().unwrap();
        assert!(ps[0].is_heading);
        assert_eq!(ps[0].outline_level, Some(0));
        // Now defined by the document, it is listed once
        assert_eq!(m.paragraph_styles().iter().filter(|s| s.builtin.as_deref() == Some("heading 1")).count(), 1);

        m.set_paragraph_style(&[0], &gallery[0].id).unwrap();
        assert!(!m.extract_paragraphs().unwrap()[0].is_heading);
        assert!(m.set_paragraph_style(&[0], "NoExiste").is_err());
    }
}
