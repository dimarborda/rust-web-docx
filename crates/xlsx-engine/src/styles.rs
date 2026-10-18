//! `xl/styles.xml`: which cell formats show dates, and adding a date format based on a cell's
//! current one so that a date written there keeps its font, fill and borders

use std::collections::HashMap;

use quick_xml::events::Event;
use quick_xml::reader::Reader;

use ooxml_core::xml::{element_ranges, get_attr_value, set_attr, tag_is};

use crate::dates::{is_builtin_date_format, is_date_format_code, DATE_FORMAT_ID, DATE_TIME_FORMAT_ID};

/// Number formats of the cell formats (`<cellXfs>`) of a stylesheet
#[derive(Debug, Default)]
pub struct Styles {
    /// Custom number formats (`<numFmt>`) by id
    custom_formats: HashMap<u32, String>,
    /// `numFmtId` of each cell format, by index (a cell's `s` attribute)
    xf_formats: Vec<u32>,
}

impl Styles {
    pub fn parse(xml: &str) -> Result<Styles, String> {
        let mut styles = Styles::default();
        let mut reader = Reader::from_str(xml);
        let mut in_cell_xfs = false;
        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let name = e.name();
                    if tag_is(name.as_ref(), "numFmt") {
                        let id = get_attr_value(&e, "numFmtId").and_then(|v| v.parse().ok());
                        if let (Some(id), Some(code)) = (id, get_attr_value(&e, "formatCode")) {
                            styles.custom_formats.insert(id, code);
                        }
                    } else if tag_is(name.as_ref(), "cellXfs") {
                        in_cell_xfs = true;
                    } else if in_cell_xfs && tag_is(name.as_ref(), "xf") {
                        styles.xf_formats.push(get_attr_value(&e, "numFmtId").and_then(|v| v.parse().ok()).unwrap_or(0));
                    }
                }
                Ok(Event::End(e)) if tag_is(e.name().as_ref(), "cellXfs") => in_cell_xfs = false,
                Ok(Event::Eof) => break,
                Err(e) => return Err(format!("La hoja de estilos no es XML válido: {}", e)),
                _ => {}
            }
        }
        Ok(styles)
    }

    /// Whether the number format id shows a date or a time
    pub fn is_date_format(&self, id: u32) -> bool {
        match self.custom_formats.get(&id) {
            Some(code) => is_date_format_code(code),
            None => is_builtin_date_format(id),
        }
    }

    /// Whether cells with the cell format `xf` (their `s`) show dates
    pub fn is_date_style(&self, xf: u32) -> bool {
        self.xf_formats.get(xf as usize).is_some_and(|id| self.is_date_format(*id))
    }
}

/// A cell format showing dates (or dates and times) that looks like the cell format `base`: the
/// index of `base` itself when it already shows dates, of an identical existing format, or of a
/// copy of `base` with a date number format appended to `<cellXfs>`. Returns the stylesheet
/// when it had to change.
pub fn date_style(xml: &str, base: u32, with_time: bool) -> Result<(u32, Option<String>), String> {
    let styles = Styles::parse(xml)?;
    if styles.is_date_style(base) {
        return Ok((base, None));
    }
    let cell_xfs = element_ranges(xml, "cellXfs", |_| true)?
        .into_iter()
        .next()
        .ok_or("La hoja de estilos no tiene formatos de celda (<cellXfs>).")?;
    let section = &xml[cell_xfs.clone()];
    let xfs = element_ranges(section, "xf", |_| true)?;
    let format_id = if with_time { DATE_TIME_FORMAT_ID } else { DATE_FORMAT_ID };

    let base_xf = xfs
        .get(base as usize)
        .or(xfs.first())
        .map(|r| &section[r.clone()])
        .unwrap_or(r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>"#);
    let wanted = set_attr(&set_attr(base_xf, "numFmtId", &format_id.to_string()), "applyNumberFormat", "1");
    if let Some(existing) = xfs.iter().position(|r| section[r.clone()] == wanted) {
        return Ok((existing as u32, None));
    }

    let index = xfs.len() as u32;
    let insert_at = xfs.last().map(|r| r.end).ok_or("La hoja de estilos no tiene formatos de celda (<xf>).")?;
    let mut new_section = format!("{}{}{}", &section[..insert_at], wanted, &section[insert_at..]);
    new_section = set_attr(&new_section, "count", &(index + 1).to_string());
    Ok((index, Some(format!("{}{}{}", &xml[..cell_xfs.start], new_section, &xml[cell_xfs.end..]))))
}

#[cfg(test)]
mod tests {
    use super::*;

    const STYLES: &str = r#"<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts><cellStyleXfs count="1"><xf numFmtId="14"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" xfId="0"/><xf numFmtId="164" fontId="1" xfId="0" applyNumberFormat="1"/><xf numFmtId="4" fontId="2" xfId="0"><alignment horizontal="left"/></xf></cellXfs></styleSheet>"#;

    #[test]
    fn knows_which_styles_show_dates() {
        let styles = Styles::parse(STYLES).unwrap();
        assert!(!styles.is_date_style(0), "cellStyleXfs are not cell formats");
        assert!(styles.is_date_style(1));
        assert!(!styles.is_date_style(2));
        assert!(!styles.is_date_style(9));
    }

    #[test]
    fn adds_a_date_style_once_keeping_the_rest_of_the_format() {
        assert_eq!(date_style(STYLES, 1, false).unwrap(), (1, None));
        let (index, xml) = date_style(STYLES, 2, true).unwrap();
        let xml = xml.unwrap();
        assert_eq!(index, 3);
        assert!(xml.contains(r#"<cellXfs count="4">"#), "{}", xml);
        assert!(xml.contains(r#"<xf numFmtId="22" fontId="2" xfId="0" applyNumberFormat="1"><alignment horizontal="left"/></xf></cellXfs>"#), "{}", xml);
        assert!(Styles::parse(&xml).unwrap().is_date_style(3));
        assert_eq!(date_style(&xml, 2, true).unwrap(), (3, None), "reuses the format it added");
    }
}
