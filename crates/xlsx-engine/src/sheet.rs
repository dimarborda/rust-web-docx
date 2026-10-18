//! A worksheet part (`xl/worksheets/sheetN.xml`): its cells read with their byte positions, and
//! edits spliced into the original text so that rows, cells and everything else not edited stay
//! exactly as written

use std::collections::BTreeMap;
use std::ops::Range;

use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;

use ooxml_core::xml::{get_attr_value, prefix_of, remove_attr, set_attr, tag_is, utf8};

use crate::cell_ref::{CellRange, CellRef};
use crate::text::decode;

/// The formula of a cell
#[derive(Debug, Clone, PartialEq)]
pub struct Formula {
    /// Formula text without `=`; empty in cells that reuse a shared formula from another cell
    pub text: String,
    /// The cell holds a shared formula that other cells reuse (`<f t="shared" ref="…">`)
    pub shared_master: bool,
}

/// One `<c>` element
#[derive(Debug, Clone)]
pub struct CellNode {
    pub at: CellRef,
    pub range: Range<usize>,
    /// Whether the element has its `r` attribute (cells without it follow the previous one)
    pub has_ref: bool,
    /// Value type (`t`): `s`, `str`, `inlineStr`, `b`, `e`, `d`, `n` or none
    pub kind: Option<String>,
    /// Cell format index (`s`)
    pub style: u32,
    /// Raw `s` attribute, kept as written when the style does not change
    pub style_attr: Option<String>,
    /// Phonetic guide flag (`ph`)
    pub phonetic: Option<String>,
    /// Content of `<v>`
    pub value: Option<String>,
    pub formula: Option<Formula>,
    /// Text of an inline string (`<is>`)
    pub inline: Option<String>,
}

#[derive(Debug, Clone)]
pub struct RowNode {
    pub number: u32,
    pub range: Range<usize>,
    pub start_tag: Range<usize>,
    pub has_ref: bool,
    pub self_closing: bool,
    pub cells: Vec<CellNode>,
}

#[derive(Debug, Clone)]
enum DataSpan {
    /// `<sheetData>…</sheetData>`: range of the content between the tags
    Open(Range<usize>),
    /// `<sheetData/>`
    Empty(Range<usize>),
}

/// A worksheet read for cell access
#[derive(Debug)]
pub struct SheetXml<'a> {
    pub xml: &'a str,
    /// Namespace prefix of the SpreadsheetML elements (usually empty)
    prefix: String,
    data: DataSpan,
    pub rows: Vec<RowNode>,
    /// The `<dimension>` element (used range), if any
    dimension: Option<Range<usize>>,
}

fn attr_u32(e: &BytesStart, name: &str) -> Option<u32> {
    get_attr_value(e, name).and_then(|v| v.trim().parse().ok())
}

impl<'a> SheetXml<'a> {
    pub fn parse(xml: &'a str) -> Result<SheetXml<'a>, String> {
        let invalid = |e: quick_xml::Error| format!("La hoja no es XML válido: {}", e);
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(false);
        let mut prefix = String::new();
        let mut data = None;
        let mut data_start = 0;
        let mut dimension = None;
        let mut rows: Vec<RowNode> = Vec::new();
        let mut cell: Option<CellNode> = None;
        // Open element inside a cell whose raw content is being read: (name, content start)
        let mut reading: Option<(&'static str, usize)> = None;
        let mut in_inline = false;
        let mut phonetic_depth = 0usize;
        let mut inline_text = String::new();

        loop {
            let before = reader.buffer_position() as usize;
            let event = reader.read_event().map_err(invalid)?;
            let after = reader.buffer_position() as usize;
            match event {
                Event::Eof => break,
                Event::Start(e) | Event::Empty(e) if data.is_none() && cell.is_none() => {
                    let empty = matches!(xml.as_bytes().get(after.wrapping_sub(2)), Some(b'/'));
                    let name = e.name();
                    let name = name.as_ref();
                    if tag_is(name, "sheetData") {
                        prefix = prefix_of(utf8(name)).to_string();
                        if empty {
                            data = Some(DataSpan::Empty(before..after));
                        } else {
                            data_start = after;
                            data = Some(DataSpan::Open(after..after));
                        }
                    } else if tag_is(name, "dimension") && rows.is_empty() {
                        dimension = Some(before..after);
                    }
                }
                Event::End(e) if tag_is(e.name().as_ref(), "sheetData") && cell.is_none() => {
                    data = Some(DataSpan::Open(data_start..before));
                    // Nothing after sheetData holds cells
                    break;
                }
                Event::Start(e) | Event::Empty(e) if cell.is_none() && tag_is(e.name().as_ref(), "row") => {
                    let self_closing = xml[before..after].ends_with("/>");
                    let previous = rows.last().map(|r| r.number).unwrap_or(0);
                    let number = attr_u32(&e, "r");
                    rows.push(RowNode {
                        number: number.unwrap_or(previous + 1),
                        range: before..after,
                        start_tag: before..after,
                        has_ref: number.is_some(),
                        self_closing,
                        cells: Vec::new(),
                    });
                }
                Event::End(e) if cell.is_none() && tag_is(e.name().as_ref(), "row") => {
                    if let Some(row) = rows.last_mut() {
                        row.range.end = after;
                    }
                }
                Event::Start(e) | Event::Empty(e) if cell.is_none() && tag_is(e.name().as_ref(), "c") => {
                    let Some(row) = rows.last_mut() else { continue };
                    let previous_col = row.cells.last().map(|c| c.at.col).unwrap_or(0);
                    let explicit = get_attr_value(&e, "r").and_then(|r| CellRef::parse(&r).ok());
                    let at = explicit.unwrap_or(CellRef { row: row.number, col: previous_col + 1 });
                    let node = CellNode {
                        at,
                        range: before..after,
                        has_ref: explicit.is_some(),
                        kind: get_attr_value(&e, "t"),
                        style: attr_u32(&e, "s").unwrap_or(0),
                        style_attr: get_attr_value(&e, "s"),
                        phonetic: get_attr_value(&e, "ph"),
                        value: None,
                        formula: None,
                        inline: None,
                    };
                    if xml[before..after].ends_with("/>") {
                        row.cells.push(node);
                    } else {
                        cell = Some(node);
                    }
                }
                Event::Start(e) if cell.is_some() => {
                    let name = e.name();
                    let name = name.as_ref();
                    let current = cell.as_mut().unwrap();
                    if tag_is(name, "v") {
                        reading = Some(("v", after));
                    } else if tag_is(name, "f") {
                        current.formula = Some(Formula {
                            text: String::new(),
                            shared_master: get_attr_value(&e, "t").as_deref() == Some("shared") && get_attr_value(&e, "ref").is_some(),
                        });
                        reading = Some(("f", after));
                    } else if tag_is(name, "is") {
                        in_inline = true;
                        inline_text.clear();
                    } else if tag_is(name, "rPh") {
                        phonetic_depth += 1;
                    } else if tag_is(name, "t") && in_inline && phonetic_depth == 0 {
                        reading = Some(("t", after));
                    }
                }
                Event::Empty(e) if cell.is_some() => {
                    let name = e.name();
                    let current = cell.as_mut().unwrap();
                    if tag_is(name.as_ref(), "f") {
                        current.formula = Some(Formula {
                            text: String::new(),
                            shared_master: get_attr_value(&e, "t").as_deref() == Some("shared") && get_attr_value(&e, "ref").is_some(),
                        });
                    } else if tag_is(name.as_ref(), "v") {
                        current.value = Some(String::new());
                    } else if tag_is(name.as_ref(), "is") {
                        current.inline = Some(String::new());
                    }
                }
                Event::End(e) if cell.is_some() => {
                    let name = e.name();
                    let name = name.as_ref();
                    let current = cell.as_mut().unwrap();
                    if let Some((open, start)) = reading {
                        if tag_is(name, open) {
                            let raw = &xml[start..before];
                            match open {
                                "v" => current.value = Some(decode(raw)),
                                "f" => current.formula.as_mut().unwrap().text = decode(raw),
                                _ => inline_text.push_str(&decode(raw)),
                            }
                            reading = None;
                        }
                    } else if tag_is(name, "rPh") {
                        phonetic_depth = phonetic_depth.saturating_sub(1);
                    } else if tag_is(name, "is") {
                        in_inline = false;
                        current.inline = Some(std::mem::take(&mut inline_text));
                    } else if tag_is(name, "c") {
                        let mut done = cell.take().unwrap();
                        done.range.end = after;
                        if let Some(row) = rows.last_mut() {
                            row.cells.push(done);
                        }
                    }
                }
                _ => {}
            }
        }
        let data = data.ok_or("La hoja no tiene <sheetData>.")?;
        Ok(SheetXml { xml, prefix, data, rows, dimension })
    }

    /// The cell at `at`, if the sheet has an element for it
    pub fn cell(&self, at: CellRef) -> Option<&CellNode> {
        let row = self.rows.iter().find(|r| r.number == at.row)?;
        row.cells.iter().find(|c| c.at == at)
    }

    pub fn prefix(&self) -> &str {
        &self.prefix
    }

    /// The sheet with the given cells replaced: each entry is the new `<c>` element, or `None`
    /// to remove the cell. Missing rows and cells are created in order; the rest of the text is
    /// copied as it was.
    pub fn splice(&self, edits: &BTreeMap<CellRef, Option<String>>) -> String {
        let xml = self.xml;
        let p = &self.prefix;
        let mut by_row: BTreeMap<u32, Vec<(u32, &Option<String>)>> = BTreeMap::new();
        for (at, content) in edits {
            by_row.entry(at.row).or_default().push((at.col, content));
        }

        let new_row = |number: u32, cells: &[(u32, &Option<String>)]| -> String {
            let body: String = cells.iter().filter_map(|(_, c)| c.as_deref()).collect();
            if body.is_empty() {
                String::new()
            } else {
                format!(r#"<{p}row r="{}">{}</{p}row>"#, number, body, p = p)
            }
        };

        let (inner, wrap) = match &self.data {
            DataSpan::Open(inner) => (inner.clone(), None),
            DataSpan::Empty(whole) => (whole.end..whole.end, Some(whole.clone())),
        };
        let mut body = String::new();
        let mut cursor = inner.start;
        let mut pending = by_row.iter().peekable();
        for row in &self.rows {
            while let Some((number, cells)) = pending.peek().filter(|(n, _)| **n < row.number) {
                body.push_str(&xml[cursor..row.range.start]);
                cursor = row.range.start;
                body.push_str(&new_row(**number, cells));
                pending.next();
            }
            let edits_here = pending.peek().filter(|(n, _)| **n == row.number).map(|(_, cells)| *cells);
            if edits_here.is_some() {
                pending.next();
            }
            if edits_here.is_none() && row.has_ref {
                continue;
            }
            body.push_str(&xml[cursor..row.range.start]);
            body.push_str(&self.rebuild_row(row, edits_here.map(Vec::as_slice).unwrap_or(&[])));
            cursor = row.range.end;
        }
        let tail_at = self.rows.last().map(|r| r.range.end.max(cursor)).unwrap_or(inner.start);
        body.push_str(&xml[cursor..tail_at]);
        cursor = tail_at;
        for (number, cells) in pending {
            body.push_str(&new_row(*number, cells));
        }
        body.push_str(&xml[cursor..inner.end]);

        let mut out = String::with_capacity(xml.len() + body.len());
        let head_end = wrap.as_ref().map(|w| w.start).unwrap_or(inner.start);
        let head = &xml[..head_end];
        match &self.dimension {
            Some(range) => {
                out.push_str(&head[..range.start]);
                out.push_str(&self.updated_dimension(&head[range.clone()], edits));
                out.push_str(&head[range.end..]);
            }
            None => out.push_str(head),
        }
        match wrap {
            Some(whole) if !body.is_empty() => {
                out.push_str(&format!("<{p}sheetData>{}</{p}sheetData>", body, p = p));
                out.push_str(&xml[whole.end..]);
            }
            Some(whole) => out.push_str(&xml[whole.start..]),
            None => {
                out.push_str(&body);
                out.push_str(&xml[inner.end..]);
            }
        }
        out
    }

    /// A row with its edited cells replaced and its new cells inserted in column order
    fn rebuild_row(&self, row: &RowNode, edits: &[(u32, &Option<String>)]) -> String {
        let xml = self.xml;
        let mut start_tag = xml[row.start_tag.clone()].to_string();
        if !row.has_ref {
            start_tag = set_attr(&start_tag, "r", &row.number.to_string());
        }
        let adds_cells = edits.iter().any(|(col, c)| c.is_some() && !row.cells.iter().any(|cell| cell.at.col == *col));
        if adds_cells {
            // `spans` is only an optimization hint; a stale one would be wrong
            start_tag = remove_attr(&start_tag, "spans");
        }
        if edits.is_empty() {
            return format!("{}{}", start_tag, &xml[row.start_tag.end..row.range.end]);
        }
        if row.self_closing {
            let new_cells: String = edits.iter().filter_map(|(_, c)| c.as_deref()).collect();
            if new_cells.is_empty() {
                return start_tag;
            }
            let open = format!("{}>", start_tag.trim_end_matches("/>").trim_end());
            return format!("{}{}</{p}row>", open, new_cells, p = self.prefix);
        }

        let mut out = start_tag;
        let mut cursor = row.start_tag.end;
        let mut pending = edits.iter().peekable();
        for cell in &row.cells {
            while let Some((_, content)) = pending.peek().filter(|(col, _)| *col < cell.at.col) {
                out.push_str(&xml[cursor..cell.range.start]);
                cursor = cell.range.start;
                out.push_str(content.as_deref().unwrap_or(""));
                pending.next();
            }
            out.push_str(&xml[cursor..cell.range.start]);
            cursor = cell.range.end;
            match pending.peek().filter(|(col, _)| *col == cell.at.col) {
                Some((_, content)) => {
                    out.push_str(content.as_deref().unwrap_or(""));
                    pending.next();
                }
                None if !cell.has_ref => out.push_str(&set_attr(&xml[cell.range.clone()], "r", &cell.at.to_string())),
                None => out.push_str(&xml[cell.range.clone()]),
            }
        }
        let tail_at = row.cells.last().map(|c| c.range.end).unwrap_or(row.start_tag.end);
        out.push_str(&xml[cursor..tail_at]);
        for (_, content) in pending {
            out.push_str(content.as_deref().unwrap_or(""));
        }
        out.push_str(&xml[tail_at..row.range.end]);
        out
    }

    /// The `<dimension>` element grown to hold the written cells
    fn updated_dimension(&self, element: &str, edits: &BTreeMap<CellRef, Option<String>>) -> String {
        let mut reader = Reader::from_str(element);
        let current = match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => get_attr_value(&e, "ref").and_then(|r| CellRange::parse(&r).ok()),
            _ => None,
        };
        let written = edits
            .iter()
            .filter(|(_, c)| c.is_some())
            .map(|(at, _)| CellRange::around(*at, *at))
            .reduce(CellRange::union);
        let grown = match (current, written) {
            (Some(current), Some(written)) => current.union(written),
            (None, Some(written)) if self.rows.is_empty() => written,
            _ => return element.to_string(),
        };
        if Some(grown) == current {
            element.to_string()
        } else {
            set_attr(element, "ref", &grown.to_string())
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const SHEET: &str = r#"<worksheet xmlns="m"><dimension ref="A1:C3"/><sheetData><row r="1" spans="1:3"><c r="A1" t="s"><v>0</v></c><c r="C1" s="2"><f>A2*2</f><v>84</v></c></row><row r="3"><c r="B3" t="inlineStr"><is><r><t>Hola </t></r><r><t>mundo</t></r><rPh><t>x</t></rPh></is></c><c r="C3" t="b"><v>1</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A5:B5"/></mergeCells></worksheet>"#;

    fn edits(pairs: &[(&str, Option<&str>)]) -> BTreeMap<CellRef, Option<String>> {
        pairs.iter().map(|(r, c)| (CellRef::parse(r).unwrap(), c.map(str::to_string))).collect()
    }

    #[test]
    fn reads_cells_with_their_content() {
        let sheet = SheetXml::parse(SHEET).unwrap();
        assert_eq!(sheet.rows.len(), 2);
        let c1 = sheet.cell(CellRef::parse("C1").unwrap()).unwrap();
        assert_eq!((c1.style, c1.value.as_deref()), (2, Some("84")));
        assert_eq!(c1.formula.as_ref().unwrap().text, "A2*2");
        let b3 = sheet.cell(CellRef::parse("B3").unwrap()).unwrap();
        assert_eq!(b3.inline.as_deref(), Some("Hola mundo"));
        assert!(sheet.cell(CellRef::parse("A2").unwrap()).is_none());
    }

    #[test]
    fn infers_positions_without_references() {
        let xml = r#"<worksheet><sheetData><row><c><v>1</v></c><c><v>2</v></c></row><row><c r="C2"><v>3</v></c><c><v>4</v></c></row></sheetData></worksheet>"#;
        let sheet = SheetXml::parse(xml).unwrap();
        assert_eq!(sheet.cell(CellRef::parse("B1").unwrap()).unwrap().value.as_deref(), Some("2"));
        assert_eq!(sheet.cell(CellRef::parse("D2").unwrap()).unwrap().value.as_deref(), Some("4"));

        let out = sheet.splice(&edits(&[("A1", Some(r#"<c r="A1"><v>9</v></c>"#))]));
        assert_eq!(out, r#"<worksheet><sheetData><row r="1"><c r="A1"><v>9</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="C2"><v>3</v></c><c><v>4</v></c></row></sheetData></worksheet>"#);
    }

    #[test]
    fn splices_cells_and_rows_in_order() {
        let sheet = SheetXml::parse(SHEET).unwrap();
        let out = sheet.splice(&edits(&[
            ("B1", Some(r#"<c r="B1"><v>5</v></c>"#)),
            ("C1", Some(r#"<c r="C1" s="2"><v>7</v></c>"#)),
            ("A2", Some(r#"<c r="A2"><v>1</v></c>"#)),
            ("C3", None),
            ("D3", Some(r#"<c r="D3"><v>2</v></c>"#)),
            ("A4", None),
            ("E9", Some(r#"<c r="E9"><v>3</v></c>"#)),
        ]));
        assert_eq!(
            out,
            r#"<worksheet xmlns="m"><dimension ref="A1:E9"/><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>5</v></c><c r="C1" s="2"><v>7</v></c></row><row r="2"><c r="A2"><v>1</v></c></row><row r="3"><c r="B3" t="inlineStr"><is><r><t>Hola </t></r><r><t>mundo</t></r><rPh><t>x</t></rPh></is></c><c r="D3"><v>2</v></c></row><row r="9"><c r="E9"><v>3</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A5:B5"/></mergeCells></worksheet>"#
        );
        assert_eq!(sheet.splice(&BTreeMap::new()), SHEET, "no edits, same text");
    }

    #[test]
    fn writes_into_empty_sheets_and_rows() {
        let empty = r#"<x:worksheet xmlns:x="m"><x:sheetData/></x:worksheet>"#;
        let sheet = SheetXml::parse(empty).unwrap();
        assert_eq!(
            sheet.splice(&edits(&[("B2", Some(r#"<x:c r="B2"><x:v>1</x:v></x:c>"#))])),
            r#"<x:worksheet xmlns:x="m"><x:sheetData><x:row r="2"><x:c r="B2"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>"#
        );
        assert_eq!(sheet.splice(&edits(&[("B2", None)])), empty);

        let blank_row = r#"<worksheet><sheetData><row r="2" ht="30" customHeight="1"/></sheetData></worksheet>"#;
        let sheet = SheetXml::parse(blank_row).unwrap();
        assert_eq!(
            sheet.splice(&edits(&[("A2", Some(r#"<c r="A2"><v>1</v></c>"#))])),
            r#"<worksheet><sheetData><row r="2" ht="30" customHeight="1"><c r="A2"><v>1</v></c></row></sheetData></worksheet>"#
        );
    }
}
