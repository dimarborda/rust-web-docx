//! Lossless paragraph editing.
//!
//! The editor only models text plus a handful of run properties (bold, italic,
//! underline, color, size, font). Instead of regenerating a paragraph from that
//! reduced model, an edit is applied as a minimal diff on the original XML:
//! untouched runs are copied byte-for-byte, changed runs keep their full `w:rPr`
//! (highlight, rStyle, lang, ...), and everything that is not text — hyperlinks,
//! fields, bookmarks, comments, drawings — stays exactly where it was.

use crate::docx_parser::{parse_paragraph_with, tag_is, ParagraphInfo, RunInfo};
use crate::styles::StyleSheet;
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::ops::Range;

/// Requested formatting for one character. `None` keeps the run's current value.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct FormatTarget {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub color: Option<String>,
    pub font_size: Option<f64>,
    pub font_family: Option<String>,
}

impl FormatTarget {
    pub fn from_run(run: &RunInfo) -> Self {
        FormatTarget {
            bold: Some(run.bold),
            italic: Some(run.italic),
            underline: Some(run.underline),
            color: Some(run.color.clone()),
            font_size: run.font_size,
            font_family: run.font_family.clone(),
        }
    }

    /// Keeps only the properties that differ from `base`; `None` when nothing changes
    fn diff_against(&self, base: &RunInfo) -> Option<FormatTarget> {
        let mut d = FormatTarget::default();
        if self.bold.is_some_and(|b| b != base.bold) {
            d.bold = self.bold;
        }
        if self.italic.is_some_and(|i| i != base.italic) {
            d.italic = self.italic;
        }
        if self.underline.is_some_and(|u| u != base.underline) {
            d.underline = self.underline;
        }
        if let Some(c) = &self.color {
            if normalize_color(c) != normalize_color(&base.color) {
                d.color = Some(normalize_color(c));
            }
        }
        if let Some(sz) = self.font_size {
            if base.font_size.map_or(true, |b| (b - sz).abs() > 0.01) {
                d.font_size = Some(sz);
            }
        }
        if let Some(fam) = &self.font_family {
            if base.font_family.as_deref().map_or(true, |b| !b.eq_ignore_ascii_case(fam)) {
                d.font_family = Some(fam.clone());
            }
        }
        if d == FormatTarget::default() {
            None
        } else {
            Some(d)
        }
    }
}

fn normalize_color(c: &str) -> String {
    let c = c.trim().trim_start_matches('#').to_uppercase();
    if c == "AUTO" {
        String::new()
    } else {
        c
    }
}

fn normalize_align(a: &str) -> &'static str {
    match a {
        "center" => "center",
        "right" | "end" => "right",
        "both" | "justify" | "distribute" => "both",
        _ => "left",
    }
}

/// Canonical child order of `w:rPr` (CT_RPr). Word rejects files with out-of-order children.
const RPR_ORDER: &[&str] = &[
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
    "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
    "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
    "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath", "rPrChange",
];

/// Canonical child order of `w:pPr` (CT_PPr)
const PPR_ORDER: &[&str] = &[
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
    "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
    "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap",
    "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId",
    "cnfStyle", "rPr", "sectPr", "pPrChange",
];

struct Token<'a> {
    ev: Event<'a>,
    span: Range<usize>,
}

fn tokenize(xml: &str) -> Result<Vec<Token<'_>>, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut tokens = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event() {
            Ok(Event::Eof) => break,
            Ok(ev) => {
                let end = reader.buffer_position() as usize;
                tokens.push(Token { ev, span: start..end });
            }
            Err(e) => return Err(format!("XML error: {:?}", e)),
        }
    }
    Ok(tokens)
}

/// Index of the token closing the element opened at `i` (`i` itself for empty elements)
fn element_end(tokens: &[Token], i: usize) -> usize {
    if !matches!(tokens[i].ev, Event::Start(_)) {
        return i;
    }
    let mut depth = 0usize;
    for (j, tok) in tokens.iter().enumerate().skip(i) {
        match tok.ev {
            Event::Start(_) => depth += 1,
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    return j;
                }
            }
            _ => {}
        }
    }
    tokens.len() - 1
}

/// Qualified name of an element token, if it is one
fn element_name<'t>(tok: &'t Token) -> Option<&'t [u8]> {
    match &tok.ev {
        Event::Start(e) | Event::Empty(e) => Some(e.name().into_inner()),
        _ => None,
    }
}

fn is_element(tok: &Token, local: &str) -> bool {
    element_name(tok).is_some_and(|n| tag_is(n, local))
}

fn local_name_of(name: &[u8]) -> &str {
    let s = std::str::from_utf8(name).unwrap_or("");
    s.rsplit(':').next().unwrap_or(s)
}

/// Namespace prefix of an element name ("w" for "w:r"), used to name generated elements
fn prefix_of(name: &[u8]) -> String {
    let s = std::str::from_utf8(name).unwrap_or("");
    s.split_once(':').map(|(p, _)| p.to_string()).unwrap_or_default()
}

fn qualified(prefix: &str, local: &str) -> String {
    if prefix.is_empty() {
        local.to_string()
    } else {
        format!("{}:{}", prefix, local)
    }
}

fn escape_text(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;")
}

fn escape_attr(s: &str) -> String {
    escape_text(s).replace('"', "&quot;")
}

/// Byte ranges of the body paragraphs, indexed exactly like `parse_document_elements`
pub fn body_paragraph_ranges(xml: &str) -> Result<Vec<Range<usize>>, String> {
    let tokens = tokenize(xml)?;
    let mut ranges = Vec::new();
    let mut in_body = false;
    let mut i = 0;
    while i < tokens.len() {
        match &tokens[i].ev {
            Event::Start(e) if tag_is(e.name().as_ref(), "body") => in_body = true,
            Event::End(e) if tag_is(e.name().as_ref(), "body") => in_body = false,
            Event::Start(e) if in_body && tag_is(e.name().as_ref(), "tbl") => {
                i = element_end(&tokens, i) + 1;
                continue;
            }
            Event::Start(e) | Event::Empty(e) if in_body && tag_is(e.name().as_ref(), "p") => {
                let end = element_end(&tokens, i);
                ranges.push(tokens[i].span.start..tokens[end].span.end);
                i = end + 1;
                continue;
            }
            _ => {}
        }
        i += 1;
    }
    Ok(ranges)
}

/// Byte ranges of the paragraphs inside cell (row, col) of the `table_index`-th body table,
/// with the same table/row/cell numbering as `parse_table_from_reader`
pub fn table_cell_paragraph_ranges(
    xml: &str,
    table_index: usize,
    row: usize,
    col: usize,
) -> Result<Vec<Range<usize>>, String> {
    let tokens = tokenize(xml)?;
    let not_found = || format!("No se encontró la celda ({}, {}) de la tabla {}", row, col, table_index);

    // 1. Locate the top-level body table
    let mut in_body = false;
    let mut tbl_count = 0;
    let mut table = None;
    let mut i = 0;
    while i < tokens.len() {
        match &tokens[i].ev {
            Event::Start(e) if tag_is(e.name().as_ref(), "body") => in_body = true,
            Event::Start(e) if in_body && tag_is(e.name().as_ref(), "p") => {
                i = element_end(&tokens, i) + 1;
                continue;
            }
            Event::Start(e) if in_body && tag_is(e.name().as_ref(), "tbl") => {
                let end = element_end(&tokens, i);
                if tbl_count == table_index {
                    table = Some((i, end));
                    break;
                }
                tbl_count += 1;
                i = end + 1;
                continue;
            }
            _ => {}
        }
        i += 1;
    }
    let (tbl_start, tbl_end) = table.ok_or_else(not_found)?;

    // 2. Locate the row and cell
    let mut row_idx = 0;
    let mut i = tbl_start + 1;
    while i < tbl_end {
        if matches!(tokens[i].ev, Event::Start(_)) && is_element(&tokens[i], "tr") {
            let tr_end = element_end(&tokens, i);
            if row_idx == row {
                let mut col_idx = 0;
                let mut j = i + 1;
                while j < tr_end {
                    if is_element(&tokens[j], "tc") {
                        let tc_end = element_end(&tokens, j);
                        if col_idx == col {
                            return cell_paragraphs(&tokens, j, tc_end);
                        }
                        col_idx += 1;
                        j = tc_end + 1;
                        continue;
                    }
                    j += 1;
                }
                return Err(not_found());
            }
            row_idx += 1;
            i = tr_end + 1;
            continue;
        }
        i += 1;
    }
    Err(not_found())
}

fn cell_paragraphs(tokens: &[Token], tc_start: usize, tc_end: usize) -> Result<Vec<Range<usize>>, String> {
    let mut ranges = Vec::new();
    let mut j = tc_start + 1;
    while j < tc_end {
        if is_element(&tokens[j], "tbl") {
            return Err("Las celdas con tablas anidadas aún no se pueden editar.".to_string());
        }
        if is_element(&tokens[j], "p") {
            let end = element_end(tokens, j);
            ranges.push(tokens[j].span.start..tokens[end].span.end);
            j = end + 1;
            continue;
        }
        j += 1;
    }
    if ranges.is_empty() {
        return Err("La celda no contiene párrafos.".to_string());
    }
    Ok(ranges)
}

/// Parses a single `<w:p>` fragment with the regular paragraph parser
pub fn parse_paragraph_fragment(p_xml: &str, styles: &StyleSheet) -> ParagraphInfo {
    let mut reader = Reader::from_str(p_xml);
    reader.config_mut().trim_text(false);
    match reader.read_event() {
        Ok(Event::Start(_)) => parse_paragraph_with(&mut reader, 0, styles, None),
        _ => ParagraphInfo {
            style: "Normal".to_string(),
            align: "left".to_string(),
            ..Default::default()
        },
    }
}

enum AtomKind {
    Text(String),
    Tab,
    Break,
    /// Zero-width content kept verbatim (drawing, fldChar, footnote reference, ...)
    Opaque,
}

/// A direct child of a run, spanning tokens `first..=last`
struct Atom {
    kind: AtomKind,
    first: usize,
    last: usize,
}

struct RunModel {
    start: usize,
    end: usize,
    rpr: Option<(usize, usize)>,
    atoms: Vec<Atom>,
    /// Formatting as reported by the parser (what the editor was shown)
    info: Option<RunInfo>,
}

impl RunModel {
    fn text(&self) -> String {
        let mut s = String::new();
        for atom in &self.atoms {
            match &atom.kind {
                AtomKind::Text(t) => s.push_str(t),
                AtomKind::Tab => s.push('\t'),
                AtomKind::Break => s.push('\n'),
                AtomKind::Opaque => {}
            }
        }
        s
    }
}

fn parse_run(tokens: &[Token], start: usize, end: usize) -> Result<RunModel, String> {
    let mut run = RunModel { start, end, rpr: None, atoms: Vec::new(), info: None };
    let mut j = start + 1;
    while j < end {
        let child_end = element_end(tokens, j);
        if let Some(name) = element_name(&tokens[j]) {
            let kind = if tag_is(name, "rPr") {
                run.rpr = Some((j, child_end));
                None
            } else if tag_is(name, "t") {
                let mut text = String::new();
                for tok in &tokens[j + 1..child_end] {
                    match &tok.ev {
                        Event::Text(t) => text.push_str(&t.unescape().map_err(|e| e.to_string())?),
                        Event::CData(c) => text.push_str(&String::from_utf8_lossy(c)),
                        _ => {}
                    }
                }
                Some(AtomKind::Text(text))
            } else if tag_is(name, "tab") {
                Some(AtomKind::Tab)
            } else if tag_is(name, "br") || tag_is(name, "cr") {
                Some(AtomKind::Break)
            } else {
                Some(AtomKind::Opaque)
            };
            if let Some(kind) = kind {
                run.atoms.push(Atom { kind, first: j, last: child_end });
            }
        }
        j = child_end + 1;
    }
    Ok(run)
}

/// One piece of a rebuilt run
enum Item {
    /// A character; `atom` holds the original tokens of a kept tab/break
    Char { ch: char, atom: Option<(usize, usize)>, format: Option<FormatTarget> },
    Opaque { first: usize, last: usize },
}

/// Applies `new_text` (and optionally per-character formatting and paragraph alignment)
/// to the paragraph fragment `p_xml`, returning the edited fragment.
pub fn edit_paragraph(
    p_xml: &str,
    styles: &StyleSheet,
    new_text: &str,
    formats: Option<&[FormatTarget]>,
    align: Option<&str>,
) -> Result<String, String> {
    let tokens = tokenize(p_xml)?;
    let new_chars: Vec<char> = new_text.chars().collect();
    if let Some(f) = formats {
        if f.len() != new_chars.len() {
            return Err("El número de formatos no coincide con el texto.".to_string());
        }
    }
    let p_is_empty_tag = match tokens.first().map(|t| &t.ev) {
        Some(Event::Start(e)) if tag_is(e.name().as_ref(), "p") => false,
        Some(Event::Empty(e)) if tag_is(e.name().as_ref(), "p") => true,
        _ => return Err("El fragmento no es un párrafo.".to_string()),
    };
    let p_name = element_name(&tokens[0]).unwrap_or(b"w:p");
    let p_prefix = prefix_of(p_name);
    let content_end = if p_is_empty_tag { 1 } else { tokens.len() - 1 };

    // 1. Build the run model
    let mut runs: Vec<RunModel> = Vec::new();
    let mut ppr: Option<(usize, usize)> = None;
    let mut i = 1;
    while i < content_end {
        if let Some(name) = element_name(&tokens[i]) {
            let end = element_end(&tokens, i);
            if tag_is(name, "pPr") && ppr.is_none() && runs.is_empty() {
                ppr = Some((i, end));
                i = end + 1;
                continue;
            }
            if tag_is(name, "r") {
                runs.push(parse_run(&tokens, i, end)?);
                i = end + 1;
                continue;
            }
        }
        i += 1;
    }

    // 2. Attach the parser's view of each run; both must agree on the text
    let info = parse_paragraph_fragment(p_xml, styles);
    let mut parsed_runs = info.runs.iter();
    for run in runs.iter_mut() {
        let text = run.text();
        if text.is_empty() {
            continue;
        }
        match parsed_runs.next() {
            Some(pr) if pr.text == text => run.info = Some(pr.clone()),
            _ => return Err(unsupported()),
        }
    }
    if parsed_runs.next().is_some() {
        return Err(unsupported());
    }

    // 3. Character-level diff (common prefix/suffix)
    let mut owners: Vec<usize> = Vec::new();
    for (r, run) in runs.iter().enumerate() {
        owners.extend(std::iter::repeat(r).take(run.text().chars().count()));
    }
    let old_chars: Vec<char> = runs.iter().flat_map(|r| r.text().chars().collect::<Vec<_>>()).collect();
    let (old_len, new_len) = (old_chars.len(), new_chars.len());
    let mut prefix = 0;
    while prefix < old_len && prefix < new_len && old_chars[prefix] == new_chars[prefix] {
        prefix += 1;
    }
    let mut suffix = 0;
    while suffix < old_len - prefix
        && suffix < new_len - prefix
        && old_chars[old_len - 1 - suffix] == new_chars[new_len - 1 - suffix]
    {
        suffix += 1;
    }
    let deleted = prefix..old_len - suffix;
    let inserted = prefix..new_len - suffix;
    let new_pos = |k: usize| if k < prefix { k } else { k + new_len - old_len };
    let insert_owner = if inserted.is_empty() {
        None
    } else if prefix > 0 {
        Some(owners[prefix - 1])
    } else {
        owners.first().copied()
    };

    let format_for = |pos: usize, run: Option<&RunInfo>| -> Option<FormatTarget> {
        let target = formats?.get(pos)?;
        match run {
            Some(base) => target.diff_against(base),
            None => target.diff_against(&paragraph_base_format(&info)),
        }
    };
    let insert_items = |run: Option<&RunInfo>| -> Vec<Item> {
        inserted
            .clone()
            .map(|p| Item::Char { ch: new_chars[p], atom: None, format: format_for(p, run) })
            .collect()
    };

    // 4. Distribute kept, deleted and inserted characters over the runs
    let mut items: Vec<Vec<Item>> = (0..runs.len()).map(|_| Vec::new()).collect();
    let mut dirty = vec![false; runs.len()];
    let mut k = 0usize;
    for (r, run) in runs.iter().enumerate() {
        let info = run.info.as_ref();
        for atom in &run.atoms {
            let chars: Vec<(char, Option<(usize, usize)>)> = match &atom.kind {
                AtomKind::Opaque => {
                    items[r].push(Item::Opaque { first: atom.first, last: atom.last });
                    continue;
                }
                AtomKind::Text(t) => t.chars().map(|c| (c, None)).collect(),
                AtomKind::Tab => vec![('\t', Some((atom.first, atom.last)))],
                AtomKind::Break => vec![('\n', Some((atom.first, atom.last)))],
            };
            for (ch, original) in chars {
                if prefix == 0 && k == 0 && insert_owner == Some(r) {
                    items[r].extend(insert_items(info));
                    dirty[r] = true;
                }
                if deleted.contains(&k) {
                    dirty[r] = true;
                } else {
                    let format = format_for(new_pos(k), info);
                    dirty[r] |= format.is_some();
                    items[r].push(Item::Char { ch, atom: original, format });
                }
                if prefix > 0 && k + 1 == prefix && insert_owner == Some(r) {
                    items[r].extend(insert_items(info));
                    dirty[r] = true;
                }
                k += 1;
            }
        }
    }

    let new_align = align
        .map(normalize_align)
        .filter(|a| *a != normalize_align(&info.align));

    if !dirty.contains(&true) && new_align.is_none() && (inserted.is_empty() || insert_owner.is_some()) {
        return Ok(p_xml.to_string());
    }

    // 5. Emit the paragraph
    let raw = |a: usize, b: usize| &p_xml[tokens[a].span.start..tokens[b].span.end];
    let mut out = String::with_capacity(p_xml.len() + 128);
    if p_is_empty_tag {
        let tag = raw(0, 0);
        out.push_str(tag.trim_end_matches("/>").trim_end());
        out.push('>');
    } else {
        out.push_str(raw(0, 0));
    }

    if let (Some(a), None) = (new_align, ppr) {
        out.push_str(&format!(
            "<{ppr}><{jc} {val}=\"{a}\"/></{ppr}>",
            ppr = qualified(&p_prefix, "pPr"),
            jc = qualified(&p_prefix, "jc"),
            val = qualified(&p_prefix, "val"),
        ));
    }

    let mut next_run = 0;
    let mut i = 1;
    while i < content_end {
        if let (Some(a), Some((ps, pe))) = (new_align, ppr) {
            if i == ps {
                out.push_str(&rebuild_ppr_with_jc(&tokens, p_xml, ps, pe, a));
                i = pe + 1;
                continue;
            }
        }
        if next_run < runs.len() && runs[next_run].start == i {
            let run = &runs[next_run];
            if dirty[next_run] {
                emit_run(&mut out, &tokens, p_xml, run, &items[next_run]);
            } else {
                out.push_str(raw(run.start, run.end));
            }
            i = run.end + 1;
            next_run += 1;
            continue;
        }
        out.push_str(raw(i, i));
        i += 1;
    }

    // Paragraph without any text: inserted text goes into a new run that inherits the
    // paragraph mark formatting, like Word does when typing into an empty paragraph
    if !inserted.is_empty() && insert_owner.is_none() {
        let run_items = insert_items(None);
        let base_rpr = ppr.and_then(|(ps, pe)| paragraph_mark_rpr(&tokens, p_xml, ps, pe));
        emit_new_run(&mut out, &p_prefix, base_rpr, &run_items);
    }

    if p_is_empty_tag {
        out.push_str(&format!("</{}>", String::from_utf8_lossy(p_name)));
    } else {
        out.push_str(raw(tokens.len() - 1, tokens.len() - 1));
    }
    Ok(out)
}

fn unsupported() -> String {
    "El párrafo contiene contenido que el editor aún no soporta; no se modificó.".to_string()
}

/// Formatting a new run inherits in a paragraph that has no text yet
fn paragraph_base_format(info: &ParagraphInfo) -> RunInfo {
    RunInfo {
        text: String::new(),
        bold: info.bold,
        italic: info.italic,
        underline: false,
        color: info.color.clone(),
        font_size: info.font_size,
        font_family: info.font_family.clone(),
    }
}

/// Children of an element (`first..=last`) as (local name, raw xml), skipping whitespace
fn children(tokens: &[Token], xml: &str, first: usize, last: usize) -> Vec<(String, String)> {
    let mut out = Vec::new();
    if first == last {
        return out;
    }
    let mut j = first + 1;
    while j < last {
        let end = element_end(tokens, j);
        if let Some(name) = element_name(&tokens[j]) {
            out.push((
                local_name_of(name).to_string(),
                xml[tokens[j].span.start..tokens[end].span.end].to_string(),
            ));
        }
        j = end + 1;
    }
    out
}

/// Stable-sorts children into schema order; unknown elements stay after their predecessor
fn sort_children(children: &mut Vec<(String, String)>, order: &[&str]) {
    let mut last_known = 0;
    let mut keyed: Vec<(usize, (String, String))> = children
        .drain(..)
        .map(|c| {
            if let Some(pos) = order.iter().position(|o| *o == c.0) {
                last_known = pos;
            }
            (last_known, c)
        })
        .collect();
    keyed.sort_by_key(|(k, _)| *k);
    children.extend(keyed.into_iter().map(|(_, c)| c));
}

fn open_tag_of(tokens: &[Token], xml: &str, first: usize, last: usize) -> String {
    let tag = &xml[tokens[first].span.clone()];
    if first == last {
        format!("{}>", tag.trim_end_matches("/>").trim_end())
    } else {
        tag.to_string()
    }
}

fn rebuild_ppr_with_jc(tokens: &[Token], xml: &str, first: usize, last: usize, align: &str) -> String {
    let name = element_name(&tokens[first]).unwrap_or(b"w:pPr");
    let prefix = prefix_of(name);
    let mut kids = children(tokens, xml, first, last);
    kids.retain(|(n, _)| n != "jc");
    kids.push((
        "jc".to_string(),
        format!("<{} {}=\"{}\"/>", qualified(&prefix, "jc"), qualified(&prefix, "val"), align),
    ));
    sort_children(&mut kids, PPR_ORDER);
    let body: String = kids.into_iter().map(|(_, x)| x).collect();
    format!("{}{}</{}>", open_tag_of(tokens, xml, first, last), body, String::from_utf8_lossy(name))
}

/// The `w:rPr` inside `w:pPr` (paragraph mark formatting) as a list of children
fn paragraph_mark_rpr(tokens: &[Token], xml: &str, first: usize, last: usize) -> Option<Vec<(String, String)>> {
    let mut j = first + 1;
    while j < last {
        let end = element_end(tokens, j);
        if is_element(&tokens[j], "rPr") {
            let mut kids = children(tokens, xml, j, end);
            kids.retain(|(n, _)| n != "ins" && n != "del" && n != "rPrChange" && n != "moveFrom" && n != "moveTo");
            return Some(kids);
        }
        j = end + 1;
    }
    None
}

/// Applies a formatting diff to rPr children, keeping every property the editor doesn't model
fn patch_rpr(kids: &mut Vec<(String, String)>, prefix: &str, diff: &FormatTarget) {
    let el = |local: &str, attrs: &str| format!("<{}{}/>", qualified(prefix, local), attrs);
    let val = |v: &str| format!(" {}=\"{}\"", qualified(prefix, "val"), escape_attr(v));
    let set = |kids: &mut Vec<(String, String)>, local: &str, xml: Option<String>| {
        kids.retain(|(n, _)| n != local);
        if let Some(x) = xml {
            kids.push((local.to_string(), x));
        }
    };

    if let Some(b) = diff.bold {
        set(kids, "b", Some(if b { el("b", "") } else { el("b", &val("0")) }));
    }
    if let Some(i) = diff.italic {
        set(kids, "i", Some(if i { el("i", "") } else { el("i", &val("0")) }));
    }
    if let Some(u) = diff.underline {
        set(kids, "u", Some(el("u", &val(if u { "single" } else { "none" }))));
    }
    if let Some(c) = &diff.color {
        set(kids, "color", if c.is_empty() { None } else { Some(el("color", &val(c))) });
    }
    if let Some(sz) = diff.font_size {
        set(kids, "sz", Some(el("sz", &val(&((sz * 2.0).round() as i64).to_string()))));
    }
    if let Some(fam) = &diff.font_family {
        let f = escape_attr(fam);
        let attrs = format!(
            " {}=\"{f}\" {}=\"{f}\"",
            qualified(prefix, "ascii"),
            qualified(prefix, "hAnsi")
        );
        set(kids, "rFonts", Some(el("rFonts", &attrs)));
    }
    sort_children(kids, RPR_ORDER);
}

fn emit_run(out: &mut String, tokens: &[Token], xml: &str, run: &RunModel, items: &[Item]) {
    if items.is_empty() {
        // Every character of this run was deleted and it held nothing else
        return;
    }
    let run_name = element_name(&tokens[run.start]).unwrap_or(b"w:r");
    let prefix = prefix_of(run_name);
    let open = &xml[tokens[run.start].span.clone()];
    let close = format!("</{}>", String::from_utf8_lossy(run_name));
    let original_rpr = run.rpr.map(|(a, b)| &xml[tokens[a].span.start..tokens[b].span.end]);

    for group in group_items(items) {
        out.push_str(open);
        match group_format(group) {
            None => {
                if let Some(rpr) = original_rpr {
                    out.push_str(rpr);
                }
            }
            Some(diff) => {
                let mut kids = run.rpr.map(|(a, b)| children(tokens, xml, a, b)).unwrap_or_default();
                patch_rpr(&mut kids, &prefix, diff);
                let rpr_open = run
                    .rpr
                    .map(|(a, b)| open_tag_of(tokens, xml, a, b))
                    .unwrap_or_else(|| format!("<{}>", qualified(&prefix, "rPr")));
                write_rpr(out, &prefix, &rpr_open, &kids);
            }
        }
        write_items(out, &prefix, group, |first, last| &xml[tokens[first].span.start..tokens[last].span.end]);
        out.push_str(&close);
    }
}

fn emit_new_run(out: &mut String, prefix: &str, base_rpr: Option<Vec<(String, String)>>, items: &[Item]) {
    for group in group_items(items) {
        out.push_str(&format!("<{}>", qualified(prefix, "r")));
        let mut kids = base_rpr.clone().unwrap_or_default();
        if let Some(diff) = group_format(group) {
            patch_rpr(&mut kids, prefix, diff);
        }
        write_rpr(out, prefix, &format!("<{}>", qualified(prefix, "rPr")), &kids);
        write_items(out, prefix, group, |_, _| "");
        out.push_str(&format!("</{}>", qualified(prefix, "r")));
    }
}

fn write_rpr(out: &mut String, prefix: &str, open: &str, kids: &[(String, String)]) {
    if kids.is_empty() {
        return;
    }
    out.push_str(open);
    for (_, x) in kids {
        out.push_str(x);
    }
    out.push_str(&format!("</{}>", qualified(prefix, "rPr")));
}

/// Splits items into consecutive groups sharing the same formatting diff
fn group_items(items: &[Item]) -> Vec<&[Item]> {
    let mut groups = Vec::new();
    let mut start = 0;
    let mut current: Option<&Option<FormatTarget>> = None;
    for (idx, item) in items.iter().enumerate() {
        if let Item::Char { format, .. } = item {
            match current {
                Some(c) if c != format => {
                    groups.push(&items[start..idx]);
                    start = idx;
                    current = Some(format);
                }
                None => current = Some(format),
                _ => {}
            }
        }
    }
    if start < items.len() {
        groups.push(&items[start..]);
    }
    groups
}

fn group_format(group: &[Item]) -> Option<&FormatTarget> {
    group.iter().find_map(|item| match item {
        Item::Char { format, .. } => Some(format.as_ref()),
        _ => None,
    })?
}

fn write_items<'x>(out: &mut String, prefix: &str, group: &[Item], raw: impl Fn(usize, usize) -> &'x str) {
    let mut pending = String::new();
    let flush = |out: &mut String, pending: &mut String| {
        if !pending.is_empty() {
            let t = qualified(prefix, "t");
            out.push_str(&format!("<{t} xml:space=\"preserve\">{}</{t}>", escape_text(pending)));
            pending.clear();
        }
    };
    for item in group {
        match item {
            Item::Char { atom: Some((first, last)), .. } => {
                flush(out, &mut pending);
                out.push_str(raw(*first, *last));
            }
            Item::Char { ch: '\t', .. } => {
                flush(out, &mut pending);
                out.push_str(&format!("<{}/>", qualified(prefix, "tab")));
            }
            Item::Char { ch: '\n', .. } => {
                flush(out, &mut pending);
                out.push_str(&format!("<{}/>", qualified(prefix, "br")));
            }
            Item::Char { ch, .. } => pending.push(*ch),
            Item::Opaque { first, last } => {
                flush(out, &mut pending);
                out.push_str(raw(*first, *last));
            }
        }
    }
    flush(out, &mut pending);
}

#[cfg(test)]
mod tests {
    use super::*;

    fn plain(text: &str) -> String {
        text.to_string()
    }

    fn formats_for(runs: &[(&str, bool)]) -> (String, Vec<FormatTarget>) {
        let text: String = runs.iter().map(|(t, _)| *t).collect();
        let formats = runs
            .iter()
            .flat_map(|(t, bold)| {
                std::iter::repeat(FormatTarget { bold: Some(*bold), ..Default::default() }).take(t.chars().count())
            })
            .collect();
        (text, formats)
    }

    #[test]
    fn test_noop_edit_is_byte_identical() {
        let p = r#"<w:p w:rsidR="00A1"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>Hola</w:t></w:r></w:p>"#;
        assert_eq!(edit_paragraph(p, &StyleSheet::default(), "Hola", None, Some("center")).unwrap(), p);
    }

    #[test]
    fn test_typing_preserves_hyperlink_and_run_properties() {
        let link = r#"<w:hyperlink r:id="rId5"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>enlace</w:t></w:r></w:hyperlink>"#;
        let tail = r#"<w:r><w:t xml:space="preserve"> final</w:t></w:r>"#;
        let p = format!(
            r#"<w:p><w:r><w:rPr><w:highlight w:val="yellow"/><w:lang w:val="es-CO"/></w:rPr><w:t xml:space="preserve">Hola </w:t></w:r>{link}{tail}</w:p>"#
        );
        let out = edit_paragraph(&p, &StyleSheet::default(), &plain("Hola mundo enlace final"), None, None).unwrap();
        assert!(out.contains(r#"<w:rPr><w:highlight w:val="yellow"/><w:lang w:val="es-CO"/></w:rPr><w:t xml:space="preserve">Hola mundo </w:t>"#), "{}", out);
        assert!(out.contains(link), "hyperlink must be untouched: {}", out);
        assert!(out.contains(tail), "untouched runs are copied verbatim: {}", out);
        assert_eq!(parse_paragraph_fragment(&out, &StyleSheet::default()).text, "Hola mundo enlace final");
    }

    #[test]
    fn test_drawing_keeps_its_position() {
        let drawing = r#"<w:r><w:drawing><wp:inline><a:graphic/></wp:inline></w:drawing></w:r>"#;
        let p = format!(r#"<w:p><w:r><w:t>Antes</w:t></w:r>{drawing}<w:r><w:t>Después</w:t></w:r></w:p>"#);
        let out = edit_paragraph(&p, &StyleSheet::default(), "AhoraDespués", None, None).unwrap();
        let expected = format!(r#"<w:p><w:r><w:t xml:space="preserve">Ahora</w:t></w:r>{drawing}<w:r><w:t>Después</w:t></w:r></w:p>"#);
        assert_eq!(out, expected);
    }

    #[test]
    fn test_bold_one_word_splits_run_and_keeps_rpr_order() {
        let p = r#"<w:p><w:r><w:rPr><w:highlight w:val="cyan"/><w:lang w:val="es-ES"/></w:rPr><w:t>uno dos tres</w:t></w:r></w:p>"#;
        let (text, formats) = formats_for(&[("uno ", false), ("dos", true), (" tres", false)]);
        let out = edit_paragraph(p, &StyleSheet::default(), &text, Some(&formats), None).unwrap();
        assert!(out.contains(r#"<w:rPr><w:b/><w:highlight w:val="cyan"/><w:lang w:val="es-ES"/></w:rPr><w:t xml:space="preserve">dos</w:t>"#), "{}", out);
        let info = parse_paragraph_fragment(&out, &StyleSheet::default());
        let bold: Vec<(&str, bool)> = info.runs.iter().map(|r| (r.text.as_str(), r.bold)).collect();
        assert_eq!(bold, vec![("uno ", false), ("dos", true), (" tres", false)]);
    }

    #[test]
    fn test_font_size_not_lost_when_editor_omits_it() {
        let p = r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="28"/></w:rPr><w:t>Título</w:t></w:r></w:p>"#;
        let run = RunInfo { text: "Título nuevo".into(), ..Default::default() };
        let formats = vec![FormatTarget::from_run(&run); run.text.chars().count()];
        let out = edit_paragraph(p, &StyleSheet::default(), &run.text, Some(&formats), None).unwrap();
        assert!(out.contains(r#"<w:sz w:val="28"/>"#) && out.contains("Georgia"), "{}", out);
    }

    #[test]
    fn test_fields_survive_edits_outside_them() {
        let field = r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#;
        let p = format!(r#"<w:p><w:r><w:t xml:space="preserve">Página </w:t></w:r>{field}</w:p>"#);
        let out = edit_paragraph(&p, &StyleSheet::default(), "Pág. 3", None, None).unwrap();
        assert!(out.contains(field), "{}", out);
        assert_eq!(parse_paragraph_fragment(&out, &StyleSheet::default()).text, "Pág. 3");
    }

    #[test]
    fn test_delete_across_runs() {
        let p = r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>AB</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>CD</w:t></w:r></w:p>"#;
        let out = edit_paragraph(p, &StyleSheet::default(), "AD", None, None).unwrap();
        let runs = parse_paragraph_fragment(&out, &StyleSheet::default()).runs;
        assert_eq!(runs.len(), 2);
        assert_eq!((runs[0].text.as_str(), runs[0].bold), ("A", true));
        assert_eq!((runs[1].text.as_str(), runs[1].italic), ("D", true));
    }

    #[test]
    fn test_tabs_and_breaks_keep_original_markup() {
        let p = r#"<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br w:type="page"/><w:t>c</w:t></w:r></w:p>"#;
        let out = edit_paragraph(p, &StyleSheet::default(), "a\tb\ncX", None, None).unwrap();
        assert!(out.contains(r#"<w:tab/>"#) && out.contains(r#"<w:br w:type="page"/>"#), "{}", out);
        assert_eq!(parse_paragraph_fragment(&out, &StyleSheet::default()).text, "a\tb\ncX");
    }

    #[test]
    fn test_typing_into_empty_paragraph_uses_paragraph_mark_format() {
        let p = r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:rPr><w:color w:val="1F6F6B"/></w:rPr></w:pPr></w:p>"#;
        let out = edit_paragraph(p, &StyleSheet::default(), "Nuevo", None, None).unwrap();
        assert!(out.contains(r#"<w:r><w:rPr><w:color w:val="1F6F6B"/></w:rPr><w:t xml:space="preserve">Nuevo</w:t></w:r></w:p>"#), "{}", out);

        let out = edit_paragraph("<w:p/>", &StyleSheet::default(), "x", None, Some("center")).unwrap();
        assert_eq!(out, r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p>"#);
    }

    #[test]
    fn test_alignment_inserted_in_schema_order() {
        let p = r#"<w:p><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:after="120"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>"#;
        let out = edit_paragraph(p, &StyleSheet::default(), "x", None, Some("both")).unwrap();
        assert!(out.contains(r#"<w:spacing w:after="120"/><w:jc w:val="both"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r>"#), "{}", out);
    }

    #[test]
    fn test_text_box_paragraphs_do_not_shift_indices() {
        let xml = r#"<w:document><w:body><w:p><w:r><w:t>uno</w:t></w:r><w:r><w:drawing><wps:txbx><w:txbxContent><w:p><w:r><w:t>caja</w:t></w:r></w:p><w:p/></w:txbxContent></wps:txbx></w:drawing></w:r></w:p><w:p/><w:p><w:r><w:t>tres</w:t></w:r></w:p></w:body></w:document>"#;
        let ranges = body_paragraph_ranges(xml).unwrap();
        let parsed = crate::docx_parser::parse_document_elements(xml);
        assert_eq!(ranges.len(), 3);
        assert_eq!(parsed.len(), 3);
        assert_eq!(parse_paragraph_fragment(&xml[ranges[0].clone()], &StyleSheet::default()).text, "uno");
        assert_eq!(parse_paragraph_fragment(&xml[ranges[2].clone()], &StyleSheet::default()).text, "tres");
    }

    #[test]
    fn test_edits_compare_against_style_resolved_formatting() {
        let styles = StyleSheet::load(
            Some(r#"<w:styles><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:color w:val="2F5496"/></w:rPr></w:style></w:styles>"#),
            None,
            None,
        );
        let p = r#"<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Título</w:t></w:r></w:p>"#;
        let info = parse_paragraph_fragment(p, &styles);
        assert!(info.runs[0].bold, "bold comes from the style");
        assert_eq!(info.runs[0].color, "2F5496");

        // The editor sends back what it was shown: nothing to write
        let same: Vec<FormatTarget> = vec![FormatTarget::from_run(&info.runs[0]); 6];
        assert_eq!(edit_paragraph(p, &styles, "Título", Some(&same), None).unwrap(), p);

        // Turning bold off must override the style explicitly
        let mut plain = info.runs[0].clone();
        plain.bold = false;
        let formats = vec![FormatTarget::from_run(&plain); 6];
        let out = edit_paragraph(p, &styles, "Título", Some(&formats), None).unwrap();
        assert!(out.contains(r#"<w:rPr><w:b w:val="0"/></w:rPr>"#), "{}", out);
        assert!(!parse_paragraph_fragment(&out, &styles).runs[0].bold);
    }
}

