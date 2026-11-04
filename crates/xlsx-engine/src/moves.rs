//! Moving whole rows or columns, as Excel's "cut and insert": the block is taken out and put back
//! before another line, and everything that points at those lines follows. Formulas are rewritten
//! as text, reference by reference, so what the evaluator does not understand (`[1]Hoja!A1`,
//! `Tabla[Col]`, `1:1`, unions) survives and every byte that does not change stays as written.
//!
//! - A reference to a cell follows the cell.
//! - A range keeps its lines: a move inside it leaves it as it is; a block that leaves it shrinks
//!   it and a block put inside it grows it. A range inside the block goes with the block.
//! - Merged cells, multi-cell array formulas, tables and filters cannot be torn: a move that would
//!   do so is refused and nothing changes. The exceptions, which do not change their shape, are
//!   reordering inside a merged range without touching its first line, the data rows of a table or
//!   filter, and the columns of a table or filter.
//! - Sheets with pivot tables are refused: their layout depends on positions this engine does not
//!   rebuild.
//!
//! Inserting and deleting lines are the same operation with the lines at the end of the sheet:
//! to insert, the empty last lines are moved to the place (so ranges grow as they do when a block
//! enters them) and take the format of the line before; to delete, the block is moved to the end
//! and dropped, and what pointed only at it becomes `#REF!`. Merged cells and rows of tables and
//! filters then grow and shrink. Columns of a table and a whole table cannot be inserted or
//! deleted, and inserting is refused when the last lines of the sheet have values.

use std::collections::HashMap;
use std::ops::Range;

use quick_xml::events::{BytesStart, Event};
use quick_xml::reader::Reader;

use ooxml_core::rels::{parse_relationships, rels_part_for, resolve_target, type_is};
use ooxml_core::xml::{element_ranges, escape_xml, get_attr_value, remove_attr, remove_elements, set_attr, tag_is, unescape_text};

use crate::areas::parse_table;
use crate::cell_ref::{column_name, CellRange, CellRef, MAX_COL, MAX_ROW};
use crate::formula::{is_word_char, parse_col, parse_pos, Pos};
use crate::sheet::{CellNode, SheetXml};
use crate::sheets::{rewrite_texts, FORMULA_ELEMENTS};
use crate::workbook::{SheetKind, Workbook};

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Axis {
    Rows,
    Cols,
}

impl Axis {
    fn limit(self) -> u32 {
        match self {
            Axis::Rows => MAX_ROW,
            Axis::Cols => MAX_COL,
        }
    }

    /// `fila` or `columna`, for messages
    fn noun(self) -> &'static str {
        match self {
            Axis::Rows => "fila",
            Axis::Cols => "columna",
        }
    }
}

/// What the block of a `LineMove` is
#[derive(Debug, Clone, Copy, PartialEq)]
enum EditKind {
    Move,
    /// The block is the empty end of the sheet, which becomes the new lines
    Insert,
    /// The block goes to the end of the sheet and disappears
    Delete,
}

/// The lines `first..=last` taken out and inserted before the line `before` (numbered as they were
/// before the move)
#[derive(Debug, Clone, Copy)]
pub struct LineMove {
    pub axis: Axis,
    first: u32,
    last: u32,
    before: u32,
    kind: EditKind,
}

impl LineMove {
    pub fn new(axis: Axis, first: u32, count: u32, before: u32) -> Result<LineMove, String> {
        let limit = axis.limit();
        if count == 0 {
            return Err(format!("Hay que mover al menos una {}.", axis.noun()));
        }
        let last = first.checked_add(count - 1).filter(|l| first >= 1 && *l <= limit);
        let Some(last) = last else {
            return Err(format!("Las {} {}s desde la {} se salen de la hoja.", count, axis.noun(), first));
        };
        if before == 0 || before > limit + 1 {
            return Err(format!("El destino {} está fuera de la hoja.", before));
        }
        Ok(LineMove { axis, first, last, before, kind: EditKind::Move })
    }

    /// `count` new lines before the line `before`
    pub fn insert(axis: Axis, before: u32, count: u32) -> Result<LineMove, String> {
        let limit = axis.limit();
        if count == 0 {
            return Err(format!("Hay que insertar al menos una {}.", axis.noun()));
        }
        if before == 0 || before as u64 + count as u64 - 1 > limit as u64 {
            return Err(format!("No caben {} {}s nuevas en la {} {}.", count, axis.noun(), axis.noun(), before));
        }
        Ok(LineMove { axis, first: limit - count + 1, last: limit, before, kind: EditKind::Insert })
    }

    /// The lines from `first` go away
    pub fn delete(axis: Axis, first: u32, count: u32) -> Result<LineMove, String> {
        let limit = axis.limit();
        let mut edit = LineMove::new(axis, first, count, limit + 1)?;
        edit.kind = EditKind::Delete;
        Ok(edit)
    }

    /// Whether the line `i` does not exist after the edit
    fn removes(&self, i: u32) -> bool {
        self.kind != EditKind::Move && i >= self.first && i <= self.last
    }

    /// Whether the lines `first..=last` all disappear
    fn removes_all(&self, first: u32, last: u32) -> bool {
        self.kind != EditKind::Move && first >= self.first && last <= self.last
    }

    /// The line whose format new lines take: the one before them
    fn format_source(&self) -> Option<u32> {
        (self.kind == EditKind::Insert && self.before > 1).then(|| self.before - 1)
    }

    /// Where the first line of the block ends up (or the new lines start, or the next one goes)
    fn result_line(&self) -> u32 {
        match self.kind {
            EditKind::Move if self.is_noop() => self.first,
            EditKind::Move => self.line(self.first),
            EditKind::Insert => self.before,
            EditKind::Delete => self.first,
        }
    }

    fn count(&self) -> u32 {
        self.last - self.first + 1
    }

    /// Whether the block stays where it is
    fn is_noop(&self) -> bool {
        self.kind != EditKind::Delete && self.before >= self.first && self.before <= self.last + 1
    }

    /// Where the line `i` ends up
    pub fn line(&self, i: u32) -> u32 {
        if i >= self.first && i <= self.last {
            if self.before > self.last {
                i + (self.before - 1 - self.last)
            } else {
                i - (self.first - self.before)
            }
        } else if self.before > self.last && i > self.last && i < self.before {
            i - self.count()
        } else if self.before < self.first && i >= self.before && i < self.first {
            i + self.count()
        } else {
            i
        }
    }

    /// The lines where the shift changes: a shape over them cannot move as one piece
    fn boundaries(&self) -> [u32; 3] {
        [self.first, self.last + 1, self.before]
    }

    /// Whether the lines `first..=last` would not stay together
    fn crosses(&self, first: u32, last: u32) -> bool {
        self.boundaries().iter().any(|b| first < *b && *b <= last)
    }

    /// Whether the block and its destination are inside `first..=last`
    fn inside(&self, first: u32, last: u32) -> bool {
        self.first >= first && self.last <= last && self.before >= first && self.before <= last + 1
    }

    /// What happens to a range of references to the lines `first..=last` (see the module
    /// documentation)
    fn span(&self, first: u32, last: u32) -> (u32, u32) {
        if self.inside(first, last) {
            return (first, last);
        }
        if first >= self.first && last <= self.last {
            return (self.line(first), self.line(last));
        }
        let count = self.count();
        // The block taken out…
        let after_removal = |i: u32, end: bool| {
            if i < self.first {
                i
            } else if i > self.last {
                i - count
            } else if end {
                self.first - 1
            } else {
                self.first
            }
        };
        let (f, l) = (after_removal(first, false), after_removal(last, true));
        // …and put back
        let at = if self.before > self.last { self.before - count } else { self.before };
        let limit = self.axis.limit();
        if at <= f {
            ((f + count).min(limit), (l + count).min(limit))
        } else if at <= l {
            (f, (l + count).min(limit))
        } else {
            (f, l)
        }
    }

    /// The lines `first..=last` where they end up, in order, as runs of adjacent lines
    fn pieces(&self, first: u32, last: u32) -> Vec<(u32, u32)> {
        let mut cuts: Vec<u32> = self.boundaries().into_iter().filter(|b| first < *b && *b <= last).collect();
        cuts.sort_unstable();
        cuts.dedup();
        let mut start = first;
        let mut out = Vec::with_capacity(cuts.len() + 1);
        for cut in cuts.into_iter().chain(std::iter::once(last + 1)) {
            if !self.removes(start) {
                out.push((self.line(start), self.line(cut - 1)));
            }
            start = cut;
        }
        out.sort_unstable();
        // Runs that end up side by side are one
        let mut merged: Vec<(u32, u32)> = Vec::with_capacity(out.len());
        for (a, b) in out {
            match merged.last_mut() {
                Some(previous) if previous.1 + 1 == a => previous.1 = b,
                _ => merged.push((a, b)),
            }
        }
        merged
    }

    fn at(&self, cell: CellRef) -> CellRef {
        match self.axis {
            Axis::Rows => CellRef { row: self.line(cell.row), col: cell.col },
            Axis::Cols => CellRef { row: cell.row, col: self.line(cell.col) },
        }
    }

    fn bounds(&self, range: &CellRange) -> (u32, u32) {
        match self.axis {
            Axis::Rows => (range.first.row, range.last.row),
            Axis::Cols => (range.first.col, range.last.col),
        }
    }

    fn with_bounds(&self, range: &CellRange, first: u32, last: u32) -> CellRange {
        match self.axis {
            Axis::Rows => CellRange { first: CellRef { row: first, col: range.first.col }, last: CellRef { row: last, col: range.last.col } },
            Axis::Cols => CellRange { first: CellRef { row: range.first.row, col: first }, last: CellRef { row: range.last.row, col: last } },
        }
    }

    /// A range that holds together, where it ends up
    fn rigid(&self, range: &CellRange) -> CellRange {
        let (first, last) = self.bounds(range);
        self.with_bounds(range, self.line(first), self.line(last))
    }

    fn moved_span(&self, range: &CellRange) -> CellRange {
        let (first, last) = self.bounds(range);
        let (first, last) = self.span(first, last);
        self.with_bounds(range, first, last)
    }
}

// ---------- References in formula text ----------

/// A reference found in the text of a formula
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) enum RefKind {
    Cell(Pos),
    Range(Pos, Pos),
    /// `A:C`: (column, absolute) of each end
    Cols((u32, bool), (u32, bool)),
    /// `1:5`: (row, absolute) of each end
    Rows((u32, bool), (u32, bool)),
}

impl RefKind {
    fn text(&self) -> String {
        let pos = |p: &Pos| format!("{}{}{}{}", if p.col_abs { "$" } else { "" }, column_name(p.col), if p.row_abs { "$" } else { "" }, p.row);
        let col = |(c, abs): &(u32, bool)| format!("{}{}", if *abs { "$" } else { "" }, column_name(*c));
        let row = |(r, abs): &(u32, bool)| format!("{}{}", if *abs { "$" } else { "" }, r);
        match self {
            RefKind::Cell(p) => pos(p),
            RefKind::Range(p, q) => format!("{}:{}", pos(p), pos(q)),
            RefKind::Cols(a, b) => format!("{}:{}", col(a), col(b)),
            RefKind::Rows(a, b) => format!("{}:{}", row(a), row(b)),
        }
    }
}

/// `$3` → (row, absolute)
fn parse_row(word: &str) -> Option<(u32, bool)> {
    let (abs, digits) = match word.strip_prefix('$') {
        Some(rest) => (true, rest),
        None => (false, word),
    };
    if digits.is_empty() || digits.len() > 7 || digits.starts_with('0') || !digits.bytes().all(|b| b.is_ascii_digit()) {
        return None;
    }
    let row: u32 = digits.parse().ok()?;
    (row <= MAX_ROW).then_some((row, abs))
}

fn ref_pair(a: &str, b: &str) -> Option<RefKind> {
    if let (Some(p), Some(q)) = (parse_pos(a), parse_pos(b)) {
        return Some(RefKind::Range(p, q));
    }
    if let (Some(p), Some(q)) = (parse_col(a), parse_col(b)) {
        return Some(RefKind::Cols(p, q));
    }
    if let (Some(p), Some(q)) = (parse_row(a), parse_row(b)) {
        return Some(RefKind::Rows(p, q));
    }
    None
}

/// `formula` with each reference passed through `rewrite`, which gets the sheet written before it
/// (when any), whether that sheet is outside the plain sheets of this workbook (`[1]Hoja!`,
/// `Hoja1:Hoja3!`) and the reference, and returns the new text of the reference or `None` to leave
/// it as written. Texts, table references and everything else are copied as they are.
pub(crate) fn rewrite_refs(formula: &str, rewrite: &mut impl FnMut(Option<&str>, bool, &RefKind) -> Option<String>) -> String {
    let chars: Vec<char> = formula.chars().collect();
    let mut out = String::with_capacity(formula.len());
    let mut sheet: Option<String> = None;
    let mut foreign = false;
    let mut after_bracket = false;
    let mut i = 0;
    let word_end = |from: usize| {
        let mut j = from;
        while j < chars.len() && is_word_char(chars[j]) {
            j += 1;
        }
        j
    };
    while i < chars.len() {
        let c = chars[i];
        match c {
            '"' => {
                out.push(c);
                i += 1;
                while i < chars.len() {
                    out.push(chars[i]);
                    i += 1;
                    if chars[i - 1] == '"' {
                        if chars.get(i) == Some(&'"') {
                            out.push('"');
                            i += 1;
                        } else {
                            break;
                        }
                    }
                }
            }
            '\'' => {
                let mut j = i + 1;
                let mut name = String::new();
                while j < chars.len() {
                    if chars[j] == '\'' {
                        if chars.get(j + 1) == Some(&'\'') {
                            name.push('\'');
                            j += 2;
                            continue;
                        }
                        break;
                    }
                    name.push(chars[j]);
                    j += 1;
                }
                let end = (j + 1).min(chars.len());
                out.extend(&chars[i..end]);
                i = end;
                if chars.get(i) == Some(&'!') {
                    out.push('!');
                    i += 1;
                    foreign = name.contains('[') || name.contains(':');
                    sheet = Some(name);
                }
            }
            '[' => {
                let mut depth = 0;
                while i < chars.len() {
                    out.push(chars[i]);
                    match chars[i] {
                        '[' => depth += 1,
                        ']' => depth -= 1,
                        _ => {}
                    }
                    i += 1;
                    if depth == 0 {
                        break;
                    }
                }
                after_bracket = true;
                continue;
            }
            '#' => {
                // An error literal such as #REF! (or the # of a spilled range)
                out.push(c);
                i += 1;
                while i < chars.len() && (chars[i].is_alphanumeric() || "/!?_.".contains(chars[i])) {
                    out.push(chars[i]);
                    i += 1;
                }
            }
            c if is_word_char(c) => {
                let j = word_end(i);
                let word: String = chars[i..j].iter().collect();
                let next = chars.get(j).copied();
                if next == Some('!') {
                    foreign = after_bracket;
                    sheet = Some(word.clone());
                    out.push_str(&word);
                    out.push('!');
                    i = j + 1;
                    after_bracket = false;
                    continue;
                }
                if matches!(next, Some('(') | Some('[')) {
                    out.push_str(&word);
                    i = j;
                    sheet = None;
                    after_bracket = false;
                    continue;
                }
                let mut kind = None;
                let mut end = j;
                if next == Some(':') {
                    let k = word_end(j + 1);
                    let second: String = chars[j + 1..k].iter().collect();
                    if chars.get(k) == Some(&'!') {
                        // Hoja1:Hoja3!A1
                        out.extend(&chars[i..=k]);
                        i = k + 1;
                        sheet = Some(format!("{}:{}", word, second));
                        foreign = true;
                        after_bracket = false;
                        continue;
                    }
                    if !matches!(chars.get(k), Some('(') | Some('[')) {
                        kind = ref_pair(&word, &second);
                        if kind.is_some() {
                            end = k;
                        }
                    }
                }
                if kind.is_none() {
                    kind = parse_pos(&word).map(RefKind::Cell);
                }
                let raw: String = chars[i..end].iter().collect();
                match kind.and_then(|k| rewrite(sheet.as_deref(), foreign, &k)) {
                    Some(text) => out.push_str(&text),
                    None => out.push_str(&raw),
                }
                i = end;
                sheet = None;
                foreign = false;
            }
            _ => {
                out.push(c);
                i += 1;
            }
        }
        after_bracket = false;
    }
    out
}

/// The formula of a shared group as written for the cell `rows` below and `cols` to the right of
/// its first cell
pub(crate) fn shift_formula(formula: &str, rows: i64, cols: i64) -> String {
    let line = |value: u32, abs: bool, delta: i64, limit: u32| -> Option<u32> {
        if abs {
            return Some(value);
        }
        let shifted = value as i64 + delta;
        (shifted >= 1 && shifted <= limit as i64).then_some(shifted as u32)
    };
    let pos = |p: Pos| -> Option<Pos> {
        Some(Pos { row: line(p.row, p.row_abs, rows, MAX_ROW)?, col: line(p.col, p.col_abs, cols, MAX_COL)?, ..p })
    };
    rewrite_refs(formula, &mut |_, _, kind| {
        let shifted = match *kind {
            RefKind::Cell(p) => pos(p).map(RefKind::Cell),
            RefKind::Range(p, q) => pos(p).zip(pos(q)).map(|(p, q)| RefKind::Range(p, q)),
            RefKind::Cols(a, b) => line(a.0, a.1, cols, MAX_COL).zip(line(b.0, b.1, cols, MAX_COL)).map(|(x, y)| RefKind::Cols((x, a.1), (y, b.1))),
            RefKind::Rows(a, b) => line(a.0, a.1, rows, MAX_ROW).zip(line(b.0, b.1, rows, MAX_ROW)).map(|(x, y)| RefKind::Rows((x, a.1), (y, b.1))),
        };
        Some(shifted.map(|k| k.text()).unwrap_or_else(|| "#REF!".to_string()))
    })
}

/// Where a reference goes when the lines of its sheet move; `None` when the axis does not touch it
fn moved_ref(kind: &RefKind, mv: &LineMove) -> Option<RefKind> {
    let ends = |a: u32, b: u32| {
        let (first, last) = mv.span(a.min(b), a.max(b));
        if a <= b {
            (first, last)
        } else {
            (last, first)
        }
    };
    match (*kind, mv.axis) {
        (RefKind::Cell(p), Axis::Rows) => Some(RefKind::Cell(Pos { row: mv.line(p.row), ..p })),
        (RefKind::Cell(p), Axis::Cols) => Some(RefKind::Cell(Pos { col: mv.line(p.col), ..p })),
        (RefKind::Range(p, q), Axis::Rows) => {
            let (a, b) = ends(p.row, q.row);
            Some(RefKind::Range(Pos { row: a, ..p }, Pos { row: b, ..q }))
        }
        (RefKind::Range(p, q), Axis::Cols) => {
            let (a, b) = ends(p.col, q.col);
            Some(RefKind::Range(Pos { col: a, ..p }, Pos { col: b, ..q }))
        }
        (RefKind::Cols(p, q), Axis::Cols) => {
            let (a, b) = ends(p.0, q.0);
            Some(RefKind::Cols((a, p.1), (b, q.1)))
        }
        (RefKind::Rows(p, q), Axis::Rows) => {
            let (a, b) = ends(p.0, q.0);
            Some(RefKind::Rows((a, p.1), (b, q.1)))
        }
        _ => None,
    }
}

/// Whether a reference points only at lines that disappear
fn removed_ref(kind: &RefKind, mv: &LineMove) -> bool {
    match (*kind, mv.axis) {
        (RefKind::Cell(p), Axis::Rows) => mv.removes(p.row),
        (RefKind::Cell(p), Axis::Cols) => mv.removes(p.col),
        (RefKind::Range(p, q), Axis::Rows) => mv.removes_all(p.row.min(q.row), p.row.max(q.row)),
        (RefKind::Range(p, q), Axis::Cols) => mv.removes_all(p.col.min(q.col), p.col.max(q.col)),
        (RefKind::Cols(p, q), Axis::Cols) => mv.removes_all(p.0.min(q.0), p.0.max(q.0)),
        (RefKind::Rows(p, q), Axis::Rows) => mv.removes_all(p.0.min(q.0), p.0.max(q.0)),
        _ => false,
    }
}

/// The sheet whose lines move, as formulas name it
#[derive(Clone, Copy)]
struct Target<'a> {
    name: &'a str,
    mv: &'a LineMove,
}

impl Target<'_> {
    fn names(&self, sheet: Option<&str>, host_is_target: bool) -> bool {
        match sheet {
            Some(s) => s.to_lowercase() == self.name.to_lowercase(),
            None => host_is_target,
        }
    }

    /// `formula` of a cell of the moved sheet (`host_is_target`) or of another place
    fn formula(&self, formula: &str, host_is_target: bool) -> String {
        if !host_is_target {
            // Most formulas of other sheets do not mention it: skip scanning them
            let needle = self.name.replace('\'', "''").to_lowercase();
            if !formula.to_lowercase().contains(&needle) {
                return formula.to_string();
            }
        }
        rewrite_refs(formula, &mut |sheet, foreign, kind| {
            if foreign || !self.names(sheet, host_is_target) {
                return None;
            }
            if removed_ref(kind, self.mv) {
                return Some("#REF!".to_string());
            }
            moved_ref(kind, self.mv).map(|k| k.text())
        })
    }

    /// A formula of a conditional format or data validation of the moved sheet. It is written for
    /// the top left cell of the cells it applies to, and its relative references are distances
    /// from each cell: they stay the same distance (`$A1` on a row keeps looking at its own row)
    /// when that cell moves by `delta`. Absolute references follow their lines.
    fn anchored_formula(&self, formula: &str, delta: i64) -> String {
        let mv = self.mv;
        let limit = mv.axis.limit();
        let relative = |value: u32| (value as i64 + delta).clamp(1, limit as i64) as u32;
        let line = |value: u32, abs: bool, on_target: bool| match (abs, on_target) {
            (true, true) => mv.line(value),
            (true, false) => value,
            (false, _) => relative(value),
        };
        rewrite_refs(formula, &mut |sheet, foreign, kind| {
            if foreign {
                return None;
            }
            let on_target = self.names(sheet, true);
            let kind = *kind;
            let absolute = match (kind, mv.axis) {
                (RefKind::Cell(p), Axis::Rows) => p.row_abs,
                (RefKind::Cell(p), Axis::Cols) => p.col_abs,
                (RefKind::Range(p, q), Axis::Rows) => p.row_abs && q.row_abs,
                (RefKind::Range(p, q), Axis::Cols) => p.col_abs && q.col_abs,
                (RefKind::Cols(p, q), Axis::Cols) | (RefKind::Rows(p, q), Axis::Rows) => p.1 && q.1,
                _ => false,
            };
            if on_target && absolute && removed_ref(&kind, mv) {
                return Some("#REF!".to_string());
            }
            let moved = match (kind, mv.axis) {
                // Absolute at both ends: a plain range of lines
                (RefKind::Range(p, q), Axis::Rows) if p.row_abs && q.row_abs => return on_target.then(|| moved_ref(&kind, mv).map(|k| k.text())).flatten(),
                (RefKind::Range(p, q), Axis::Cols) if p.col_abs && q.col_abs => return on_target.then(|| moved_ref(&kind, mv).map(|k| k.text())).flatten(),
                (RefKind::Cols(p, q), Axis::Cols) if p.1 && q.1 => return on_target.then(|| moved_ref(&kind, mv).map(|k| k.text())).flatten(),
                (RefKind::Rows(p, q), Axis::Rows) if p.1 && q.1 => return on_target.then(|| moved_ref(&kind, mv).map(|k| k.text())).flatten(),
                (RefKind::Cell(p), Axis::Rows) => RefKind::Cell(Pos { row: line(p.row, p.row_abs, on_target), ..p }),
                (RefKind::Cell(p), Axis::Cols) => RefKind::Cell(Pos { col: line(p.col, p.col_abs, on_target), ..p }),
                (RefKind::Range(p, q), Axis::Rows) => RefKind::Range(Pos { row: line(p.row, p.row_abs, on_target), ..p }, Pos { row: line(q.row, q.row_abs, on_target), ..q }),
                (RefKind::Range(p, q), Axis::Cols) => RefKind::Range(Pos { col: line(p.col, p.col_abs, on_target), ..p }, Pos { col: line(q.col, q.col_abs, on_target), ..q }),
                (RefKind::Cols(p, q), Axis::Cols) => RefKind::Cols((line(p.0, p.1, on_target), p.1), (line(q.0, q.1, on_target), q.1)),
                (RefKind::Rows(p, q), Axis::Rows) => RefKind::Rows((line(p.0, p.1, on_target), p.1), (line(q.0, q.1, on_target), q.1)),
                _ => return None,
            };
            (moved != kind).then(|| moved.text())
        })
    }
}

// ---------- Text edits ----------

/// `xml` with the ranges replaced; the ranges do not overlap
fn apply_edits(xml: &str, mut edits: Vec<(Range<usize>, String)>) -> String {
    edits.sort_by_key(|(range, _)| range.start);
    let mut out = String::with_capacity(xml.len());
    let mut cursor = 0;
    for (range, text) in edits {
        out.push_str(&xml[cursor..range.start]);
        out.push_str(&text);
        cursor = range.end;
    }
    out.push_str(&xml[cursor..]);
    out
}

/// The `sqref` list `A1:A9 C3` where its cells end up
fn moved_sqref(text: &str, mv: &LineMove) -> String {
    let mut out: Vec<String> = Vec::new();
    for token in text.split_whitespace() {
        match CellRange::parse(token) {
            // What is inserted in the middle of a range takes part in it
            Ok(range) if mv.kind == EditKind::Insert => out.push(mv.moved_span(&range).to_string()),
            Ok(range) => {
                let (first, last) = mv.bounds(&range);
                out.extend(mv.pieces(first, last).into_iter().map(|(a, b)| mv.with_bounds(&range, a, b).to_string()));
            }
            Err(_) => out.push(token.to_string()),
        }
    }
    out.join(" ")
}

/// How far the top left of the cells of a `sqref` moves along the axis
fn sqref_delta(text: &str, mv: &LineMove) -> i64 {
    let mut before = u32::MAX;
    let mut after = u32::MAX;
    for range in text.split_whitespace().filter_map(|t| CellRange::parse(t).ok()) {
        let (first, last) = mv.bounds(&range);
        before = before.min(first);
        after = mv.pieces(first, last).into_iter().map(|p| p.0).fold(after, u32::min);
    }
    if before == u32::MAX || after == u32::MAX {
        0
    } else {
        after as i64 - before as i64
    }
}

/// How a table or filter takes a move
#[derive(PartialEq)]
enum ShapeMove {
    /// All of it moves (or stays), as one piece
    Rigid,
    /// Lines change places inside it; its range stays
    Inner,
    /// Rows inserted or deleted inside it: its range grows or shrinks
    Resized,
}

/// A table or filter over `range`, with `header` and `totals` rows that must stay in place
fn shape_move(what: &str, range: &CellRange, header: u32, totals: u32, mv: &LineMove) -> Result<ShapeMove, String> {
    let (first, last) = mv.bounds(range);
    if !mv.crosses(first, last) {
        return Ok(ShapeMove::Rigid);
    }
    let (body_first, body_last) = match mv.axis {
        Axis::Rows => (first + header, last.saturating_sub(totals)),
        Axis::Cols => (first, last),
    };
    if mv.inside(body_first, body_last) {
        return Ok(ShapeMove::Inner);
    }
    Err(format!("No se puede mover: la {} rompería {} {}.", mv.axis.noun(), what, range))
}

// ---------- The moved sheet ----------

/// What changes in the `<f>` of a cell
#[derive(Default)]
struct FormulaEdit {
    text: Option<String>,
    unshare: bool,
}

/// The `<f>` of the cells whose formula changes. A shared formula stands for many cells: when any
/// of them changes or moves, its cells get a formula of their own.
fn formula_edits(parsed: &SheetXml, host_is_target: bool, target: Target) -> HashMap<CellRef, FormulaEdit> {
    let mut masters: HashMap<u32, (CellRef, &str)> = HashMap::new();
    let cells = || parsed.rows.iter().flat_map(|r| &r.cells);
    for cell in cells() {
        if let Some(f) = &cell.formula {
            if f.kind.as_deref() == Some("shared") && !f.text.trim().is_empty() {
                if let Some(si) = f.shared_index {
                    masters.insert(si, (cell.at, &f.text));
                }
            }
        }
    }
    let moves = |at: CellRef| host_is_target && target.mv.at(at) != at;
    let mut edits = HashMap::new();
    let mut groups: HashMap<u32, Vec<(CellRef, String, bool)>> = HashMap::new();
    for cell in cells() {
        let Some(f) = &cell.formula else { continue };
        if f.kind.as_deref() == Some("shared") {
            if let Some((si, (anchor, text))) = f.shared_index.and_then(|si| masters.get(&si).map(|m| (si, *m))) {
                let own = shift_formula(text, cell.at.row as i64 - anchor.row as i64, cell.at.col as i64 - anchor.col as i64);
                let mapped = target.formula(&own, host_is_target);
                let changed = mapped != own || moves(cell.at);
                groups.entry(si).or_default().push((cell.at, mapped, changed));
                continue;
            }
        }
        if !f.text.trim().is_empty() {
            let mapped = target.formula(&f.text, host_is_target);
            if mapped != f.text {
                edits.insert(cell.at, FormulaEdit { text: Some(mapped), unshare: false });
            }
        }
    }
    for members in groups.into_values() {
        if members.iter().any(|m| m.2) {
            for (at, text, _) in members {
                edits.insert(at, FormulaEdit { text: Some(text), unshare: true });
            }
        }
    }
    edits
}

/// `xml` with the shared formulas whose first cell will be overwritten written out in each of
/// their cells, since the others read it from there. The rest of the sheet stays as it is.
pub(crate) fn without_shared_sources(xml: String, overwritten: impl Fn(CellRef) -> bool) -> Result<String, String> {
    let parsed = SheetXml::parse(&xml)?;
    let cells = || parsed.rows.iter().flat_map(|r| &r.cells);
    let mut masters: HashMap<u32, (CellRef, &str)> = HashMap::new();
    for cell in cells() {
        if let Some(f) = cell.formula.as_ref().filter(|f| f.shared_master && overwritten(cell.at)) {
            if let Some(si) = f.shared_index {
                masters.insert(si, (cell.at, &f.text));
            }
        }
    }
    if masters.is_empty() {
        drop(parsed);
        return Ok(xml);
    }
    let mut edits = std::collections::BTreeMap::new();
    for cell in cells() {
        let Some(f) = cell.formula.as_ref().filter(|f| f.kind.as_deref() == Some("shared")) else { continue };
        let Some((anchor, text)) = f.shared_index.and_then(|si| masters.get(&si)) else { continue };
        let own = shift_formula(text, cell.at.row as i64 - anchor.row as i64, cell.at.col as i64 - anchor.col as i64);
        let edit = FormulaEdit { text: Some(own), unshare: true };
        edits.insert(cell.at, Some(with_formula(&xml[cell.range.clone()], &edit)?));
    }
    Ok(parsed.splice(&edits))
}

/// A `<c>` element with its `<f>` changed
fn with_formula(cell: &str, edit: &FormulaEdit) -> Result<String, String> {
    let Some(range) = element_ranges(cell, "f", |_| true)?.into_iter().next() else { return Ok(cell.to_string()) };
    let mut f = cell[range.clone()].to_string();
    if edit.unshare {
        for attr in ["t", "ref", "si"] {
            f = remove_attr(&f, attr);
        }
    }
    if let Some(text) = &edit.text {
        let tag_end = f.find('>').map(|i| i + 1).unwrap_or(f.len());
        let self_closing = f[..tag_end].ends_with("/>");
        let open = if self_closing { format!("{}>", f[..tag_end - 2].trim_end()) } else { f[..tag_end].to_string() };
        let name_end = open.find(|c: char| c.is_whitespace() || c == '>').unwrap_or(open.len());
        f = format!("{}{}</{}>", open, escape_xml(text), &open[1..name_end]);
    }
    Ok(format!("{}{}{}", &cell[..range.start], f, &cell[range.end..]))
}

/// The attributes of an `<f>` that hold ranges of the sheet: an array or data table cannot be torn
fn moved_formula_attrs(cell: &str, node: &CellNode, mv: &LineMove) -> Result<String, String> {
    let Some(f) = &node.formula else { return Ok(cell.to_string()) };
    if !matches!(f.kind.as_deref(), Some("array") | Some("dataTable")) {
        return Ok(cell.to_string());
    }
    let Some(range) = element_ranges(cell, "f", |_| true)?.into_iter().next() else { return Ok(cell.to_string()) };
    let mut element = cell[range.clone()].to_string();
    let attrs: Vec<(&str, String)> = {
        let mut reader = Reader::from_str(&element);
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => ["ref", "r1", "r2"].iter().filter_map(|a| get_attr_value(&e, a).map(|v| (*a, v))).collect(),
            _ => Vec::new(),
        }
    };
    for (name, value) in attrs {
        let Ok(area) = CellRange::parse(&value) else { continue };
        let (first, last) = mv.bounds(&area);
        if mv.crosses(first, last) {
            return Err(format!("No se puede mover: la {} rompería la fórmula matricial de {}.", mv.axis.noun(), area));
        }
        element = set_attr(&element, name, &mv.rigid(&area).to_string());
    }
    Ok(format!("{}{}{}", &cell[..range.start], element, &cell[range.end..]))
}

/// The `<sheetData>` content of the moved sheet
fn moved_rows(parsed: &SheetXml, target: Target) -> Result<(String, Option<CellRange>), String> {
    let mv = target.mv;
    let xml = parsed.xml;
    let edits = formula_edits(parsed, true, target);
    let prefix = parsed.prefix();
    let source = mv.format_source();
    let grow = |range: Option<CellRange>, at: CellRef| match range {
        Some(range) => Some(range.union(CellRange::around(at, at))),
        None => Some(CellRange::around(at, at)),
    };
    let mut used: Option<CellRange> = None;
    let mut rows: Vec<(u32, String)> = Vec::with_capacity(parsed.rows.len());
    for row in &parsed.rows {
        if mv.axis == Axis::Rows && mv.removes(row.number) {
            refuse_pushing_off(row.cells.iter(), mv)?;
            continue;
        }
        // New rows take the height and format of the one before them, without its values
        if mv.axis == Axis::Rows && source == Some(row.number) {
            let mut tag = remove_attr(&remove_attr(&xml[row.start_tag.clone()], "spans"), "hidden");
            if !row.self_closing && row.cells.iter().all(|c| c.style_attr.is_none()) {
                tag = format!("{}/>", tag.trim_end_matches('>').trim_end());
            }
            for number in mv.before..mv.before + mv.count() {
                let tag = set_attr(&tag, "r", &number.to_string());
                if tag.ends_with("/>") {
                    rows.push((number, tag));
                    continue;
                }
                let mut cells = String::new();
                for cell in &row.cells {
                    if let Some(style) = &cell.style_attr {
                        let at = CellRef { row: number, col: cell.at.col };
                        used = grow(used, at);
                        cells.push_str(&format!(r#"<{}c r="{}" s="{}"/>"#, prefix, at, escape_xml(style)));
                    }
                }
                rows.push((number, format!("{}{}</{}row>", tag, cells, prefix)));
            }
        }
        let number = if mv.axis == Axis::Rows { mv.line(row.number) } else { row.number };
        let mut tag = xml[row.start_tag.clone()].to_string();
        if number != row.number || !row.has_ref {
            tag = set_attr(&tag, "r", &number.to_string());
        }
        if row.self_closing {
            rows.push((number, tag));
            continue;
        }
        if mv.axis == Axis::Cols {
            refuse_pushing_off(row.cells.iter().filter(|c| mv.removes(c.at.col)), mv)?;
        }
        let mut cells: Vec<(u32, String)> = Vec::with_capacity(row.cells.len());
        for cell in &row.cells {
            if mv.axis == Axis::Cols && mv.removes(cell.at.col) {
                continue;
            }
            // New columns take the format of the cell to their left
            if let (Axis::Cols, true, Some(style)) = (mv.axis, source == Some(cell.at.col), &cell.style_attr) {
                for col in mv.before..mv.before + mv.count() {
                    let at = CellRef { row: row.number, col };
                    used = grow(used, at);
                    cells.push((col, format!(r#"<{}c r="{}" s="{}"/>"#, prefix, at, escape_xml(style))));
                }
            }
            let at = mv.at(cell.at);
            let mut text = xml[cell.range.clone()].to_string();
            if at != cell.at || !cell.has_ref {
                text = set_attr(&text, "r", &at.to_string());
            }
            text = moved_formula_attrs(&text, cell, mv)?;
            if let Some(edit) = edits.get(&cell.at) {
                text = with_formula(&text, edit)?;
            }
            used = grow(used, at);
            cells.push((at.col, text));
        }
        if mv.axis == Axis::Cols && (cells.len() != row.cells.len() || row.cells.iter().any(|c| mv.at(c.at) != c.at)) {
            tag = remove_attr(&tag, "spans");
            cells.sort_by_key(|(col, _)| *col);
        }
        let tail = row.cells.last().map(|c| c.range.end).unwrap_or(row.start_tag.end);
        let body: String = cells.into_iter().map(|(_, text)| text).collect();
        rows.push((number, format!("{}{}{}", tag, body, &xml[tail..row.range.end])));
    }
    rows.sort_by_key(|(number, _)| *number);
    Ok((rows.into_iter().map(|(_, text)| text).collect(), used))
}

/// Inserting pushes the last lines off the sheet: only if they hold no values
fn refuse_pushing_off<'c>(cells: impl Iterator<Item = &'c CellNode>, mv: &LineMove) -> Result<(), String> {
    if mv.kind != EditKind::Insert {
        return Ok(());
    }
    for cell in cells {
        if cell.formula.is_some() || cell.inline.is_some() || cell.value.as_deref().is_some_and(|v| !v.is_empty()) {
            return Err(format!("No se puede insertar: la celda {} tiene datos y se saldría de la hoja.", cell.at));
        }
    }
    Ok(())
}

/// The frozen lines of a sheet in `head` grow or shrink when lines are inserted or deleted in them
fn moved_pane(head: &str, mv: &LineMove) -> Result<String, String> {
    let Some(range) = element_ranges(head, "pane", |e| {
        matches!(get_attr_value(e, "state").as_deref(), Some("frozen") | Some("frozenSplit"))
    })?
    .into_iter()
    .next() else {
        return Ok(head.to_string());
    };
    let element = &head[range.clone()];
    let attr = match mv.axis {
        Axis::Rows => "ySplit",
        Axis::Cols => "xSplit",
    };
    let read = |name: &str| -> Option<String> {
        let mut reader = Reader::from_str(element);
        match reader.read_event() {
            Ok(Event::Start(e)) | Ok(Event::Empty(e)) => get_attr_value(&e, name),
            _ => None,
        }
    };
    let Some(split) = read(attr).and_then(|v| v.parse::<f64>().ok()).map(|v| v as u32).filter(|v| *v > 0) else {
        return Ok(head.to_string());
    };
    let new_split = match mv.kind {
        EditKind::Move => split,
        EditKind::Insert if mv.before <= split => split + mv.count(),
        EditKind::Insert => split,
        EditKind::Delete => split - (mv.first..=mv.last.min(split)).count() as u32,
    };
    if new_split == split {
        return Ok(head.to_string());
    }
    let other = read(if mv.axis == Axis::Rows { "xSplit" } else { "ySplit" }).and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0);
    if new_split == 0 && other == 0.0 {
        return Ok(format!("{}{}", &head[..range.start], &head[range.end..]));
    }
    let mut written = if new_split == 0 { remove_attr(element, attr) } else { set_attr(element, attr, &new_split.to_string()) };
    if let Some(corner) = read("topLeftCell").and_then(|r| CellRef::parse(&r).ok()) {
        let corner = match mv.axis {
            Axis::Rows => CellRef { row: new_split + 1, col: corner.col },
            Axis::Cols => CellRef { row: corner.row, col: new_split + 1 },
        };
        written = set_attr(&written, "topLeftCell", &corner.to_string());
    }
    Ok(format!("{}{}{}", &head[..range.start], written, &head[range.end..]))
}

/// Takes out of a worksheet what a deletion left without cells (marked with an empty range), and
/// the lists left without elements
fn without_removed(xml: String) -> Result<String, String> {
    let mut xml = xml;
    for (item, attr, list) in [
        ("mergeCell", "ref", Some("mergeCells")),
        ("conditionalFormatting", "sqref", None),
        ("dataValidation", "sqref", Some("dataValidations")),
        ("hyperlink", "ref", Some("hyperlinks")),
        ("protectedRange", "sqref", Some("protectedRanges")),
        ("ignoredError", "sqref", Some("ignoredErrors")),
        ("autoFilter", "ref", None),
    ] {
        let written = remove_elements(&xml, item, |e| get_attr_value(e, attr).as_deref() == Some(""))?;
        if written == xml {
            continue;
        }
        xml = written;
        if let Some(list) = list {
            let Some(range) = element_ranges(&xml, list, |_| true)?.into_iter().next() else { continue };
            let left = element_ranges(&xml[range.clone()], item, |_| true)?.len();
            let element = if left == 0 {
                String::new()
            } else if xml[range.clone()].split('>').next().is_some_and(|tag| tag.contains("count=")) {
                set_attr(&xml[range.clone()], "count", &left.to_string())
            } else {
                xml[range.clone()].to_string()
            };
            xml.replace_range(range, &element);
        }
    }
    Ok(xml)
}

/// The `<cols>` of a sheet whose columns move, in `head` (the sheet up to its data)
fn moved_cols(head: &str, mv: &LineMove) -> Result<String, String> {
    let Some(range) = element_ranges(head, "cols", |_| true)?.into_iter().next() else { return Ok(head.to_string()) };
    let text = &head[range.clone()];
    let tag_end = text.find('>').map(|i| i + 1).unwrap_or(text.len());
    if text[..tag_end].ends_with("/>") {
        return Ok(head.to_string());
    }
    let mut cols: Vec<(u32, String)> = Vec::new();
    for element in element_ranges(text, "col", |_| true)? {
        let element = &text[element];
        let (min, max) = {
            let mut reader = Reader::from_str(element);
            match reader.read_event() {
                Ok(Event::Start(e)) | Ok(Event::Empty(e)) => {
                    let n = |a: &str| get_attr_value(&e, a).and_then(|v| v.trim().parse::<u32>().ok());
                    (n("min").unwrap_or(0), n("max").unwrap_or(0))
                }
                _ => (0, 0),
            }
        };
        if min == 0 || max < min {
            continue;
        }
        for (a, b) in mv.pieces(min, max.min(MAX_COL)) {
            cols.push((a, set_attr(&set_attr(element, "min", &a.to_string()), "max", &b.to_string())));
        }
        // New columns are as wide as the one to their left
        if mv.format_source().is_some_and(|c| (min..=max).contains(&c)) {
            let (a, b) = (mv.before, mv.before + mv.count() - 1);
            let copy = remove_attr(element, "hidden");
            cols.push((a, set_attr(&set_attr(&copy, "min", &a.to_string()), "max", &b.to_string())));
        }
    }
    cols.sort_by_key(|(min, _)| *min);
    let close = text.rfind("</").unwrap_or(text.len());
    let body: String = cols.into_iter().map(|(_, element)| element).collect();
    Ok(format!("{}{}{}{}{}", &head[..range.start], &text[..tag_end], body, &text[close..], &head[range.end..]))
}

/// A merged range after a move, or the refusal: it may only keep its shape. Lines inserted or
/// deleted in it make it grow or shrink; `None` when no merged range is left.
fn moved_merge(merge: &CellRange, mv: &LineMove) -> Result<Option<CellRange>, String> {
    let (first, last) = mv.bounds(merge);
    if mv.kind != EditKind::Move {
        if mv.removes_all(first, last) {
            return Ok(None);
        }
        let moved = mv.moved_span(merge);
        return Ok((moved.cell_count() > 1).then_some(moved));
    }
    if !mv.crosses(first, last) {
        return Ok(Some(mv.rigid(merge)));
    }
    // Lines reordered under a merged title: the first line holds its value and must stay first
    if mv.first > first && mv.last <= last && mv.before > first && mv.before <= last + 1 {
        return Ok(Some(*merge));
    }
    Err(format!("No se puede mover: la {} rompería las celdas combinadas {}.", mv.axis.noun(), merge))
}

fn attr_of(e: &BytesStart, name: &str) -> Option<String> {
    get_attr_value(e, name)
}

/// What follows `<sheetData>` in a worksheet. In the moved sheet (`host_is_target`) ranges of
/// cells follow their cells; in any sheet the formulas follow the lines.
fn moved_tail(tail: &str, host_is_target: bool, target: Target) -> Result<String, String> {
    let mv = target.mv;
    let mut reader = Reader::from_str(tail);
    let config = reader.config_mut();
    config.trim_text(false);
    // It starts and ends in the middle of the worksheet
    config.check_end_names = false;
    config.allow_unmatched_ends = true;
    let mut edits: Vec<(Range<usize>, String)> = Vec::new();
    let mut delta: i64 = 0;
    // First line of the filter being read, when its columns change places
    let mut filter_first: Option<(u32, u32)> = None;
    let mut open: Option<(&'static str, usize)> = None;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|e| format!("La hoja no es XML válido: {}", e))?;
        let after = reader.buffer_position() as usize;
        match &event {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) => {
                let name = e.name();
                let name = name.as_ref();
                if host_is_target {
                    let mut set = |attr: &str, value: String| edits.push((before..after, set_attr(&tail[before..after], attr, &value)));
                    if tag_is(name, "mergeCell") {
                        if let Some(merge) = attr_of(e, "ref").and_then(|r| CellRange::parse(&r).ok()) {
                            set("ref", moved_merge(&merge, mv)?.map(|m| m.to_string()).unwrap_or_default());
                        }
                    } else if tag_is(name, "conditionalFormatting") || tag_is(name, "dataValidation") || tag_is(name, "protectedRange") || tag_is(name, "ignoredError") {
                        if let Some(sqref) = attr_of(e, "sqref") {
                            delta = sqref_delta(&sqref, mv);
                            set("sqref", moved_sqref(&sqref, mv));
                        }
                    } else if tag_is(name, "hyperlink") {
                        let mut element = tail[before..after].to_string();
                        if let Some(reference) = attr_of(e, "ref") {
                            element = set_attr(&element, "ref", &moved_sqref(&reference, mv));
                        }
                        if let Some(location) = attr_of(e, "location") {
                            element = set_attr(&element, "location", &target.formula(&location, true));
                        }
                        edits.push((before..after, element));
                    } else if tag_is(name, "autoFilter") {
                        if let Some(range) = attr_of(e, "ref").and_then(|r| CellRange::parse(&r).ok()) {
                            if mv.kind != EditKind::Move {
                                let (first, last) = mv.bounds(&range);
                                if mv.axis == Axis::Cols && mv.crosses(first, last) {
                                    filter_first = Some((first, mv.span(first, last).0));
                                }
                                set("ref", if mv.removes_all(first, last) { String::new() } else { mv.moved_span(&range).to_string() });
                            } else if shape_move("el filtro", &range, 1, 0, mv)? == ShapeMove::Rigid {
                                set("ref", mv.rigid(&range).to_string());
                            } else if mv.axis == Axis::Cols {
                                filter_first = Some((range.first.col, range.first.col));
                            }
                        }
                    } else if tag_is(name, "filterColumn") {
                        if let (Some((first, new_first)), Some(id)) = (filter_first, attr_of(e, "colId").and_then(|v| v.parse::<u32>().ok())) {
                            if mv.removes(first + id) {
                                return Err("No se puede eliminar: la columna tiene un filtro aplicado.".to_string());
                            }
                            set("colId", (mv.line(first + id) - new_first).to_string());
                        }
                    } else if tag_is(name, "sortState") || tag_is(name, "sortCondition") {
                        if let Some(range) = attr_of(e, "ref").and_then(|r| CellRange::parse(&r).ok()) {
                            set("ref", mv.moved_span(&range).to_string());
                        }
                    } else if tag_is(name, "cellWatch") || tag_is(name, "inputCells") {
                        if let Some(at) = attr_of(e, "r").and_then(|r| CellRef::parse(&r).ok()) {
                            set("r", mv.at(at).to_string());
                        }
                    }
                }
                if matches!(event, Event::Start(_)) {
                    for (element, kind) in [("formula", "formula"), ("formula1", "formula"), ("formula2", "formula"), ("f", "f"), ("sqref", "sqref")] {
                        if tag_is(name, element) {
                            open = Some((kind, after));
                        }
                    }
                }
            }
            Event::End(e) => {
                let name = e.name();
                let name = name.as_ref();
                if tag_is(name, "conditionalFormatting") || tag_is(name, "dataValidation") {
                    delta = 0;
                } else if tag_is(name, "autoFilter") {
                    filter_first = None;
                }
                if let Some((kind, start)) = open.take() {
                    let raw = &tail[start..before];
                    let text = unescape_text(raw);
                    let written = match kind {
                        "formula" if host_is_target => target.anchored_formula(&text, delta),
                        // Nothing marks these for removal: without cells they keep their text
                        "sqref" if host_is_target => Some(moved_sqref(&text, mv)).filter(|t| !t.is_empty()).unwrap_or_else(|| text.to_string()),
                        "sqref" => text.to_string(),
                        _ => target.formula(&text, host_is_target),
                    };
                    if written != text {
                        edits.push((start..before, escape_xml(&written)));
                    }
                }
            }
            _ => {}
        }
    }
    Ok(apply_edits(tail, edits))
}

/// The whole worksheet whose lines move
fn moved_sheet(xml: &str, target: Target) -> Result<String, String> {
    let parsed = SheetXml::parse(xml)?;
    let Some(inner) = parsed.data_inner() else {
        // No cells at all: only what is around them
        let end = parsed.data_end();
        let mut head = if target.mv.axis == Axis::Cols { moved_cols(&xml[..end], target.mv)? } else { xml[..end].to_string() };
        head = moved_pane(&head, target.mv)?;
        return without_removed(format!("{}{}", head, moved_tail(&xml[end..], true, target)?));
    };
    let (rows, used) = moved_rows(&parsed, target)?;
    let mut head = xml[..inner.start].to_string();
    if target.mv.axis == Axis::Cols {
        head = moved_cols(&head, target.mv)?;
    }
    if let (Some(range), Some(used)) = (element_ranges(&head, "dimension", |_| true)?.into_iter().next(), used) {
        let element = set_attr(&head[range.clone()], "ref", &used.to_string());
        head.replace_range(range, &element);
    }
    head = moved_pane(&head, target.mv)?;
    without_removed(format!("{}{}{}", head, rows, moved_tail(&xml[inner.end..], true, target)?))
}

/// A worksheet that does not move but may refer to the one that does
fn other_sheet(xml: &str, target: Target) -> Result<Option<String>, String> {
    let needle = target.name.replace('\'', "''").to_lowercase();
    if !xml.to_lowercase().contains(&needle) {
        return Ok(None);
    }
    let parsed = SheetXml::parse(xml)?;
    let end = parsed.data_end();
    let tail = moved_tail(&xml[end..], false, target)?;
    let mut cells = std::collections::BTreeMap::new();
    let formulas = formula_edits(&parsed, false, target);
    for cell in parsed.rows.iter().flat_map(|r| &r.cells) {
        if let Some(edit) = formulas.get(&cell.at) {
            cells.insert(cell.at, Some(with_formula(&xml[cell.range.clone()], edit)?));
        }
    }
    if cells.is_empty() && tail == xml[end..] {
        return Ok(None);
    }
    let spliced = parsed.splice(&cells);
    let keep = spliced.len() - (xml.len() - end);
    Ok(Some(format!("{}{}", &spliced[..keep], tail)))
}

// ---------- Tables, drawings and comments ----------

/// The part of a table on the moved sheet (`host_is_target`) or on another one
fn moved_table(xml: &str, host_is_target: bool, target: Target) -> Result<Option<String>, String> {
    let mv = target.mv;
    let mut written = xml.to_string();
    if host_is_target {
        let Some((name, range, header, totals, _)) = parse_table(xml) else { return Ok(None) };
        let shape = if mv.kind == EditKind::Move {
            shape_move(&format!("la tabla {},", name), &range, header, totals, mv)?
        } else {
            resized_table(&name, &range, header, totals, mv)?
        };
        if shape == ShapeMove::Resized {
            written = table_resized(xml, &range, totals, mv)?;
        } else if shape == ShapeMove::Rigid {
            let mut edits = Vec::new();
            let mut reader = Reader::from_str(xml);
            loop {
                let before = reader.buffer_position() as usize;
                let event = reader.read_event().map_err(|e| format!("La tabla no es XML válido: {}", e))?;
                let after = reader.buffer_position() as usize;
                match &event {
                    Event::Eof => break,
                    Event::Start(e) | Event::Empty(e) => {
                        let tag = e.name();
                        let is_area = ["table", "autoFilter", "sortState", "sortCondition"].iter().any(|n| tag_is(tag.as_ref(), n));
                        if let Some(area) = is_area.then(|| attr_of(e, "ref").and_then(|r| CellRange::parse(&r).ok())).flatten() {
                            edits.push((before..after, set_attr(&xml[before..after], "ref", &mv.rigid(&area).to_string())));
                        }
                    }
                    _ => {}
                }
            }
            written = apply_edits(xml, edits);
        } else if mv.axis == Axis::Cols {
            written = table_columns_reordered(xml, &range, mv)?;
        }
    }
    let (written, _) = rewrite_texts(&written, &["calculatedColumnFormula", "totalsRowFormula"], &|f| target.formula(f, host_is_target))?;
    Ok((written != xml).then_some(written))
}

/// How a table takes lines inserted or deleted, or the refusal
fn resized_table(name: &str, range: &CellRange, header: u32, totals: u32, mv: &LineMove) -> Result<ShapeMove, String> {
    let (first, last) = mv.bounds(range);
    if mv.removes_all(first, last) {
        return Err(format!("No se puede eliminar: la {} contiene toda la tabla {}.", mv.axis.noun(), name));
    }
    if !mv.crosses(first, last) {
        return Ok(ShapeMove::Rigid);
    }
    if mv.axis == Axis::Cols {
        return Err(format!("No se pueden insertar ni eliminar columnas dentro de la tabla {}.", name));
    }
    if mv.kind == EditKind::Delete {
        let (body_first, body_last) = (first + header, last.saturating_sub(totals));
        if mv.first < body_first || mv.last > body_last {
            return Err(format!("No se puede eliminar el encabezado ni los totales de la tabla {}.", name));
        }
        if mv.first == body_first && mv.last == body_last {
            return Err(format!("No se pueden eliminar todas las filas de datos de la tabla {}.", name));
        }
    } else if mv.before < first + header {
        return Err(format!("No se puede insertar dentro del encabezado de la tabla {}.", name));
    }
    Ok(ShapeMove::Resized)
}

/// The table with its rows counted again: the filter reaches the last row of data
fn table_resized(xml: &str, range: &CellRange, totals: u32, mv: &LineMove) -> Result<String, String> {
    let resized = mv.moved_span(range);
    let filter = CellRange { first: resized.first, last: CellRef { row: resized.last.row - totals, col: resized.last.col } };
    let mut edits = Vec::new();
    let mut reader = Reader::from_str(xml);
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|e| format!("La tabla no es XML válido: {}", e))?;
        let after = reader.buffer_position() as usize;
        match &event {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) => {
                let tag = e.name();
                let tag = tag.as_ref();
                let written = if tag_is(tag, "table") {
                    Some(resized.to_string())
                } else if tag_is(tag, "autoFilter") {
                    Some(filter.to_string())
                } else if tag_is(tag, "sortState") || tag_is(tag, "sortCondition") {
                    attr_of(e, "ref").and_then(|r| CellRange::parse(&r).ok()).map(|r| mv.moved_span(&r).to_string())
                } else {
                    None
                };
                if let Some(written) = written {
                    edits.push((before..after, set_attr(&xml[before..after], "ref", &written)));
                }
            }
            _ => {}
        }
    }
    Ok(apply_edits(xml, edits))
}

/// The table with its columns in their new order: `tableColumn` goes by position, its id stays
fn table_columns_reordered(xml: &str, range: &CellRange, mv: &LineMove) -> Result<String, String> {
    let first = range.first.col;
    let slots = element_ranges(xml, "tableColumn", |_| true)?;
    let mut edits: Vec<(Range<usize>, String)> = Vec::new();
    for (index, slot) in slots.iter().enumerate() {
        let new_index = (mv.line(first + index as u32) - first) as usize;
        let destination = slots.get(new_index).ok_or("Las columnas de la tabla no coinciden con su rango.")?;
        edits.push((destination.clone(), xml[slot.clone()].to_string()));
    }
    let mut reader = Reader::from_str(xml);
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|e| format!("La tabla no es XML válido: {}", e))?;
        let after = reader.buffer_position() as usize;
        match &event {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if tag_is(e.name().as_ref(), "filterColumn") => {
                if let Some(id) = attr_of(e, "colId").and_then(|v| v.parse::<u32>().ok()) {
                    edits.push((before..after, set_attr(&xml[before..after], "colId", &(mv.line(first + id) - first).to_string())));
                }
            }
            _ => {}
        }
    }
    Ok(apply_edits(xml, edits))
}

/// A drawing part with its anchors following their first cell (keeping their size); those marked
/// `editAs="absolute"` do not move with cells
fn moved_drawing(xml: &str, mv: &LineMove) -> Result<Option<String>, String> {
    let line_name = match mv.axis {
        Axis::Rows => "row",
        Axis::Cols => "col",
    };
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut edits = Vec::new();
    let mut absolute = false;
    let mut corner: Option<bool> = None; // Some(true) inside <from>, Some(false) inside <to>
    let mut shift: i64 = 0;
    let mut open: Option<usize> = None;
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|e| format!("El dibujo no es XML válido: {}", e))?;
        let after = reader.buffer_position() as usize;
        match &event {
            Event::Eof => break,
            Event::Start(e) => {
                let name = e.name();
                let name = name.as_ref();
                if tag_is(name, "twoCellAnchor") || tag_is(name, "oneCellAnchor") {
                    absolute = get_attr_value(e, "editAs").as_deref() == Some("absolute");
                    shift = 0;
                } else if tag_is(name, "from") {
                    corner = Some(true);
                } else if tag_is(name, "to") {
                    corner = Some(false);
                } else if corner.is_some() && tag_is(name, line_name) {
                    open = Some(after);
                }
            }
            Event::End(e) => {
                let name = e.name();
                let name = name.as_ref();
                if tag_is(name, "from") || tag_is(name, "to") {
                    corner = None;
                } else if let Some(start) = open.take() {
                    let Ok(value) = xml[start..before].trim().parse::<u32>() else { continue };
                    if absolute {
                        continue;
                    }
                    if corner == Some(true) {
                        // What was anchored in deleted lines stays, in the line that takes their place
                        let line = if mv.removes(value + 1) { mv.first } else { mv.line(value + 1) };
                        shift = line as i64 - (value as i64 + 1);
                    }
                    if shift != 0 {
                        edits.push((start..before, (value as i64 + shift).max(0).to_string()));
                    }
                }
            }
            _ => {}
        }
    }
    Ok((!edits.is_empty()).then(|| apply_edits(xml, edits)))
}

/// A part whose elements `name` say in `ref` which cell they belong to (comments)
fn moved_cell_refs(xml: &str, name: &str, mv: &LineMove) -> Result<Option<String>, String> {
    let removed = |at: CellRef| match mv.axis {
        Axis::Rows => mv.removes(at.row),
        Axis::Cols => mv.removes(at.col),
    };
    let kept = remove_elements(xml, name, |e| attr_of(e, "ref").and_then(|r| CellRef::parse(&r).ok()).is_some_and(removed))?;
    let changed = kept != xml;
    let xml = kept.as_str();
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut edits = Vec::new();
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|e| format!("Los comentarios no son XML válido: {}", e))?;
        let after = reader.buffer_position() as usize;
        match &event {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if tag_is(e.name().as_ref(), name) => {
                let Some(at) = attr_of(e, "ref").and_then(|r| CellRef::parse(&r).ok()) else { continue };
                let moved = mv.at(at);
                if moved != at {
                    edits.push((before..after, set_attr(&xml[before..after], "ref", &moved.to_string())));
                }
            }
            _ => {}
        }
    }
    Ok((changed || !edits.is_empty()).then(|| apply_edits(xml, edits)))
}

/// The shapes of old-style notes say their cell in `<x:Row>` / `<x:Column>`, from 0. VML is not
/// always well-formed XML: these elements are found as text.
fn moved_vml(xml: &str, mv: &LineMove) -> Option<String> {
    let (open, close) = match mv.axis {
        Axis::Rows => ("<x:Row>", "</x:Row>"),
        Axis::Cols => ("<x:Column>", "</x:Column>"),
    };
    let mut edits = Vec::new();
    let mut from = 0;
    while let Some(at) = xml[from..].find(open) {
        let start = from + at + open.len();
        let Some(length) = xml[start..].find(close) else { break };
        let end = start + length;
        if let Ok(value) = xml[start..end].trim().parse::<u32>() {
            if mv.removes(value + 1) {
                // The note goes with its cell: the whole shape around these elements
                const CLOSE: &str = "</v:shape>";
                if let (Some(shape), Some(length)) = (xml[..start].rfind("<v:shape"), xml[end..].find(CLOSE)) {
                    edits.push((shape..end + length + CLOSE.len(), String::new()));
                    from = end + length + CLOSE.len();
                    continue;
                }
            }
            let moved = mv.line(value + 1) - 1;
            if moved != value {
                edits.push((start..end, moved.to_string()));
            }
        }
        from = end;
    }
    (!edits.is_empty()).then(|| apply_edits(xml, edits))
}

impl Workbook {
    /// Moves the rows `first..first + count` before the row `before` (numbered as they are now):
    /// see the module documentation. Returns where the first of them ends up. Nothing changes when
    /// the move is refused.
    pub fn move_rows(&mut self, sheet: usize, first: u32, count: u32, before: u32) -> Result<u32, String> {
        self.move_lines(sheet, LineMove::new(Axis::Rows, first, count, before)?)
    }

    /// Moves the columns `first..first + count` before the column `before` (see `move_rows`)
    pub fn move_columns(&mut self, sheet: usize, first: u32, count: u32, before: u32) -> Result<u32, String> {
        self.move_lines(sheet, LineMove::new(Axis::Cols, first, count, before)?)
    }

    /// `count` empty rows before the row `before`, with the height and format of the one above
    /// (see the module documentation). Returns `before`.
    pub fn insert_rows(&mut self, sheet: usize, before: u32, count: u32) -> Result<u32, String> {
        self.move_lines(sheet, LineMove::insert(Axis::Rows, before, count)?)
    }

    /// `count` empty columns before the column `before`, like the one to their left
    pub fn insert_columns(&mut self, sheet: usize, before: u32, count: u32) -> Result<u32, String> {
        self.move_lines(sheet, LineMove::insert(Axis::Cols, before, count)?)
    }

    /// Deletes `count` rows from `first`; what pointed only at them becomes `#REF!`
    pub fn delete_rows(&mut self, sheet: usize, first: u32, count: u32) -> Result<u32, String> {
        self.move_lines(sheet, LineMove::delete(Axis::Rows, first, count)?)
    }

    /// Deletes `count` columns from `first` (see `delete_rows`)
    pub fn delete_columns(&mut self, sheet: usize, first: u32, count: u32) -> Result<u32, String> {
        self.move_lines(sheet, LineMove::delete(Axis::Cols, first, count)?)
    }

    fn move_lines(&mut self, sheet: usize, mv: LineMove) -> Result<u32, String> {
        let part = self.worksheet_part(sheet)?;
        if mv.is_noop() {
            return Ok(mv.result_line());
        }
        let sheets = self.sheets()?;
        let name = sheets[sheet].name.clone();
        let target = Target { name: &name, mv: &mv };
        let rels = |part: &str| -> Vec<(String, String)> {
            self.package
                .get_string(&rels_part_for(part))
                .map(|xml| parse_relationships(&xml))
                .unwrap_or_default()
                .into_iter()
                .filter(|r| !r.external)
                .map(|r| (r.rel_type.clone(), resolve_target(part, &r.target)))
                .filter(|(_, target)| self.package.contains(target))
                .collect()
        };
        let own = rels(&part);
        if own.iter().any(|(kind, _)| type_is(kind, "pivotTable")) {
            return Err(format!("No se pueden mover, insertar ni eliminar filas o columnas de la hoja '{}': tiene tablas dinámicas.", name));
        }

        // Everything is worked out first, so a refusal leaves the workbook untouched
        let mut changes: Vec<(String, String)> = vec![(part.clone(), moved_sheet(&self.package.get_string(&part)?, target)?)];
        for (index, info) in sheets.iter().enumerate() {
            if info.kind != SheetKind::Worksheet || !self.package.contains(&info.part) {
                continue;
            }
            if index != sheet {
                if let Some(xml) = other_sheet(&self.package.get_string(&info.part)?, target)? {
                    changes.push((info.part.clone(), xml));
                }
            }
            for (_, table) in rels(&info.part).into_iter().filter(|(kind, _)| type_is(kind, "table")) {
                if let Some(xml) = moved_table(&self.package.get_string(&table)?, index == sheet, target)? {
                    changes.push((table, xml));
                }
            }
        }
        for (kind, related) in &own {
            let xml = if type_is(kind, "drawing") {
                moved_drawing(&self.package.get_string(related)?, &mv)?
            } else if type_is(kind, "comments") {
                moved_cell_refs(&self.package.get_string(related)?, "comment", &mv)?
            } else if type_is(kind, "threadedComment") {
                moved_cell_refs(&self.package.get_string(related)?, "threadedComment", &mv)?
            } else if type_is(kind, "vmlDrawing") {
                moved_vml(&self.package.get_string(related)?, &mv)
            } else {
                None
            };
            if let Some(xml) = xml {
                changes.push((related.clone(), xml));
            }
        }
        let (workbook, _) = rewrite_texts(&self.workbook_xml()?, &["definedName"], &|f| target.formula(f, false))?;
        changes.push((self.workbook_part.clone(), workbook));
        let mut others: Vec<String> = self
            .package
            .parts()
            .keys()
            .filter(|p| p.ends_with(".xml") && !p.contains("/_rels/") && (p.starts_with("xl/charts/") || p.starts_with("xl/pivotCache/pivotCacheDefinition")))
            .cloned()
            .collect();
        others.sort();
        for other in others {
            let xml = self.package.get_string(&other)?;
            let written = if other.starts_with("xl/charts/") {
                rewrite_texts(&xml, FORMULA_ELEMENTS, &|f| target.formula(f, false))?.0
            } else {
                moved_pivot_source(&xml, target)?
            };
            if written != xml {
                changes.push((other, written));
            }
        }

        for (changed, xml) in changes {
            self.package.put(changed, xml.into_bytes());
        }
        self.recalculate_after_edit();
        self.prepare_recalculation()?;
        Ok(mv.result_line())
    }
}

/// A pivot cache whose data comes from a range of the moved sheet
fn moved_pivot_source(xml: &str, target: Target) -> Result<String, String> {
    let mut edits = Vec::new();
    let mut reader = Reader::from_str(xml);
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(|e| format!("La caché no es XML válido: {}", e))?;
        let after = reader.buffer_position() as usize;
        match &event {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if tag_is(e.name().as_ref(), "worksheetSource") => {
                let on_target = attr_of(e, "sheet").is_some_and(|s| target.names(Some(&s), false));
                if let (true, Some(range)) = (on_target, attr_of(e, "ref").and_then(|r| CellRange::parse(&r).ok())) {
                    edits.push((before..after, set_attr(&xml[before..after], "ref", &target.mv.moved_span(&range).to_string())));
                }
            }
            _ => {}
        }
    }
    Ok(apply_edits(xml, edits))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::workbook::{CellInput, CellValue};

    fn blank() -> Workbook {
        Workbook::from_bytes(&crate::blank::blank_xlsx().unwrap()).unwrap()
    }

    fn number(n: f64) -> Option<CellInput> {
        Some(CellInput::Number(n))
    }

    fn formula(f: &str) -> Option<CellInput> {
        Some(CellInput::Formula(f.into()))
    }

    fn sheet_xml(wb: &Workbook) -> String {
        wb.package.get_string("xl/worksheets/sheet1.xml").unwrap()
    }

    /// Replaces the sheet's XML after `</sheetData>` start with `tail`, and keeps the data
    fn with_tail(wb: &mut Workbook, tail: &str) {
        let xml = sheet_xml(wb);
        let at = xml.find("</sheetData>").unwrap() + "</sheetData>".len();
        let end = xml.rfind("</worksheet>").unwrap();
        let written = format!("{}{}{}", &xml[..at], tail, &xml[end..]);
        wb.package.put("xl/worksheets/sheet1.xml".into(), written.into_bytes());
    }

    fn formula_at(wb: &Workbook, sheet: usize, at: &str) -> String {
        wb.cell(sheet, at).unwrap().formula.unwrap_or_default()
    }

    fn value(wb: &Workbook, at: &str) -> CellValue {
        wb.cell(0, at).unwrap().value
    }

    fn rows_mv(first: u32, count: u32, before: u32) -> LineMove {
        LineMove::new(Axis::Rows, first, count, before).unwrap()
    }

    #[test]
    fn lines_go_where_they_should() {
        // Row 2 before row 5: 3, 4 go up
        let down = rows_mv(2, 1, 5);
        assert_eq!([1, 2, 3, 4, 5].map(|i| down.line(i)), [1, 4, 2, 3, 5]);
        // Rows 4–5 before row 2: 2, 3 go down
        let up = rows_mv(4, 2, 2);
        assert_eq!([1, 2, 3, 4, 5, 6].map(|i| up.line(i)), [1, 4, 5, 2, 3, 6]);
        assert!(rows_mv(2, 2, 3).is_noop() && rows_mv(2, 2, 4).is_noop() && rows_mv(2, 2, 2).is_noop());
        assert!(!rows_mv(2, 2, 5).is_noop());
        assert!(LineMove::new(Axis::Rows, 0, 1, 2).is_err());
        assert!(LineMove::new(Axis::Rows, 1, 0, 2).is_err());
        assert!(LineMove::new(Axis::Cols, 16_384, 2, 1).is_err());
        assert!(LineMove::new(Axis::Cols, 1, 1, 16_386).is_err());
    }

    #[test]
    fn ranges_keep_their_lines() {
        let span = |first, count, before, a, b| rows_mv(first, count, before).span(a, b);
        // Inside it, also to its ends: the same set of rows
        assert_eq!(span(3, 1, 8, 1, 10), (1, 10));
        assert_eq!(span(5, 1, 2, 2, 10), (2, 10));
        assert_eq!(span(5, 1, 11, 2, 10), (2, 10));
        assert_eq!(span(2, 1, 11, 2, 10), (2, 10));
        // Leaving shrinks, entering grows
        assert_eq!(span(3, 1, 20, 1, 10), (1, 9));
        assert_eq!(span(20, 2, 3, 1, 10), (1, 12));
        // Elsewhere: only pushed
        assert_eq!(span(20, 1, 1, 5, 10), (6, 11));
        assert_eq!(span(1, 1, 20, 5, 10), (4, 9));
        assert_eq!(span(20, 1, 30, 5, 10), (5, 10));
        // Inside the block: goes with it
        assert_eq!(span(3, 4, 20, 4, 5), (17, 18));
        // Part of it leaves with the block
        assert_eq!(span(8, 5, 30, 5, 10), (5, 7));
    }

    #[test]
    fn rewrites_every_kind_of_reference_and_nothing_else() {
        let wb_name = "Datos";
        let mv = rows_mv(2, 1, 5);
        let target = Target { name: wb_name, mv: &mv };
        let cases = [
            ("A2*2", "A4*2"),
            ("$A$2+A$3+$A3", "$A$4+A$2+$A2"),
            ("SUM(A1:B9)", "SUM(A1:B9)"),
            ("SUM(A2:A3)", "SUM(A2:A2)"),
            (r#"IF(A2="A2","A2",LOG10(A2))"#, r#"IF(A4="A2","A2",LOG10(A4))"#),
            ("Datos!A2+Otra!A2", "Datos!A4+Otra!A2"),
            ("'Datos'!A2", "'Datos'!A4"),
            ("datos!A2", "datos!A4"),
            ("[1]Datos!A2", "[1]Datos!A2"),
            ("Datos:Otra!A2", "Datos:Otra!A2"),
            ("Tabla1[A2]+Tabla1[[#This Row],[B3]]", "Tabla1[A2]+Tabla1[[#This Row],[B3]]"),
            ("SUM(A:C)", "SUM(A:C)"),
            ("SUM(2:2)+SUM($2:$3)", "SUM(4:4)+SUM($2:$2)"),
            ("#REF!+A2", "#REF!+A4"),
            ("1.5E2+Nombre2_x+A2", "1.5E2+Nombre2_x+A4"),
            ("TRUE", "TRUE"),
            ("(A2,A3)", "(A4,A2)"),
        ];
        for (input, expected) in cases {
            assert_eq!(target.formula(input, true), expected, "{}", input);
        }
        // From another sheet only what names it
        assert_eq!(target.formula("A2+Datos!A2", false), "A2+Datos!A4");

        let cols = LineMove::new(Axis::Cols, 1, 1, 4).unwrap();
        let target = Target { name: wb_name, mv: &cols };
        assert_eq!(target.formula("A1+B$1+SUM(A:A)+SUM($B:$C)+SUM(1:1)", true), "C1+A$1+SUM(C:C)+SUM($A:$B)+SUM(1:1)");
    }

    #[test]
    fn shifts_shared_formulas_as_text() {
        assert_eq!(shift_formula("A1+$B$1+Otra!C2+[1]H!A1+SUM(A:A)", 2, 1), "B3+$B$1+Otra!D4+[1]H!B3+SUM(B:B)");
        assert_eq!(shift_formula("A1", -1, 0), "#REF!");
    }

    #[test]
    fn moves_rows_with_their_format_and_formulas() {
        let mut wb = blank();
        wb.set_values(
            0,
            "A1",
            vec![
                vec![Some(CellInput::Text("Producto".into())), Some(CellInput::Text("Total".into()))],
                vec![Some(CellInput::Text("uno".into())), number(1.0), formula("B2*10")],
                vec![Some(CellInput::Text("dos".into())), number(2.0), formula("B3*10")],
                vec![Some(CellInput::Text("tres".into())), number(3.0), formula("B4*10")],
                vec![Some(CellInput::Text("suma".into())), formula("SUM(B2:B4)"), formula("C2+C4")],
            ],
        )
        .unwrap();
        wb.set_format(0, "A4:B4", &crate::format::Format { bold: Some(true), ..Default::default() }).unwrap();
        let xml = sheet_xml(&wb).replace(r#"<row r="4""#, r#"<row r="4" ht="30" customHeight="1""#);
        wb.package.put("xl/worksheets/sheet1.xml".into(), xml.into_bytes());
        let style = wb.cell_views(0, &crate::numfmt::Locale::from_tag("es")).unwrap().into_iter().find(|c| c.row == 4 && c.col == 1).unwrap().style;

        wb.checkpoint(None);
        assert_eq!(wb.move_rows(0, 4, 1, 2).unwrap(), 2);
        assert_eq!(value(&wb, "A2"), CellValue::Text("tres".into()));
        assert_eq!(value(&wb, "A3"), CellValue::Text("uno".into()));
        assert_eq!(value(&wb, "A4"), CellValue::Text("dos".into()));
        assert_eq!(formula_at(&wb, 0, "C2"), "B2*10");
        assert_eq!(formula_at(&wb, 0, "C3"), "B3*10");
        assert_eq!(value(&wb, "C2"), CellValue::Number(30.0));
        // The same rows: the total does not change
        assert_eq!(formula_at(&wb, 0, "B5"), "SUM(B2:B4)");
        assert_eq!(value(&wb, "B5"), CellValue::Number(6.0));
        // C2 was the row "uno", C4 the row "tres": each reference follows its cell
        assert_eq!(formula_at(&wb, 0, "C5"), "C3+C2");
        assert_eq!(value(&wb, "C5"), CellValue::Number(40.0));
        let xml = sheet_xml(&wb);
        assert!(xml.contains(r#"<row r="2" ht="30" customHeight="1""#), "{}", xml);
        let moved = wb.cell_views(0, &crate::numfmt::Locale::from_tag("es")).unwrap().into_iter().find(|c| c.row == 2 && c.col == 1).unwrap().style;
        assert_eq!(moved, style);
        // Rows stay in order in the file
        let numbers: Vec<u32> = SheetXml::parse(&xml).unwrap().rows.iter().map(|r| r.number).collect();
        assert!(numbers.windows(2).all(|w| w[0] < w[1]), "{:?}", numbers);

        // One undo step, and back the same
        wb.undo().unwrap();
        assert_eq!(value(&wb, "A2"), CellValue::Text("uno".into()));
        assert_eq!(formula_at(&wb, 0, "C5"), "C2+C4");
        wb.redo().unwrap();
        assert_eq!(value(&wb, "A2"), CellValue::Text("tres".into()));
        // Still a workbook after a round trip
        let again = Workbook::from_bytes(&wb.to_bytes().unwrap()).unwrap();
        assert_eq!(again.cell(0, "A2").unwrap().value, CellValue::Text("tres".into()));
    }

    #[test]
    fn moves_columns_with_their_widths_and_references() {
        let mut wb = blank();
        wb.set_values(0, "A1", vec![vec![number(1.0), number(2.0), number(3.0), formula("A1+B1*C1")], vec![number(4.0), formula("A2*2")]]).unwrap();
        wb.set_column_width(0, 1, 1, 30.0).unwrap();
        wb.set_column_width(0, 3, 16_384, 12.0).unwrap();
        let xml = sheet_xml(&wb).replace(r#"<row r="1""#, r#"<row r="1" spans="1:4""#);
        wb.package.put("xl/worksheets/sheet1.xml".into(), xml.into_bytes());

        // A goes after C: B C A D
        assert_eq!(wb.move_columns(0, 1, 1, 4).unwrap(), 3);
        assert_eq!(value(&wb, "A1"), CellValue::Number(2.0));
        assert_eq!(value(&wb, "B1"), CellValue::Number(3.0));
        assert_eq!(value(&wb, "C1"), CellValue::Number(1.0));
        assert_eq!(formula_at(&wb, 0, "D1"), "C1+A1*B1");
        assert_eq!(value(&wb, "D1"), CellValue::Number(7.0));
        assert_eq!(formula_at(&wb, 0, "A2"), "C2*2");
        assert_eq!(value(&wb, "A2"), CellValue::Number(8.0));
        let xml = sheet_xml(&wb);
        assert!(!xml.contains("spans"), "{}", xml);
        let layout = wb.sheet_layout(0).unwrap();
        let width_of = |c: u32| layout.cols.iter().find(|(a, b, ..)| *a <= c && c <= *b).and_then(|x| x.2);
        assert_eq!(width_of(3), Some(30.0));
        assert_eq!(width_of(2), Some(12.0));
        assert_eq!(width_of(4), Some(12.0));
        assert_eq!(width_of(16_384), Some(12.0));
        assert_eq!(width_of(1), None);
    }

    #[test]
    fn shared_formulas_get_their_own_text_when_they_change() {
        let mut wb = blank();
        let xml = sheet_xml(&wb);
        let data = r#"<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row><row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row><row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row></sheetData>"#;
        let start = xml.find("<sheetData").unwrap();
        let end = xml.find("</worksheet>").unwrap();
        wb.package.put("xl/worksheets/sheet1.xml".into(), format!("{}{}{}", &xml[..start], data, &xml[end..]).into_bytes());

        wb.move_rows(0, 3, 1, 1).unwrap();
        assert_eq!(formula_at(&wb, 0, "B1"), "A1*2");
        assert_eq!(formula_at(&wb, 0, "B2"), "A2*2");
        assert_eq!(formula_at(&wb, 0, "B3"), "A3*2");
        assert_eq!(value(&wb, "B1"), CellValue::Number(6.0));
        assert_eq!(value(&wb, "B2"), CellValue::Number(2.0));
        assert!(!sheet_xml(&wb).contains("shared"), "{}", sheet_xml(&wb));
    }

    #[test]
    fn other_sheets_and_names_follow() {
        let mut wb = blank();
        wb.add_sheet("Resumen", None).unwrap();
        wb.set_values(0, "A1", vec![vec![number(1.0)], vec![number(2.0)], vec![number(3.0)]]).unwrap();
        wb.set_values(1, "A1", vec![vec![formula("Hoja1!A3*2"), formula("A1")]]).unwrap();
        let xml = wb.workbook_xml().unwrap().replace("</sheets>", r#"</sheets><definedNames><definedName name="Ultima">Hoja1!$A$3</definedName><definedName name="Otra">Resumen!$A$3</definedName></definedNames>"#);
        wb.package.put(wb.workbook_part.clone(), xml.into_bytes());
        wb.package.put("xl/charts/chart1.xml".into(), br#"<c:chartSpace xmlns:c="c"><c:f>Hoja1!$A$1:$A$2</c:f><c:f>Resumen!$A$1</c:f></c:chartSpace>"#.to_vec());

        wb.move_rows(0, 3, 1, 1).unwrap();
        assert_eq!(formula_at(&wb, 1, "A1"), "Hoja1!A1*2");
        assert_eq!(formula_at(&wb, 1, "B1"), "A1");
        assert_eq!(wb.cell(1, "A1").unwrap().value, CellValue::Number(6.0));
        let book = wb.workbook_xml().unwrap();
        assert!(book.contains(">Hoja1!$A$1<"), "{}", book);
        assert!(book.contains(">Resumen!$A$3<"), "{}", book);
        let chart = wb.package.get_string("xl/charts/chart1.xml").unwrap();
        assert!(chart.contains("<c:f>Hoja1!$A$2:$A$3</c:f><c:f>Resumen!$A$1</c:f>"), "{}", chart);
    }

    #[test]
    fn refuses_to_tear_merged_cells_and_arrays() {
        let mut wb = blank();
        wb.set_values(0, "A1", vec![vec![number(1.0)], vec![number(2.0)], vec![number(3.0)], vec![number(4.0)]]).unwrap();
        with_tail(&mut wb, r#"<mergeCells count="2"><mergeCell ref="A2:B3"/><mergeCell ref="C1:F1"/></mergeCells>"#);
        let before = sheet_xml(&wb);
        wb.checkpoint(None);
        let error = wb.move_rows(0, 3, 1, 1).unwrap_err();
        assert!(error.contains("A2:B3"), "{}", error);
        assert!(wb.move_rows(0, 1, 1, 3).is_err(), "into the middle");
        assert_eq!(sheet_xml(&wb), before);
        assert!(!wb.can_undo(), "a refused move leaves no undo step");

        // The whole merge moves
        wb.move_rows(0, 2, 2, 6).unwrap();
        assert!(sheet_xml(&wb).contains(r#"<mergeCell ref="A4:B5"/>"#), "{}", sheet_xml(&wb));
        assert!(sheet_xml(&wb).contains(r#"<mergeCell ref="C1:F1"/>"#));
        // Columns under a merged title, not its first one
        wb.move_columns(0, 4, 1, 6).unwrap();
        assert!(sheet_xml(&wb).contains(r#"<mergeCell ref="C1:F1"/>"#));
        assert!(wb.move_columns(0, 3, 1, 6).is_err(), "its first column holds the value");
        assert!(wb.move_columns(0, 8, 1, 4).is_err(), "from outside");

        let mut wb = blank();
        let xml = sheet_xml(&wb);
        let start = xml.find("<sheetData").unwrap();
        let end = xml.find("</worksheet>").unwrap();
        let data = r#"<sheetData><row r="1"><c r="A1"><f t="array" ref="A1:A2">B1:B2*2</f><v>0</v></c></row><row r="2"><c r="A2"><v>0</v></c></row><row r="3"><c r="B3"><f t="array" ref="B3">SUM(A1:A2*2)</f><v>0</v></c></row></sheetData>"#;
        wb.package.put("xl/worksheets/sheet1.xml".into(), format!("{}{}{}", &xml[..start], data, &xml[end..]).into_bytes());
        assert!(wb.move_rows(0, 1, 1, 4).unwrap_err().contains("matricial"));
        wb.move_rows(0, 3, 1, 1).unwrap();
        let xml = sheet_xml(&wb);
        assert!(xml.contains(r#"<f t="array" ref="B1">SUM(A2:A3*2)</f>"#), "{}", xml);
        assert!(xml.contains(r#"<f t="array" ref="A2:A3">B2:B3*2</f>"#), "{}", xml);
    }

    #[test]
    fn formats_validations_and_links_follow_their_cells() {
        let mut wb = blank();
        wb.set_values(0, "A1", vec![vec![number(1.0)], vec![number(2.0)], vec![number(3.0)]]).unwrap();
        with_tail(
            &mut wb,
            r#"<conditionalFormatting sqref="A1:C5"><cfRule type="expression" priority="1"><formula>$A1&gt;$F$2</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="A2"><cfRule type="expression" priority="2"><formula>A2&gt;0</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="list" sqref="D2"><formula1>$F$1:$F$3</formula1></dataValidation></dataValidations><hyperlinks><hyperlink ref="B2" location="Hoja1!A2"/></hyperlinks>"#,
        );
        wb.move_rows(0, 2, 1, 4).unwrap();
        let xml = sheet_xml(&wb);
        // The rule keeps its rows; its relative reference still looks at each row's own A
        assert!(xml.contains(r#"sqref="A1:C5"><cfRule type="expression" priority="1"><formula>$A1&gt;$F$3</formula>"#), "{}", xml);
        // The rule of A2 goes with the cell, looking at it from there
        assert!(xml.contains(r#"sqref="A3"><cfRule type="expression" priority="2"><formula>A3&gt;0</formula>"#), "{}", xml);
        // $F$1:$F$3 holds the same lines
        assert!(xml.contains(r#"sqref="D3"><formula1>$F$1:$F$3</formula1>"#), "{}", xml);
        assert!(xml.contains(r#"<hyperlink ref="B3" location="Hoja1!A3"/>"#), "{}", xml);

        // A rule over rows that scatter is split
        let mut wb = blank();
        wb.set_value(0, "A1", CellInput::Number(1.0)).unwrap();
        with_tail(&mut wb, r#"<conditionalFormatting sqref="A1:A3"><cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule></conditionalFormatting>"#);
        wb.move_rows(0, 2, 1, 10).unwrap();
        assert!(sheet_xml(&wb).contains(r#"sqref="A1:A2 A9""#), "{}", sheet_xml(&wb));
    }

    fn with_table(wb: &mut Workbook, table: &str) {
        wb.package.put("xl/tables/table1.xml".into(), table.as_bytes().to_vec());
        wb.package.put(
            "xl/worksheets/_rels/sheet1.xml.rels".into(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#.to_vec(),
        );
    }

    const TABLE: &str = r#"<table xmlns="m" id="1" name="T" displayName="T" ref="B2:D5" totalsRowCount="1"><autoFilter ref="B2:D4"><filterColumn colId="0"/></autoFilter><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"><calculatedColumnFormula>F3*2</calculatedColumnFormula></tableColumn><tableColumn id="3" name="c"/></tableColumns></table>"#;

    #[test]
    fn tables_move_whole_or_reorder_inside() {
        let mut wb = blank();
        with_table(&mut wb, TABLE);
        wb.set_value(0, "A1", CellInput::Number(1.0)).unwrap();
        for (first, count, before) in [(2, 1, 4), (3, 1, 10), (10, 1, 4), (5, 1, 3), (4, 1, 6), (3, 1, 2)] {
            let error = wb.move_rows(0, first, count, before).unwrap_err();
            assert!(error.contains("la tabla T"), "{} {} {}: {}", first, count, before, error);
        }
        assert!(wb.move_columns(0, 3, 1, 9).unwrap_err().contains("la tabla T"));

        // Data rows change places inside it: the table stays
        wb.move_rows(0, 4, 1, 3).unwrap();
        assert_eq!(wb.package.get_string("xl/tables/table1.xml").unwrap().replace("F4*2", "F3*2"), TABLE);

        // Columns in the table: the order of tableColumn and the filter's positions follow
        wb.move_columns(0, 2, 1, 5).unwrap();
        let table = wb.package.get_string("xl/tables/table1.xml").unwrap();
        assert!(table.contains(r#"ref="B2:D5""#), "{}", table);
        let order: Vec<_> = ["name=\"b\"", "name=\"c\"", "name=\"a\""].iter().map(|n| table.find(n).unwrap()).collect();
        assert!(order.windows(2).all(|w| w[0] < w[1]), "{}", table);
        assert!(table.contains(r#"<filterColumn colId="2"/>"#), "{}", table);

        // The whole table
        wb.move_rows(0, 2, 4, 1).unwrap();
        let table = wb.package.get_string("xl/tables/table1.xml").unwrap();
        assert!(table.contains(r#"ref="B1:D4""#) && table.contains(r#"<autoFilter ref="B1:D3">"#), "{}", table);
    }

    #[test]
    fn drawings_comments_and_pivots() {
        let mut wb = blank();
        wb.set_value(0, "B3", CellInput::Number(1.0)).unwrap();
        wb.package.put(
            "xl/drawings/drawing1.xml".into(),
            br#"<xdr:wsDr xmlns:xdr="d"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:row>4</xdr:row></xdr:to></xdr:twoCellAnchor><xdr:twoCellAnchor editAs="absolute"><xdr:from><xdr:row>2</xdr:row></xdr:from><xdr:to><xdr:row>4</xdr:row></xdr:to></xdr:twoCellAnchor></xdr:wsDr>"#.to_vec(),
        );
        wb.package.put("xl/comments1.xml".into(), br#"<comments xmlns="m"><commentList><comment ref="B3" authorId="0"/><comment ref="B9" authorId="0"/></commentList></comments>"#.to_vec());
        wb.package.put("xl/drawings/vmlDrawing1.vml".into(), b"<xml><x:ClientData><x:Row>2</x:Row><x:Column>1</x:Column></x:ClientData></xml>".to_vec());
        wb.package.put(
            "xl/worksheets/_rels/sheet1.xml.rels".into(),
            br#"<Relationships xmlns="r"><Relationship Id="rId1" Type="x/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="rId2" Type="x/comments" Target="../comments1.xml"/><Relationship Id="rId3" Type="x/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>"#.to_vec(),
        );
        wb.move_rows(0, 3, 1, 1).unwrap();
        let drawing = wb.package.get_string("xl/drawings/drawing1.xml").unwrap();
        assert!(drawing.contains("<xdr:col>1</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:row>2</xdr:row>"), "{}", drawing);
        assert!(drawing.contains(r#"editAs="absolute"><xdr:from><xdr:row>2</xdr:row></xdr:from><xdr:to><xdr:row>4</xdr:row>"#), "{}", drawing);
        let comments = wb.package.get_string("xl/comments1.xml").unwrap();
        assert!(comments.contains(r#"<comment ref="B1" authorId="0"/><comment ref="B9" authorId="0"/>"#), "{}", comments);
        assert!(wb.package.get_string("xl/drawings/vmlDrawing1.vml").unwrap().contains("<x:Row>0</x:Row><x:Column>1</x:Column>"));

        let mut wb = blank();
        wb.package.put("xl/pivotTables/pivotTable1.xml".into(), b"<pivotTableDefinition/>".to_vec());
        wb.package.put(
            "xl/worksheets/_rels/sheet1.xml.rels".into(),
            br#"<Relationships xmlns="r"><Relationship Id="rId1" Type="x/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>"#.to_vec(),
        );
        assert!(wb.move_rows(0, 1, 1, 3).unwrap_err().contains("dinámicas"));
    }

    #[test]
    fn inserts_rows_growing_ranges_and_taking_the_format_above() {
        let mut wb = blank();
        wb.set_values(0, "A1", vec![vec![number(1.0), number(10.0)], vec![number(2.0), number(20.0)], vec![number(3.0), formula("SUM(B1:B2)+A2")]]).unwrap();
        wb.set_format(0, "A2:B2", &crate::format::Format { bold: Some(true), ..Default::default() }).unwrap();
        let xml = sheet_xml(&wb).replace(r#"<row r="2""#, r#"<row r="2" ht="30" customHeight="1""#);
        wb.package.put("xl/worksheets/sheet1.xml".into(), xml.into_bytes());
        with_tail(&mut wb, r#"<mergeCells count="1"><mergeCell ref="D1:D3"/></mergeCells><conditionalFormatting sqref="A1:A3"><cfRule type="expression" priority="1"><formula>$A$2&gt;0</formula></cfRule></conditionalFormatting>"#);

        wb.checkpoint(None);
        assert_eq!(wb.insert_rows(0, 2, 2).unwrap(), 2);
        assert_eq!(value(&wb, "A4"), CellValue::Number(2.0));
        assert_eq!(value(&wb, "A2"), CellValue::Empty);
        assert_eq!(formula_at(&wb, 0, "B5"), "SUM(B1:B4)+A4");
        assert_eq!(value(&wb, "B5"), CellValue::Number(32.0));
        let xml = sheet_xml(&wb);
        assert!(xml.contains(r#"<mergeCell ref="D1:D5"/>"#), "{}", xml);
        assert!(xml.contains(r#"sqref="A1:A5"><cfRule type="expression" priority="1"><formula>$A$4&gt;0</formula>"#), "{}", xml);
        let numbers: Vec<u32> = SheetXml::parse(&xml).unwrap().rows.iter().map(|r| r.number).collect();
        assert_eq!(numbers, [1, 2, 3, 4, 5]);

        // At the first row there is no format to take
        let mut other = blank();
        other.set_format(0, "A1", &crate::format::Format { bold: Some(true), ..Default::default() }).unwrap();
        other.insert_rows(0, 1, 1).unwrap();
        assert!(!sheet_xml(&other).contains(r#"<c r="A1""#), "{}", sheet_xml(&other));

        // Below the row of height 30: they take it, as the format of the row above
        let mut wb2 = blank();
        wb2.set_values(0, "A1", vec![vec![number(1.0)], vec![number(2.0)]]).unwrap();
        wb2.set_format(0, "A1", &crate::format::Format { bold: Some(true), ..Default::default() }).unwrap();
        let xml = sheet_xml(&wb2).replace(r#"<row r="1""#, r#"<row r="1" ht="30" customHeight="1""#);
        wb2.package.put("xl/worksheets/sheet1.xml".into(), xml.into_bytes());
        wb2.insert_rows(0, 2, 1).unwrap();
        let xml = sheet_xml(&wb2);
        let style = SheetXml::parse(&xml).unwrap().cell(CellRef::parse("A1").unwrap()).unwrap().style;
        assert!(xml.contains(&format!(r#"<row r="2" ht="30" customHeight="1"><c r="A2" s="{}"/></row>"#, style)), "{}", xml);
        assert_eq!(value(&wb2, "A2"), CellValue::Empty);
        assert_eq!(value(&wb2, "A3"), CellValue::Number(2.0));

        wb.undo().unwrap();
        assert_eq!(value(&wb, "A2"), CellValue::Number(2.0));
    }

    #[test]
    fn inserts_columns_with_the_width_on_their_left() {
        let mut wb = blank();
        wb.set_values(0, "A1", vec![vec![number(1.0), number(2.0), formula("A1+B1")]]).unwrap();
        wb.set_column_width(0, 2, 2, 25.0).unwrap();
        wb.set_column_width(0, 3, 16_384, 9.0).unwrap();
        assert_eq!(wb.insert_columns(0, 3, 2).unwrap(), 3);
        assert_eq!(formula_at(&wb, 0, "E1"), "A1+B1");
        assert_eq!(value(&wb, "E1"), CellValue::Number(3.0));
        let layout = wb.sheet_layout(0).unwrap();
        let width_of = |c: u32| layout.cols.iter().find(|(a, b, ..)| *a <= c && c <= *b).and_then(|x| x.2);
        assert_eq!([width_of(2), width_of(3), width_of(4), width_of(5), width_of(16_384)], [Some(25.0), Some(25.0), Some(25.0), Some(9.0), Some(9.0)]);

        // Values at the end of the sheet would fall off
        wb.set_value(0, "XFD1", CellInput::Number(1.0)).unwrap();
        assert!(wb.insert_columns(0, 1, 1).unwrap_err().contains("XFD1"));
        assert!(wb.insert_rows(0, 1_048_576, 2).is_err());
    }

    #[test]
    fn deletes_rows_leaving_ref_errors_where_nothing_is_left() {
        let mut wb = blank();
        wb.add_sheet("Resumen", None).unwrap();
        wb.set_values(
            0,
            "A1",
            vec![vec![number(1.0)], vec![number(2.0)], vec![number(3.0)], vec![number(4.0)], vec![formula("SUM(A1:A4)"), formula("A2*2"), formula("SUM(A2:A3)"), formula("A4*2")]],
        )
        .unwrap();
        wb.set_values(1, "A1", vec![vec![formula("Hoja1!A2"), formula("Hoja1!A4+1")]]).unwrap();
        let xml = wb.workbook_xml().unwrap().replace("</sheets>", r#"</sheets><definedNames><definedName name="Dos">Hoja1!$A$2</definedName></definedNames>"#);
        wb.package.put(wb.workbook_part.clone(), xml.into_bytes());
        with_tail(
            &mut wb,
            r#"<mergeCells count="2"><mergeCell ref="F2:F3"/><mergeCell ref="F7:F9"/></mergeCells><conditionalFormatting sqref="B2:B3"><cfRule type="expression" priority="1"><formula>B2&gt;0</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation sqref="C2"><formula1>$G$1</formula1></dataValidation></dataValidations><hyperlinks><hyperlink ref="D2" location="Hoja1!A2"/></hyperlinks>"#,
        );

        wb.checkpoint(None);
        assert_eq!(wb.delete_rows(0, 2, 2).unwrap(), 2);
        assert_eq!(value(&wb, "A2"), CellValue::Number(4.0));
        assert_eq!(formula_at(&wb, 0, "A3"), "SUM(A1:A2)");
        assert_eq!(value(&wb, "A3"), CellValue::Number(5.0));
        assert_eq!(formula_at(&wb, 0, "B3"), "#REF!*2");
        assert_eq!(value(&wb, "B3"), CellValue::Error("#REF!".into()));
        assert_eq!(formula_at(&wb, 0, "C3"), "SUM(#REF!)");
        assert_eq!(formula_at(&wb, 0, "D3"), "A2*2");
        assert_eq!(formula_at(&wb, 1, "A1"), "Hoja1!#REF!");
        assert_eq!(formula_at(&wb, 1, "B1"), "Hoja1!A2+1");
        assert!(wb.workbook_xml().unwrap().contains(">Hoja1!#REF!<"));
        let xml = sheet_xml(&wb);
        let empty_attr = format!("ref={0}{0}", '"');
        assert!(!xml.contains(&empty_attr), "{}", xml);
        assert!(xml.contains(r#"<mergeCells count="1"><mergeCell ref="F5:F7"/></mergeCells>"#), "{}", xml);
        assert!(!xml.contains("conditionalFormatting") && !xml.contains("dataValidations") && !xml.contains("hyperlinks"), "{}", xml);
        let again = Workbook::from_bytes(&wb.to_bytes().unwrap()).unwrap();
        assert_eq!(again.cell(0, "A2").unwrap().value, CellValue::Number(4.0));

        wb.undo().unwrap();
        assert_eq!(value(&wb, "A2"), CellValue::Number(2.0));
        assert_eq!(formula_at(&wb, 1, "A1"), "Hoja1!A2");
    }

    #[test]
    fn deleting_keeps_ranges_comments_and_columns_right() {
        let mut wb = blank();
        wb.set_values(0, "A1", vec![vec![number(1.0), number(2.0), number(3.0), formula("SUM(A1:C1)+SUM(B:B)")]]).unwrap();
        wb.set_column_width(0, 2, 2, 25.0).unwrap();
        wb.package.put("xl/comments1.xml".into(), br#"<comments xmlns="m"><commentList><comment ref="B1" authorId="0"/><comment ref="C1" authorId="0"/></commentList></comments>"#.to_vec());
        wb.package.put(
            "xl/drawings/vmlDrawing1.vml".into(),
            br#"<xml><v:shapetype id="t"/><v:shape id="a"><x:ClientData><x:Row>0</x:Row><x:Column>1</x:Column></x:ClientData></v:shape><v:shape id="b"><x:ClientData><x:Row>0</x:Row><x:Column>2</x:Column></x:ClientData></v:shape></xml>"#.to_vec(),
        );
        wb.package.put(
            "xl/worksheets/_rels/sheet1.xml.rels".into(),
            br#"<Relationships xmlns="r"><Relationship Id="rId2" Type="x/comments" Target="../comments1.xml"/><Relationship Id="rId3" Type="x/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>"#.to_vec(),
        );
        wb.delete_columns(0, 2, 1).unwrap();
        assert_eq!(value(&wb, "B1"), CellValue::Number(3.0));
        assert_eq!(formula_at(&wb, 0, "C1"), "SUM(A1:B1)+SUM(#REF!)");
        assert!(wb.sheet_layout(0).unwrap().cols.is_empty(), "{:?}", wb.sheet_layout(0).unwrap().cols);
        assert_eq!(wb.package.get_string("xl/comments1.xml").unwrap(), r#"<comments xmlns="m"><commentList><comment ref="B1" authorId="0"/></commentList></comments>"#);
        let vml = wb.package.get_string("xl/drawings/vmlDrawing1.vml").unwrap();
        assert!(!vml.contains(r#"id="a""#) && vml.contains(r#"<v:shape id="b"><x:ClientData><x:Row>0</x:Row><x:Column>1</x:Column>"#), "{}", vml);

        // Deleting at the very end works too
        let mut wb = blank();
        wb.set_value(0, "A1048576", CellInput::Number(1.0)).unwrap();
        wb.delete_rows(0, 1_048_576, 1).unwrap();
        assert_eq!(value(&wb, "A1048576"), CellValue::Empty);
    }

    #[test]
    fn tables_grow_and_shrink_with_their_rows() {
        let mut wb = blank();
        with_table(&mut wb, TABLE);
        wb.set_value(0, "A1", CellInput::Number(1.0)).unwrap();
        let table = |wb: &Workbook| wb.package.get_string("xl/tables/table1.xml").unwrap();
        // Table B2:D5: header in 2, data in 3–4, totals in 5
        assert!(wb.delete_rows(0, 2, 1).unwrap_err().contains("encabezado"));
        assert!(wb.delete_rows(0, 5, 1).unwrap_err().contains("totales"));
        assert!(wb.delete_rows(0, 3, 2).unwrap_err().contains("todas"));
        assert!(wb.delete_rows(0, 2, 4).unwrap_err().contains("toda la tabla"));
        assert!(wb.insert_columns(0, 3, 1).unwrap_err().contains("columnas"));
        assert!(wb.delete_columns(0, 3, 1).unwrap_err().contains("columnas"));

        wb.insert_rows(0, 4, 2).unwrap();
        assert!(table(&wb).contains(r#"ref="B2:D7""#) && table(&wb).contains(r#"<autoFilter ref="B2:D6">"#), "{}", table(&wb));
        wb.delete_rows(0, 4, 1).unwrap();
        assert!(table(&wb).contains(r#"ref="B2:D6""#) && table(&wb).contains(r#"<autoFilter ref="B2:D5">"#), "{}", table(&wb));
        // Before it or beside it, it moves whole
        wb.insert_rows(0, 1, 1).unwrap();
        assert!(table(&wb).contains(r#"ref="B3:D7""#), "{}", table(&wb));
        wb.insert_columns(0, 1, 1).unwrap();
        assert!(table(&wb).contains(r#"ref="C3:E7""#), "{}", table(&wb));
    }

    #[test]
    fn frozen_lines_count_what_is_inserted_or_deleted_in_them() {
        let mut wb = blank();
        wb.set_value(0, "A1", CellInput::Number(1.0)).unwrap();
        let xml = sheet_xml(&wb);
        let start = xml.find("<sheetData").unwrap();
        let view = r#"<sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>"#;
        let head = match (xml.find("<sheetViews"), xml.find("</sheetViews>")) {
            (Some(a), Some(b)) => format!("{}{}{}", &xml[..a], view, &xml[b + "</sheetViews>".len()..start]),
            _ => format!("{}{}", &xml[..start], view),
        };
        wb.package.put("xl/worksheets/sheet1.xml".into(), format!("{}{}", head, &xml[start..]).into_bytes());
        wb.insert_rows(0, 2, 2).unwrap();
        assert_eq!(wb.sheet_layout(0).unwrap().frozen, Some((5, 0)));
        assert!(sheet_xml(&wb).contains(r#"topLeftCell="A6""#), "{}", sheet_xml(&wb));
        wb.insert_rows(0, 20, 1).unwrap();
        assert_eq!(wb.sheet_layout(0).unwrap().frozen, Some((5, 0)));
        wb.delete_rows(0, 4, 4).unwrap();
        assert_eq!(wb.sheet_layout(0).unwrap().frozen, Some((3, 0)));
        wb.delete_rows(0, 1, 3).unwrap();
        assert_eq!(wb.sheet_layout(0).unwrap().frozen, None);
        assert!(!sheet_xml(&wb).contains("<pane"));
    }

    #[test]
    fn a_move_to_where_it_is_changes_nothing() {
        let mut wb = blank();
        wb.set_value(0, "A1", CellInput::Number(1.0)).unwrap();
        let before = sheet_xml(&wb);
        wb.checkpoint(None);
        assert_eq!(wb.move_rows(0, 1, 2, 2).unwrap(), 1);
        assert_eq!(wb.move_columns(0, 1, 1, 2).unwrap(), 1);
        assert_eq!(sheet_xml(&wb), before);
        assert!(!wb.can_undo());
    }
}
