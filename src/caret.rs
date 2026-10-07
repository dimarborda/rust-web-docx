//! Caret geometry on a computed layout: point → text position, text position → caret box,
//! vertical caret movement and selection rectangles.
//!
//! Positions are `(paragraph, offset)` with the offset counted in Unicode scalar values of
//! the paragraph text (what `ParagraphInfo::text` holds).

use crate::layout_engine::{DocumentLayout, FontSpec, RenderCommand, TextMeasurer, TextRun};
use serde::{Deserialize, Serialize};

#[derive(Serialize, Deserialize, Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextPosition {
    pub paragraph: usize,
    pub offset: usize,
}

#[derive(Serialize, Debug, Clone, PartialEq)]
pub struct CaretBox {
    pub page: usize,
    pub x: f64,
    /// Top of the line box
    pub y: f64,
    pub height: f64,
    /// First and last caret offsets on the caret's line (for Home / End)
    pub line_start: usize,
    pub line_end: usize,
}

#[derive(Serialize, Debug, Clone, PartialEq)]
pub struct Rect {
    pub page: usize,
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
}

/// A paragraph body line as laid out on a page
struct Line<'a> {
    page: usize,
    paragraph: usize,
    x: f64,
    /// Horizontal bounds of the line box (the paragraph's column)
    left: f64,
    right: f64,
    top: f64,
    height: f64,
    start: usize,
    end: usize,
    space_extra: f64,
    font_size: f64,
    family: &'a str,
    runs: &'a [TextRun],
    /// Ends its paragraph or a manual break (false for automatically wrapped lines)
    ends_with_break: bool,
}

fn char_len(s: &str) -> usize {
    s.chars().count()
}

impl Line<'_> {
    fn last_run_end(&self) -> usize {
        self.runs.last().map_or(self.start, |r| r.start + char_len(&r.text))
    }

    /// Last caret offset drawn on this line: a wrapped line hands `end` to the next line
    fn caret_end(&self) -> usize {
        if !self.ends_with_break && self.end > self.last_run_end() {
            self.end - 1
        } else {
            self.end
        }
    }

    fn prefix_width(&self, run: &TextRun, chars: usize, m: &mut dyn TextMeasurer) -> f64 {
        if chars == 0 {
            return 0.0;
        }
        if run.text == "\t" {
            return run.width;
        }
        let prefix: String = run.text.chars().take(chars).collect();
        let font = FontSpec {
            family: run.font_family.as_deref().unwrap_or(self.family),
            size: run.font_size.unwrap_or(self.font_size),
            bold: run.bold,
            italic: run.italic,
        };
        m.measure(&prefix, &font) + self.space_extra * prefix.matches(' ').count() as f64
    }

    fn x_at(&self, offset: usize, m: &mut dyn TextMeasurer) -> f64 {
        let Some(first) = self.runs.first() else { return self.x };
        if offset <= first.start {
            return first.x;
        }
        for run in self.runs {
            let len = char_len(&run.text);
            if offset <= run.start + len {
                return run.x + self.prefix_width(run, offset - run.start, m);
            }
        }
        let last = &self.runs[self.runs.len() - 1];
        last.x + self.prefix_width(last, char_len(&last.text), m)
    }

    /// Nearest caret offset to `x` on this line
    fn offset_at(&self, x: f64, m: &mut dyn TextMeasurer) -> usize {
        let Some(first) = self.runs.first() else { return self.start };
        if x <= first.x {
            return first.start;
        }
        for run in self.runs {
            let len = char_len(&run.text);
            if x > run.x + self.prefix_width(run, len, m) {
                continue;
            }
            // Last character boundary at or before x, then whichever side is closer
            let (mut lo, mut hi) = (0usize, len);
            while lo < hi {
                let mid = (lo + hi).div_ceil(2);
                if run.x + self.prefix_width(run, mid, m) <= x {
                    lo = mid;
                } else {
                    hi = mid - 1;
                }
            }
            if lo < len {
                let left = run.x + self.prefix_width(run, lo, m);
                let right = run.x + self.prefix_width(run, lo + 1, m);
                if x - left > right - x {
                    lo += 1;
                }
            }
            return run.start + lo;
        }
        self.caret_end()
    }

    fn distance_x(&self, x: f64) -> f64 {
        if x < self.left {
            self.left - x
        } else {
            (x - self.right).max(0.0)
        }
    }

    fn distance_y(&self, y: f64) -> f64 {
        if y < self.top {
            self.top - y
        } else {
            (y - (self.top + self.height)).max(0.0)
        }
    }
}

/// Body lines in reading order (pages in order, lines as they were placed)
fn lines(layout: &DocumentLayout) -> Vec<Line<'_>> {
    let mut out = Vec::new();
    for page in &layout.pages {
        for item in &page.items {
            if let RenderCommand::Text {
                paragraph_index,
                x,
                height,
                font_size,
                font_family,
                runs,
                is_last_line,
                line: Some(range),
                ..
            } = item
            {
                out.push(Line {
                    page: page.page_number,
                    paragraph: *paragraph_index,
                    x: *x,
                    left: range.left,
                    right: range.right.max(range.left),
                    top: range.top,
                    height: *height,
                    start: range.start,
                    end: range.end,
                    space_extra: range.space_extra,
                    font_size: *font_size,
                    family: font_family,
                    runs,
                    ends_with_break: *is_last_line,
                });
            }
        }
    }
    out
}

/// The line that displays the caret at `pos`. A wrapped line's `end` belongs to the next
/// line; `end` of a line closed by a break (or of the last line) belongs to that line.
fn line_index(lines: &[Line], pos: TextPosition) -> Option<usize> {
    let mut at_end = None;
    let mut last = None;
    for (i, line) in lines.iter().enumerate() {
        if line.paragraph != pos.paragraph {
            continue;
        }
        if line.start <= pos.offset && pos.offset < line.end {
            return Some(i);
        }
        if pos.offset == line.end && at_end.is_none() {
            at_end = Some(i);
        }
        last = Some(i);
    }
    at_end.or(last)
}

/// Text position under a point of a page (`page` is 1-based, coordinates in page px).
/// Lines in the clicked column win over closer lines in other columns (table cells side by
/// side share the same height band); clicks in a margin go to the line at that height.
pub fn hit_test(layout: &DocumentLayout, page: usize, x: f64, y: f64, m: &mut dyn TextMeasurer) -> Option<TextPosition> {
    let all = lines(layout);
    let score = |l: &Line| (l.distance_x(x) > 0.5, l.distance_y(y), l.distance_x(x));
    let line = all.iter().filter(|l| l.page == page).min_by(|a, b| {
        score(a).partial_cmp(&score(b)).unwrap_or(std::cmp::Ordering::Equal)
    })?;
    Some(TextPosition { paragraph: line.paragraph, offset: line.offset_at(x, m) })
}

pub fn caret_box(layout: &DocumentLayout, pos: TextPosition, m: &mut dyn TextMeasurer) -> Option<CaretBox> {
    let all = lines(layout);
    let line = &all[line_index(&all, pos)?];
    let offset = pos.offset.clamp(line.start, line.caret_end().max(line.start));
    Some(CaretBox {
        page: line.page,
        x: line.x_at(offset, m),
        y: line.top,
        height: line.height,
        line_start: line.start,
        line_end: line.caret_end(),
    })
}

/// Caret one line up (`direction < 0`) or down, keeping the horizontal position `goal_x`.
/// The next line is found geometrically: the nearest line above/below whose column contains
/// `goal_x` (so inside a table ↓ goes to the cell below, not the neighbour), else simply
/// the nearest line in that direction.
pub fn move_vertical(
    layout: &DocumentLayout,
    pos: TextPosition,
    direction: i32,
    goal_x: f64,
    m: &mut dyn TextMeasurer,
) -> Option<TextPosition> {
    let all = lines(layout);
    let current = &all[line_index(&all, pos)?];
    let here = (current.page, current.top);
    let beyond = |l: &&Line| {
        let there = (l.page, l.top);
        if direction < 0 {
            there.0 < here.0 || (there.0 == here.0 && there.1 < here.1 - 0.5)
        } else {
            there.0 > here.0 || (there.0 == here.0 && there.1 > here.1 + 0.5)
        }
    };
    let nearest = |candidates: Vec<&Line<'_>>| -> Option<usize> {
        let mut best: Option<(usize, (usize, f64))> = None;
        for (i, l) in candidates.iter().enumerate() {
            let key = (l.page, l.top);
            let better = match best {
                None => true,
                Some((_, b)) if direction < 0 => key.0 > b.0 || (key.0 == b.0 && key.1 > b.1),
                Some((_, b)) => key.0 < b.0 || (key.0 == b.0 && key.1 < b.1),
            };
            if better {
                best = Some((i, key));
            }
        }
        best.map(|(i, _)| i)
    };
    let candidates: Vec<&Line> = all.iter().filter(beyond).collect();
    let in_column: Vec<&Line> = candidates.iter().copied().filter(|l| l.distance_x(goal_x) == 0.0).collect();
    let target = match nearest(in_column.clone()) {
        Some(i) => in_column[i],
        None => candidates[nearest(candidates.clone())?],
    };
    Some(TextPosition { paragraph: target.paragraph, offset: target.offset_at(goal_x, m) })
}

/// Highlight rectangles for the characters `start..end` of a paragraph
pub fn selection_rects(
    layout: &DocumentLayout,
    paragraph: usize,
    start: usize,
    end: usize,
    m: &mut dyn TextMeasurer,
) -> Vec<Rect> {
    selection_rects_range(layout, TextPosition { paragraph, offset: start }, TextPosition { paragraph, offset: end }, m)
}

/// Highlight rectangles for a selection that may span several paragraphs
pub fn selection_rects_range(
    layout: &DocumentLayout,
    from: TextPosition,
    to: TextPosition,
    m: &mut dyn TextMeasurer,
) -> Vec<Rect> {
    let (from, to) = if (to.paragraph, to.offset) < (from.paragraph, from.offset) { (to, from) } else { (from, to) };
    let mut rects = Vec::new();
    if from == to {
        return rects;
    }
    for line in lines(layout).iter().filter(|l| l.paragraph >= from.paragraph && l.paragraph <= to.paragraph) {
        // Whole paragraphs in the middle of the selection, including their paragraph mark
        let start = if line.paragraph == from.paragraph { from.offset } else { 0 };
        let end = if line.paragraph == to.paragraph { to.offset } else { usize::MAX };
        if start >= end {
            continue;
        }
        let (ls, le) = (line.start, line.caret_end());
        if end < ls || start > le || (end == ls && ls < le) {
            continue;
        }
        let x0 = line.x_at(start.max(ls), m);
        let mut x1 = line.x_at(end.min(le), m);
        // The selection continues past this line: show the line break as selected
        if end > le {
            x1 += line.font_size * 0.3;
        }
        if x1 > x0 {
            rects.push(Rect { page: line.page, x: x0, y: line.top, width: x1 - x0, height: line.height });
        }
    }
    rects
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx_parser::{DocumentElement, HeaderFooterInfo, PageSetup, ParagraphInfo, RunInfo};
    use crate::layout_engine::{EstimateMeasurer, LayoutEngine};

    /// Every character is 10px wide: easy to reason about positions
    struct Mono;
    impl TextMeasurer for Mono {
        fn measure(&mut self, text: &str, _font: &FontSpec) -> f64 {
            text.chars().count() as f64 * 10.0
        }
    }

    fn para(index: usize, text: &str, align: &str) -> ParagraphInfo {
        ParagraphInfo {
            index,
            text: text.to_string(),
            style: "Normal".to_string(),
            align: align.to_string(),
            font_size: Some(12.0),
            spacing_resolved: true,
            widow_control: true,
            runs: vec![RunInfo { text: text.to_string(), font_size: Some(12.0), ..Default::default() }],
            ..Default::default()
        }
    }

    fn layout(paragraphs: Vec<ParagraphInfo>) -> DocumentLayout {
        let elements: Vec<DocumentElement> = paragraphs.into_iter().map(DocumentElement::Paragraph).collect();
        LayoutEngine::new().compute_layout_with(
            &elements,
            "FFFFFF",
            &PageSetup::default(),
            &HeaderFooterInfo::default(),
            None,
            None,
            0.0,
            &mut Mono,
        )
    }

    fn pos(paragraph: usize, offset: usize) -> TextPosition {
        TextPosition { paragraph, offset }
    }

    /// Default page: left margin 65px, 670px of line width → 67 characters per line
    const LEFT: f64 = 65.0;

    #[test]
    fn test_caret_x_and_hit_test_round_trip() {
        let l = layout(vec![para(0, "hola mundo", "left")]);
        let caret = caret_box(&l, pos(0, 4), &mut Mono).unwrap();
        assert_eq!((caret.page, caret.x), (1, LEFT + 40.0));

        // Clicking just right of the 4th boundary snaps back to it; past the middle goes forward
        let y = caret.y + 1.0;
        assert_eq!(hit_test(&l, 1, LEFT + 43.0, y, &mut Mono), Some(pos(0, 4)));
        assert_eq!(hit_test(&l, 1, LEFT + 46.0, y, &mut Mono), Some(pos(0, 5)));
        assert_eq!(hit_test(&l, 1, 5.0, y, &mut Mono), Some(pos(0, 0)), "left margin → line start");
        assert_eq!(hit_test(&l, 1, 700.0, y, &mut Mono), Some(pos(0, 10)), "past the end → paragraph end");
    }

    #[test]
    fn test_wrapped_lines_hand_over_at_word_start() {
        // 60 a's, a space, then 10 b's: "bbbbbbbbbb" wraps to line 2 starting at offset 61
        let text = format!("{} {}", "a".repeat(60), "b".repeat(10));
        let l = layout(vec![para(0, &text, "left")]);
        let first = caret_box(&l, pos(0, 60), &mut Mono).unwrap();
        let space = caret_box(&l, pos(0, 61), &mut Mono).unwrap();
        assert!(space.y > first.y, "offset 61 starts the second line");
        assert_eq!(space.x, LEFT);
        assert_eq!(first.line_end, 60, "the swallowed space is the last caret stop of line 1");

        // Clicking past the end of line 1 stays on line 1
        let hit = hit_test(&l, 1, 900.0, first.y + 1.0, &mut Mono).unwrap();
        assert_eq!(hit, pos(0, 60));

        // Down from offset 5 keeps the column on the next line
        let down = move_vertical(&l, pos(0, 5), 1, LEFT + 50.0, &mut Mono).unwrap();
        assert_eq!(down, pos(0, 66));
        assert_eq!(move_vertical(&l, pos(0, 66), -1, LEFT + 50.0, &mut Mono), Some(pos(0, 5)));
    }

    #[test]
    fn test_vertical_moves_cross_paragraphs_and_stop_at_edges() {
        let l = layout(vec![para(0, "uno", "left"), para(1, "dos tres", "left")]);
        assert_eq!(move_vertical(&l, pos(0, 2), 1, LEFT + 20.0, &mut Mono), Some(pos(1, 2)));
        assert_eq!(move_vertical(&l, pos(0, 2), -1, LEFT + 20.0, &mut Mono), None);
        assert_eq!(move_vertical(&l, pos(1, 7), 1, LEFT + 70.0, &mut Mono), None);
    }

    #[test]
    fn test_manual_breaks_and_empty_lines_hold_the_caret() {
        let l = layout(vec![para(0, "ab\n\ncd", "left")]);
        let a = caret_box(&l, pos(0, 2), &mut Mono).unwrap();
        let empty = caret_box(&l, pos(0, 3), &mut Mono).unwrap();
        let c = caret_box(&l, pos(0, 4), &mut Mono).unwrap();
        assert_eq!(a.x, LEFT + 20.0, "offset of '\\n' is the end of its line");
        assert!(empty.y > a.y && c.y > empty.y);
        assert_eq!((empty.x, c.x), (LEFT, LEFT));

        // An empty paragraph still has a caret position
        let l = layout(vec![para(0, "", "left")]);
        assert_eq!(caret_box(&l, pos(0, 0), &mut Mono).unwrap().x, LEFT);
        assert_eq!(hit_test(&l, 1, 300.0, 80.0, &mut Mono), Some(pos(0, 0)));
    }

    #[test]
    fn test_justified_lines_spread_caret_positions() {
        let text = format!("{} {} {}", "a".repeat(30), "b".repeat(30), "c".repeat(10));
        let l = layout(vec![para(0, &text, "both")]);
        // Line 1 = "aaa… bbb…" (61 chars = 610px) stretched to 670px: its one space gets +60px
        let b_start = caret_box(&l, pos(0, 31), &mut Mono).unwrap();
        assert_eq!(b_start.x, LEFT + 310.0 + 60.0);
        assert_eq!(hit_test(&l, 1, LEFT + 375.0, b_start.y + 1.0, &mut Mono), Some(pos(0, 31)));
        // The last line is not justified
        let c = caret_box(&l, pos(0, 63), &mut Mono).unwrap();
        assert_eq!(c.x, LEFT + 10.0);
    }

    #[test]
    fn test_selection_rects_cover_the_range_per_line() {
        let text = format!("{} {}", "a".repeat(60), "b".repeat(10));
        let l = layout(vec![para(0, &text, "left")]);
        let rects = selection_rects(&l, 0, 58, 63, &mut Mono);
        assert_eq!(rects.len(), 2);
        assert_eq!(rects[0].x, LEFT + 580.0);
        assert!((rects[0].width - (20.0 + 16.0 * 0.3)).abs() < 0.01, "2 chars + line-break marker");
        assert_eq!((rects[1].x, rects[1].width), (LEFT, 20.0));
        assert!(selection_rects(&l, 0, 5, 5, &mut Mono).is_empty());
    }

    #[test]
    fn test_real_document_positions_round_trip() {
        let bytes = crate::sample_generator::generate_sample_docx().unwrap();
        let m = crate::docx_parser::DocxModifier::from_bytes(&bytes).unwrap();
        let elements = m.extract_elements().unwrap();
        let s = m.get_statistics().unwrap();
        let mut est = EstimateMeasurer;
        let l = LayoutEngine::new().compute_layout_with(
            &elements, &s.background_color, &s.page_setup, &s.header_footer, None, None, 0.0, &mut est,
        );
        for p in m.extract_paragraphs().unwrap() {
            for offset in 0..=p.text.chars().count() {
                let at = pos(p.index, offset);
                let caret = caret_box(&l, at, &mut est).expect("every position has a caret");
                // Clicking where the caret is drawn finds the same position again
                let hit = hit_test(&l, caret.page, caret.x, caret.y + caret.height / 2.0, &mut est).unwrap();
                assert_eq!(hit, at, "paragraph {:?}", p.text);
            }
        }
    }

    #[test]
    fn test_blank_document_has_a_caret_at_the_start() {
        use crate::blank_generator::{generate_blank_docx, PageSize};
        let m = crate::docx_parser::DocxModifier::from_bytes(&generate_blank_docx(PageSize::Letter).unwrap()).unwrap();
        let elements = m.extract_elements().unwrap();
        let s = m.get_statistics().unwrap();
        let mut est = EstimateMeasurer;
        let l = LayoutEngine::new().compute_layout_with(
            &elements, &s.background_color, &s.page_setup, &s.header_footer, None, None, 0.0, &mut est,
        );
        let caret = caret_box(&l, pos(0, 0), &mut est).expect("an empty paragraph still has a caret");
        assert_eq!(caret.page, 1);
        assert!(caret.height > 0.0);
        let hit = hit_test(&l, 1, caret.x + 200.0, caret.y + caret.height / 2.0, &mut est).unwrap();
        assert_eq!(hit, pos(0, 0), "clicking anywhere on the empty line lands on it");
    }

    /// Body paragraph 0, a 2×2 table (paragraphs 1–4, row by row), body paragraph 5
    fn table_layout() -> DocumentLayout {
        use crate::docx_parser::{TableCellData, TableInfo, TableRowData};
        let cell = |i: usize, text: &str| TableCellData { paragraphs: vec![para(i, text, "left")], ..Default::default() };
        let table = TableInfo {
            index: 0,
            rows: vec![],
            rich_rows: vec![
                TableRowData { cells: vec![cell(1, "a1 primera línea larga"), cell(2, "b1")], is_header: false, ..Default::default() },
                TableRowData { cells: vec![cell(3, "a2"), cell(4, "b2")], is_header: false, ..Default::default() },
            ],
            // Two 3000-twip (200px) columns
            grid_cols: vec![3000.0, 3000.0],
            header_row: false,
            borders: Default::default(),
            ..Default::default()
        };
        let elements = vec![
            DocumentElement::Paragraph(para(0, "antes", "left")),
            DocumentElement::Table(table),
            DocumentElement::Paragraph(para(5, "después", "left")),
        ];
        LayoutEngine::new().compute_layout_with(
            &elements,
            "FFFFFF",
            &PageSetup::default(),
            &HeaderFooterInfo::default(),
            None,
            None,
            0.0,
            &mut Mono,
        )
    }

    #[test]
    fn test_table_cells_get_their_own_caret_positions() {
        let l = table_layout();
        let a1 = caret_box(&l, pos(1, 0), &mut Mono).unwrap();
        let b1 = caret_box(&l, pos(2, 0), &mut Mono).unwrap();
        assert_eq!(a1.y, b1.y, "cells of a row share the line band");
        assert_eq!(a1.x, LEFT + 7.2, "text starts after the cell margin");
        assert_eq!(b1.x, LEFT + 200.0 + 7.2, "second column starts 200px later");

        // A click picks the cell under the pointer, not the first line at that height
        let y = a1.y + 1.0;
        assert_eq!(hit_test(&l, 1, LEFT + 230.0, y, &mut Mono), Some(pos(2, 2)));
        assert_eq!(hit_test(&l, 1, LEFT + 27.0, y, &mut Mono), Some(pos(1, 2)));
    }

    #[test]
    fn test_vertical_moves_stay_in_the_column() {
        let l = table_layout();
        let x_b = LEFT + 200.0 + 7.2 + 10.0;
        // ↓ from b1 goes to b2 (below), not to a2 (next in reading order)
        assert_eq!(move_vertical(&l, pos(2, 1), 1, x_b, &mut Mono), Some(pos(4, 1)));
        // ↓ from the last row leaves the table into the following paragraph
        assert_eq!(move_vertical(&l, pos(4, 1), 1, x_b, &mut Mono).map(|p| p.paragraph), Some(5));
        // ↑ from the paragraph above the table enters the column under the caret
        assert_eq!(move_vertical(&l, pos(0, 1), 1, x_b, &mut Mono).map(|p| p.paragraph), Some(2));
    }

    #[test]
    fn test_words_wider_than_the_line_break_between_characters() {
        // 100 characters without spaces on a 67-character line: 67 + 33, like Word
        let l = layout(vec![para(0, &"x".repeat(100), "left")]);
        let first = caret_box(&l, pos(0, 66), &mut Mono).unwrap();
        let second = caret_box(&l, pos(0, 67), &mut Mono).unwrap();
        assert!(second.y > first.y, "the word continues on the next line");
        assert_eq!(second.x, LEFT);
        assert_eq!(caret_box(&l, pos(0, 100), &mut Mono).unwrap().x, LEFT + 330.0);
        assert_eq!(hit_test(&l, 1, LEFT + 50.0, second.y + 1.0, &mut Mono), Some(pos(0, 72)));
    }
}

