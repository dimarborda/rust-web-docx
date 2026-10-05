use crate::docx_parser::{DocumentElement, HeaderFooterInfo, PageSetup, ParagraphInfo, RunInfo, TableInfo, PAGE_BREAK};
use std::collections::HashMap;
use serde::{Deserialize, Serialize};

pub const PAGE_WIDTH: f64 = 800.0;
pub const PAGE_HEIGHT: f64 = 1130.0; // A4 standard ratio
pub const MARGIN_LEFT: f64 = 65.0;
pub const MARGIN_RIGHT: f64 = 65.0;
pub const MARGIN_TOP: f64 = 70.0;
pub const MARGIN_BOTTOM: f64 = 70.0;
pub const PRINTABLE_WIDTH: f64 = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT; // 670.0
pub const PRINTABLE_HEIGHT: f64 = PAGE_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM; // 990.0

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct DocumentLayout {
    pub total_pages: usize,
    pub page_width: f64,
    pub page_height: f64,
    pub pages: Vec<PageLayout>,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct PageLayout {
    pub page_number: usize,
    pub width: f64,
    pub height: f64,
    pub margin_left: f64,
    pub margin_right: f64,
    pub printable_width: f64,
    pub bg_color: String,
    pub items: Vec<RenderCommand>,
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq, Default)]
pub struct TextRun {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub color: String,
    pub font_size: Option<f64>,
    pub font_family: Option<String>,
    pub width: f64,
}

#[derive(Serialize, Deserialize, Debug, Clone)]
#[serde(tag = "type")]
pub enum RenderCommand {
    #[serde(rename = "text")]
    Text {
        text: String,
        x: f64,
        y: f64,
        width: f64,
        height: f64,
        font_size: f64,
        font_family: String,
        font_weight: String,
        font_style: String,
        color: String,
        align: String,
        paragraph_index: usize,
        line_index: usize,
        #[serde(default)]
        is_last_line: bool,
        #[serde(default)]
        max_width: f64,
        #[serde(default)]
        runs: Vec<TextRun>,
    },
    #[serde(rename = "table_cell")]
    TableCell {
        table_index: usize,
        row: usize,
        col: usize,
        x: f64,
        y: f64,
        width: f64,
        height: f64,
        text: String,
        #[serde(default)]
        lines: Vec<String>,
        is_header: bool,
        bg_color: Option<String>,
        color: String,
        font_size: f64,
        font_family: String,
        font_weight: String,
        align: String,
        border_color: String,
    },
    #[serde(rename = "line")]
    Line {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        color: String,
        line_width: f64,
    },
    #[serde(rename = "image")]
    Image {
        data_url: String,
        x: f64,
        y: f64,
        width: f64,
        height: f64,
        is_background: bool,
        opacity: f64,
    },
    #[serde(rename = "watermark")]
    Watermark {
        text: String,
        x: f64,
        y: f64,
        opacity: f64,
        font_size: f64,
        color: String,
        rotation_deg: f64,
    },
}

pub struct LayoutEngine {
    pub page_width: f64,
    pub page_height: f64,
}

impl LayoutEngine {
    pub fn new() -> Self {
        Self {
            page_width: PAGE_WIDTH,
            page_height: PAGE_HEIGHT,
        }
    }

    /// Computes the layout with estimated glyph widths (native builds and tests)
    #[allow(clippy::too_many_arguments)]
    pub fn compute_layout(
        &self,
        elements: &[DocumentElement],
        bg_color: &str,
        page_setup: &PageSetup,
        header_footer: &HeaderFooterInfo,
        bg_image_data_url: Option<&str>,
        watermark: Option<&str>,
        watermark_opacity: f64,
    ) -> DocumentLayout {
        let mut measurer = CachedMeasurer::new(EstimateMeasurer);
        self.compute_layout_with(
            elements,
            bg_color,
            page_setup,
            header_footer,
            bg_image_data_url,
            watermark,
            watermark_opacity,
            &mut measurer,
        )
    }

    /// Computes multi-page layout using `measurer` for text widths
    #[allow(clippy::too_many_arguments)]
    pub fn compute_layout_with(
        &self,
        elements: &[DocumentElement],
        bg_color: &str,
        page_setup: &PageSetup,
        header_footer: &HeaderFooterInfo,
        bg_image_data_url: Option<&str>,
        watermark: Option<&str>,
        watermark_opacity: f64,
        measurer: &mut dyn TextMeasurer,
    ) -> DocumentLayout {
        let page_w = page_setup.width.max(320.0);
        let page_h = page_setup.height.max(400.0);
        let margin_left = page_setup.margin_left.max(20.0);
        let margin_right = page_setup.margin_right.max(20.0);
        let margin_top = page_setup.margin_top.max(24.0);
        let margin_bottom = page_setup.margin_bottom.max(24.0);

        let geo = PageGeometry {
            page_w,
            page_h,
            margin_left,
            margin_right,
            printable_w: (page_w - margin_left - margin_right).max(100.0),
            // If a header logo is present, page 1 content starts below the logo area
            top_first: if header_footer.header_image_data_url.is_some() { margin_top.max(120.0) } else { margin_top },
            top: margin_top,
            bottom: page_h - margin_bottom,
            bg: css_color(bg_color, "#FFFFFF"),
        };

        // Pass 1: measure every block; pass 2: paginate with look-ahead for keep rules
        let blocks: Vec<Block> = elements
            .iter()
            .map(|el| match el {
                DocumentElement::Paragraph(p) => Block::Paragraph(prepare_paragraph(p, &geo, measurer)),
                DocumentElement::Table(t) => Block::Table(prepare_table(t, &geo, measurer)),
            })
            .collect();

        let mut pag = Paginator::new(&geo, watermark, watermark_opacity);
        let mut flow = FlowState::default();
        for (i, block) in blocks.iter().enumerate() {
            match block {
                Block::Paragraph(b) => place_paragraph(&mut pag, &blocks, i, b, &mut flow),
                Block::Table(t) => place_table(&mut pag, t, &mut flow),
            }
        }
        let mut pages = pag.finish();

        // Decorate pages with Background Images, Header Logo/Text, and Footers
        let total_pages = pages.len();
        for (page_idx, page) in pages.iter_mut().enumerate() {
            let page_num = page_idx + 1;

            // 1. Full page background image (e.g. diploma / certificate)
            if let Some(bg_url) = bg_image_data_url {
                page.items.insert(
                    0,
                    RenderCommand::Image {
                        data_url: bg_url.to_string(),
                        x: 0.0,
                        y: 0.0,
                        width: page_w,
                        height: page_h,
                        is_background: true,
                        opacity: 1.0,
                    },
                );
            }

            // 2. Header logo & text (top-left placement matching OpenXML)
            if let Some(ref header_img) = header_footer.header_image_data_url {
                page.items.push(RenderCommand::Image {
                    data_url: header_img.clone(),
                    x: 20.0,
                    y: 15.0,
                    width: 140.0,
                    height: 95.0,
                    is_background: false,
                    opacity: 1.0,
                });
            }

            if !header_footer.header_text.is_empty() {
                page.items.push(RenderCommand::Text {
                    text: header_footer.header_text.clone(),
                    x: page_w - margin_right,
                    y: page_setup.header_margin.max(12.0) + 12.0,
                    width: 250.0,
                    height: 14.0,
                    font_size: 9.0,
                    font_family: "Inter, sans-serif".to_string(),
                    font_weight: "500".to_string(),
                    font_style: "normal".to_string(),
                    color: "#64748B".to_string(),
                    align: "right".to_string(),
                    paragraph_index: 0,
                    line_index: 0,
                    is_last_line: true,
                    max_width: 250.0,
                    runs: Vec::new(),
                });
            }

            // 3. Footer (only if document has footer text or multiple pages)
            if !header_footer.footer_text.is_empty() {
                page.items.push(RenderCommand::Text {
                    text: header_footer.footer_text.clone(),
                    x: margin_left,
                    y: page_h - page_setup.footer_margin.max(14.0),
                    width: 250.0,
                    height: 14.0,
                    font_size: 9.0,
                    font_family: "Inter, sans-serif".to_string(),
                    font_weight: "400".to_string(),
                    font_style: "normal".to_string(),
                    color: "#64748B".to_string(),
                    align: "left".to_string(),
                    paragraph_index: 0,
                    line_index: 0,
                    is_last_line: true,
                    max_width: 250.0,
                    runs: Vec::new(),
                });
            }

            if total_pages > 1 {
                page.items.push(RenderCommand::Text {
                    text: format!("Página {} de {}", page_num, total_pages),
                    x: page_w - margin_right,
                    y: page_h - page_setup.footer_margin.max(14.0),
                    width: 120.0,
                    height: 14.0,
                    font_size: 9.0,
                    font_family: "Inter, sans-serif".to_string(),
                    font_weight: "500".to_string(),
                    font_style: "normal".to_string(),
                    color: "#64748B".to_string(),
                    align: "right".to_string(),
                    paragraph_index: 0,
                    line_index: 0,
                    is_last_line: true,
                    max_width: 120.0,
                    runs: Vec::new(),
                });
            }
        }

        DocumentLayout {
            total_pages,
            page_width: page_w,
            page_height: page_h,
            pages,
        }
    }
}


/// Measures the advance width (px) of `text` in a font. The web build measures with the same
/// canvas API that draws the page, so line breaks match what is rendered exactly.
pub trait TextMeasurer {
    fn measure(&mut self, text: &str, font: &FontSpec) -> f64;
}

#[derive(Debug, Clone, PartialEq)]
pub struct FontSpec<'a> {
    /// Raw Word family ("Arial") or a CSS stack; the renderer maps both the same way
    pub family: &'a str,
    /// px
    pub size: f64,
    pub bold: bool,
    pub italic: bool,
}

/// Approximate glyph widths, for native builds and tests
pub struct EstimateMeasurer;

impl TextMeasurer for EstimateMeasurer {
    fn measure(&mut self, text: &str, font: &FontSpec) -> f64 {
        estimate_text_width(text, font.size) * if font.bold { 1.05 } else { 1.0 }
    }
}

/// Memoizes another measurer: a document measures the same words over and over
pub struct CachedMeasurer<M: TextMeasurer> {
    inner: M,
    cache: HashMap<String, f64>,
}

impl<M: TextMeasurer> CachedMeasurer<M> {
    pub fn new(inner: M) -> Self {
        CachedMeasurer { inner, cache: HashMap::new() }
    }
}

impl<M: TextMeasurer> TextMeasurer for CachedMeasurer<M> {
    fn measure(&mut self, text: &str, font: &FontSpec) -> f64 {
        let key = format!("{}|{}|{}{}|{}", font.family, font.size, font.bold as u8, font.italic as u8, text);
        if let Some(w) = self.cache.get(&key) {
            return *w;
        }
        let w = self.inner.measure(text, font);
        self.cache.insert(key, w);
        w
    }
}

const PX_PER_PT: f64 = 96.0 / 72.0;

/// Ascent and descent (em) behind Word's "single" line height: OS/2 win metrics plus line gap
fn font_vertical_metrics(family: &str) -> (f64, f64) {
    let first = family
        .split(',')
        .next()
        .unwrap_or("")
        .trim()
        .trim_matches(|c| c == '"' || c == '\'')
        .to_lowercase();
    let has = |names: &[&str]| names.iter().any(|n| first.contains(n));
    if has(&["calibri", "carlito"]) {
        (0.952, 0.269)
    } else if has(&["arial", "arimo", "helvetica", "liberation sans"]) {
        (0.938, 0.212)
    } else if has(&["times", "tinos", "liberation serif"]) {
        (0.933, 0.216)
    } else if has(&["cambria", "caladea"]) {
        (0.950, 0.222)
    } else if has(&["georgia"]) {
        (0.917, 0.219)
    } else if has(&["aptos"]) {
        (0.939, 0.261)
    } else if has(&["courier", "cousine", "consolas"]) {
        (0.833, 0.300)
    } else if has(&["verdana"]) {
        (1.005, 0.210)
    } else if has(&["tahoma"]) {
        (1.000, 0.207)
    } else if has(&["segoe"]) {
        (1.079, 0.251)
    } else {
        (0.920, 0.250)
    }
}

struct PageGeometry {
    page_w: f64,
    page_h: f64,
    margin_left: f64,
    margin_right: f64,
    printable_w: f64,
    /// Content top of the first page (lower when a header logo is drawn)
    top_first: f64,
    top: f64,
    bottom: f64,
    bg: String,
}

/// A laid out line plus its vertical metrics
struct LineBox {
    line: LayoutLine,
    height: f64,
    /// Baseline offset from the top of the line box
    baseline: f64,
}

struct ParagraphBox<'a> {
    p: &'a ParagraphInfo,
    lines: Vec<LineBox>,
    font_size: f64,
    font_weight: &'static str,
    font_style: &'static str,
    font_family: String,
    color: String,
    space_before: f64,
    space_after: f64,
    label: Option<ListLabelGeometry>,
    top_border_h: f64,
}

impl ParagraphBox<'_> {
    fn lines_height(&self) -> f64 {
        self.lines.iter().map(|l| l.height).sum()
    }
}

struct RowBox {
    height: f64,
    cell_lines: Vec<Vec<String>>,
    is_header: bool,
}

struct TableBox<'a> {
    tbl: &'a TableInfo,
    col_widths: Vec<f64>,
    rows: Vec<RowBox>,
}

enum Block<'a> {
    Paragraph(ParagraphBox<'a>),
    Table(TableBox<'a>),
}

fn css_color(c: &str, fallback: &str) -> String {
    if c.is_empty() || c.eq_ignore_ascii_case("auto") {
        fallback.to_string()
    } else if c.starts_with('#') {
        c.to_string()
    } else {
        format!("#{}", c)
    }
}

fn prepare_paragraph<'a>(p: &'a ParagraphInfo, geo: &PageGeometry, m: &mut dyn TextMeasurer) -> ParagraphBox<'a> {
    let (font_size, _, font_weight, font_family, default_color) = get_paragraph_typography(p);
    let color = css_color(&p.color, &default_color);
    let lines = layout_paragraph_lines(p, font_size, &color, &font_family, geo.margin_left, geo.printable_w, m);
    let label = list_label_geometry(p, font_size, &font_family, m);

    let lines = lines
        .into_iter()
        .map(|line| {
            // The tallest run sets the line's natural height
            let (size, family) = line
                .runs
                .iter()
                .filter(|r| r.text != "\t")
                .map(|r| (r.font_size.unwrap_or(font_size), r.font_family.as_deref().unwrap_or(&font_family)))
                .fold((font_size, font_family.as_str()), |best, cur| if cur.0 > best.0 { cur } else { best });
            let (ascent, descent) = font_vertical_metrics(family);
            let natural = size * (ascent + descent);
            let height = if let Some(pt) = p.line_exact {
                pt * PX_PER_PT
            } else if let Some(pt) = p.line_at_least {
                natural.max(pt * PX_PER_PT)
            } else {
                natural * p.line_spacing.unwrap_or(if p.spacing_resolved { 1.0 } else { 1.15 })
            };
            // Extra leading sits above the text, as in Word
            LineBox { line, height, baseline: height - descent * size }
        })
        .collect();

    let (space_before, space_after) = if p.spacing_resolved {
        (p.space_before * PX_PER_PT, p.space_after * PX_PER_PT)
    } else {
        let after = if p.space_after > 0.0 { p.space_after } else if p.is_heading { 12.0 } else { 6.0 };
        (p.space_before, after)
    };
    let top_border_h = p.borders.top.as_ref().map_or(0.0, |b| b.space.max(3.0) + b.sz_px);

    ParagraphBox {
        p,
        lines,
        font_size,
        font_weight,
        font_style: if p.italic { "italic" } else { "normal" },
        font_family,
        color,
        space_before,
        space_after,
        label,
        top_border_h,
    }
}

fn prepare_table<'a>(tbl: &'a TableInfo, geo: &PageGeometry, m: &mut dyn TextMeasurer) -> TableBox<'a> {
    let printable_w = geo.printable_w;
    let num_cols = if !tbl.grid_cols.is_empty() {
        tbl.grid_cols.len()
    } else {
        tbl.rows.first().map(|r| r.len()).unwrap_or(1).max(1)
    };
    let col_widths: Vec<f64> = if !tbl.grid_cols.is_empty() {
        let total: f64 = tbl.grid_cols.iter().sum::<f64>().max(1.0);
        tbl.grid_cols.iter().map(|w| (w / total) * printable_w).collect()
    } else {
        vec![printable_w / num_cols as f64; num_cols]
    };

    let rows = tbl
        .rows
        .iter()
        .enumerate()
        .map(|(row_idx, row_strings)| {
            let rich_row = tbl.rich_rows.get(row_idx);
            let is_header = rich_row.map(|r| r.is_header).unwrap_or(row_idx == 0 && tbl.header_row);
            let cell_lines: Vec<Vec<String>> = row_strings
                .iter()
                .enumerate()
                .map(|(col_idx, cell_text)| {
                    let col_w = col_widths.get(col_idx).copied().unwrap_or(printable_w / num_cols as f64);
                    let cell = rich_row.and_then(|r| r.cells.get(col_idx));
                    let size = cell.map(|c| c.font_size * PX_PER_PT).unwrap_or(if is_header { 13.33 } else { 12.0 });
                    let family = cell.map(|c| c.font_family.as_str()).unwrap_or("Calibri, Inter, sans-serif");
                    let bold = cell.map(|c| c.bold).unwrap_or(is_header);
                    let font = FontSpec { family, size, bold, italic: false };
                    let mut lines: Vec<String> = cell_text
                        .split('\n')
                        .flat_map(|part| wrap_text(part, (col_w - 16.0).max(10.0), &font, m))
                        .collect();
                    if lines.is_empty() {
                        lines.push(String::new());
                    }
                    lines
                })
                .collect();
            let max_lines = cell_lines.iter().map(|l| l.len()).max().unwrap_or(1).max(1);
            // Must match the renderer's cell line height (font size × 1.35)
            let base_fs = rich_row.and_then(|r| r.cells.first()).map(|c| c.font_size * PX_PER_PT).unwrap_or(13.33);
            let height = ((max_lines as f64) * base_fs * 1.35 + 14.0).max(28.0);
            RowBox { height, cell_lines, is_header }
        })
        .collect();

    TableBox { tbl, col_widths, rows }
}

/// Places content top to bottom and starts new pages as needed
struct Paginator<'a> {
    geo: &'a PageGeometry,
    watermark: Option<&'a str>,
    watermark_opacity: f64,
    pages: Vec<PageLayout>,
    items: Vec<RenderCommand>,
    cursor_y: f64,
}

impl<'a> Paginator<'a> {
    fn new(geo: &'a PageGeometry, watermark: Option<&'a str>, watermark_opacity: f64) -> Self {
        Paginator { geo, watermark, watermark_opacity, pages: Vec::new(), items: Vec::new(), cursor_y: geo.top_first }
    }

    fn page_top(&self) -> f64 {
        if self.pages.is_empty() {
            self.geo.top_first
        } else {
            self.geo.top
        }
    }

    fn at_top(&self) -> bool {
        self.cursor_y <= self.page_top() + 0.01
    }

    fn remaining(&self) -> f64 {
        self.geo.bottom - self.cursor_y
    }

    fn capacity(&self) -> f64 {
        self.geo.bottom - self.geo.top
    }

    fn new_page(&mut self) {
        let mut items = std::mem::take(&mut self.items);
        if let Some(wm) = self.watermark {
            add_watermark_command(&mut items, wm, self.watermark_opacity, self.geo.page_w, self.geo.page_h);
        }
        self.pages.push(PageLayout {
            page_number: self.pages.len() + 1,
            width: self.geo.page_w,
            height: self.geo.page_h,
            margin_left: self.geo.margin_left,
            margin_right: self.geo.margin_right,
            printable_width: self.geo.printable_w,
            bg_color: self.geo.bg.clone(),
            items,
        });
        self.cursor_y = self.geo.top;
    }

    fn finish(mut self) -> Vec<PageLayout> {
        if !self.items.is_empty() || self.pages.is_empty() {
            self.new_page();
        }
        self.pages
    }
}

/// What the previous block left for the spacing between it and the next one
#[derive(Default)]
struct FlowState {
    pending_after: f64,
    prev_style: Option<String>,
    prev_contextual: bool,
}

/// Height that must stay on one page with a `keepNext` paragraph: the chain of following
/// keepNext paragraphs plus the first lines of the paragraph (or row of the table) after them
fn keep_chain_height(blocks: &[Block], start: usize) -> f64 {
    let mut h = 0.0;
    let mut j = start;
    while j < blocks.len() && j - start <= 20 {
        match &blocks[j] {
            Block::Paragraph(b) if j == start || b.p.keep_next => {
                if j > start {
                    h += b.space_before + b.top_border_h;
                }
                h += b.lines_height() + b.space_after;
                if !b.p.keep_next {
                    break;
                }
            }
            Block::Paragraph(b) => {
                let first_lines = if b.p.widow_control { 2 } else { 1 };
                h += b.space_before + b.top_border_h + b.lines.iter().take(first_lines).map(|l| l.height).sum::<f64>();
                break;
            }
            Block::Table(t) => {
                h += 8.0 + t.rows.first().map_or(0.0, |r| r.height);
                break;
            }
        }
        j += 1;
    }
    h
}

fn place_paragraph(pag: &mut Paginator, blocks: &[Block], i: usize, b: &ParagraphBox, flow: &mut FlowState) {
    let p = b.p;
    if p.page_break_before && !pag.at_top() {
        pag.new_page();
    }

    // contextualSpacing drops the spacing between paragraphs of the same style
    let same_style = flow.prev_style.as_deref() == Some(p.style.as_str());
    let after_prev = if same_style && flow.prev_contextual { 0.0 } else { flow.pending_after };
    let before = if same_style && p.contextual_spacing { 0.0 } else { b.space_before };
    let lead = |pag: &Paginator| if pag.at_top() { 0.0 } else { after_prev } + before + b.top_border_h;

    // keepLines / keepNext: move to a new page when the block fits on a page but not here
    let total = b.lines_height();
    let keep_height = if p.keep_next {
        Some(keep_chain_height(blocks, i) - b.space_after)
    } else if p.keep_lines {
        Some(total)
    } else {
        None
    };
    if let Some(h) = keep_height {
        if !pag.at_top() && h + lead(pag) > pag.remaining() && h + before + b.top_border_h <= pag.capacity() {
            pag.new_page();
        }
    }
    pag.cursor_y += lead(pag) - b.top_border_h;

    if let Some(ref top) = p.borders.top {
        let y = pag.cursor_y;
        pag.items.push(RenderCommand::Line {
            x1: pag.geo.margin_left,
            y1: y,
            x2: pag.geo.page_w - pag.geo.margin_right,
            y2: y,
            color: css_color(&top.color, "#CBD5E1"),
            line_width: top.sz_px.max(0.75),
        });
        pag.cursor_y += b.top_border_h;
    }

    // Lines, split into chunks at manual page breaks, with widow/orphan control
    let n = b.lines.len();
    let mut k0 = 0;
    let mut seg_top = pag.cursor_y;
    while k0 < n {
        let chunk_end = (k0..n).find(|&k| b.lines[k].line.page_break_after).map_or(n, |k| k + 1);
        let mut fit = 0;
        let mut h = 0.0;
        while k0 + fit < chunk_end && h + b.lines[k0 + fit].height <= pag.remaining() + 0.01 {
            h += b.lines[k0 + fit].height;
            fit += 1;
        }
        let rest = chunk_end - k0;
        let mut k = fit;
        if fit < rest {
            if p.widow_control {
                // No single last line at the top of the next page (widow)...
                if rest - k == 1 && k >= 2 {
                    k -= 1;
                }
                // ...and no single first line at the bottom of this one (orphan)
                if k == 1 && k0 == 0 && rest > 1 {
                    k = 0;
                }
            }
            if k == 0 && pag.at_top() {
                k = fit.max(1);
            }
        }
        for idx in k0..k0 + k {
            draw_line(pag, b, idx);
        }
        k0 += k;
        let hard_break = k0 == chunk_end && k > 0 && b.lines[chunk_end - 1].line.page_break_after;
        if k0 < n || hard_break {
            draw_left_border(pag, b, seg_top);
            pag.new_page();
            seg_top = pag.cursor_y;
        }
    }
    draw_left_border(pag, b, seg_top);

    if let Some(ref bottom) = p.borders.bottom {
        let empty = b.lines.len() == 1 && b.lines[0].line.text.is_empty();
        let y = if empty { pag.cursor_y } else { pag.cursor_y + 3.0 };
        pag.items.push(RenderCommand::Line {
            x1: pag.geo.margin_left,
            y1: y,
            x2: pag.geo.page_w - pag.geo.margin_right,
            y2: y,
            color: css_color(&bottom.color, "#CBD5E1"),
            line_width: bottom.sz_px.max(1.0),
        });
        pag.cursor_y += bottom.space.max(3.0) + bottom.sz_px;
    }

    flow.pending_after = b.space_after;
    flow.prev_style = Some(p.style.clone());
    flow.prev_contextual = p.contextual_spacing;

    if p.section_break_after && !pag.at_top() {
        pag.new_page();
    }
}

fn draw_line(pag: &mut Paginator, b: &ParagraphBox, idx: usize) {
    let lb = &b.lines[idx];
    let baseline_y = pag.cursor_y + lb.baseline;
    if !lb.line.text.is_empty() || !lb.line.runs.is_empty() {
        pag.items.push(RenderCommand::Text {
            text: lb.line.text.clone(),
            x: lb.line.x,
            y: baseline_y,
            width: lb.line.width,
            height: lb.height,
            font_size: b.font_size,
            font_family: b.font_family.clone(),
            font_weight: b.font_weight.to_string(),
            font_style: b.font_style.to_string(),
            color: b.color.clone(),
            align: b.p.align.clone(),
            paragraph_index: b.p.index,
            line_index: idx,
            is_last_line: lb.line.is_last_line,
            max_width: lb.line.max_width,
            runs: lb.line.runs.clone(),
        });
    }

    // List number / bullet, drawn in the hanging indent of the first line
    if idx == 0 {
        if let (Some(label), Some(geo)) = (b.p.list_label.as_ref(), b.label.as_ref()) {
            pag.items.push(RenderCommand::Text {
                text: label.text.clone(),
                x: pag.geo.margin_left + geo.label_offset,
                y: baseline_y,
                width: geo.label_width,
                height: lb.height,
                font_size: geo.font_size,
                font_family: b.font_family.clone(),
                font_weight: if label.bold { "700" } else { "400" }.to_string(),
                font_style: if label.italic { "italic" } else { "normal" }.to_string(),
                color: b.color.clone(),
                align: "left".to_string(),
                paragraph_index: b.p.index,
                line_index: 0,
                is_last_line: true,
                max_width: geo.label_width,
                runs: vec![TextRun {
                    text: label.text.clone(),
                    bold: label.bold,
                    italic: label.italic,
                    underline: label.underline,
                    color: label.color.clone(),
                    font_size: Some(geo.font_size),
                    font_family: label.font_family.clone(),
                    width: geo.label_width,
                }],
            });
        }
    }
    pag.cursor_y += lb.height;
}

/// Vertical accent bar for the part of the paragraph on the current page
fn draw_left_border(pag: &mut Paginator, b: &ParagraphBox, seg_top: f64) {
    let Some(ref left) = b.p.borders.left else { return };
    if pag.cursor_y <= seg_top {
        return;
    }
    let x = (pag.geo.margin_left + b.p.indent_left.max(0.0) - 7.0).max(10.0);
    pag.items.push(RenderCommand::Line {
        x1: x,
        y1: seg_top + 1.0,
        x2: x,
        y2: pag.cursor_y - 1.0,
        color: css_color(&left.color, "#334155"),
        line_width: left.sz_px.max(2.5),
    });
}

fn place_table(pag: &mut Paginator, t: &TableBox, flow: &mut FlowState) {
    if !pag.at_top() {
        pag.cursor_y += flow.pending_after;
    }
    pag.cursor_y += 8.0;
    let tbl = t.tbl;
    let num_cols = t.col_widths.len().max(1);

    for (row_idx, row) in t.rows.iter().enumerate() {
        if row.height > pag.remaining() && !pag.at_top() {
            pag.new_page();
        }
        let rich_row = tbl.rich_rows.get(row_idx);
        let mut x = pag.geo.margin_left;
        for (col_idx, cell_text) in tbl.rows[row_idx].iter().enumerate() {
            let col_w = t.col_widths.get(col_idx).copied().unwrap_or(pag.geo.printable_w / num_cols as f64);
            let cell = rich_row.and_then(|r| r.cells.get(col_idx));
            pag.items.push(RenderCommand::TableCell {
                table_index: tbl.index,
                row: row_idx,
                col: col_idx,
                x,
                y: pag.cursor_y,
                width: col_w,
                height: row.height,
                text: cell_text.clone(),
                lines: row.cell_lines.get(col_idx).cloned().unwrap_or_else(|| vec![cell_text.clone()]),
                is_header: row.is_header,
                bg_color: cell.and_then(|c| c.bg_color.clone()),
                color: cell
                    .map(|c| c.color.clone())
                    .unwrap_or_else(|| if row.is_header { "FAF7F0".to_string() } else { "1B1F1E".to_string() }),
                font_size: cell.map(|c| c.font_size * PX_PER_PT).unwrap_or(if row.is_header { 13.33 } else { 12.0 }),
                font_family: cell.map(|c| c.font_family.clone()).unwrap_or_else(|| "Calibri, Inter, sans-serif".to_string()),
                font_weight: if cell.map(|c| c.bold).unwrap_or(row.is_header) { "700" } else { "400" }.to_string(),
                align: cell.map(|c| c.align.clone()).unwrap_or_else(|| "left".to_string()),
                border_color: cell.map(|c| c.border_color.clone()).unwrap_or_else(|| "DDD5C2".to_string()),
            });
            x += col_w;
        }
        pag.cursor_y += row.height;
    }

    flow.pending_after = 16.0;
    flow.prev_style = None;
    flow.prev_contextual = false;
}

#[derive(Debug, Clone)]
struct LayoutLine {
    pub text: String,
    pub x: f64,
    pub width: f64,
    pub max_width: f64,
    pub is_last_line: bool,
    /// A manual page break ends this line
    pub page_break_after: bool,
    pub runs: Vec<TextRun>,
}

#[derive(Clone, Debug)]
enum SegmentKind {
    Newline,
    PageBreak,
    Tab,
    Space(String),
    Text(String),
}

#[derive(Clone, Debug)]
struct LayoutSegment {
    kind: SegmentKind,
    bold: bool,
    italic: bool,
    underline: bool,
    color: String,
    font_size: f64,
    font_family: Option<String>,
}

enum AtomicUnit {
    Newline,
    PageBreak,
    Tab,
    Spaces(Vec<LayoutSegment>),
    Word(Vec<LayoutSegment>),
}

fn extract_segments(runs: &[RunInfo], default_color: &str, default_fs: f64) -> Vec<LayoutSegment> {
    let mut segments = Vec::new();

    for r in runs {
        let run_color = if !r.color.is_empty() {
            if r.color.starts_with('#') {
                r.color.clone()
            } else {
                format!("#{}", r.color)
            }
        } else {
            default_color.to_string()
        };
        let run_fs = r.font_size.unwrap_or(default_fs);

        let mut current_text = String::new();
        let mut current_spaces = String::new();

        for ch in r.text.chars() {
            if ch == '\n' || ch == PAGE_BREAK {
                if !current_text.is_empty() {
                    segments.push(LayoutSegment {
                        kind: SegmentKind::Text(current_text.clone()),
                        bold: r.bold,
                        italic: r.italic,
                        underline: r.underline,
                        color: run_color.clone(),
                        font_size: run_fs,
                        font_family: r.font_family.clone(),
                    });
                    current_text.clear();
                }
                if !current_spaces.is_empty() {
                    segments.push(LayoutSegment {
                        kind: SegmentKind::Space(current_spaces.clone()),
                        bold: r.bold,
                        italic: r.italic,
                        underline: r.underline,
                        color: run_color.clone(),
                        font_size: run_fs,
                        font_family: r.font_family.clone(),
                    });
                    current_spaces.clear();
                }
                segments.push(LayoutSegment {
                    kind: if ch == PAGE_BREAK { SegmentKind::PageBreak } else { SegmentKind::Newline },
                    bold: r.bold,
                    italic: r.italic,
                    underline: r.underline,
                    color: run_color.clone(),
                    font_size: run_fs,
                    font_family: r.font_family.clone(),
                });
            } else if ch == '\t' {
                if !current_text.is_empty() {
                    segments.push(LayoutSegment {
                        kind: SegmentKind::Text(current_text.clone()),
                        bold: r.bold,
                        italic: r.italic,
                        underline: r.underline,
                        color: run_color.clone(),
                        font_size: run_fs,
                        font_family: r.font_family.clone(),
                    });
                    current_text.clear();
                }
                if !current_spaces.is_empty() {
                    segments.push(LayoutSegment {
                        kind: SegmentKind::Space(current_spaces.clone()),
                        bold: r.bold,
                        italic: r.italic,
                        underline: r.underline,
                        color: run_color.clone(),
                        font_size: run_fs,
                        font_family: r.font_family.clone(),
                    });
                    current_spaces.clear();
                }
                segments.push(LayoutSegment {
                    kind: SegmentKind::Tab,
                    bold: r.bold,
                    italic: r.italic,
                    underline: r.underline,
                    color: run_color.clone(),
                    font_size: run_fs,
                    font_family: r.font_family.clone(),
                });
            } else if ch.is_whitespace() {
                if !current_text.is_empty() {
                    segments.push(LayoutSegment {
                        kind: SegmentKind::Text(current_text.clone()),
                        bold: r.bold,
                        italic: r.italic,
                        underline: r.underline,
                        color: run_color.clone(),
                        font_size: run_fs,
                        font_family: r.font_family.clone(),
                    });
                    current_text.clear();
                }
                current_spaces.push(ch);
            } else {
                if !current_spaces.is_empty() {
                    segments.push(LayoutSegment {
                        kind: SegmentKind::Space(current_spaces.clone()),
                        bold: r.bold,
                        italic: r.italic,
                        underline: r.underline,
                        color: run_color.clone(),
                        font_size: run_fs,
                        font_family: r.font_family.clone(),
                    });
                    current_spaces.clear();
                }
                current_text.push(ch);
            }
        }

        if !current_text.is_empty() {
            segments.push(LayoutSegment {
                kind: SegmentKind::Text(current_text),
                bold: r.bold,
                italic: r.italic,
                underline: r.underline,
                color: run_color.clone(),
                font_size: run_fs,
                font_family: r.font_family.clone(),
            });
        }
        if !current_spaces.is_empty() {
            segments.push(LayoutSegment {
                kind: SegmentKind::Space(current_spaces),
                bold: r.bold,
                italic: r.italic,
                underline: r.underline,
                color: run_color,
                font_size: run_fs,
                font_family: r.font_family.clone(),
            });
        }
    }

    segments
}

fn group_into_units(segments: Vec<LayoutSegment>) -> Vec<AtomicUnit> {
    let mut units = Vec::new();
    let mut current_word: Vec<LayoutSegment> = Vec::new();
    let mut current_spaces: Vec<LayoutSegment> = Vec::new();

    for seg in segments {
        match seg.kind {
            SegmentKind::Newline | SegmentKind::PageBreak => {
                if !current_word.is_empty() {
                    units.push(AtomicUnit::Word(std::mem::take(&mut current_word)));
                }
                if !current_spaces.is_empty() {
                    units.push(AtomicUnit::Spaces(std::mem::take(&mut current_spaces)));
                }
                units.push(if matches!(seg.kind, SegmentKind::PageBreak) {
                    AtomicUnit::PageBreak
                } else {
                    AtomicUnit::Newline
                });
            }
            SegmentKind::Tab => {
                if !current_word.is_empty() {
                    units.push(AtomicUnit::Word(std::mem::take(&mut current_word)));
                }
                if !current_spaces.is_empty() {
                    units.push(AtomicUnit::Spaces(std::mem::take(&mut current_spaces)));
                }
                units.push(AtomicUnit::Tab);
            }
            SegmentKind::Space(_) => {
                if !current_word.is_empty() {
                    units.push(AtomicUnit::Word(std::mem::take(&mut current_word)));
                }
                current_spaces.push(seg);
            }
            SegmentKind::Text(_) => {
                if !current_spaces.is_empty() {
                    units.push(AtomicUnit::Spaces(std::mem::take(&mut current_spaces)));
                }
                current_word.push(seg);
            }
        }
    }

    if !current_word.is_empty() {
        units.push(AtomicUnit::Word(current_word));
    }
    if !current_spaces.is_empty() {
        units.push(AtomicUnit::Spaces(current_spaces));
    }

    units
}

fn append_segment_to_runs(runs: &mut Vec<TextRun>, text: &str, seg: &LayoutSegment, width: f64) {
    if let Some(last) = runs.last_mut() {
        if last.bold == seg.bold
            && last.italic == seg.italic
            && last.underline == seg.underline
            && last.color == seg.color
            && last.font_size == Some(seg.font_size)
            && last.font_family == seg.font_family
            && last.text != "\t"
        {
            last.text.push_str(text);
            last.width += width;
            return;
        }
    }
    runs.push(TextRun {
        text: text.to_string(),
        bold: seg.bold,
        italic: seg.italic,
        underline: seg.underline,
        color: seg.color.clone(),
        font_size: Some(seg.font_size),
        font_family: seg.font_family.clone(),
        width,
    });
}

fn seg_width(m: &mut dyn TextMeasurer, seg: &LayoutSegment, text: &str, family_css: &str) -> f64 {
    m.measure(
        text,
        &FontSpec {
            family: seg.font_family.as_deref().unwrap_or(family_css),
            size: seg.font_size,
            bold: seg.bold,
            italic: seg.italic,
        },
    )
}

fn layout_paragraph_lines(
    p: &ParagraphInfo,
    font_size: f64,
    default_color: &str,
    family_css: &str,
    margin_left: f64,
    printable_width: f64,
    m: &mut dyn TextMeasurer,
) -> Vec<LayoutLine> {
    let first_line_indent = list_label_geometry(p, font_size, family_css, m)
        .map(|g| g.text_offset)
        .unwrap_or_else(|| (p.indent_left + p.indent_first_line).max(0.0));
    let line_indent = |idx: usize| if idx == 0 { first_line_indent } else { p.indent_left.max(0.0) };
    let max_width = |idx: usize| (printable_width - line_indent(idx) - p.indent_right.max(0.0)).max(60.0);
    let flush = |runs: Vec<TextRun>, line_w: f64, idx: usize, is_last: bool, page_break_after: bool| {
        let indent = line_indent(idx);
        let max_w = max_width(idx);
        let x = match p.align.as_str() {
            "center" => margin_left + indent + (max_w - line_w).max(0.0) / 2.0,
            "right" => margin_left + indent + (max_w - line_w).max(0.0),
            _ => margin_left + indent,
        };
        LayoutLine {
            text: runs.iter().map(|r| r.text.as_str()).collect(),
            x,
            width: line_w,
            max_width: max_w,
            is_last_line: is_last,
            page_break_after,
            runs,
        }
    };

    if p.text.trim().is_empty() && p.runs.is_empty() {
        return vec![flush(Vec::new(), 0.0, 0, true, false)];
    }

    // Raw runs with font sizes scaled to px
    let source_runs: Vec<RunInfo> = if !p.runs.is_empty() {
        p.runs
            .iter()
            .map(|r| RunInfo { font_size: r.font_size.map(|pt| pt * PX_PER_PT), ..r.clone() })
            .collect()
    } else {
        vec![RunInfo {
            text: p.text.clone(),
            bold: p.bold,
            italic: p.italic,
            underline: false,
            color: if !p.color.is_empty() { p.color.clone() } else { default_color.to_string() },
            font_size: p.font_size.map(|pt| pt * PX_PER_PT),
            font_family: p.font_family.clone(),
        }]
    };

    let units = group_into_units(extract_segments(&source_runs, default_color, font_size));

    let mut lines = Vec::new();
    let mut runs: Vec<TextRun> = Vec::new();
    let mut line_w = 0.0;
    let mut pending_spaces: Vec<LayoutSegment> = Vec::new();
    let mut idx = 0;

    for unit in units {
        match unit {
            AtomicUnit::Newline | AtomicUnit::PageBreak => {
                let page_break = matches!(unit, AtomicUnit::PageBreak);
                lines.push(flush(std::mem::take(&mut runs), line_w, idx, true, page_break));
                line_w = 0.0;
                pending_spaces.clear();
                idx += 1;
            }
            AtomicUnit::Tab => {
                // Default tab stops every 0.5in, measured from the left margin
                const DEFAULT_TAB: f64 = 48.0;
                let pos = line_indent(idx) + line_w;
                let advance = (((pos / DEFAULT_TAB).floor() + 1.0) * DEFAULT_TAB - pos).max(1.0);
                if line_w + advance > max_width(idx) && !runs.is_empty() {
                    lines.push(flush(std::mem::take(&mut runs), line_w, idx, false, false));
                    line_w = 0.0;
                    idx += 1;
                }
                runs.push(TextRun {
                    text: "\t".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                    color: default_color.to_string(),
                    font_size: Some(font_size),
                    font_family: None,
                    width: advance,
                });
                line_w += advance;
                pending_spaces.clear();
            }
            AtomicUnit::Spaces(spaces) => pending_spaces = spaces,
            AtomicUnit::Word(word_segs) => {
                let pieces: Vec<(String, LayoutSegment, f64)> = word_segs
                    .into_iter()
                    .filter_map(|seg| match seg.kind {
                        SegmentKind::Text(ref t) => {
                            let w = seg_width(m, &seg, t, family_css);
                            Some((t.clone(), seg, w))
                        }
                        _ => None,
                    })
                    .collect();
                let word_w: f64 = pieces.iter().map(|(_, _, w)| w).sum();
                let spaces: Vec<(String, LayoutSegment, f64)> = if runs.is_empty() {
                    Vec::new()
                } else {
                    pending_spaces
                        .drain(..)
                        .filter_map(|sp| match sp.kind {
                            SegmentKind::Space(ref s) => {
                                let w = seg_width(m, &sp, s, family_css);
                                Some((s.clone(), sp, w))
                            }
                            _ => None,
                        })
                        .collect()
                };
                pending_spaces.clear();
                let spaces_w: f64 = spaces.iter().map(|(_, _, w)| w).sum();

                if line_w + spaces_w + word_w > max_width(idx) && !runs.is_empty() {
                    lines.push(flush(std::mem::take(&mut runs), line_w, idx, false, false));
                    line_w = 0.0;
                    idx += 1;
                } else {
                    for (s, seg, w) in &spaces {
                        append_segment_to_runs(&mut runs, s, seg, *w);
                        line_w += w;
                    }
                }
                for (t, seg, w) in &pieces {
                    append_segment_to_runs(&mut runs, t, seg, *w);
                    line_w += w;
                }
            }
        }
    }

    if !runs.is_empty() {
        lines.push(flush(runs, line_w, idx, true, false));
    }
    if lines.is_empty() {
        lines.push(flush(Vec::new(), 0.0, 0, true, false));
    }
    lines
}

fn add_watermark_command(items: &mut Vec<RenderCommand>, text: &str, opacity: f64, page_w: f64, page_h: f64) {
    items.insert(
        0,
        RenderCommand::Watermark {
            text: text.to_string(),
            x: page_w / 2.0,
            y: page_h / 2.0,
            opacity: if opacity <= 0.0 { 0.2 } else { opacity },
            font_size: (page_w * 0.09).clamp(48.0, 84.0),
            color: "#64748B".to_string(),
            rotation_deg: -30.0,
        },
    );
}

fn get_paragraph_typography(p: &ParagraphInfo) -> (f64, f64, &'static str, String, String) {
    let lower_style = p.style.to_lowercase();
    let effective_family = p.font_family.as_deref()
        .or_else(|| p.runs.iter().find_map(|r| r.font_family.as_deref()));

    let default_family = if let Some(fam) = effective_family {
        if fam.eq_ignore_ascii_case("consolas") || fam.eq_ignore_ascii_case("courier") || fam.eq_ignore_ascii_case("monospace") {
            "Consolas, Cousine, 'JetBrains Mono', monospace".to_string()
        } else if fam.eq_ignore_ascii_case("calibri") {
            "Calibri, Carlito, 'Segoe UI', Inter, sans-serif".to_string()
        } else if fam.eq_ignore_ascii_case("cambria") {
            "Cambria, Caladea, Georgia, serif".to_string()
        } else if fam.eq_ignore_ascii_case("times new roman") || fam.eq_ignore_ascii_case("times") {
            "'Times New Roman', Tinos, Georgia, serif".to_string()
        } else if fam.eq_ignore_ascii_case("arial") {
            "Arial, Arimo, Helvetica, sans-serif".to_string()
        } else if fam.eq_ignore_ascii_case("aptos") {
            "Aptos, Calibri, Carlito, 'Segoe UI', sans-serif".to_string()
        } else {
            format!("{}, Calibri, Carlito, Inter, sans-serif", fam)
        }
    } else {
        "Calibri, Carlito, 'Segoe UI', Inter, sans-serif".to_string()
    };

    if let Some(sz_pt) = p.font_size {
        let font_size_px = sz_pt * 1.3333;
        let weight = if p.bold || p.is_heading { "700" } else { "400" };
        let lh = font_size_px * p.line_spacing.unwrap_or(1.35).clamp(1.15, 2.0);
        let color = if !p.color.is_empty() {
            if p.color.starts_with('#') { p.color.clone() } else { format!("#{}", p.color) }
        } else {
            "#1B1F1E".to_string()
        };
        (font_size_px, lh, weight, default_family, color)
    } else if lower_style.contains("heading1") || lower_style.contains("title") || lower_style.contains("título") {
        (24.0, 32.0, "700", "Calibri, Carlito, Outfit, sans-serif".to_string(), "#1B1F1E".to_string())
    } else if lower_style.contains("heading2") || lower_style.contains("subtítulo") {
        (18.0, 26.0, "600", "Calibri, Carlito, Outfit, sans-serif".to_string(), "#1E40AF".to_string())
    } else if lower_style.contains("heading3") {
        (15.0, 22.0, "600", "Calibri, Carlito, Inter, sans-serif".to_string(), "#334155".to_string())
    } else {
        let weight = if p.bold { "700" } else { "400" };
        (14.66, 20.0, weight, default_family, "#1E293B".to_string())
    }
}

/// Where a paragraph's list label goes, relative to the left margin (px)
struct ListLabelGeometry {
    label_offset: f64,
    label_width: f64,
    /// Where the first line's text starts
    text_offset: f64,
    font_size: f64,
}

/// Word places the label at `left - hanging` and tabs to `left`; when the label is wider
/// than the hanging indent the text moves to the next default tab stop (0.5in)
fn list_label_geometry(
    p: &ParagraphInfo,
    paragraph_font_size: f64,
    family_css: &str,
    m: &mut dyn TextMeasurer,
) -> Option<ListLabelGeometry> {
    let label = p.list_label.as_ref()?;
    let font_size = label.font_size.map(|pt| pt * PX_PER_PT).unwrap_or(paragraph_font_size);
    let label_offset = (p.indent_left + p.indent_first_line).max(0.0);
    let font = FontSpec {
        family: label.font_family.as_deref().unwrap_or(family_css),
        size: font_size,
        bold: label.bold,
        italic: label.italic,
    };
    let label_width = m.measure(&label.text, &font);
    let label_end = label_offset + label_width + font_size * 0.2;
    let text_offset = if label.text.ends_with(' ') {
        label_offset + label_width
    } else if p.indent_first_line < 0.0 && label_end <= p.indent_left {
        p.indent_left
    } else {
        const DEFAULT_TAB: f64 = 48.0;
        (label_end / DEFAULT_TAB).ceil() * DEFAULT_TAB
    };
    Some(ListLabelGeometry { label_offset, label_width, text_offset, font_size })
}

/// Greedy word wrapping for table cells
fn wrap_text(text: &str, max_width: f64, font: &FontSpec, m: &mut dyn TextMeasurer) -> Vec<String> {
    let words: Vec<&str> = text.split_whitespace().collect();
    if words.is_empty() {
        return vec![String::new()];
    }

    let mut lines = Vec::new();
    let mut current_line = String::new();
    for word in words {
        let test_line = if current_line.is_empty() {
            word.to_string()
        } else {
            format!("{} {}", current_line, word)
        };
        if m.measure(&test_line, font) > max_width && !current_line.is_empty() {
            lines.push(std::mem::replace(&mut current_line, word.to_string()));
        } else {
            current_line = test_line;
        }
    }
    if !current_line.is_empty() {
        lines.push(current_line);
    }
    lines
}

/// Estimates text width with proportional font metrics
fn estimate_text_width(text: &str, font_size: f64) -> f64 {
    let avg_char_width = font_size * 0.54;
    text.chars()
        .map(|c| match c {
            'i' | 'l' | 'j' | '!' | '|' | ':' | ';' | '.' | ',' | '\'' => font_size * 0.28,
            'f' | 't' | 'r' | 'I' => font_size * 0.35,
            'm' | 'w' | 'M' | 'W' => font_size * 0.85,
            'A'..='Z' => font_size * 0.65,
            ' ' => font_size * 0.3,
            _ => avg_char_width,
        })
        .sum()
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Arial 12pt: one line is 16px × 1.15 = 18.4px; a default page holds 990px of content
    fn para(index: usize, text: &str) -> ParagraphInfo {
        ParagraphInfo {
            index,
            text: text.to_string(),
            style: "Normal".to_string(),
            align: "left".to_string(),
            font_size: Some(12.0),
            font_family: Some("Arial".to_string()),
            spacing_resolved: true,
            widow_control: true,
            runs: vec![RunInfo {
                text: text.to_string(),
                font_size: Some(12.0),
                font_family: Some("Arial".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        }
    }

    fn lines(index: usize, n: usize) -> ParagraphInfo {
        let text = (0..n).map(|i| format!("l{}", i)).collect::<Vec<_>>().join("\n");
        para(index, &text)
    }

    fn layout(paragraphs: Vec<ParagraphInfo>) -> DocumentLayout {
        let elements: Vec<DocumentElement> = paragraphs.into_iter().map(DocumentElement::Paragraph).collect();
        LayoutEngine::new().compute_layout(
            &elements,
            "FFFFFF",
            &PageSetup::default(),
            &HeaderFooterInfo::default(),
            None,
            None,
            0.0,
        )
    }

    /// (page number, y, height) of every text line of a paragraph
    fn placed(layout: &DocumentLayout, paragraph: usize) -> Vec<(usize, f64, f64)> {
        layout
            .pages
            .iter()
            .flat_map(|page| {
                page.items.iter().filter_map(move |item| match item {
                    // Page decorations ("Página 1 de 2") carry no runs
                    RenderCommand::Text { paragraph_index, y, height, runs, .. }
                        if *paragraph_index == paragraph && !runs.is_empty() =>
                    {
                        Some((page.page_number, *y, *height))
                    }
                    _ => None,
                })
            })
            .collect()
    }

    fn pages_of(layout: &DocumentLayout, paragraph: usize) -> Vec<usize> {
        placed(layout, paragraph).into_iter().map(|(p, _, _)| p).collect()
    }

    #[test]
    fn test_manual_page_break_moves_following_text() {
        let l = layout(vec![para(0, &format!("uno{}dos", crate::docx_parser::PAGE_BREAK))]);
        assert_eq!(pages_of(&l, 0), vec![1, 2]);
    }

    #[test]
    fn test_page_break_before_and_section_breaks() {
        let mut second = para(1, "segundo");
        second.page_break_before = true;
        let l = layout(vec![para(0, "primero"), second]);
        assert_eq!((pages_of(&l, 0), pages_of(&l, 1)), (vec![1], vec![2]));

        let mut first = para(0, "fin de sección");
        first.section_break_after = true;
        let l = layout(vec![first, para(1, "nueva sección")]);
        assert_eq!(pages_of(&l, 1), vec![2]);
    }

    #[test]
    fn test_widow_control_moves_two_lines() {
        // 50 lines leave room for 3 more; a 4-line paragraph would leave 1 line alone
        let l = layout(vec![lines(0, 50), lines(1, 4)]);
        assert_eq!(pages_of(&l, 1), vec![1, 1, 2, 2]);

        let mut no_control = lines(1, 4);
        no_control.widow_control = false;
        let l = layout(vec![lines(0, 50), no_control]);
        assert_eq!(pages_of(&l, 1), vec![1, 1, 1, 2]);
    }

    #[test]
    fn test_orphan_control_moves_whole_paragraph() {
        // 52 lines leave room for exactly one more line
        let mut b = lines(1, 4);
        b.widow_control = true;
        let l = layout(vec![lines(0, 52), b]);
        assert_eq!(pages_of(&l, 1), vec![2, 2, 2, 2]);
    }

    #[test]
    fn test_keep_next_moves_heading_with_its_paragraph() {
        // Room for 2 lines: the heading fits, but not with the first lines of its paragraph
        let mut heading = para(1, "Título");
        heading.keep_next = true;
        let l = layout(vec![lines(0, 51), heading.clone(), lines(2, 3)]);
        assert_eq!(pages_of(&l, 1), vec![2]);
        assert_eq!(pages_of(&l, 2), vec![2, 2, 2]);

        heading.keep_next = false;
        let l = layout(vec![lines(0, 51), heading, lines(2, 3)]);
        assert_eq!(pages_of(&l, 1), vec![1]);
    }

    #[test]
    fn test_keep_lines_keeps_paragraph_together() {
        let mut b = lines(1, 5);
        b.keep_lines = true;
        b.widow_control = false;
        let l = layout(vec![lines(0, 50), b]);
        assert_eq!(pages_of(&l, 1), vec![2; 5]);
    }

    #[test]
    fn test_line_spacing_rules() {
        let mut exact = para(0, "exacto");
        exact.line_exact = Some(24.0);
        let mut at_least = para(1, "mínimo");
        at_least.line_at_least = Some(6.0);
        let mut double = para(2, "doble");
        double.line_spacing = Some(2.0);
        let l = layout(vec![exact, at_least, double]);
        let h = |i| placed(&l, i)[0].2;
        assert!((h(0) - 32.0).abs() < 0.01, "24pt exact = 32px, got {}", h(0));
        assert!((h(1) - 18.4).abs() < 0.01, "atLeast below natural keeps natural, got {}", h(1));
        assert!((h(2) - 36.8).abs() < 0.01, "double = 2 × natural, got {}", h(2));
    }

    #[test]
    fn test_spacing_in_points_and_contextual_spacing() {
        let spaced = |i: usize, contextual: bool| {
            let mut p = para(i, "x");
            p.style = "ListParagraph".to_string();
            p.space_before = 6.0;
            p.space_after = 6.0;
            p.contextual_spacing = contextual;
            p
        };
        let gap = |l: &DocumentLayout| placed(l, 1)[0].1 - placed(l, 0)[0].1;

        // 6pt after + 6pt before = 16px between the two lines (18.4px each)
        let l = layout(vec![spaced(0, false), spaced(1, false)]);
        assert!((gap(&l) - (18.4 + 16.0)).abs() < 0.01, "got {}", gap(&l));

        let l = layout(vec![spaced(0, true), spaced(1, true)]);
        assert!((gap(&l) - 18.4).abs() < 0.01, "same-style contextual spacing collapses, got {}", gap(&l));
    }

    #[test]
    fn test_custom_measurer_drives_line_breaks() {
        struct Wide;
        impl TextMeasurer for Wide {
            fn measure(&mut self, text: &str, _font: &FontSpec) -> f64 {
                text.chars().count() as f64 * 100.0
            }
        }
        let p = para(0, "aaa bbb ccc");
        let elements = vec![DocumentElement::Paragraph(p)];
        let l = LayoutEngine::new().compute_layout_with(
            &elements,
            "FFFFFF",
            &PageSetup::default(),
            &HeaderFooterInfo::default(),
            None,
            None,
            0.0,
            &mut Wide,
        );
        // Each 3-letter word is 300px wide and the line holds 670px: one word per line... plus a space
        assert_eq!(placed(&l, 0).len(), 3);
    }
}
