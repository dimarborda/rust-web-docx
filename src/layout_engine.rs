use crate::docx_parser::{
    is_dark_hex_str, DocumentElement, HeaderFooterInfo, ImageRef, PageSetup, ParagraphInfo, RunInfo, TableCellData,
    TableInfo, PAGE_BREAK,
};
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
    /// Absolute x where the run starts (justification included)
    #[serde(default)]
    pub x: f64,
    /// Character offset of the run in its paragraph's text
    #[serde(default)]
    pub start: usize,
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
        /// Present on paragraph body lines (not on list labels or page decorations)
        #[serde(default, skip_serializing_if = "Option::is_none")]
        line: Option<LineRange>,
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

/// Which part of its paragraph a laid out line shows, for caret placement and hit testing
#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
pub struct LineRange {
    /// Character offsets in the paragraph text: `[start, end)`
    pub start: usize,
    pub end: usize,
    /// Top of the line box (the text item's `y` is the baseline)
    pub top: f64,
    /// Extra px added to each space on justified lines
    pub space_extra: f64,
    /// Horizontal bounds of the line box (the paragraph's column, e.g. a table cell)
    #[serde(default)]
    pub left: f64,
    #[serde(default)]
    pub right: f64,
}

pub struct LayoutEngine {
    pub page_width: f64,
    pub page_height: f64,
    /// Pixels (data URLs) of the body's pictures by relationship id
    images: HashMap<String, String>,
}

impl LayoutEngine {
    pub fn new() -> Self {
        Self {
            page_width: PAGE_WIDTH,
            page_height: PAGE_HEIGHT,
            images: HashMap::new(),
        }
    }

    /// Provides the pixels of the body's pictures (relationship id → data URL)
    pub fn with_images(mut self, images: HashMap<String, String>) -> Self {
        self.images = images;
        self
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
        let header_inline_bottom = header_footer
            .header_images
            .iter()
            .filter(|p| !p.image.anchored)
            .map(|p| page_setup.header_margin + p.image.height)
            .fold(0.0, f64::max);
        let body_top = margin_top.max(header_inline_bottom);

        let geo = PageGeometry {
            page_w,
            page_h,
            margin_left,
            margin_right,
            printable_w: (page_w - margin_left - margin_right).max(100.0),
            top_first: body_top,
            top: body_top,
            bottom: page_h - margin_bottom,
            bg: css_color(bg_color, "#FFFFFF"),
        };

        // Pass 1: measure every block; pass 2: paginate with look-ahead for keep rules
        let blocks: Vec<Block> = elements
            .iter()
            .map(|el| match el {
                DocumentElement::Paragraph(p) => {
                    Block::Paragraph(prepare_paragraph(p, geo.margin_left, geo.printable_w, None, measurer))
                }
                DocumentElement::Table(t) => Block::Table(prepare_table(t, &geo, measurer)),
            })
            .collect();

        let mut pag = Paginator::new(&geo, watermark, watermark_opacity, &self.images);
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

            // 1. Full page background chosen in the editor ("Fondo / Marca de agua")
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

            // 2. Header and footer pictures, positioned like the paragraph that anchors them
            let column = (geo.margin_left, page_w - geo.margin_right);
            let header_top = page_setup.header_margin;
            let footer_top = page_h - page_setup.footer_margin;
            for (placed, para_top) in header_footer
                .header_images
                .iter()
                .map(|p| (p, header_top))
                .chain(header_footer.footer_images.iter().map(|p| (p, footer_top - p.image.height)))
            {
                let (x, y) = if placed.image.anchored {
                    anchored_position(&placed.image, &geo, column, para_top)
                } else {
                    (column.0, para_top)
                };
                let cmd = image_command(&placed.image, &placed.data_url, x, y);
                if placed.image.behind_text {
                    page.items.insert(0, cmd);
                } else {
                    page.items.push(cmd);
                }
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
                    line: None,
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
                    line: None,
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
                    line: None,
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

    /// Reuses measurements from earlier layouts
    pub fn with_cache(inner: M, cache: HashMap<String, f64>) -> Self {
        CachedMeasurer { inner, cache }
    }

    /// Hands the measurements back for the next layout (dropped when they grow too large)
    pub fn into_cache(self) -> HashMap<String, f64> {
        if self.cache.len() > 300_000 {
            HashMap::new()
        } else {
            self.cache
        }
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
    /// Left edge and width of the paragraph's column (page margin, or a table cell's content box)
    left: f64,
    width: f64,
    /// Where the first line's inline pictures start (x) when the paragraph has any
    inline_images_x: f64,
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

/// A table cell: its box and its paragraphs laid out inside it
struct CellBox<'a> {
    x: f64,
    width: f64,
    cell: &'a TableCellData,
    paragraphs: Vec<ParagraphBox<'a>>,
    content_height: f64,
}

struct RowBox<'a> {
    height: f64,
    cells: Vec<CellBox<'a>>,
    is_header: bool,
}

struct TableBox<'a> {
    tbl: &'a TableInfo,
    rows: Vec<RowBox<'a>>,
}

/// Word's default cell margins: 0.08" left/right, none top/bottom
const CELL_PAD_X: f64 = 7.2;
const CELL_PAD_Y: f64 = 1.0;

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

/// Lays out a paragraph in the column starting at `left` with `width` px. `auto_color`
/// replaces the default text color (e.g. white text in dark table cells, like Word's "auto").
fn prepare_paragraph<'a>(
    p: &'a ParagraphInfo,
    left: f64,
    width: f64,
    auto_color: Option<&str>,
    m: &mut dyn TextMeasurer,
) -> ParagraphBox<'a> {
    let (font_size, _, font_weight, font_family, default_color) = get_paragraph_typography(p);
    let color = css_color(&p.color, auto_color.unwrap_or(&default_color));
    let lines = layout_paragraph_lines(p, font_size, &color, &font_family, left, width, m);
    let label = list_label_geometry(p, font_size, &font_family, m);

    let lines: Vec<LineBox> = lines
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

    // Inline pictures sit on the first line (approximation: before its text), which grows
    // to fit them with the text on the baseline, as in Word
    let inline: Vec<&ImageRef> = p.images.iter().filter(|i| !i.anchored).collect();
    let mut lines = lines;
    let mut inline_images_x = left;
    if !inline.is_empty() {
        let total_w: f64 = inline.iter().map(|i| i.width).sum();
        let tallest = inline.iter().map(|i| i.height).fold(0.0, f64::max);
        if let Some(first) = lines.first_mut() {
            if tallest > first.height {
                first.baseline += tallest - first.height;
                first.height = tallest;
            }
            let indent = (p.indent_left + p.indent_first_line).max(0.0);
            let free = (width - indent - total_w).max(0.0);
            inline_images_x = left + indent + match p.align.as_str() {
                "center" => free / 2.0,
                "right" => free,
                _ => 0.0,
            };
        }
    }

    ParagraphBox {
        p,
        left,
        width,
        inline_images_x,
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
    // Column widths come from the grid (twips → px); tables wider than the text area shrink
    let columns = tbl
        .rich_rows
        .iter()
        .map(|r| r.cells.len())
        .max()
        .unwrap_or(0)
        .max(tbl.grid_cols.len())
        .max(1);
    let grid_px: Vec<f64> = tbl.grid_cols.iter().map(|w| w / 15.0).collect();
    let grid_total: f64 = grid_px.iter().sum();
    let col_widths: Vec<f64> = if grid_total > 0.0 {
        let scale = if grid_total > geo.printable_w * 1.05 { geo.printable_w / grid_total } else { 1.0 };
        let mut widths: Vec<f64> = grid_px.iter().map(|w| w * scale).collect();
        while widths.len() < columns {
            widths.push(geo.printable_w / columns as f64);
        }
        widths
    } else {
        vec![geo.printable_w / columns as f64; columns]
    };

    let rows = tbl
        .rich_rows
        .iter()
        .map(|row| {
            let mut x = geo.margin_left;
            let cells: Vec<CellBox> = row
                .cells
                .iter()
                .enumerate()
                .map(|(col, cell)| {
                    let width = col_widths.get(col).copied().unwrap_or(geo.printable_w / columns as f64);
                    let dark = cell.bg_color.as_deref().is_some_and(is_dark_hex_str);
                    let paragraphs: Vec<ParagraphBox> = cell
                        .paragraphs
                        .iter()
                        .map(|p| {
                            let inner = (width - 2.0 * CELL_PAD_X).max(10.0);
                            prepare_paragraph(p, x + CELL_PAD_X, inner, dark.then_some("#FFFFFF"), m)
                        })
                        .collect();
                    let content_height = 2.0 * CELL_PAD_Y
                        + paragraphs.iter().map(|b| b.space_before + b.lines_height() + b.space_after).sum::<f64>();
                    let cell_box = CellBox { x, width, cell, paragraphs, content_height };
                    x += width;
                    cell_box
                })
                .collect();
            let height = cells.iter().map(|c| c.content_height).fold(8.0, f64::max);
            RowBox { height, cells, is_header: row.is_header }
        })
        .collect();

    TableBox { tbl, rows }
}

/// Places content top to bottom and starts new pages as needed
struct Paginator<'a> {
    geo: &'a PageGeometry,
    watermark: Option<&'a str>,
    watermark_opacity: f64,
    /// Body pictures by relationship id
    images: &'a HashMap<String, String>,
    pages: Vec<PageLayout>,
    items: Vec<RenderCommand>,
    /// Pictures in front of the text, drawn after everything else on the page
    front: Vec<RenderCommand>,
    cursor_y: f64,
}

impl<'a> Paginator<'a> {
    fn new(
        geo: &'a PageGeometry,
        watermark: Option<&'a str>,
        watermark_opacity: f64,
        images: &'a HashMap<String, String>,
    ) -> Self {
        Paginator {
            geo,
            watermark,
            watermark_opacity,
            images,
            pages: Vec::new(),
            items: Vec::new(),
            front: Vec::new(),
            cursor_y: geo.top_first,
        }
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
        items.append(&mut self.front);
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
        if !self.items.is_empty() || !self.front.is_empty() || self.pages.is_empty() {
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
    if idx == 0 {
        place_paragraph_images(pag, b, pag.cursor_y);
    }
    push_line_items(&mut pag.items, b, idx, pag.cursor_y);
    pag.cursor_y += b.lines[idx].height;
}

/// Puts a paragraph's pictures on the current page: floating ones where they are anchored
/// (behind or in front of the text), inline ones on the first line
fn place_paragraph_images(pag: &mut Paginator, b: &ParagraphBox, para_top: f64) {
    let mut inline_x = b.inline_images_x;
    for img in &b.p.images {
        let Some(data_url) = pag.images.get(&img.rel_id) else { continue };
        let (x, y) = if img.anchored {
            anchored_position(img, pag.geo, (b.left, b.left + b.width), para_top)
        } else {
            let first = &b.lines[0];
            let pos = (inline_x, para_top + first.height - img.height);
            inline_x += img.width;
            pos
        };
        let cmd = image_command(img, data_url, x, y);
        if img.behind_text {
            pag.items.insert(0, cmd);
        } else if img.anchored {
            pag.front.push(cmd);
        } else {
            pag.items.push(cmd);
        }
    }
}

/// Top-left corner of a floating picture, from what its position is relative to
fn anchored_position(img: &ImageRef, geo: &PageGeometry, column: (f64, f64), para_top: f64) -> (f64, f64) {
    let (left, right) = match img.h_relative.as_str() {
        "page" => (0.0, geo.page_w),
        "leftMargin" | "insideMargin" => (0.0, geo.margin_left),
        "rightMargin" | "outsideMargin" => (geo.page_w - geo.margin_right, geo.page_w),
        "column" | "character" => column,
        _ => (geo.margin_left, geo.page_w - geo.margin_right),
    };
    let x = match img.h_align.as_deref() {
        Some("center") => left + (right - left - img.width) / 2.0,
        Some("right") | Some("outside") => right - img.width,
        Some(_) => left,
        None => left + img.h_offset,
    };
    let (top, bottom) = match img.v_relative.as_str() {
        "page" => (0.0, geo.page_h),
        "topMargin" => (0.0, geo.top),
        "bottomMargin" => (geo.bottom, geo.page_h),
        "paragraph" | "line" => (para_top, para_top),
        _ => (geo.top, geo.bottom),
    };
    let y = match img.v_align.as_deref() {
        Some("center") => top + (bottom - top - img.height) / 2.0,
        Some("bottom") | Some("outside") => bottom - img.height,
        Some(_) => top,
        None => top + img.v_offset,
    };
    (x, y)
}

fn image_command(img: &ImageRef, data_url: &str, x: f64, y: f64) -> RenderCommand {
    RenderCommand::Image {
        data_url: data_url.to_string(),
        x,
        y,
        width: img.width,
        height: img.height,
        is_background: img.behind_text,
        opacity: 1.0,
    }
}

/// Render commands for one line of a paragraph whose line box starts at `top`
fn push_line_items(items: &mut Vec<RenderCommand>, b: &ParagraphBox, idx: usize, top: f64) {
    let lb = &b.lines[idx];
    let baseline_y = top + lb.baseline;
    // Empty lines are emitted too, so the caret can be placed on them
    {
        items.push(RenderCommand::Text {
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
            line: Some(LineRange {
                start: lb.line.start,
                end: lb.line.end,
                top,
                space_extra: lb.line.space_extra,
                left: lb.line.box_left,
                right: lb.line.box_right,
            }),
        });
    }

    // List number / bullet, drawn in the hanging indent of the first line
    if idx == 0 {
        if let (Some(label), Some(geo)) = (b.p.list_label.as_ref(), b.label.as_ref()) {
            items.push(RenderCommand::Text {
                text: label.text.clone(),
                x: b.left + geo.label_offset,
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
                    x: b.left + geo.label_offset,
                    start: 0,
                }],
                line: None,
            });
        }
    }
}

/// Vertical accent bar for the part of the paragraph on the current page
fn draw_left_border(pag: &mut Paginator, b: &ParagraphBox, seg_top: f64) {
    let Some(ref left) = b.p.borders.left else { return };
    if pag.cursor_y <= seg_top {
        return;
    }
    let x = (b.left + b.p.indent_left.max(0.0) - 7.0).max(10.0);
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
    pag.cursor_y += 4.0;
    let tbl = t.tbl;

    for (row_idx, row) in t.rows.iter().enumerate() {
        // Rows are not split across pages
        if row.height > pag.remaining() && !pag.at_top() {
            pag.new_page();
        }
        let top = pag.cursor_y;
        for (col_idx, cell) in row.cells.iter().enumerate() {
            // Background and borders; the text is drawn as paragraph lines below
            pag.items.push(RenderCommand::TableCell {
                table_index: tbl.index,
                row: row_idx,
                col: col_idx,
                x: cell.x,
                y: top,
                width: cell.width,
                height: row.height,
                text: String::new(),
                lines: Vec::new(),
                is_header: row.is_header,
                bg_color: cell.cell.bg_color.clone(),
                color: cell.cell.color.clone(),
                font_size: cell.cell.font_size * PX_PER_PT,
                font_family: cell.cell.font_family.clone(),
                font_weight: if cell.cell.bold { "700" } else { "400" }.to_string(),
                align: cell.cell.align.clone(),
                border_color: cell.cell.border_color.clone(),
            });
            let mut y = top + CELL_PAD_Y;
            for b in &cell.paragraphs {
                y += b.space_before;
                place_paragraph_images(pag, b, y);
                for idx in 0..b.lines.len() {
                    push_line_items(&mut pag.items, b, idx, y);
                    y += b.lines[idx].height;
                }
                y += b.space_after;
            }
        }
        pag.cursor_y += row.height;
    }

    flow.pending_after = 12.0;
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
    /// Character range of the paragraph text covered by this line. For wrapped lines `end` is
    /// where the next line starts (trailing spaces belong to this line but are not drawn)
    pub start: usize,
    pub end: usize,
    /// Extra px added to every space when the line is justified
    pub space_extra: f64,
    /// Horizontal bounds of the line box
    pub box_left: f64,
    pub box_right: f64,
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
    /// Character offset of the segment in the paragraph text
    start: usize,
    bold: bool,
    italic: bool,
    underline: bool,
    color: String,
    font_size: f64,
    font_family: Option<String>,
}

enum AtomicUnit {
    Newline(usize),
    PageBreak(usize),
    Tab(usize),
    Spaces(Vec<LayoutSegment>),
    Word(Vec<LayoutSegment>),
}

/// Splits runs into words, spaces and breaks, remembering where each piece starts in the text
fn extract_segments(runs: &[RunInfo], default_color: &str, default_fs: f64) -> Vec<LayoutSegment> {
    let mut segments = Vec::new();
    let mut pos = 0usize;

    for r in runs {
        let color = css_color(&r.color, default_color);
        let seg = |kind: SegmentKind, start: usize| LayoutSegment {
            kind,
            start,
            bold: r.bold,
            italic: r.italic,
            underline: r.underline,
            color: color.clone(),
            font_size: r.font_size.unwrap_or(default_fs),
            font_family: r.font_family.clone(),
        };
        let mut text = String::new();
        let mut text_start = pos;
        let mut spaces = String::new();
        let mut spaces_start = pos;

        for ch in r.text.chars() {
            let is_break = ch == '\n' || ch == PAGE_BREAK || ch == '\t';
            if (is_break || ch.is_whitespace()) && !text.is_empty() {
                segments.push(seg(SegmentKind::Text(std::mem::take(&mut text)), text_start));
            }
            if (is_break || !ch.is_whitespace()) && !spaces.is_empty() {
                segments.push(seg(SegmentKind::Space(std::mem::take(&mut spaces)), spaces_start));
            }
            match ch {
                '\n' => segments.push(seg(SegmentKind::Newline, pos)),
                PAGE_BREAK => segments.push(seg(SegmentKind::PageBreak, pos)),
                '\t' => segments.push(seg(SegmentKind::Tab, pos)),
                c if c.is_whitespace() => {
                    if spaces.is_empty() {
                        spaces_start = pos;
                    }
                    spaces.push(c);
                }
                c => {
                    if text.is_empty() {
                        text_start = pos;
                    }
                    text.push(c);
                }
            }
            pos += 1;
        }
        if !text.is_empty() {
            segments.push(seg(SegmentKind::Text(text), text_start));
        }
        if !spaces.is_empty() {
            segments.push(seg(SegmentKind::Space(spaces), spaces_start));
        }
    }

    segments
}

fn group_into_units(segments: Vec<LayoutSegment>) -> Vec<AtomicUnit> {
    let mut units = Vec::new();
    let mut word: Vec<LayoutSegment> = Vec::new();
    let mut spaces: Vec<LayoutSegment> = Vec::new();

    for seg in segments {
        match seg.kind {
            SegmentKind::Newline | SegmentKind::PageBreak | SegmentKind::Tab => {
                if !word.is_empty() {
                    units.push(AtomicUnit::Word(std::mem::take(&mut word)));
                }
                if !spaces.is_empty() {
                    units.push(AtomicUnit::Spaces(std::mem::take(&mut spaces)));
                }
                units.push(match seg.kind {
                    SegmentKind::Newline => AtomicUnit::Newline(seg.start),
                    SegmentKind::PageBreak => AtomicUnit::PageBreak(seg.start),
                    _ => AtomicUnit::Tab(seg.start),
                });
            }
            SegmentKind::Space(_) => {
                if !word.is_empty() {
                    units.push(AtomicUnit::Word(std::mem::take(&mut word)));
                }
                spaces.push(seg);
            }
            SegmentKind::Text(_) => {
                if !spaces.is_empty() {
                    units.push(AtomicUnit::Spaces(std::mem::take(&mut spaces)));
                }
                word.push(seg);
            }
        }
    }
    if !word.is_empty() {
        units.push(AtomicUnit::Word(word));
    }
    if !spaces.is_empty() {
        units.push(AtomicUnit::Spaces(spaces));
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
            && last.start + last.text.chars().count() == seg.start
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
        x: 0.0,
        start: seg.start,
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
    let flush = |mut runs: Vec<TextRun>, line_w: f64, idx: usize, ends_paragraph_line: bool, page_break_after: bool, range: (usize, usize)| {
        let indent = line_indent(idx);
        let max_w = max_width(idx);
        let x = match p.align.as_str() {
            "center" => margin_left + indent + (max_w - line_w).max(0.0) / 2.0,
            "right" => margin_left + indent + (max_w - line_w).max(0.0),
            _ => margin_left + indent,
        };
        // Justified lines (except the last one) stretch their spaces to fill the line
        let space_count: usize = runs.iter().filter(|r| r.text != "\t").map(|r| r.text.matches(' ').count()).sum();
        let gap = max_w - line_w;
        let space_extra = if p.align == "both" && !ends_paragraph_line && space_count > 0 && gap > 0.0 {
            gap / space_count as f64
        } else {
            0.0
        };
        let mut cur = x;
        for r in runs.iter_mut() {
            r.x = cur;
            cur += r.width + if r.text == "\t" { 0.0 } else { space_extra * r.text.matches(' ').count() as f64 };
        }
        LayoutLine {
            text: runs.iter().map(|r| r.text.as_str()).collect(),
            x,
            width: line_w,
            max_width: max_w,
            is_last_line: ends_paragraph_line,
            page_break_after,
            start: range.0,
            end: range.1,
            space_extra,
            box_left: margin_left + indent,
            box_right: margin_left + indent + max_w,
            runs,
        }
    };

    let total_chars = p.text.chars().count();
    if p.text.trim().is_empty() && p.runs.is_empty() {
        return vec![flush(Vec::new(), 0.0, 0, true, false, (0, total_chars))];
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
    let total_chars = source_runs.iter().map(|r| r.text.chars().count()).sum::<usize>();

    let units = group_into_units(extract_segments(&source_runs, default_color, font_size));

    let mut lines = Vec::new();
    let mut runs: Vec<TextRun> = Vec::new();
    let mut line_w = 0.0;
    let mut line_start = 0usize;
    let mut pending_spaces: Vec<LayoutSegment> = Vec::new();
    // Spaces are only swallowed at automatic line wraps, not after manual breaks
    let mut just_wrapped = false;
    let mut ended_with_break = false;
    let mut idx = 0;

    for unit in units {
        ended_with_break = false;
        match unit {
            AtomicUnit::Newline(at) | AtomicUnit::PageBreak(at) => {
                let page_break = matches!(unit, AtomicUnit::PageBreak(_));
                // Spaces before a manual break stay on the line
                for sp in pending_spaces.drain(..) {
                    if let SegmentKind::Space(ref s) = sp.kind {
                        if !runs.is_empty() || !just_wrapped {
                            let w = seg_width(m, &sp, s, family_css);
                            append_segment_to_runs(&mut runs, s, &sp, w);
                            line_w += w;
                        }
                    }
                }
                lines.push(flush(std::mem::take(&mut runs), line_w, idx, true, page_break, (line_start, at)));
                line_w = 0.0;
                line_start = at + 1;
                just_wrapped = false;
                ended_with_break = true;
                idx += 1;
            }
            AtomicUnit::Tab(at) => {
                // Default tab stops every 0.5in, measured from the left margin
                const DEFAULT_TAB: f64 = 48.0;
                for sp in pending_spaces.drain(..) {
                    if let SegmentKind::Space(ref s) = sp.kind {
                        if !runs.is_empty() || !just_wrapped {
                            let w = seg_width(m, &sp, s, family_css);
                            append_segment_to_runs(&mut runs, s, &sp, w);
                            line_w += w;
                        }
                    }
                }
                let advance = |indent: f64, w: f64| {
                    let pos = indent + w;
                    (((pos / DEFAULT_TAB).floor() + 1.0) * DEFAULT_TAB - pos).max(1.0)
                };
                let mut tab_w = advance(line_indent(idx), line_w);
                if line_w + tab_w > max_width(idx) && !runs.is_empty() {
                    lines.push(flush(std::mem::take(&mut runs), line_w, idx, false, false, (line_start, at)));
                    line_w = 0.0;
                    line_start = at;
                    idx += 1;
                    tab_w = advance(line_indent(idx), 0.0);
                }
                runs.push(TextRun {
                    text: "\t".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                    color: default_color.to_string(),
                    font_size: Some(font_size),
                    font_family: None,
                    width: tab_w,
                    x: 0.0,
                    start: at,
                });
                line_w += tab_w;
                just_wrapped = false;
            }
            AtomicUnit::Spaces(spaces) => pending_spaces = spaces,
            AtomicUnit::Word(word_segs) => {
                let word_start = word_segs.first().map_or(line_start, |s| s.start);
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
                let keep_spaces = !runs.is_empty() || !just_wrapped;
                let spaces: Vec<(String, LayoutSegment, f64)> = if keep_spaces {
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
                } else {
                    Vec::new()
                };
                pending_spaces.clear();
                let spaces_w: f64 = spaces.iter().map(|(_, _, w)| w).sum();

                if line_w + spaces_w + word_w > max_width(idx) && !runs.is_empty() {
                    lines.push(flush(std::mem::take(&mut runs), line_w, idx, false, false, (line_start, word_start)));
                    line_w = 0.0;
                    line_start = word_start;
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
                just_wrapped = false;
            }
        }
        if let Some(last) = lines.last() {
            if !last.is_last_line && runs.is_empty() {
                just_wrapped = true;
            }
        }
    }

    // Trailing spaces stay on the last line (they are visible to the caret)
    for sp in pending_spaces.drain(..) {
        if let SegmentKind::Space(ref s) = sp.kind {
            if !runs.is_empty() || !just_wrapped {
                let w = seg_width(m, &sp, s, family_css);
                append_segment_to_runs(&mut runs, s, &sp, w);
                line_w += w;
            }
        }
    }
    if !runs.is_empty() || ended_with_break || lines.is_empty() {
        // A paragraph ending in a manual break still shows an (empty) last line
        lines.push(flush(runs, line_w, idx, true, false, (line_start, total_chars)));
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
                    // Body lines only: not list labels or page decorations ("Página 1 de 2")
                    RenderCommand::Text { paragraph_index, y, height, line: Some(_), .. }
                        if *paragraph_index == paragraph =>
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

    fn picture(behind: bool) -> ImageRef {
        ImageRef {
            rel_id: "rId7".into(),
            width: 300.0,
            height: 200.0,
            anchored: true,
            behind_text: behind,
            h_relative: "page".into(),
            h_offset: 10.0,
            v_relative: "paragraph".into(),
            v_offset: -20.0,
            ..Default::default()
        }
    }

    fn images_per_page(l: &DocumentLayout) -> Vec<Vec<(f64, f64, bool, usize)>> {
        l.pages
            .iter()
            .map(|page| {
                page.items
                    .iter()
                    .enumerate()
                    .filter_map(|(i, item)| match item {
                        RenderCommand::Image { x, y, is_background, .. } => Some((*x, *y, *is_background, i)),
                        _ => None,
                    })
                    .collect()
            })
            .collect()
    }

    #[test]
    fn test_anchored_picture_only_on_its_page() {
        // The picture is anchored in paragraph 1, which sits on page 2 of 3
        let mut anchor = para(1, "con imagen");
        anchor.images = vec![picture(true)];
        let elements: Vec<DocumentElement> = vec![lines(0, 60), anchor, lines(2, 60)]
            .into_iter()
            .map(DocumentElement::Paragraph)
            .collect();
        let images = HashMap::from([("rId7".to_string(), "data:image/png;base64,AAAA".to_string())]);
        let l = LayoutEngine::new().with_images(images).compute_layout(
            &elements, "FFFFFF", &PageSetup::default(), &HeaderFooterInfo::default(), None, None, 0.0,
        );
        let per_page = images_per_page(&l);
        assert_eq!(per_page.iter().map(|p| p.len()).collect::<Vec<_>>(), vec![0, 1, 0], "one page only");

        // Positioned from the page edge and the paragraph's top, and drawn before the text
        let (x, y, behind, index) = per_page[1][0];
        let para_top = l.pages[1]
            .items
            .iter()
            .find_map(|item| match item {
                RenderCommand::Text { paragraph_index: 1, line: Some(range), .. } => Some(range.top),
                _ => None,
            })
            .unwrap();
        assert_eq!(x, 10.0, "10px from the page edge");
        assert!((y - (para_top - 20.0)).abs() < 1e-9, "20px above its paragraph");
        assert!(behind);
        assert_eq!(index, 0, "behind the text: first thing drawn on the page");
    }

    #[test]
    fn test_header_pictures_repeat_on_every_page_at_their_position() {
        use crate::docx_parser::PlacedImage;
        let header = HeaderFooterInfo {
            has_header: true,
            header_images: vec![PlacedImage { image: picture(false), data_url: "data:image/png;base64,AAAA".into() }],
            ..Default::default()
        };
        let elements: Vec<DocumentElement> = vec![lines(0, 120)].into_iter().map(DocumentElement::Paragraph).collect();
        let l = LayoutEngine::new().compute_layout(&elements, "FFFFFF", &PageSetup::default(), &header, None, None, 0.0);
        let per_page = images_per_page(&l);
        assert!(per_page.len() >= 3);
        let setup = PageSetup::default();
        for page in per_page {
            assert_eq!(page.len(), 1);
            let (x, y, behind, _) = page[0];
            assert_eq!((x, y, behind), (10.0, setup.header_margin - 20.0, false));
        }
    }
}

