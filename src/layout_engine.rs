use crate::docx_parser::{DocumentElement, HeaderFooterInfo, PageSetup, ParagraphInfo, RunInfo};
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

    /// Computes multi-page layout from document elements using document page setup, headers/footers, and media
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
        let page_w = page_setup.width.max(320.0);
        let page_h = page_setup.height.max(400.0);
        let margin_left = page_setup.margin_left.max(20.0);
        let margin_right = page_setup.margin_right.max(20.0);
        let margin_top = page_setup.margin_top.max(24.0);
        let margin_bottom = page_setup.margin_bottom.max(24.0);
        let printable_w = (page_w - margin_left - margin_right).max(100.0);

        let mut pages: Vec<PageLayout> = Vec::new();
        let mut current_page_items: Vec<RenderCommand> = Vec::new();
        let mut current_page_num = 1;

        // If header logo is present, start content below the logo area on page 1
        let mut cursor_y = if header_footer.header_image_data_url.is_some() {
            margin_top.max(120.0)
        } else {
            margin_top
        };

        let clean_bg = if bg_color.starts_with('#') {
            bg_color.to_string()
        } else {
            format!("#{}", bg_color)
        };

        for element in elements {
            match element {
                DocumentElement::Paragraph(p) => {
                    let (font_size, line_height, font_weight, font_family, default_color) =
                        get_paragraph_typography(p);

                    let p_color = if !p.color.is_empty() {
                        if p.color.starts_with('#') {
                            p.color.clone()
                        } else {
                            format!("#{}", p.color)
                        }
                    } else {
                        default_color.to_string()
                    };

                    let p_style = if p.italic { "italic" } else { "normal" };

                    // 1. Spacing before paragraph
                    if p.space_before > 0.0 {
                        cursor_y += p.space_before;
                    }

                    // 2. Top border (if present)
                    if let Some(ref top_bdr) = p.borders.top {
                        let top_color = if top_bdr.color.starts_with('#') {
                            top_bdr.color.clone()
                        } else if !top_bdr.color.is_empty() {
                            format!("#{}", top_bdr.color)
                        } else {
                            "#CBD5E1".to_string()
                        };
                        current_page_items.push(RenderCommand::Line {
                            x1: margin_left,
                            y1: cursor_y,
                            x2: page_w - margin_right,
                            y2: cursor_y,
                            color: top_color,
                            line_width: top_bdr.sz_px.max(0.75),
                        });
                        cursor_y += top_bdr.space.max(3.0) + top_bdr.sz_px;
                    }

                    let p_start_y = cursor_y;

                    // 3. Layout paragraph lines respecting runs, bold/italic/underline, indents, tabs, fonts
                    let lines = layout_paragraph_lines(p, font_size, &p_color, margin_left, printable_w);

                    // Check if new page is needed before starting paragraph
                    if cursor_y + line_height > page_h - margin_bottom {
                        if let Some(wm_text) = watermark {
                            add_watermark_command(&mut current_page_items, wm_text, watermark_opacity, page_w, page_h);
                        }
                        pages.push(PageLayout {
                            page_number: current_page_num,
                            width: page_w,
                            height: page_h,
                            margin_left,
                            margin_right,
                            printable_width: printable_w,
                            bg_color: clean_bg.clone(),
                            items: current_page_items,
                        });

                        current_page_items = Vec::new();
                        current_page_num += 1;
                        cursor_y = margin_top;
                    }

                    for (line_idx, line) in lines.iter().enumerate() {
                        // Check for page overflow within paragraph lines
                        if cursor_y + line_height > page_h - margin_bottom {
                            if let Some(wm_text) = watermark {
                                add_watermark_command(&mut current_page_items, wm_text, watermark_opacity, page_w, page_h);
                            }
                            pages.push(PageLayout {
                                page_number: current_page_num,
                                width: page_w,
                                height: page_h,
                                margin_left,
                                margin_right,
                                printable_width: printable_w,
                                bg_color: clean_bg.clone(),
                                items: current_page_items,
                            });
                            current_page_items = Vec::new();
                            current_page_num += 1;
                            cursor_y = margin_top;
                        }

                        if !line.text.is_empty() || !line.runs.is_empty() {
                            current_page_items.push(RenderCommand::Text {
                                text: line.text.clone(),
                                x: line.x,
                                y: cursor_y + font_size * 0.85,
                                width: line.width,
                                height: line_height,
                                font_size,
                                font_family: font_family.clone(),
                                font_weight: font_weight.to_string(),
                                font_style: p_style.to_string(),
                                color: p_color.clone(),
                                align: p.align.clone(),
                                paragraph_index: p.index,
                                line_index: line_idx,
                                is_last_line: line.is_last_line,
                                max_width: line.max_width,
                                runs: line.runs.clone(),
                            });
                        }

                        // List number / bullet, drawn in the hanging indent of the first line
                        if line_idx == 0 {
                            if let (Some(label), Some(geo)) = (p.list_label.as_ref(), list_label_geometry(p, font_size)) {
                                let mut run = label.clone();
                                run.font_size = Some(geo.font_size);
                                current_page_items.push(RenderCommand::Text {
                                    text: label.text.clone(),
                                    x: margin_left + geo.label_offset,
                                    y: cursor_y + font_size * 0.85,
                                    width: geo.label_width,
                                    height: line_height,
                                    font_size: geo.font_size,
                                    font_family: font_family.clone(),
                                    font_weight: if label.bold { "700" } else { "400" }.to_string(),
                                    font_style: if label.italic { "italic" } else { "normal" }.to_string(),
                                    color: p_color.clone(),
                                    align: "left".to_string(),
                                    paragraph_index: p.index,
                                    line_index: 0,
                                    is_last_line: true,
                                    max_width: geo.label_width,
                                    runs: vec![TextRun {
                                        text: run.text.clone(),
                                        bold: run.bold,
                                        italic: run.italic,
                                        underline: run.underline,
                                        color: run.color.clone(),
                                        font_size: run.font_size,
                                        font_family: run.font_family.clone(),
                                        width: geo.label_width,
                                    }],
                                });
                            }
                        }

                        cursor_y += line_height;
                    }

                    let p_end_y = cursor_y;

                    // 4. Left border (vertical accent bar, e.g. on headings)
                    if let Some(ref left_bdr) = p.borders.left {
                        let left_color = if left_bdr.color.starts_with('#') {
                            left_bdr.color.clone()
                        } else if !left_bdr.color.is_empty() && left_bdr.color.to_lowercase() != "auto" {
                            format!("#{}", left_bdr.color)
                        } else {
                            "#334155".to_string()
                        };
                        let bar_x = (margin_left + p.indent_left.max(0.0) - 7.0).max(10.0);
                        current_page_items.push(RenderCommand::Line {
                            x1: bar_x,
                            y1: p_start_y + 1.0,
                            x2: bar_x,
                            y2: p_end_y.max(p_start_y + line_height) - 1.0,
                            color: left_color,
                            line_width: left_bdr.sz_px.max(2.5),
                        });
                    }

                    // 5. Bottom border (e.g. title divider or metadata bottom border)
                    if let Some(ref btm_bdr) = p.borders.bottom {
                        let btm_color = if btm_bdr.color.starts_with('#') {
                            btm_bdr.color.clone()
                        } else if !btm_bdr.color.is_empty() && btm_bdr.color.to_lowercase() != "auto" {
                            format!("#{}", btm_bdr.color)
                        } else {
                            "#CBD5E1".to_string()
                        };
                        let line_y = if lines.is_empty() || (lines.len() == 1 && lines[0].text.is_empty()) {
                            cursor_y
                        } else {
                            cursor_y + 3.0
                        };
                        current_page_items.push(RenderCommand::Line {
                            x1: margin_left,
                            y1: line_y,
                            x2: page_w - margin_right,
                            y2: line_y,
                            color: btm_color,
                            line_width: btm_bdr.sz_px.max(1.0),
                        });
                        cursor_y += btm_bdr.space.max(3.0) + btm_bdr.sz_px;
                    }

                    // 6. Spacing after paragraph
                    let spacing_after = if p.space_after > 0.0 {
                        p.space_after
                    } else if p.is_heading {
                        12.0
                    } else {
                        6.0
                    };
                    cursor_y += spacing_after;
                }

                DocumentElement::Table(tbl) => {
                    let num_cols = if !tbl.grid_cols.is_empty() {
                        tbl.grid_cols.len()
                    } else {
                        tbl.rows.first().map(|r| r.len()).unwrap_or(1).max(1)
                    };

                    let total_grid_dxa = if !tbl.grid_cols.is_empty() {
                        tbl.grid_cols.iter().sum::<f64>().max(1.0)
                    } else {
                        num_cols as f64
                    };

                    let col_widths: Vec<f64> = if !tbl.grid_cols.is_empty() {
                        tbl.grid_cols.iter().map(|w| (w / total_grid_dxa) * printable_w).collect()
                    } else {
                        vec![printable_w / num_cols as f64; num_cols]
                    };

                    cursor_y += 8.0;

                    for (row_idx, row_strings) in tbl.rows.iter().enumerate() {
                        let rich_row = tbl.rich_rows.get(row_idx);
                        let is_header = rich_row.map(|r| r.is_header).unwrap_or(row_idx == 0 && tbl.header_row);

                        // Pre-calculate wrapped lines for all cells in this row to determine dynamic row height
                        let mut row_cell_lines: Vec<Vec<String>> = Vec::new();
                        for (col_idx, cell_text) in row_strings.iter().enumerate() {
                            let col_w = col_widths.get(col_idx).copied().unwrap_or(printable_w / num_cols as f64);
                            let inner_cell_width = (col_w - 16.0).max(10.0);
                            let cell_font_size = rich_row
                                .and_then(|r| r.cells.get(col_idx))
                                .map(|c| c.font_size * (96.0 / 72.0))
                                .unwrap_or(if is_header { 13.33 } else { 12.0 });

                            let mut cell_lines = Vec::new();
                            for part in cell_text.split('\n') {
                                let wrapped = wrap_text(part, inner_cell_width, cell_font_size);
                                for w in wrapped {
                                    cell_lines.push(w);
                                }
                            }
                            if cell_lines.is_empty() {
                                cell_lines.push(String::new());
                            }
                            row_cell_lines.push(cell_lines);
                        }

                        let max_lines = row_cell_lines
                            .iter()
                            .map(|lines| lines.len())
                            .max()
                            .unwrap_or(1)
                            .max(1);

                        let base_fs = rich_row.and_then(|r| r.cells.first()).map(|c| c.font_size * (96.0 / 72.0)).unwrap_or(13.33);
                        let line_height = base_fs * 1.35;
                        let row_height = ((max_lines as f64) * line_height + 14.0).max(28.0);

                        // Check if row fits on current page
                        if cursor_y + row_height > page_h - margin_bottom {
                            if let Some(wm_text) = watermark {
                                add_watermark_command(&mut current_page_items, wm_text, watermark_opacity, page_w, page_h);
                            }
                            pages.push(PageLayout {
                                page_number: current_page_num,
                                width: page_w,
                                height: page_h,
                                margin_left,
                                margin_right,
                                printable_width: printable_w,
                                bg_color: clean_bg.clone(),
                                items: current_page_items,
                            });
                            current_page_items = Vec::new();
                            current_page_num += 1;
                            cursor_y = margin_top;
                        }

                        let mut current_col_x = margin_left;
                        for (col_idx, cell_text) in row_strings.iter().enumerate() {
                            let col_w = col_widths.get(col_idx).copied().unwrap_or(printable_w / num_cols as f64);
                            let lines = row_cell_lines.get(col_idx).cloned().unwrap_or_else(|| vec![cell_text.clone()]);
                            let rich_cell = rich_row.and_then(|r| r.cells.get(col_idx));

                            let bg_color = rich_cell.and_then(|c| c.bg_color.clone());
                            let color = rich_cell
                                .map(|c| c.color.clone())
                                .unwrap_or_else(|| if is_header { "FAF7F0".to_string() } else { "1B1F1E".to_string() });

                            let cell_font_size = rich_cell.map(|c| c.font_size * (96.0 / 72.0)).unwrap_or(if is_header { 13.33 } else { 12.0 });
                            let cell_font_weight = if rich_cell.map(|c| c.bold).unwrap_or(is_header) { "700" } else { "400" };
                            let cell_font_family = rich_cell
                                .map(|c| c.font_family.clone())
                                .unwrap_or_else(|| "Calibri, Inter, sans-serif".to_string());
                            let cell_align = rich_cell.map(|c| c.align.clone()).unwrap_or_else(|| "left".to_string());
                            let border_color = rich_cell
                                .map(|c| c.border_color.clone())
                                .unwrap_or_else(|| "DDD5C2".to_string());

                            current_page_items.push(RenderCommand::TableCell {
                                table_index: tbl.index,
                                row: row_idx,
                                col: col_idx,
                                x: current_col_x,
                                y: cursor_y,
                                width: col_w,
                                height: row_height,
                                text: cell_text.clone(),
                                lines,
                                is_header,
                                bg_color,
                                color,
                                font_size: cell_font_size,
                                font_family: cell_font_family,
                                font_weight: cell_font_weight.to_string(),
                                align: cell_align,
                                border_color,
                            });

                            current_col_x += col_w;
                        }

                        cursor_y += row_height;
                    }

                    cursor_y += 16.0;
                }
            }
        }

        // Push final page
        if !current_page_items.is_empty() || pages.is_empty() {
            if let Some(wm_text) = watermark {
                add_watermark_command(&mut current_page_items, wm_text, watermark_opacity, page_w, page_h);
            }
            pages.push(PageLayout {
                page_number: current_page_num,
                width: page_w,
                height: page_h,
                margin_left,
                margin_right,
                printable_width: printable_w,
                bg_color: clean_bg,
                items: current_page_items,
            });
        }

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

#[derive(Debug, Clone)]
struct LayoutLine {
    pub text: String,
    pub x: f64,
    pub width: f64,
    pub max_width: f64,
    pub is_last_line: bool,
    pub runs: Vec<TextRun>,
}

#[derive(Clone, Debug)]
enum SegmentKind {
    Newline,
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
            if ch == '\n' {
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
                    kind: SegmentKind::Newline,
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
            SegmentKind::Newline => {
                if !current_word.is_empty() {
                    units.push(AtomicUnit::Word(std::mem::take(&mut current_word)));
                }
                if !current_spaces.is_empty() {
                    units.push(AtomicUnit::Spaces(std::mem::take(&mut current_spaces)));
                }
                units.push(AtomicUnit::Newline);
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

fn layout_paragraph_lines(
    p: &ParagraphInfo,
    font_size: f64,
    default_color: &str,
    margin_left: f64,
    printable_width: f64,
) -> Vec<LayoutLine> {
    if p.text.trim().is_empty() && p.runs.is_empty() {
        let max_w = (printable_width - p.indent_left.max(0.0) - p.indent_right.max(0.0)).max(60.0);
        return vec![LayoutLine {
            text: String::new(),
            x: margin_left + p.indent_left.max(0.0),
            width: 0.0,
            max_width: max_w,
            is_last_line: true,
            runs: Vec::new(),
        }];
    }

    // Build raw runs list with font sizes scaled to px
    let source_runs: Vec<RunInfo> = if !p.runs.is_empty() {
        p.runs.iter().map(|r| {
            let mut r_clone = r.clone();
            if let Some(sz_pt) = r_clone.font_size {
                r_clone.font_size = Some(sz_pt * (96.0 / 72.0));
            }
            r_clone
        }).collect()
    } else {
        vec![RunInfo {
            text: p.text.clone(),
            bold: p.bold,
            italic: p.italic,
            underline: false,
            color: if !p.color.is_empty() { p.color.clone() } else { default_color.to_string() },
            font_size: p.font_size.map(|sz_pt| sz_pt * (96.0 / 72.0)),
            font_family: p.font_family.clone(),
        }]
    };

    let units = group_into_units(extract_segments(&source_runs, default_color, font_size));

    let mut result_lines = Vec::new();
    let mut current_line_runs: Vec<TextRun> = Vec::new();
    let mut current_line_w = 0.0;
    let mut pending_spaces: Vec<LayoutSegment> = Vec::new();
    let mut line_idx = 0;

    let first_line_indent = list_label_geometry(p, font_size)
        .map(|g| g.text_offset)
        .unwrap_or_else(|| (p.indent_left + p.indent_first_line).max(0.0));
    let get_line_indent = |idx: usize| -> f64 {
        if idx == 0 {
            first_line_indent
        } else {
            p.indent_left.max(0.0)
        }
    };

    let get_max_width = |idx: usize| -> f64 {
        let ind = get_line_indent(idx);
        (printable_width - ind - p.indent_right.max(0.0)).max(60.0)
    };

    let flush_line = |runs: Vec<TextRun>, line_w: f64, idx: usize, is_last: bool| -> LayoutLine {
        let line_indent = get_line_indent(idx);
        let max_w = get_max_width(idx);
        let full_text = runs.iter().map(|r| r.text.as_str()).collect::<String>();
        let x = match p.align.as_str() {
            "center" => margin_left + line_indent + (max_w - line_w).max(0.0) / 2.0,
            "right" => margin_left + line_indent + (max_w - line_w).max(0.0),
            _ => margin_left + line_indent,
        };
        LayoutLine {
            text: full_text,
            x,
            width: line_w,
            max_width: max_w,
            is_last_line: is_last,
            runs,
        }
    };

    for unit in units {
        match unit {
            AtomicUnit::Newline => {
                result_lines.push(flush_line(current_line_runs, current_line_w, line_idx, true));
                current_line_runs = Vec::new();
                current_line_w = 0.0;
                pending_spaces.clear();
                line_idx += 1;
            }
            AtomicUnit::Tab => {
                let tab_interval = 48.0;
                let next_stop = ((current_line_w / tab_interval).floor() + 1.0) * tab_interval;
                let tab_advance = (next_stop - current_line_w).max(16.0);
                let max_w = get_max_width(line_idx);

                if current_line_w + tab_advance > max_w && !current_line_runs.is_empty() {
                    result_lines.push(flush_line(current_line_runs, current_line_w, line_idx, false));
                    current_line_runs = Vec::new();
                    current_line_w = 0.0;
                    pending_spaces.clear();
                    line_idx += 1;
                }

                current_line_runs.push(TextRun {
                    text: "\t".to_string(),
                    bold: false,
                    italic: false,
                    underline: false,
                    color: default_color.to_string(),
                    font_size: Some(font_size),
                    font_family: None,
                    width: tab_advance,
                });
                current_line_w += tab_advance;
                pending_spaces.clear();
            }
            AtomicUnit::Spaces(spaces) => {
                pending_spaces = spaces;
            }
            AtomicUnit::Word(word_segs) => {
                let mut word_w = 0.0;
                for seg in &word_segs {
                    if let SegmentKind::Text(ref t) = seg.kind {
                        word_w += estimate_text_width(t, seg.font_size) * (if seg.bold { 1.05 } else { 1.0 });
                    }
                }

                let mut spaces_w = 0.0;
                if !current_line_runs.is_empty() {
                    for sp in &pending_spaces {
                        if let SegmentKind::Space(ref s) = sp.kind {
                            spaces_w += estimate_text_width(s, sp.font_size);
                        }
                    }
                }

                let max_w = get_max_width(line_idx);

                if current_line_w + spaces_w + word_w > max_w && !current_line_runs.is_empty() {
                    result_lines.push(flush_line(current_line_runs, current_line_w, line_idx, false));
                    current_line_runs = Vec::new();
                    current_line_w = 0.0;
                    pending_spaces.clear();
                    line_idx += 1;

                    // Add word pieces to new line
                    for seg in word_segs {
                        if let SegmentKind::Text(ref t) = seg.kind {
                            let piece_w = estimate_text_width(t, seg.font_size) * (if seg.bold { 1.05 } else { 1.0 });
                            append_segment_to_runs(&mut current_line_runs, t, &seg, piece_w);
                            current_line_w += piece_w;
                        }
                    }
                } else {
                    // Commit pending spaces if on same line
                    if !current_line_runs.is_empty() && !pending_spaces.is_empty() {
                        for sp in pending_spaces.drain(..) {
                            if let SegmentKind::Space(ref s) = sp.kind {
                                let sp_w = estimate_text_width(s, sp.font_size);
                                append_segment_to_runs(&mut current_line_runs, s, &sp, sp_w);
                                current_line_w += sp_w;
                            }
                        }
                    }

                    // Add word pieces to current line
                    for seg in word_segs {
                        if let SegmentKind::Text(ref t) = seg.kind {
                            let piece_w = estimate_text_width(t, seg.font_size) * (if seg.bold { 1.05 } else { 1.0 });
                            append_segment_to_runs(&mut current_line_runs, t, &seg, piece_w);
                            current_line_w += piece_w;
                        }
                    }
                }
            }
        }
    }

    if !current_line_runs.is_empty() {
        result_lines.push(flush_line(current_line_runs, current_line_w, line_idx, true));
    }

    if result_lines.is_empty() {
        let max_w = get_max_width(0);
        result_lines.push(LayoutLine {
            text: String::new(),
            x: margin_left + p.indent_left.max(0.0),
            width: 0.0,
            max_width: max_w,
            is_last_line: true,
            runs: Vec::new(),
        });
    }

    result_lines
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
fn list_label_geometry(p: &ParagraphInfo, paragraph_font_size: f64) -> Option<ListLabelGeometry> {
    let label = p.list_label.as_ref()?;
    let font_size = label.font_size.map(|pt| pt * (96.0 / 72.0)).unwrap_or(paragraph_font_size);
    let label_offset = (p.indent_left + p.indent_first_line).max(0.0);
    let label_width = estimate_text_width(&label.text, font_size);
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

/// Simple greedy word wrapping calculation based on approximate glyph widths
fn wrap_text(text: &str, max_width: f64, font_size: f64) -> Vec<String> {
    if text.trim().is_empty() {
        return vec![String::new()];
    }

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

        let line_w = estimate_text_width(&test_line, font_size);
        if line_w > max_width && !current_line.is_empty() {
            lines.push(current_line);
            current_line = word.to_string();
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
