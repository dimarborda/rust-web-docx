// Draws a region of a sheet on a 2D canvas: cell backgrounds, gridlines, merged cells, text
// with Excel's alignment and overflow rules, and borders.
import { ptToPx } from './grid.js';

/**
 * CSS font stack for a workbook font: the font itself when installed, otherwise a
 * metric-compatible substitute (Carlito ≈ Calibri, Caladea ≈ Cambria, Arimo ≈ Arial,
 * Tinos ≈ Times New Roman, Cousine ≈ Courier New/Consolas)
 */
export function fontStack(family) {
  if (!family) return '"Carlito", "Calibri", "Arimo", "Arial", sans-serif';
  const lower = family.toLowerCase().trim();
  if (lower.includes('calibri')) return '"Calibri", "Carlito", "Segoe UI", Roboto, sans-serif';
  if (lower.includes('aptos')) return '"Aptos", "Calibri", "Carlito", "Segoe UI", sans-serif';
  if (lower.includes('consolas') || lower.includes('courier') || lower.includes('mono')) return '"Consolas", "Cousine", monospace';
  if (lower.includes('cambria')) return '"Cambria", "Caladea", Georgia, serif';
  if (lower.includes('times')) return '"Times New Roman", "Tinos", Georgia, serif';
  if (lower.includes('arial') || lower.includes('helvetica')) return '"Arial", "Arimo", Helvetica, sans-serif';
  return `"${family}", "Carlito", "Arimo", sans-serif`;
}

/** Canvas font of a cell font at `zoom` */
export function canvasFont(font, zoom = 1) {
  const px = ptToPx(font.size || 11) * zoom;
  return `${font.italic ? 'italic ' : ''}${font.bold ? '700 ' : ''}${px.toFixed(2)}px ${fontStack(font.name)}`;
}

const GRIDLINE = '#d9d9d9';
const PAD = 3; // horizontal text padding, px at zoom 1

/** Line width and dash of an Excel border style, at zoom 1 */
function borderLook(style) {
  switch (style) {
    case 'medium': return { width: 2, dash: [] };
    case 'thick': return { width: 3, dash: [] };
    case 'double': return { width: 1, dash: [], double: true };
    case 'hair': return { width: 1, dash: [1, 1] };
    case 'dotted': return { width: 1, dash: [1, 2] };
    case 'dashed': return { width: 1, dash: [3, 1] };
    case 'mediumDashed': return { width: 2, dash: [6, 2] };
    case 'dashDot': return { width: 1, dash: [6, 2, 2, 2] };
    case 'mediumDashDot': return { width: 2, dash: [6, 2, 2, 2] };
    case 'dashDotDot': return { width: 1, dash: [6, 2, 2, 2, 2, 2] };
    case 'mediumDashDotDot': return { width: 2, dash: [6, 2, 2, 2, 2, 2] };
    case 'slantDashDot': return { width: 2, dash: [6, 2, 2, 2] };
    default: return { width: 1, dash: [] };
  }
}

function strokeLine(ctx, x1, y1, x2, y2, side, zoom) {
  const [style, color] = side;
  const look = borderLook(style);
  const width = Math.max(1, Math.round(look.width * zoom));
  ctx.strokeStyle = color || '#000000';
  ctx.lineWidth = width;
  ctx.setLineDash(look.dash.map(d => d * zoom));
  // Odd widths sit on half pixels to stay crisp
  const half = width % 2 ? 0.5 : 0;
  const horizontal = y1 === y2;
  ctx.beginPath();
  if (look.double) {
    const gap = Math.max(1, Math.round(zoom));
    for (const d of [-gap, gap]) {
      if (horizontal) {
        ctx.moveTo(x1, Math.round(y1 + d) + half);
        ctx.lineTo(x2, Math.round(y2 + d) + half);
      } else {
        ctx.moveTo(Math.round(x1 + d) + half, y1);
        ctx.lineTo(Math.round(x2 + d) + half, y2);
      }
    }
  } else if (horizontal) {
    ctx.moveTo(x1, Math.round(y1) + half);
    ctx.lineTo(x2, Math.round(y2) + half);
  } else {
    ctx.moveTo(Math.round(x1) + half, y1);
    ctx.lineTo(Math.round(x2) + half, y2);
  }
  ctx.stroke();
}

/** Splits `text` into lines no wider than `width` (words longer than a line are cut) */
export function wrapLines(ctx, text, width) {
  const lines = [];
  for (const paragraph of String(text).split('\n')) {
    let line = '';
    for (const word of paragraph.split(/(\s+)/)) {
      const candidate = line + word;
      if (!line || ctx.measureText(candidate).width <= width) {
        line = candidate;
        continue;
      }
      lines.push(line.trimEnd());
      line = word.trimStart();
      while (line && ctx.measureText(line).width > width) {
        let cut = line.length - 1;
        while (cut > 1 && ctx.measureText(line.slice(0, cut)).width > width) cut--;
        lines.push(line.slice(0, cut));
        line = line.slice(cut);
      }
    }
    lines.push(line);
  }
  return lines;
}

/**
 * Draws the cells of rows `r0..r1` and columns `c0..c1` with their top-left corner at
 * (`x0`, `y0`) on the canvas, clipped to `clip` ({x, y, w, h}).
 * @param {CanvasRenderingContext2D} ctx
 * @param {object} sheet `{cells: Map<key, cell>, rows: Axis, cols: Axis, merges, mergeAt, covered, gridlines}`
 * @param {object} styles the workbook's `styleView()`
 */
export function drawRegion(ctx, sheet, styles, { r0, r1, c0, c1, x0, y0, clip, zoom }) {
  const { rows, cols } = sheet;
  const xOf = c => x0 + (cols.offset(c) - cols.offset(c0)) * zoom;
  const yOf = r => y0 + (rows.offset(r) - rows.offset(r0)) * zoom;
  const xs = [];
  for (let c = c0; c <= c1 + 1; c++) xs.push(xOf(c));
  const ys = [];
  for (let r = r0; r <= r1 + 1; r++) ys.push(yOf(r));
  const X = c => xs[c - c0];
  const Y = r => ys[r - r0];
  const xf = s => styles.xfs[s] || styles.xfs[0];

  ctx.save();
  ctx.beginPath();
  ctx.rect(clip.x, clip.y, clip.w, clip.h);
  ctx.clip();
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(clip.x, clip.y, clip.w, clip.h);

  // Gridlines
  if (sheet.gridlines) {
    ctx.strokeStyle = GRIDLINE;
    ctx.lineWidth = 1;
    ctx.setLineDash([]);
    ctx.beginPath();
    for (let c = c0; c <= c1 + 1; c++) {
      if (c <= c1 && cols.size(c) === 0) continue;
      const x = Math.round(X(c)) - 0.5;
      ctx.moveTo(x, Y(r0));
      ctx.lineTo(x, Y(r1 + 1));
    }
    for (let r = r0; r <= r1 + 1; r++) {
      if (r <= r1 && rows.size(r) === 0) continue;
      const y = Math.round(Y(r)) - 0.5;
      ctx.moveTo(X(c0), y);
      ctx.lineTo(X(c1 + 1), y);
    }
    ctx.stroke();
  }

  // Merged ranges touching the region, clipped to it
  const merges = sheet.merges.filter(m => m.last.row >= r0 && m.first.row <= r1 && m.last.col >= c0 && m.first.col <= c1);
  const rectOf = (fr, fc, lr, lc) => {
    const x = fc >= c0 ? X(fc) : xOf(fc);
    const y = fr >= r0 ? Y(fr) : yOf(fr);
    const right = lc + 1 <= c1 + 1 ? X(lc + 1) : xOf(lc + 1);
    const bottom = lr + 1 <= r1 + 1 ? Y(lr + 1) : yOf(lr + 1);
    return { x, y, w: right - x, h: bottom - y };
  };

  // Backgrounds: filled cells hide the gridlines, merged ranges hide their inner lines
  for (const m of merges) {
    const cell = sheet.cells.get(key(m.first.row, m.first.col));
    const fill = cell ? xf(cell.style).fill : null;
    const r = rectOf(m.first.row, m.first.col, m.last.row, m.last.col);
    ctx.fillStyle = fill || '#ffffff';
    ctx.fillRect(r.x, r.y, r.w - 1, r.h - 1);
  }
  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      const cell = sheet.cells.get(key(r, c));
      if (!cell || sheet.covered.has(key(r, c))) continue;
      const fill = xf(cell.style).fill;
      if (fill) {
        ctx.fillStyle = fill;
        ctx.fillRect(X(c) - 1, Y(r) - 1, X(c + 1) - X(c) + 1, Y(r + 1) - Y(r) + 1);
      }
    }
  }

  // Text
  ctx.textBaseline = 'alphabetic';
  for (let r = r0; r <= r1; r++) {
    if (rows.size(r) === 0) continue;
    for (let c = c0; c <= c1; c++) {
      const k = key(r, c);
      const cell = sheet.cells.get(k);
      if (!cell || !cell.text || sheet.covered.has(k)) continue;
      if (sheet.mergeAt.has(k)) continue;
      const box = { x: X(c), y: Y(r), w: X(c + 1) - X(c), h: Y(r + 1) - Y(r) };
      if (box.w <= 0 || box.h <= 0) continue;
      drawCellText(ctx, sheet, xf(cell.style), cell, box, { r, c, c1, xOf, zoom, merged: false });
    }
    // Text of a cell left of the region (scrolled away or in the frozen pane) that runs over
    // the empty cells into it
    for (let c = c0 - 1; c >= Math.max(1, c0 - 60); c--) {
      const k = key(r, c);
      const cell = sheet.cells.get(k);
      if (sheet.covered.has(k) || sheet.mergeAt.has(k)) break;
      if (!cell || !cell.text) continue;
      if (cell.kind === 's' && !xf(cell.style).wrap) {
        const box = { x: xOf(c), y: Y(r), w: xOf(c + 1) - xOf(c), h: Y(r + 1) - Y(r) };
        if (box.w > 0) drawCellText(ctx, sheet, xf(cell.style), cell, box, { r, c, c1, xOf, zoom, merged: false });
      }
      break;
    }
  }
  // Merged ranges show the text of their first cell, even when that cell is scrolled away
  for (const m of merges) {
    const cell = sheet.cells.get(key(m.first.row, m.first.col));
    if (!cell || !cell.text) continue;
    const box = rectOf(m.first.row, m.first.col, m.last.row, m.last.col);
    if (box.w > 0 && box.h > 0) drawCellText(ctx, sheet, xf(cell.style), cell, box, { r: m.first.row, c: m.first.col, c1, xOf, zoom, merged: true });
  }

  // Borders (after the text, as Excel paints them on top)
  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      const cell = sheet.cells.get(key(r, c));
      if (!cell) continue;
      const [left, right, top, bottom] = xf(cell.style).border;
      if (!left && !right && !top && !bottom) continue;
      const x1 = X(c) - 1;
      const x2 = X(c + 1) - 1;
      const y1 = Y(r) - 1;
      const y2 = Y(r + 1) - 1;
      if (top) strokeLine(ctx, x1, y1, x2 + 1, y1, top, zoom);
      if (bottom) strokeLine(ctx, x1, y2, x2 + 1, y2, bottom, zoom);
      if (left) strokeLine(ctx, x1, y1, x1, y2 + 1, left, zoom);
      if (right) strokeLine(ctx, x2, y1, x2, y2 + 1, right, zoom);
    }
  }
  ctx.setLineDash([]);
  ctx.restore();
}

export const key = (row, col) => row * 16_385 + col;

function drawCellText(ctx, sheet, style, cell, box, { r, c, c1, xOf, zoom, merged }) {
  const font = style.font;
  ctx.font = canvasFont(font, zoom);
  const color = cell.color || font.color || '#000000';
  const pad = PAD * zoom + (style.indent || 0) * 9 * zoom;
  const fontPx = ptToPx(font.size || 11) * zoom;
  const lineHeight = fontPx * 1.22;

  // Excel's General alignment: numbers right, text left, booleans and errors centered
  let h = style.h;
  if (!h || h === 'general') h = cell.kind === 'n' ? 'right' : cell.kind === 'b' || cell.kind === 'e' ? 'center' : 'left';
  if (h === 'fill' || h === 'justify' || h === 'distributed') h = 'left';
  if (h === 'centerContinuous') h = 'center';
  const v = style.v || 'bottom';

  let text = cell.text;
  let lines = [text];
  const inner = box.w - pad * 2;
  if (style.wrap && cell.kind !== 'n') {
    lines = wrapLines(ctx, text, Math.max(inner, 1));
  } else if (cell.kind === 'n' && ctx.measureText(text).width > box.w - zoom) {
    // A number that does not fit shows ### instead of being cut
    const hashWidth = ctx.measureText('#').width;
    text = '#'.repeat(Math.max(1, Math.floor((box.w - zoom) / hashWidth)));
    lines = [text];
    h = 'left';
  }

  // Where the text may spill: left-aligned text runs over empty cells to its right,
  // right-aligned over empty cells to its left, centered both ways
  let clipX = box.x;
  let clipRight = box.x + box.w;
  if (!style.wrap && !merged && cell.kind === 's') {
    const width = ctx.measureText(text).width + pad * 2;
    if (width > box.w) {
      const empty = col => {
        const other = sheet.cells.get(key(r, col));
        return !(other && other.text) && !sheet.covered.has(key(r, col)) && !sheet.mergeAt.has(key(r, col));
      };
      if (h === 'left' || h === 'center') {
        for (let col = c + 1; col <= Math.max(c1, c + 1) + 60 && empty(col) && clipRight < box.x + width; col++) clipRight = xOf(col + 1);
      }
      if (h === 'right' || h === 'center') {
        for (let col = c - 1; col >= 1 && empty(col) && clipX > box.x + box.w - width; col--) clipX = xOf(col);
      }
    }
  }

  const blockHeight = lines.length * lineHeight;
  const descent = fontPx * 0.22;
  let y;
  if (v === 'top') y = box.y + lineHeight - descent + 1 * zoom;
  else if (v === 'center' || v === 'distributed' || v === 'justify') y = box.y + (box.h - blockHeight) / 2 + lineHeight - descent;
  else y = box.y + box.h - blockHeight + lineHeight - descent - 2 * zoom;

  ctx.save();
  ctx.beginPath();
  ctx.rect(clipX, box.y, clipRight - clipX - 1, box.h - 1);
  ctx.clip();
  ctx.fillStyle = color;
  for (const line of lines) {
    const w = ctx.measureText(line).width;
    let x;
    if (h === 'right') x = box.x + box.w - pad - w;
    else if (h === 'center') x = box.x + (box.w - w) / 2;
    else x = box.x + pad;
    ctx.fillText(line, x, y);
    const thickness = Math.max(1, zoom);
    if (font.underline) ctx.fillRect(x, y + Math.max(1, fontPx * 0.08), w, thickness);
    if (font.strike) ctx.fillRect(x, y - fontPx * 0.3, w, thickness);
    y += lineHeight;
  }
  ctx.restore();
}

/**
 * Column letters and row numbers around the grid. `colSegs` and `rowSegs` are the visible
 * stretches (`{from, to, start}`, `start` in canvas pixels): one per pane.
 */
export function drawHeaders(ctx, { rows, cols, colSegs, rowSegs, rowHeaderW, colHeaderH, width, height, zoom, selection, columnName }) {
  const HEADER_BG = '#f3f4f6';
  const HEADER_LINE = '#c8ccd2';
  const HEADER_TEXT = '#4b5563';
  const ACTIVE_BG = '#d3e3d9';
  const ACTIVE_TEXT = '#185c37';
  const font = `${(11 * zoom).toFixed(2)}px system-ui, -apple-system, "Segoe UI", sans-serif`;
  ctx.save();
  ctx.font = font;
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';
  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(0, 0, width, colHeaderH);
  ctx.fillRect(0, 0, rowHeaderW, height);

  const inSel = (i, a, b) => selection && i >= a && i <= b;
  for (const seg of colSegs) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(seg.clipStart, 0, seg.clipEnd - seg.clipStart, colHeaderH);
    ctx.clip();
    for (let c = seg.from; c <= seg.to; c++) {
      const size = cols.size(c) * zoom;
      if (size === 0) continue;
      const x = seg.start + (cols.offset(c) - cols.offset(seg.from)) * zoom;
      const active = selection && inSel(c, selection.first.col, selection.last.col);
      if (active) {
        ctx.fillStyle = ACTIVE_BG;
        ctx.fillRect(x, 0, size, colHeaderH);
        ctx.fillStyle = ACTIVE_TEXT;
        ctx.fillRect(x, colHeaderH - 2, size, 2);
      }
      ctx.fillStyle = active ? ACTIVE_TEXT : HEADER_TEXT;
      ctx.fillText(columnName(c), x + size / 2, colHeaderH / 2);
      ctx.fillStyle = HEADER_LINE;
      ctx.fillRect(Math.round(x + size) - 1, 0, 1, colHeaderH);
    }
    ctx.restore();
  }
  for (const seg of rowSegs) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, seg.clipStart, rowHeaderW, seg.clipEnd - seg.clipStart);
    ctx.clip();
    for (let r = seg.from; r <= seg.to; r++) {
      const size = rows.size(r) * zoom;
      if (size === 0) continue;
      const y = seg.start + (rows.offset(r) - rows.offset(seg.from)) * zoom;
      const active = selection && inSel(r, selection.first.row, selection.last.row);
      if (active) {
        ctx.fillStyle = ACTIVE_BG;
        ctx.fillRect(0, y, rowHeaderW, size);
        ctx.fillStyle = ACTIVE_TEXT;
        ctx.fillRect(rowHeaderW - 2, y, 2, size);
      }
      ctx.fillStyle = active ? ACTIVE_TEXT : HEADER_TEXT;
      ctx.fillText(String(r), rowHeaderW / 2, y + size / 2);
      ctx.fillStyle = HEADER_LINE;
      ctx.fillRect(0, Math.round(y + size) - 1, rowHeaderW, 1);
    }
    ctx.restore();
  }
  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(0, 0, rowHeaderW - 1, colHeaderH - 1);
  ctx.fillStyle = HEADER_LINE;
  ctx.fillRect(0, colHeaderH - 1, width, 1);
  ctx.fillRect(rowHeaderW - 1, 0, 1, height);
  ctx.restore();
}
