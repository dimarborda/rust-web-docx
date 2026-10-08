// Draws the render commands produced by the Rust layout engine onto a 2D canvas, and
// measures text for that engine with the very same fonts, so line breaks and caret positions
// match the pixels.
import { hasCustomFont } from './fonts_registry.js';

/**
 * CSS font stack for a document font: the font itself when installed or uploaded, otherwise a
 * metric-compatible substitute (Carlito ≈ Calibri, Caladea ≈ Cambria, Arimo ≈ Arial,
 * Tinos ≈ Times New Roman, Cousine ≈ Courier New/Consolas).
 */
export function fontStack(family) {
  if (!family) return '"Carlito", "Calibri", "Arimo", "Arial", sans-serif';
  const lower = family.toLowerCase().trim();
  if (hasCustomFont(lower) || (document.fonts && document.fonts.check(`16px "${family}"`))) {
    return `"${family}", "Carlito", "Arimo", sans-serif`;
  }
  if (lower.includes('calibri')) return '"Calibri", "Carlito", "Segoe UI", Roboto, sans-serif';
  if (lower.includes('consolas') || lower.includes('courier') || lower.includes('mono')) return '"Consolas", "Cousine", monospace';
  if (lower.includes('cambria')) return '"Caladea", "Cambria", Georgia, serif';
  if (lower.includes('times')) return '"Tinos", "Times New Roman", Georgia, serif';
  if (lower.includes('arial') || lower.includes('helvetica')) return '"Arimo", "Arial", Helvetica, sans-serif';
  if (lower.includes('aptos')) return '"Aptos", "Calibri", "Carlito", "Segoe UI", sans-serif';
  return `"${family}", "Carlito", "Arimo", sans-serif`;
}

/** True when the document font is drawn with a substitute rather than the real font */
export function usesSubstitute(family) {
  const lower = (family || '').toLowerCase().trim();
  if (hasCustomFont(lower)) return false;
  const installed = document.fonts && document.fonts.check(`16px "${family}"`);
  return lower.includes('calibri') || lower.includes('aptos') || !installed;
}

export function canvasFont(weight, style, size, family) {
  const s = style === 'italic' || style === 'oblique' ? 'italic ' : '';
  const w = weight && weight !== 'normal' && weight !== '400' ? `${weight} ` : '';
  return `${s}${w}${size}px ${fontStack(family)}`;
}

// Word applies neither ligatures ("fi") nor kerning by default; with both off every glyph
// sits exactly at the sum of the advances measured for the caret
export function disableLigatures(ctx) {
  if ('textRendering' in ctx) ctx.textRendering = 'optimizeSpeed';
  if ('fontKerning' in ctx) ctx.fontKerning = 'none';
}

const measureCtx = document.createElement('canvas').getContext('2d');
disableLigatures(measureCtx);

/** `(text, family, sizePx, bold, italic) => width`, the measurer the Rust layout calls */
export function measureText(text, family, size, bold, italic) {
  measureCtx.font = canvasFont(bold ? '700' : '400', italic ? 'italic' : 'normal', size, family);
  return measureCtx.measureText(text).width;
}

const PROBE_TEXT = 'Contrato de prestación de servicios 0123456789';

/**
 * Widths of a probe text in each canvas font: changes when a web font becomes usable for
 * canvas text, which some engines (Safari) only allow a moment after reporting it loaded
 */
export function fontWidthsSignature(faces) {
  return [...faces].map(face => {
    measureCtx.font = face;
    return measureCtx.measureText(PROBE_TEXT).width.toFixed(2);
  }).join('|');
}

/** Every font face a layout uses, as canvas font strings */
export function layoutFontFaces(layout) {
  const faces = new Set();
  layout.pages.forEach(page => page.items.forEach(item => {
    if (item.type === 'text' && item.runs) {
      item.runs.forEach(run => faces.add(
        canvasFont(run.bold ? '700' : '400', run.italic ? 'italic' : 'normal', 16, run.font_family || item.font_family),
      ));
    } else if (item.type === 'table_cell') {
      faces.add(canvasFont(item.font_weight, 'normal', 16, item.font_family));
    }
  }));
  return faces;
}

export function cssColor(color, fallback = '#000000') {
  if (!color || color === 'auto') return fallback;
  if (color.startsWith('#') || color.startsWith('rgb') || color.startsWith('hsl')) return color;
  return `#${color}`;
}

const images = new Map(); // data URL → HTMLImageElement, shared by every editor

function drawCellSide(ctx, border, gridlines, x1, y1, x2, y2) {
  if (!border && !gridlines) return;
  ctx.save();
  ctx.beginPath();
  if (border) {
    const width = border.sz_px || 0.75;
    ctx.strokeStyle = cssColor(border.color);
    ctx.lineWidth = width;
    if (border.val === 'dashed' || border.val === 'dashSmallGap') ctx.setLineDash([width * 4, width * 2]);
    else if (border.val === 'dotted') ctx.setLineDash([width, width * 1.5]);
    if (border.val === 'double') {
      // Two thin lines, as Word draws "double"
      const dx = y1 === y2 ? 0 : width;
      const dy = y1 === y2 ? width : 0;
      ctx.lineWidth = Math.max(0.5, width / 2);
      ctx.moveTo(x1 - dx, y1 - dy); ctx.lineTo(x2 - dx, y2 - dy);
      ctx.moveTo(x1 + dx, y1 + dy); ctx.lineTo(x2 + dx, y2 + dy);
    } else {
      ctx.moveTo(x1, y1); ctx.lineTo(x2, y2);
    }
  } else {
    // Word's non-printing gridlines (View → Gridlines) where a cell side has no border
    ctx.strokeStyle = 'rgba(100, 116, 139, 0.45)';
    ctx.lineWidth = 0.5;
    ctx.setLineDash([2, 2]);
    ctx.moveTo(x1, y1); ctx.lineTo(x2, y2);
  }
  ctx.stroke();
  ctx.restore();
}

/**
 * Draws one page's render commands.
 * @returns {Promise<void>[]} images still loading; draw the page again when they settle
 */
export function drawPageItems(ctx, items, { gridlines = true } = {}) {
  const pending = [];
  items.forEach(item => {
    if (item.type === 'watermark') {
      ctx.save();
      ctx.translate(item.x, item.y);
      ctx.rotate((item.rotation_deg * Math.PI) / 180);
      ctx.globalAlpha = item.opacity;
      ctx.font = `700 ${item.font_size}px "Arimo", "Arial", sans-serif`;
      ctx.fillStyle = cssColor(item.color, '#64748B');
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(item.text, 0, 0);
      ctx.restore();
    } else if (item.type === 'image' && item.data_url) {
      let img = images.get(item.data_url);
      if (!img) {
        img = new Image();
        img.src = item.data_url;
        images.set(item.data_url, img);
      }
      if (img.complete && img.naturalWidth > 0) {
        ctx.save();
        if (item.opacity !== undefined && item.opacity < 1.0) ctx.globalAlpha = item.opacity;
        ctx.drawImage(img, item.x, item.y, item.width, item.height);
        ctx.restore();
      } else {
        pending.push(img.decode().catch(() => {}));
      }
    } else if (item.type === 'line') {
      ctx.beginPath();
      ctx.moveTo(item.x1, item.y1);
      ctx.lineTo(item.x2, item.y2);
      ctx.strokeStyle = cssColor(item.color, '#CBD5E1');
      ctx.lineWidth = item.line_width;
      ctx.stroke();
    } else if (item.type === 'table_cell') {
      // Shading only when the document sets it; the cell's text comes as regular lines
      if (item.bg_color) {
        ctx.fillStyle = cssColor(item.bg_color, '#FFFFFF');
        ctx.fillRect(item.x, item.y, item.width, item.height);
      }
      const { x, y, width: w, height: h } = item;
      const sides = item.borders || {};
      drawCellSide(ctx, sides.top, gridlines, x, y, x + w, y);
      drawCellSide(ctx, sides.bottom, gridlines, x, y + h, x + w, y + h);
      drawCellSide(ctx, sides.left, gridlines, x, y, x, y + h);
      drawCellSide(ctx, sides.right, gridlines, x + w, y, x + w, y + h);
    } else if (item.type === 'text') {
      drawText(ctx, item);
    }
  });
  return pending;
}

function drawText(ctx, item) {
  if (item.runs && item.runs.length > 0) {
    // Runs go exactly where the layout placed them (run.x), with its justification spacing,
    // so the caret geometry computed in Rust matches the pixels
    const extra = item.line ? item.line.space_extra : 0;
    let nextX = item.x;
    item.runs.forEach(run => {
      const startX = run.x ?? nextX;
      if (run.text === '\t') {
        nextX = startX + (run.width || 0);
        return;
      }
      const color = cssColor(run.color || item.color);
      ctx.font = canvasFont(run.bold ? '700' : '400', run.italic ? 'italic' : 'normal',
        run.font_size || item.font_size || 14.66, run.font_family || item.font_family);
      ctx.fillStyle = color;
      ctx.textAlign = 'left';
      ctx.textBaseline = 'alphabetic';

      let curX = startX;
      if (extra > 0 && run.text.includes(' ')) {
        run.text.split(' ').forEach((word, i) => {
          if (i > 0) curX += ctx.measureText(' ').width + extra;
          if (word) {
            ctx.fillText(word, curX, item.y);
            curX += ctx.measureText(word).width;
          }
        });
      } else {
        ctx.fillText(run.text, curX, item.y);
        curX += ctx.measureText(run.text).width;
      }

      if (run.underline) {
        ctx.beginPath();
        ctx.moveTo(startX, item.y + 2);
        ctx.lineTo(curX, item.y + 2);
        ctx.strokeStyle = color;
        ctx.lineWidth = 1;
        ctx.stroke();
      }
      nextX = curX;
    });
  } else if (item.text) {
    // Page decorations (header, footer, page numbers)
    ctx.font = canvasFont(item.font_weight, item.font_style, item.font_size || 14.66, item.font_family);
    ctx.fillStyle = cssColor(item.color);
    ctx.textAlign = item.align === 'right' ? 'right' : item.align === 'center' ? 'center' : 'left';
    ctx.textBaseline = 'alphabetic';
    ctx.fillText(item.text, item.x, item.y);
  }
}
