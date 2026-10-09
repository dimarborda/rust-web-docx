// Vector PDF export: the pages laid out by the engine (the same render commands the canvases
// draw) written as a PDF with real, selectable text. Loaded on demand by `exportPdf()`, so
// pdf-lib and fontkit are only downloaded when someone exports.
//
// Text is placed word by word at the positions the layout computed, in the font the page
// shows: an uploaded font, or the metric-compatible substitute (Carlito, Arimo, Tinos,
// Caladea, Cousine) found in the page's @font-face rules. When neither is available the PDF
// falls back to its standard fonts (Helvetica, Times, Courier).
import { customFontRecord } from './fonts_registry.js';

const PT = 0.75; // 1 CSS px = 0.75 pt

/**
 * @param {{pages: object[]}} layout  the editor's current layout
 * @param {{title?: string, watermark?: boolean}} [options]
 * @returns {Promise<Uint8Array>}
 */
export async function exportPdf(layout, { title, watermark = true } = {}) {
  const lib = await import('pdf-lib');
  const fontkit = (await import('@pdf-lib/fontkit')).default;
  const doc = await lib.PDFDocument.create();
  doc.registerFontkit(fontkit);
  if (title) doc.setTitle(title);
  doc.setCreator('@dimarborda/docx-editor');
  doc.setProducer('Rust DOCX (pdf-lib)');

  const fonts = new FontSet(doc, lib);
  const images = new ImageSet(doc);
  for (const page of layout.pages) {
    const pdfPage = doc.addPage([page.width * PT, page.height * PT]);
    const ctx = { lib, page: pdfPage, height: page.height, fonts, images };
    const bg = parseColor(page.bg_color, lib);
    if (bg && page.bg_color.replace('#', '').toUpperCase() !== 'FFFFFF') {
      pdfPage.drawRectangle({ x: 0, y: 0, width: page.width * PT, height: page.height * PT, color: bg });
    }
    for (const item of page.items) {
      if (item.type === 'watermark' && !watermark) continue;
      await drawItem(ctx, item);
    }
  }
  return doc.save();
}

// ---------- Commands ----------

async function drawItem(ctx, item) {
  switch (item.type) {
    case 'text':
      return item.runs?.length ? drawRuns(ctx, item) : drawPlainText(ctx, item);
    case 'line':
      return line(ctx, item.x1, item.y1, item.x2, item.y2, item.color || '#CBD5E1', item.line_width);
    case 'table_cell':
      return drawCell(ctx, item);
    case 'image':
      return drawImage(ctx, item);
    case 'shape':
      return drawShape(ctx, item);
    case 'watermark':
      return drawWatermark(ctx, item);
    default:
      return undefined;
  }
}

/** Page px → PDF points, with the y axis pointing up */
const px = (ctx, x, y) => ({ x: x * PT, y: (ctx.height - y) * PT });

function line(ctx, x1, y1, x2, y2, color, width, dashArray) {
  const c = parseColor(color, ctx.lib);
  if (!c || !(width > 0)) return;
  ctx.page.drawLine({ start: px(ctx, x1, y1), end: px(ctx, x2, y2), thickness: width * PT, color: c, dashArray });
}

async function drawRuns(ctx, item) {
  const extra = item.line ? item.line.space_extra : 0;
  for (const run of item.runs) {
    if (run.text === '\t' || !run.text) continue;
    const family = run.font_family || item.font_family;
    const size = run.font_size || item.font_size || 14.66;
    const color = parseColor(run.color || item.color, ctx.lib) || ctx.lib.rgb(0, 0, 0);
    const face = await ctx.fonts.face(family, run.bold, run.italic);
    // The layout measured the run with the browser's font: scale the PDF font's advances to
    // that width, so words land exactly where the page shows them
    const natural = face.width(run.text, size);
    // The layout measured the run with the browser's font. When the PDF font differs (a
    // system font that cannot be embedded), its glyphs are squeezed or stretched to the same
    // width (PDF horizontal scaling), so every word lands where the page shows it.
    const scale = natural > 0 && run.width > 0 ? Math.min(1.5, Math.max(0.67, run.width / natural)) : 1;
    const squeeze = Math.abs(scale - 1) > 0.005;
    if (squeeze) ctx.page.pushOperators(ctx.lib.pushGraphicsState(), ctx.lib.setCharacterSqueeze(scale * 100));
    let x = run.x;
    const words = run.text.split(' ');
    words.forEach((word, i) => {
      if (i > 0) x += face.width(' ', size) * scale + extra;
      if (word) {
        face.draw(ctx, word, x, item.y, size, color);
        x += face.width(word, size) * scale;
      }
    });
    if (squeeze) ctx.page.pushOperators(ctx.lib.popGraphicsState());
    if (run.underline) line(ctx, run.x, item.y + 2, x, item.y + 2, run.color || item.color || '#000000', 1);
  }
}

/** Page decorations without runs (old-style header text, page labels) */
async function drawPlainText(ctx, item) {
  if (!item.text) return;
  const bold = item.font_weight === 'bold' || Number(item.font_weight) >= 600;
  const face = await ctx.fonts.face(item.font_family, bold, item.font_style === 'italic');
  const size = item.font_size || 14.66;
  const width = face.width(item.text, size);
  const x = item.align === 'right' ? item.x - width : item.align === 'center' ? item.x - width / 2 : item.x;
  face.draw(ctx, item.text, x, item.y, size, parseColor(item.color, ctx.lib) || ctx.lib.rgb(0, 0, 0));
}

function drawCell(ctx, item) {
  const { x, y, width: w, height: h } = item;
  const fill = item.bg_color && parseColor(item.bg_color, ctx.lib);
  if (fill) {
    const p = px(ctx, x, y + h);
    ctx.page.drawRectangle({ x: p.x, y: p.y, width: w * PT, height: h * PT, color: fill });
  }
  const sides = item.borders || {};
  const side = (border, x1, y1, x2, y2) => {
    if (!border) return; // gridlines are not printed
    const width = border.sz_px || 0.75;
    const dash = border.val === 'dashed' || border.val === 'dashSmallGap' ? [width * 4 * PT, width * 2 * PT]
      : border.val === 'dotted' ? [width * PT, width * 1.5 * PT] : undefined;
    if (border.val === 'double') {
      const dx = y1 === y2 ? 0 : width;
      const dy = y1 === y2 ? width : 0;
      line(ctx, x1 - dx, y1 - dy, x2 - dx, y2 - dy, border.color, Math.max(0.5, width / 2));
      line(ctx, x1 + dx, y1 + dy, x2 + dx, y2 + dy, border.color, Math.max(0.5, width / 2));
    } else {
      line(ctx, x1, y1, x2, y2, border.color || '#000000', width, dash);
    }
  };
  side(sides.top, x, y, x + w, y);
  side(sides.bottom, x, y + h, x + w, y + h);
  side(sides.left, x, y, x, y + h);
  side(sides.right, x + w, y, x + w, y + h);
}

async function drawImage(ctx, item) {
  const image = await ctx.images.get(item.data_url);
  if (!image) return;
  const p = px(ctx, item.x, item.y + item.height);
  ctx.page.drawImage(image, { x: p.x, y: p.y, width: item.width * PT, height: item.height * PT, opacity: item.opacity ?? 1 });
}

function drawShape(ctx, item) {
  const fill = item.fill && item.geometry !== 'line' ? parseColor(item.fill, ctx.lib) : undefined;
  const stroke = item.stroke && item.stroke_width > 0 ? parseColor(item.stroke, ctx.lib) : undefined;
  if (!fill && !stroke) return;
  const { x, y, width: w, height: h } = item;
  const border = stroke ? { borderColor: stroke, borderWidth: item.stroke_width * PT } : {};
  if (item.geometry === 'line') {
    line(ctx, x, y, x + w, y + h, item.stroke, item.stroke_width);
  } else if (item.geometry === 'ellipse') {
    const c = px(ctx, x + w / 2, y + h / 2);
    ctx.page.drawEllipse({ x: c.x, y: c.y, xScale: (w / 2) * PT, yScale: (h / 2) * PT, color: fill, ...border });
  } else if (item.geometry === 'roundRect') {
    const r = Math.min(w, h) * 0.1667;
    const path = `M ${r} 0 H ${w - r} A ${r} ${r} 0 0 1 ${w} ${r} V ${h - r} A ${r} ${r} 0 0 1 ${w - r} ${h} `
      + `H ${r} A ${r} ${r} 0 0 1 0 ${h - r} V ${r} A ${r} ${r} 0 0 1 ${r} 0 Z`;
    const o = px(ctx, x, y);
    ctx.page.drawSvgPath(path, { x: o.x, y: o.y, scale: PT, color: fill, ...border });
  } else {
    const p = px(ctx, x, y + h);
    ctx.page.drawRectangle({ x: p.x, y: p.y, width: w * PT, height: h * PT, color: fill, ...border });
  }
}

async function drawWatermark(ctx, item) {
  const face = await ctx.fonts.face('Arial', true, false);
  const size = item.font_size;
  const width = face.width(item.text, size);
  // The canvas turns the text by rotation_deg (y down); in the PDF (y up) that is the opposite
  const angle = (-item.rotation_deg * Math.PI) / 180;
  const dx = -width / 2;
  const dy = -size * 0.35; // from the middle of the letters down to the baseline
  const ox = item.x + dx * Math.cos(angle) - dy * Math.sin(angle);
  const oy = item.y - (dx * Math.sin(angle) + dy * Math.cos(angle));
  face.draw(ctx, item.text, ox, oy, size, parseColor(item.color || '#64748B', ctx.lib), {
    rotate: ctx.lib.degrees((angle * 180) / Math.PI),
    opacity: item.opacity,
  });
}

// ---------- Fonts ----------

/** Metric-compatible substitute the renderer falls back to for a Word font */
function substituteFamily(family) {
  const f = (family || '').toLowerCase();
  if (f.includes('consolas') || f.includes('courier') || f.includes('mono')) return 'Cousine';
  if (f.includes('cambria')) return 'Caladea';
  if (f.includes('times') || f.includes('tinos') || f.includes('georgia') || f.includes('serif') && !f.includes('sans')) return 'Tinos';
  if (f.includes('arial') || f.includes('helvetica') || f.includes('arimo')) return 'Arimo';
  return 'Carlito';
}

function standardFont(family, bold, italic, StandardFonts) {
  const sub = substituteFamily(family);
  const pick = (regular, b, i, bi) => (bold && italic ? bi : bold ? b : italic ? i : regular);
  if (sub === 'Cousine') return pick(StandardFonts.Courier, StandardFonts.CourierBold, StandardFonts.CourierOblique, StandardFonts.CourierBoldOblique);
  if (sub === 'Tinos' || sub === 'Caladea') {
    return pick(StandardFonts.TimesRoman, StandardFonts.TimesRomanBold, StandardFonts.TimesRomanItalic, StandardFonts.TimesRomanBoldItalic);
  }
  return pick(StandardFonts.Helvetica, StandardFonts.HelveticaBold, StandardFonts.HelveticaOblique, StandardFonts.HelveticaBoldOblique);
}

/** First family of a CSS font stack, without quotes */
const firstFamily = family => (family || '').split(',')[0].trim().replace(/^["']|["']$/g, '');

/** @font-face rules of the page for `family` (case-insensitive), with their sources */
function fontFaceRules(family) {
  const out = [];
  const wanted = family.toLowerCase();
  const visit = (rules, base) => {
    for (const rule of rules) {
      if (rule.cssRules && !(rule instanceof CSSFontFaceRule)) {
        try {
          visit(rule.cssRules, rule.href || base);
        } catch {
          // cross-origin stylesheet: not readable
        }
        continue;
      }
      if (!(rule instanceof CSSFontFaceRule)) continue;
      const name = rule.style.getPropertyValue('font-family').trim().replace(/^["']|["']$/g, '').toLowerCase();
      if (name !== wanted) continue;
      const sources = [...rule.style.getPropertyValue('src').matchAll(/url\(\s*["']?([^"')]+)["']?\s*\)\s*(?:format\(\s*["']?([\w-]+)["']?\s*\))?/g)]
        .map(m => ({ url: new URL(m[1], base).href, format: (m[2] || m[1].split('.').pop() || '').toLowerCase() }));
      out.push({
        weight: Number(rule.style.getPropertyValue('font-weight')) || 400,
        italic: /italic|oblique/.test(rule.style.getPropertyValue('font-style')),
        range: rule.style.getPropertyValue('unicode-range'),
        sources,
      });
    }
  };
  for (const sheet of document.styleSheets) {
    try {
      visit(sheet.cssRules, sheet.href || location.href);
    } catch {
      // cross-origin stylesheet
    }
  }
  return out;
}

/** True for unicode ranges that cover basic Latin (the face to try first) */
const coversLatin = range => !range || /U\+0(-|0[0-9A-F]{0,2}-)/i.test(range);

/**
 * The fonts a run is drawn with: one or more embedded files of the same face (fontsource
 * splits faces by script), plus a standard font for anything they lack
 */
class FontSet {
  constructor(doc, lib) {
    this.doc = doc;
    this.lib = lib;
    this.faces = new Map();
    this.files = new Map();
  }

  face(family, bold, italic) {
    const key = `${firstFamily(family).toLowerCase()}|${!!bold}|${!!italic}`;
    if (!this.faces.has(key)) this.faces.set(key, this.#build(firstFamily(family), !!bold, !!italic));
    return this.faces.get(key);
  }

  async #build(family, bold, italic) {
    const fonts = [];
    // 1. A font the user uploaded for this family (e.g. the real Calibri)
    const record = customFontRecord(family, bold, italic);
    if (record) {
      const font = await this.#embed(`custom:${record.id}`, () => record.buffer, !/woff2/i.test(record.fileName || ''));
      if (font) fonts.push(font);
    }
    // 2. The substitute the page draws with, from its @font-face rules
    if (!fonts.length) {
      const rules = fontFaceRules(substituteFamily(family));
      const weight = bold ? 700 : 400;
      const matching = rules
        .filter(r => r.italic === italic)
        .sort((a, b) => Math.abs(a.weight - weight) - Math.abs(b.weight - weight));
      const best = matching.length ? matching.filter(r => r.weight === matching[0].weight) : [];
      // Basic Latin first (browsers may not expose unicode-range: the file name tells too)
      const latin = r => coversLatin(r.range) && !r.sources.some(s => /latin-ext|cyrillic|greek|vietnamese/i.test(s.url));
      best.sort((a, b) => Number(latin(b)) - Number(latin(a)));
      for (const rule of best) {
        // WOFF subsets well; WOFF2 is embedded whole (fontkit cannot subset it reliably)
        const source = rule.sources.find(s => s.format === 'woff') || rule.sources.find(s => /truetype|opentype|ttf|otf/.test(s.format))
          || rule.sources.find(s => s.format === 'woff2');
        if (!source) continue;
        const font = await this.#embed(source.url, async () => (await fetch(source.url)).arrayBuffer(), source.format !== 'woff2');
        if (font) fonts.push(font);
      }
    }
    // 3. The PDF's own fonts for whatever is left
    const standard = await this.doc.embedFont(standardFont(family, bold, italic, this.lib.StandardFonts));
    return new Face([...fonts, standard], standard);
  }

  async #embed(key, load, subset) {
    if (!this.files.has(key)) {
      this.files.set(key, (async () => {
        try {
          return await this.doc.embedFont(new Uint8Array(await load()), { subset });
        } catch (err) {
          console.warn('docx-editor: font could not be embedded in the PDF', key, err);
          return null;
        }
      })());
    }
    return this.files.get(key);
  }
}

/** Draws text with the first font of the list that has each character */
class Face {
  constructor(fonts, standard) {
    this.fonts = fonts;
    this.standard = standard;
    this.sets = fonts.map(f => new Set(f.getCharacterSet()));
  }

  /** Pieces of `text` that one font can draw: [[font, text], …] */
  pieces(text) {
    const out = [];
    for (const ch of text) {
      const code = ch.codePointAt(0);
      let i = this.sets.findIndex(set => set.has(code));
      let c = ch;
      if (i < 0) {
        i = this.fonts.length - 1;
        c = '?';
      }
      const last = out[out.length - 1];
      if (last && last[0] === this.fonts[i]) last[1] += c;
      else out.push([this.fonts[i], c]);
    }
    return out;
  }

  /** Width in page px of `text` at `size` px */
  width(text, size) {
    return this.pieces(text).reduce((sum, [font, piece]) => sum + font.widthOfTextAtSize(piece, size * PT) / PT, 0);
  }

  draw(ctx, text, x, baseline, size, color, extra = {}) {
    let cursor = x;
    for (const [font, piece] of this.pieces(text)) {
      const p = px(ctx, cursor, baseline);
      ctx.page.drawText(piece, { x: p.x, y: p.y, size: size * PT, font, color, ...extra });
      cursor += font.widthOfTextAtSize(piece, size * PT) / PT;
    }
  }
}

// ---------- Images ----------

class ImageSet {
  constructor(doc) {
    this.doc = doc;
    this.cache = new Map();
  }

  get(dataUrl) {
    if (!dataUrl) return null;
    if (!this.cache.has(dataUrl)) this.cache.set(dataUrl, this.#embed(dataUrl).catch(err => {
      console.warn('docx-editor: picture could not be embedded in the PDF', err);
      return null;
    }));
    return this.cache.get(dataUrl);
  }

  async #embed(dataUrl) {
    const mime = /^data:([^;,]+)/.exec(dataUrl)?.[1] || '';
    if (mime === 'image/png') return this.doc.embedPng(dataUrl);
    if (mime === 'image/jpeg' || mime === 'image/jpg') return this.doc.embedJpg(dataUrl);
    // GIF, SVG, WebP…: drawn once to a canvas and embedded as PNG
    const img = new Image();
    img.src = dataUrl;
    await img.decode();
    const canvas = document.createElement('canvas');
    canvas.width = img.naturalWidth || 1;
    canvas.height = img.naturalHeight || 1;
    canvas.getContext('2d').drawImage(img, 0, 0);
    return this.doc.embedPng(canvas.toDataURL('image/png'));
  }
}

// ---------- Colors ----------

function parseColor(color, lib) {
  if (!color || color === 'auto') return null;
  const hex = /^#?([0-9a-f]{6})$/i.exec(color.trim());
  if (hex) {
    const v = parseInt(hex[1], 16);
    return lib.rgb(((v >> 16) & 255) / 255, ((v >> 8) & 255) / 255, (v & 255) / 255);
  }
  const short = /^#([0-9a-f]{3})$/i.exec(color.trim());
  if (short) {
    const [r, g, b] = short[1].split('').map(c => parseInt(c + c, 16) / 255);
    return lib.rgb(r, g, b);
  }
  const rgbMatch = /^rgba?\(([^)]+)\)/i.exec(color.trim());
  if (rgbMatch) {
    const [r, g, b] = rgbMatch[1].split(',').map(n => parseFloat(n) / 255);
    return lib.rgb(r || 0, g || 0, b || 0);
  }
  return lib.rgb(0, 0, 0);
}
