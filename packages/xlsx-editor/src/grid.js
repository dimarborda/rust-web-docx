// Geometry of a sheet: where each row and column starts, in CSS pixels, given Excel's default
// sizes and the ones the sheet sets. No DOM here, so it is testable in Node.

/** Points to CSS pixels (96 per inch) */
export const ptToPx = pt => (pt * 96) / 72;

/**
 * Column width in characters (as stored in `<col width>`) to pixels, with `mdw` the maximum
 * digit width of the workbook's default font in pixels (7 for Calibri 11)
 */
export function colWidthToPx(width, mdw) {
  return Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw);
}

/**
 * Default column width in pixels: `defaultColWidth` when the sheet sets it, otherwise
 * `baseColWidth` characters (8 by default) plus 5 pixels of padding, rounded up to a multiple
 * of 8 as Excel does (64 px for Calibri 11)
 */
export function defaultColPx(layout, mdw) {
  if (layout.defaultColWidth) return colWidthToPx(layout.defaultColWidth, mdw);
  const base = layout.baseColWidth ?? 8;
  return Math.ceil((base * mdw + 5) / 8) * 8;
}

/**
 * Sizes along one axis (rows or columns, counted from 1): a default size and the indices that
 * differ from it. Offsets are found by binary search over the differing ones, so a sheet with a
 * million rows costs nothing more than one with ten.
 */
export class Axis {
  /**
   * @param {number} count how many rows or columns the grid shows
   * @param {number} defaultSize pixels
   * @param {Map<number, number>} custom index → size in pixels (0 hides it)
   */
  constructor(count, defaultSize, custom = new Map()) {
    this.count = count;
    this.defaultSize = defaultSize;
    this.custom = custom;
    this.indices = [...custom.keys()].filter(i => i >= 1 && i <= count).sort((a, b) => a - b);
    // prefix[k] = sum of (size - default) for indices[0..k)
    this.prefix = [0];
    for (const i of this.indices) this.prefix.push(this.prefix[this.prefix.length - 1] + custom.get(i) - defaultSize);
  }

  size(i) {
    return this.custom.has(i) ? this.custom.get(i) : this.defaultSize;
  }

  /** Start of row/column `i` (1-based); `offset(count + 1)` is the total size */
  offset(i) {
    // Custom indices below i
    let lo = 0;
    let hi = this.indices.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.indices[mid] < i) lo = mid + 1;
      else hi = mid;
    }
    return (i - 1) * this.defaultSize + this.prefix[lo];
  }

  get total() {
    return this.offset(this.count + 1);
  }

  /** The same axis with room for `count` rows/columns, up to `max` (sizes are kept) */
  grown(count, max) {
    const next = Math.min(max, Math.max(this.count, count));
    return next === this.count ? this : new Axis(next, this.defaultSize, this.custom);
  }

  /** The row/column under position `pos` (clamped to the grid) */
  indexAt(pos) {
    if (pos <= 0) return 1;
    let lo = 1;
    let hi = this.count;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.offset(mid) <= pos) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }
}

/** Last row and column of a worksheet */
export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

/** Column letters: 1 → A, 27 → AA */
export function columnName(col) {
  let name = '';
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

/** `B3` → { row: 3, col: 2 } */
export function parseRef(ref) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(String(ref).trim());
  if (!m) throw new Error(`'${ref}' no es una referencia de celda válida (por ejemplo, B3).`);
  const col = [...m[1].toUpperCase()].reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0);
  return { row: Number(m[2]), col };
}

/** `A1:D3` → corners */
export function parseRange(range) {
  const [a, b = a] = String(range).split(':');
  const p = parseRef(a);
  const q = parseRef(b);
  return {
    first: { row: Math.min(p.row, q.row), col: Math.min(p.col, q.col) },
    last: { row: Math.max(p.row, q.row), col: Math.max(p.col, q.col) },
  };
}

/**
 * Rows and columns of a sheet as axes: sizes from the layout, extent from the used cells plus
 * some room to scroll past them
 * @param {import('./engine.js').SheetLayout} layout
 * @param {{maxRow: number, maxCol: number}} used
 * @param {number} mdw maximum digit width of the default font, pixels
 * @param {number} defaultRowPt default row height when the sheet sets none, points
 */
export function sheetAxes(layout, used, mdw, defaultRowPt = 15) {
  const rowPx = Math.round(ptToPx(layout.defaultRowHeight || defaultRowPt));
  const colPx = defaultColPx(layout, mdw);
  const rowSizes = new Map();
  for (const [row, height, hidden] of layout.rows) rowSizes.set(row, hidden ? 0 : Math.round(ptToPx(height ?? (layout.defaultRowHeight || defaultRowPt))));
  const colSizes = new Map();
  let lastCol = 0;
  for (const [min, max, width, hidden] of layout.cols) {
    // A <col> spanning to XFD sets the rest of the sheet; only the shown part matters
    const end = Math.min(max, Math.max(min, used.maxCol + 50));
    for (let c = min; c <= end; c++) colSizes.set(c, hidden ? 0 : width == null ? colPx : colWidthToPx(width, mdw));
    lastCol = Math.max(lastCol, Math.min(max, end));
  }
  let lastRow = used.maxRow;
  for (const [row] of layout.rows) lastRow = Math.max(lastRow, row);
  for (const m of layout.merges) {
    const { last } = parseRange(m);
    lastRow = Math.max(lastRow, last.row);
    lastCol = Math.max(lastCol, last.col);
  }
  const rows = Math.min(MAX_ROWS, Math.max(lastRow + 100, 200));
  const cols = Math.min(MAX_COLS, Math.max(lastCol, used.maxCol) + 12);
  return { rows: new Axis(rows, rowPx, rowSizes), cols: new Axis(Math.max(cols, 26), colPx, colSizes) };
}
