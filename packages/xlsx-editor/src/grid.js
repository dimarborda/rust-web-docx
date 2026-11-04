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

/** Pixels to the width in characters `colWidthToPx` turns back into those pixels */
export function pxToColWidth(px, mdw) {
  return Math.round((px / mdw) * 256) / 256;
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

/**
 * What Excel's status bar says about the cells of a range: how many have something and the sum,
 * average, smallest and largest of the numbers. Hidden rows and columns do not count, and any
 * error leaves only the count (as Excel does). `style` is the format of the first number. `null`
 * when the range is empty.
 * @param {Map<number, {kind: string, style: number, number?: number}>} cells by `row * 16385 + col`
 * @param {{first: {row: number, col: number}, last: {row: number, col: number}}} range
 * @param {{rows?: Axis, cols?: Axis}} [axes]
 */
export function rangeStats(cells, { first, last }, { rows, cols } = {}) {
  let count = 0;
  let numbers = 0;
  let sum = 0;
  let min = Infinity;
  let max = -Infinity;
  let error = false;
  let style = 0;
  let firstAt = Infinity;
  const visit = (cell, row, col) => {
    if (!cell || !cell.kind || rows?.size(row) === 0 || cols?.size(col) === 0) return;
    count++;
    if (cell.kind === 'e') error = true;
    if (cell.number === undefined) return;
    numbers++;
    sum += cell.number;
    min = Math.min(min, cell.number);
    max = Math.max(max, cell.number);
    const at = row * 16_385 + col;
    if (at < firstAt) {
      firstAt = at;
      style = cell.style;
    }
  };
  // Whichever is smaller: the cells of the range or those of the sheet
  if ((last.row - first.row + 1) * (last.col - first.col + 1) <= cells.size) {
    for (let r = first.row; r <= last.row; r++) {
      for (let c = first.col; c <= last.col; c++) visit(cells.get(r * 16_385 + c), r, c);
    }
  } else {
    for (const [at, cell] of cells) {
      const r = Math.floor(at / 16_385);
      const c = at - r * 16_385;
      if (r >= first.row && r <= last.row && c >= first.col && c <= last.col) visit(cell, r, c);
    }
  }
  if (!count) return null;
  if (error || !numbers) return { count, numbers: error ? 0 : numbers, sum: null, average: null, min: null, max: null, style };
  return { count, numbers, sum, average: sum / numbers, min, max, style };
}

/**
 * Where Ctrl+arrow lands from the line `from`, as Excel does: to the end of the block of filled
 * cells it is in, else to the next filled cell, else to the end of the sheet
 * @param {(line: number) => boolean} filled
 * @param {number} from
 * @param {1 | -1} step
 * @param {number} limit last line of the sheet
 */
export function dataEdge(filled, from, step, limit) {
  const inside = i => i >= 1 && i <= limit;
  if (!inside(from + step)) return from;
  let i = from;
  if (filled(from) && filled(from + step)) {
    while (inside(i + step) && filled(i + step)) i += step;
    return i;
  }
  for (i = from + step; inside(i); i += step) if (filled(i)) return i;
  return step < 0 ? 1 : limit;
}

/**
 * The far end of the run of numbers next to the line `from`, going by `step`; null when the next
 * line has none. Autosum takes that run.
 * @param {(line: number) => boolean} isNumber
 * @param {number} from
 * @param {1 | -1} step
 * @param {number} limit last line of the sheet
 */
export function numberRun(isNumber, from, step, limit) {
  let end = null;
  for (let i = from + step; i >= 1 && i <= limit && isNumber(i); i += step) end = i;
  return end;
}
