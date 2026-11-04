// Grid geometry: Excel's pixel sizes, offsets with custom and hidden rows, and the sheet extent.
// Run with: npm test -w packages/xlsx-editor
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Axis, MAX_ROWS, colWidthToPx, columnName, dataEdge, defaultColPx, parseRange, parseRef, ptToPx, pxToColWidth, rangeStats, sheetAxes } from '../src/grid.js';

test('Excel sizes in pixels', () => {
  assert.equal(Math.round(ptToPx(15)), 20);
  assert.equal(colWidthToPx(9.140625, 7), 64, 'the stored default width of Calibri 11');
  assert.equal(colWidthToPx(22, 7), 154);
  assert.equal(defaultColPx({ baseColWidth: 8 }, 7), 64);
  assert.equal(defaultColPx({}, 7), 64);
  assert.equal(defaultColPx({ baseColWidth: 10 }, 7), 80);
  assert.equal(defaultColPx({ defaultColWidth: 9.140625 }, 7), 64);
});

test('offsets and positions with custom and hidden sizes', () => {
  const axis = new Axis(1_048_576, 20, new Map([[3, 40], [5, 0], [1_000_000, 100]]));
  assert.equal(axis.offset(1), 0);
  assert.equal(axis.offset(3), 40);
  assert.equal(axis.offset(4), 80);
  assert.equal(axis.offset(5), 100);
  assert.equal(axis.offset(6), 100, 'row 5 is hidden');
  assert.equal(axis.size(5), 0);
  assert.equal(axis.indexAt(0), 1);
  assert.equal(axis.indexAt(45), 3);
  assert.equal(axis.indexAt(79.9), 3);
  assert.equal(axis.indexAt(80), 4);
  assert.equal(axis.indexAt(100), 6, 'a hidden row is never under the pointer');
  assert.equal(axis.total, 1_048_576 * 20 + 20 - 20 + 80);
  assert.equal(axis.indexAt(axis.offset(999_999) + 5), 999_999);
  assert.equal(axis.indexAt(1e12), 1_048_576);
});

test('references', () => {
  assert.equal(columnName(1), 'A');
  assert.equal(columnName(28), 'AB');
  assert.equal(columnName(16384), 'XFD');
  assert.deepEqual(parseRef('$b$3'), { row: 3, col: 2 });
  assert.deepEqual(parseRange('D3:A1'), { first: { row: 1, col: 1 }, last: { row: 3, col: 4 } });
  assert.throws(() => parseRef('3B'));
});

test('sheet axes cover the used cells and some room to scroll', () => {
  const layout = {
    defaultRowHeight: 15, defaultColWidth: null, baseColWidth: null,
    cols: [[1, 1, 22, false], [3, 16384, 12, false]], rows: [[2, 30, false], [4, null, true]],
    merges: ['A1:F1'], gridlines: true, headings: true, rightToLeft: false, frozen: null,
  };
  const { rows, cols } = sheetAxes(layout, { maxRow: 40, maxCol: 6 }, 7);
  assert.equal(rows.count, 200);
  assert.equal(rows.size(2), 40);
  assert.equal(rows.size(4), 0);
  assert.equal(cols.size(1), 154);
  assert.equal(cols.size(2), 64);
  assert.equal(cols.size(3), colWidthToPx(12, 7));
  assert.ok(cols.count >= 26 && cols.count < 100, `a <col> to XFD does not make 16384 columns (${cols.count})`);
  const big = sheetAxes({ ...layout, cols: [] }, { maxRow: 5000, maxCol: 60 }, 7);
  assert.equal(big.rows.count, 5100);
  assert.equal(big.cols.count, 72);
});

test('an axis grows keeping its sizes, up to the end of the sheet', () => {
  const axis = new Axis(200, 20, new Map([[3, 0], [10, 40]]));
  const grown = axis.grown(500, MAX_ROWS);
  assert.equal(grown.count, 500);
  assert.equal(grown.offset(11), axis.offset(11));
  assert.equal(grown.total, axis.total + 300 * 20);
  assert.equal(axis.grown(100, MAX_ROWS), axis, 'never shrinks');
  assert.equal(axis.grown(MAX_ROWS + 10, MAX_ROWS).count, MAX_ROWS);
});

test('rangeStats says what the status bar shows', () => {
  const at = (row, col) => row * 16_385 + col;
  const cells = new Map([
    [at(1, 1), { kind: 's', style: 0 }],
    [at(2, 1), { kind: 'n', style: 3, number: 10 }],
    [at(3, 1), { kind: 'n', style: 5, number: 20 }],
    [at(4, 1), { kind: '', style: 2 }],
    [at(5, 1), { kind: 'n', style: 5, number: 90 }],
    [at(2, 2), { kind: 'e', style: 0 }],
    [at(9, 9), { kind: 'n', style: 1, number: 1000 }],
  ]);
  const range = (a, b, c, d) => ({ first: { row: a, col: b }, last: { row: c, col: d } });
  assert.deepEqual(rangeStats(cells, range(1, 1, 3, 1)), { count: 3, numbers: 2, sum: 30, average: 15, min: 10, max: 20, style: 3 });
  // Text only, an error, nothing
  assert.equal(rangeStats(cells, range(1, 1, 1, 1)).sum, null);
  assert.deepEqual(rangeStats(cells, range(2, 1, 2, 2)), { count: 2, numbers: 0, sum: null, average: null, min: null, max: null, style: 3 });
  assert.equal(rangeStats(cells, range(20, 1, 30, 5)), null);
  // A range far larger than the sheet walks the cells of the sheet
  assert.equal(rangeStats(cells, range(1, 1, 1_048_576, 1)).sum, 120);
  // Hidden rows do not count
  const rows = new Axis(100, 20, new Map([[3, 0]]));
  assert.equal(rangeStats(cells, range(1, 1, 5, 1), { rows }).sum, 100);
});

test('pxToColWidth gives back the pixels', () => {
  for (const mdw of [6, 7, 8, 11]) {
    for (let px = 4; px < 1000; px += 7) assert.equal(colWidthToPx(pxToColWidth(px, mdw), mdw), px);
  }
});

test('dataEdge lands where Ctrl+arrow does', () => {
  const filled = new Set([3, 4, 5, 9]);
  const has = i => filled.has(i);
  assert.equal(dataEdge(has, 3, 1, 100), 5, 'to the end of the block');
  assert.equal(dataEdge(has, 5, 1, 100), 9, 'from its end to the next one');
  assert.equal(dataEdge(has, 1, 1, 100), 3, 'from empty to the next one');
  assert.equal(dataEdge(has, 9, 1, 100), 100, 'nothing else: the end of the sheet');
  assert.equal(dataEdge(has, 5, -1, 100), 3);
  assert.equal(dataEdge(has, 3, -1, 100), 1, 'nothing before: the start');
  assert.equal(dataEdge(has, 1, -1, 100), 1, 'already there');
  assert.equal(dataEdge(has, 100, 1, 100), 100);
});
