// Grid geometry: Excel's pixel sizes, offsets with custom and hidden rows, and the sheet extent.
// Run with: npm test -w packages/xlsx-editor
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Axis, colWidthToPx, columnName, defaultColPx, parseRange, parseRef, ptToPx, sheetAxes } from '../src/grid.js';

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
