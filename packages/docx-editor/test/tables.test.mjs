// Existing-table API: the engine bindings DocxEditor calls (tables_json, set_table_cell,
// insert_table_rows, delete_table_rows) and the argument checks of table_api.js.
// Run after `npm run build:wasm` with: npm test -w packages/docx-editor
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { initSync, DocxSession } from '../wasm/docx_engine.js';
import { checkIndex, engineCall, tableRowsPayload } from '../src/table_api.js';

initSync({ module: readFileSync(new URL('../wasm/docx_engine_bg.wasm', import.meta.url)) });

/** Sample contract: one table whose first row is a shaded header */
const sample = () => DocxSession.new_sample();
const tables = s => JSON.parse(s.tables_json());
const cells = (s, table) => JSON.parse(s.get_tables_json())[table].rows;
const reopen = s => {
  // The engine's own reader: reopen the saved bytes and read the table back
  const reopened = new DocxSession(s.export_bytes());
  return { tables: tables(reopened), rows: cells(reopened, 0) };
};

test('tables() lists the body tables with header and paragraph range', () => {
  const s = sample();
  const [t] = tables(s);
  assert.equal(t.index, 0);
  assert.ok(t.rows >= 2 && t.cols >= 2);
  assert.equal(t.header.length, t.cols);
  assert.deepEqual(t.header, cells(s, 0)[0]);
  const paragraphs = JSON.parse(s.get_paragraphs_json()).map(p => p.index);
  assert.ok(t.first <= t.last && !paragraphs.includes(t.first), 'cell paragraphs are not body paragraphs');
});

test('setTableCell writes a cell as one undo step', () => {
  const s = sample();
  const before = cells(s, 0);
  s.set_table_cell(0, 1, 1, 'Nuevo\nvalor', undefined);
  assert.equal(cells(s, 0)[1][1], 'Nuevo\nvalor');
  s.undo();
  assert.deepEqual(cells(s, 0), before);
});

test('insertTableRows at the start, in the middle and at the end', () => {
  const s = sample();
  const { rows, cols } = tables(s)[0];
  const row = text => [text, ...Array(cols - 1).fill('')];
  assert.deepEqual(JSON.parse(s.insert_table_rows(0, 1, JSON.stringify([row('medio')]), undefined, undefined)), { first: 1, count: 1 });
  s.insert_table_rows(0, rows + 1, JSON.stringify([row('fin 1'), ['fin 2']]), undefined, undefined);
  s.insert_table_rows(0, 0, JSON.stringify([row('inicio')]), undefined, undefined);
  const after = cells(s, 0);
  assert.equal(after.length, rows + 4);
  assert.equal(after[0][0], 'inicio');
  assert.equal(after[2][0], 'medio');
  assert.deepEqual(after.at(-1), ['fin 2', ...Array(cols - 1).fill('')], 'short rows are completed');
  // The saved file reopens with the new rows
  assert.equal(reopen(s).rows.length, rows + 4);
});

test('insertTableRows copies the formatting of the template row', () => {
  const s = sample();
  const rich = () => JSON.parse(s.get_tables_json())[0].rich_rows;
  const header = rich()[0].cells[0];
  s.insert_table_rows(0, 1, JSON.stringify([['copia del encabezado']]), 0, undefined);
  const copy = rich()[1].cells[0];
  assert.equal(copy.bg_color, header.bg_color);
  assert.equal(copy.bold, header.bold);
  assert.equal(copy.font_size, header.font_size);
  assert.equal(copy.color, header.color);
});

test('deleteTableRows removes rows and is undone in one step', () => {
  const s = sample();
  const before = cells(s, 0);
  s.delete_table_rows(0, 1, 1, undefined);
  assert.equal(cells(s, 0).length, before.length - 1);
  assert.deepEqual(cells(s, 0)[1], before[2]);
  s.undo();
  assert.deepEqual(cells(s, 0), before);
  assert.throws(() => s.delete_table_rows(0, 0, before.length, undefined), /sin filas/);
});

test('each insert is one undo step even with several rows', () => {
  const s = sample();
  const before = cells(s, 0);
  s.insert_table_rows(0, 1, JSON.stringify([['a'], ['b'], ['c']]), undefined, undefined);
  s.undo();
  assert.deepEqual(cells(s, 0), before);
  assert.equal(s.can_undo(), false);
});

test('invalid indices throw clear errors without changing the document', () => {
  const s = sample();
  const bytes = () => s.export_bytes().length;
  const { rows, cols } = tables(s)[0];
  const size = bytes();
  assert.throws(() => s.set_table_cell(3, 0, 0, 'x', undefined), /No existe la tabla 3/);
  assert.throws(() => s.set_table_cell(0, rows, 0, 'x', undefined), /No existe la fila/);
  assert.throws(() => s.set_table_cell(0, 0, cols, 'x', undefined), /No existe la columna/);
  assert.throws(() => s.insert_table_rows(0, rows + 1, '[["x"]]', undefined, undefined), /posición/);
  assert.throws(() => s.insert_table_rows(0, 1, JSON.stringify([Array(cols + 1).fill('x')]), undefined, undefined), /valores/);
  assert.throws(() => s.delete_table_rows(0, rows, 1, undefined), /No se pueden eliminar/);
  assert.equal(bytes(), size);
  assert.equal(s.can_undo(), false, 'failed calls leave no undo step');
});

test('argument checks of the JS layer', () => {
  assert.throws(() => checkIndex('setTableCell', { table: 0, row: -1 }), /`row` must be an integer ≥ 0 \(got -1\)/);
  assert.throws(() => checkIndex('deleteTableRows', { count: 1.5 }), /`count`/);
  assert.throws(() => checkIndex('insertTableRows', { at: '1' }), /`at`/);
  assert.doesNotThrow(() => checkIndex('insertTableRows', { table: 0, at: 3 }));
  assert.throws(() => tableRowsPayload([]), /at least one row/);
  assert.throws(() => tableRowsPayload(['a']), /row 0 must be an array/);
  assert.deepEqual(tableRowsPayload([['a', null, 3, 'x\r\ny']]), [['a', '', '3', 'x\ny']]);
  // Engine errors arrive as strings; the API throws them as Error
  const s = sample();
  assert.throws(() => engineCall(() => s.set_table_cell(7, 0, 0, 'x', undefined)), err => err instanceof Error && /^docx-editor: No existe la tabla 7/.test(err.message));
});
