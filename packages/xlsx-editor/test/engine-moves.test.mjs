// Moving rows and columns through the public API.
// Run after `npm run build:wasm:xlsx` with: npm test -w packages/xlsx-editor
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createWorkbook, openWorkbook } from '../src/engine.js';

async function sales() {
  const wb = await createWorkbook();
  const sheet = wb.sheet(0);
  sheet.setValues('A1', [
    ['Producto', 'Precio', 'Doble'],
    ['uno', 1, { formula: 'B2*2' }],
    ['dos', 2, { formula: 'B3*2' }],
    ['tres', 3, { formula: 'B4*2' }],
    ['Total', { formula: 'SUM(B2:B4)' }, { formula: 'C2+C4' }],
  ]);
  return wb;
}

test('moveRows() follows formulas and keeps ranges', async () => {
  const wb = await sales();
  const sheet = wb.sheet(0);
  assert.equal(sheet.moveRows(4, 1, 2), 2);
  assert.deepEqual(sheet.getValues('A2:A4').flat(), ['tres', 'uno', 'dos']);
  assert.equal(sheet.getCell('C2').formula, 'B2*2');
  assert.equal(sheet.getCell('B5').formula, 'SUM(B2:B4)');
  assert.equal(sheet.getValue('B5'), 6);
  assert.equal(sheet.getCell('C5').formula, 'C3+C2');
  assert.equal(sheet.getValue('C5'), 8);
  const again = (await openWorkbook(wb.toBytes())).sheet(0);
  assert.equal(again.getValue('A2'), 'tres');
});

test('moveColumns() takes letters and goes in one undo step', async () => {
  const wb = await sales();
  const sheet = wb.sheet(0);
  wb.checkpoint('antes');
  assert.equal(sheet.moveColumns('C', 1, 'A'), 1);
  assert.deepEqual(sheet.getValues('A1:C1').flat(), ['Doble', 'Producto', 'Precio']);
  assert.equal(sheet.getCell('A2').formula, 'C2*2');
  assert.equal(sheet.getCell('C5').formula, 'SUM(C2:C4)');
  assert.equal(wb.undo().selection, 'antes');
  assert.equal(sheet.getValue('A1'), 'Producto');
});

test('a move that would tear something is refused and changes nothing', async () => {
  const wb = await sales();
  const sheet = wb.sheet(0);
  sheet.mergeCells('A7:B8');
  const before = wb.toBytes();
  wb.checkpoint();
  assert.throws(() => sheet.moveRows(7, 1, 2), /combinadas A7:B8/);
  assert.equal(wb.canUndo, false, 'a refused move leaves no undo step');
  assert.deepEqual(wb.toBytes(), before);
  assert.throws(() => sheet.moveRows(-1, 1, 2), TypeError);
  assert.throws(() => sheet.moveRows(1, 0, 2), /al menos una fila/);
});

test('formatNumbers() and cellViews() give what the status bar needs', async () => {
  const wb = await sales();
  const sheet = wb.sheet(0);
  sheet.setFormat('B2', { numberFormat: '#,##0.00' });
  const views = sheet.cellViews('es');
  const b2 = views.find(([row, col]) => row === 2 && col === 2);
  assert.equal(b2[6], 1);
  assert.equal(views.find(([row, col]) => row === 1 && col === 1)[6], undefined);
  assert.deepEqual(wb.formatNumbers(b2[2], [1234.5, 2], 'es'), ['1.234,50', '2,00']);
});

test('insertRows() and deleteRows() update formulas in one undo step', async () => {
  const wb = await sales();
  const sheet = wb.sheet(0);
  wb.checkpoint('antes');
  assert.equal(sheet.insertRows(3, 2), 3);
  assert.equal(sheet.getValue('A5'), 'dos');
  assert.equal(sheet.getValue('A3'), null);
  assert.equal(sheet.getCell('B7').formula, 'SUM(B2:B6)');
  assert.equal(sheet.getValue('B7'), 6);
  wb.undo();
  assert.equal(sheet.getValue('A3'), 'dos');

  assert.equal(sheet.deleteRows(3), 3);
  assert.equal(sheet.getCell('B4').formula, 'SUM(B2:B3)');
  assert.equal(sheet.getValue('B4'), 4);
  sheet.setValue('E1', { formula: 'A2&"!"' });
  assert.equal(sheet.deleteRows(2), 2);
  assert.equal(sheet.getCell('E1').formula, '#REF!&"!"');
  assert.equal(sheet.getValue('E1'), '#REF!');
  assert.throws(() => sheet.insertRows(1, 0), /al menos una fila/);
});

test('insertColumns(), deleteColumns() and setRowHeight()', async () => {
  const wb = await sales();
  const sheet = wb.sheet(0);
  assert.equal(sheet.insertColumns('B'), 2);
  assert.equal(sheet.getValue('C1'), 'Precio');
  assert.equal(sheet.getCell('C5').formula, 'SUM(C2:C4)');
  sheet.deleteColumns('A', 2);
  assert.equal(sheet.getValue('A1'), 'Precio');
  assert.equal(sheet.getCell('B2').formula, 'A2*2');

  sheet.setRowHeight('2:3', 30);
  sheet.setRowHeight(4, 12.5);
  assert.deepEqual(sheet.layout().rows.map(([row, height]) => [row, height]), [[2, 30], [3, 30], [4, 12.5]]);
  sheet.setRowHeight(4, null);
  assert.deepEqual(sheet.layout().rows.map(([row]) => row), [2, 3]);
  assert.throws(() => sheet.setRowHeight(1, 'alto'), TypeError);
  assert.equal(sheet.getValue('A2'), 1);
});

test('fill() continues series and moves formulas', async () => {
  const wb = await createWorkbook();
  const sheet = wb.sheet(0);
  sheet.setValues('A1', [[1, 'Lote 1', 'enero', { formula: 'A1*2' }], [2]]);
  sheet.fill('A1:A2', 'A1:A4');
  assert.deepEqual(sheet.getValues('A3:A4').flat(), [3, 4]);
  sheet.fill('B1:D1', 'B1:D3');
  assert.deepEqual(sheet.getValues('B2:D2')[0], ['Lote 2', 'febrero', 4]);
  assert.equal(sheet.getCell('D3').formula, 'A3*2');
  sheet.fill('A1:A2', 'A1:A4', { series: false });
  assert.deepEqual(sheet.getValues('A3:A4').flat(), [1, 2]);
  assert.throws(() => sheet.fill('A1', 'B2'), /extremo/);
  assert.equal(sheet.getCells('D3:D3')[0][0].formula, 'A3*2');
});

test('setRowsHidden() and setColumnsHidden() keep cells and sizes', async () => {
  const wb = await sales();
  const sheet = wb.sheet(0);
  sheet.setColumnWidth('B', 20);
  sheet.setColumnsHidden('B', true);
  sheet.setRowsHidden('2:3', true);
  assert.deepEqual(sheet.layout().cols, [[2, 2, 20, true]]);
  assert.deepEqual(sheet.layout().rows.map(([row, , hidden]) => [row, hidden]), [[2, true], [3, true]]);
  assert.equal(sheet.getValue('B5'), 6);
  sheet.setColumnsHidden('B', false);
  sheet.setRowsHidden(2, false);
  assert.deepEqual(sheet.layout().cols, [[2, 2, 20, false]]);
  assert.deepEqual(sheet.layout().rows.map(([row]) => row), [3]);
});
