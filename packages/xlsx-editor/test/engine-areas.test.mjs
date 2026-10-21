// Areas for agents and basic formatting through the public API.
// Run after `npm run build:wasm:xlsx` with: npm test -w packages/xlsx-editor
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createWorkbook, openWorkbook } from '../src/engine.js';

async function report() {
  const wb = await createWorkbook();
  const sheet = wb.sheet(0);
  sheet.setValues('A1', [
    ['Ventas del trimestre'],
    ['Región', 'Pedidos', 'Importe'],
    ['Norte', 3, 120.5],
    ['Sur', 5, 80],
  ]);
  sheet.setValue('C5', { formula: 'SUM(C3:C4)' });
  sheet.setValues('F1', [['Notas'], ['revisar']]);
  return wb;
}

test('areas() finds blocks with their title and header', async () => {
  const wb = await report();
  const [sales, notes] = wb.areas();
  assert.equal(sales.kind, 'region');
  assert.equal(sales.address, 'Hoja1!A1:C5');
  assert.equal(sales.title, 'Ventas del trimestre');
  assert.equal(sales.headerRow, 2);
  assert.equal(sales.body, 'A3:C5');
  assert.deepEqual(sales.columns, ['Región', 'Pedidos', 'Importe']);
  assert.equal(notes.ref, 'F1:F2');
  assert.equal(wb.sheet(0).regionAt('B4').ref, 'A1:C5');
  assert.deepEqual(wb.areas({ sheet: 'Hoja1' }).length, 2);
});

test('area() gives the context an agent needs', async () => {
  const wb = await report();
  const context = wb.area('Hoja1!A1:C5', { maxRows: 2 });
  assert.equal(context.rowCount, 3);
  assert.equal(context.truncated, true);
  assert.deepEqual(context.rows, [['Norte', '3', '120.5'], ['Sur', '5', '80']]);
  const importe = context.columnsInfo[2];
  assert.equal(importe.type, 'number');
  assert.equal(importe.numbers.sum, 401);
  assert.equal(importe.exampleFormula, 'C5: =SUM(C3:C4)');
  assert.match(context.markdown, /^\| fila \| Región \(A\) \| Pedidos \(B\) \| Importe \(C\) \|/);
  assert.equal(wb.sheet(0).area('A1:C5').address, 'Hoja1!A1:C5');
  assert.throws(() => wb.area('A1:C5'), /hoja/);
});

test('setFormat() changes the look and keeps values', async () => {
  const wb = await report();
  const sheet = wb.sheet(0);
  sheet.setFormat('A2:C2', { bold: true, fill: '#DDEBF7', border: { style: 'thin', edges: 'outline' }, align: 'center' });
  sheet.setFormat('C3:C5', { numberFormat: '#,##0.00', color: '#C00000' });
  sheet.setColumnWidth('A:C', 14);
  sheet.mergeCells('A1:C1');

  const again = (await openWorkbook(wb.toBytes())).sheet(0);
  const xfs = (await openWorkbook(wb.toBytes())).styleView().xfs;
  const cells = new Map(again.cellViews('en').map(([row, col, style, text]) => [`${row},${col}`, { style, text }]));
  assert.equal(cells.get('5,3').text, '200.50');
  assert.equal(xfs[cells.get('2,1').style].h, 'center');
  assert.deepEqual(again.layout().merges, ['A1:C1']);
  assert.equal(again.layout().cols[0][2], 14);
  assert.equal(again.getValue('C5'), 200.5);
  assert.throws(() => sheet.setFormat('A1', { negrita: true }), /negrita/);
  assert.throws(() => sheet.mergeCells('A3:B3'), /B3/);
  assert.equal(sheet.unmergeCells('A1'), 1);
});
