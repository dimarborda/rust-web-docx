// Reading and writing cells through the public API. The engine's Rust tests cover templates
// saved by Excel and the byte-for-byte round trip; these check the JS side of the boundary.
// Run after `npm run build:wasm:xlsx` with: npm test -w packages/xlsx
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createWorkbook, openWorkbook } from '../src/index.js';

const reopen = async wb => openWorkbook(wb.toBytes());

test('values of every type survive a save', async () => {
  const wb = await createWorkbook();
  const sheet = wb.sheet('Hoja1');
  const day = new Date(Date.UTC(2024, 0, 15));
  const moment = new Date(Date.UTC(2024, 0, 15, 10, 30));
  sheet.setValues('A2', [
    ['Texto', 42.5, true, day, moment],
    ['Texto', -1, false, null, 'ñandú & <cía>'],
  ]);

  const again = (await reopen(wb)).sheet(0);
  assert.deepEqual(again.getValues('A2:E3'), [
    ['Texto', 42.5, true, day, moment],
    ['Texto', -1, false, null, 'ñandú & <cía>'],
  ]);
  assert.deepEqual(again.getCell('D2'), { ref: 'D2', type: 'date', value: day });
  assert.deepEqual(again.getCell('A1'), { ref: 'A1', type: 'empty', value: null });
  assert.equal(again.getValue('B2'), 42.5);
});

test('undefined leaves a cell as it is and null clears it', async () => {
  const wb = await createWorkbook();
  const sheet = wb.sheets[0];
  sheet.setValues('A1', [[1, 2, 3]]);
  sheet.setValues('A1', [[undefined, null, 'c']]);
  // eslint-disable-next-line no-sparse-arrays
  sheet.setValues('A2', [[, 'hueco']]);
  assert.deepEqual(sheet.getValues('A1:C2'), [
    [1, null, 'c'],
    [null, 'hueco', null],
  ]);
});

test('fillTemplate() replaces placeholders and counts them', async () => {
  const wb = await createWorkbook();
  wb.sheet(0).setValues('A1', [['Hola {{NOMBRE}}'], ['Total: {{ TOTAL }} {{MONEDA}}'], ['{{SIN_VALOR}}']]);
  const template = await reopen(wb);
  assert.equal(template.fillTemplate({ NOMBRE: 'Ana', TOTAL: 1500, MONEDA: 'EUR' }), 3);
  const filled = (await reopen(template)).sheet('hoja1');
  assert.deepEqual(filled.getValues('A1:A3'), [['Hola Ana'], ['Total: 1500 EUR'], ['{{SIN_VALOR}}']]);
  assert.throws(() => template.fillTemplate({ X: {} }), TypeError);
});

test('markers() finds fields and pending content in text cells', async () => {
  const wb = await createWorkbook();
  wb.sheet(0).setValues('A1', [['Cliente: {{ CLIENTE: razón social }}', 42], [null, 'Plazo [Completar] ver [1]']]);
  const template = await reopen(wb);
  assert.deepEqual(template.markers().map(m => [m.ref, m.row, m.col, m.kind, m.key, m.hint, m.offset]), [
    ['A1', 1, 1, 'field', 'CLIENTE', 'razón social', 9],
    ['B2', 2, 2, 'pending', null, 'Completar', 6],
  ]);
  assert.equal(template.markers({ kind: 'field' }).length, 1);
  assert.equal(template.fillTemplate({ CLIENTE: 'Acme' }), 1);
  assert.equal((await reopen(template)).sheet(0).getValue('A1'), 'Cliente: Acme');
});

test('formulas are written and recalculated after every edit', async () => {
  const wb = await createWorkbook();
  const sheet = wb.sheet(0);
  sheet.setValues('A1', [[2, 3, { formula: '=SUM(A1:B1)*10%' }, { formula: 'IF(C1>0.4,"alto","bajo")' }]]);
  assert.deepEqual(sheet.getValues('C1:D1'), [[0.5, 'alto']]);
  assert.deepEqual(sheet.getCell('C1'), { ref: 'C1', type: 'number', value: 0.5, formula: 'SUM(A1:B1)*10%' });

  wb.checkpoint();
  sheet.setValue('A1', 1);
  assert.deepEqual(sheet.getValues('C1:D1'), [[0.4, 'bajo']]);
  wb.undo();
  assert.equal(sheet.getValue('C1'), 0.5);

  const again = (await reopen(wb)).sheet(0);
  assert.deepEqual(again.getValues('C1:D1'), [[0.5, 'alto']]);
  assert.equal(wb.recalculate(), 0);
  assert.throws(() => sheet.setValue('E1', { formula: '=SUM(A1' }), /E1/);
});

test('invalid input is refused with a reason', async () => {
  const wb = await createWorkbook();
  const sheet = wb.sheet(0);
  assert.throws(() => sheet.setValue('A1', NaN), /finito/);
  assert.throws(() => sheet.setValue('A1', new Date('no')), /fecha/);
  assert.throws(() => sheet.setValue('A1', { a: 1 }), TypeError);
  assert.throws(() => sheet.setValues('A1', ['no es una fila']), TypeError);
  assert.throws(() => sheet.setValue('A0', 1), /referencia/);
  assert.throws(() => sheet.getValue('1A'), err => err instanceof Error && /referencia/.test(err.message));
  assert.throws(() => sheet.setValue('A1', new Date(Date.UTC(1800, 0, 1))), /1900/);
  assert.throws(() => wb.sheet('Otra'), /Otra/);
  assert.throws(() => wb.sheet(3), RangeError);
});

test('checkpoint(), undo() and redo() step through the edits', async () => {
  const wb = await createWorkbook();
  const sheet = wb.sheet(0);
  assert.equal(wb.canUndo, false);
  assert.equal(wb.undo(), null);
  wb.checkpoint('A1');
  sheet.setValues('A1', [['uno', 1]]);
  wb.setSelectionAfter('A2');
  wb.checkpoint('B5');
  sheet.setValue('B5', true);
  assert.equal(wb.canUndo, true);
  assert.equal(wb.undo().selection, 'B5');
  assert.equal(sheet.getValue('B5'), null);
  assert.equal(sheet.getValue('A1'), 'uno');
  const first = wb.undo();
  assert.equal(first.selection, 'A1');
  assert.ok(first.parts.length > 0);
  assert.deepEqual(sheet.getValues('A1:B1'), [[null, null]]);
  assert.equal(wb.canRedo, true);
  assert.equal(wb.redo().selection, 'A2');
  assert.deepEqual(sheet.getValues('A1:B1'), [['uno', 1]]);
  wb.checkpoint();
  sheet.setValue('C1', 3);
  assert.equal(wb.canRedo, false);
});
