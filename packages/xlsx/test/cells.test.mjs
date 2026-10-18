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
