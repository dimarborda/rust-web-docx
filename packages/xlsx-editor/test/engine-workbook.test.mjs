// Public API on Node: the engine loads from the package and workbooks round-trip.
// Run after `npm run build:wasm:xlsx` with: npm test -w packages/xlsx-editor
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createWorkbook, openWorkbook } from '../src/engine.js';

test('createWorkbook() has one sheet', async () => {
  const wb = await createWorkbook();
  assert.deepEqual(wb.sheetNames, ['Hoja1']);
  wb.close();
});

test('a saved workbook opens again with the same sheets', async () => {
  const bytes = (await createWorkbook()).toBytes();
  assert.ok(bytes instanceof Uint8Array && bytes.length > 0);
  const reopened = await openWorkbook(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
  assert.deepEqual(reopened.sheetNames, ['Hoja1']);
});

test('a file that is not a workbook is refused with a reason', async () => {
  await assert.rejects(openWorkbook(new Uint8Array([1, 2, 3, 4, 5])), /XLSX|ZIP/);
});

test('addSheet() and renameSheet() keep formulas pointing at the right sheet', async () => {
  const wb = await createWorkbook();
  const datos = wb.addSheet('Datos');
  assert.equal(datos.name, 'Datos');
  wb.addSheet('Portada', 'Hoja1');
  assert.deepEqual(wb.sheetNames, ['Hoja1', 'Portada', 'Datos']);
  wb.sheet('Datos').setValues('A1', [[4]]);
  wb.sheet('Hoja1').setValues('A1', [[{ formula: 'Datos!A1*2' }]]);
  assert.equal(wb.renameSheet('datos', 'Precios 2026'), 1);
  assert.equal(wb.sheet(0).getCell('A1').formula, "'Precios 2026'!A1*2");
  assert.throws(() => wb.addSheet('portada'), /Ya hay una hoja/);
  assert.throws(() => wb.renameSheet(0, 'a/b'), /no puede llevar/);
  const reopened = await openWorkbook(wb.toBytes());
  assert.deepEqual(reopened.sheetNames, ['Hoja1', 'Portada', 'Precios 2026']);
});
