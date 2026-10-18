// Public API on Node: the engine loads from the package and workbooks round-trip.
// Run after `npm run build:wasm:xlsx` with: npm test -w packages/xlsx
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createWorkbook, openWorkbook } from '../src/index.js';

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
