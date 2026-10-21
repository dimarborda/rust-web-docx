// What typed text means, by language, and the clipboard's TSV.
// Run with: npm test -w packages/xlsx-editor
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { editText, numberText, parseInput, parseTSV, toTSV } from '../src/input.js';

const today = new Date(2026, 9, 17);
const parse = (text, locale) => parseInput(text, locale, { today });
const utc = (...args) => new Date(Date.UTC(...args));

test('numbers with the separators of the language', () => {
  assert.deepEqual(parse('1234,5', 'es'), { type: 'number', value: 1234.5 });
  assert.deepEqual(parse('1.234.567,25', 'es'), { type: 'number', value: 1234567.25 });
  assert.deepEqual(parse('1,234,567.25', 'en'), { type: 'number', value: 1234567.25 });
  assert.deepEqual(parse('1.5', 'en-US'), { type: 'number', value: 1.5 });
  assert.deepEqual(parse('-0,5', 'es'), { type: 'number', value: -0.5 });
  assert.deepEqual(parse(',5', 'es'), { type: 'number', value: 0.5 });
  assert.deepEqual(parse('+42', 'en'), { type: 'number', value: 42 });
  assert.deepEqual(parse('(300)', 'en'), { type: 'number', value: -300 });
  assert.deepEqual(parse('1,5E+3', 'es'), { type: 'number', value: 1500 });
  assert.deepEqual(parse('  007 ', 'en'), { type: 'number', value: 7 });
  assert.deepEqual(parse('$1,200.50', 'en'), { type: 'number', value: 1200.5 });
  assert.deepEqual(parse('-$5', 'en'), { type: 'number', value: -5 });
  assert.deepEqual(parse('1.234,50 €', 'es'), { type: 'number', value: 1234.5 });
  // Separators of the other language, or misplaced ones, are text
  assert.equal(parse('1,5', 'en').type, 'string');
  assert.equal(parse('1.23', 'es').type, 'string');
  assert.equal(parse('12,34,5', 'en').type, 'string');
  assert.equal(parse('-', 'en').type, 'string');
  assert.equal(parse('$', 'en').type, 'string');
  assert.equal(parse('12 abc', 'en').type, 'string');
});

test('percentages are divided by 100', () => {
  assert.deepEqual(parse('15%', 'es'), { type: 'percent', value: 0.15 });
  assert.deepEqual(parse('12,5 %', 'es'), { type: 'percent', value: 0.125 });
  assert.deepEqual(parse('-7.1%', 'en'), { type: 'percent', value: -0.071 });
  assert.equal(parse('%', 'en').type, 'string');
});

test('dates: day first in Spanish, month first in English', () => {
  assert.deepEqual(parse('15/03/2024', 'es'), { type: 'date', value: utc(2024, 2, 15) });
  assert.deepEqual(parse('3/15/2024', 'en'), { type: 'date', value: utc(2024, 2, 15) });
  assert.deepEqual(parse('03/04/2024', 'es').value, utc(2024, 3, 3));
  assert.deepEqual(parse('03/04/2024', 'en').value, utc(2024, 2, 4));
  assert.deepEqual(parse('15-3-24', 'es').value, utc(2024, 2, 15));
  assert.deepEqual(parse('15.3.2024', 'es').value, utc(2024, 2, 15));
  assert.deepEqual(parse('1/2/95', 'en').value, utc(1995, 0, 2));
  assert.deepEqual(parse('2024-03-15', 'es').value, utc(2024, 2, 15));
  // Without a year, this year
  assert.deepEqual(parse('15/3', 'es').value, utc(2026, 2, 15));
  // With a time
  assert.deepEqual(parse('15/03/2024 14:30', 'es').value, utc(2024, 2, 15, 14, 30));
  assert.deepEqual(parse('3/15/2024 2:30 PM', 'en').value, utc(2024, 2, 15, 14, 30));
  assert.deepEqual(parse('2024-03-15T08:05:09', 'en').value, utc(2024, 2, 15, 8, 5, 9));
  // Impossible dates are text
  assert.equal(parse('31/02/2024', 'es').type, 'string');
  assert.equal(parse('15/03/2024', 'en').type, 'string');
  assert.equal(parse('13/13/2024', 'es').type, 'string');
  assert.equal(parse('1/1/1850', 'es').type, 'string');
  assert.equal(parse('15/03/2024 25:00', 'es').type, 'string');
});

test('times without a date are fractions of a day', () => {
  assert.deepEqual(parse('12:00', 'es'), { type: 'time', value: 0.5 });
  assert.deepEqual(parse('6:00 pm', 'en'), { type: 'time', value: 0.75 });
  assert.equal(parse('0:00:30', 'es').value, 30 / 86400);
  assert.equal(parse('24:00', 'es').type, 'string');
  assert.equal(parse('13:00 PM', 'en').type, 'string');
});

test('booleans in the language of the sheet', () => {
  assert.deepEqual(parse('verdadero', 'es'), { type: 'boolean', value: true });
  assert.deepEqual(parse('FALSO', 'es-MX'), { type: 'boolean', value: false });
  assert.deepEqual(parse('True', 'en'), { type: 'boolean', value: true });
  assert.deepEqual(parse('false', 'en-GB'), { type: 'boolean', value: false });
  assert.equal(parse('TRUE', 'es').type, 'string');
  assert.equal(parse('VERDADERO', 'en').type, 'string');
});

test('text, formulas, apostrophes and empty input', () => {
  assert.deepEqual(parse('Hola mundo', 'es'), { type: 'string', value: 'Hola mundo' });
  assert.deepEqual(parse("'0012", 'es'), { type: 'string', value: '0012' });
  assert.deepEqual(parse("'15/03/2024", 'es'), { type: 'string', value: '15/03/2024' });
  assert.deepEqual(parse('=SUM(A1:A3)', 'es'), { type: 'formula', value: { formula: '=SUM(A1:A3)' } });
  assert.deepEqual(parse('=', 'es'), { type: 'string', value: '=' });
  assert.deepEqual(parse('', 'es'), { type: 'empty', value: null });
  assert.deepEqual(parse('   ', 'es'), { type: 'empty', value: null });
  assert.deepEqual(parse('línea 1\nlínea 2', 'es'), { type: 'string', value: 'línea 1\nlínea 2' });
});

test('the text to edit a cell reads back as the same value', () => {
  assert.equal(numberText(1234.5, 'es'), '1234,5');
  assert.equal(numberText(0.1 + 0.2, 'en'), '0.3');
  assert.equal(numberText(1.5e21, 'es'), '1,5E+21');
  const cases = [
    [{ type: 'number', value: -1234.56 }, ''],
    [{ type: 'number', value: 0.155 }, '15,5 %'],
    [{ type: 'boolean', value: true }, ''],
    [{ type: 'date', value: utc(2024, 2, 15) }, ''],
    [{ type: 'date', value: utc(2024, 2, 15, 9, 5, 7) }, ''],
    [{ type: 'string', value: '15/03/2024' }, ''],
    [{ type: 'string', value: '00123' }, ''],
    [{ type: 'string', value: 'texto' }, ''],
  ];
  for (const locale of ['es', 'en']) {
    for (const [cell, shown] of cases) {
      const back = parse(editText(cell, locale, shown), locale);
      const expected = cell.type === 'date' ? cell.value.getTime() : cell.value;
      const actual = back.value instanceof Date ? back.value.getTime() : back.value;
      assert.equal(actual, expected, `${locale}: ${JSON.stringify(cell)} → ${editText(cell, locale, shown)}`);
    }
  }
  assert.equal(editText({ type: 'date', value: utc(2024, 2, 15) }, 'es'), '15/03/2024');
  assert.equal(editText({ type: 'date', value: utc(2024, 2, 15, 14, 30) }, 'en'), '3/15/2024 14:30:00');
  assert.equal(editText({ type: 'date', value: utc(1899, 11, 31, 14, 30) }, 'es'), '14:30:00');
  assert.equal(editText({ type: 'number', value: 3, formula: 'A1+A2' }, 'es'), '=A1+A2');
  assert.equal(editText({ type: 'string', value: '00123' }, 'es'), "'00123");
  assert.equal(editText({ type: 'empty', value: null }, 'es'), '');
  assert.equal(editText({ type: 'error', value: '#DIV/0!' }, 'es'), '#DIV/0!');
});

test('TSV as Excel and Google Sheets copy it', () => {
  const rows = [['a', '1,5', ''], ['con\ttab', 'dos\nlíneas', 'di "hola"']];
  const tsv = toTSV(rows);
  assert.equal(tsv, 'a\t1,5\t\r\n"con\ttab"\t"dos\nlíneas"\t"di ""hola"""');
  assert.deepEqual(parseTSV(tsv), rows);
  assert.deepEqual(parseTSV('a\tb\r\nc\td\r\n'), [['a', 'b'], ['c', 'd']]);
  assert.deepEqual(parseTSV('solo'), [['solo']]);
  assert.deepEqual(parseTSV('\t\n\t'), [['', ''], ['', '']]);
  assert.deepEqual(parseTSV('"abierto\tsin cerrar'), [['"abierto', 'sin cerrar']]);
  assert.deepEqual(parseTSV('5" pantalla\t"x"'), [['5" pantalla', 'x']]);
});
