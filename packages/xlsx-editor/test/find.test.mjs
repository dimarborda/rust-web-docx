// Searching texts as Excel does. Run with: npm test -w packages/xlsx-editor
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { matches, replaceMatches, searchPattern } from '../src/find.js';

test('searchPattern understands wildcards and options', () => {
  assert.equal(searchPattern(''), null);
  const has = (query, text, options) => matches(searchPattern(query, options), text);
  assert.ok(has('ven', 'Ventas'));
  assert.ok(!has('ven', 'Ventas', { matchCase: true }));
  assert.ok(has('ventas', 'Ventas', { wholeCell: true }));
  assert.ok(!has('venta', 'Ventas', { wholeCell: true }));
  assert.ok(has('v*s', 'Ventas', { wholeCell: true }));
  assert.ok(has('v?nt', 'Ventas'));
  assert.ok(!has('v?t', 'Ventas'));
  assert.ok(has('~*', 'a*b'));
  assert.ok(!has('~*', 'ab'));
  assert.ok(has('~?', '¿qué?'));
  assert.ok(has('a~~b', 'a~b'));
  assert.ok(has('1.5', '1.5'));
  assert.ok(!has('1.5', '125'));
  assert.ok(has('(x)$[', 'a(x)$[b'));
  // The same expression gives the same answer every time
  const pattern = searchPattern('a');
  assert.ok(matches(pattern, 'a'));
  assert.ok(matches(pattern, 'a'));
});

test('replaceMatches takes the replacement literally', () => {
  assert.equal(replaceMatches(searchPattern('a'), 'banana', 'o'), 'bonono');
  assert.equal(replaceMatches(searchPattern('A'), 'banana', '$&$1'), 'b$&$1n$&$1n$&$1');
  assert.equal(replaceMatches(searchPattern('A', { matchCase: true }), 'banana', 'o'), 'banana');
  assert.equal(replaceMatches(searchPattern('*'), 'todo', 'x'), 'x', 'once, not again at the end');
  assert.equal(replaceMatches(searchPattern('A1'), '=A1+A10', 'B2'), '=B2+B20');
});
