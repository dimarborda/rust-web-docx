// Marker lexicon ({{KEY}}, {key}, [ … ]) and activity marks: the engine's scan_markers_json
// and the helpers of marker_api.js and activity.js that DocxEditor calls.
// Run after `npm run build:wasm` with: npm test -w packages/docx-editor
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { initSync, DocxSession, scan_markers_json } from '../wasm/docx_engine.js';
import { collectMarkers, markerAfter, relocateMarker, variablePairs } from '../src/marker_api.js';
import { highlightKinds, markerBoxes } from '../src/marker_layer.js';
import { normalizeTarget, remapTarget, targetSpan, textDiff } from '../src/activity.js';

initSync({ module: readFileSync(new URL('../wasm/docx_engine_bg.wasm', import.meta.url)) });

const paragraphs = texts => texts.map((text, index) => ({ index, text, container: 'body' }));

test('the engine scans several texts in one call', () => {
  const found = JSON.parse(scan_markers_json(JSON.stringify(['Hola {{CLIENTE}}', 'nada', 'Ver [1] y [Completar: plazo]'])));
  assert.deepEqual(found, [
    { item: 0, kind: 'field', syntax: 'double', raw: '{{CLIENTE}}', key: 'CLIENTE', hint: null, offset: 5, length: 11 },
    { item: 2, kind: 'pending', syntax: 'bracket', raw: '[Completar: plazo]', key: null, hint: 'Completar: plazo', offset: 10, length: 18 },
  ]);
});

test('markers carry paragraph, place and section', () => {
  const list = [
    ...paragraphs(['Contrato', 'Entre {{ CLIENTE }} y {proveedor}.', 'Objeto: [...]']),
    { index: 3, text: 'Celda {{VALOR: moneda}}', container: '0:1:2' },
  ];
  const sections = [{ key: 'objeto', first: 2, last: 3, depth: 0 }];
  const markers = collectMarkers(list, sections, scan_markers_json);
  assert.deepEqual(markers.map(m => [m.id, m.kind, m.key, m.hint, m.place, m.section]), [
    ['1:6', 'field', 'CLIENTE', null, 'body', null],
    ['1:22', 'field', 'proveedor', null, 'body', null],
    ['2:8', 'pending', null, null, 'body', 'objeto'],
    ['3:6', 'field', 'VALOR', 'moneda', 'table', 'objeto'],
  ]);
  assert.equal(collectMarkers(list, sections, scan_markers_json, { kind: 'pending' }).length, 1);
});

test('a marker is found again after text was typed before it', () => {
  const marker = { paragraph: 0, offset: 5, raw: '{{X}}' };
  assert.deepEqual(relocateMarker(marker, 'Hola {{X}}'), { offset: 5, length: 5 });
  assert.deepEqual(relocateMarker(marker, 'Hola señor {{X}}'), { offset: 11, length: 5 });
  assert.equal(relocateMarker(marker, 'Hola Ana'), null);
});

test('a bare key fills every spelling of the field, longest first', () => {
  const markers = [
    { kind: 'field', key: 'CLIENTE', raw: '{CLIENTE}' },
    { kind: 'field', key: 'CLIENTE', raw: '{{ CLIENTE: nombre }}' },
    { kind: 'field', key: 'OTRO', raw: '{{OTRO}}' },
  ];
  assert.deepEqual(variablePairs({ CLIENTE: 'Acme', '{x}': 1 }, markers).map(p => p.key), [
    '{{ CLIENTE: nombre }}', '{{CLIENTE}}', '{CLIENTE}', '{x}',
  ]);
});

test('replaceVariables pairs fill a real document without breaking {{ }}', () => {
  const s = DocxSession.new_blank('letter');
  s.insert_paragraphs(0, 0, JSON.stringify([{ text: 'A {{CLIENTE}} B {CLIENTE} C {{ CLIENTE: nombre }}' }]), undefined);
  const list = JSON.parse(s.get_document_elements_json()).filter(e => e.type === 'paragraph').map(e => ({ ...e, container: 'body' }));
  const pairs = variablePairs({ CLIENTE: 'Acme' }, collectMarkers(list, [], scan_markers_json));
  s.batch_replace(JSON.stringify(pairs));
  assert.match(s.get_raw_text(), /A Acme B Acme C Acme/);
});

test('activity targets', () => {
  assert.deepEqual(normalizeTarget('intro'), { section: 'intro' });
  assert.deepEqual(normalizeTarget({ first: 4, count: 3 }), { from: 4, to: 6 });
  assert.deepEqual(normalizeTarget({ paragraph: 2, offset: 3, length: 5 }), { paragraph: 2, offset: 3, length: 5 });
  assert.deepEqual(normalizeTarget({ paragraph: 2 }), { from: 2, to: 2 });
  assert.throws(() => normalizeTarget(null));
  const len = i => ['abc', 'defg', ''][i] ? ['abc', 'defg', ''][i].length : 0;
  assert.deepEqual(targetSpan({ from: 0, to: 1 }, [], len), { start: { paragraph: 0, offset: 0 }, end: { paragraph: 1, offset: 4 } });
  assert.equal(targetSpan({ section: 'gone' }, [], len), null);
});

test('activity ranges follow paragraphs inserted or removed before them', () => {
  const before = ['a', 'b', 'c', 'd'];
  // Two paragraphs typed after "a": the mark on "c"–"d" moves down
  const inserted = textDiff(before, ['a', 'x', 'y', 'b', 'c', 'd']);
  assert.deepEqual(remapTarget({ from: 2, to: 3 }, inserted), { from: 4, to: 5 });
  assert.deepEqual(remapTarget({ from: 0, to: 0 }, inserted), { from: 0, to: 0 });
  // A paragraph replaced by three inside the mark: the mark grows
  const grown = textDiff(before, ['a', 'b1', 'b2', 'b3', 'c', 'd']);
  assert.deepEqual(remapTarget({ from: 1, to: 2 }, grown), { from: 1, to: 4 });
  // "a" deleted: everything moves up
  assert.deepEqual(remapTarget({ paragraph: 2, offset: 1, length: 2 }, textDiff(before, ['b', 'c', 'd'])), { paragraph: 1, offset: 1, length: 2 });
  assert.deepEqual(remapTarget({ section: 's' }, inserted), { section: 's' });
});

test('highlightPending picks the kinds of marker to paint', () => {
  assert.deepEqual(highlightKinds(true), ['pending']);
  assert.deepEqual(highlightKinds({ fields: true }), ['pending', 'field']);
  assert.deepEqual(highlightKinds({}), ['pending']);
  assert.deepEqual(highlightKinds(false), []);
  assert.deepEqual(highlightKinds(undefined), []);
});

test('markerBoxes paints each line of each marker, and nothing once it is gone', () => {
  const markers = collectMarkers(paragraphs(['Plazo: [Completar: plazo] días', 'Cliente {{CLIENTE}}', 'Sin nada']), [], scan_markers_json);
  const calls = [];
  // A marker that wraps onto the next line gives two rects; empty rects are skipped
  const rects = (start, end) => {
    calls.push([start, end]);
    return start.paragraph === 0
      ? [{ page: 1, x: 90, y: 100, width: 60, height: 14 }, { page: 1, x: 72, y: 114, width: 30, height: 14 }]
      : [{ page: 2, x: 120, y: 40, width: 0, height: 14 }, { page: 2, x: 125, y: 40, width: 70, height: 14 }];
  };
  assert.deepEqual(markerBoxes(markers, rects).map(b => [b.kind, b.page, b.x]), [
    ['pending', 1, 90],
    ['pending', 1, 72],
    ['field', 2, 125],
  ]);
  assert.deepEqual(calls, [
    [{ paragraph: 0, offset: 7 }, { paragraph: 0, offset: 25 }],
    [{ paragraph: 1, offset: 8 }, { paragraph: 1, offset: 19 }],
  ]);
  // Filled: the scan finds no marker, so there is nothing left to paint
  const filled = collectMarkers(paragraphs(['Plazo: 30 días', 'Cliente Acme']), [], scan_markers_json);
  assert.deepEqual(markerBoxes(filled, rects), []);
});

test('markerAfter walks the markers in document order', () => {
  const markers = collectMarkers(paragraphs(['[…] y [Completar: a]', 'nada', 'Fin: [...]']), [], scan_markers_json);
  const at = m => m && [m.paragraph, m.offset];
  assert.deepEqual(at(markerAfter(markers, null)), [0, 0]);
  // A marker starting exactly at `from` does not count: calling again moves on
  assert.deepEqual(at(markerAfter(markers, { paragraph: 0, offset: 0 })), [0, 6]);
  assert.deepEqual(at(markerAfter(markers, { paragraph: 0, offset: 6 })), [2, 5]);
  assert.deepEqual(at(markerAfter(markers, { paragraph: 1, offset: 2 })), [2, 5]);
  assert.equal(markerAfter(markers, { paragraph: 2, offset: 5 }), null);
  assert.deepEqual(at(markerAfter(markers, { paragraph: 2, offset: 5 }, { wrap: true })), [0, 0]);
  assert.equal(markerAfter([], null, { wrap: true }), null);
});
