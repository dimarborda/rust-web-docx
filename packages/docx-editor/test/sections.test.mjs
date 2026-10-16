// Logical sections: the engine bindings DocxEditor calls (sections_json, define_section,
// replace_section, unwrap_section) and the helpers of section_api.js.
// Run after `npm run build:wasm` with: npm test -w packages/docx-editor
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { initSync, DocxSession } from '../wasm/docx_engine.js';
import { checkSectionKey, sectionContent } from '../src/section_api.js';

initSync({ module: readFileSync(new URL('../wasm/docx_engine_bg.wasm', import.meta.url)) });

const sections = s => JSON.parse(s.sections_json());
const elements = s => JSON.parse(s.get_document_elements_json());
const paragraphs = s => JSON.parse(s.get_paragraphs_json());
/** Every paragraph text by index, table cells included */
const allTexts = s => {
  const out = [];
  for (const el of elements(s)) {
    if (el.type === 'paragraph') out[el.index] = el.text;
    else for (const row of el.rich_rows) for (const cell of row.cells) for (const p of cell.paragraphs) out[p.index] = p.text;
  }
  return out;
};
/** Minimal structure() blocks, as DocxEditor builds them */
const blocks = s => elements(s).map(el => {
  if (el.type === 'paragraph') return { kind: 'paragraph', paragraph: el.index };
  const idx = el.rich_rows.flatMap(r => r.cells.flatMap(c => c.paragraphs.map(p => p.index)));
  return { kind: 'table', first: Math.min(...idx), last: Math.max(...idx) };
});

/** Blank document with five paragraphs */
function report() {
  const s = DocxSession.new_blank('letter');
  s.insert_paragraphs(0, 0, JSON.stringify(['Informe', '1. Introducción', 'Objeto.', '2. Conclusiones', 'Fin.'].map(text => ({ text }))), undefined);
  return s;
}

test('a blank document has no sections', () => {
  assert.deepEqual(sections(DocxSession.new_blank('letter')), []);
});

test('defineSection marks paragraphs and the saved file keeps the section', () => {
  const s = report();
  const before = allTexts(s);
  const first = before.indexOf('1. Introducción');
  const info = JSON.parse(s.define_section('introduccion', first, first + 1, 'Introducción', true, undefined));
  assert.deepEqual(info, { key: 'introduccion', title: 'Introducción', first, last: first + 1, depth: 0, parent: null, locked: true });
  assert.deepEqual(allTexts(s), before, 'the text does not change');
  assert.deepEqual(sections(new DocxSession(s.export_bytes())), [info]);
  s.undo();
  assert.deepEqual(sections(s), []);
});

test('sections nest and sectionContent reads blocks and text', () => {
  const s = report();
  const t = allTexts(s);
  const intro = t.indexOf('1. Introducción');
  const end = t.indexOf('Fin.');
  s.define_section('cuerpo', intro, end, undefined, false, undefined);
  s.define_section('conclusiones', end - 1, end, 'Conclusiones', false, undefined);
  const [cuerpo, conclusiones] = sections(s);
  assert.equal(cuerpo.title, 'cuerpo', 'the title defaults to the key');
  assert.deepEqual([conclusiones.depth, conclusiones.parent], [1, 'cuerpo']);
  const content = sectionContent(conclusiones, blocks(s), elements(s));
  assert.equal(content.text, '2. Conclusiones\nFin.');
  assert.deepEqual(content.blocks.map(b => b.paragraph), [end - 1, end]);
});

test('replaceSection rewrites the content and keeps the section', () => {
  const s = report();
  const t = allTexts(s);
  const from = t.indexOf('2. Conclusiones');
  s.define_section('conclusiones', from, from + 1, 'Conclusiones', false, undefined);
  const result = JSON.parse(s.replace_section('conclusiones', JSON.stringify([{ text: 'Uno.' }, { text: 'Dos.' }, { text: 'Tres.' }]), false, undefined));
  assert.deepEqual([result.first, result.count], [from, 3]);
  const [section] = sections(s);
  assert.deepEqual([section.first, section.last, section.title], [from, from + 2, 'Conclusiones']);
  assert.deepEqual(allTexts(s).slice(from, from + 3), ['Uno.', 'Dos.', 'Tres.']);
  s.undo();
  assert.deepEqual(allTexts(s), t, 'one undo step');
});

test('unwrapSection removes the marker and keeps the text', () => {
  const s = report();
  const t = allTexts(s);
  s.define_section('x', 1, 2, undefined, false, undefined);
  s.unwrap_section('x', undefined);
  assert.deepEqual(sections(s), []);
  assert.deepEqual(allTexts(s), t);
});

test('engine errors explain the problem and change nothing', () => {
  const s = report();
  s.define_section('a', 1, 1, undefined, false, undefined);
  const xml = () => paragraphs(s).map(p => p.text).join('|') + JSON.stringify(sections(s));
  const before = xml();
  assert.throws(() => s.define_section('a', 2, 2, undefined, false, undefined), /Ya existe/);
  assert.throws(() => s.define_section('con espacio', 2, 2, undefined, false, undefined), /no válida/);
  assert.throws(() => s.replace_section('nada', '[{"text":"x"}]', false, undefined), /No existe la sección «nada»/);
  assert.throws(() => s.unwrap_section('nada', undefined), /Secciones del documento: «a»/);
  assert.equal(xml(), before);
});

test('checkSectionKey refuses keys that are not strings', () => {
  assert.throws(() => checkSectionKey('replaceSection', ''), /`key` must be a non-empty string/);
  assert.throws(() => checkSectionKey('replaceSection', 3), /got 3/);
  assert.doesNotThrow(() => checkSectionKey('replaceSection', 'introduccion'));
});
