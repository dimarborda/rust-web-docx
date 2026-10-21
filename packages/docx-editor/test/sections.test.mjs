// Logical sections: the engine bindings DocxEditor calls (sections_json, define_section,
// replace_section, unwrap_section) and the helpers of section_api.js.
// Run after `npm run build:wasm` with: npm test -w packages/docx-editor
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { initSync, DocxSession } from '../wasm/docx_engine.js';
import { checkSectionKey, planMissingSections, schemaPayload, sectionContent, sectionIsEmpty, validateSections } from '../src/section_api.js';
import { sectionBoxes } from '../src/section_layer.js';

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
  const info = JSON.parse(s.define_section('introduccion', first, first + 1, 'Introducción', true, false, undefined));
  assert.deepEqual(info, { key: 'introduccion', title: 'Introducción', first, last: first + 1, depth: 0, parent: null, locked: true, contentLocked: false });
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
  s.define_section('cuerpo', intro, end, undefined, false, false, undefined);
  s.define_section('conclusiones', end - 1, end, 'Conclusiones', false, false, undefined);
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
  s.define_section('conclusiones', from, from + 1, 'Conclusiones', false, false, undefined);
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
  s.define_section('x', 1, 2, undefined, false, false, undefined);
  s.unwrap_section('x', undefined);
  assert.deepEqual(sections(s), []);
  assert.deepEqual(allTexts(s), t);
});

test('engine errors explain the problem and change nothing', () => {
  const s = report();
  s.define_section('a', 1, 1, undefined, false, false, undefined);
  const xml = () => paragraphs(s).map(p => p.text).join('|') + JSON.stringify(sections(s));
  const before = xml();
  assert.throws(() => s.define_section('a', 2, 2, undefined, false, false, undefined), /Ya existe/);
  assert.throws(() => s.define_section('con espacio', 2, 2, undefined, false, false, undefined), /no válida/);
  assert.throws(() => s.replace_section('nada', '[{"text":"x"}]', false, undefined), /No existe la sección «nada»/);
  assert.throws(() => s.unwrap_section('nada', undefined), /Secciones del documento: «a»/);
  assert.equal(xml(), before);
});

test('checkSectionKey refuses keys that are not strings', () => {
  assert.throws(() => checkSectionKey('replaceSection', ''), /`key` must be a non-empty string/);
  assert.throws(() => checkSectionKey('replaceSection', 3), /got 3/);
  assert.doesNotThrow(() => checkSectionKey('replaceSection', 'introduccion'));
});

// ---------- Locks, headings, insertion and schema ----------

/** Blank document: a title, then three Heading 1 sections and a Heading 2 inside the first */
function technicalReport() {
  const s = DocxSession.new_blank('letter');
  const items = [
    { text: 'Informe técnico' },
    { text: '1. Introducción', style: 'Heading1' },
    { text: 'Objeto.' },
    { text: '1.1 Alcance', style: 'Heading2' },
    { text: 'Edificio A.' },
    { text: '2. Caracterización', style: 'Heading1' },
    { text: 'Datos.' },
    { text: '3. Conclusiones', style: 'Heading1' },
    { text: 'Apto.' },
  ];
  s.insert_paragraphs(0, 0, JSON.stringify(items), undefined);
  return s;
}

test('set_section_lock reports both locks and survives saving', () => {
  const s = report();
  s.define_section('x', 1, 2, undefined, false, true, undefined);
  assert.equal(sections(s)[0].contentLocked, true);
  const info = JSON.parse(s.set_section_lock('x', true, false, undefined));
  assert.deepEqual([info.locked, info.contentLocked], [true, false]);
  const reopened = sections(new DocxSession(s.export_bytes()))[0];
  assert.deepEqual([reopened.locked, reopened.contentLocked], [true, false]);
});

test('sections_from_headings nests levels and reports what it skipped', () => {
  const s = technicalReport();
  const result = JSON.parse(s.sections_from_headings(1, 2, undefined));
  assert.deepEqual(result.created.map(x => [x.key, x.parent]), [
    ['introduccion', null], ['alcance', 'introduccion'], ['caracterizacion', null], ['conclusiones', null],
  ]);
  assert.deepEqual(result.skipped, []);
  const again = JSON.parse(s.sections_from_headings(1, 2, undefined));
  assert.equal(again.created.length, 0);
  assert.equal(again.skipped.length, 4);
  s.undo();
  assert.deepEqual(sections(s), [], 'one undo step for every section');
});

test('insert_sections adds several sections in one undo step, or none on error', () => {
  const s = technicalReport();
  s.sections_from_headings(1, 1, undefined);
  const heading = text => ({ text, style: 'Heading1' });
  const result = JSON.parse(s.insert_sections(JSON.stringify([
    { key: 'resumen', title: 'Resumen', place: 'before', relativeTo: 'introduccion', paragraphs: [heading('Resumen'), { text: '' }] },
    { key: 'anexos', title: 'Anexos', place: 'end', paragraphs: [heading('Anexos'), { text: '' }] },
  ]), undefined));
  assert.deepEqual(result.map(r => r.section.key), ['resumen', 'anexos']);
  assert.deepEqual(sections(s).map(x => x.key), ['resumen', 'introduccion', 'caracterizacion', 'conclusiones', 'anexos']);
  s.undo();
  assert.deepEqual(sections(s).map(x => x.key), ['introduccion', 'caracterizacion', 'conclusiones']);
  const before = JSON.stringify(sections(s));
  assert.throws(() => s.insert_sections(JSON.stringify([
    { key: 'ok', place: 'end', paragraphs: [{ text: 'x' }] },
    { key: 'mal', place: 'after', relativeTo: 'nada', paragraphs: [{ text: 'y' }] },
  ]), undefined), /No existe la sección «nada»/);
  assert.equal(JSON.stringify(sections(s)), before, 'all or nothing');
});

const SCHEMA = {
  type: 'informe-tecnico',
  title: 'Informe técnico',
  sections: [
    { key: 'introduccion', title: 'Introducción' },
    { key: 'alcance', title: 'Alcance', parent: 'introduccion', required: false },
    { key: 'caracterizacion', title: 'Caracterización' },
    { key: 'resultados', title: 'Resultados' },
    { key: 'conclusiones', title: 'Conclusiones' },
  ],
};

test('the section schema is stored in the .docx and read back', () => {
  const s = report();
  assert.equal(JSON.parse(s.section_schema_json()), null);
  s.set_section_schema(JSON.stringify(schemaPayload(SCHEMA)), undefined);
  const stored = JSON.parse(new DocxSession(s.export_bytes()).section_schema_json());
  assert.deepEqual(stored, schemaPayload(SCHEMA));
  assert.throws(() => s.set_section_schema(JSON.stringify({ type: 'x', sections: [{ key: 'a', parent: 'b', required: true }] }), undefined), /debe aparecer antes/);
  s.set_section_schema('null', undefined);
  assert.equal(JSON.parse(s.section_schema_json()), null);
});

const MESSAGES = Object.fromEntries(['missing', 'duplicated', 'parent', 'empty', 'order', 'unexpected'].map(k => [k, t => `${k}:${t}`]));
const sec = (key, extra = {}) => ({ key, title: key, first: 0, last: 0, depth: 0, parent: null, locked: false, contentLocked: false, ...extra });

test('validateSections finds missing, empty, misplaced and extra sections', () => {
  const schema = schemaPayload(SCHEMA);
  const ok = validateSections(schema, ['introduccion', 'caracterizacion', 'resultados', 'conclusiones'].map(k => sec(k)), () => false, MESSAGES);
  assert.deepEqual(ok, { valid: true, type: 'informe-tecnico', issues: [] });

  const found = validateSections(schema, [
    sec('conclusiones'),
    sec('introduccion'),
    sec('alcance'),
    sec('caracterizacion'),
    sec('caracterizacion'),
    sec('anexo'),
  ], s => s.key === 'conclusiones', MESSAGES);
  assert.equal(found.valid, false);
  assert.deepEqual(found.issues.map(i => [i.kind, i.key]), [
    ['parent', 'alcance'],
    ['duplicated', 'caracterizacion'],
    ['missing', 'resultados'],
    ['empty', 'conclusiones'],
    ['order', 'conclusiones'],
    ['unexpected', 'anexo'],
  ]);
  assert.equal(found.issues.find(i => i.kind === 'parent').expected, 'introduccion');
  assert.equal(found.issues[0].message, 'parent:Alcance');

  const onlyExtra = validateSections(schema, [...['introduccion', 'caracterizacion', 'resultados', 'conclusiones'].map(k => sec(k)), sec('anexo')], () => false, MESSAGES);
  assert.equal(onlyExtra.valid, true, 'unexpected sections do not make it invalid');
  assert.deepEqual(validateSections(null, [], () => false, MESSAGES), { valid: true, type: null, issues: [] });
});

test('planMissingSections follows the schema order', () => {
  const schema = schemaPayload(SCHEMA);
  const plan = planMissingSections(schema, [sec('caracterizacion'), sec('conclusiones')], { optional: true });
  assert.deepEqual(plan.map(p => [p.key, p.place, p.relativeTo, p.depth]), [
    ['introduccion', 'before', 'caracterizacion', 0],
    ['alcance', 'inside', 'introduccion', 1],
    ['resultados', 'after', 'caracterizacion', 0],
  ]);
  assert.deepEqual(planMissingSections(schema, []).map(p => [p.key, p.place, p.relativeTo]), [
    ['introduccion', 'end', null], ['caracterizacion', 'after', 'introduccion'], ['resultados', 'after', 'caracterizacion'], ['conclusiones', 'after', 'resultados'],
  ], 'without optional, alcance is left out');
});

test('the planned sections can be inserted into a real document', () => {
  const s = technicalReport();
  s.sections_from_headings(1, 1, undefined);
  const plan = planMissingSections(schemaPayload(SCHEMA), sections(s));
  assert.deepEqual(plan.map(p => p.key), ['resultados']);
  s.insert_sections(JSON.stringify(plan.map(p => ({ ...p, paragraphs: [{ text: p.title, style: 'Heading1' }, { text: '' }] }))), undefined);
  assert.deepEqual(sections(s).map(x => x.key), ['introduccion', 'caracterizacion', 'resultados', 'conclusiones']);
});

test('sectionIsEmpty ignores the heading', () => {
  const blocks = (...b) => ({ blocks: b });
  const h = { kind: 'paragraph', text: 'Título', headingLevel: 1 };
  assert.equal(sectionIsEmpty(blocks(h)), true);
  assert.equal(sectionIsEmpty(blocks(h, { kind: 'paragraph', text: '  ', headingLevel: null })), true);
  assert.equal(sectionIsEmpty(blocks(h, { kind: 'paragraph', text: 'Algo', headingLevel: null })), false);
  assert.equal(sectionIsEmpty(blocks(h, { kind: 'table', first: 3, last: 4 })), false);
});

test('sectionBoxes groups line rects by page', () => {
  const boxes = sectionBoxes([
    { page: 1, x: 96, y: 500, width: 300, height: 20 },
    { page: 1, x: 120, y: 520, width: 200, height: 20 },
    { page: 2, x: 96, y: 96, width: 100, height: 18 },
  ]);
  assert.deepEqual(boxes, [{ page: 1, top: 500, bottom: 540, left: 96 }, { page: 2, top: 96, bottom: 114, left: 96 }]);
});
