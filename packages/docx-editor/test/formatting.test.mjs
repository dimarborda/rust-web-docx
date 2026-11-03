// Character and paragraph formatting: the engine bindings DocxEditor calls for the toolbar
// (strikethrough, superscript/subscript, highlight, clear formatting, paragraph styles, line
// spacing, indent and lists). Run after `npm run build:wasm` with: npm test -w packages/docx-editor
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { initSync, DocxSession } from '../wasm/docx_engine.js';

initSync({ module: readFileSync(new URL('../wasm/docx_engine_bg.wasm', import.meta.url)) });

/** A blank document with `texts` as its paragraphs */
const doc = texts => {
  const s = DocxSession.new_blank('A4');
  s.replace_paragraphs(0, 0, JSON.stringify(texts.map(text => ({ text }))), undefined);
  return s;
};
const paragraphs = s => JSON.parse(s.get_paragraphs_json());
const reopen = s => new DocxSession(s.export_bytes());

test('strikethrough, subscript and highlight are written and read back', () => {
  const s = doc(['H2O']);
  const [run] = paragraphs(s)[0].runs;
  const runs = [
    { ...run, text: 'H', strike: true },
    { ...run, text: '2', vert_align: 'subscript' },
    { ...run, text: 'O', highlight: 'FFFF00' },
  ];
  s.update_paragraphs_runs(JSON.stringify([{ index: 0, runs, align: 'left' }]));
  const back = paragraphs(reopen(s))[0].runs.map(r => [r.text, !!r.strike, r.vert_align ?? null, r.highlight ?? null]);
  assert.deepEqual(back, [['H', true, null, null], ['2', false, 'subscript', null], ['O', false, null, 'FFFF00']]);

  // The layout draws the subscript smaller and lower
  const layout = JSON.parse(s.compute_canvas_layout_json(() => 10, undefined));
  const line = layout.pages[0].items.find(i => i.type === 'text' && i.runs?.length);
  const sub = line.runs.find(r => r.text === '2');
  const plain = line.runs.find(r => r.text === 'H');
  assert.ok(sub.font_size < plain.font_size && sub.baseline_shift > 0);
  assert.equal(plain.strike, true);
});

test('clear_formatting removes direct formatting in the range only', () => {
  const s = doc(['uno dos']);
  const [run] = paragraphs(s)[0].runs;
  s.update_paragraphs_runs(JSON.stringify([{ index: 0, runs: [{ ...run, bold: true, strike: true }], align: 'left' }]));
  s.clear_formatting(0, 4, 0, 7);
  const runs = paragraphs(s)[0].runs.map(r => [r.text, r.bold, !!r.strike]);
  assert.deepEqual(runs, [['uno ', true, true], ['dos', false, false]]);
  s.undo();
  assert.equal(paragraphs(s)[0].runs.length, 1);
});

test('lists: bullets, numbering, levels and removal as undo steps', () => {
  const s = doc(['a', 'b', 'c']);
  s.set_list('[0,1,2]', 'number');
  s.change_indent('[1]', 1);
  let ps = paragraphs(reopen(s));
  assert.deepEqual(ps.map(p => p.list_label?.text), ['1.', 'a.', '2.']);
  assert.deepEqual(ps.map(p => p.list_kind), ['number', 'number', 'number']);
  s.set_list('[0,1,2]', 'bullet');
  ps = paragraphs(s);
  assert.deepEqual(ps.map(p => p.list_kind), ['bullet', 'bullet', 'bullet']);
  assert.equal(ps[1].list_level, 1);
  s.set_list('[0,1,2]', 'none');
  assert.ok(paragraphs(s).every(p => !p.list_kind));
  s.undo();
  assert.ok(paragraphs(s).every(p => p.list_kind === 'bullet'));
});

test('paragraph styles, line spacing and indent', () => {
  const s = doc(['Título', 'texto']);
  const styles = JSON.parse(s.paragraph_styles_json());
  assert.equal(styles[0].builtin, 'normal');
  assert.ok(styles.some(st => st.builtin === 'heading 1'));

  s.set_paragraph_style('[0]', 'Heading1');
  s.set_line_spacing('[1]', 1.5);
  s.change_indent('[1]', 1);
  const [title, body] = paragraphs(reopen(s));
  assert.equal(title.is_heading, true);
  assert.equal(body.line_spacing, 1.5);
  assert.equal(body.indent_left, 48);
  assert.throws(() => s.set_paragraph_style('[0]', 'NoExiste'));
});
