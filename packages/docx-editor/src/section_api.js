// Logical sections of DocxEditor (sections, sectionContent, defineSection, replaceSection,
// unwrapSection): argument checks and reading a section's blocks and text from the engine's
// document elements, kept apart from the DOM so they run under Node.

/** Throws unless `key` is a non-empty string (the engine checks the allowed characters) */
export function checkSectionKey(method, key) {
  if (typeof key !== 'string' || !key) {
    throw new Error(`docx-editor: ${method}: \`key\` must be a non-empty string such as 'introduccion' (got ${String(key)})`);
  }
}

/** True when block `b` of `structure()` lies inside paragraphs `first`–`last` */
const blockInside = (b, first, last) => {
  const [from, to] = b.kind === 'table' ? [b.first, b.last] : [b.paragraph, b.paragraph];
  return from >= first && to <= last;
};

/**
 * A section with its blocks (from `structure()`) and its plain text: paragraphs on lines of
 * their own, table rows as lines with cells separated by tabs.
 * @param {object} section from `sections()`
 * @param {Array<object>} blocks from `structure()`
 * @param {Array<object>} elements the engine's document elements
 */
export function sectionContent(section, blocks, elements) {
  const { first, last } = section;
  if (first == null || last == null) return { ...section, blocks: [], text: '' };
  const lines = [];
  for (const el of elements) {
    if (el.type === 'paragraph' && el.index >= first && el.index <= last) {
      lines.push(el.text || '');
    } else if (el.type === 'table') {
      for (const row of el.rich_rows || []) {
        const cells = row.cells.map(cell => (cell.paragraphs || []).filter(p => p.index >= first && p.index <= last));
        if (cells.every(ps => !ps.length)) continue;
        lines.push(cells.map(ps => ps.map(p => p.text || '').join('\n')).join('\t'));
      }
    }
  }
  return { ...section, blocks: blocks.filter(b => blockInside(b, first, last)), text: lines.join('\n') };
}

/** True when a section holds nothing but (maybe) its heading: no text, no tables */
export function sectionIsEmpty(content) {
  const blocks = content.blocks.slice();
  if (blocks[0]?.kind === 'paragraph' && blocks[0].headingLevel != null) blocks.shift();
  return blocks.every(b => b.kind === 'paragraph' && !(b.text || '').trim());
}

/** A schema as the engine stores it, checking its shape (the engine checks keys and parents) */
export function schemaPayload(schema) {
  if (schema == null) return null;
  if (typeof schema !== 'object' || !Array.isArray(schema.sections)) {
    throw new Error("docx-editor: setSectionSchema needs { type, sections: [{ key, title, required, parent }] } or null");
  }
  return {
    type: String(schema.type ?? ''),
    title: schema.title == null ? null : String(schema.title),
    sections: schema.sections.map((s, i) => {
      if (!s || typeof s.key !== 'string') throw new Error(`docx-editor: setSectionSchema: sections[${i}] needs a string key`);
      return {
        key: s.key,
        title: s.title == null ? null : String(s.title),
        required: s.required !== false,
        parent: s.parent == null ? null : String(s.parent),
      };
    }),
  };
}

/** Indices of `seq` outside one of its longest increasing subsequences */
function outOfOrder(seq) {
  const n = seq.length;
  const len = Array(n).fill(1);
  const prev = Array(n).fill(-1);
  for (let i = 0; i < n; i++) {
    for (let j = 0; j < i; j++) {
      if (seq[j] < seq[i] && len[j] + 1 > len[i]) {
        len[i] = len[j] + 1;
        prev[i] = j;
      }
    }
  }
  const keep = new Set();
  let i = len.indexOf(Math.max(0, ...len));
  while (i >= 0) {
    keep.add(i);
    i = prev[i];
  }
  return seq.map((_, k) => k).filter(k => !keep.has(k));
}

/**
 * Checks the document's sections against its schema.
 * @param {object|null} schema from `sectionSchema()`
 * @param {Array<object>} sections from `sections()`
 * @param {(section: object) => boolean} isEmpty
 * @param {object} messages the editor's `labels.sectionIssues`
 * @returns {{valid: boolean, type: string|null, issues: Array<{kind: string, key: string, message: string}>}}
 *   `valid` ignores `unexpected` issues (sections the schema does not list)
 */
export function validateSections(schema, sections, isEmpty, messages) {
  if (!schema) return { valid: true, type: null, issues: [] };
  const issues = [];
  const add = (kind, s, extra = {}) => {
    const title = s.title || s.key;
    issues.push({ kind, key: s.key, title, ...extra, message: messages[kind](title, extra) });
  };
  const byKey = new Map();
  sections.forEach(s => byKey.set(s.key, [...(byKey.get(s.key) || []), s]));
  const listed = new Map(schema.sections.map((s, index) => [s.key, { ...s, index }]));
  for (const s of schema.sections) {
    const found = byKey.get(s.key) || [];
    if (!found.length) {
      if (s.required) add('missing', s);
      continue;
    }
    if (found.length > 1) add('duplicated', s, { count: found.length });
    const actual = found[0];
    const expected = s.parent ?? null;
    if ((actual.parent ?? null) !== expected) add('parent', s, { expected, actual: actual.parent ?? null });
    if (s.required && isEmpty(actual)) add('empty', s);
  }
  // Order among the sections that share a parent in the schema, as they appear in the document
  const groups = new Map();
  const seen = new Set();
  for (const s of sections) {
    const entry = listed.get(s.key);
    if (!entry || seen.has(s.key)) continue;
    seen.add(s.key);
    const group = entry.parent ?? '';
    groups.set(group, [...(groups.get(group) || []), entry]);
  }
  for (const entries of groups.values()) {
    outOfOrder(entries.map(e => e.index)).forEach(k => add('order', entries[k]));
  }
  for (const s of sections) if (!listed.has(s.key)) add('unexpected', s);
  return { valid: issues.every(i => i.kind === 'unexpected'), type: schema.type, issues };
}

/**
 * Where each section missing from the document goes, following the schema's order: after
 * the previous sibling, else before the next one, else at the end of its parent (or of the
 * document). Sections whose parent is missing and not added are left out.
 * @returns {Array<{key: string, title: string, place: 'after'|'before'|'inside'|'end', relativeTo: string|null, depth: number}>}
 */
export function planMissingSections(schema, sections, { optional = false } = {}) {
  if (!schema) return [];
  const present = new Set(sections.map(s => s.key));
  const depthOf = new Map();
  const plan = [];
  schema.sections.forEach((s, index) => {
    const depth = s.parent == null ? 0 : (depthOf.get(s.parent) ?? 0) + 1;
    depthOf.set(s.key, depth);
    if (present.has(s.key) || (!s.required && !optional)) return;
    if (s.parent != null && !present.has(s.parent)) return;
    const siblings = schema.sections.filter(x => (x.parent ?? null) === (s.parent ?? null));
    const at = siblings.findIndex(x => x.key === s.key);
    const before = siblings.slice(0, at).reverse().find(x => present.has(x.key));
    const after = siblings.slice(at + 1).find(x => present.has(x.key) && sections.some(y => y.key === x.key));
    let place;
    if (before) place = { place: 'after', relativeTo: before.key };
    else if (after) place = { place: 'before', relativeTo: after.key };
    else if (s.parent != null) place = { place: 'inside', relativeTo: s.parent };
    else place = { place: 'end', relativeTo: null };
    plan.push({ key: s.key, title: s.title || s.key, ...place, depth, index });
    present.add(s.key);
  });
  return plan.map(({ index, ...p }) => p);
}
