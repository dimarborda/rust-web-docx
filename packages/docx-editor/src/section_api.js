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
