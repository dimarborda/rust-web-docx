// Markers of DocxEditor (markers, fillMarker, variables): the shared lexicon of template
// fields ({{KEY}}, {key}) and pending content ([ … ]) located in the document's paragraphs.
// The engine scans the texts (ooxml-core::markers); this module only maps the results to
// paragraphs, containers and sections, and stays apart from the DOM so it runs under Node.

/** First index of text box paragraphs and of header/footer paragraphs (as in the engine) */
const TEXT_BOX_BASE = 1 << 24;

/** Where a paragraph lives: 'body', 'table', 'textbox' or 'headerFooter' */
export function paragraphPlace(p) {
  if (p.container === 'body') return 'body';
  if (p.container.startsWith('tb:')) return 'textbox';
  if (p.container.startsWith('hf:')) return 'headerFooter';
  return 'table';
}

/**
 * Markers of the paragraphs, in document order.
 * @param {Array<{index: number, text?: string, container: string}>} paragraphs every paragraph
 * @param {Array<{key: string, first: number|null, last: number|null, depth: number}>} sections
 * @param {(textsJson: string) => string} scan the engine's `scan_markers_json`
 * @param {{kind?: 'field'|'pending'}} [filter]
 */
export function collectMarkers(paragraphs, sections, scan, { kind } = {}) {
  const candidates = paragraphs.filter(p => p.text && /[[{]/.test(p.text));
  if (!candidates.length) return [];
  const found = JSON.parse(scan(JSON.stringify(candidates.map(p => p.text))));
  return found
    .filter(m => !kind || m.kind === kind)
    .map(m => {
      const p = candidates[m.item];
      return {
        id: `${p.index}:${m.offset}`,
        kind: m.kind,
        syntax: m.syntax,
        raw: m.raw,
        key: m.key ?? null,
        hint: m.hint ?? null,
        paragraph: p.index,
        offset: m.offset,
        length: m.length,
        place: paragraphPlace(p),
        section: innermostSection(sections, p.index),
      };
    });
}

/** Key of the deepest section holding body paragraph `index`, or null */
export function innermostSection(sections, index) {
  if (index >= TEXT_BOX_BASE) return null;
  let best = null;
  for (const s of sections) {
    if (s.first == null || index < s.first || index > s.last) continue;
    if (!best || s.depth > best.depth) best = s;
  }
  return best ? best.key : null;
}

/**
 * Where `marker` is now in the text of its paragraph: at its offset when unchanged, else the
 * nearest copy of the same `raw` in that paragraph (text typed before it moves it). Null when
 * it is gone.
 * @param {{paragraph: number, offset: number, raw: string}} marker
 * @param {string} text current text of the paragraph
 * @returns {{offset: number, length: number} | null} in code points
 */
export function relocateMarker(marker, text) {
  const chars = Array.from(text || '');
  const raw = Array.from(marker.raw);
  const at = i => raw.every((c, k) => chars[i + k] === c);
  if (at(marker.offset)) return { offset: marker.offset, length: raw.length };
  let best = null;
  for (let i = 0; i + raw.length <= chars.length; i++) {
    if (at(i) && (best == null || Math.abs(i - marker.offset) < Math.abs(best - marker.offset))) best = i;
  }
  return best == null ? null : { offset: best, length: raw.length };
}

/**
 * The first of `markers` (in document order) starting after `from`; a marker starting exactly
 * at `from` does not count, so calling it again from the marker found moves on. With `wrap`,
 * the first marker when none is left after `from`. No `from`: the first marker.
 * @param {Array<{paragraph: number, offset: number}>} markers
 * @param {{paragraph: number, offset: number} | null} from
 * @param {{wrap?: boolean}} [options]
 */
export function markerAfter(markers, from, { wrap = false } = {}) {
  if (!markers.length) return null;
  if (!from) return markers[0];
  const after = markers.find(m => m.paragraph > from.paragraph || (m.paragraph === from.paragraph && m.offset > from.offset));
  return after ?? (wrap ? markers[0] : null);
}

/**
 * Pairs for the engine's batch replace: a key with braces is replaced as written; a bare key
 * fills every spelling of that field found in the document ({{KEY}}, {{ KEY }}, {KEY},
 * {{KEY: hint}}) and `{{KEY}}` itself.
 * @param {Record<string, unknown> | {key: string, value: unknown}[]} values
 * @param {Array<{kind: string, key: string|null, raw: string}>} markers
 */
export function variablePairs(values, markers) {
  const entries = Array.isArray(values) ? values : Object.entries(values ?? {}).map(([key, value]) => ({ key, value }));
  const pairs = [];
  for (const { key, value } of entries) {
    if (!key) continue;
    const text = String(value ?? '');
    if (key.includes('{')) {
      pairs.push({ key, value: text });
      continue;
    }
    const spellings = new Set([`{{${key}}}`]);
    markers.forEach(m => m.kind === 'field' && m.key === key && spellings.add(m.raw));
    spellings.forEach(raw => pairs.push({ key: raw, value: text }));
  }
  // Longest first: `{CLIENTE}` must not eat the inside of `{{CLIENTE}}`
  return pairs.sort((a, b) => b.key.length - a.key.length);
}
