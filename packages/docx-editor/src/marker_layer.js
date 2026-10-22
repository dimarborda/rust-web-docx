// Marker highlight of DocxEditor (option `highlightPending`, `setHighlightPending()`): a soft
// background behind each pending `[ … ]` marker (and, when asked, each `{{field}}`), so the
// person filling a long document sees at a glance what is still missing. View state only:
// nothing is written into the .docx, it is not printed nor exported to PDF, and nothing is
// painted (or scanned) while it is off. It is repainted after every render, so a marker that
// is filled or deleted loses its highlight with it.

/**
 * Kinds of marker to highlight for an option value: `true` → pending only,
 * `{ fields: true }` → pending and fields, anything falsy → none.
 * @param {boolean | {fields?: boolean}} option
 * @returns {Array<'pending'|'field'>}
 */
export function highlightKinds(option) {
  if (!option) return [];
  return typeof option === 'object' && option.fields ? ['pending', 'field'] : ['pending'];
}

/**
 * Boxes to paint: one per line rect of each marker, with its kind.
 * @param {Array<{kind: string, paragraph: number, offset: number, length: number}>} markers
 * @param {(start: object, end: object) => Array<{page: number, x: number, y: number, width: number, height: number}>} rects
 */
export function markerBoxes(markers, rects) {
  const out = [];
  for (const m of markers) {
    if (!m.length) continue;
    const end = { paragraph: m.paragraph, offset: m.offset + m.length };
    for (const r of rects({ paragraph: m.paragraph, offset: m.offset }, end)) {
      if (r.width > 0 && r.height > 0) out.push({ ...r, kind: m.kind });
    }
  }
  return out;
}

/**
 * @param {object} env
 * @param {() => Map<number, HTMLElement>} env.cards page cards by page number
 * @param {() => number} env.zoom
 * @param {(kind: 'pending'|'field') => Array<object>} env.markers markers of one kind
 * @param {(start: object, end: object) => Array<{page: number, x: number, y: number, width: number, height: number}>} env.rects
 */
export function createMarkerLayer(env) {
  let kinds = [];

  function layerOf(card) {
    let layer = card.querySelector(':scope > .docx-marker-layer');
    if (!layer) {
      const canvas = card.querySelector('canvas');
      layer = document.createElement('div');
      layer.className = 'docx-marker-layer';
      Object.assign(layer.style, {
        left: `${canvas.offsetLeft}px`,
        top: `${canvas.offsetTop}px`,
        width: canvas.style.width,
        height: canvas.style.height,
      });
      // Under the caret and the other layers, right above the page
      canvas.after(layer);
    }
    return layer;
  }

  function paint() {
    const cards = env.cards();
    cards.forEach(card => card.querySelector(':scope > .docx-marker-layer')?.remove());
    if (!kinds.length) return;
    const zoom = env.zoom();
    const markers = kinds.flatMap(kind => env.markers(kind));
    for (const box of markerBoxes(markers, env.rects)) {
      const card = cards.get(box.page);
      if (!card) continue;
      const mark = document.createElement('div');
      mark.className = `docx-marker-highlight ${box.kind}`;
      Object.assign(mark.style, {
        left: `${box.x * zoom}px`,
        top: `${box.y * zoom}px`,
        width: `${box.width * zoom}px`,
        height: `${box.height * zoom}px`,
      });
      layerOf(card).appendChild(mark);
    }
  }

  return {
    paint,
    /** Kinds highlighted now (empty when off) */
    get kinds() {
      return [...kinds];
    },
    setOption(option) {
      kinds = highlightKinds(option);
      paint();
    },
  };
}
