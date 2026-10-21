// Section outline of DocxEditor (option `showSections`, `setShowSections()`): a bracket in the
// left margin beside each logical section's lines, with its title at the top, so the person
// reading sees where each section starts and ends. Nested sections go further out into the
// margin. View state only: nothing is written into the .docx, and nothing is painted (or
// computed) while it is hidden.

/** Margin (px at 100 %) between the text and the innermost bracket, and between levels */
const GAP = 22;
const STEP = 8;
/** Narrowest room (px) left of a bracket where its title fits */
const MIN_LABEL_ROOM = 40;

/** Pages a section's lines fall on, with their top, bottom and left edge (unzoomed px) */
export function sectionBoxes(rects) {
  const byPage = new Map();
  for (const r of rects) {
    const box = byPage.get(r.page);
    if (!box) byPage.set(r.page, { page: r.page, top: r.y, bottom: r.y + r.height, left: r.x });
    else {
      box.top = Math.min(box.top, r.y);
      box.bottom = Math.max(box.bottom, r.y + r.height);
      box.left = Math.min(box.left, r.x);
    }
  }
  return [...byPage.values()].sort((a, b) => a.page - b.page);
}

/**
 * @param {object} env
 * @param {() => Map<number, HTMLElement>} env.cards page cards by page number
 * @param {() => number} env.zoom
 * @param {() => Array<object>} env.sections
 * @param {(index: number) => number} env.lengthOf code points of a paragraph
 * @param {(start: object, end: object) => Array<{page: number, x: number, y: number, width: number, height: number}>} env.rects
 * @param {(section: object) => void} env.select called when a title is clicked
 * @param {object} env.labels
 */
export function createSectionLayer(env) {
  let visible = false;

  function layerOf(card) {
    let layer = card.querySelector(':scope > .docx-section-layer');
    if (!layer) {
      const canvas = card.querySelector('canvas');
      layer = document.createElement('div');
      layer.className = 'docx-section-layer';
      Object.assign(layer.style, {
        left: `${canvas.offsetLeft}px`,
        top: `${canvas.offsetTop}px`,
        width: canvas.style.width,
        height: canvas.style.height,
      });
      card.appendChild(layer);
    }
    return layer;
  }

  function paint() {
    const cards = env.cards();
    cards.forEach(card => card.querySelector(':scope > .docx-section-layer')?.remove());
    if (!visible) return;
    const zoom = env.zoom();
    for (const section of env.sections()) {
      if (section.first == null) continue;
      const end = { paragraph: section.last, offset: env.lengthOf(section.last) };
      const boxes = sectionBoxes(env.rects({ paragraph: section.first, offset: 0 }, end));
      boxes.forEach((box, i) => {
        const card = cards.get(box.page);
        if (!card) return;
        const x = Math.max(2, box.left * zoom - GAP - section.depth * STEP);
        const bracket = document.createElement('div');
        bracket.className = 'docx-section-bracket';
        if (i === 0) bracket.classList.add('starts');
        if (i === boxes.length - 1) bracket.classList.add('ends');
        if (section.contentLocked) bracket.classList.add('content-locked');
        bracket.dataset.section = section.key;
        Object.assign(bracket.style, {
          left: `${x}px`, top: `${box.top * zoom}px`, height: `${(box.bottom - box.top) * zoom}px`,
        });
        const layer = layerOf(card);
        layer.appendChild(bracket);
        if (i !== 0) return;
        const name = document.createElement('button');
        name.type = 'button';
        name.className = 'docx-section-name';
        name.dataset.section = section.key;
        const lock = section.contentLocked ? env.labels.sectionContentLockedMark : section.locked ? env.labels.sectionLockedMark : '';
        name.textContent = lock ? `${section.title} ${lock}` : section.title;
        name.title = `${section.title} (${section.key})`;
        name.style.top = `${box.top * zoom}px`;
        const room = x - 4;
        if (room >= MIN_LABEL_ROOM) {
          name.style.left = `${x - 3}px`;
          name.style.maxWidth = `${room}px`;
        } else {
          // No room in the margin: only the bracket, the title shows on hover
          name.classList.add('compact');
          name.style.left = `${x}px`;
        }
        name.addEventListener('mousedown', e => e.preventDefault());
        name.addEventListener('click', () => env.select(section));
        layer.appendChild(name);
      });
    }
  }

  return {
    paint,
    get visible() {
      return visible;
    },
    setVisible(on) {
      visible = !!on;
      paint();
    },
  };
}
