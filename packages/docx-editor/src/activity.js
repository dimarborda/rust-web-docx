// Activity marks of DocxEditor (markActivity, activities, revealActivity): where an agent or
// any code is working, painted over the pages so the person reading can find it. They are
// view state only: nothing is written into the .docx, and they cost nothing while unused.
//
// - "work" marks stay until `done()` (or their `ttl`): a tinted area, a bar in the margin and
//   a label, plus a chip pinned to the top of the editor that says where it is and jumps
//   there on click without moving the caret.
// - "flash" marks highlight what a call from code just changed and fade out by themselves.
//
// A mark targets a section key (it follows the section as it grows or moves), a paragraph
// range `{from, to}` or a span `{paragraph, offset, length}` (e.g. a marker). Ranges are kept
// in place across edits by comparing the paragraph texts before and after each change.

/** Paragraphs from this index on are text boxes, headers and footers (as in the engine) */
const TEXT_BOX_BASE = 1 << 24;

/**
 * Normalizes a target to `{section}`, `{from, to}` or `{paragraph, offset, length}`.
 * @param {string | {from: number, to?: number} | {paragraph: number, offset?: number, length?: number} | {first: number, count: number}} target
 */
export function normalizeTarget(target) {
  if (typeof target === 'string' && target) return { section: target };
  if (target && typeof target === 'object') {
    if (typeof target.section === 'string') return { section: target.section };
    if (Number.isInteger(target.from)) return { from: target.from, to: Number.isInteger(target.to) ? Math.max(target.to, target.from) : target.from };
    // `{first, count}` as returned by insertParagraphs, replaceParagraphs, replaceSection…
    if (Number.isInteger(target.first) && Number.isInteger(target.count)) {
      return { from: target.first, to: target.first + Math.max(target.count, 1) - 1 };
    }
    if (Number.isInteger(target.paragraph)) {
      const offset = Math.max(0, target.offset ?? 0);
      return Number.isInteger(target.length)
        ? { paragraph: target.paragraph, offset, length: Math.max(0, target.length) }
        : { from: target.paragraph, to: target.paragraph };
    }
  }
  throw new Error('docx-editor: an activity needs a section key, {from, to}, {first, count} or {paragraph, offset, length}');
}

/**
 * Where the paragraphs of `after` came from in `before`: the unchanged head and tail of the
 * list, and how many paragraphs the edit added (negative: removed)
 */
export function textDiff(before, after) {
  let head = 0;
  while (head < before.length && head < after.length && before[head] === after[head]) head++;
  let tail = 0;
  while (
    tail < before.length - head && tail < after.length - head
    && before[before.length - 1 - tail] === after[after.length - 1 - tail]
  ) tail++;
  return { head, tail, oldLength: before.length, delta: after.length - before.length };
}

/** Paragraph `index` of the old list in the new one, given a `textDiff` */
function remapIndex(index, { head, tail, oldLength, delta }, edge) {
  if (index < head) return index;
  if (index >= oldLength - tail) return index + delta;
  // Inside the changed block: the start stays, the end follows the block's new size
  return edge === 'end' ? Math.max(head, index + delta) : index;
}

/** A normalized target moved through an edit (sections need no help: they are looked up) */
export function remapTarget(target, diff) {
  if (target.section != null || diff.delta === 0) return target;
  if (target.from != null) {
    if (target.from >= TEXT_BOX_BASE) return target;
    const from = remapIndex(target.from, diff, 'start');
    return { from, to: Math.max(from, remapIndex(target.to, diff, 'end')) };
  }
  if (target.paragraph >= TEXT_BOX_BASE) return target;
  return { ...target, paragraph: remapIndex(target.paragraph, diff, 'start') };
}

/**
 * Character span of a target today: `{start, end}` positions for selection rects, or null
 * when its section is gone or empty.
 * @param {object} target normalized
 * @param {Array<{key: string, first: number|null, last: number|null}>} sections
 * @param {(index: number) => number} lengthOf code points of a paragraph
 */
export function targetSpan(target, sections, lengthOf) {
  if (target.section != null) {
    const s = sections.find(x => x.key === target.section);
    if (!s || s.first == null) return null;
    return { start: { paragraph: s.first, offset: 0 }, end: { paragraph: s.last, offset: lengthOf(s.last) } };
  }
  if (target.from != null) {
    return { start: { paragraph: target.from, offset: 0 }, end: { paragraph: target.to, offset: lengthOf(target.to) } };
  }
  const length = lengthOf(target.paragraph);
  const offset = Math.min(target.offset, length);
  return {
    start: { paragraph: target.paragraph, offset },
    end: { paragraph: target.paragraph, offset: Math.min(length, offset + target.length) },
  };
}

/**
 * The marks of one editor and their painting.
 * @param {object} env
 * @param {HTMLElement} env.root        the scrolling editor element
 * @param {() => Map<number, HTMLElement>} env.cards page cards by page number
 * @param {() => number} env.zoom
 * @param {() => Array<object>} env.sections
 * @param {(index: number) => number} env.lengthOf
 * @param {(start: object, end: object) => Array<{page: number, x: number, y: number, width: number, height: number}>} env.rects
 * @param {(index: number) => string|null} env.sectionAt key of the deepest section holding a paragraph
 * @param {object} env.labels
 * @param {(list: Array<object>) => void} env.changed
 */
export function createActivityLayer(env) {
  const marks = new Map(); // id → { id, kind, label, color, target, expires, timer }
  let nextId = 1;
  let dock = null;
  let scrollFrame = 0;
  /** Painted geometry by id: first page and its top, for reveal and for the chips */
  const painted = new Map();

  const onScroll = () => {
    cancelAnimationFrame(scrollFrame);
    scrollFrame = requestAnimationFrame(updateChips);
  };
  // The page may scroll in the editor or in any ancestor (or the window): listen to them all,
  // but only while there are chips to update
  const listen = on => {
    const method = on ? 'addEventListener' : 'removeEventListener';
    document[method]('scroll', onScroll, { capture: true, passive: true });
    window[method]('resize', onScroll, { passive: true });
  };

  function add(kind, target, { label, color, ttl } = {}) {
    const id = `a${nextId++}`;
    const mark = { id, kind, label: label ?? env.labels.agent, color: color || null, target: normalizeTarget(target), timer: 0 };
    marks.set(id, mark);
    if (ttl > 0) mark.timer = setTimeout(() => remove(id), ttl);
    paint();
    return id;
  }

  function remove(id) {
    const mark = marks.get(id);
    if (!mark) return false;
    clearTimeout(mark.timer);
    marks.delete(id);
    paint();
    return true;
  }

  /** Public description of the marks, in document order */
  function list() {
    const sections = env.sections();
    return [...marks.values()]
      .map(m => {
        const span = targetSpan(m.target, sections, env.lengthOf);
        const where = painted.get(m.id);
        return {
          id: m.id,
          kind: m.kind,
          label: m.label,
          from: span ? span.start.paragraph : null,
          to: span ? span.end.paragraph : null,
          section: m.target.section ?? (span ? env.sectionAt(span.start.paragraph) : null),
          page: where?.page ?? null,
        };
      })
      .sort((a, b) => (a.from ?? Infinity) - (b.from ?? Infinity));
  }

  function layerOf(card) {
    let layer = card.querySelector(':scope > .docx-activity-layer');
    if (!layer) {
      const canvas = card.querySelector('canvas');
      layer = document.createElement('div');
      layer.className = 'docx-activity-layer';
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
    cards.forEach(card => card.querySelector(':scope > .docx-activity-layer')?.replaceChildren());
    painted.clear();
    if (marks.size) {
      const zoom = env.zoom();
      const sections = env.sections();
      for (const mark of marks.values()) {
        const span = targetSpan(mark.target, sections, env.lengthOf);
        if (!span) continue;
        const byPage = new Map();
        env.rects(span.start, span.end).forEach(r => {
          if (!byPage.has(r.page)) byPage.set(r.page, []);
          byPage.get(r.page).push(r);
        });
        // An empty paragraph has no selection rect: fall back to nothing rather than guess
        let first = true;
        for (const [page, rects] of [...byPage].sort((a, b) => a[0] - b[0])) {
          const card = cards.get(page);
          if (!card) continue;
          const layer = layerOf(card);
          const top = Math.min(...rects.map(r => r.y));
          const bottom = Math.max(...rects.map(r => r.y + r.height));
          const left = Math.min(...rects.map(r => r.x));
          const group = document.createElement('div');
          group.className = `docx-activity docx-activity-${mark.kind}`;
          group.dataset.activity = mark.id;
          if (mark.color) group.style.setProperty('--docx-activity', mark.color);
          rects.forEach(r => {
            const area = document.createElement('div');
            area.className = 'docx-activity-area';
            Object.assign(area.style, {
              left: `${r.x * zoom}px`, top: `${r.y * zoom}px`, width: `${r.width * zoom}px`, height: `${r.height * zoom}px`,
            });
            group.appendChild(area);
          });
          const bar = document.createElement('div');
          bar.className = 'docx-activity-bar';
          Object.assign(bar.style, {
            left: `${Math.max(2, left * zoom - 10)}px`, top: `${top * zoom}px`, height: `${(bottom - top) * zoom}px`,
          });
          group.appendChild(bar);
          if (first && mark.kind === 'work' && mark.label) {
            const tag = document.createElement('span');
            tag.className = 'docx-activity-label';
            tag.textContent = mark.label;
            tag.style.left = `${Math.max(2, left * zoom - 10)}px`;
            tag.style.top = `${top * zoom}px`;
            group.appendChild(tag);
          }
          layer.appendChild(group);
          if (first) painted.set(mark.id, { page, element: group, top });
          first = false;
        }
      }
    }
    renderDock();
    env.changed(list());
  }

  function renderDock() {
    const work = [...marks.values()].filter(m => m.kind === 'work');
    if (!work.length) {
      if (dock) listen(false);
      dock?.remove();
      dock = null;
      return;
    }
    if (!dock) {
      listen(true);
      dock = document.createElement('div');
      dock.className = 'docx-activity-dock';
      env.root.prepend(dock);
    }
    dock.replaceChildren(...work.map(mark => {
      const chip = document.createElement('button');
      chip.type = 'button';
      chip.className = 'docx-activity-chip';
      chip.dataset.activity = mark.id;
      if (mark.color) chip.style.setProperty('--docx-activity', mark.color);
      const where = mark.target.section ?? null;
      const dot = document.createElement('span');
      dot.className = 'docx-activity-dot';
      const text = document.createElement('span');
      text.textContent = where ? `${mark.label} · ${sectionLabel(where)}` : mark.label;
      const arrow = document.createElement('span');
      arrow.className = 'docx-activity-arrow';
      chip.append(dot, text, arrow);
      chip.title = env.labels.goToActivity;
      chip.addEventListener('click', () => reveal(mark.id));
      return chip;
    }));
    updateChips();
  }

  function sectionLabel(key) {
    return env.sections().find(s => s.key === key)?.title || key;
  }

  /** Arrows on the chips: is the work above, below or on screen? */
  function updateChips() {
    if (!dock) return;
    const root = env.root.getBoundingClientRect();
    const view = { top: Math.max(root.top, 0), bottom: Math.min(root.bottom, window.innerHeight) };
    dock.querySelectorAll('.docx-activity-chip').forEach(chip => {
      // The bar spans the mark's lines on its first page (the group itself has no size)
      const bar = painted.get(chip.dataset.activity)?.element.querySelector('.docx-activity-bar');
      const box = bar?.getBoundingClientRect();
      const arrow = chip.querySelector('.docx-activity-arrow');
      const state = !box ? '' : box.bottom < view.top ? '↑' : box.top > view.bottom ? '↓' : '';
      arrow.textContent = state;
      chip.classList.toggle('in-view', !!box && !state);
    });
  }

  /** Scrolls a mark (the last work mark when `id` is omitted) into view; the caret stays */
  function reveal(id) {
    const target = id ?? [...marks.values()].filter(m => m.kind === 'work').at(-1)?.id;
    const element = target && painted.get(target)?.element;
    if (!element) return false;
    const area = element.querySelector('.docx-activity-bar') || element;
    area.scrollIntoView({ block: 'center', behavior: 'smooth' });
    return true;
  }

  return {
    add,
    remove,
    list,
    paint,
    reveal,
    has: id => marks.has(id),
    get size() {
      return marks.size;
    },
    /** Updates a mark's target or label */
    update(id, target, { label } = {}) {
      const mark = marks.get(id);
      if (!mark) return false;
      if (target != null) mark.target = normalizeTarget(target);
      if (label != null) mark.label = label;
      paint();
      return true;
    },
    /** Moves the marks' paragraph ranges through an edit (see `textDiff`) */
    remap(diff) {
      marks.forEach(m => { m.target = remapTarget(m.target, diff); });
    },
    clear() {
      marks.forEach(m => clearTimeout(m.timer));
      marks.clear();
      paint();
    },
    destroy() {
      marks.forEach(m => clearTimeout(m.timer));
      marks.clear();
      cancelAnimationFrame(scrollFrame);
      if (dock) listen(false);
      dock?.remove();
      dock = null;
    },
  };
}
