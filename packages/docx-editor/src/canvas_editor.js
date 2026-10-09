// Canvas-native editing: the caret, the selection and keyboard input live on top of the
// rendered pages, so there is no per-paragraph edit box. Rust owns the geometry (hit testing,
// caret boxes, vertical moves), the edits and the undo history; this module routes input and
// paints overlays.
//
// Positions are { paragraph, offset } with offsets in Unicode code points of the paragraph
// text, the same unit Rust uses. A selection runs from `anchor` (where it started) to `focus`
// (where the caret is) and may span paragraphs.

const WORD_CHAR = /[\p{L}\p{N}_]/u;
/** First index of header and footer paragraphs (the engine numbers them apart) */
const HEADER_FOOTER_BASE = 1 << 25;
const IS_MAC = /Mac|iPhone|iPad/.test(navigator.platform);
const LINE_BREAK = '\u000B'; // Shift+Enter: line break inside the paragraph (Word's ^l)
const TYPING_GROUP_MS = 1500;

/**
 * @param {object} env
 * @param {() => any} env.session            current DocxSession
 * @param {() => {list: any[], byIndex: Map}} env.paragraphs  every paragraph (body and table
 *        cells) in document order, each with `index`, `text` and `container`
 * @param {() => number} env.zoom
 * @param {Function} env.measure             text measurer passed to Rust
 * @param {() => void} env.documentChanged   re-render after an edit
 * @param {(sel) => void} env.selectionChanged
 * @param {(msg: string) => void} env.hint
 * @param {HTMLElement} env.root             element that hosts the hidden input
 * @param {object} env.labels                { editor, nothingToUndo, nothingToRedo }
 * @param {(ref) => void} env.imageSelected   a picture was selected (ref = { paragraph, index }) or deselected (null)
 * @param {(ref, change) => void} env.imageEdited  the user resized or moved a picture with the mouse or
 *        keyboard: change = { width?, height?, dx?, dy? } in page px
 * @param {(ref) => void} env.imageDeleted    Delete / Backspace on a selected picture
 * @param {() => boolean} env.readOnly        when true the user can move the caret, select and copy, but
 *        not change the document (typing, deleting, pasting, undo and picture editing are ignored)
 */
export function createCanvasEditor(env) {
  let anchor = null;
  let focus = null;
  let goalX = null;        // remembered column for ↑/↓
  let dragging = false;
  let composing = false;
  let lastEdit = null;     // { kind, caret, time } for grouping keystrokes into one undo step
  let image = null;        // selected picture: { paragraph, index }
  let imageDrag = null;    // { kind: 'move' | 'resize', dir, start: {x, y, width, height}, from: {x, y}, frame, ratio, moved }
  const pages = new Map(); // page number → { canvas, overlay, page }

  // A hidden textarea owns keyboard focus: it receives typing, IME/dead-key composition
  // (´ + e = é) and clipboard events
  const input = document.createElement('textarea');
  input.className = 'canvas-editor-input';
  input.setAttribute('autocomplete', 'off');
  input.setAttribute('autocorrect', 'off');
  input.setAttribute('autocapitalize', 'off');
  input.setAttribute('spellcheck', 'false');
  input.setAttribute('aria-label', env.labels.editor);
  env.root.appendChild(input);

  // ---------- Geometry (Rust) ----------

  const parse = json => (json ? JSON.parse(json) : null);
  const session = () => env.session();
  // Header mode (double click on a header or footer): clicks reach header and footer text,
  // and `partPage` says which page's copy shows the caret
  let headerMode = false;
  let partPage = null;
  let bodyCaret = null; // caret to go back to when leaving header mode
  const hitTest = (page, x, y, header = headerMode) => parse(session()?.hit_test(page, x, y, env.measure, header));
  const pageHint = pos => (pos && pos.paragraph >= HEADER_FOOTER_BASE ? partPage ?? undefined : undefined);
  const caretBox = pos => parse(session()?.caret_box(pos.paragraph, pos.offset, env.measure, pageHint(pos)));
  const moveVertical = (pos, dir, x) => parse(session()?.move_vertical(pos.paragraph, pos.offset, dir, x, env.measure, pageHint(pos)));
  const selectionRects = (a, b) =>
    parse(session()?.selection_rects_range(a.paragraph, a.offset, b.paragraph, b.offset, env.measure, pageHint(a))) || [];

  // ---------- Text helpers ----------

  function paragraph(index) {
    return env.paragraphs().byIndex.get(index);
  }

  function chars(index) {
    const p = paragraph(index);
    return p ? Array.from(p.text || '') : [];
  }

  // Text box paragraphs live in containers "tb:<anchor>:<index>"; the caret never wanders
  // between a text box and the rest of the document
  // Header and footer paragraphs ("hf:<part>") are separate flows too
  const isBox = p => !!p && (p.container.startsWith('tb:') || p.container.startsWith('hf:'));
  const sameFlow = (a, b) => !!a && !!b && (isBox(a) || isBox(b) ? a.container === b.container : true);

  /** Paragraphs the caret at `index` can reach: its text box, or the body and its tables */
  function flowOf(index) {
    const here = paragraph(index);
    return env.paragraphs().list.filter(p => (here ? sameFlow(here, p) : !isBox(p)));
  }

  function lastParagraphIndex(index) {
    const list = index === undefined ? env.paragraphs().list.filter(p => !isBox(p)) : flowOf(index);
    return list.length ? list[list.length - 1].index : -1;
  }

  function firstParagraphIndex(index) {
    const list = index === undefined ? env.paragraphs().list.filter(p => !isBox(p)) : flowOf(index);
    return list.length ? list[0].index : 0;
  }

  /** True when paragraphs `a` and `a + 1` can be joined: same container, nothing between */
  function adjacent(a) {
    const p = paragraph(a);
    const next = paragraph(a + 1);
    return !!p && !!next && p.container === next.container;
  }

  /** First paragraph of the next (dir > 0) or previous table cell, for Tab / Shift+Tab */
  function neighbourCell(index, dir) {
    const current = paragraph(index);
    if (!current || current.container === 'body' || isBox(current)) return null;
    const table = current.container.split(':')[0];
    const list = env.paragraphs().list;
    const inTable = list.filter(p => p.container !== 'body' && p.container.split(':')[0] === table);
    const cells = [...new Set(inTable.map(p => p.container))];
    const target = cells[cells.indexOf(current.container) + dir];
    return target ? inTable.find(p => p.container === target) : null;
  }

  const before = (a, b) => a.paragraph < b.paragraph || (a.paragraph === b.paragraph && a.offset < b.offset);
  const samePos = (a, b) => !!a && !!b && a.paragraph === b.paragraph && a.offset === b.offset;

  function ordered() {
    return before(focus, anchor) ? [focus, anchor] : [anchor, focus];
  }

  function collapsed() {
    return !anchor || !focus || samePos(anchor, focus);
  }

  function selection() {
    if (!anchor || !focus) return null;
    const [start, end] = ordered();
    return { anchor, focus, start, end, collapsed: collapsed() };
  }

  function selectedText() {
    if (collapsed()) return '';
    const [start, end] = ordered();
    const parts = [];
    for (let p = start.paragraph; p <= end.paragraph; p++) {
      if (!paragraph(p)) continue;
      const text = chars(p);
      parts.push(text.slice(p === start.paragraph ? start.offset : 0, p === end.paragraph ? end.offset : text.length).join(''));
    }
    return parts.join('\n');
  }

  // Word boundaries for Ctrl/Alt+arrows and double click
  function wordStart(text, offset) {
    let i = offset;
    while (i > 0 && !WORD_CHAR.test(text[i - 1])) i--;
    while (i > 0 && WORD_CHAR.test(text[i - 1])) i--;
    return i;
  }

  function wordEnd(text, offset) {
    let i = offset;
    while (i < text.length && !WORD_CHAR.test(text[i])) i++;
    while (i < text.length && WORD_CHAR.test(text[i])) i++;
    return i;
  }

  // ---------- Selection state ----------

  function setCaret(pos, extend = false) {
    if (extend && anchor) {
      focus = pos;
    } else {
      anchor = pos;
      focus = pos;
    }
  }

  function stepHorizontal(pos, dir, byWord) {
    const text = chars(pos.paragraph);
    if (dir < 0) {
      if (pos.offset > 0) {
        return { paragraph: pos.paragraph, offset: byWord ? wordStart(text, pos.offset) : pos.offset - 1 };
      }
      const prev = paragraph(pos.paragraph - 1);
      return prev && sameFlow(prev, paragraph(pos.paragraph)) ? { paragraph: prev.index, offset: chars(prev.index).length } : pos;
    }
    if (pos.offset < text.length) {
      return { paragraph: pos.paragraph, offset: byWord ? wordEnd(text, pos.offset) : pos.offset + 1 };
    }
    const next = paragraph(pos.paragraph + 1);
    return next && sameFlow(next, paragraph(pos.paragraph)) ? { paragraph: next.index, offset: 0 } : pos;
  }

  // ---------- Editing ----------

  /**
   * Replaces `range` (default: the selection) with `text` through Rust's single edit primitive.
   * `kind` decides undo grouping: consecutive typing (or deleting) at the caret joins one step,
   * starting a new step at each space so undo goes back word by word.
   */
  function edit(text, kind, range = ordered()) {
    if (!session() || !focus || env.readOnly()) return;
    const [start, end] = range;
    const continues = lastEdit && lastEdit.kind === kind && Date.now() - lastEdit.time < TYPING_GROUP_MS &&
      (samePos(lastEdit.caret, start) || samePos(lastEdit.caret, end));
    const coalesce = continues && (kind === 'delete' || (kind === 'type' && !/\s/.test(text)));

    let caret;
    try {
      caret = parse(session().edit(
        start.paragraph, start.offset, end.paragraph, end.offset, text,
        JSON.stringify({ anchor, focus }), coalesce,
      ));
    } catch (err) {
      console.error('Edit failed:', err);
      env.hint(String(err));
      return;
    }
    anchor = focus = caret;
    goalX = null;
    lastEdit = { kind, caret, time: Date.now() };
    env.documentChanged();
  }

  function deleteBackward(byWord) {
    if (!collapsed()) return edit('', 'cut');
    const text = chars(focus.paragraph);
    if (focus.offset > 0) {
      const from = { paragraph: focus.paragraph, offset: byWord ? wordStart(text, focus.offset) : focus.offset - 1 };
      return edit('', 'delete', [from, focus]);
    }
    // At the start of a paragraph: join it with the previous one
    const prev = focus.paragraph - 1;
    if (prev < 0 || !sameFlow(paragraph(prev), paragraph(focus.paragraph))) return;
    const prevEnd = { paragraph: prev, offset: chars(prev).length };
    if (adjacent(prev)) {
      edit('', 'join', [prevEnd, focus]);
    } else {
      setCaret(prevEnd); // a table sits between: never delete it with Backspace
    }
  }

  function deleteForward(byWord) {
    if (!collapsed()) return edit('', 'cut');
    const text = chars(focus.paragraph);
    if (focus.offset < text.length) {
      const to = { paragraph: focus.paragraph, offset: byWord ? wordEnd(text, focus.offset) : focus.offset + 1 };
      return edit('', 'delete', [focus, to]);
    }
    const next = focus.paragraph + 1;
    if (!sameFlow(paragraph(next), paragraph(focus.paragraph))) return;
    if (adjacent(focus.paragraph)) {
      edit('', 'join', [focus, { paragraph: next, offset: 0 }]);
    } else {
      setCaret({ paragraph: next, offset: 0 });
    }
  }

  function restoreSelection(json) {
    const saved = parse(json);
    if (saved?.anchor && saved?.focus) {
      anchor = saved.anchor;
      focus = saved.focus;
    } else if (saved && 'paragraph' in saved) {
      anchor = focus = saved; // a caret position (after an edit)
    }
  }

  function undoRedo(redo) {
    if (!session()) return;
    const result = parse(redo ? session().redo() : session().undo());
    if (!result?.done) {
      env.hint(redo ? env.labels.nothingToRedo : env.labels.nothingToUndo);
      return;
    }
    restoreSelection(result.selection);
    lastEdit = null;
    goalX = null;
    env.documentChanged();
  }

  // ---------- Painting ----------

  function paint({ reveal = false } = {}) {
    pages.forEach(({ overlay }) => overlay.replaceChildren());
    if (image) {
      const found = imageItem(image);
      if (found) {
        paintImageFrame(found.number, found.item, reveal);
        return;
      }
      selectImage(null);
    }
    if (!focus) {
      env.selectionChanged(null);
      return;
    }
    const zoom = env.zoom();

    if (!collapsed()) {
      const [start, end] = ordered();
      selectionRects(start, end).forEach(r => {
        const page = pages.get(r.page);
        if (!page) return;
        const div = document.createElement('div');
        div.className = 'canvas-selection';
        Object.assign(div.style, {
          left: `${r.x * zoom}px`,
          top: `${r.y * zoom}px`,
          width: `${r.width * zoom}px`,
          height: `${r.height * zoom}px`,
        });
        page.overlay.appendChild(div);
      });
    }

    paintTextBoxOutline(zoom);
    paintHeaderMode(zoom);
    const box = caretBox(focus);
    const page = box && pages.get(box.page);
    if (page) {
      const caret = document.createElement('div');
      caret.className = 'canvas-caret';
      if (!collapsed()) caret.classList.add('with-selection');
      Object.assign(caret.style, {
        left: `${box.x * zoom}px`,
        top: `${box.y * zoom}px`,
        height: `${box.height * zoom}px`,
      });
      page.overlay.appendChild(caret);

      // Keep the hidden input at the caret so IME candidate windows open in place. It is
      // positioned inside the root (not the viewport), so transformed or contained ancestors
      // cannot push it away, and focusing or typing scrolls to the caret, not elsewhere.
      const canvasRect = page.canvas.getBoundingClientRect();
      const rootRect = env.root.getBoundingClientRect();
      input.style.left = `${canvasRect.left - rootRect.left - env.root.clientLeft + env.root.scrollLeft + box.x * zoom}px`;
      input.style.top = `${canvasRect.top - rootRect.top - env.root.clientTop + env.root.scrollTop + box.y * zoom}px`;
      // Only a visible caret can be scrolled to (with a selection it is hidden)
      if (reveal && collapsed()) caret.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
    env.selectionChanged(selection());
  }

  // ---------- Pictures ----------

  const sameImage = (a, b) => !!a && !!b && a.paragraph === b.paragraph && a.index === b.index;
  // Pictures, text boxes and shapes of the body are selected and edited the same way
  const isBodyImage = item => (item.type === 'image' || item.type === 'shape') && item.paragraph_index != null && item.image_index != null;
  const RESIZE_DIRS = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
  const BOX_EDGE = 5; // px of a text box's border that select the box instead of its text
  const MIN_IMAGE_PX = 8;

  function selectImage(ref) {
    const next = ref ? { paragraph: ref.paragraph, index: ref.index } : null;
    if (sameImage(next, image) || (!next && !image)) return;
    image = next;
    imageDrag = null;
    env.imageSelected(image);
  }

  /** Where the last render drew a picture: { number, item } */
  function imageItem(ref) {
    for (const [number, { page }] of pages) {
      const item = page.items?.find(it => isBodyImage(it) && it.paragraph_index === ref.paragraph && it.image_index === ref.index);
      if (item) return { number, item };
    }
    return null;
  }

  /** Topmost picture under a page point. Pictures behind the text are only picked where
   *  there is no text, so clicking on text over a letterhead still places the caret. */
  function imageAt(page, x, y) {
    const items = page.items || [];
    const inside = (it, w, h) => x >= it.x && x <= it.x + w && y >= it.y && y <= it.y + h;
    const overText = () => items.some(it => it.type === 'text' && it.line &&
      x >= it.line.left && x <= it.line.right && y >= it.line.top && y <= it.line.top + it.height && it.text.trim());
    for (let i = items.length - 1; i >= 0; i--) {
      const it = items[i];
      if (!isBodyImage(it) || !inside(it, it.width, it.height)) continue;
      // Inside an editable text box a click edits its text; its border selects the box
      if (it.editable_text) {
        const edge = Math.min(BOX_EDGE, it.width / 4, it.height / 4);
        const nearEdge = x - it.x < edge || it.x + it.width - x < edge || y - it.y < edge || it.y + it.height - y < edge;
        return nearEdge ? it : null;
      }
      if (it.is_background && overText()) continue;
      return it;
    }
    return null;
  }

  function paintImageFrame(number, item, reveal) {
    const { overlay } = pages.get(number);
    const zoom = env.zoom();
    const frame = document.createElement('div');
    frame.className = 'canvas-image-frame';
    if (item.anchored) frame.classList.add('movable');
    placeFrame(frame, item, zoom);
    RESIZE_DIRS.forEach(dir => {
      const handle = document.createElement('div');
      handle.className = `canvas-image-handle ${dir}`;
      handle.dataset.dir = dir;
      frame.appendChild(handle);
    });
    frame.addEventListener('mousedown', e => startImageDrag(e, item, frame, zoom));
    overlay.appendChild(frame);
    if (reveal) frame.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    env.selectionChanged(selection());
  }

  function placeFrame(frame, box, zoom) {
    Object.assign(frame.style, {
      left: `${box.x * zoom}px`,
      top: `${box.y * zoom}px`,
      width: `${box.width * zoom}px`,
      height: `${box.height * zoom}px`,
    });
  }

  function startImageDrag(e, item, frame, zoom) {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    const dir = e.target.dataset?.dir;
    // Pictures in line with the text move with it; only floating ones can be dragged
    if (!dir && !item.anchored) return input.focus({ preventScroll: true });
    imageDrag = {
      kind: dir ? 'resize' : 'move',
      dir,
      start: { x: item.x, y: item.y, width: item.width, height: item.height },
      from: { x: e.clientX, y: e.clientY },
      box: { x: item.x, y: item.y, width: item.width, height: item.height },
      frame,
      zoom,
      ratio: item.height / item.width,
      moved: false,
    };
    frame.classList.add('dragging');
    input.focus({ preventScroll: true });
  }

  function dragImage(e) {
    const d = imageDrag;
    const dx = (e.clientX - d.from.x) / d.zoom;
    const dy = (e.clientY - d.from.y) / d.zoom;
    if (Math.abs(dx) + Math.abs(dy) > 1) d.moved = true;
    const s = d.start;
    if (d.kind === 'move') {
      d.box = { ...s, x: s.x + dx, y: s.y + dy };
    } else {
      let { x, y, width, height } = s;
      const dir = d.dir;
      if (dir.includes('e')) width = s.width + dx;
      if (dir.includes('w')) width = s.width - dx;
      if (dir.includes('s')) height = s.height + dy;
      if (dir.includes('n')) height = s.height - dy;
      width = Math.max(MIN_IMAGE_PX, width);
      height = Math.max(MIN_IMAGE_PX, height);
      // Corners keep the proportions (Shift frees them)
      if (dir.length === 2 && !e.shiftKey) {
        const scale = Math.max(width / s.width, height / s.height);
        width = Math.max(MIN_IMAGE_PX, s.width * scale);
        height = Math.max(MIN_IMAGE_PX, s.height * scale);
      }
      if (dir.includes('w')) x = s.x + s.width - width;
      if (dir.includes('n')) y = s.y + s.height - height;
      d.box = { x, y, width, height };
    }
    placeFrame(d.frame, d.box, d.zoom);
  }

  function endImageDrag() {
    const d = imageDrag;
    imageDrag = null;
    d.frame.classList.remove('dragging');
    if (!d.moved || !image) return;
    const round = v => Math.round(v * 100) / 100;
    const change = {};
    if (d.kind === 'resize') {
      change.width = round(d.box.width);
      change.height = round(d.box.height);
    }
    const dx = round(d.box.x - d.start.x);
    const dy = round(d.box.y - d.start.y);
    // A floating picture resized from its left or top edge keeps its opposite edge in place
    if (dx || dy) {
      const item = imageItem(image)?.item;
      if (item?.anchored) Object.assign(change, { dx, dy });
    }
    env.imageEdited(image, change);
  }

  /** Dashed outline around the text box the caret is in, as Word shows it */
  function paintTextBoxOutline(zoom) {
    const p = paragraph(focus.paragraph);
    if (!isBox(p)) return;
    const [, anchorIndex, k] = p.container.split(':').map(Number);
    for (const { page, overlay } of pages.values()) {
      const shape = page.items?.find(it => it.type === 'shape' && it.paragraph_index === anchorIndex && it.image_index === k);
      if (!shape) continue;
      const outline = document.createElement('div');
      outline.className = 'canvas-textbox-outline';
      Object.assign(outline.style, {
        left: `${shape.x * zoom}px`,
        top: `${shape.y * zoom}px`,
        width: `${shape.width * zoom}px`,
        height: `${shape.height * zoom}px`,
      });
      overlay.appendChild(outline);
      return;
    }
  }

  // ---------- Pointer input ----------

  function pagePoint(canvas, e) {
    const rect = canvas.getBoundingClientRect();
    const zoom = env.zoom();
    return { x: (e.clientX - rect.left) / zoom, y: (e.clientY - rect.top) / zoom };
  }

  function positionAt(clientX, clientY) {
    const el = document.elementFromPoint(clientX, clientY);
    if (!el || el.tagName !== 'CANVAS' || !el.dataset.pageNum) return null;
    if (![...pages.values()].some(p => p.canvas === el)) return null; // another editor's page
    const pt = pagePoint(el, { clientX, clientY });
    return hitTest(Number(el.dataset.pageNum), pt.x, pt.y);
  }

  function onMouseDown(e, page, card, canvas) {
    if (e.button !== 0) return;
    const pt = pagePoint(canvas, e);
    // Read-only: pictures are not selected, clicks always reach the text
    const picture = !env.readOnly() && imageAt(pages.get(page.page_number)?.page || page, pt.x, pt.y);
    if (picture) {
      e.preventDefault();
      selectImage({ paragraph: picture.paragraph_index, index: picture.image_index });
      if (!focus) anchor = focus = { paragraph: picture.paragraph_index, offset: 0 };
      lastEdit = null;
      input.focus({ preventScroll: true });
      paint();
      // Press and drag in one go moves a floating picture
      const frame = pages.get(page.page_number)?.overlay.querySelector('.canvas-image-frame');
      if (frame && picture.anchored) startImageDrag(e, picture, frame, env.zoom());
      return;
    }
    selectImage(null);
    // Double click on a header or footer enters header mode; a click elsewhere leaves it
    let pos = null;
    if (headerMode || e.detail === 2) {
      pos = hitTest(page.page_number, pt.x, pt.y, true);
      if (pos && !headerMode) {
        enterHeaderMode(pos, page.page_number);
        e.preventDefault();
        return;
      }
      if (pos) partPage = page.page_number;
      else if (headerMode) leaveHeaderMode(false);
    }
    pos = pos || hitTest(page.page_number, pt.x, pt.y, false);
    if (!pos) return;
    e.preventDefault();

    if (e.detail === 2) {
      const text = chars(pos.paragraph);
      const atWord = WORD_CHAR.test(text[pos.offset] || '');
      anchor = { paragraph: pos.paragraph, offset: atWord ? wordStart(text, pos.offset + 1) : pos.offset };
      focus = { paragraph: pos.paragraph, offset: atWord ? wordEnd(text, pos.offset) : pos.offset };
    } else if (e.detail >= 3) {
      anchor = { paragraph: pos.paragraph, offset: 0 };
      focus = { paragraph: pos.paragraph, offset: chars(pos.paragraph).length };
    } else {
      setCaret(pos, e.shiftKey);
      dragging = true;
    }
    goalX = null;
    lastEdit = null;
    input.focus({ preventScroll: true });
    paint();
  }

  function enterHeaderMode(pos, pageNumber) {
    bodyCaret = focus && focus.paragraph < HEADER_FOOTER_BASE ? focus : bodyCaret;
    headerMode = true;
    partPage = pageNumber;
    anchor = focus = pos;
    goalX = null;
    lastEdit = null;
    input.focus({ preventScroll: true });
    paint();
  }

  /** Back to the body; with `restore` the caret returns where it was before header mode */
  function leaveHeaderMode(restore) {
    headerMode = false;
    partPage = null;
    if (restore) {
      anchor = focus = bodyCaret || { paragraph: firstParagraphIndex(), offset: 0 };
    }
  }

  /** Dashed boundary and label of the header and footer on each page while editing them */
  function paintHeaderMode(zoom) {
    if (!headerMode) return;
    for (const { page, overlay } of pages.values()) {
      const lines = (page.items || []).filter(it => it.type === 'text' && it.line && it.paragraph_index >= HEADER_FOOTER_BASE);
      const half = page.height / 2;
      for (const [kind, group] of [['header', lines.filter(l => l.line.top < half)], ['footer', lines.filter(l => l.line.top >= half)]]) {
        if (!group.length) continue;
        const y = kind === 'header'
          ? Math.max(...group.map(l => l.line.top + l.height)) + 4
          : Math.min(...group.map(l => l.line.top)) - 4;
        const boundary = document.createElement('div');
        boundary.className = `canvas-part-boundary ${kind}`;
        boundary.style.top = `${y * zoom}px`;
        const label = document.createElement('span');
        label.textContent = env.labels[kind] || (kind === 'header' ? 'Encabezado' : 'Pie de página');
        boundary.appendChild(label);
        overlay.appendChild(boundary);
      }
    }
  }

  function onMouseMove(e) {
    if (imageDrag) {
      if (e.buttons & 1) dragImage(e);
      else endImageDrag();
      return;
    }
    if (!dragging || !(e.buttons & 1)) {
      dragging = false;
      return;
    }
    const pos = positionAt(e.clientX, e.clientY);
    // A selection stays in the text box (or the body) it started in
    if (pos && sameFlow(paragraph(pos.paragraph), paragraph(anchor?.paragraph))) {
      setCaret(pos, true);
      paint();
    }
  }
  function onMouseUp() {
    dragging = false;
    if (imageDrag) endImageDrag();
  }
  document.addEventListener('mousemove', onMouseMove);
  document.addEventListener('mouseup', onMouseUp);

  // ---------- Keyboard input ----------

  /** Keys for a selected picture; returns true when handled */
  function imageKey(e) {
    const key = e.key;
    if (key === 'Delete' || key === 'Backspace') {
      const ref = image;
      selectImage(null);
      env.imageDeleted(ref);
      return true;
    }
    if (key === 'Escape') {
      selectImage(null);
      paint();
      return true;
    }
    const step = e.shiftKey ? 10 : 1;
    const nudge = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[key];
    if (nudge && imageItem(image)?.item.anchored) {
      env.imageEdited(image, { dx: nudge[0], dy: nudge[1] });
      return true;
    }
    // Anything else (typing, arrows on an inline picture) goes back to the text
    if (!e.metaKey && !e.ctrlKey) {
      selectImage(null);
      paint();
    }
    return false;
  }

  input.addEventListener('keydown', e => {
    if (image && !composing && !e.isComposing && imageKey(e)) {
      e.preventDefault();
      return;
    }
    if (!focus || composing || e.isComposing) return;
    const mod = e.metaKey || e.ctrlKey;
    const byWord = IS_MAC ? e.altKey : e.ctrlKey;
    const lineEdge = IS_MAC ? e.metaKey : false;
    const key = e.key.length === 1 ? e.key.toLowerCase() : e.key;
    let handled = true;
    let moved = true; // caret moved without editing: start a new undo group

    switch (key) {
      case 'ArrowLeft':
      case 'ArrowRight': {
        const dir = key === 'ArrowLeft' ? -1 : 1;
        if (lineEdge) {
          const box = caretBox(focus);
          if (box) setCaret({ paragraph: focus.paragraph, offset: dir < 0 ? box.line_start : box.line_end }, e.shiftKey);
        } else if (!collapsed() && !e.shiftKey) {
          const [start, end] = ordered();
          setCaret(dir < 0 ? start : end);
        } else {
          setCaret(stepHorizontal(focus, dir, byWord), e.shiftKey);
        }
        goalX = null;
        break;
      }
      case 'ArrowUp':
      case 'ArrowDown': {
        const dir = key === 'ArrowUp' ? -1 : 1;
        if (lineEdge) {
          const last = lastParagraphIndex(focus.paragraph);
          const target = dir < 0
            ? { paragraph: firstParagraphIndex(focus.paragraph), offset: 0 }
            : { paragraph: last, offset: chars(last).length };
          setCaret(target, e.shiftKey);
          break;
        }
        if (goalX === null) goalX = caretBox(focus)?.x ?? 0;
        const next = moveVertical(focus, dir, goalX);
        // At the first/last line, ↑/↓ go to the start/end of that paragraph
        const fallback = { paragraph: focus.paragraph, offset: dir < 0 ? 0 : chars(focus.paragraph).length };
        setCaret(next || fallback, e.shiftKey);
        break;
      }
      case 'Home':
      case 'End': {
        const box = caretBox(focus);
        if (box) setCaret({ paragraph: focus.paragraph, offset: key === 'Home' ? box.line_start : box.line_end }, e.shiftKey);
        goalX = null;
        break;
      }
      case 'Backspace':
        deleteBackward(byWord);
        moved = false;
        break;
      case 'Delete':
        deleteForward(byWord);
        moved = false;
        break;
      case 'Enter':
        edit(e.shiftKey ? LINE_BREAK : '\n', 'enter');
        moved = false;
        break;
      case 'Tab': {
        // In a table Tab moves between cells, as in Word; elsewhere it inserts a tab
        const cell = neighbourCell(focus.paragraph, e.shiftKey ? -1 : 1);
        if (cell) {
          anchor = { paragraph: cell.index, offset: 0 };
          focus = { paragraph: cell.index, offset: chars(cell.index).length };
        } else if (paragraph(focus.paragraph)?.container !== 'body' && !isBox(paragraph(focus.paragraph))) {
          // Last cell (or first with Shift): stay put rather than typing a tab
        } else {
          edit('\t', 'type');
          moved = false;
        }
        break;
      }
      case 'Escape':
        if (headerMode) {
          leaveHeaderMode(true);
          break;
        }
        anchor = focus = null;
        input.blur();
        break;
      case 'a':
        if (!mod) { handled = false; break; }
        {
          // Select all: the whole text box when the caret is in one, else the document
          const last = lastParagraphIndex(focus.paragraph);
          anchor = { paragraph: firstParagraphIndex(focus.paragraph), offset: 0 };
          focus = { paragraph: last, offset: chars(last).length };
        }
        break;
      case 'z':
        if (!mod) { handled = false; break; }
        if (!env.readOnly()) undoRedo(e.shiftKey);
        moved = false;
        break;
      case 'y':
        if (!mod || IS_MAC) { handled = false; break; }
        if (!env.readOnly()) undoRedo(true);
        moved = false;
        break;
      default:
        handled = false; // typing and clipboard arrive as input / clipboard events
    }

    if (handled) {
      e.preventDefault();
      if (moved) lastEdit = null;
      paint({ reveal: true });
    }
  });

  input.addEventListener('compositionstart', () => {
    composing = true;
  });

  input.addEventListener('compositionend', () => {
    composing = false;
    const text = input.value;
    input.value = '';
    if (text && focus) edit(text, 'type');
  });

  input.addEventListener('input', e => {
    if (composing || e.isComposing) return;
    const text = input.value;
    input.value = '';
    if (text && focus) edit(text, 'type');
  });

  input.addEventListener('copy', e => {
    const text = selectedText();
    if (!text) return;
    e.preventDefault();
    e.clipboardData.setData('text/plain', text);
  });

  input.addEventListener('cut', e => {
    const text = selectedText();
    if (!text) return;
    e.preventDefault();
    e.clipboardData.setData('text/plain', text);
    edit('', 'cut'); // read-only: copies without deleting
  });

  input.addEventListener('paste', e => {
    e.preventDefault();
    if (!focus) return;
    // Each pasted line becomes a paragraph, as in Word
    const text = (e.clipboardData.getData('text/plain') || '').replace(/\r\n?/g, '\n');
    if (text) edit(text, 'paste');
  });

  input.addEventListener('blur', () => {
    pages.forEach(({ overlay }) => overlay.classList.add('inactive'));
  });
  input.addEventListener('focus', () => {
    pages.forEach(({ overlay }) => overlay.classList.remove('inactive'));
  });

  // ---------- Public API ----------

  return {
    /** Call before re-creating page canvases */
    beginRender() {
      pages.clear();
    },

    /** Registers a freshly rendered page canvas */
    attachPage(page, card, canvas) {
      const existing = card.querySelector('.canvas-overlay');
      if (existing && canvas.dataset.editorAttached) {
        pages.set(page.page_number, { canvas, overlay: existing, page }); // unchanged page kept as is
        return;
      }
      canvas.dataset.editorAttached = '1';
      card.style.position = 'relative';
      const overlay = document.createElement('div');
      overlay.className = 'canvas-overlay';
      if (document.activeElement !== input) overlay.classList.add('inactive');
      Object.assign(overlay.style, {
        left: `${canvas.offsetLeft}px`,
        top: `${canvas.offsetTop}px`,
        width: canvas.style.width,
        height: canvas.style.height,
      });
      card.appendChild(overlay);
      pages.set(page.page_number, { canvas, overlay, page });
      canvas.addEventListener('mousedown', e => onMouseDown(e, page, card, canvas));
    },

    /** Call after all pages are rendered */
    endRender() {
      // Keep the selection inside the (possibly changed) document
      if (focus) {
        const last = Math.max(0, lastParagraphIndex());
        const clamp = pos => {
          // A paragraph that no longer exists (e.g. a text box paragraph that was deleted)
          // falls back to the end of the body
          const p = paragraph(pos.paragraph) ? pos.paragraph : Math.min(pos.paragraph, last);
          return { paragraph: p, offset: Math.min(pos.offset, chars(p).length) };
        };
        focus = clamp(focus);
        anchor = clamp(anchor);
      }
      paint({ reveal: true });
    },

    /** `{ anchor, focus, start, end, collapsed }` with start ≤ end, or null */
    selection,

    /** Focuses the editor; without a caret yet, puts it at the start of the document */
    focus() {
      if (!focus) {
        const first = env.paragraphs().list[0];
        if (!first) return;
        anchor = focus = { paragraph: first.index, offset: 0 };
        goalX = null;
        paint({ reveal: true });
      }
      input.focus({ preventScroll: true });
    },

    /** Applies `env.readOnly()`: drops a selected picture and keeps the IME / virtual keyboard away */
    readOnlyChanged() {
      input.readOnly = env.readOnly();
      if (env.readOnly()) selectImage(null);
      lastEdit = null;
      paint();
    },

    clear() {
      anchor = focus = null;
      lastEdit = null;
      selectImage(null);
      paint();
    },

    /** The selected picture `{ paragraph, index }`, or null */
    selectedImage: () => image,

    /** Selects a picture (null deselects it) and scrolls it into view */
    selectImage(ref) {
      selectImage(ref);
      paint({ reveal: true });
      input.focus({ preventScroll: true });
    },

    /** Selects a range, e.g. to restore a selection or select everything */
    select(newAnchor, newFocus = newAnchor) {
      selectImage(null);
      if (headerMode && newFocus?.paragraph < HEADER_FOOTER_BASE) leaveHeaderMode(false);
      if (newFocus?.paragraph >= HEADER_FOOTER_BASE) headerMode = true;
      anchor = newAnchor;
      focus = newFocus;
      goalX = null;
      lastEdit = null; // typing after a programmatic change starts a new undo step
      paint({ reveal: true });
    },

    undo: () => undoRedo(false),
    redo: () => undoRedo(true),

    destroy() {
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
      input.remove();
      pages.clear();
    },
  };
}
