// Canvas-native editing: the caret, the selection and keyboard input live on top of the
// rendered pages, so there is no per-paragraph edit box. Rust owns the geometry (hit testing,
// caret boxes, vertical moves), the edits and the undo history; this module routes input and
// paints overlays.
//
// Positions are { paragraph, offset } with offsets in Unicode code points of the paragraph
// text, the same unit Rust uses. A selection runs from `anchor` (where it started) to `focus`
// (where the caret is) and may span paragraphs.

const WORD_CHAR = /[\p{L}\p{N}_]/u;
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
 */
export function createCanvasEditor(env) {
  let anchor = null;
  let focus = null;
  let goalX = null;        // remembered column for ↑/↓
  let dragging = false;
  let composing = false;
  let lastEdit = null;     // { kind, caret, time } for grouping keystrokes into one undo step
  const pages = new Map(); // page number → { canvas, overlay }

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
  const hitTest = (page, x, y) => parse(session()?.hit_test(page, x, y, env.measure));
  const caretBox = pos => parse(session()?.caret_box(pos.paragraph, pos.offset, env.measure));
  const moveVertical = (pos, dir, x) => parse(session()?.move_vertical(pos.paragraph, pos.offset, dir, x, env.measure));
  const selectionRects = (a, b) =>
    parse(session()?.selection_rects_range(a.paragraph, a.offset, b.paragraph, b.offset, env.measure)) || [];

  // ---------- Text helpers ----------

  function paragraph(index) {
    return env.paragraphs().byIndex.get(index);
  }

  function chars(index) {
    const p = paragraph(index);
    return p ? Array.from(p.text || '') : [];
  }

  function lastParagraphIndex() {
    const list = env.paragraphs().list;
    return list.length ? list[list.length - 1].index : -1;
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
    if (!current || current.container === 'body') return null;
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
      return pos.paragraph > 0 ? { paragraph: pos.paragraph - 1, offset: chars(pos.paragraph - 1).length } : pos;
    }
    if (pos.offset < text.length) {
      return { paragraph: pos.paragraph, offset: byWord ? wordEnd(text, pos.offset) : pos.offset + 1 };
    }
    return pos.paragraph < lastParagraphIndex() ? { paragraph: pos.paragraph + 1, offset: 0 } : pos;
  }

  // ---------- Editing ----------

  /**
   * Replaces `range` (default: the selection) with `text` through Rust's single edit primitive.
   * `kind` decides undo grouping: consecutive typing (or deleting) at the caret joins one step,
   * starting a new step at each space so undo goes back word by word.
   */
  function edit(text, kind, range = ordered()) {
    if (!session() || !focus) return;
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
    if (prev < 0) return;
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
    if (next > lastParagraphIndex()) return;
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

      // Keep the hidden input at the caret so IME candidate windows open in place
      const rect = page.canvas.getBoundingClientRect();
      input.style.left = `${rect.left + box.x * zoom}px`;
      input.style.top = `${rect.top + box.y * zoom}px`;
      // Only a visible caret can be scrolled to (with a selection it is hidden)
      if (reveal && collapsed()) caret.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
    env.selectionChanged(selection());
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
    const pos = hitTest(page.page_number, pt.x, pt.y);
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

  function onMouseMove(e) {
    if (!dragging || !(e.buttons & 1)) {
      dragging = false;
      return;
    }
    const pos = positionAt(e.clientX, e.clientY);
    if (pos) {
      setCaret(pos, true);
      paint();
    }
  }
  function onMouseUp() {
    dragging = false;
  }
  document.addEventListener('mousemove', onMouseMove);
  document.addEventListener('mouseup', onMouseUp);

  // ---------- Keyboard input ----------

  input.addEventListener('keydown', e => {
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
          const target = dir < 0
            ? { paragraph: 0, offset: 0 }
            : { paragraph: lastParagraphIndex(), offset: chars(lastParagraphIndex()).length };
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
        } else if (paragraph(focus.paragraph)?.container !== 'body') {
          // Last cell (or first with Shift): stay put rather than typing a tab
        } else {
          edit('\t', 'type');
          moved = false;
        }
        break;
      }
      case 'Escape':
        anchor = focus = null;
        input.blur();
        break;
      case 'a':
        if (!mod) { handled = false; break; }
        anchor = { paragraph: 0, offset: 0 };
        focus = { paragraph: lastParagraphIndex(), offset: chars(lastParagraphIndex()).length };
        break;
      case 'z':
        if (!mod) { handled = false; break; }
        undoRedo(e.shiftKey);
        moved = false;
        break;
      case 'y':
        if (!mod || IS_MAC) { handled = false; break; }
        undoRedo(true);
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
    edit('', 'cut');
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
        pages.set(page.page_number, { canvas, overlay: existing }); // unchanged page kept as is
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
      pages.set(page.page_number, { canvas, overlay });
      canvas.addEventListener('mousedown', e => onMouseDown(e, page, card, canvas));
    },

    /** Call after all pages are rendered */
    endRender() {
      // Keep the selection inside the (possibly changed) document
      if (focus) {
        const last = Math.max(0, lastParagraphIndex());
        const clamp = pos => {
          const p = Math.min(pos.paragraph, last);
          return { paragraph: p, offset: Math.min(pos.offset, chars(p).length) };
        };
        focus = clamp(focus);
        anchor = clamp(anchor);
      }
      paint({ reveal: true });
    },

    /** `{ anchor, focus, start, end, collapsed }` with start ≤ end, or null */
    selection,

    focus() {
      if (focus) input.focus({ preventScroll: true });
    },

    clear() {
      anchor = focus = null;
      lastEdit = null;
      paint();
    },

    /** Selects a range, e.g. to restore a selection or select everything */
    select(newAnchor, newFocus = newAnchor) {
      anchor = newAnchor;
      focus = newFocus;
      goalX = null;
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
