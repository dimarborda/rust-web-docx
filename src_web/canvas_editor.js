// Canvas-native editing: the caret, the selection and keyboard input live on top of the
// rendered pages, so there is no per-paragraph edit box. Rust owns the geometry (hit testing,
// caret boxes, vertical moves) and the edits; this module routes input and paints overlays.
//
// Positions are { paragraph, offset } with offsets in Unicode code points of the paragraph
// text, the same unit Rust uses. In this phase a selection stays inside one paragraph.

const WORD_CHAR = /[\p{L}\p{N}_]/u;
const IS_MAC = /Mac|iPhone|iPad/.test(navigator.platform);

/**
 * @param {object} env
 * @param {() => any} env.session            current DocxSession
 * @param {() => any[]} env.elements         document elements (paragraphs with `text`, `runs`)
 * @param {() => number} env.zoom
 * @param {Function} env.measure             text measurer passed to Rust
 * @param {() => void} env.documentChanged   re-render after an edit
 * @param {(sel) => void} env.selectionChanged
 * @param {(page, point, card, canvas) => boolean} env.handleTableClick  true when a cell took the click
 * @param {(msg: string) => void} env.hint
 */
export function createCanvasEditor(env) {
  let anchor = null;
  let focus = null;
  let goalX = null;        // remembered column for ↑/↓
  let dragging = false;
  let composing = false;
  const pages = new Map(); // page number → { canvas, overlay }

  // A hidden textarea owns keyboard focus: it receives typing, IME/dead-key composition
  // (´ + e = é) and clipboard events
  const input = document.createElement('textarea');
  input.className = 'canvas-editor-input';
  input.setAttribute('autocomplete', 'off');
  input.setAttribute('autocorrect', 'off');
  input.setAttribute('autocapitalize', 'off');
  input.setAttribute('spellcheck', 'false');
  input.setAttribute('aria-label', 'Editor del documento');
  document.body.appendChild(input);

  // ---------- Geometry (Rust) ----------

  const parse = json => (json ? JSON.parse(json) : null);
  const session = () => env.session();
  const hitTest = (page, x, y) => parse(session()?.hit_test(page, x, y, env.measure));
  const caretBox = pos => parse(session()?.caret_box(pos.paragraph, pos.offset, env.measure));
  const moveVertical = (pos, dir, x) => parse(session()?.move_vertical(pos.paragraph, pos.offset, dir, x, env.measure));
  const selectionRects = (p, s, e) => parse(session()?.selection_rects(p, s, e, env.measure)) || [];

  // ---------- Text helpers ----------

  function paragraph(index) {
    return env.elements().find(el => el.type === 'paragraph' && el.index === index);
  }

  function chars(index) {
    const p = paragraph(index);
    return p ? Array.from(p.text || '') : [];
  }

  function lastParagraphIndex() {
    let last = -1;
    env.elements().forEach(el => {
      if (el.type === 'paragraph') last = Math.max(last, el.index);
    });
    return last;
  }

  function selection() {
    if (!anchor || !focus) return null;
    return {
      paragraph: focus.paragraph,
      start: Math.min(anchor.offset, focus.offset),
      end: Math.max(anchor.offset, focus.offset),
    };
  }

  function selectedText() {
    const sel = selection();
    return sel ? chars(sel.paragraph).slice(sel.start, sel.end).join('') : '';
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
      // Phase 1: selections stay within the anchor's paragraph
      if (pos.paragraph !== anchor.paragraph) {
        pos = {
          paragraph: anchor.paragraph,
          offset: pos.paragraph > anchor.paragraph ? chars(anchor.paragraph).length : 0,
        };
      }
      focus = pos;
    } else {
      anchor = pos;
      focus = pos;
    }
  }

  function collapsed() {
    return !anchor || !focus || (anchor.paragraph === focus.paragraph && anchor.offset === focus.offset);
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

  function replaceSelection(text) {
    const sel = selection();
    if (!sel || !session()) return;
    try {
      session().replace_text(sel.paragraph, sel.start, sel.end, text);
    } catch (err) {
      console.error('Edit failed:', err);
      env.hint(String(err));
      return;
    }
    const caret = { paragraph: sel.paragraph, offset: sel.start + Array.from(text).length };
    anchor = focus = caret;
    goalX = null;
    env.documentChanged();
  }

  function deleteBackward(byWord) {
    if (!collapsed()) return replaceSelection('');
    const text = chars(focus.paragraph);
    if (focus.offset === 0) {
      env.hint('Unir párrafos con Retroceso llega en la fase 2.');
      return;
    }
    anchor = { paragraph: focus.paragraph, offset: byWord ? wordStart(text, focus.offset) : focus.offset - 1 };
    replaceSelection('');
  }

  function deleteForward(byWord) {
    if (!collapsed()) return replaceSelection('');
    const text = chars(focus.paragraph);
    if (focus.offset >= text.length) {
      env.hint('Unir párrafos con Suprimir llega en la fase 2.');
      return;
    }
    anchor = { paragraph: focus.paragraph, offset: byWord ? wordEnd(text, focus.offset) : focus.offset + 1 };
    replaceSelection('');
  }

  // ---------- Painting ----------

  function paint({ reveal = false } = {}) {
    pages.forEach(({ overlay }) => overlay.replaceChildren());
    if (!focus) {
      env.selectionChanged(null);
      return;
    }
    const zoom = env.zoom();
    const sel = selection();

    if (sel && sel.start !== sel.end) {
      selectionRects(sel.paragraph, sel.start, sel.end).forEach(r => {
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
    env.selectionChanged(sel);
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
    const pt = pagePoint(el, { clientX, clientY });
    return hitTest(Number(el.dataset.pageNum), pt.x, pt.y);
  }

  function onMouseDown(e, page, card, canvas) {
    if (e.button !== 0) return;
    const pt = pagePoint(canvas, e);
    if (env.handleTableClick(page, pt, card, canvas)) {
      // Keep focus in the cell editor that just opened
      e.preventDefault();
      anchor = focus = null;
      paint();
      return;
    }
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
    input.focus({ preventScroll: true });
    paint();
  }

  document.addEventListener('mousemove', e => {
    if (!dragging || !(e.buttons & 1)) {
      dragging = false;
      return;
    }
    const pos = positionAt(e.clientX, e.clientY);
    if (pos) {
      setCaret(pos, true);
      paint();
    }
  });
  document.addEventListener('mouseup', () => {
    dragging = false;
  });

  // ---------- Keyboard input ----------

  input.addEventListener('keydown', e => {
    if (!focus || composing || e.isComposing) return;
    const mod = e.metaKey || e.ctrlKey;
    const byWord = IS_MAC ? e.altKey : e.ctrlKey;
    const lineEdge = IS_MAC ? e.metaKey : false;
    let handled = true;

    switch (e.key) {
      case 'ArrowLeft':
      case 'ArrowRight': {
        const dir = e.key === 'ArrowLeft' ? -1 : 1;
        if (lineEdge) {
          const box = caretBox(focus);
          if (box) setCaret({ paragraph: focus.paragraph, offset: dir < 0 ? box.line_start : box.line_end }, e.shiftKey);
        } else if (!collapsed() && !e.shiftKey) {
          const sel = selection();
          setCaret({ paragraph: sel.paragraph, offset: dir < 0 ? sel.start : sel.end });
        } else {
          setCaret(stepHorizontal(focus, dir, byWord), e.shiftKey);
        }
        goalX = null;
        break;
      }
      case 'ArrowUp':
      case 'ArrowDown': {
        const dir = e.key === 'ArrowUp' ? -1 : 1;
        if (lineEdge) {
          const target = dir < 0
            ? { paragraph: 0, offset: 0 }
            : { paragraph: lastParagraphIndex(), offset: chars(lastParagraphIndex()).length };
          setCaret(target, e.shiftKey);
          break;
        }
        if (goalX === null) goalX = caretBox(focus)?.x ?? 0;
        const next = moveVertical(focus, dir, goalX);
        // At the first/last line, ↑/↓ go to the start/end of that line
        const fallback = { paragraph: focus.paragraph, offset: dir < 0 ? 0 : chars(focus.paragraph).length };
        setCaret(next || fallback, e.shiftKey);
        break;
      }
      case 'Home':
      case 'End': {
        const box = caretBox(focus);
        if (box) setCaret({ paragraph: focus.paragraph, offset: e.key === 'Home' ? box.line_start : box.line_end }, e.shiftKey);
        goalX = null;
        break;
      }
      case 'Backspace':
        deleteBackward(byWord);
        break;
      case 'Delete':
        deleteForward(byWord);
        break;
      case 'Enter':
        if (e.shiftKey) {
          replaceSelection('\n');
        } else {
          env.hint('Partir párrafos con Enter llega en la fase 2; Shift+Enter inserta un salto de línea.');
        }
        break;
      case 'Tab':
        replaceSelection('\t');
        break;
      case 'Escape':
        anchor = focus = null;
        input.blur();
        break;
      default:
        if (mod && e.key.toLowerCase() === 'a') {
          anchor = { paragraph: focus.paragraph, offset: 0 };
          focus = { paragraph: focus.paragraph, offset: chars(focus.paragraph).length };
        } else if (mod && e.key.toLowerCase() === 'z') {
          env.hint('Deshacer llega en la fase 2.');
        } else {
          handled = false; // typing and clipboard arrive as input / clipboard events
        }
    }

    if (handled) {
      e.preventDefault();
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
    if (text && focus) replaceSelection(text);
  });

  input.addEventListener('input', e => {
    if (composing || e.isComposing) return;
    const text = input.value;
    input.value = '';
    if (text && focus) replaceSelection(text);
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
    replaceSelection('');
  });

  input.addEventListener('paste', e => {
    e.preventDefault();
    if (!focus) return;
    // Phase 1: pasted line breaks become soft breaks inside the paragraph
    const text = (e.clipboardData.getData('text/plain') || '').replace(/\r\n?/g, '\n');
    if (text) replaceSelection(text);
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
      // Keep the caret inside the (possibly shorter) document
      if (focus) {
        const clamp = pos => ({ paragraph: pos.paragraph, offset: Math.min(pos.offset, chars(pos.paragraph).length) });
        focus = clamp(focus);
        anchor = clamp(anchor);
      }
      paint({ reveal: true });
    },

    /** Current selection `{paragraph, start, end}` (collapsed when start === end), or null */
    selection,

    focus() {
      if (focus) input.focus({ preventScroll: true });
    },

    clear() {
      anchor = focus = null;
      paint();
    },
  };
}
