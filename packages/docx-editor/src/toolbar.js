// Optional formatting toolbar for a DocxEditor: undo/redo, bold/italic/underline, text color,
// alignment and zoom. Projects with their own design can skip it and call the editor API.

const ICONS = {
  undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>',
  redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 0 0 0 11H13"/>',
  bold: '<path d="M6 12h9a4 4 0 0 1 0 8H6V4h8a4 4 0 0 1 0 8"/>',
  italic: '<line x1="19" x2="10" y1="4" y2="4"/><line x1="14" x2="5" y1="20" y2="20"/><line x1="15" x2="9" y1="4" y2="20"/>',
  underline: '<path d="M6 4v6a6 6 0 0 0 12 0V4"/><line x1="4" x2="20" y1="20" y2="20"/>',
  left: '<path d="M15 12H3"/><path d="M17 18H3"/><path d="M21 6H3"/>',
  center: '<path d="M17 12H7"/><path d="M19 18H5"/><path d="M21 6H3"/>',
  right: '<path d="M21 12H9"/><path d="M21 18H7"/><path d="M21 6H3"/>',
  both: '<path d="M3 12h18"/><path d="M3 18h18"/><path d="M3 6h18"/>',
  zoomOut: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/><path d="M8 11h6"/>',
  zoomIn: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/><path d="M8 11h6"/><path d="M11 8v6"/>',
};

const LABELS = {
  es: {
    undo: 'Deshacer', redo: 'Rehacer', bold: 'Negrita', italic: 'Cursiva', underline: 'Subrayado',
    color: 'Color de texto', left: 'Alinear a la izquierda', center: 'Centrar', right: 'Alinear a la derecha',
    both: 'Justificar', zoomOut: 'Reducir zoom', zoomIn: 'Aumentar zoom',
  },
  en: {
    undo: 'Undo', redo: 'Redo', bold: 'Bold', italic: 'Italic', underline: 'Underline',
    color: 'Text color', left: 'Align left', center: 'Center', right: 'Align right',
    both: 'Justify', zoomOut: 'Zoom out', zoomIn: 'Zoom in',
  },
};

export const DEFAULT_TOOLBAR_ITEMS = ['undo', 'redo', '|', 'bold', 'italic', 'underline', 'color', '|', 'left', 'center', 'right', 'both', '|', 'zoom'];

const ZOOM_STEPS = [0.5, 0.75, 1, 1.25, 1.5, 2];

/**
 * Renders a toolbar for `editor` inside `container`.
 * @param {import('./docx_editor.js').DocxEditor} editor
 * @param {HTMLElement} container
 * @param {{items?: string[], locale?: 'es'|'en', labels?: object}} [options]
 *        items: any of DEFAULT_TOOLBAR_ITEMS, '|' draws a separator
 * @returns {{destroy(): void}}
 */
export function createDocxToolbar(editor, container, { items = DEFAULT_TOOLBAR_ITEMS, locale = editor.options.locale, labels } = {}) {
  const text = { ...(LABELS[locale] || LABELS.en), ...labels };
  const bar = document.createElement('div');
  bar.className = 'docx-toolbar';
  bar.setAttribute('role', 'toolbar');
  const buttons = {};
  let colorInput = null;
  let zoomLabel = null;

  const button = (name, onClick) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'docx-toolbar-button';
    b.title = text[name];
    b.setAttribute('aria-label', text[name]);
    b.innerHTML = `<svg viewBox="0 0 24 24" aria-hidden="true">${ICONS[name]}</svg>`;
    // Keep the caret (and the selection) in the document while clicking
    b.addEventListener('mousedown', e => e.preventDefault());
    b.addEventListener('click', () => editor.hasDocument && onClick());
    buttons[name] = b;
    return b;
  };

  const actions = {
    undo: () => editor.undo(),
    redo: () => editor.redo(),
    bold: () => editor.toggleBold(),
    italic: () => editor.toggleItalic(),
    underline: () => editor.toggleUnderline(),
    left: () => editor.setAlignment('left'),
    center: () => editor.setAlignment('center'),
    right: () => editor.setAlignment('right'),
    both: () => editor.setAlignment('both'),
  };

  items.forEach(item => {
    if (item === '|') {
      const sep = document.createElement('span');
      sep.className = 'docx-toolbar-separator';
      bar.appendChild(sep);
    } else if (item === 'color') {
      const label = document.createElement('label');
      label.className = 'docx-toolbar-color';
      label.title = text.color;
      colorInput = document.createElement('input');
      colorInput.type = 'color';
      colorInput.value = '#000000';
      colorInput.setAttribute('aria-label', text.color);
      colorInput.addEventListener('input', () => editor.hasDocument && editor.setColor(colorInput.value, { refocus: false }));
      label.appendChild(colorInput);
      bar.appendChild(label);
    } else if (item === 'zoom') {
      const zoom = factor => {
        const steps = factor > 0 ? ZOOM_STEPS : [...ZOOM_STEPS].reverse();
        const next = steps.find(z => (factor > 0 ? z > editor.zoom + 1e-6 : z < editor.zoom - 1e-6));
        if (next) editor.setZoom(next);
        zoomLabel.textContent = `${Math.round(editor.zoom * 100)}%`;
      };
      bar.appendChild(button('zoomOut', () => zoom(-1)));
      zoomLabel = document.createElement('span');
      zoomLabel.className = 'docx-toolbar-zoom';
      zoomLabel.textContent = `${Math.round(editor.zoom * 100)}%`;
      bar.appendChild(zoomLabel);
      bar.appendChild(button('zoomIn', () => zoom(1)));
    } else if (actions[item]) {
      bar.appendChild(button(item, actions[item]));
    }
  });

  const onSelection = e => {
    const format = e.detail.format;
    ['bold', 'italic', 'underline'].forEach(f => buttons[f]?.classList.toggle('active', !!format?.[f]));
    ['left', 'center', 'right', 'both'].forEach(a => buttons[a]?.classList.toggle('active', format?.align === a));
    if (colorInput && format?.color) colorInput.value = format.color;
  };
  editor.addEventListener('selectionchange', onSelection);
  container.appendChild(bar);

  return {
    destroy() {
      editor.removeEventListener('selectionchange', onSelection);
      bar.remove();
    },
  };
}
