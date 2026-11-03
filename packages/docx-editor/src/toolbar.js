// Optional formatting toolbar for a DocxEditor: undo/redo, paragraph style, font and size,
// bold/italic/underline/strikethrough, superscript/subscript, text color, highlight, clear
// formatting, alignment, line spacing, lists, indent, text wrapping of the selected picture
// and zoom. Projects with their own design can skip it and call the editor API.

const ICONS = {
  undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>',
  redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H9.5a5.5 5.5 0 0 0 0 11H13"/>',
  bold: '<path d="M6 12h9a4 4 0 0 1 0 8H6V4h8a4 4 0 0 1 0 8"/>',
  italic: '<line x1="19" x2="10" y1="4" y2="4"/><line x1="14" x2="5" y1="20" y2="20"/><line x1="15" x2="9" y1="4" y2="20"/>',
  underline: '<path d="M6 4v6a6 6 0 0 0 12 0V4"/><line x1="4" x2="20" y1="20" y2="20"/>',
  strike: '<path d="M16 4H9a3 3 0 0 0-2.83 4"/><path d="M14 12a4 4 0 0 1 0 8H6"/><line x1="4" x2="20" y1="12" y2="12"/>',
  superscript: '<path d="m4 19 8-8"/><path d="m12 19-8-8"/><path d="M20 12h-4c0-1.5.44-2 1.5-2.5S20 8.33 20 7c0-.47-.17-.93-.48-1.29a2.11 2.11 0 0 0-2.62-.44c-.42.24-.74.62-.9 1.07"/>',
  subscript: '<path d="m4 5 8 8"/><path d="m12 5-8 8"/><path d="M20 19h-4c0-1.5.44-2 1.5-2.5S20 15.33 20 14c0-.47-.17-.93-.48-1.29a2.11 2.11 0 0 0-2.62-.44c-.42.24-.74.62-.9 1.07"/>',
  highlight: '<path d="m9 11-6 6v3h9l3-3"/><path d="m22 12-4.6 4.6a2 2 0 0 1-2.8 0l-5.2-5.2a2 2 0 0 1 0-2.8L14 4"/>',
  clearFormat: '<path d="M4 7V4h16v3"/><path d="M5 20h6"/><path d="M13 4 8 20"/><path d="m15 15 5 5"/><path d="m20 15-5 5"/>',
  left: '<path d="M15 12H3"/><path d="M17 18H3"/><path d="M21 6H3"/>',
  center: '<path d="M17 12H7"/><path d="M19 18H5"/><path d="M21 6H3"/>',
  right: '<path d="M21 12H9"/><path d="M21 18H7"/><path d="M21 6H3"/>',
  both: '<path d="M3 12h18"/><path d="M3 18h18"/><path d="M3 6h18"/>',
  bulletList: '<path d="M3 12h.01"/><path d="M3 18h.01"/><path d="M3 6h.01"/><path d="M8 12h13"/><path d="M8 18h13"/><path d="M8 6h13"/>',
  numberedList: '<path d="M10 12h11"/><path d="M10 18h11"/><path d="M10 6h11"/><path d="M4 10h2"/><path d="M4 6h1v4"/><path d="M6 18H4c0-1 2-2 2-3s-1-1.5-2-1"/>',
  outdent: '<path d="M21 12H11"/><path d="M21 18H11"/><path d="M21 6H11"/><path d="m7 8-4 4 4 4"/>',
  indent: '<path d="M21 12H11"/><path d="M21 18H11"/><path d="M21 6H11"/><path d="m3 8 4 4-4 4"/>',
  zoomOut: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/><path d="M8 11h6"/>',
  zoomIn: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/><path d="M8 11h6"/><path d="M11 8v6"/>',
  chevron: '<path d="m6 9 6 6 6-6"/>',
};

const LABELS = {
  es: {
    undo: 'Deshacer', redo: 'Rehacer', bold: 'Negrita', italic: 'Cursiva', underline: 'Subrayado',
    strike: 'Tachado', superscript: 'Superíndice', subscript: 'Subíndice',
    color: 'Color de texto', highlight: 'Color de resaltado', highlightColors: 'Elegir color de resaltado',
    noHighlight: 'Sin color', clearFormat: 'Borrar formato',
    style: 'Estilo de párrafo', font: 'Fuente', fontSize: 'Tamaño de fuente',
    left: 'Alinear a la izquierda', center: 'Centrar', right: 'Alinear a la derecha', both: 'Justificar',
    lineSpacing: 'Interlineado', bulletList: 'Viñetas', numberedList: 'Numeración',
    outdent: 'Disminuir sangría', indent: 'Aumentar sangría',
    zoomOut: 'Reducir zoom', zoomIn: 'Aumentar zoom',
    imageWrap: 'Ajuste de texto de la imagen',
    wraps: {
      inline: 'En línea con el texto', square: 'Cuadrado', tight: 'Estrecho', through: 'Transparente',
      topAndBottom: 'Arriba y abajo', behind: 'Detrás del texto', inFront: 'Delante del texto',
    },
  },
  en: {
    undo: 'Undo', redo: 'Redo', bold: 'Bold', italic: 'Italic', underline: 'Underline',
    strike: 'Strikethrough', superscript: 'Superscript', subscript: 'Subscript',
    color: 'Text color', highlight: 'Highlight color', highlightColors: 'Choose highlight color',
    noHighlight: 'No color', clearFormat: 'Clear formatting',
    style: 'Paragraph style', font: 'Font', fontSize: 'Font size',
    left: 'Align left', center: 'Center', right: 'Align right', both: 'Justify',
    lineSpacing: 'Line spacing', bulletList: 'Bullets', numberedList: 'Numbering',
    outdent: 'Decrease indent', indent: 'Increase indent',
    zoomOut: 'Zoom out', zoomIn: 'Zoom in',
    imageWrap: 'Picture text wrapping',
    wraps: {
      inline: 'In line with text', square: 'Square', tight: 'Tight', through: 'Through',
      topAndBottom: 'Top and bottom', behind: 'Behind text', inFront: 'In front of text',
    },
  },
};

export const DEFAULT_TOOLBAR_ITEMS = [
  'undo', 'redo', '|', 'style', 'font', 'fontSize', '|',
  'bold', 'italic', 'underline', 'strike', 'superscript', 'subscript', 'color', 'highlight', 'clearFormat', '|',
  'left', 'center', 'right', 'both', 'lineSpacing', '|', 'bulletList', 'numberedList', 'outdent', 'indent', '|',
  'imageWrap', '|', 'zoom',
];

const WRAPS = ['inline', 'square', 'tight', 'through', 'topAndBottom', 'behind', 'inFront'];

const ZOOM_STEPS = [0.5, 0.75, 1, 1.25, 1.5, 2];

const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

const LINE_SPACINGS = [1, 1.15, 1.5, 2, 2.5, 3];

/** Word's highlight colors, most used first (the editor writes the closest one) */
const HIGHLIGHTS = ['#FFFF00', '#00FF00', '#00FFFF', '#FF00FF', '#0000FF', '#FF0000', '#000080', '#008080',
  '#008000', '#800080', '#800000', '#808000', '#808080', '#C0C0C0', '#000000'];

/** Toggle buttons and the `format` flag (from the selectionchange event) that lights them */
const TOGGLES = { bold: 'bold', italic: 'italic', underline: 'underline', strike: 'strike', superscript: 'superscript', subscript: 'subscript' };

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
  const fields = []; // selects and inputs disabled in read-only mode
  let colorInput = null;
  let zoomLabel = null;
  let wrapSelect = null;
  let styleSelect = null;
  let fontInput = null;
  let sizeInput = null;
  let spacingSelect = null;
  let highlightBar = null;
  let palette = null;
  let highlightColor = HIGHLIGHTS[0];
  let format = null;

  const svg = name => `<svg viewBox="0 0 24 24" aria-hidden="true">${ICONS[name]}</svg>`;

  const button = (name, onClick) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'docx-toolbar-button';
    b.title = text[name];
    b.setAttribute('aria-label', text[name]);
    b.innerHTML = svg(name);
    // Keep the caret (and the selection) in the document while clicking
    b.addEventListener('mousedown', e => e.preventDefault());
    b.addEventListener('click', () => editor.hasDocument && onClick());
    buttons[name] = b;
    return b;
  };

  const select = (name, options, onChange) => {
    const s = document.createElement('select');
    s.className = `docx-toolbar-select docx-toolbar-${name}`;
    s.title = text[name];
    s.setAttribute('aria-label', text[name]);
    options.forEach(([value, label]) => s.add(new Option(label, value)));
    s.addEventListener('change', () => {
      if (editor.hasDocument) onChange(s.value);
      editor.focus();
    });
    fields.push(s);
    return s;
  };

  // A text field with suggestions (any font name or size can be typed)
  const combo = (name, suggestions, onCommit) => {
    const input = document.createElement('input');
    input.className = `docx-toolbar-input docx-toolbar-${name}`;
    input.title = text[name];
    input.setAttribute('aria-label', text[name]);
    input.autocomplete = 'off';
    const list = document.createElement('datalist');
    list.id = `docx-toolbar-${name}-${Math.random().toString(36).slice(2)}`;
    input.setAttribute('list', list.id);
    const fill = values => { list.replaceChildren(...values.map(v => new Option(String(v)))); };
    fill(suggestions);
    const commit = () => {
      const value = input.value.trim();
      if (value && editor.hasDocument) onCommit(value);
      editor.focus();
    };
    input.addEventListener('change', commit);
    input.addEventListener('keydown', e => {
      if (e.key === 'Escape') editor.focus();
    });
    input.addEventListener('focus', () => input.select());
    fields.push(input);
    return { input, list, fill };
  };

  const align = value => {
    const image = editor.selectedImage;
    if (image && value !== 'both') editor.alignImage(image, value);
    else editor.setAlignment(value);
  };

  const actions = {
    undo: () => editor.undo(),
    redo: () => editor.redo(),
    bold: () => editor.toggleBold(),
    italic: () => editor.toggleItalic(),
    underline: () => editor.toggleUnderline(),
    strike: () => editor.toggleStrikethrough(),
    superscript: () => editor.toggleSuperscript(),
    subscript: () => editor.toggleSubscript(),
    clearFormat: () => editor.clearFormatting(),
    left: () => align('left'),
    center: () => align('center'),
    right: () => align('right'),
    both: () => align('both'),
    bulletList: () => editor.toggleBulletList(),
    numberedList: () => editor.toggleNumberedList(),
    outdent: () => editor.outdent(),
    indent: () => editor.indent(),
  };

  // Styles can be added to the document (Heading 1 the first time it is used): read them again
  const refreshStyles = () => {
    if (!styleSelect || !editor.hasDocument) return;
    const current = styleSelect.value;
    styleSelect.replaceChildren(...editor.paragraphStyles().map(s => new Option(s.name, s.id)));
    styleSelect.value = current;
  };

  const showStyle = id => {
    if (!styleSelect) return;
    if (id && ![...styleSelect.options].some(o => o.value === id)) refreshStyles();
    // A paragraph in the default style may carry no id, or one the gallery lists first
    styleSelect.value = [...styleSelect.options].some(o => o.value === id) ? id : styleSelect.options[0]?.value ?? '';
  };

  const closePalette = () => {
    if (palette) palette.hidden = true;
  };

  const applyHighlight = color => {
    closePalette();
    if (color) {
      highlightColor = color;
      if (highlightBar) highlightBar.style.background = color;
    }
    editor.setHighlight(color);
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
    } else if (item === 'highlight') {
      // Applies the last color (or removes it when the text already has it); the arrow opens
      // Word's highlight colors
      const group = document.createElement('span');
      group.className = 'docx-toolbar-group';
      const main = button('highlight', () => {
        applyHighlight(format?.highlight === highlightColor ? null : highlightColor);
      });
      highlightBar = document.createElement('span');
      highlightBar.className = 'docx-toolbar-swatch';
      highlightBar.style.background = highlightColor;
      main.appendChild(highlightBar);
      const more = document.createElement('button');
      more.type = 'button';
      more.className = 'docx-toolbar-button docx-toolbar-more';
      more.title = text.highlightColors;
      more.setAttribute('aria-label', text.highlightColors);
      more.setAttribute('aria-haspopup', 'true');
      more.innerHTML = svg('chevron');
      more.addEventListener('mousedown', e => e.preventDefault());
      more.addEventListener('click', () => {
        if (editor.hasDocument) palette.hidden = !palette.hidden;
      });
      buttons.highlightColors = more;
      palette = document.createElement('div');
      palette.className = 'docx-toolbar-palette';
      palette.hidden = true;
      HIGHLIGHTS.forEach(color => {
        const swatch = document.createElement('button');
        swatch.type = 'button';
        swatch.className = 'docx-toolbar-palette-color';
        swatch.style.background = color;
        swatch.title = color;
        swatch.setAttribute('aria-label', color);
        swatch.addEventListener('mousedown', e => e.preventDefault());
        swatch.addEventListener('click', () => applyHighlight(color));
        palette.appendChild(swatch);
      });
      const none = document.createElement('button');
      none.type = 'button';
      none.className = 'docx-toolbar-palette-none';
      none.textContent = text.noHighlight;
      none.addEventListener('mousedown', e => e.preventDefault());
      none.addEventListener('click', () => applyHighlight(null));
      palette.appendChild(none);
      group.append(main, more, palette);
      bar.appendChild(group);
    } else if (item === 'style') {
      styleSelect = select('style', [], id => editor.setParagraphStyle(id));
      // The list is read when opened, so styles added meanwhile show up
      styleSelect.addEventListener('mousedown', refreshStyles);
      styleSelect.addEventListener('focus', refreshStyles);
      bar.appendChild(styleSelect);
    } else if (item === 'font') {
      const font = combo('font', editor.fontFamilies?.() || [], value => editor.setFontFamily(value));
      fontInput = font.input;
      fontInput.addEventListener('focus', () => font.fill(editor.fontFamilies?.() || []));
      bar.append(font.input, font.list);
    } else if (item === 'fontSize') {
      const size = combo('fontSize', FONT_SIZES, value => editor.setFontSize(parseFloat(value.replace(',', '.'))));
      sizeInput = size.input;
      sizeInput.inputMode = 'decimal';
      bar.append(size.input, size.list);
    } else if (item === 'lineSpacing') {
      spacingSelect = select('lineSpacing', LINE_SPACINGS.map(v => [String(v), v.toLocaleString(locale)]), v => editor.setLineSpacing(Number(v)));
      bar.appendChild(spacingSelect);
    } else if (item === 'imageWrap') {
      // Enabled while a picture is selected; alignment buttons then align the picture
      wrapSelect = document.createElement('select');
      wrapSelect.className = 'docx-toolbar-select';
      wrapSelect.title = text.imageWrap;
      wrapSelect.setAttribute('aria-label', text.imageWrap);
      wrapSelect.disabled = true;
      WRAPS.forEach(wrap => wrapSelect.add(new Option(text.wraps[wrap], wrap)));
      wrapSelect.addEventListener('change', () => {
        const image = editor.hasDocument && editor.selectedImage;
        if (image) editor.setImageWrap(image, wrapSelect.value);
        editor.focus();
      });
      bar.appendChild(wrapSelect);
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
    format = e.detail.format;
    Object.entries(TOGGLES).forEach(([name, flag]) => buttons[name]?.classList.toggle('active', !!format?.[flag]));
    ['left', 'center', 'right', 'both'].forEach(a => buttons[a]?.classList.toggle('active', format?.align === a));
    buttons.bulletList?.classList.toggle('active', format?.list === 'bullet');
    buttons.numberedList?.classList.toggle('active', format?.list === 'number');
    buttons.highlight?.classList.toggle('active', !!format?.highlight);
    if (colorInput && format?.color) colorInput.value = format.color;
    // Fields being edited keep what the user is typing
    if (fontInput && document.activeElement !== fontInput) fontInput.value = format?.fontFamily || '';
    if (sizeInput && document.activeElement !== sizeInput) sizeInput.value = format?.fontSize ? String(format.fontSize) : '';
    if (format) showStyle(format.style);
    if (spacingSelect) {
      const spacing = format?.lineSpacing ?? 1;
      const known = LINE_SPACINGS.find(v => Math.abs(v - spacing) < 0.01);
      spacingSelect.value = known != null ? String(known) : '';
    }
  };
  // Read-only: only zoom stays usable
  const applyReadOnly = () => {
    const readOnly = editor.readOnly;
    Object.entries(buttons).forEach(([name, b]) => {
      if (name !== 'zoomIn' && name !== 'zoomOut') b.disabled = readOnly;
    });
    fields.forEach(f => { f.disabled = readOnly; });
    if (colorInput) colorInput.disabled = readOnly;
    if (wrapSelect) wrapSelect.disabled = readOnly || !editor.selectedImage;
    if (readOnly) closePalette();
    bar.classList.toggle('read-only', readOnly);
  };

  const onImage = e => {
    if (!wrapSelect) return;
    const image = e.detail.image;
    wrapSelect.disabled = !image || editor.readOnly;
    if (image) wrapSelect.value = image.wrap;
  };
  const onLoad = () => refreshStyles();
  const onOutside = e => {
    if (palette && !palette.hidden && !palette.parentElement.contains(e.target)) closePalette();
  };
  editor.addEventListener('selectionchange', onSelection);
  editor.addEventListener('imageselect', onImage);
  editor.addEventListener('readonlychange', applyReadOnly);
  editor.addEventListener('load', onLoad);
  document.addEventListener('mousedown', onOutside);
  applyReadOnly();
  refreshStyles();
  container.appendChild(bar);

  return {
    destroy() {
      editor.removeEventListener('selectionchange', onSelection);
      editor.removeEventListener('imageselect', onImage);
      editor.removeEventListener('readonlychange', applyReadOnly);
      editor.removeEventListener('load', onLoad);
      document.removeEventListener('mousedown', onOutside);
      bar.remove();
    },
  };
}
