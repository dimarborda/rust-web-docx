import init, { DocxSession } from '../pkg/rust_web_docx.js';

// Application State
let wasmReady = false;
let currentSession = null;
let currentFileName = 'documento.docx';
let activeDocumentElements = [];
let activeParagraphIndex = null;
let activeTarget = null; // { type: 'paragraph' | 'cell', paragraphIndex?, tableIndex?, row?, col?, item? }
let detectedVariables = [];
let updateDebounceTimer = null;
let currentZoom = 1.0; // 0.75, 1.0, 1.25, 1.5
let canvasPagesLayout = null;

let currentWatermark = {
  type: 'none', // 'none' | 'preset'
  text: '',
  opacity: 0.2
};
let currentPageBgColor = '#FFFFFF';

// DOM Elements - Shell & Nav
const emptyStateView = document.getElementById('empty-state-view');
const canvasDocumentView = document.getElementById('canvas-document-view');
const canvasPagesWrapper = document.getElementById('canvas-pages-wrapper');
const canvasParagraphEditor = document.getElementById('canvas-paragraph-editor');
const canvasCellEditor = document.getElementById('canvas-cell-editor');

const docStatusContainer = document.getElementById('doc-status-container');
const docToolbar = document.getElementById('doc-toolbar');
const formattingRibbon = document.getElementById('formatting-ribbon');
const appStatusbar = document.getElementById('app-statusbar');
const currentDocTitle = document.getElementById('current-doc-title');
const syncStatus = document.getElementById('sync-status');

// Zoom Controls
const btnZoomIn = document.getElementById('btn-zoom-in');
const btnZoomOut = document.getElementById('btn-zoom-out');
const zoomLevelLabel = document.getElementById('zoom-level-label');

const dropTarget = document.getElementById('drop-target');

// Inputs & Buttons
const fileInput = document.getElementById('file-input');
const fileInputWelcome = document.getElementById('file-input-welcome');
const btnLoadSample = document.getElementById('btn-load-sample');
const btnWelcomeDemo = document.getElementById('btn-welcome-demo');
const btnDownload = document.getElementById('btn-download');

// Search & Replace
const searchReplaceBar = document.getElementById('search-replace-bar');
const btnToggleFind = document.getElementById('btn-toggle-find');
const btnCloseSearch = document.getElementById('btn-close-search');
const findQuery = document.getElementById('find-query');
const replaceQuery = document.getElementById('replace-query');
const chkCaseSensitive = document.getElementById('chk-case-sensitive');
const chkRegex = document.getElementById('chk-regex');
const btnDoReplace = document.getElementById('btn-do-replace');

// Variables Modal
const variablesModal = document.getElementById('variables-modal');
const btnToggleVariables = document.getElementById('btn-toggle-variables');
const btnCloseVariables = document.getElementById('btn-close-variables');
const varsCountBadge = document.getElementById('vars-count-badge');
const variablesListContainer = document.getElementById('variables-list-container');
const btnAddCustomVar = document.getElementById('btn-add-custom-var');
const btnApplyAllVariables = document.getElementById('btn-apply-all-variables');

// Background & Watermark Modal
const backgroundModal = document.getElementById('background-modal');
const btnToggleBgModal = document.getElementById('btn-toggle-background-modal');
const btnCloseBgModal = document.getElementById('btn-close-bg-modal');
const watermarkOpacity = document.getElementById('watermark-opacity');
const opacityValLabel = document.getElementById('opacity-val-label');
const btnSaveBgConfig = document.getElementById('btn-save-bg-config');
const btnWmConfidential = document.getElementById('btn-wm-confidential');
const btnWmDraft = document.getElementById('btn-wm-draft');
const btnWmClear = document.getElementById('btn-wm-clear');
const bgPresetBtns = document.querySelectorAll('.bg-preset-btn');

// Table Modal
const tableModal = document.getElementById('table-modal');
const btnInsertTable = document.getElementById('btn-insert-table');
const btnCloseTableModal = document.getElementById('btn-close-table-modal');
const btnCreateTableConfirm = document.getElementById('btn-create-table-confirm');
const tblInputRows = document.getElementById('tbl-input-rows');
const tblInputCols = document.getElementById('tbl-input-cols');

// Formatting Ribbon Buttons
const btnFmtBold = document.getElementById('btn-fmt-bold');
const btnFmtItalic = document.getElementById('btn-fmt-italic');
const btnFmtUnderline = document.getElementById('btn-fmt-underline');
const textColorInput = document.getElementById('text-color-input');
const colorSwatchDots = document.querySelectorAll('.color-swatch-dot');
const alignButtons = {
  left: document.getElementById('btn-align-left'),
  center: document.getElementById('btn-align-center'),
  right: document.getElementById('btn-align-right'),
  both: document.getElementById('btn-align-justify'),
};

// Live Statistics
const statWords = document.getElementById('stat-words');
const statChars = document.getElementById('stat-chars');
const statParagraphs = document.getElementById('stat-paragraphs');
const statTables = document.getElementById('stat-tables');
const statPages = document.getElementById('stat-pages');
const statFormat = document.getElementById('stat-format');

// Toast
const toast = document.getElementById('toast');
const toastText = document.getElementById('toast-text');

// 1. Initialize WebAssembly
async function initializeWasm() {
  if (wasmReady) return;
  try {
    await init();
    wasmReady = true;
    console.log('⚡ Rust WebAssembly DOCX High-Fidelity Engine loaded successfully.');
  } catch (err) {
    console.error('Failed to init WebAssembly:', err);
    showToast('Error al inicializar WebAssembly: ' + err, true);
  }
}

// 2. Load File into WASM Session
async function loadDocxFile(file) {
  await initializeWasm();
  try {
    const arrayBuffer = await file.arrayBuffer();
    const bytes = new Uint8Array(arrayBuffer);
    currentFileName = file.name;
    
    currentSession = new DocxSession(bytes);
    onDocumentLoaded();
    showToast(`"${currentFileName}" cargado con vista fiel multi-página.`);
  } catch (err) {
    console.error('Error opening DOCX:', err);
    showToast('Error al abrir el documento DOCX: ' + err, true);
  }
}

// Load Demo Sample File
async function loadSampleDocx() {
  await initializeWasm();
  try {
    currentFileName = 'contrato_ejemplo.docx';
    currentSession = DocxSession.new_sample();
    onDocumentLoaded();
    showToast('Documento demo cargado en vista fiel multi-página.');
  } catch (err) {
    console.error('Error generating sample:', err);
    showToast('Error al cargar documento demo: ' + err, true);
  }
}

// Load Real Document from examples/
async function loadExampleDocx(filename) {
  await initializeWasm();
  try {
    showToast(`Cargando "${filename}"...`);
    const res = await fetch(`/examples/${encodeURIComponent(filename)}`);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const arrayBuffer = await res.arrayBuffer();
    const bytes = new Uint8Array(arrayBuffer);
    currentFileName = filename;
    currentSession = new DocxSession(bytes);
    onDocumentLoaded();
    showToast(`"${filename}" cargado con éxito en vista multi-página.`);
  } catch (err) {
    console.error('Error loading real example:', err);
    showToast('Error al cargar archivo de ejemplo: ' + err, true);
  }
}

// 3. Document Loaded Transition
function onDocumentLoaded() {
  if (!currentSession) return;

  emptyStateView.style.display = 'none';
  canvasDocumentView.style.display = 'flex';
  docStatusContainer.style.display = 'flex';
  docToolbar.style.display = 'flex';
  formattingRibbon.style.display = 'flex';
  appStatusbar.style.display = 'flex';
  btnDownload.disabled = false;

  currentDocTitle.textContent = currentFileName;

  refreshDocumentView();
}

// 4. Refresh Document Multi-Page Canvas View
function refreshDocumentView() {
  if (!currentSession) return;

  try {
    // Sync local element cache
    const elementsJson = currentSession.get_document_elements_json();
    activeDocumentElements = JSON.parse(elementsJson);

    renderCanvasPagesFromWasm();
    updateLiveStats();
    scanAndHighlightVariables();
  } catch (err) {
    console.error('Error refreshing document view:', err);
  }
}

// 5. Render Multi-Page A4 Canvas Layout from Rust WASM
function renderCanvasPagesFromWasm() {
  if (!currentSession) return;

  try {
    const elementsJson = currentSession.get_document_elements_json();
    activeDocumentElements = JSON.parse(elementsJson);

    const wmText = currentWatermark.type === 'preset' ? currentWatermark.text : null;
    const wmOpacity = currentWatermark.opacity || 0.2;

    const layoutJson = currentSession.compute_canvas_layout_json(wmText, wmOpacity, measureTextForLayout);
    canvasPagesLayout = JSON.parse(layoutJson);

    statPages.textContent = canvasPagesLayout.total_pages;
    canvasPagesWrapper.innerHTML = '';

    const dpr = window.devicePixelRatio || 1;

    canvasPagesLayout.pages.forEach(page => {
      const pageCard = document.createElement('div');
      pageCard.className = 'canvas-page-card';
      pageCard.dataset.pageNumber = page.page_number;

      const pageHeader = document.createElement('div');
      pageHeader.className = 'canvas-page-header';
      pageHeader.innerHTML = `<span class="page-badge">Página ${page.page_number} de ${canvasPagesLayout.total_pages}</span>`;
      pageCard.appendChild(pageHeader);

      const canvas = document.createElement('canvas');
      canvas.className = 'page-canvas';
      canvas.dataset.pageNum = page.page_number;

      // Scaled dimensions with Retina High-DPI support
      const scaledW = page.width * currentZoom;
      const scaledH = page.height * currentZoom;

      canvas.width = Math.round(scaledW * dpr);
      canvas.height = Math.round(scaledH * dpr);
      canvas.style.width = `${scaledW}px`;
      canvas.style.height = `${scaledH}px`;

      const ctx = canvas.getContext('2d');
      ctx.scale(dpr * currentZoom, dpr * currentZoom);

      // 1. Draw page background
      ctx.fillStyle = page.bg_color || currentPageBgColor || '#FFFFFF';
      ctx.fillRect(0, 0, page.width, page.height);

      // 2. Draw page content items
      drawCanvasPageItems(ctx, page.items);

      // 3. Click handler for in-place paragraph / table cell block editing
      canvas.addEventListener('click', (e) => {
        handleCanvasClick(e, page, canvas, pageCard);
      });

      pageCard.appendChild(canvas);
      canvasPagesWrapper.appendChild(pageCard);
    });

    ensureLayoutFonts(canvasPagesLayout);

  } catch (err) {
    console.error('Error rendering Canvas layout:', err);
    showToast('Error al renderizar páginas: ' + err, true);
  }
}

// Rust lays out text with these measurements, so line breaks match what the canvas draws
const layoutMeasureCtx = document.createElement('canvas').getContext('2d');

function measureTextForLayout(text, family, size, bold, italic) {
  layoutMeasureCtx.font = buildCanvasFont(bold ? '700' : '400', italic ? 'italic' : 'normal', size, family);
  return layoutMeasureCtx.measureText(text).width;
}

// Web fonts (Carlito, Arimo, Tinos...) load lazily: request every face the layout uses and
// lay out again once they arrive, since measurements taken with a fallback font are wrong
const requestedFontFaces = new Set();

function ensureLayoutFonts(layout) {
  const faces = new Set();
  layout.pages.forEach(page => page.items.forEach(item => {
    if (item.type === 'text' && item.runs) {
      item.runs.forEach(run => faces.add(
        buildCanvasFont(run.bold ? '700' : '400', run.italic ? 'italic' : 'normal', 16, run.font_family || item.font_family)
      ));
    } else if (item.type === 'table_cell') {
      faces.add(buildCanvasFont(item.font_weight, 'normal', 16, item.font_family));
    }
  }));

  const pending = [...faces].filter(face => !requestedFontFaces.has(face));
  if (pending.length === 0) return;
  pending.forEach(face => requestedFontFaces.add(face));

  Promise.all(pending.map(face => document.fonts.load(face).catch(() => [])))
    .then(results => {
      if (results.some(loaded => loaded.length > 0)) renderCanvasPagesFromWasm();
    });
}

function formatFontFamily(family) {
  if (!family) return '"Calibri", "Carlito", "Segoe UI", Inter, sans-serif';
  const fLower = family.toLowerCase();
  if (fLower.includes('calibri')) {
    return '"Calibri", "Carlito", "Segoe UI", Roboto, sans-serif';
  } else if (fLower.includes('consolas') || fLower.includes('courier') || fLower.includes('mono')) {
    return '"Consolas", "Cousine", "JetBrains Mono", monospace';
  } else if (fLower.includes('cambria')) {
    return '"Cambria", "Caladea", Georgia, serif';
  } else if (fLower.includes('times')) {
    return '"Times New Roman", "Tinos", Georgia, serif';
  } else if (fLower.includes('arial')) {
    return '"Arial", "Arimo", Helvetica, sans-serif';
  } else if (fLower.includes('aptos')) {
    return '"Aptos", "Calibri", "Carlito", "Segoe UI", sans-serif';
  }
  return `"${family}", "Calibri", "Carlito", Inter, sans-serif`;
}

function buildCanvasFont(weight, style, size, family) {
  const s = (style && (style === 'italic' || style === 'oblique')) ? 'italic ' : '';
  const w = (weight && weight !== 'normal' && weight !== '400') ? `${weight} ` : '';
  const fam = formatFontFamily(family);
  return `${s}${w}${size}px ${fam}`;
}

// Image element cache for smooth 60fps canvas rendering
function formatCssColor(color, fallback = '#1E293B') {
  if (!color || color === 'auto' || color === '') return fallback;
  if (color.startsWith('#') || color.startsWith('rgb') || color.startsWith('hsl')) return color;
  return `#${color}`;
}

const imageElementCache = new Map();

// 6. Draw Content on High-Fidelity Canvas
function drawCanvasPageItems(ctx, items) {
  items.forEach(item => {
    if (item.type === 'watermark') {
      ctx.save();
      ctx.translate(item.x, item.y);
      ctx.rotate((item.rotation_deg * Math.PI) / 180);
      ctx.globalAlpha = item.opacity;
      ctx.font = `800 ${item.font_size}px Outfit, sans-serif`;
      ctx.fillStyle = formatCssColor(item.color, '#64748B');
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(item.text, 0, 0);
      ctx.restore();

    } else if (item.type === 'image') {
      if (item.data_url) {
        let img = imageElementCache.get(item.data_url);
        if (!img) {
          img = new Image();
          img.src = item.data_url;
          img.onload = () => {
            // Redraw canvas when image finishes loading
            renderCanvasPagesFromWasm();
          };
          imageElementCache.set(item.data_url, img);
        }
        if (img.complete && img.naturalWidth > 0) {
          ctx.save();
          if (item.opacity !== undefined && item.opacity < 1.0) {
            ctx.globalAlpha = item.opacity;
          }
          ctx.drawImage(img, item.x, item.y, item.width, item.height);
          ctx.restore();
        }
      }

    } else if (item.type === 'line') {
      ctx.beginPath();
      ctx.moveTo(item.x1, item.y1);
      ctx.lineTo(item.x2, item.y2);
      ctx.strokeStyle = formatCssColor(item.color, '#CBD5E1');
      ctx.lineWidth = item.line_width;
      ctx.stroke();

    } else if (item.type === 'table_cell') {
      // 1. Draw cell background fill if specified
      if (item.bg_color) {
        ctx.fillStyle = formatCssColor(item.bg_color, '#F8FAFC');
        ctx.fillRect(item.x, item.y, item.width, item.height);
      } else if (item.is_header) {
        ctx.fillStyle = '#F8FAFC';
        ctx.fillRect(item.x, item.y, item.width, item.height);
      }

      // 2. Draw cell borders
      const borderColor = formatCssColor(item.border_color, '#DDD5C2');
      ctx.strokeStyle = borderColor;
      ctx.lineWidth = 1;
      ctx.strokeRect(item.x, item.y, item.width, item.height);

      // 3. Draw cell text content with exact color and alignment
      const textColor = formatCssColor(item.color, '#1E293B');
      const weight = item.font_weight || '400';
      const fontSize = item.font_size || 9.5;
      const fontFamily = item.font_family || 'Calibri, sans-serif';

      ctx.font = buildCanvasFont(weight, 'normal', fontSize, fontFamily);
      ctx.fillStyle = textColor;
      ctx.textAlign = item.align === 'center' ? 'center' : (item.align === 'right' ? 'right' : 'left');
      ctx.textBaseline = 'alphabetic';

      const lines = (item.lines && item.lines.length > 0) ? item.lines : (item.text ? item.text.split('\n') : ['']);
      const lineHeight = fontSize * 1.35;
      const totalTextH = lines.length * lineHeight;
      const startY = item.y + (item.height - totalTextH) / 2 + fontSize * 0.88;

      const padX = 8;
      const textX = item.align === 'center' 
        ? item.x + item.width / 2 
        : (item.align === 'right' ? item.x + item.width - padX : item.x + padX);

      lines.forEach((lineText, lIdx) => {
        ctx.fillText(lineText, textX, startY + lIdx * lineHeight);
      });

    } else if (item.type === 'text') {
      const isJustified = (item.align === 'both' || item.align === 'justify') && !item.is_last_line;

      if (item.runs && item.runs.length > 0) {
        let curX = item.x;
        let extraSpacePerSpace = 0;

        if (isJustified && item.max_width && item.max_width > 0) {
          let naturalWidth = 0;
          let spaceCount = 0;

          item.runs.forEach(run => {
            if (run.text === '\t') {
              naturalWidth += run.width || 0;
              return;
            }
            const weight = run.bold ? '700' : '400';
            const style = run.italic ? 'italic' : 'normal';
            const fontSize = run.font_size || item.font_size || 14.66;
            const fontFamily = run.font_family || item.font_family || 'Calibri, sans-serif';
            ctx.font = buildCanvasFont(weight, style, fontSize, fontFamily);
            naturalWidth += ctx.measureText(run.text).width;
            for (let i = 0; i < run.text.length; i++) {
              if (run.text[i] === ' ') spaceCount++;
            }
          });

          if (spaceCount > 0 && item.max_width > naturalWidth) {
            const gap = item.max_width - naturalWidth;
            if (gap / spaceCount < (item.font_size || 14) * 2.0) {
              extraSpacePerSpace = gap / spaceCount;
            }
          }
        }

        item.runs.forEach(run => {
          // Tabs were resolved to stop positions by the layout engine
          if (run.text === '\t') {
            curX += run.width || 0;
            return;
          }
          const weight = run.bold ? '700' : '400';
          const style = run.italic ? 'italic' : 'normal';
          const fontSize = run.font_size || item.font_size || 14.66;
          const fontFamily = run.font_family || item.font_family || 'Calibri, sans-serif';
          const runColor = formatCssColor(run.color || item.color, '#1E293B');

          ctx.font = buildCanvasFont(weight, style, fontSize, fontFamily);
          ctx.fillStyle = runColor;
          ctx.textAlign = 'left';
          ctx.textBaseline = 'alphabetic';

          const runStartX = curX;

          if (run.text.includes('\t')) {
            const parts = run.text.split('\t');
            parts.forEach((pStr, pIdx) => {
              if (pIdx > 0) {
                const tabInterval = 48;
                const relX = curX - (item.x < 100 ? item.x : 65);
                curX = 65 + Math.floor((relX + tabInterval) / tabInterval) * tabInterval;
              }
              if (pStr) {
                if (extraSpacePerSpace > 0 && pStr.includes(' ')) {
                  const words = pStr.split(' ');
                  words.forEach((w, wIdx) => {
                    if (wIdx > 0) {
                      curX += ctx.measureText(' ').width + extraSpacePerSpace;
                    }
                    if (w) {
                      ctx.fillText(w, curX, item.y);
                      curX += ctx.measureText(w).width;
                    }
                  });
                } else {
                  ctx.fillText(pStr, curX, item.y);
                  curX += ctx.measureText(pStr).width;
                }
              }
            });
          } else if (extraSpacePerSpace > 0 && run.text.includes(' ')) {
            const words = run.text.split(' ');
            words.forEach((w, wIdx) => {
              if (wIdx > 0) {
                curX += ctx.measureText(' ').width + extraSpacePerSpace;
              }
              if (w) {
                ctx.fillText(w, curX, item.y);
                curX += ctx.measureText(w).width;
              }
            });
          } else {
            ctx.fillText(run.text, curX, item.y);
            curX += ctx.measureText(run.text).width;
          }

          if (run.underline) {
            ctx.beginPath();
            ctx.moveTo(runStartX, item.y + 2);
            ctx.lineTo(curX, item.y + 2);
            ctx.strokeStyle = runColor;
            ctx.lineWidth = 1;
            ctx.stroke();
          }
        });
      } else {
        const fontSize = item.font_size || 14.66;
        const textColor = formatCssColor(item.color, '#1E293B');
        ctx.font = buildCanvasFont(item.font_weight, item.font_style, fontSize, item.font_family);
        ctx.fillStyle = textColor;
        ctx.textAlign = 'left';
        ctx.textBaseline = 'alphabetic';

        let extraSpacePerSpace = 0;
        if (isJustified && item.max_width && item.max_width > 0) {
          const naturalWidth = ctx.measureText(item.text).width;
          let spaceCount = 0;
          for (let i = 0; i < item.text.length; i++) {
            if (item.text[i] === ' ') spaceCount++;
          }
          if (spaceCount > 0 && item.max_width > naturalWidth) {
            const gap = item.max_width - naturalWidth;
            if (gap / spaceCount < fontSize * 2.0) {
              extraSpacePerSpace = gap / spaceCount;
            }
          }
        }

        if (extraSpacePerSpace > 0 && item.text.includes(' ')) {
          let curX = item.x;
          const words = item.text.split(' ');
          words.forEach((w, wIdx) => {
            if (wIdx > 0) {
              curX += ctx.measureText(' ').width + extraSpacePerSpace;
            }
            if (w) {
              ctx.fillText(w, curX, item.y);
              curX += ctx.measureText(w).width;
            }
          });
        } else if (item.text.includes('\t')) {
          let curX = item.x;
          const parts = item.text.split('\t');
          parts.forEach((pStr, pIdx) => {
            if (pIdx > 0) {
              const tabInterval = 48;
              const relX = curX - 65;
              curX = 65 + Math.floor((relX + tabInterval) / tabInterval) * tabInterval;
            }
            if (pStr) {
              ctx.fillText(pStr, curX, item.y);
              curX += ctx.measureText(pStr).width;
            }
          });
        } else {
          ctx.fillText(item.text, item.x, item.y);
        }
      }
    }
  });
}

let blurTimeout = null;

// 7. Interactive Click: Activate In-Place Full Paragraph / Table Cell Editor
function handleCanvasClick(e, page, canvas, pageCard) {
  if (blurTimeout) {
    clearTimeout(blurTimeout);
    blurTimeout = null;
  }

  const rect = canvas.getBoundingClientRect();
  const clickX = (e.clientX - rect.left) / currentZoom;
  const clickY = (e.clientY - rect.top) / currentZoom;

  // Hit test table cells first
  for (const item of page.items) {
    if (item.type === 'table_cell') {
      if (
        clickX >= item.x &&
        clickX <= item.x + item.width &&
        clickY >= item.y &&
        clickY <= item.y + item.height
      ) {
        if (
          activeTarget &&
          activeTarget.type === 'cell' &&
          activeTarget.tableIndex === item.table_index &&
          activeTarget.row === item.row &&
          activeTarget.col === item.col
        ) {
          closeActiveEditors();
          return;
        }
        commitCurrentEditor();
        openInPlaceCellEditor(item, pageCard, canvas);
        return;
      }
    }
  }

  // Hit test paragraph text items
  for (const item of page.items) {
    if (item.type === 'text') {
      if (
        clickX >= item.x - 25 &&
        clickX <= item.x + item.width + 50 &&
        clickY >= item.y - item.font_size - 6 &&
        clickY <= item.y + 8
      ) {
        // If clicking on the same paragraph that is currently being edited, close it to reveal the rendered canvas
        if (
          activeTarget &&
          activeTarget.type === 'paragraph' &&
          activeTarget.paragraphIndex === item.paragraph_index
        ) {
          closeActiveEditors();
          return;
        }
        commitCurrentEditor();
        openInPlaceParagraphEditor(item.paragraph_index, page, pageCard, canvas);
        return;
      }
    }
  }

  // If clicked on blank area, commit and close open editors
  closeActiveEditors();
}

function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function rgbToHex(color) {
  if (!color) return '';
  if (color.startsWith('#')) return color.toUpperCase();
  const rgbMatch = color.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
  if (rgbMatch) {
    const r = parseInt(rgbMatch[1], 10).toString(16).padStart(2, '0');
    const g = parseInt(rgbMatch[2], 10).toString(16).padStart(2, '0');
    const b = parseInt(rgbMatch[3], 10).toString(16).padStart(2, '0');
    return `#${r}${g}${b}`.toUpperCase();
  }
  return color;
}

function buildParagraphEditorHtml(p) {
  if (!p) return '';
  if (p.runs && p.runs.length > 0) {
    return p.runs.map(run => {
      let text = escapeHtml(run.text || '');
      if (!text) return '';
      text = text.replace(/\n/g, '<br>');
      
      let style = '';
      if (run.color && run.color.toLowerCase() !== 'auto' && run.color !== '') {
        const hex = run.color.startsWith('#') ? run.color : `#${run.color}`;
        style += `color: ${hex};`;
      }
      
      let inner = text;
      if (run.bold) inner = `<b>${inner}</b>`;
      if (run.italic) inner = `<i>${inner}</i>`;
      if (run.underline) inner = `<u>${inner}</u>`;
      if (style) inner = `<span style="${style}">${inner}</span>`;
      return inner;
    }).join('');
  }
  return escapeHtml(p.text || '').replace(/\n/g, '<br>');
}

function extractRunsFromEditor(rootEl) {
  const rawRuns = [];

  function traverse(node, currentStyle) {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent;
      if (text && text.length > 0) {
        rawRuns.push({
          text: text,
          bold: Boolean(currentStyle.bold),
          italic: Boolean(currentStyle.italic),
          underline: Boolean(currentStyle.underline),
          color: currentStyle.color || '',
          font_size: currentStyle.fontSize || null,
          font_family: currentStyle.fontFamily || null
        });
      }
      return;
    }

    if (node.nodeType === Node.ELEMENT_NODE) {
      const tag = node.tagName.toLowerCase();
      if (tag === 'br') {
        rawRuns.push({
          text: '\n',
          bold: Boolean(currentStyle.bold),
          italic: Boolean(currentStyle.italic),
          underline: Boolean(currentStyle.underline),
          color: currentStyle.color || '',
          font_size: currentStyle.fontSize || null,
          font_family: currentStyle.fontFamily || null
        });
        return;
      }

      const nextStyle = { ...currentStyle };
      const style = node.style;

      if (tag === 'b' || tag === 'strong' || (style && (style.fontWeight === 'bold' || parseInt(style.fontWeight, 10) >= 600))) {
        nextStyle.bold = true;
      }
      if (tag === 'i' || tag === 'em' || (style && style.fontStyle === 'italic')) {
        nextStyle.italic = true;
      }
      if (tag === 'u' || (style && style.textDecoration && style.textDecoration.includes('underline'))) {
        nextStyle.underline = true;
      }
      if (tag === 'font' && node.color) {
        nextStyle.color = rgbToHex(node.color).replace('#', '');
      } else if (style && style.color) {
        const hex = rgbToHex(style.color);
        if (hex) nextStyle.color = hex.replace('#', '');
      }

      for (let child = node.firstChild; child; child = child.nextSibling) {
        traverse(child, nextStyle);
      }
    }
  }

  const baseStyle = {
    bold: false,
    italic: false,
    underline: false,
    color: '',
    fontSize: null,
    fontFamily: null
  };

  traverse(rootEl, baseStyle);

  // Merge consecutive runs with identical formatting
  const mergedRuns = [];
  for (const r of rawRuns) {
    if (mergedRuns.length > 0) {
      const last = mergedRuns[mergedRuns.length - 1];
      if (
        last.bold === r.bold &&
        last.italic === r.italic &&
        last.underline === r.underline &&
        (last.color || '') === (r.color || '') &&
        last.font_size === r.font_size &&
        last.font_family === r.font_family
      ) {
        last.text += r.text;
        continue;
      }
    }
    mergedRuns.push({ ...r });
  }

  if (mergedRuns.length === 0) {
    mergedRuns.push({
      text: '',
      bold: false,
      italic: false,
      underline: false,
      color: '',
      font_size: null,
      font_family: null
    });
  }

  return mergedRuns;
}

function getParagraphElement(pIndex) {
  if (currentSession) {
    try {
      const elementsJson = currentSession.get_document_elements_json();
      activeDocumentElements = JSON.parse(elementsJson);
    } catch (e) {
      console.warn('Could not refresh elements:', e);
    }
  }
  return activeDocumentElements?.find(
    el => el.type === 'paragraph' && el.index === pIndex
  );
}

// 8. Open In-Place Full Paragraph Block Editor
function openInPlaceParagraphEditor(paragraphIndex, page, pageCard, canvas) {
  if (blurTimeout) {
    clearTimeout(blurTimeout);
    blurTimeout = null;
  }
  clearTimeout(updateDebounceTimer);

  activeParagraphIndex = paragraphIndex;

  // Find all text lines belonging to this paragraph on this page
  const pLines = page.items.filter(
    it => it.type === 'text' && it.paragraph_index === paragraphIndex
  );

  if (pLines.length === 0) return;

  const firstLine = pLines[0];
  const minY = Math.min(...pLines.map(it => it.y - it.font_size * 0.85));
  const maxY = Math.max(...pLines.map(it => it.y + it.height - it.font_size * 0.85));
  const blockHeight = Math.max(34, maxY - minY);

  const pElement = getParagraphElement(paragraphIndex);
  const marginL = (page.margin_left !== undefined) ? page.margin_left : 65;
  const marginR = (page.margin_right !== undefined) ? page.margin_right : 65;
  const printableW = (page.printable_width !== undefined) ? page.printable_width : Math.max(100, page.width - marginL - marginR);

  const indentL = (pElement && pElement.indent_left) ? pElement.indent_left : 0;
  const indentR = (pElement && pElement.indent_right) ? pElement.indent_right : 0;

  const blockX = marginL + indentL;
  const blockWidth = Math.max(80, printableW - indentL - indentR);

  const rawAlign = pElement?.align || firstLine.align || 'left';
  const align = (rawAlign === 'both' || rawAlign === 'justify') ? 'justify' : (rawAlign === 'center' || rawAlign === 'right' ? rawAlign : 'left');

  const fullHtml = buildParagraphEditorHtml(pElement || { text: getParagraphFullText(paragraphIndex) });

  const pageCardRect = pageCard.getBoundingClientRect();
  const containerRect = canvasDocumentView.getBoundingClientRect();

  const cardOffsetLeft = pageCardRect.left - containerRect.left + canvasDocumentView.scrollLeft;
  const cardOffsetTop = pageCardRect.top - containerRect.top + canvasDocumentView.scrollTop;

  const canvasLeft = canvas.offsetLeft;
  const canvasTop = canvas.offsetTop;

  const pageBg = page.bg_color || currentPageBgColor || '#FFFFFF';
  const cleanBg = pageBg.startsWith('#') ? pageBg : `#${pageBg}`;
  const textColor = firstLine.color ? (firstLine.color.startsWith('#') ? firstLine.color : `#${firstLine.color}`) : '#1E293B';
  const lineHeightPx = firstLine.height || (firstLine.font_size * 1.35);

  canvasParagraphEditor.style.left = `${cardOffsetLeft + canvasLeft + blockX * currentZoom}px`;
  canvasParagraphEditor.style.top = `${cardOffsetTop + canvasTop + minY * currentZoom}px`;
  canvasParagraphEditor.style.width = `${blockWidth * currentZoom}px`;
  canvasParagraphEditor.style.minHeight = `${blockHeight * currentZoom}px`;
  canvasParagraphEditor.style.fontSize = `${firstLine.font_size * currentZoom}px`;
  canvasParagraphEditor.style.fontFamily = formatFontFamily(firstLine.font_family);
  canvasParagraphEditor.style.fontWeight = (firstLine.font_weight === '700' || pElement?.bold) ? 'bold' : 'normal';
  canvasParagraphEditor.style.fontStyle = (firstLine.font_style === 'italic' || pElement?.italic) ? 'italic' : 'normal';
  canvasParagraphEditor.style.color = textColor;
  canvasParagraphEditor.style.caretColor = textColor;
  canvasParagraphEditor.style.background = cleanBg;
  canvasParagraphEditor.style.textAlign = align;
  canvasParagraphEditor.style.lineHeight = `${lineHeightPx * currentZoom}px`;
  canvasParagraphEditor.style.letterSpacing = 'normal';
  canvasParagraphEditor.style.border = 'none';
  canvasParagraphEditor.style.outline = 'none';
  canvasParagraphEditor.style.boxShadow = 'none';
  canvasParagraphEditor.style.padding = '0';
  canvasParagraphEditor.style.margin = '0';
  canvasParagraphEditor.style.boxSizing = 'border-box';
  canvasParagraphEditor.style.wordBreak = 'normal';
  canvasParagraphEditor.style.overflowWrap = 'break-word';
  canvasParagraphEditor.style.whiteSpace = 'pre-wrap';

  canvasParagraphEditor.innerHTML = fullHtml;
  canvasParagraphEditor.style.display = 'block';

  activeTarget = {
    type: 'paragraph',
    paragraphIndex,
    item: firstLine
  };

  // Focus and select at end
  canvasParagraphEditor.focus();
  placeCaretAtEnd(canvasParagraphEditor);

  // Update Ribbon buttons to reflect paragraph formatting
  updateRibbonStateForParagraph(firstLine, pElement);
}

// 9. Open In-Place Table Cell Editor
function openInPlaceCellEditor(cellItem, pageCard, canvas) {
  if (blurTimeout) {
    clearTimeout(blurTimeout);
    blurTimeout = null;
  }
  clearTimeout(updateDebounceTimer);

  const pageCardRect = pageCard.getBoundingClientRect();
  const containerRect = canvasDocumentView.getBoundingClientRect();

  const cardOffsetLeft = pageCardRect.left - containerRect.left + canvasDocumentView.scrollLeft;
  const cardOffsetTop = pageCardRect.top - containerRect.top + canvasDocumentView.scrollTop;

  const canvasLeft = canvas.offsetLeft;
  const canvasTop = canvas.offsetTop;

  const textColor = cellItem.color ? (cellItem.color.startsWith('#') ? cellItem.color : `#${cellItem.color}`) : '#1E293B';
  const bgColor = cellItem.bg_color ? (cellItem.bg_color.startsWith('#') ? cellItem.bg_color : `#${cellItem.bg_color}`) : '#FFFFFF';
  const borderColor = cellItem.border_color ? (cellItem.border_color.startsWith('#') ? cellItem.border_color : `#${cellItem.border_color}`) : '#DDD5C2';

  canvasCellEditor.style.left = `${cardOffsetLeft + canvasLeft + cellItem.x * currentZoom}px`;
  canvasCellEditor.style.top = `${cardOffsetTop + canvasTop + cellItem.y * currentZoom}px`;
  canvasCellEditor.style.width = `${cellItem.width * currentZoom}px`;
  canvasCellEditor.style.minHeight = `${cellItem.height * currentZoom}px`;
  canvasCellEditor.style.fontSize = `${cellItem.font_size * currentZoom}px`;
  canvasCellEditor.style.fontFamily = formatFontFamily(cellItem.font_family);
  canvasCellEditor.style.fontWeight = cellItem.font_weight || '400';
  canvasCellEditor.style.color = textColor;
  canvasCellEditor.style.caretColor = textColor;
  canvasCellEditor.style.background = bgColor;
  canvasCellEditor.style.border = 'none';
  canvasCellEditor.style.outline = 'none';
  canvasCellEditor.style.boxShadow = 'none';
  canvasCellEditor.style.textAlign = cellItem.align || 'left';
  canvasCellEditor.style.lineHeight = '1.35';

  canvasCellEditor.innerText = cellItem.text;
  canvasCellEditor.style.display = 'flex';

  activeTarget = {
    type: 'cell',
    tableIndex: cellItem.table_index,
    row: cellItem.row,
    col: cellItem.col,
    item: cellItem
  };

  canvasCellEditor.focus();
  placeCaretAtEnd(canvasCellEditor);
}

// 10. Live Synchronize Paragraph Editor Input
function handleParagraphEditorInput() {
  if (!canvasParagraphEditor || !activeTarget || activeTarget.type !== 'paragraph' || !currentSession) return;

  const runs = extractRunsFromEditor(canvasParagraphEditor);
  const pIndex = activeTarget.paragraphIndex;
  setSyncStatus(false);

  // Update local memory cache
  const p = activeDocumentElements.find(
    el => el.type === 'paragraph' && el.index === pIndex
  );
  if (p) {
    p.runs = runs;
    p.text = runs.map(r => r.text).join('');
  }

  clearTimeout(updateDebounceTimer);
  updateDebounceTimer = setTimeout(() => {
    try {
      const align = p?.align || 'left';
      currentSession.update_paragraph_runs(pIndex, JSON.stringify(runs), align);
      setSyncStatus(true);
      updateLiveStats();
    } catch (err) {
      console.error('Error syncing paragraph runs:', err);
    }
  }, 200);
}

function handleCellEditorInput() {
  if (!canvasCellEditor || !activeTarget || activeTarget.type !== 'cell' || !currentSession) return;

  const newText = canvasCellEditor.innerText;
  const { tableIndex, row, col } = activeTarget;
  setSyncStatus(false);

  clearTimeout(updateDebounceTimer);
  updateDebounceTimer = setTimeout(() => {
    try {
      currentSession.update_table_cell(tableIndex, row, col, newText);
      setSyncStatus(true);
      updateLiveStats();
    } catch (err) {
      console.error('Error syncing cell:', err);
    }
  }, 250);
}

// 11. Commit and Close Active Editors
function commitCurrentEditor() {
  clearTimeout(updateDebounceTimer);

  if (canvasParagraphEditor && canvasParagraphEditor.style.display !== 'none') {
    if (activeTarget && activeTarget.type === 'paragraph' && currentSession) {
      const runs = extractRunsFromEditor(canvasParagraphEditor);
      const pIndex = activeTarget.paragraphIndex;
      try {
        const p = activeDocumentElements.find(
          el => el.type === 'paragraph' && el.index === pIndex
        );
        const align = p?.align || 'left';
        currentSession.update_paragraph_runs(pIndex, JSON.stringify(runs), align);
        if (p) {
          p.runs = runs;
          p.text = runs.map(r => r.text).join('');
        }
      } catch (err) {
        console.error('Error committing paragraph runs:', err);
      }
    }
    canvasParagraphEditor.style.display = 'none';
  }

  if (canvasCellEditor && canvasCellEditor.style.display !== 'none') {
    if (activeTarget && activeTarget.type === 'cell' && currentSession) {
      const finalVal = canvasCellEditor.innerText;
      try {
        currentSession.update_table_cell(activeTarget.tableIndex, activeTarget.row, activeTarget.col, finalVal);
      } catch (err) {
        console.error('Error committing cell:', err);
      }
    }
    canvasCellEditor.style.display = 'none';
  }

  setSyncStatus(true);
  activeTarget = null;
}

function closeActiveEditors() {
  commitCurrentEditor();
  renderCanvasPagesFromWasm();
}

function navigateToAdjacentCell(direction) {
  if (!activeTarget || activeTarget.type !== 'cell' || !canvasPagesLayout) return;
  const { tableIndex, row, col } = activeTarget;
  commitCurrentEditor();

  const allCells = [];
  const cellPageMap = new Map();
  canvasPagesLayout.pages.forEach(p => {
    p.items.forEach(it => {
      if (it.type === 'table_cell' && it.table_index === tableIndex) {
        allCells.push(it);
        cellPageMap.set(it, p.page_number);
      }
    });
  });

  const curIdx = allCells.findIndex(c => c.row === row && c.col === col);
  if (curIdx !== -1) {
    const nextIdx = curIdx + direction;
    if (nextIdx >= 0 && nextIdx < allCells.length) {
      const nextCell = allCells[nextIdx];
      const pageNum = cellPageMap.get(nextCell);
      const pageCard = document.querySelector(`.canvas-page-card[data-page-number="${pageNum}"]`);
      const canvas = pageCard?.querySelector('.page-canvas');
      if (pageCard && canvas) {
        openInPlaceCellEditor(nextCell, pageCard, canvas);
        return;
      }
    }
  }
  renderCanvasPagesFromWasm();
}

function navigateToNextRowCell() {
  if (!activeTarget || activeTarget.type !== 'cell' || !canvasPagesLayout) return;
  const { tableIndex, row, col } = activeTarget;
  commitCurrentEditor();

  let targetCell = null;
  let targetPageNum = null;
  canvasPagesLayout.pages.forEach(p => {
    p.items.forEach(it => {
      if (it.type === 'table_cell' && it.table_index === tableIndex && it.row === row + 1 && it.col === col) {
        targetCell = it;
        targetPageNum = p.page_number;
      }
    });
  });

  if (targetCell) {
    const pageCard = document.querySelector(`.canvas-page-card[data-page-number="${targetPageNum}"]`);
    const canvas = pageCard?.querySelector('.page-canvas');
    if (pageCard && canvas) {
      openInPlaceCellEditor(targetCell, pageCard, canvas);
      return;
    }
  }
  renderCanvasPagesFromWasm();
}

function getParagraphFullText(pIndex) {
  if (currentSession) {
    try {
      const elementsJson = currentSession.get_document_elements_json();
      activeDocumentElements = JSON.parse(elementsJson);
    } catch (e) {
      console.warn('Could not refresh elements:', e);
    }
  }
  const p = activeDocumentElements?.find(
    el => el.type === 'paragraph' && el.index === pIndex
  );
  return p ? p.text : '';
}

function placeCaretAtEnd(el) {
  if (typeof window.getSelection !== 'undefined' && typeof document.createRange !== 'undefined') {
    const range = document.createRange();
    range.selectNodeContents(el);
    range.collapse(false);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
  }
}

function setSyncStatus(isSynced) {
  if (isSynced) {
    syncStatus.className = 'status-badge';
    syncStatus.innerHTML = '<span class="status-dot"></span> Sincronizado';
  } else {
    syncStatus.className = 'status-badge saving';
    syncStatus.innerHTML = '<span class="status-dot"></span> Guardando...';
  }
}

// 12. Formatting Ribbon Actions (Applies to Selection or Active Paragraph Block)
function applyAlignment(alignValue) {
  if (activeParagraphIndex === null) {
    showToast('Haz clic en un párrafo de la página para alinearlo.');
    return;
  }

  const cssAlign = (alignValue === 'both' || alignValue === 'justify') ? 'justify' : (alignValue === 'center' || alignValue === 'right' ? alignValue : 'left');

  if (canvasParagraphEditor && canvasParagraphEditor.style.display !== 'none') {
    canvasParagraphEditor.style.textAlign = cssAlign;
  }

  const p = activeDocumentElements.find(el => el.type === 'paragraph' && el.index === activeParagraphIndex);
  if (p) p.align = alignValue;

  if (canvasParagraphEditor && canvasParagraphEditor.style.display !== 'none') {
    const runs = extractRunsFromEditor(canvasParagraphEditor);
    currentSession.update_paragraph_runs(activeParagraphIndex, JSON.stringify(runs), alignValue);
  } else {
    const text = p?.text || '';
    const color = p?.color || '';
    const bold = p?.bold || false;
    const italic = p?.italic || false;
    currentSession.update_paragraph_rich(activeParagraphIndex, text, alignValue, color, bold, italic);
  }

  updateRibbonAlignUI(alignValue);

  if (!canvasParagraphEditor || canvasParagraphEditor.style.display === 'none') {
    renderCanvasPagesFromWasm();
  }
}

function toggleBold() {
  if (canvasParagraphEditor && canvasParagraphEditor.style.display !== 'none') {
    canvasParagraphEditor.focus();
    document.execCommand('bold', false, null);
    handleParagraphEditorInput();
    updateSelectionFormattingState();
    return;
  }

  if (activeParagraphIndex === null) {
    showToast('Haz clic en un párrafo o texto para aplicar negrita.');
    return;
  }

  const p = activeDocumentElements.find(el => el.type === 'paragraph' && el.index === activeParagraphIndex);
  if (!p) return;
  const isBold = !p.bold;
  p.bold = isBold;
  if (p.runs && p.runs.length > 0) {
    p.runs.forEach(r => r.bold = isBold);
    currentSession.update_paragraph_runs(activeParagraphIndex, JSON.stringify(p.runs), p.align || 'left');
  } else {
    currentSession.update_paragraph_rich(activeParagraphIndex, p.text || '', p.align || 'left', p.color || '', isBold, p.italic || false);
  }
  btnFmtBold.classList.toggle('active', isBold);
  renderCanvasPagesFromWasm();
}

function toggleItalic() {
  if (canvasParagraphEditor && canvasParagraphEditor.style.display !== 'none') {
    canvasParagraphEditor.focus();
    document.execCommand('italic', false, null);
    handleParagraphEditorInput();
    updateSelectionFormattingState();
    return;
  }

  if (activeParagraphIndex === null) {
    showToast('Haz clic en un párrafo o texto para aplicar cursiva.');
    return;
  }

  const p = activeDocumentElements.find(el => el.type === 'paragraph' && el.index === activeParagraphIndex);
  if (!p) return;
  const isItalic = !p.italic;
  p.italic = isItalic;
  if (p.runs && p.runs.length > 0) {
    p.runs.forEach(r => r.italic = isItalic);
    currentSession.update_paragraph_runs(activeParagraphIndex, JSON.stringify(p.runs), p.align || 'left');
  } else {
    currentSession.update_paragraph_rich(activeParagraphIndex, p.text || '', p.align || 'left', p.color || '', p.bold || false, isItalic);
  }
  btnFmtItalic.classList.toggle('active', isItalic);
  renderCanvasPagesFromWasm();
}

function toggleUnderline() {
  if (canvasParagraphEditor && canvasParagraphEditor.style.display !== 'none') {
    canvasParagraphEditor.focus();
    document.execCommand('underline', false, null);
    handleParagraphEditorInput();
    updateSelectionFormattingState();
    return;
  }

  if (activeParagraphIndex === null) {
    showToast('Haz clic en un párrafo o texto para aplicar subrayado.');
    return;
  }

  const p = activeDocumentElements.find(el => el.type === 'paragraph' && el.index === activeParagraphIndex);
  if (p && p.runs && p.runs.length > 0) {
    const isUnderline = !p.runs[0].underline;
    p.runs.forEach(r => r.underline = isUnderline);
    currentSession.update_paragraph_runs(activeParagraphIndex, JSON.stringify(p.runs), p.align || 'left');
    if (btnFmtUnderline) btnFmtUnderline.classList.toggle('active', isUnderline);
    renderCanvasPagesFromWasm();
  }
}

function applyTextColor(hexColor) {
  if (!hexColor) return;
  const cleanHex = hexColor.startsWith('#') ? hexColor : `#${hexColor}`;

  if (canvasParagraphEditor && canvasParagraphEditor.style.display !== 'none') {
    canvasParagraphEditor.focus();
    document.execCommand('foreColor', false, cleanHex);
    handleParagraphEditorInput();
    return;
  }

  if (activeParagraphIndex === null) {
    showToast('Haz clic en un párrafo o texto para cambiar su color.');
    return;
  }

  const p = activeDocumentElements.find(el => el.type === 'paragraph' && el.index === activeParagraphIndex);
  if (!p) return;
  const rawCol = cleanHex.replace('#', '');
  p.color = rawCol;
  if (p.runs && p.runs.length > 0) {
    p.runs.forEach(r => r.color = cleanHex);
    currentSession.update_paragraph_runs(activeParagraphIndex, JSON.stringify(p.runs), p.align || 'left');
  } else {
    currentSession.update_paragraph_rich(activeParagraphIndex, p.text || '', p.align || 'left', cleanHex, p.bold || false, p.italic || false);
  }
  renderCanvasPagesFromWasm();
}

function updateSelectionFormattingState() {
  if (canvasParagraphEditor && canvasParagraphEditor.style.display !== 'none') {
    try {
      const isBold = document.queryCommandState('bold');
      const isItalic = document.queryCommandState('italic');
      const isUnderline = document.queryCommandState('underline');
      btnFmtBold.classList.toggle('active', Boolean(isBold));
      btnFmtItalic.classList.toggle('active', Boolean(isItalic));
      if (btnFmtUnderline) btnFmtUnderline.classList.toggle('active', Boolean(isUnderline));

      const foreColor = document.queryCommandValue('foreColor');
      if (foreColor) {
        const hex = rgbToHex(foreColor);
        if (hex && hex.startsWith('#')) {
          textColorInput.value = hex;
        }
      }
    } catch (e) {
      // ignore
    }
  }
}

function updateRibbonAlignUI(align) {
  Object.keys(alignButtons).forEach(k => {
    if (alignButtons[k]) {
      alignButtons[k].classList.toggle('active', k === align || (k === 'both' && (align === 'justify' || align === 'both')));
    }
  });
}

function updateRibbonStateForParagraph(item, pElement) {
  updateRibbonAlignUI(item.align || pElement?.align || 'left');

  const isBold = item.font_weight === '700' || pElement?.bold;
  const isItalic = item.font_style === 'italic' || pElement?.italic;
  const isUnderline = pElement?.runs?.[0]?.underline || false;

  btnFmtBold.classList.toggle('active', Boolean(isBold));
  btnFmtItalic.classList.toggle('active', Boolean(isItalic));
  if (btnFmtUnderline) btnFmtUnderline.classList.toggle('active', Boolean(isUnderline));

  const color = item.color || pElement?.color;
  if (color) {
    textColorInput.value = color.startsWith('#') ? color : `#${color}`;
  }
}

// 13. Statistics Calculations
function updateLiveStats() {
  if (!currentSession) return;
  try {
    const statsJson = currentSession.get_stats_json();
    const stats = JSON.parse(statsJson);

    statWords.textContent = stats.word_count.toLocaleString();
    statChars.textContent = stats.char_count.toLocaleString();
    statParagraphs.textContent = stats.paragraph_count.toLocaleString();
    statTables.textContent = (stats.table_count || 0).toLocaleString();

    if (stats.page_setup && statFormat) {
      const isLandscape = stats.page_setup.orientation === 'landscape';
      statFormat.textContent = isLandscape ? 'Horizontal (Landscape)' : 'Vertical';
    }
  } catch (err) {
    console.warn('Could not read stats:', err);
  }
}

// 14. Search and Replace
function executeSearchAndReplace() {
  if (!currentSession) return;

  const search = findQuery.value;
  const replace = replaceQuery.value;
  const matchCase = chkCaseSensitive.checked;
  const useRegex = chkRegex.checked;

  if (!search) {
    showToast('Introduce un término a buscar.', true);
    return;
  }

  try {
    commitCurrentEditor();
    const resultJson = currentSession.find_and_replace(search, replace, matchCase, useRegex);
    const result = JSON.parse(resultJson);

    if (result.occurrences_replaced > 0) {
      showToast(result.message);
      refreshDocumentView();
    } else {
      showToast(`No se encontraron coincidencias para "${search}".`, true);
    }
  } catch (err) {
    console.error('Find/replace error:', err);
    showToast('Error en reemplazo: ' + err, true);
  }
}

// 15. Template Variables
function scanAndHighlightVariables() {
  if (!currentSession) return;
  try {
    const rawText = currentSession.get_raw_text();
    // Supports both {{VARIABLE}} and VEAD-style {variable} placeholders
    const regex = /\{\{[^{}]+\}\}|\{[^{}\s]+\}/g;
    const found = new Set();
    let match;

    while ((match = regex.exec(rawText)) !== null) {
      found.add(match[0]);
    }

    detectedVariables = Array.from(found);

    if (detectedVariables.length > 0) {
      varsCountBadge.textContent = detectedVariables.length;
      varsCountBadge.style.display = 'inline-block';
    } else {
      varsCountBadge.style.display = 'none';
    }
  } catch (err) {
    console.warn('Could not scan variables:', err);
  }
}

function openVariablesModal() {
  variablesListContainer.innerHTML = '';
  if (detectedVariables.length === 0) {
    addVariableInputRow('{{EJEMPLO}}', 'Valor');
  } else {
    detectedVariables.forEach(varTag => {
      addVariableInputRow(varTag, '');
    });
  }
  variablesModal.style.display = 'flex';
}

function addVariableInputRow(key = '', val = '') {
  const row = document.createElement('div');
  row.className = 'var-item-row';
  row.innerHTML = `
    <input type="text" class="var-key-input" value="${escapeHtml(key)}" placeholder="{{CAMPO}}" />
    <input type="text" class="var-val-input" value="${escapeHtml(val)}" placeholder="Nuevo valor..." />
    <button class="btn-icon btn-del-var" title="Eliminar">✕</button>
  `;
  row.querySelector('.btn-del-var').addEventListener('click', () => row.remove());
  variablesListContainer.appendChild(row);
}

function applyVariablesFromModal() {
  if (!currentSession) return;
  const rows = variablesListContainer.querySelectorAll('.var-item-row');
  const pairs = [];

  rows.forEach(r => {
    const key = r.querySelector('.var-key-input').value.trim();
    const value = r.querySelector('.var-val-input').value;
    if (key) pairs.push({ key, value });
  });

  if (pairs.length === 0) {
    showToast('No hay variables válidas para aplicar.', true);
    return;
  }

  try {
    commitCurrentEditor();
    const resultJson = currentSession.batch_replace(JSON.stringify(pairs));
    const result = JSON.parse(resultJson);
    showToast(result.message);
    variablesModal.style.display = 'none';
    refreshDocumentView();
  } catch (err) {
    console.error('Batch replace error:', err);
    showToast('Error al aplicar variables: ' + err, true);
  }
}

// 16. Insert Table
function handleCreateTableConfirm() {
  if (!currentSession) return;
  const rows = parseInt(tblInputRows.value, 10) || 3;
  const cols = parseInt(tblInputCols.value, 10) || 3;

  const headers = [];
  for (let i = 1; i <= cols; i++) {
    headers.push(`Columna ${i}`);
  }

  try {
    commitCurrentEditor();
    currentSession.add_table(rows, cols, JSON.stringify(headers));
    tableModal.style.display = 'none';
    showToast(`Tabla de ${rows}x${cols} insertada exitosamente.`);
    refreshDocumentView();
  } catch (err) {
    console.error('Error inserting table:', err);
    showToast('Error al insertar tabla: ' + err, true);
  }
}

// 17. Download Modified DOCX
function downloadDocx() {
  if (!currentSession) return;

  try {
    commitCurrentEditor();
    const bytes = currentSession.export_bytes();
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    
    const dotIdx = currentFileName.lastIndexOf('.');
    const baseName = dotIdx !== -1 ? currentFileName.substring(0, dotIdx) : currentFileName;
    a.download = `${baseName}_modificado.docx`;

    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast('¡Archivo DOCX descargado exitosamente!');
  } catch (err) {
    console.error('Export error:', err);
    showToast('Error al exportar DOCX: ' + err, true);
  }
}

// 18. Toast & Utilities
function showToast(msg, isError = false) {
  toastText.textContent = msg;
  toast.style.background = isError 
    ? 'rgba(239, 68, 68, 0.15)' 
    : 'rgba(16, 185, 129, 0.15)';
  toast.style.borderColor = isError 
    ? 'rgba(239, 68, 68, 0.4)' 
    : 'rgba(16, 185, 129, 0.4)';
  toast.style.color = isError ? '#FCA5A5' : '#A7F3D0';
  toast.style.display = 'flex';

  setTimeout(() => {
    toast.style.display = 'none';
  }, 4000);
}

// 19. Setup Event Listeners
function setupEventListeners() {
  // Selection change to reflect active formatting (bold, italic, underline, color) in the ribbon
  document.addEventListener('selectionchange', updateSelectionFormattingState);

  // Zoom Controls
  btnZoomIn.addEventListener('click', () => {
    if (currentZoom < 1.5) {
      currentZoom += 0.25;
      zoomLevelLabel.textContent = `${Math.round(currentZoom * 100)}%`;
      closeActiveEditors();
    }
  });

  btnZoomOut.addEventListener('click', () => {
    if (currentZoom > 0.5) {
      currentZoom -= 0.25;
      zoomLevelLabel.textContent = `${Math.round(currentZoom * 100)}%`;
      closeActiveEditors();
    }
  });

  // Paragraph & Cell Editor inputs and keyboard shortcuts
  canvasParagraphEditor.addEventListener('input', handleParagraphEditorInput);
  canvasCellEditor.addEventListener('input', handleCellEditorInput);

  canvasParagraphEditor.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      closeActiveEditors();
    }
  });

  canvasCellEditor.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      closeActiveEditors();
    } else if (e.key === 'Tab') {
      e.preventDefault();
      navigateToAdjacentCell(e.shiftKey ? -1 : 1);
    } else if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      navigateToNextRowCell();
    }
  });

  // Blur handlers to commit changes when clicking away
  canvasParagraphEditor.addEventListener('blur', () => {
    if (blurTimeout) clearTimeout(blurTimeout);
    blurTimeout = setTimeout(() => {
      const activeEl = document.activeElement;
      if (
        activeEl !== canvasParagraphEditor &&
        activeEl !== canvasCellEditor &&
        !activeEl?.closest('.formatting-ribbon') &&
        !activeEl?.closest('.modal-card') &&
        !activeEl?.closest('.modal-backdrop')
      ) {
        closeActiveEditors();
      }
    }, 100);
  });

  canvasCellEditor.addEventListener('blur', () => {
    if (blurTimeout) clearTimeout(blurTimeout);
    blurTimeout = setTimeout(() => {
      const activeEl = document.activeElement;
      if (
        activeEl !== canvasParagraphEditor &&
        activeEl !== canvasCellEditor &&
        !activeEl?.closest('.formatting-ribbon') &&
        !activeEl?.closest('.modal-card') &&
        !activeEl?.closest('.modal-backdrop')
      ) {
        closeActiveEditors();
      }
    }, 100);
  });

  // Prevent ribbon buttons from stealing focus from the active paragraph editor
  document.querySelectorAll('.formatting-ribbon button, .formatting-ribbon input').forEach(btn => {
    btn.addEventListener('mousedown', (e) => {
      if (e.target.tagName !== 'INPUT') {
        e.preventDefault();
      }
    });
  });

  // File inputs
  [fileInput, fileInputWelcome].forEach(input => {
    if (!input) return;
    input.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        loadDocxFile(e.target.files[0]);
      }
    });
  });

  // Demo buttons
  [btnLoadSample, btnWelcomeDemo].forEach(btn => {
    if (!btn) return;
    btn.addEventListener('click', loadSampleDocx);
  });

  // Real Example Cards
  document.querySelectorAll('.example-card-btn').forEach(card => {
    card.addEventListener('click', () => {
      const docName = card.dataset.example;
      if (docName) {
        loadExampleDocx(docName);
      }
    });
  });

  // Download
  btnDownload.addEventListener('click', downloadDocx);

  // Drag and Drop
  if (dropTarget) {
    ['dragenter', 'dragover'].forEach(name => {
      dropTarget.addEventListener(name, (e) => {
        e.preventDefault();
        dropTarget.classList.add('drag-over');
      });
    });

    ['dragleave', 'drop'].forEach(name => {
      dropTarget.addEventListener(name, (e) => {
        e.preventDefault();
        dropTarget.classList.remove('drag-over');
      });
    });

    dropTarget.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      if (dt && dt.files && dt.files[0]) {
        loadDocxFile(dt.files[0]);
      }
    });
  }

  // Formatting Ribbon
  btnFmtBold.addEventListener('click', toggleBold);
  btnFmtItalic.addEventListener('click', toggleItalic);
  if (btnFmtUnderline) {
    btnFmtUnderline.addEventListener('click', toggleUnderline);
  }

  Object.keys(alignButtons).forEach(key => {
    const btn = alignButtons[key];
    if (btn) {
      btn.addEventListener('click', () => applyAlignment(key));
    }
  });

  textColorInput.addEventListener('input', (e) => {
    applyTextColor(e.target.value);
  });

  colorSwatchDots.forEach(dot => {
    dot.addEventListener('click', () => {
      const color = dot.dataset.color;
      textColorInput.value = color;
      applyTextColor(color);
    });
  });

  // Table Modal
  btnInsertTable.addEventListener('click', () => {
    tableModal.style.display = 'flex';
  });
  btnCloseTableModal.addEventListener('click', () => {
    tableModal.style.display = 'none';
  });
  btnCreateTableConfirm.addEventListener('click', handleCreateTableConfirm);

  // Background / Watermark Modal
  btnToggleBgModal.addEventListener('click', () => {
    backgroundModal.style.display = 'flex';
  });
  btnCloseBgModal.addEventListener('click', () => {
    backgroundModal.style.display = 'none';
  });
  btnSaveBgConfig.addEventListener('click', () => {
    backgroundModal.style.display = 'none';
    refreshDocumentView();
    showToast('Configuración de fondo aplicada.');
  });

  bgPresetBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      bgPresetBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentPageBgColor = btn.dataset.bg;
      if (currentSession) {
        currentSession.set_background_color(currentPageBgColor);
      }
      refreshDocumentView();
    });
  });

  btnWmConfidential.addEventListener('click', () => {
    currentWatermark = {
      type: 'preset',
      text: 'CONFIDENCIAL',
      opacity: parseInt(watermarkOpacity.value, 10) / 100
    };
    refreshDocumentView();
    showToast('Marca de agua "CONFIDENCIAL" aplicada.');
  });

  btnWmDraft.addEventListener('click', () => {
    currentWatermark = {
      type: 'preset',
      text: 'BORRADOR',
      opacity: parseInt(watermarkOpacity.value, 10) / 100
    };
    refreshDocumentView();
    showToast('Marca de agua "BORRADOR" aplicada.');
  });

  btnWmClear.addEventListener('click', () => {
    currentWatermark = { type: 'none', text: '', opacity: 0.2 };
    refreshDocumentView();
    showToast('Marca de agua eliminada.');
  });

  watermarkOpacity.addEventListener('input', (e) => {
    opacityValLabel.textContent = `${e.target.value}%`;
    currentWatermark.opacity = parseInt(e.target.value, 10) / 100;
    refreshDocumentView();
  });

  // Search Toggle
  btnToggleFind.addEventListener('click', () => {
    const isVisible = searchReplaceBar.style.display === 'flex';
    searchReplaceBar.style.display = isVisible ? 'none' : 'flex';
    if (!isVisible) {
      findQuery.focus();
      btnToggleFind.classList.add('active');
    } else {
      btnToggleFind.classList.remove('active');
    }
  });

  btnCloseSearch.addEventListener('click', () => {
    searchReplaceBar.style.display = 'none';
    btnToggleFind.classList.remove('active');
  });

  btnDoReplace.addEventListener('click', executeSearchAndReplace);
  [findQuery, replaceQuery].forEach(inp => {
    inp.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') executeSearchAndReplace();
    });
  });

  // Variables
  btnToggleVariables.addEventListener('click', openVariablesModal);
  btnCloseVariables.addEventListener('click', () => {
    variablesModal.style.display = 'none';
  });
  btnAddCustomVar.addEventListener('click', () => addVariableInputRow());
  btnApplyAllVariables.addEventListener('click', applyVariablesFromModal);

  // Close modals on backdrop click
  [variablesModal, backgroundModal, tableModal].forEach(modal => {
    modal.addEventListener('click', (e) => {
      if (e.target === modal) modal.style.display = 'none';
    });
  });

  // Keyboard Shortcuts
  window.addEventListener('keydown', (e) => {
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'f') {
      e.preventDefault();
      searchReplaceBar.style.display = 'flex';
      btnToggleFind.classList.add('active');
      findQuery.focus();
      findQuery.select();
    }
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'b') {
      e.preventDefault();
      toggleBold();
    }
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'i') {
      e.preventDefault();
      toggleItalic();
    }
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'u') {
      e.preventDefault();
      toggleUnderline();
    }
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      if (currentSession) downloadDocx();
    }
    if (e.key === 'Escape') {
      searchReplaceBar.style.display = 'none';
      btnToggleFind.classList.remove('active');
      variablesModal.style.display = 'none';
      backgroundModal.style.display = 'none';
      tableModal.style.display = 'none';
      closeActiveEditors();
    }
  });
}

// Start
initializeWasm();
setupEventListeners();
