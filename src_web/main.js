import init, { DocxSession } from '../pkg/rust_web_docx.js';
import { createCanvasEditor } from './canvas_editor.js';

// Application State
let wasmReady = false;
let currentSession = null;
let currentFileName = 'documento.docx';
let activeDocumentElements = [];
let activeParagraphIndex = null;
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

// Folders of documents shown on the welcome screen like a file browser: examples/ (listed
// by the dev server on every page load, see vite.config.js) and any folder the user picks
// with "Elegir carpeta…". Nothing is uploaded; files are read in the browser.
const documentFolders = { examples: null, picked: null };

const isDocx = name => /\.docx$/i.test(name) && !name.startsWith('~$'); // skip Word lock files

async function loadExamplesFolder() {
  try {
    const res = await fetch('/__examples.json', { cache: 'no-store' });
    if (!res.ok || !(res.headers.get('content-type') || '').includes('json')) return null;
    const { folder, files } = await res.json();
    return {
      name: folder,
      hint: 'Coloca aquí tus archivos .docx y recarga la página.',
      files: files.filter(f => isDocx(f.name)).map(f => ({
        ...f,
        open: async () => new File([await (await fetch(`/${folder}/${encodeURIComponent(f.name)}`)).blob()], f.name),
      })),
    };
  } catch {
    return null; // production build: there is no examples/ listing
  }
}

function usePickedFolder(fileList) {
  const name = (fileList[0]?.webkitRelativePath || '').split('/')[0] || 'Carpeta seleccionada';
  const files = [...fileList]
    .filter(f => isDocx(f.name))
    .sort((a, b) => a.name.localeCompare(b.name))
    .map(file => ({ name: file.name, size: file.size, modified: file.lastModified, open: async () => file }));
  documentFolders.picked = { name, hint: 'Esta carpeta no tiene archivos .docx.', files };
  renderDocumentFolders();
}

function formatFileSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

const fileDateFormat = new Intl.DateTimeFormat('es', { day: 'numeric', month: 'short', year: 'numeric' });

function renderDocumentFolders() {
  const container = document.getElementById('document-folders');
  if (!container) return;
  const folders = [documentFolders.picked, documentFolders.examples].filter(Boolean);

  container.replaceChildren(...folders.map(folder => {
    const block = document.createElement('div');
    block.className = 'folder-block';
    const count = folder.files.length;
    block.innerHTML = `
      <div class="folder-header">
        <span class="folder-name">📁 ${escapeHtml(folder.name)}/</span>
        <span class="folder-count">${count} archivo${count === 1 ? '' : 's'}</span>
      </div>`;
    if (count === 0) {
      const empty = document.createElement('p');
      empty.className = 'folder-empty';
      empty.textContent = folder.hint;
      block.appendChild(empty);
    }
    const list = document.createElement('ul');
    list.className = 'folder-files';
    folder.files.forEach(file => {
      const item = document.createElement('li');
      const row = document.createElement('button');
      row.className = 'folder-file';
      row.title = `Abrir ${file.name}`;
      row.innerHTML = `
        <span class="file-icon">📄</span>
        <span class="file-name">${escapeHtml(file.name)}</span>
        <span class="file-meta">${formatFileSize(file.size)}</span>
        <span class="file-meta">${file.modified ? fileDateFormat.format(new Date(file.modified)) : ''}</span>`;
      row.addEventListener('click', async () => loadDocxFile(await file.open()));
      item.appendChild(row);
      list.appendChild(item);
    });
    block.appendChild(list);
    return block;
  }));

  if (folders.length === 0) {
    container.innerHTML = '<p class="folder-empty">Elige una carpeta con documentos <code>.docx</code> para verlos aquí.</p>';
  }
}

async function initDocumentFolders() {
  documentFolders.examples = await loadExamplesFolder();
  renderDocumentFolders();
}

// 3. Document Loaded Transition
function onDocumentLoaded() {
  canvasEditor.clear();
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
// What the page canvases currently show, to redraw only pages whose content changed
let renderedState = { session: null, zoom: null, cards: new Map() };

function renderCanvasPagesFromWasm({ full = false } = {}) {
  if (!currentSession) return;

  try {
    const elementsJson = currentSession.get_document_elements_json();
    activeDocumentElements = JSON.parse(elementsJson);

    const wmText = currentWatermark.type === 'preset' ? currentWatermark.text : null;
    const wmOpacity = currentWatermark.opacity || 0.2;

    // Zoom, a new document or new fonts change every page: redraw them all
    const redrawAll = full || renderedState.session !== currentSession || renderedState.zoom !== currentZoom;
    const previous = canvasPagesLayout;
    const layoutJson = currentSession.compute_canvas_layout_json(wmText, wmOpacity, measureTextForLayout, !redrawAll);
    const layout = JSON.parse(layoutJson);
    const freshPages = new Set(layout.pages.filter(p => !p.unchanged).map(p => p.page_number));
    layout.pages = layout.pages.map((page, i) => (page.unchanged ? previous.pages[i] : page));
    canvasPagesLayout = layout;

    statPages.textContent = layout.total_pages;
    canvasEditor.beginRender();
    if (redrawAll) {
      canvasPagesWrapper.innerHTML = '';
      renderedState = { session: currentSession, zoom: currentZoom, cards: new Map() };
    }

    const dpr = window.devicePixelRatio || 1;

    layout.pages.forEach(page => {
      let card = renderedState.cards.get(page.page_number);
      if (!card || freshPages.has(page.page_number)) {
        const fresh = buildPageCard(page, layout.total_pages, dpr);
        if (card) card.replaceWith(fresh);
        else canvasPagesWrapper.appendChild(fresh);
        card = fresh;
        renderedState.cards.set(page.page_number, card);
      }
      // Caret, selection and table cell editing on top of the canvas
      canvasEditor.attachPage(page, card, card.querySelector('canvas'));
    });

    // The document got shorter: drop the extra pages
    renderedState.cards.forEach((card, number) => {
      if (number > layout.total_pages) {
        card.remove();
        renderedState.cards.delete(number);
      }
    });

    ensureLayoutFonts(layout);
    canvasEditor.endRender();

  } catch (err) {
    console.error('Error rendering Canvas layout:', err);
    showToast('Error al renderizar páginas: ' + err, true);
  }
}

function buildPageCard(page, totalPages, dpr) {
  const pageCard = document.createElement('div');
  pageCard.className = 'canvas-page-card';
  pageCard.dataset.pageNumber = page.page_number;

  const pageHeader = document.createElement('div');
  pageHeader.className = 'canvas-page-header';
  pageHeader.innerHTML = `<span class="page-badge">Página ${page.page_number} de ${totalPages}</span>`;
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
  disableLigatures(ctx);
  ctx.scale(dpr * currentZoom, dpr * currentZoom);
  ctx.fillStyle = page.bg_color || currentPageBgColor || '#FFFFFF';
  ctx.fillRect(0, 0, page.width, page.height);
  drawCanvasPageItems(ctx, page.items);

  pageCard.appendChild(canvas);
  return pageCard;
}

// Rust lays out text with these measurements, so line breaks match what the canvas draws
const layoutMeasureCtx = document.createElement('canvas').getContext('2d');
disableLigatures(layoutMeasureCtx);

// Word does not apply ligatures ("fi") or kerning by default, and with them off every glyph
// sits exactly at the sum of the advances measured for the caret
function disableLigatures(ctx) {
  if ('textRendering' in ctx) ctx.textRendering = 'optimizeSpeed';
  if ('fontKerning' in ctx) ctx.fontKerning = 'none';
}

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
      if (results.some(loaded => loaded.length > 0) && currentSession) {
        // Widths measured with fallback fonts are stale now
        currentSession.reset_measurements();
        renderCanvasPagesFromWasm({ full: true });
      }
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
            // Redraw the pages once the image can be drawn
            renderCanvasPagesFromWasm({ full: true });
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
      if (item.runs && item.runs.length > 0) {
        // Runs are drawn exactly where the layout placed them (run.x), with the layout's
        // justification spacing, so the caret geometry computed in Rust matches the pixels
        const extra = item.line ? item.line.space_extra : 0;
        let nextX = item.x;
        item.runs.forEach(run => {
          const startX = run.x ?? nextX;
          if (run.text === '\t') {
            nextX = startX + (run.width || 0);
            return;
          }
          const fontSize = run.font_size || item.font_size || 14.66;
          const fontFamily = run.font_family || item.font_family || 'Calibri, sans-serif';
          const runColor = formatCssColor(run.color || item.color, '#1E293B');
          ctx.font = buildCanvasFont(run.bold ? '700' : '400', run.italic ? 'italic' : 'normal', fontSize, fontFamily);
          ctx.fillStyle = runColor;
          ctx.textAlign = 'left';
          ctx.textBaseline = 'alphabetic';

          let curX = startX;
          if (extra > 0 && run.text.includes(' ')) {
            run.text.split(' ').forEach((word, i) => {
              if (i > 0) curX += ctx.measureText(' ').width + extra;
              if (word) {
                ctx.fillText(word, curX, item.y);
                curX += ctx.measureText(word).width;
              }
            });
          } else {
            ctx.fillText(run.text, curX, item.y);
            curX += ctx.measureText(run.text).width;
          }

          if (run.underline) {
            ctx.beginPath();
            ctx.moveTo(startX, item.y + 2);
            ctx.lineTo(curX, item.y + 2);
            ctx.strokeStyle = runColor;
            ctx.lineWidth = 1;
            ctx.stroke();
          }
          nextX = curX;
        });
      } else if (item.text) {
        // Page decorations (header, footer, page numbers)
        ctx.font = buildCanvasFont(item.font_weight, item.font_style, item.font_size || 14.66, item.font_family);
        ctx.fillStyle = formatCssColor(item.color, '#1E293B');
        ctx.textAlign = item.align === 'right' ? 'right' : (item.align === 'center' ? 'center' : 'left');
        ctx.textBaseline = 'alphabetic';
        ctx.fillText(item.text, item.x, item.y);
      }
    }
  });
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

function setSyncStatus(isSynced) {
  if (isSynced) {
    syncStatus.className = 'status-badge';
    syncStatus.innerHTML = '<span class="status-dot"></span> Sincronizado';
  } else {
    syncStatus.className = 'status-badge saving';
    syncStatus.innerHTML = '<span class="status-dot"></span> Guardando...';
  }
}

// 12. Formatting: applies to the selected characters (across paragraphs), or to the whole
// paragraph at the caret when nothing is selected. Each action is one undo step.
// Every paragraph of the document in order (body and table cells), each tagged with its
// container ("body" or "table:row:col") so the editor never joins paragraphs across cells
let paragraphListCache = { source: null, list: [], byIndex: new Map() };

function documentParagraphs() {
  if (paragraphListCache.source !== activeDocumentElements) {
    const list = [];
    activeDocumentElements.forEach(el => {
      if (el.type === 'paragraph') {
        list.push({ ...el, container: 'body' });
      } else if (el.type === 'table') {
        (el.rich_rows || []).forEach((row, r) => row.cells.forEach((cell, c) => {
          (cell.paragraphs || []).forEach(p => list.push({ ...p, container: `${el.index}:${r}:${c}` }));
        }));
      }
    });
    list.sort((a, b) => a.index - b.index);
    paragraphListCache = { source: activeDocumentElements, list, byIndex: new Map(list.map(p => [p.index, p])) };
  }
  return paragraphListCache;
}

function paragraphElement(index) {
  return documentParagraphs().byIndex.get(index);
}

/** One entry per character, carrying the formatting of its run */
function paragraphChars(p) {
  const chars = [];
  (p.runs || []).forEach(run => Array.from(run.text).forEach(ch => chars.push({ ...run, text: ch })));
  return chars;
}

function sameFormat(a, b) {
  return a.bold === b.bold && a.italic === b.italic && a.underline === b.underline &&
    (a.color || '') === (b.color || '') && a.font_size === b.font_size && a.font_family === b.font_family;
}

function mergeRuns(chars) {
  const runs = [];
  chars.forEach(c => {
    const last = runs[runs.length - 1];
    if (last && sameFormat(last, c)) last.text += c.text;
    else runs.push({ ...c });
  });
  return runs;
}

/** Paragraphs touched by the selection, each with its selected character range */
function selectedParagraphRanges() {
  const sel = canvasEditor.selection();
  if (!sel) return null;
  const items = [];
  for (let i = sel.start.paragraph; i <= sel.end.paragraph; i++) {
    const p = paragraphElement(i);
    if (!p) continue;
    const chars = paragraphChars(p);
    const start = sel.collapsed || i !== sel.start.paragraph ? 0 : sel.start.offset;
    const end = sel.collapsed || i !== sel.end.paragraph ? chars.length : sel.end.offset;
    items.push({ p, chars, start, end });
  }
  return items;
}

function applyParagraphUpdates(updates, refocus) {
  if (updates.length === 0) return;
  try {
    currentSession.update_paragraphs_runs(JSON.stringify(updates));
  } catch (err) {
    console.error('Format error:', err);
    showToast('No se pudo aplicar el formato: ' + err, true);
    return;
  }
  renderCanvasPagesFromWasm();
  if (refocus) canvasEditor.focus();
}

function applyRunFormat(change, { refocus = true } = {}) {
  const items = selectedParagraphRanges();
  if (!items) {
    showToast('Haz clic en el texto para aplicar formato.');
    return;
  }
  const range = items.flatMap(it => it.chars.slice(it.start, it.end));
  if (range.length === 0) return;
  change(range);
  const updates = items
    .filter(it => it.end > it.start)
    .map(it => ({ index: it.p.index, runs: mergeRuns(it.chars), align: it.p.align || 'left' }));
  applyParagraphUpdates(updates, refocus);
}

function toggleRunFlag(flag) {
  applyRunFormat(range => {
    const on = !range.every(c => c[flag]);
    range.forEach(c => { c[flag] = on; });
  });
}

function toggleBold() {
  toggleRunFlag('bold');
}

function toggleItalic() {
  toggleRunFlag('italic');
}

function toggleUnderline() {
  toggleRunFlag('underline');
}

function applyTextColor(hexColor) {
  if (!hexColor) return;
  const clean = hexColor.replace('#', '').toUpperCase();
  // The native color picker keeps focus while open
  applyRunFormat(range => range.forEach(c => { c.color = clean; }), { refocus: false });
}

function applyAlignment(alignValue) {
  const items = selectedParagraphRanges();
  if (!items) {
    showToast('Haz clic en un párrafo para alinearlo.');
    return;
  }
  applyParagraphUpdates(items.map(it => ({ index: it.p.index, runs: it.p.runs || [], align: alignValue })), true);
}

/** Reflects the formatting at the caret (or of the selection) in the ribbon */
function updateRibbonForSelection(sel) {
  activeParagraphIndex = sel ? sel.focus.paragraph : null;
  if (!sel) return;
  const p = paragraphElement(sel.start.paragraph);
  if (!p) return;
  updateRibbonAlignUI(p.align || 'left');

  const chars = paragraphChars(p);
  const range = sel.collapsed
    ? chars.slice(Math.max(0, sel.start.offset - 1), Math.max(1, sel.start.offset))
    : chars.slice(sel.start.offset, sel.start.paragraph === sel.end.paragraph ? sel.end.offset : chars.length);
  const all = flag => range.length > 0 && range.every(c => c[flag]);
  btnFmtBold.classList.toggle('active', all('bold'));
  btnFmtItalic.classList.toggle('active', all('italic'));
  if (btnFmtUnderline) btnFmtUnderline.classList.toggle('active', all('underline'));

  const color = (range[0]?.color || p.color || '').replace('#', '');
  if (/^[0-9a-fA-F]{6}$/.test(color)) textColorInput.value = `#${color}`;
}

function updateRibbonAlignUI(align) {
  Object.keys(alignButtons).forEach(k => {
    if (alignButtons[k]) {
      alignButtons[k].classList.toggle('active', k === align || (k === 'both' && (align === 'justify' || align === 'both')));
    }
  });
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
  // Zoom Controls
  btnZoomIn.addEventListener('click', () => {
    if (currentZoom < 1.5) {
      currentZoom += 0.25;
      zoomLevelLabel.textContent = `${Math.round(currentZoom * 100)}%`;
      renderCanvasPagesFromWasm();
    }
  });

  btnZoomOut.addEventListener('click', () => {
    if (currentZoom > 0.5) {
      currentZoom -= 0.25;
      zoomLevelLabel.textContent = `${Math.round(currentZoom * 100)}%`;
      renderCanvasPagesFromWasm();
    }
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

  // Document folders: examples/ (dev server) and any folder chosen by the user
  initDocumentFolders();
  document.getElementById('examples-folder-input')?.addEventListener('change', e => {
    if (e.target.files?.length) usePickedFolder(e.target.files);
    e.target.value = '';
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
    }
  });
}

// Canvas-native caret and selection (replaces the per-paragraph edit box)
const canvasEditor = createCanvasEditor({
  session: () => currentSession,
  paragraphs: () => documentParagraphs(),
  zoom: () => currentZoom,
  measure: measureTextForLayout,
  documentChanged: () => {
    renderCanvasPagesFromWasm();
    updateLiveStats();
    setSyncStatus(true);
  },
  selectionChanged: updateRibbonForSelection,
  hint: msg => showToast(msg),
});

// Start
if ('scrollRestoration' in history) history.scrollRestoration = 'manual';
initializeWasm();
setupEventListeners();
