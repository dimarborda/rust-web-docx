// Demo app: the full editing experience published at rust-web-docx.dimarborda.workers.dev,
// built on the reusable @dimarborda/docx-editor package (the same API other projects use).
import {
  DocxEditor,
  listStoredFonts,
  loadStoredFonts,
  registerFont,
  removeStoredFont,
} from '@dimarborda/docx-editor';
import '@dimarborda/docx-editor/style.css';
import './fonts.js';

const $ = id => document.getElementById(id);

// Shell
const emptyStateView = $('empty-state-view');
const canvasDocumentView = $('canvas-document-view');
const docStatusContainer = $('doc-status-container');
const docToolbar = $('doc-toolbar');
const formattingRibbon = $('formatting-ribbon');
const appStatusbar = $('app-statusbar');
const currentDocTitle = $('current-doc-title');
const zoomLevelLabel = $('zoom-level-label');
const dropTarget = $('drop-target');
const btnDownload = $('btn-download');

// Search & replace
const searchReplaceBar = $('search-replace-bar');
const btnToggleFind = $('btn-toggle-find');
const findQuery = $('find-query');
const replaceQuery = $('replace-query');

// Modals
const variablesModal = $('variables-modal');
const variablesListContainer = $('variables-list-container');
const varsCountBadge = $('vars-count-badge');
const backgroundModal = $('background-modal');
const watermarkOpacity = $('watermark-opacity');
const tableModal = $('table-modal');
const fontsModal = $('fonts-modal');
const fontsCountBadge = $('fonts-count-badge');
const docFontsList = $('doc-fonts-list');
const installedFontsList = $('installed-fonts-list');
const fontDropZone = $('font-drop-zone');
const fontWarningBanner = $('font-warning-banner');
const fontWarningDesc = $('font-warning-desc');

// Ribbon
const btnFmtBold = $('btn-fmt-bold');
const btnFmtItalic = $('btn-fmt-italic');
const btnFmtUnderline = $('btn-fmt-underline');
const textColorInput = $('text-color-input');
const alignButtons = {
  left: $('btn-align-left'),
  center: $('btn-align-center'),
  right: $('btn-align-right'),
  both: $('btn-align-justify'),
};

let editor = null;
let watermark = { text: null, opacity: 0.2 };
// Missing fonts the user already dismissed the banner for; edits must not bring it back
let dismissedMissingFonts = null;

// ---------- Opening documents ----------

async function openFile(file) {
  try {
    await editor.open(file);
    showToast(`"${file.name}" cargado con vista fiel multi-página.`);
  } catch (err) {
    console.error('Error opening DOCX:', err);
    showToast('Error al abrir el documento DOCX: ' + err, true);
  }
}

function openSample() {
  try {
    editor.openSample();
    showToast('Documento demo cargado en vista fiel multi-página.');
  } catch (err) {
    console.error('Error generating sample:', err);
    showToast('Error al cargar documento demo: ' + err, true);
  }
}

function onDocumentLoaded() {
  dismissedMissingFonts = null;
  emptyStateView.style.display = 'none';
  canvasDocumentView.style.display = 'flex';
  docStatusContainer.style.display = 'flex';
  docToolbar.style.display = 'flex';
  formattingRibbon.style.display = 'flex';
  appStatusbar.style.display = 'flex';
  btnDownload.disabled = false;
  currentDocTitle.textContent = editor.fileName;
  onDocumentChanged();
}

function onDocumentChanged() {
  updateStats();
  updateVariablesBadge();
  checkDocumentFonts();
}

// ---------- Welcome screen folders ----------

// examples/ (listed by the dev server, see vite.config.js) and any folder the user picks.
// Nothing is uploaded; files are read in the browser.
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
        open: async () => new File([await (await fetch(`/__examples/${encodeURIComponent(f.name)}`)).blob()], f.name),
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
  const container = $('document-folders');
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
      row.addEventListener('click', async () => openFile(await file.open()));
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

// ---------- Status bar, ribbon, variables ----------

function updateStats() {
  const stats = editor.stats();
  if (!stats) return;
  $('stat-words').textContent = stats.word_count.toLocaleString();
  $('stat-chars').textContent = stats.char_count.toLocaleString();
  $('stat-paragraphs').textContent = stats.paragraph_count.toLocaleString();
  $('stat-tables').textContent = (stats.table_count || 0).toLocaleString();
  $('stat-pages').textContent = stats.page_count;
  if (stats.page_setup) {
    $('stat-format').textContent = stats.page_setup.orientation === 'landscape' ? 'Horizontal (Landscape)' : 'Vertical';
  }
}

function updateRibbon(format) {
  if (!format) return;
  btnFmtBold.classList.toggle('active', format.bold);
  btnFmtItalic.classList.toggle('active', format.italic);
  btnFmtUnderline?.classList.toggle('active', format.underline);
  Object.entries(alignButtons).forEach(([align, btn]) => btn?.classList.toggle('active', align === format.align));
  if (format.color) textColorInput.value = format.color;
}

function updateVariablesBadge() {
  const count = editor.variables().length;
  varsCountBadge.textContent = count;
  varsCountBadge.style.display = count > 0 ? 'inline-block' : 'none';
}

function openVariablesModal() {
  variablesListContainer.innerHTML = '';
  const variables = editor.variables();
  if (variables.length === 0) addVariableRow('{{EJEMPLO}}', 'Valor');
  else variables.forEach(v => addVariableRow(v, ''));
  variablesModal.style.display = 'flex';
}

function addVariableRow(key = '', val = '') {
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

function applyVariables() {
  const pairs = [...variablesListContainer.querySelectorAll('.var-item-row')]
    .map(r => ({ key: r.querySelector('.var-key-input').value.trim(), value: r.querySelector('.var-val-input').value }))
    .filter(p => p.key);
  if (pairs.length === 0) return showToast('No hay variables válidas para aplicar.', true);
  try {
    showToast(editor.replaceVariables(pairs).message);
    variablesModal.style.display = 'none';
  } catch (err) {
    console.error('Batch replace error:', err);
    showToast('Error al aplicar variables: ' + err, true);
  }
}

function executeSearchAndReplace() {
  const search = findQuery.value;
  if (!search) return showToast('Introduce un término a buscar.', true);
  try {
    const result = editor.findReplace(search, replaceQuery.value, {
      matchCase: $('chk-case-sensitive').checked,
      regex: $('chk-regex').checked,
    });
    if (result.occurrences_replaced > 0) showToast(result.message);
    else showToast(`No se encontraron coincidencias para "${search}".`, true);
  } catch (err) {
    showToast('Error en reemplazo: ' + err, true);
  }
}

function insertTable() {
  const rows = parseInt($('tbl-input-rows').value, 10) || 3;
  const cols = parseInt($('tbl-input-cols').value, 10) || 3;
  try {
    editor.insertTable(rows, cols, Array.from({ length: cols }, (_, i) => `Columna ${i + 1}`));
    tableModal.style.display = 'none';
    showToast(`Tabla de ${rows}x${cols} insertada exitosamente.`);
  } catch (err) {
    showToast('Error al insertar tabla: ' + err, true);
  }
}

function downloadDocx() {
  if (!editor.hasDocument) return;
  try {
    const name = editor.fileName.replace(/\.docx$/i, '');
    editor.download(`${name}_modificado.docx`);
    showToast('¡Archivo DOCX descargado exitosamente!');
  } catch (err) {
    showToast('Error al exportar DOCX: ' + err, true);
  }
}

function setWatermark(text) {
  watermark = { text, opacity: parseInt(watermarkOpacity.value, 10) / 100 };
  editor.setWatermark(watermark.text, { opacity: watermark.opacity });
}

function setZoom(delta) {
  const zoom = editor.zoom + delta;
  if (zoom < 0.5 || zoom > 1.5) return;
  editor.setZoom(zoom);
  zoomLevelLabel.textContent = `${Math.round(zoom * 100)}%`;
}

// ---------- Fonts ----------

function checkDocumentFonts() {
  const fonts = editor.fonts();
  const missing = fonts.filter(f => f.substitute).map(f => f.name);
  const key = missing.join('|');
  if (missing.length > 0) {
    fontWarningBanner.style.display = key === dismissedMissingFonts ? 'none' : 'flex';
    fontWarningBanner.dataset.missingFonts = key;
    fontWarningDesc.innerHTML = `El documento solicita <strong>${missing.map(escapeHtml).join(', ')}</strong> (usando aproximación web). Sube tu archivo <code>.ttf/.otf</code> para obtener fidelidad 100% idéntica a Word.`;
    fontsCountBadge.style.display = 'inline-flex';
    fontsCountBadge.textContent = missing.length;
  } else {
    fontWarningBanner.style.display = 'none';
    fontsCountBadge.style.display = 'none';
  }

  docFontsList.innerHTML = fonts.length === 0
    ? '<div class="font-status-row"><span class="font-name-info">Ninguna fuente específica detectada</span></div>'
    : fonts.map(f => `
      <div class="font-status-row">
        <div class="font-name-info"><span>🔤</span><span>${escapeHtml(f.name)}</span></div>
        ${f.substitute
          ? '<span class="font-status-tag fallback">⚠️ Usando aproximación web</span>'
          : '<span class="font-status-tag available">✓ Disponible</span>'}
      </div>`).join('');
}

async function renderInstalledFonts() {
  const stored = await listStoredFonts();
  if (stored.length === 0) {
    installedFontsList.innerHTML = '<div class="font-status-row"><span class="font-name-info" style="color: var(--text-muted); font-size: 0.8rem;">No hay fuentes personalizadas subidas aún.</span></div>';
    return;
  }
  installedFontsList.innerHTML = stored.map(item => `
    <div class="font-status-row">
      <div class="font-name-info">
        <span>📄</span>
        <span><strong>${escapeHtml(item.family)}</strong> (${escapeHtml(item.fileName || 'fuente.ttf')})</span>
      </div>
      <button class="btn-delete-font" data-id="${escapeHtml(item.id)}" title="Eliminar fuente de la memoria">✕ Eliminar</button>
    </div>`).join('');
  installedFontsList.querySelectorAll('.btn-delete-font').forEach(btn => {
    btn.addEventListener('click', async () => {
      await removeStoredFont(btn.dataset.id);
      showToast('Fuente eliminada de la memoria');
      if (editor.hasDocument) checkDocumentFonts();
      renderInstalledFonts();
    });
  });
}

async function addFontFiles(files) {
  let count = 0;
  for (const file of files) {
    try {
      await registerFont(file);
      count++;
    } catch (err) {
      console.error('Error procesando archivo de fuente:', file.name, err);
      showToast(`Error al cargar fuente ${file.name}: ${err.message}`, true);
    }
  }
  if (count > 0) {
    showToast(`✓ ${count} fuente(s) instalada(s) y activada(s) en memoria`);
    if (editor.hasDocument) checkDocumentFonts();
    renderInstalledFonts();
  }
}

function openFontsModal() {
  fontsModal.style.display = 'flex';
  if (editor.hasDocument) checkDocumentFonts();
  renderInstalledFonts();
}

// ---------- Utilities ----------

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

let toastTimer = null;
function showToast(msg, isError = false) {
  const toast = $('toast');
  $('toast-text').textContent = msg;
  toast.style.background = isError ? 'rgba(239, 68, 68, 0.15)' : 'rgba(16, 185, 129, 0.15)';
  toast.style.borderColor = isError ? 'rgba(239, 68, 68, 0.4)' : 'rgba(16, 185, 129, 0.4)';
  toast.style.color = isError ? '#FCA5A5' : '#A7F3D0';
  toast.style.display = 'flex';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { toast.style.display = 'none'; }, 4000);
}

const show = el => { el.style.display = 'flex'; };
const hide = el => { el.style.display = 'none'; };

// ---------- Wiring ----------

function setupEventListeners() {
  editor.addEventListener('load', onDocumentLoaded);
  editor.addEventListener('change', onDocumentChanged);
  editor.addEventListener('selectionchange', e => updateRibbon(e.detail.format));
  editor.addEventListener('message', e => showToast(e.detail.message, e.detail.error));

  $('btn-zoom-in').addEventListener('click', () => setZoom(0.25));
  $('btn-zoom-out').addEventListener('click', () => setZoom(-0.25));

  // Ribbon buttons must not take focus from the document (the caret stays put)
  document.querySelectorAll('.formatting-ribbon button').forEach(btn => {
    btn.addEventListener('mousedown', e => e.preventDefault());
  });

  [$('file-input'), $('file-input-welcome')].forEach(input => input?.addEventListener('change', e => {
    if (e.target.files?.[0]) openFile(e.target.files[0]);
  }));
  [$('btn-load-sample'), $('btn-welcome-demo')].forEach(btn => btn?.addEventListener('click', openSample));

  $('examples-folder-input')?.addEventListener('change', e => {
    if (e.target.files?.length) usePickedFolder(e.target.files);
    e.target.value = '';
  });

  btnDownload.addEventListener('click', downloadDocx);

  if (dropTarget) {
    ['dragenter', 'dragover'].forEach(name => dropTarget.addEventListener(name, e => {
      e.preventDefault();
      dropTarget.classList.add('drag-over');
    }));
    ['dragleave', 'drop'].forEach(name => dropTarget.addEventListener(name, e => {
      e.preventDefault();
      dropTarget.classList.remove('drag-over');
    }));
    dropTarget.addEventListener('drop', e => {
      if (e.dataTransfer?.files?.[0]) openFile(e.dataTransfer.files[0]);
    });
  }

  // Formatting
  btnFmtBold.addEventListener('click', () => editor.toggleBold());
  btnFmtItalic.addEventListener('click', () => editor.toggleItalic());
  btnFmtUnderline?.addEventListener('click', () => editor.toggleUnderline());
  Object.entries(alignButtons).forEach(([align, btn]) => btn?.addEventListener('click', () => editor.setAlignment(align)));
  // The native color picker keeps focus while open
  textColorInput.addEventListener('input', e => editor.setColor(e.target.value, { refocus: false }));
  document.querySelectorAll('.color-swatch-dot').forEach(dot => dot.addEventListener('click', () => {
    textColorInput.value = dot.dataset.color;
    editor.setColor(dot.dataset.color);
  }));

  // Table
  $('btn-insert-table').addEventListener('click', () => show(tableModal));
  $('btn-close-table-modal').addEventListener('click', () => hide(tableModal));
  $('btn-create-table-confirm').addEventListener('click', insertTable);

  // Watermark
  $('btn-toggle-background-modal').addEventListener('click', () => show(backgroundModal));
  $('btn-close-bg-modal').addEventListener('click', () => hide(backgroundModal));
  $('btn-save-bg-config').addEventListener('click', () => {
    hide(backgroundModal);
    showToast('Configuración de fondo aplicada.');
  });
  $('btn-wm-confidential').addEventListener('click', () => {
    setWatermark('CONFIDENCIAL');
    showToast('Marca de agua "CONFIDENCIAL" aplicada.');
  });
  $('btn-wm-draft').addEventListener('click', () => {
    setWatermark('BORRADOR');
    showToast('Marca de agua "BORRADOR" aplicada.');
  });
  $('btn-wm-clear').addEventListener('click', () => {
    setWatermark(null);
    showToast('Marca de agua eliminada.');
  });
  watermarkOpacity.addEventListener('input', e => {
    $('opacity-val-label').textContent = `${e.target.value}%`;
    if (watermark.text) setWatermark(watermark.text);
  });

  // Search
  const openSearch = () => {
    show(searchReplaceBar);
    btnToggleFind.classList.add('active');
    findQuery.focus();
    findQuery.select();
  };
  const closeSearch = () => {
    hide(searchReplaceBar);
    btnToggleFind.classList.remove('active');
  };
  btnToggleFind.addEventListener('click', () => (searchReplaceBar.style.display === 'flex' ? closeSearch() : openSearch()));
  $('btn-close-search').addEventListener('click', closeSearch);
  $('btn-do-replace').addEventListener('click', executeSearchAndReplace);
  [findQuery, replaceQuery].forEach(inp => inp.addEventListener('keydown', e => {
    if (e.key === 'Enter') executeSearchAndReplace();
  }));

  // Variables
  $('btn-toggle-variables').addEventListener('click', openVariablesModal);
  $('btn-close-variables').addEventListener('click', () => hide(variablesModal));
  $('btn-add-custom-var').addEventListener('click', () => addVariableRow());
  $('btn-apply-all-variables').addEventListener('click', applyVariables);

  // Fonts
  $('btn-toggle-fonts').addEventListener('click', openFontsModal);
  $('btn-banner-upload-fonts').addEventListener('click', openFontsModal);
  $('btn-close-fonts').addEventListener('click', () => hide(fontsModal));
  $('btn-done-fonts').addEventListener('click', () => hide(fontsModal));
  $('btn-banner-dismiss-fonts').addEventListener('click', () => {
    hide(fontWarningBanner);
    dismissedMissingFonts = fontWarningBanner.dataset.missingFonts || null;
  });
  $('font-file-input').addEventListener('change', e => {
    if (e.target.files?.length) addFontFiles(e.target.files);
    e.target.value = '';
  });
  fontDropZone.addEventListener('dragover', e => {
    e.preventDefault();
    fontDropZone.classList.add('drag-over');
  });
  fontDropZone.addEventListener('dragleave', e => {
    e.preventDefault();
    fontDropZone.classList.remove('drag-over');
  });
  fontDropZone.addEventListener('drop', e => {
    e.preventDefault();
    fontDropZone.classList.remove('drag-over');
    if (e.dataTransfer?.files?.length) addFontFiles(e.dataTransfer.files);
  });

  const modals = [variablesModal, backgroundModal, tableModal, fontsModal];
  modals.forEach(modal => modal.addEventListener('click', e => {
    if (e.target === modal) hide(modal);
  }));

  // App shortcuts (bold/italic/underline are handled by the editor itself)
  window.addEventListener('keydown', e => {
    const mod = e.metaKey || e.ctrlKey;
    if (mod && e.key.toLowerCase() === 'f') {
      e.preventDefault();
      openSearch();
    } else if (mod && e.key.toLowerCase() === 's') {
      e.preventDefault();
      downloadDocx();
    } else if (e.key === 'Escape') {
      closeSearch();
      modals.forEach(hide);
    }
  });
}

// ---------- Start ----------

if ('scrollRestoration' in history) history.scrollRestoration = 'manual';
(async () => {
  await loadStoredFonts();
  editor = await DocxEditor.create($('canvas-pages-wrapper'), { locale: 'es' });
  setupEventListeners();
  documentFolders.examples = await loadExamplesFolder();
  renderDocumentFolders();
  // ?demo opens the demo contract right away (handy for shared links and previews)
  if (new URLSearchParams(location.search).has('demo')) openSample();
})();
