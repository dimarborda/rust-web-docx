import init, { DocxSession } from '../pkg/rust_web_docx.js';

// Application State
let wasmReady = false;
let currentSession = null;
let currentFileName = 'documento.docx';
let activeParagraphs = [];
let replacementHistory = [];

// DOM Elements
const dropzoneSection = document.getElementById('dropzone-section');
const editorDashboard = document.getElementById('editor-dashboard');
const dropzoneBox = document.getElementById('dropzone-box');
const fileInput = document.getElementById('file-upload-input');
const fileInput2 = document.getElementById('file-upload-input-2');
const btnSampleDocx = document.getElementById('btn-sample-docx');
const btnSampleDocx2 = document.getElementById('btn-sample-docx-2');
const btnDownloadDocx = document.getElementById('btn-download-docx');

// Meta elements
const metaDocName = document.getElementById('meta-doc-name');
const metaParagraphsCount = document.getElementById('meta-paragraphs-count');
const metaWordsCount = document.getElementById('meta-words-count');
const metaCharsCount = document.getElementById('meta-chars-count');
const metaFileSize = document.getElementById('meta-file-size');

// Toast banner
const toastBanner = document.getElementById('toast-banner');
const toastMessage = document.getElementById('toast-message');

// Tab buttons & panels
const tabButtons = document.querySelectorAll('.tab-btn');
const tabPanels = document.querySelectorAll('.tab-panel');

// Find & Replace elements
const findInput = document.getElementById('find-input');
const replaceInput = document.getElementById('replace-input');
const chkMatchCase = document.getElementById('chk-match-case');
const chkUseRegex = document.getElementById('chk-use-regex');
const btnExecuteReplace = document.getElementById('btn-execute-replace');
const replaceHistoryBox = document.getElementById('replace-history-box');
const replaceHistoryList = document.getElementById('replace-history-list');

// Paragraph Editor elements
const paragraphsContainer = document.getElementById('paragraphs-container');
const paragraphFilterInput = document.getElementById('paragraph-filter-input');
const btnSaveAllParagraphs = document.getElementById('btn-save-all-paragraphs');

// Batch Variables elements
const variablesTableBody = document.getElementById('variables-table-body');
const btnAddVariableRow = document.getElementById('btn-add-variable-row');
const btnDetectTemplateVars = document.getElementById('btn-detect-template-vars');
const btnApplyBatchVariables = document.getElementById('btn-apply-batch-variables');

// Full Text Preview
const fullTextPreview = document.getElementById('full-text-preview');
const btnCopyFullText = document.getElementById('btn-copy-full-text');

// ZIP structure
const zipFilesList = document.getElementById('zip-files-list');

// 1. Initialize WebAssembly
async function initializeWasm() {
  try {
    await init();
    wasmReady = true;
    console.log('⚡ Rust WebAssembly module initialized successfully!');
  } catch (err) {
    console.error('Failed to initialize WebAssembly:', err);
    showToast('Error al cargar WebAssembly: ' + err, true);
  }
}

// 2. Load File into WASM Session
async function loadDocxFile(file) {
  if (!wasmReady) await initializeWasm();
  try {
    const arrayBuffer = await file.arrayBuffer();
    const bytes = new Uint8Array(arrayBuffer);
    currentFileName = file.name;
    
    // Create new WASM Session
    currentSession = new DocxSession(bytes);
    refreshDocumentView();
    showToast(`Archivo "${currentFileName}" cargado correctamente.`);
  } catch (err) {
    console.error('Error loading DOCX:', err);
    showToast('Error al procesar el archivo DOCX: ' + err, true);
  }
}

// Load Sample Document
async function loadSampleDocx() {
  if (!wasmReady) await initializeWasm();
  try {
    currentFileName = 'contrato_ejemplo.docx';
    currentSession = DocxSession.new_sample();
    refreshDocumentView();
    showToast('Documento de demostración cargado exitosamente.');
  } catch (err) {
    console.error('Error generating sample docx:', err);
    showToast('Error al generar documento demo: ' + err, true);
  }
}

// 3. Refresh UI & Extract Metadata from WASM
function refreshDocumentView() {
  if (!currentSession) return;

  try {
    // Show dashboard, hide dropzone
    dropzoneSection.style.display = 'none';
    editorDashboard.style.display = 'flex';
    btnDownloadDocx.disabled = false;

    // Retrieve stats from Rust
    const statsJson = currentSession.get_stats_json();
    const stats = JSON.parse(statsJson);

    metaDocName.textContent = currentFileName;
    metaParagraphsCount.textContent = stats.paragraph_count.toLocaleString();
    metaWordsCount.textContent = stats.word_count.toLocaleString();
    metaCharsCount.textContent = stats.char_count.toLocaleString();
    metaFileSize.textContent = formatBytes(stats.original_size_bytes);

    // Retrieve paragraphs from Rust
    const paragraphsJson = currentSession.get_paragraphs_json();
    activeParagraphs = JSON.parse(paragraphsJson);
    renderParagraphs(activeParagraphs);

    // Retrieve full text
    const fullText = currentSession.get_raw_text();
    fullTextPreview.value = fullText;

    // Render zip files structure
    renderZipFiles(stats.files_in_zip);

    // Auto-detect template placeholders for the batch tab
    detectTemplateVariables(fullText);

  } catch (err) {
    console.error('Error refreshing document view:', err);
    showToast('Error al actualizar la vista: ' + err, true);
  }
}

// 4. Render Paragraph Editor
function renderParagraphs(paragraphs) {
  paragraphsContainer.innerHTML = '';

  const filterText = (paragraphFilterInput.value || '').toLowerCase().trim();

  const filtered = paragraphs.filter(p => {
    if (!filterText) return true;
    return p.text.toLowerCase().includes(filterText) || p.style.toLowerCase().includes(filterText);
  });

  if (filtered.length === 0) {
    paragraphsContainer.innerHTML = `
      <div style="text-align: center; padding: 2rem; color: var(--text-dim);">
        No se encontraron párrafos que coincidan con el filtro.
      </div>`;
    return;
  }

  filtered.forEach(p => {
    const card = document.createElement('div');
    card.className = `paragraph-card ${p.is_heading ? 'heading-type' : ''}`;
    card.dataset.index = p.index;

    card.innerHTML = `
      <div class="paragraph-meta">
        <div>
          <span class="p-tag-index">#Párrafo ${p.index + 1}</span>
          <span class="p-style-badge ${p.is_heading ? 'badge-heading' : ''}">${escapeHtml(p.style)}</span>
        </div>
        <span>${p.text.length} caracteres • ${p.run_count} fragmentos</span>
      </div>
      <textarea class="paragraph-textarea" data-index="${p.index}">${escapeHtml(p.text)}</textarea>
      <div class="paragraph-actions">
        <button class="btn btn-secondary btn-sm btn-save-p" data-index="${p.index}">
          Guardar Párrafo
        </button>
      </div>
    `;

    // Event listener for single paragraph save
    const btnSave = card.querySelector('.btn-save-p');
    const textarea = card.querySelector('.paragraph-textarea');

    textarea.addEventListener('input', () => {
      if (textarea.value !== p.text) {
        card.classList.add('modified');
      } else {
        card.classList.remove('modified');
      }
    });

    btnSave.addEventListener('click', () => {
      saveSingleParagraph(p.index, textarea.value, card);
    });

    paragraphsContainer.appendChild(card);
  });
}

function saveSingleParagraph(index, newText, cardElement) {
  try {
    const updated = currentSession.update_paragraph(index, newText);
    if (updated) {
      if (cardElement) cardElement.classList.remove('modified');
      showToast(`Párrafo #${index + 1} actualizado.`);
      refreshDocumentView();
    }
  } catch (err) {
    showToast('Error al actualizar párrafo: ' + err, true);
  }
}

// 5. Execute Find & Replace
function handleExecuteReplace() {
  if (!currentSession) return;

  const search = findInput.value;
  const replace = replaceInput.value;
  const matchCase = chkMatchCase.checked;
  const useRegex = chkUseRegex.checked;

  if (!search) {
    showToast('Por favor introduce un término de búsqueda.', true);
    return;
  }

  try {
    const resultJson = currentSession.find_and_replace(search, replace, matchCase, useRegex);
    const result = JSON.parse(resultJson);

    if (result.occurrences_replaced > 0) {
      showToast(result.message);
      
      // Add to history
      replacementHistory.unshift({
        search,
        replace,
        count: result.occurrences_replaced,
        time: new Date().toLocaleTimeString()
      });
      renderHistory();
      
      refreshDocumentView();
    } else {
      showToast('No se encontraron coincidencias para: "' + search + '"', true);
    }
  } catch (err) {
    console.error('Find/replace error:', err);
    showToast('Error en reemplazo: ' + err, true);
  }
}

function renderHistory() {
  if (replacementHistory.length === 0) {
    replaceHistoryBox.style.display = 'none';
    return;
  }

  replaceHistoryBox.style.display = 'block';
  replaceHistoryList.innerHTML = replacementHistory
    .slice(0, 5)
    .map(item => `
      <li class="history-item">
        <div>
          <strong>"${escapeHtml(item.search)}"</strong> → <span style="color: var(--accent-success);">"${escapeHtml(item.replace)}"</span>
          <span style="color: var(--text-dim); margin-left: 8px;">(${item.time})</span>
        </div>
        <span class="history-badge">${item.count} reemplazo${item.count > 1 ? 's' : ''}</span>
      </li>
    `).join('');
}

// 6. Template / Batch Variables
function detectTemplateVariables(text) {
  const regex = /\{\{([^}]+)\}\}/g;
  const found = new Set();
  let match;

  while ((match = regex.exec(text)) !== null) {
    found.add(match[0]);
  }

  variablesTableBody.innerHTML = '';

  if (found.size === 0) {
    addVariableRow('{{EJEMPLO}}', 'Valor sustituto');
    return;
  }

  found.forEach(varTag => {
    addVariableRow(varTag, '');
  });
}

function addVariableRow(key = '', val = '') {
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td>
      <input type="text" class="var-key" value="${escapeHtml(key)}" placeholder="{{VARIABLE}}" />
    </td>
    <td>
      <input type="text" class="var-val" value="${escapeHtml(val)}" placeholder="Nuevo valor..." />
    </td>
    <td style="text-align: center;">
      <button class="btn btn-danger btn-sm btn-delete-var" title="Eliminar fila">✕</button>
    </td>
  `;

  tr.querySelector('.btn-delete-var').addEventListener('click', () => {
    tr.remove();
  });

  variablesTableBody.appendChild(tr);
}

function applyBatchVariables() {
  if (!currentSession) return;

  const rows = variablesTableBody.querySelectorAll('tr');
  const pairs = [];

  rows.forEach(tr => {
    const key = tr.querySelector('.var-key').value.trim();
    const val = tr.querySelector('.var-val').value;
    if (key) {
      pairs.push({ key, value: val });
    }
  });

  if (pairs.length === 0) {
    showToast('No hay variables válidas para reemplazar.', true);
    return;
  }

  try {
    const pairsJson = JSON.stringify(pairs);
    const resultJson = currentSession.batch_replace(pairsJson);
    const result = JSON.parse(resultJson);

    showToast(result.message);
    refreshDocumentView();
  } catch (err) {
    console.error('Batch replace error:', err);
    showToast('Error en reemplazo por lotes: ' + err, true);
  }
}

// 7. Save All Modified Paragraphs in Batch
function saveAllParagraphs() {
  if (!currentSession) return;

  const textareas = paragraphsContainer.querySelectorAll('.paragraph-textarea');
  const updates = [];

  textareas.forEach(ta => {
    const index = parseInt(ta.dataset.index, 10);
    const text = ta.value;
    updates.push({ index, text });
  });

  try {
    const updatesJson = JSON.stringify(updates);
    const count = currentSession.update_paragraphs_batch(updatesJson);
    showToast(`Se guardaron los cambios en ${count} párrafos.`);
    refreshDocumentView();
  } catch (err) {
    showToast('Error al guardar párrafos: ' + err, true);
  }
}

// 8. Download Modified DOCX
function downloadModifiedDocx() {
  if (!currentSession) return;

  try {
    const bytes = currentSession.export_bytes();
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    
    // Format download file name
    const dotIdx = currentFileName.lastIndexOf('.');
    const baseName = dotIdx !== -1 ? currentFileName.substring(0, dotIdx) : currentFileName;
    a.download = `${baseName}_modificado.docx`;

    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast('¡Archivo DOCX descargado exitosamente!');
  } catch (err) {
    console.error('Download error:', err);
    showToast('Error al exportar DOCX: ' + err, true);
  }
}

// 9. Render Zip Internal Structure
function renderZipFiles(files) {
  zipFilesList.innerHTML = '';
  files.forEach(f => {
    const div = document.createElement('div');
    div.className = 'zip-file-card';
    div.innerHTML = `
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"></path>
        <polyline points="13 2 13 9 20 9"></polyline>
      </svg>
      <span>${escapeHtml(f)}</span>
    `;
    zipFilesList.appendChild(div);
  });
}

// 10. Utilities & UI Helpers
function showToast(msg, isError = false) {
  toastMessage.textContent = msg;
  toastBanner.style.background = isError 
    ? 'rgba(239, 68, 68, 0.15)' 
    : 'rgba(16, 185, 129, 0.12)';
  toastBanner.style.borderColor = isError 
    ? 'rgba(239, 68, 68, 0.4)' 
    : 'rgba(16, 185, 129, 0.35)';
  toastBanner.style.color = isError ? '#FCA5A5' : '#A7F3D0';
  toastBanner.style.display = 'flex';

  setTimeout(() => {
    toastBanner.style.display = 'none';
  }, 4500);
}

function formatBytes(bytes) {
  if (!bytes) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
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

// 11. Event Listeners Setup
function setupEventListeners() {
  // File upload inputs
  [fileInput, fileInput2].forEach(input => {
    input.addEventListener('change', (e) => {
      if (e.target.files && e.target.files[0]) {
        loadDocxFile(e.target.files[0]);
      }
    });
  });

  // Sample docx buttons
  [btnSampleDocx, btnSampleDocx2].forEach(btn => {
    btn.addEventListener('click', () => loadSampleDocx());
  });

  // Download DOCX
  btnDownloadDocx.addEventListener('click', downloadModifiedDocx);

  // Drag & Drop
  ['dragenter', 'dragover'].forEach(eventName => {
    dropzoneBox.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropzoneBox.classList.add('drag-over');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropzoneBox.addEventListener(eventName, (e) => {
      e.preventDefault();
      dropzoneBox.classList.remove('drag-over');
    });
  });

  dropzoneBox.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files[0]) {
      loadDocxFile(dt.files[0]);
    }
  });

  // Tab switching
  tabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      tabButtons.forEach(b => b.classList.remove('active'));
      tabPanels.forEach(p => p.classList.remove('active'));

      btn.classList.add('active');
      const targetId = btn.dataset.tab;
      document.getElementById(targetId).classList.add('active');
    });
  });

  // Find & Replace
  btnExecuteReplace.addEventListener('click', handleExecuteReplace);
  findInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') handleExecuteReplace();
  });
  replaceInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') handleExecuteReplace();
  });

  // Paragraph filter & batch save
  paragraphFilterInput.addEventListener('input', () => {
    renderParagraphs(activeParagraphs);
  });
  btnSaveAllParagraphs.addEventListener('click', saveAllParagraphs);

  // Variable rows & detection
  btnAddVariableRow.addEventListener('click', () => addVariableRow());
  btnDetectTemplateVars.addEventListener('click', () => {
    if (currentSession) {
      const fullText = currentSession.get_raw_text();
      detectTemplateVariables(fullText);
      showToast('Variables {{...}} auto-detectadas.');
    }
  });
  btnApplyBatchVariables.addEventListener('click', applyBatchVariables);

  // Copy Full Text
  btnCopyFullText.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(fullTextPreview.value);
      showToast('Texto copiado al portapapeles.');
    } catch {
      fullTextPreview.select();
      document.execCommand('copy');
      showToast('Texto copiado al portapapeles.');
    }
  });
}

// Kick off
initializeWasm();
setupEventListeners();
