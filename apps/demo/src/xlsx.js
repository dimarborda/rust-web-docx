// Demo of <xlsx-editor>: opens the sample workbook, files picked or dropped on the page, zoom,
// read-only mode, undo and download.
import '@dimarborda/xlsx-editor';
import '@dimarborda/xlsx-editor/fonts';
import '@dimarborda/xlsx-editor/style.css';

const element = document.getElementById('sheet');
const status = document.getElementById('status');
let fileName = 'ejemplo-ventas.xlsx';

const show = text => {
  status.textContent = text;
};

element.addEventListener('load', e => show(`${e.detail.sheets.length} hoja(s)`));
element.addEventListener('selectionchange', e => show(`${e.detail.range}${e.detail.text ? ` · ${e.detail.text}` : ''}`));
element.addEventListener('change', e => show(`${e.detail.range} cambiada (${e.detail.source})`));
element.addEventListener('error', e => show(e.detail.error.message));

async function openFile(file) {
  if (!file) return;
  fileName = file.name;
  try {
    await element.editor.open(file);
  } catch (error) {
    show(`No se pudo abrir: ${error.message}`);
  }
}

document.getElementById('file').addEventListener('change', e => openFile(e.target.files[0]));
document.getElementById('zoom').addEventListener('change', e => element.editor?.setZoom(Number(e.target.value)));
document.getElementById('readonly').addEventListener('change', e => element.toggleAttribute('readonly', e.target.checked));
document.getElementById('undo').addEventListener('click', () => element.editor?.undo());
document.getElementById('redo').addEventListener('click', () => element.editor?.redo());
document.getElementById('save').addEventListener('click', () => element.editor?.download(fileName));

document.addEventListener('dragover', e => {
  e.preventDefault();
  document.body.classList.add('drop');
});
document.addEventListener('dragleave', e => {
  if (!e.relatedTarget) document.body.classList.remove('drop');
});
document.addEventListener('drop', e => {
  e.preventDefault();
  document.body.classList.remove('drop');
  openFile(e.dataTransfer.files[0]);
});
