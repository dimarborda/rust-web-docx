// Fonts the user supplies (for example the real Calibri .ttf) so documents render with the
// exact metrics Word uses. Fonts are registered with the page's FontFace set, shared by every
// editor on the page, and can be kept in IndexedDB across visits.

const DB_NAME = 'rust_docx_fonts_db';
const DB_VERSION = 1;
const STORE = 'custom_fonts';

/** family (lower case) → stored record */
const customFonts = new Map();
const listeners = new Set();

export const hasCustomFont = family => customFonts.has((family || '').toLowerCase().trim());

/** Called whenever fonts are added or removed, so editors can lay out again */
export function onFontsChanged(callback) {
  listeners.add(callback);
  return () => listeners.delete(callback);
}

const notify = () => listeners.forEach(cb => cb());

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      if (!req.result.objectStoreNames.contains(STORE)) req.result.createObjectStore(STORE, { keyPath: 'id' });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function withStore(mode, run) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, mode);
    const result = run(tx.objectStore(STORE));
    tx.oncomplete = () => resolve(result.result ?? result);
    tx.onerror = () => reject(tx.error);
  });
}

/** Fonts saved in this browser (IndexedDB), without registering them */
export async function listStoredFonts() {
  try {
    return (await withStore('readonly', store => store.getAll())) || [];
  } catch (err) {
    console.warn('docx-editor: could not read stored fonts', err);
    return [];
  }
}

/** Registers every font saved in this browser. Call once at startup. */
export async function loadStoredFonts() {
  const stored = await listStoredFonts();
  for (const record of stored) {
    try {
      await addFace(record);
    } catch (err) {
      console.warn('docx-editor: could not register stored font', record.family, err);
    }
  }
  if (stored.length) notify();
  return stored;
}

async function addFace(record) {
  const face = new FontFace(record.family, record.buffer, { weight: record.weight, style: record.style });
  await face.load();
  document.fonts.add(face);
  customFonts.set(record.family.toLowerCase(), record);
}

/**
 * Registers a .ttf/.otf/.woff/.woff2 file. The family comes from the font's own name table
 * (falling back to the file name), the weight and style from the file name.
 * @param {File | {name: string, buffer: ArrayBuffer}} file
 * @param {{persist?: boolean}} [options] keep it in IndexedDB (default true)
 */
export async function registerFont(file, { persist = true } = {}) {
  const buffer = file.buffer instanceof ArrayBuffer ? file.buffer : await file.arrayBuffer();
  const parsed = parseFontFileName(file.name);
  const family = fontFamilyName(buffer) || parsed.family;
  const record = {
    id: `${family.toLowerCase()}_${parsed.weight}_${parsed.style}_${buffer.byteLength}`,
    family,
    fileName: file.name,
    size: buffer.byteLength,
    weight: parsed.weight,
    style: parsed.style,
    buffer,
    addedAt: Date.now(),
  };
  await addFace(record);
  if (persist) {
    try {
      await withStore('readwrite', store => store.put(record));
    } catch (err) {
      console.warn('docx-editor: could not store font', err);
    }
  }
  notify();
  return record;
}

/** Forgets a stored font. It stays usable until the page reloads (FontFace cannot be removed reliably). */
export async function removeStoredFont(id) {
  try {
    await withStore('readwrite', store => store.delete(id));
  } catch (err) {
    console.warn('docx-editor: could not delete stored font', err);
  }
  for (const [key, record] of customFonts) if (record.id === id) customFonts.delete(key);
  notify();
}

function parseFontFileName(name) {
  const clean = name.replace(/\.(ttf|otf|woff2?)$/i, '');
  const lower = clean.toLowerCase();
  const weight = /bold|bd$|bd[-_ ]|b$/.test(lower) ? 'bold' : 'normal';
  const style = /italic|oblique|it$|i$/.test(lower) ? 'italic' : 'normal';
  const known = ['Calibri', 'Aptos', 'Arial', 'Cambria', 'Times New Roman', 'Consolas'];
  const family = known.find(k => lower.includes(k.toLowerCase().split(' ')[0]))
    || clean.replace(/[-_](regular|bold|italic|bolditalic|light|semibold|medium|bd|it|z|b|i)$/i, '').trim();
  return { family, weight, style };
}

/** Family name (name IDs 1/4) from a TrueType/OpenType name table */
function fontFamilyName(buffer) {
  try {
    const view = new DataView(buffer);
    const numTables = view.getUint16(4);
    for (let i = 0; i < numTables; i++) {
      const rec = 12 + i * 16;
      const tag = String.fromCharCode(...[0, 1, 2, 3].map(k => view.getUint8(rec + k)));
      if (tag !== 'name') continue;
      const offset = view.getUint32(rec + 8);
      const count = view.getUint16(offset + 2);
      const strings = offset + view.getUint16(offset + 4);
      for (let j = 0; j < count; j++) {
        const r = offset + 6 + j * 12;
        const platform = view.getUint16(r);
        const nameId = view.getUint16(r + 6);
        const length = view.getUint16(r + 8);
        const start = strings + view.getUint16(r + 10);
        if (nameId !== 1 && nameId !== 4) continue;
        let name = '';
        if (platform === 0 || platform === 3) {
          for (let k = 0; k < length; k += 2) name += String.fromCharCode(view.getUint16(start + k));
        } else {
          for (let k = 0; k < length; k++) name += String.fromCharCode(view.getUint8(start + k));
        }
        if (name.trim()) return name.trim();
      }
    }
  } catch {
    // Not a TrueType/OpenType file we can read: the file name decides
  }
  return null;
}
