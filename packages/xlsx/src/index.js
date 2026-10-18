// Excel (.xlsx) workbooks backed by the Rust/WebAssembly engine. In Node the engine is read
// from the package itself; in the browser bundlers serve the .wasm next to the JS, or pass its
// location with `initEngine(url)`.
import init, { initSync, XlsxWorkbook } from '../wasm/xlsx_engine.js';

let loading = null;

const isNode = typeof process !== 'undefined' && !!process.versions?.node;

/**
 * Starts the engine. Safe to call many times: later calls wait for the first one.
 * @param {string | URL | BufferSource} [wasm] where `xlsx_engine_bg.wasm` is served from, or its bytes
 */
export function initEngine(wasm) {
  loading ??= (async () => {
    if (wasm && (ArrayBuffer.isView(wasm) || wasm instanceof ArrayBuffer)) {
      initSync({ module: wasm });
    } else if (!wasm && isNode) {
      const { readFile } = await import('node:fs/promises');
      initSync({ module: await readFile(new URL('../wasm/xlsx_engine_bg.wasm', import.meta.url)) });
    } else {
      await init(wasm ? { module_or_path: wasm } : undefined);
    }
  })().catch(err => {
    loading = null; // a later call may retry (e.g. after fixing the location)
    throw err;
  });
  return loading;
}

/** An open workbook */
export class Workbook {
  /** @param {XlsxWorkbook} engine */
  constructor(engine) {
    this._engine = engine;
  }

  /** Sheet names, in tab order */
  get sheetNames() {
    return this._engine.sheet_names();
  }

  /** The workbook as .xlsx bytes */
  toBytes() {
    return this._engine.export_bytes();
  }

  /** Frees the engine memory of this workbook; it cannot be used afterwards */
  close() {
    this._engine.free();
  }
}

/**
 * Opens an .xlsx
 * @param {Uint8Array | ArrayBuffer} bytes
 */
export async function openWorkbook(bytes) {
  await initEngine();
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  return new Workbook(new XlsxWorkbook(data));
}

/** A new workbook with one empty sheet */
export async function createWorkbook() {
  await initEngine();
  return new Workbook(XlsxWorkbook.new_blank());
}
