// Loads the Rust/WebAssembly engine once per page. Bundlers (Vite, webpack 5, Rollup) emit
// the .wasm file next to the JS automatically; pass `wasmUrl` to serve it from elsewhere.
import init, { DocxSession, generate_sample_docx_wasm } from '../wasm/docx_engine.js';

let loading = null;
let ready = false;

/**
 * Initializes the engine. Safe to call many times: later calls wait for the first one.
 * @param {string | URL} [wasmUrl] where `docx_engine_bg.wasm` is served from
 */
export function initEngine(wasmUrl) {
  loading ??= init(wasmUrl ? { module_or_path: wasmUrl } : undefined).then(() => {
    ready = true;
  });
  return loading;
}

export const engineReady = () => ready;

/** Bytes of the built-in demo contract, handy for tests and playgrounds */
export async function sampleDocx() {
  await initEngine();
  return generate_sample_docx_wasm();
}

export { DocxSession };
