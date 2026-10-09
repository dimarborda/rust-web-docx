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
  loading ??= init(wasmUrl ? { module_or_path: wasmUrl } : undefined).then(
    () => {
      ready = true;
    },
    err => {
      loading = null; // a later call may retry (e.g. after fixing wasmUrl)
      throw explainEngineError(err);
    },
  );
  return loading;
}

/** The engine failed to start: say why in plain words when the cause is a known setup issue */
function explainEngineError(err) {
  const text = String(err?.message || err);
  if (/Content Security Policy|unsafe-eval|wasm-unsafe-eval/i.test(text)) {
    return new Error(
      "docx-editor: the page's Content-Security-Policy blocks WebAssembly. Add 'wasm-unsafe-eval' "
        + "to script-src (see the README, section \"Content Security Policy\").",
      { cause: err },
    );
  }
  if (/magic word|expected magic|Failed to fetch|NetworkError|404/i.test(text)) {
    return new Error(
      'docx-editor: docx_engine_bg.wasm could not be loaded. Check that it is served next to the '
        + 'package\'s JavaScript, or pass its location with `wasmUrl` (see the README, section "Bundlers").',
      { cause: err },
    );
  }
  return err;
}

export const engineReady = () => ready;

/** Bytes of the built-in demo contract, handy for tests and playgrounds */
export async function sampleDocx() {
  await initEngine();
  return generate_sample_docx_wasm();
}

export { DocxSession };
