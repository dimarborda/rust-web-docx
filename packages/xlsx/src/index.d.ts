/** Starts the engine; `wasm` is where `xlsx_engine_bg.wasm` is served from, or its bytes */
export function initEngine(wasm?: string | URL | BufferSource): Promise<void>;

/** An open workbook */
export class Workbook {
  private constructor(engine: unknown);
  /** Sheet names, in tab order */
  readonly sheetNames: string[];
  /** The workbook as .xlsx bytes */
  toBytes(): Uint8Array;
  /** Frees the engine memory of this workbook; it cannot be used afterwards */
  close(): void;
}

/** Opens an .xlsx */
export function openWorkbook(bytes: Uint8Array | ArrayBuffer): Promise<Workbook>;

/** A new workbook with one empty sheet */
export function createWorkbook(): Promise<Workbook>;
