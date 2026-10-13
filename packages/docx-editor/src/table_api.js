// Argument checks for the existing-table API of DocxEditor (setTableCell, insertTableRows,
// deleteTableRows). The engine takes unsigned integers, so negative or fractional values are
// refused here with a clear message instead of wrapping around.

/** Throws unless every value of `args` is a non-negative integer */
export function checkIndex(method, args) {
  for (const [name, value] of Object.entries(args)) {
    if (!Number.isInteger(value) || value < 0) {
      throw new Error(`docx-editor: ${method}: \`${name}\` must be an integer ≥ 0 (got ${String(value)})`);
    }
  }
}

/** Rows for `insertTableRows` as arrays of strings (null and undefined become empty cells) */
export function tableRowsPayload(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    throw new Error('docx-editor: insertTableRows needs at least one row, e.g. [["A", "1"]]');
  }
  return rows.map((row, k) => {
    if (!Array.isArray(row)) {
      throw new Error(`docx-editor: insertTableRows: row ${k} must be an array of cell texts`);
    }
    return row.map(cell => (cell == null ? '' : String(cell).replace(/\r\n?/g, '\n')));
  });
}

/** Runs an engine call, turning the engine's error strings into `Error`s */
export function engineCall(fn) {
  try {
    return fn();
  } catch (err) {
    throw err instanceof Error ? err : new Error(`docx-editor: ${String(err)}`, { cause: err });
  }
}
