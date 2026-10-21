// What a typed text means, as Excel decides it in the language of the sheet: numbers with the
// local separators, percentages, dates and times, booleans; anything else is text. Also the
// reverse (the text shown while editing a cell) and the TSV that Excel and Google Sheets copy
// and paste. No DOM here, so it is testable in Node.

/**
 * Separators and words of a language. Like the engine's formatting (`numfmt.rs`), Spanish
 * (`es`, `es-MX`…) has its own conventions and every other language uses English ones.
 * @param {string} locale
 */
export function localeInfo(locale = 'en') {
  return String(locale).trim().toLowerCase().startsWith('es')
    ? { decimal: ',', group: '.', trueText: 'VERDADERO', falseText: 'FALSO', dayFirst: true }
    : { decimal: '.', group: ',', trueText: 'TRUE', falseText: 'FALSE', dayFirst: false };
}

const escape = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const CURRENCY = '[$€£¥]';
const MS_PER_DAY = 86_400_000;

/** A number written with the separators of `info`, or null */
function parseNumber(text, info) {
  let s = text.trim();
  let negative = false;
  // (123) is a negative number in accounting formats
  const paren = /^\((.*)\)$/.exec(s);
  if (paren) {
    negative = true;
    s = paren[1].trim();
  }
  const sign = re => {
    const m = re.exec(s);
    if (!m) return;
    if (m[1] === '-') negative = !negative;
    s = s.slice(m[0].length).trim();
  };
  sign(/^([+-])/);
  // A currency symbol before or after the number (`$1,234`, `1.234 €`), the sign on either side
  const before = new RegExp(`^${CURRENCY}`).exec(s);
  if (before) {
    s = s.slice(1).trim();
    sign(/^([+-])/);
  } else {
    const after = new RegExp(`\\s*${CURRENCY}$`).exec(s);
    if (after) s = s.slice(0, after.index);
  }
  let percent = false;
  const pct = /\s*%$/.exec(s);
  if (pct) {
    percent = true;
    s = s.slice(0, pct.index);
  }
  const g = escape(info.group);
  const d = escape(info.decimal);
  const m = new RegExp(`^(\\d{1,3}(?:${g}\\d{3})+|\\d+)?(?:${d}(\\d*))?(?:[eE]([+-]?\\d+))?$`).exec(s);
  if (!m || (m[1] === undefined && !m[2])) return null;
  const digits = `${(m[1] ?? '0').split(info.group).join('')}.${m[2] || '0'}${m[3] ? `e${m[3]}` : ''}`;
  let value = Number(digits);
  if (!Number.isFinite(value)) return null;
  if (negative) value = -value;
  if (percent) value = Number((value / 100).toPrecision(15));
  return { value, percent };
}

/** `14:30`, `14:30:15`, `2:30 PM` → milliseconds since midnight, or null */
function parseTime(text) {
  const m = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2})(?:[.,](\d{1,3}))?)?\s*(a\.?\s?m\.?|p\.?\s?m\.?)?$/i.exec(text.trim());
  if (!m) return null;
  let hours = Number(m[1]);
  const minutes = Number(m[2]);
  const seconds = Number(m[3] ?? 0);
  const ms = m[4] ? Number(m[4].padEnd(3, '0')) : 0;
  if (m[5]) {
    if (hours < 1 || hours > 12) return null;
    const pm = m[5].toLowerCase().startsWith('p');
    hours = (hours % 12) + (pm ? 12 : 0);
  }
  if (hours > 23 || minutes > 59 || seconds > 59) return null;
  return ((hours * 60 + minutes) * 60 + seconds) * 1000 + ms;
}

/** Two-digit years as Excel reads them: 00–29 → 2000s, 30–99 → 1900s */
const fullYear = y => (y.length <= 2 ? (Number(y) < 30 ? 2000 : 1900) + Number(y) : Number(y));

function makeDate(year, month, day) {
  if (month < 1 || month > 12 || day < 1) return null;
  const ms = Date.UTC(year, month - 1, day);
  const back = new Date(ms);
  if (back.getUTCDate() !== day || back.getUTCMonth() !== month - 1) return null; // 31/02
  return year < 1900 || year > 9999 ? null : ms;
}

/** A date (with an optional time) → milliseconds since 1970-01-01 UTC, or null */
function parseDate(text, info, today) {
  const s = text.trim();
  const [datePart, ...rest] = s.split(/\s+|T(?=\d)/);
  const timeText = rest.join(' ');
  let time = 0;
  if (timeText) {
    time = parseTime(timeText);
    if (time === null) return null;
  }
  let ms = null;
  let m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(datePart);
  if (m) {
    ms = makeDate(Number(m[1]), Number(m[2]), Number(m[3]));
  } else if ((m = /^(\d{1,2})([-/.])(\d{1,2})(?:\2(\d{2}|\d{4}))?$/.exec(datePart))) {
    // `15.3.2024` is a date in Spanish; in English the dot is the decimal point
    if (m[2] === '.' && (!info.dayFirst || !m[4])) return null;
    const [day, month] = info.dayFirst ? [m[1], m[3]] : [m[3], m[1]];
    ms = makeDate(m[4] ? fullYear(m[4]) : today.getFullYear(), Number(month), Number(day));
  }
  return ms === null ? null : ms + time;
}

/**
 * What a typed text means.
 * - `empty`: nothing (clears the cell); value `null`
 * - `number`, `percent` (value already divided by 100), `boolean`
 * - `date`: a `Date` holding the wall-clock time as UTC, as `@dimarborda/xlsx` writes it
 * - `time`: a time without date, as the fraction of a day Excel stores
 * - `formula`: starts with `=`; kept as text, since formulas are not calculated yet
 * - `string`: anything else; a leading `'` forces text and is dropped, as in Excel
 * @param {string} text
 * @param {string} [locale]
 * @param {{today?: Date}} [options] the year of dates typed without one (default: now)
 * @returns {{type: 'empty' | 'string' | 'number' | 'percent' | 'boolean' | 'date' | 'time' | 'formula', value: null | string | number | boolean | Date}}
 */
export function parseInput(text, locale = 'en', { today = new Date() } = {}) {
  const raw = String(text ?? '');
  if (raw.startsWith("'")) return { type: 'string', value: raw.slice(1) };
  const s = raw.trim();
  if (!s) return { type: 'empty', value: null };
  if (s.startsWith('=') && s.length > 1) return { type: 'formula', value: raw };
  const info = localeInfo(locale);
  const upper = s.toUpperCase();
  if (upper === info.trueText) return { type: 'boolean', value: true };
  if (upper === info.falseText) return { type: 'boolean', value: false };
  const number = parseNumber(s, info);
  if (number) return { type: number.percent ? 'percent' : 'number', value: number.value };
  const time = parseTime(s);
  if (time !== null) return { type: 'time', value: time / MS_PER_DAY };
  const date = parseDate(s, info, today);
  if (date !== null) return { type: 'date', value: new Date(date) };
  return { type: 'string', value: raw };
}

/** A number as Excel shows it while editing: no grouping, up to 15 significant digits */
export function numberText(n, locale = 'en') {
  const s = String(Number(n.toPrecision(15))).replace('e', 'E');
  return s.replace('.', localeInfo(locale).decimal);
}

const pad = n => String(n).padStart(2, '0');

function timeText(d) {
  const ms = d.getUTCMilliseconds();
  const base = `${d.getUTCHours()}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;
  return ms ? `${base}.${String(ms).padStart(3, '0').replace(/0+$/, '')}` : base;
}

/**
 * The text to edit for a cell (`sheet.getCell()`), which typed back gives the same value
 * @param {{type: string, value: any, formula?: string}} cell
 * @param {string} [locale]
 * @param {string} [shown] the text the cell shows, to keep percentages as percentages
 */
export function editText(cell, locale = 'en', shown = '') {
  if (cell.formula) return `=${cell.formula}`;
  const info = localeInfo(locale);
  switch (cell.type) {
    case 'empty':
      return '';
    case 'number':
      if (/%\s*$/.test(shown)) return `${numberText(cell.value * 100, locale)}%`;
      return numberText(cell.value, locale);
    case 'boolean':
      return cell.value ? info.trueText : info.falseText;
    case 'date': {
      const d = cell.value instanceof Date ? cell.value : new Date(cell.value);
      // Times without a date sit on Excel's day 0, before 1900
      if (d.getUTCFullYear() < 1900) return timeText(d);
      const [y, m, day] = [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()];
      const date = info.dayFirst ? `${pad(day)}/${pad(m)}/${y}` : `${m}/${day}/${y}`;
      return d.getTime() % MS_PER_DAY ? `${date} ${timeText(d)}` : date;
    }
    case 'string': {
      const text = String(cell.value);
      // Text that would be read as something else keeps it as text with Excel's apostrophe
      const type = parseInput(text, locale).type;
      return type === 'string' || type === 'formula' || type === 'empty' ? text : `'${text}`;
    }
    default:
      return String(cell.value ?? '');
  }
}

/** Rows of texts as TSV, quoting the ones with tabs, line breaks or quotes as Excel does */
export function toTSV(rows) {
  return rows
    .map(row => row.map(t => (/[\t\n\r"]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t)).join('\t'))
    .join('\r\n');
}

/** TSV (from Excel, Google Sheets or `toTSV`) to rows of texts */
export function parseTSV(text) {
  const rows = [];
  let row = [];
  let i = 0;
  const s = String(text).replace(/\r\n?/g, '\n');
  while (i <= s.length) {
    let field = '';
    if (s[i] === '"') {
      // Quoted field: "" is a quote; it ends at the quote before a tab, a line break or the end
      let j = i + 1;
      let closed = false;
      while (j < s.length) {
        if (s[j] === '"') {
          if (s[j + 1] === '"') {
            field += '"';
            j += 2;
            continue;
          }
          if (j + 1 === s.length || s[j + 1] === '\t' || s[j + 1] === '\n') {
            closed = true;
            j += 1;
            break;
          }
        }
        field += s[j];
        j += 1;
      }
      if (closed) {
        i = j;
      } else {
        // Not a quoted field after all: take it as written
        const end = s.slice(i).search(/[\t\n]/);
        field = end < 0 ? s.slice(i) : s.slice(i, i + end);
        i = end < 0 ? s.length : i + end;
      }
    } else {
      const end = s.slice(i).search(/[\t\n]/);
      field = end < 0 ? s.slice(i) : s.slice(i, i + end);
      i = end < 0 ? s.length : i + end;
    }
    row.push(field);
    if (i >= s.length) {
      rows.push(row);
      break;
    }
    if (s[i] === '\n') {
      rows.push(row);
      row = [];
      // The line break that ends the last row is not one more row
      if (i + 1 === s.length) break;
    }
    i += 1;
  }
  return rows;
}
