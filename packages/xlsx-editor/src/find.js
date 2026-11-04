// Searching cell texts as Excel does: `*` any text, `?` one character, `~` takes either literally.
// No DOM here, so it is testable in Node.

/**
 * The search as a global regular expression; null for an empty search
 * @param {string} query
 * @param {{matchCase?: boolean, wholeCell?: boolean}} [options]
 */
export function searchPattern(query, { matchCase = false, wholeCell = false } = {}) {
  if (!query) return null;
  // With the `u` flag only characters with a meaning may be escaped
  const literal = c => c.replace(/[.*+?^${}()|[\]\\/]/gu, '\\$&');
  let source = '';
  const chars = [...query];
  for (let i = 0; i < chars.length; i++) {
    const c = chars[i];
    if (c === '~' && '*?~'.includes(chars[i + 1] ?? ' ')) source += literal(chars[++i]);
    else if (c === '*') source += '[\\s\\S]*';
    else if (c === '?') source += '[\\s\\S]';
    else source += literal(c);
  }
  return new RegExp(wholeCell ? `^(?:${source})$` : source, matchCase ? 'gu' : 'giu');
}

/** Whether the pattern is in a text (a global expression remembers where it stopped: not here) */
export function matches(pattern, text) {
  pattern.lastIndex = 0;
  const found = pattern.test(text);
  pattern.lastIndex = 0;
  return found;
}

/** The text with what the pattern finds replaced, taken literally */
export function replaceMatches(pattern, text, replacement) {
  pattern.lastIndex = 0;
  return text.replace(pattern, found => (found === '' ? '' : replacement));
}
