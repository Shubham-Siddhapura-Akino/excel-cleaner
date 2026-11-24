/**
 * Cleans string cell values according to strict normalization rules:
 * - Removes double quotes (")
 * - Replaces tabs, newlines, carriage returns, non-breaking spaces (\u00A0),
 *   invisible Unicode whitespace, control characters, and multiple spaces with a single space.
 * - Trims leading and trailing spaces.
 * - Non-string values are returned unmodified.
 *
 * @param {any} text - The input value to clean
 * @param {Object} [options={}] - Configurable cleaning options
 * @returns {any} Cleaned text if string, otherwise original input
 */
export function cleanText(text, options = {}) {
  if (typeof text !== 'string') {
    return text;
  }

  let cleaned = text;

  // 1. Remove double quotes (default: true)
  if (options.removeQuotes !== false) {
    cleaned = cleaned.replace(/"/g, '');
  }

  // 2. Replace tabs, newlines, carriage returns, non-breaking spaces (\u00A0),
  //    invisible Unicode whitespace (\u200B-\u200D, \uFEFF, etc.), directional controls,
  //    C0/C1 control characters, and multiple spaces with a single space.
  if (options.normalizeWhitespace !== false) {
    cleaned = cleaned.replace(
      /[\s\xA0\u1680\u2000-\u200B\u200C\u200D\u200E\u200F\u2028\u2029\u202A-\u202E\u202F\u205F\u3000\uFEFF\u0000-\u001F\u007F-\u009F]+/g,
      ' '
    );
  }

  // 3. Trim leading and trailing spaces (default: true)
  if (options.trimSpaces !== false) {
    cleaned = cleaned.trim();
  }

  return cleaned;
}

export default cleanText;
