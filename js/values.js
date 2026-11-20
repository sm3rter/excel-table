/*
 * Cell value helpers shared by the file readers/writers and the UI.
 * A cell value is one of: null (empty), string, number, boolean, or a JS Date (UTC based).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Values = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const NUMBER = /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/;
  const DATE_TEXT = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/;

  function formatDate(date) {
    if (isNaN(date.getTime())) return '';
    const iso = date.toISOString();
    const hasTime = date.getTime() % 86400000 !== 0;
    if (!hasTime) return iso.slice(0, 10);
    return iso.slice(0, 10) + ' ' + iso.slice(11, date.getUTCSeconds() ? 19 : 16);
  }

  /** Text shown in the grid, the formula bar and CSV files. */
  function format(value) {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) return formatDate(value);
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    return String(value);
  }

  /**
   * Turn text typed by a person (or read from CSV) into a typed value.
   * Numbers become numbers, except long digit strings and values with leading zeros ("00123", phone numbers).
   * Booleans and dates are only restored when the previous value of the cell had that type.
   */
  function parse(text, previous) {
    if (text === '' || text === null || text === undefined) return null;
    const s = String(text);
    if (NUMBER.test(s) && s.length <= 15) return Number(s);
    if (typeof previous === 'boolean' && /^(true|false)$/i.test(s)) return s.toLowerCase() === 'true';
    if (previous instanceof Date) {
      const m = DATE_TEXT.exec(s);
      if (m) {
        const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)));
        if (!isNaN(d.getTime())) return d;
      }
    }
    return s;
  }

  /** Sort comparison: empty last, numbers/dates by value, text with natural (numeric) ordering. */
  function compare(a, b) {
    const aEmpty = a === null || a === undefined || a === '';
    const bEmpty = b === null || b === undefined || b === '';
    if (aEmpty || bEmpty) return aEmpty === bEmpty ? 0 : aEmpty ? 1 : -1;
    const an = a instanceof Date ? a.getTime() : typeof a === 'number' ? a : null;
    const bn = b instanceof Date ? b.getTime() : typeof b === 'number' ? b : null;
    if (an !== null && bn !== null) return an - bn;
    if (an !== null) return -1;
    if (bn !== null) return 1;
    return format(a).localeCompare(format(b), undefined, { numeric: true, sensitivity: 'base' });
  }

  /** "A", "B", ... "Z", "AA" for a zero-based column index. */
  function columnName(index) {
    let name = '';
    let n = index + 1;
    while (n > 0) {
      const m = (n - 1) % 26;
      name = String.fromCharCode(65 + m) + name;
      n = Math.floor((n - 1) / 26);
    }
    return name;
  }

  const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
    45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);

  /** Does an Excel number format (built-in id, or custom code looked up in customCodes) show a date or time? */
  function isDateFormat(id, customCodes) {
    if (BUILTIN_DATE_FORMATS.has(id)) return true;
    const code = customCodes && customCodes[id];
    if (!code) return false;
    const cleaned = code.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, '').replace(/_./g, '');
    return /[ymdhs]/i.test(cleaned);
  }

  /** Excel serial number (1900 or 1904 date system) to a UTC Date, rounded to the second. */
  function serialToDate(serial, date1904) {
    const days = date1904 ? serial + 1462 : serial;
    return new Date(Math.round(((days - 25569) * 86400000) / 1000) * 1000);
  }

  /** A valid, unique Excel sheet name (max 31 characters, none of []:*?/\\). `used` is a Set of lower-cased names. */
  function sheetName(raw, used) {
    const base = String(raw || 'Sheet').replace(/[\[\]:*?\/\\]/g, ' ').replace(/^'+|'+$/g, '').trim().slice(0, 31) || 'Sheet';
    let name = base;
    let n = 2;
    while (used.has(name.toLowerCase())) {
      const suffix = ' (' + n++ + ')';
      name = base.slice(0, 31 - suffix.length) + suffix;
    }
    used.add(name.toLowerCase());
    return name;
  }

  return { format, formatDate, parse, compare, columnName, isDateFormat, serialToDate, sheetName };
});
