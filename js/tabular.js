/*
 * CSV / TSV and JSON import and export.
 *
 *   Tabular.parseCsv(text, delimiter?)  -> rows
 *   Tabular.toCsv(rows, delimiter?)     -> string
 *   Tabular.decodeText(uint8array)      -> string   (UTF-8 / UTF-16, falls back to windows-1256 for old Arabic CSVs)
 *   Tabular.parseJson(text)             -> [{ name, rows }]
 *   Tabular.toJson(rows, header)        -> string
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./values.js'));
  else root.Tabular = factory(root.Values);
})(typeof self !== 'undefined' ? self : this, function (Values) {
  'use strict';

  function decodeText(bytes) {
    if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) return new TextDecoder('utf-8').decode(bytes.subarray(3));
    if (bytes[0] === 0xFF && bytes[1] === 0xFE) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
    if (bytes[0] === 0xFE && bytes[1] === 0xFF) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    } catch (e) {
      return new TextDecoder('windows-1256').decode(bytes); // Excel's "CSV" on Arabic Windows
    }
  }

  function detectDelimiter(text) {
    const firstLine = (text.replace(/"(?:[^"]|"")*"/g, '').split(/\r\n|\n|\r/).find((l) => l.trim() !== '')) || '';
    let best = ',';
    let bestCount = 0;
    for (const d of [',', ';', '\t', '|']) {
      const count = firstLine.split(d).length - 1;
      if (count > bestCount) { best = d; bestCount = count; }
    }
    return best;
  }

  function parseCsv(text, delimiter) {
    delimiter = delimiter || detectDelimiter(text);
    const table = [];
    let row = [];
    let field = '';
    let quoted = false;

    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (quoted) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; } else quoted = false;
        } else field += ch;
      } else if (ch === '"' && field === '') {
        quoted = true;
      } else if (ch === delimiter) {
        row.push(field); field = '';
      } else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(field); field = '';
        table.push(row); row = [];
      } else {
        field += ch;
      }
    }
    if (field !== '' || row.length) { row.push(field); table.push(row); }

    while (table.length && table[table.length - 1].every((c) => c === '')) table.pop();
    const width = table.reduce((max, r) => Math.max(max, r.length), 0);
    return table.map((r) => {
      const out = new Array(width).fill(null);
      r.forEach((cell, c) => { out[c] = Values.parse(cell); });
      return out;
    });
  }

  function toCsv(rows, delimiter) {
    delimiter = delimiter || ',';
    return rows.map((row) => row.map((value) => {
      const s = Values.format(value);
      return s.includes(delimiter) || /["\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(delimiter)).join('\r\n');
  }

  // ------------------------------------------------------------------- JSON

  function jsonCell(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === 'object') return JSON.stringify(v);
    return v;
  }

  function rowsFromJson(data) {
    if (!Array.isArray(data)) throw new Error('The JSON file must contain an array (of objects or of arrays).');
    if (data.every((item) => Array.isArray(item))) return data.map((r) => r.map(jsonCell));

    const keys = [];
    data.forEach((item) => {
      if (item === null || typeof item !== 'object' || Array.isArray(item)) {
        throw new Error('Every item in the JSON array must be an object, or every item must be an array.');
      }
      Object.keys(item).forEach((k) => { if (!keys.includes(k)) keys.push(k); });
    });
    return [keys.slice()].concat(data.map((item) => keys.map((k) => jsonCell(item[k]))));
  }

  function parseJson(text) {
    let data;
    try { data = JSON.parse(text); } catch (e) { throw new Error('Invalid JSON: ' + e.message); }
    if (Array.isArray(data)) return [{ name: 'Sheet1', rows: rowsFromJson(data) }];
    if (data && typeof data === 'object') {
      return Object.keys(data).map((name) => ({ name, rows: rowsFromJson(data[name]) }));
    }
    throw new Error('The JSON file must contain an array or an object of arrays.');
  }

  function toJson(rows, header) {
    const plain = (v) => (v instanceof Date ? Values.formatDate(v) : v === undefined ? null : v);
    if (header === false || rows.length === 0) return JSON.stringify(rows.map((r) => r.map(plain)), null, 2);

    const used = {};
    const keys = rows[0].map((h, i) => {
      const base = Values.format(h).trim() || 'column_' + (i + 1);
      used[base] = (used[base] || 0) + 1;
      return used[base] > 1 ? base + '_' + used[base] : base;
    });
    return JSON.stringify(rows.slice(1).map((r) => {
      const obj = {};
      keys.forEach((k, i) => { obj[k] = plain(r[i] === undefined ? null : r[i]); });
      return obj;
    }), null, 2);
  }

  return { decodeText, parseCsv, toCsv, parseJson, toJson, detectDelimiter };
});
