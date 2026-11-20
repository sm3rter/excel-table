/*
 * .xlsx reader and writer without dependencies.
 *
 *   XlsxLite.read(arrayBuffer)  -> Promise<{ sheets: [{ name, rows }] }>
 *   XlsxLite.write(sheets, opt) -> Promise<Uint8Array>      opt: { rtl, header }
 *
 * `rows` is an array of arrays of cell values (null | string | number | boolean | Date), see values.js.
 * Supported: shared/inline strings, numbers, booleans, dates (by number format), cached formula results,
 * multiple sheets, 1900/1904 date systems. Not supported: legacy .xls, formatting, merged cells, charts.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./zip.js'), require('./values.js'));
  else root.XlsxLite = factory(root.Zip, root.Values);
})(typeof self !== 'undefined' ? self : this, function (Zip, Values) {
  'use strict';

  // ------------------------------------------------------------ xml helpers

  function decodeEntities(s) {
    return s.replace(/&(#x[0-9a-fA-F]+|#\d+|lt|gt|amp|quot|apos);/g, function (m, e) {
      switch (e) {
        case 'lt': return '<';
        case 'gt': return '>';
        case 'amp': return '&';
        case 'quot': return '"';
        case 'apos': return "'";
      }
      const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      try { return String.fromCodePoint(code); } catch (err) { return ''; }
    });
  }

  /** Text as stored by Excel: XML entities plus _x000D_ style escapes. */
  function decodeText(s) {
    return decodeEntities(s).replace(/_x([0-9A-Fa-f]{4})_/g, (m, hex) => String.fromCharCode(parseInt(hex, 16)));
  }

  function escapeXml(s) {
    return String(s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function attrs(source) {
    const out = {};
    source.replace(/([\w:.\-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g, (m, key, a, b) => {
      out[key] = a !== undefined ? a : b;
      return m;
    });
    return out;
  }

  /** Some generators prefix every tag (<x:worksheet>). Remove that prefix so one set of regexes works. */
  function stripPrefix(xml, rootTag) {
    const m = new RegExp('<(\\w+):' + rootTag + '\\b').exec(xml);
    return m ? xml.replace(new RegExp('(</?)' + m[1] + ':', 'g'), '$1') : xml;
  }

  function textOf(xml) {
    xml = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    let out = '';
    const re = /<t\b[^>]*?(?:\/>|>([\s\S]*?)<\/t>)/g;
    let m;
    while ((m = re.exec(xml))) out += m[1] !== undefined ? decodeText(m[1]) : '';
    return out;
  }

  function columnIndex(letters) {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }

  function resolveTarget(target) {
    const parts = (target.startsWith('/') ? target.slice(1) : 'xl/' + target).split('/');
    const stack = [];
    for (const part of parts) {
      if (part === '..') stack.pop();
      else if (part !== '.' && part !== '') stack.push(part);
    }
    return stack.join('/');
  }

  // ------------------------------------------------------------------- read

  function parseStyles(xml) {
    const dateXf = [];
    if (!xml) return dateXf;
    xml = stripPrefix(xml, 'styleSheet');
    const codes = {};
    const numFmtRe = /<numFmt\b([^>]*?)\/?>/g;
    let m;
    while ((m = numFmtRe.exec(xml))) {
      const a = attrs(m[1]);
      if (a.numFmtId !== undefined) codes[a.numFmtId] = decodeEntities(a.formatCode || '');
    }
    const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
    if (xfs) {
      const xfRe = /<xf\b([^>]*?)\/?>/g;
      while ((m = xfRe.exec(xfs[1]))) {
        dateXf.push(Values.isDateFormat(parseInt(attrs(m[1]).numFmtId || '0', 10), codes));
      }
    }
    return dateXf;
  }

  function parseSharedStrings(xml) {
    const out = [];
    if (!xml) return out;
    xml = stripPrefix(xml, 'sst');
    const re = /<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g;
    let m;
    while ((m = re.exec(xml))) out.push(m[1] ? textOf(m[1]) : '');
    return out;
  }

  function readCell(a, inner, ctx) {
    if (a.t === 'inlineStr') {
      const is = /<is\b[^>]*>([\s\S]*?)<\/is>/.exec(inner);
      const text = is ? textOf(is[1]) : '';
      return text === '' ? null : text;
    }
    const vm = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(inner);
    if (!vm) return null;
    const v = vm[1];
    switch (a.t) {
      case 's': {
        const s = ctx.strings[parseInt(v, 10)];
        return s === undefined || s === '' ? null : s;
      }
      case 'str':
      case 'e':
        return v === '' ? null : decodeText(v);
      case 'b':
        return v === '1' || v === 'true';
      case 'd': {
        const d = new Date(v);
        return isNaN(d.getTime()) ? decodeText(v) : d;
      }
      default: {
        if (v === '') return null;
        const n = Number(v);
        if (!isFinite(n)) return decodeText(v);
        return ctx.dateXf[parseInt(a.s || '0', 10)] ? Values.serialToDate(n, ctx.date1904) : n;
      }
    }
  }

  function parseSheet(xml, ctx) {
    xml = stripPrefix(xml, 'worksheet');
    const sheetData = /<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/.exec(xml);
    if (!sheetData) return [];

    const sparse = [];
    let maxRow = -1;
    let maxCol = -1;
    let rowIndex = -1;
    const rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
    let rm;
    while ((rm = rowRe.exec(sheetData[1]))) {
      const ra = attrs(rm[1]);
      rowIndex = ra.r ? parseInt(ra.r, 10) - 1 : rowIndex + 1;
      if (!rm[2]) continue;

      let colIndex = -1;
      const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let cm;
      while ((cm = cellRe.exec(rm[2]))) {
        const ca = attrs(cm[1]);
        const ref = ca.r ? /^([A-Za-z]+)\d+$/.exec(ca.r) : null;
        colIndex = ref ? columnIndex(ref[1]) : colIndex + 1;
        const value = readCell(ca, cm[2] || '', ctx);
        if (value === null) continue;
        (sparse[rowIndex] || (sparse[rowIndex] = []))[colIndex] = value;
        if (rowIndex > maxRow) maxRow = rowIndex;
        if (colIndex > maxCol) maxCol = colIndex;
      }
    }

    const rows = [];
    for (let r = 0; r <= maxRow; r++) {
      const row = new Array(maxCol + 1).fill(null);
      if (sparse[r]) sparse[r].forEach((value, c) => { row[c] = value; });
      rows.push(row);
    }
    return rows;
  }

  async function read(input) {
    const zip = Zip.open(input);
    if (!zip.has('xl/workbook.xml')) throw new Error('This file is not an Excel workbook (xl/workbook.xml is missing).');

    const workbook = stripPrefix(await zip.text('xl/workbook.xml'), 'workbook');
    const relsXml = zip.has('xl/_rels/workbook.xml.rels')
      ? stripPrefix(await zip.text('xl/_rels/workbook.xml.rels'), 'Relationships') : '';

    const targets = {};
    const relRe = /<Relationship\b([^>]*?)\/?>/g;
    let m;
    while ((m = relRe.exec(relsXml))) {
      const a = attrs(m[1]);
      if (a.Id && a.Target) targets[a.Id] = resolveTarget(a.Target);
    }

    const ctx = {
      date1904: /<workbookPr\b[^>]*date1904\s*=\s*["'](?:1|true)["']/.test(workbook),
      strings: parseSharedStrings(zip.has('xl/sharedStrings.xml') ? await zip.text('xl/sharedStrings.xml') : ''),
      dateXf: parseStyles(zip.has('xl/styles.xml') ? await zip.text('xl/styles.xml') : ''),
    };

    const sheets = [];
    const sheetRe = /<sheet\b([^>]*?)\/?>/g;
    let index = 0;
    while ((m = sheetRe.exec(workbook))) {
      index++;
      const a = attrs(m[1]);
      const relId = a['r:id'] || Object.keys(a).filter((k) => k.endsWith(':id')).map((k) => a[k])[0];
      const path = (relId && targets[relId]) || 'xl/worksheets/sheet' + index + '.xml';
      if (!zip.has(path)) continue; // chart sheets and similar
      sheets.push({ name: decodeText(a.name || 'Sheet' + index), rows: parseSheet(await zip.text(path), ctx) });
    }
    if (!sheets.length) throw new Error('The workbook has no readable sheets.');
    return { sheets };
  }

  // ------------------------------------------------------------------ write

  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  const STYLES = XML_HEAD + '<styleSheet xmlns="' + NS_MAIN + '">'
    + '<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/><numFmt numFmtId="165" formatCode="yyyy\\-mm\\-dd\\ hh:mm:ss"/></numFmts>'
    + '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
    + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="4">'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    + '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    + '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + '</cellXfs>'
    + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    + '</styleSheet>';

  function sheetXml(rows, opt) {
    const columns = rows.reduce((max, r) => Math.max(max, r.length), 0);
    const header = opt.header !== false;

    // column widths from the first rows
    const widths = new Array(columns).fill(8);
    rows.slice(0, 200).forEach((row) => row.forEach((v, c) => {
      const len = Values.format(v).split('\n').reduce((max, line) => Math.max(max, line.length), 0);
      widths[c] = Math.max(widths[c], Math.min(60, len * 1.15 + 2));
    }));

    let xml = XML_HEAD + '<worksheet xmlns="' + NS_MAIN + '" xmlns:r="' + NS_REL + '">';
    xml += '<sheetViews><sheetView workbookViewId="0"' + (opt.rtl ? ' rightToLeft="1"' : '') + '>'
      + (header && rows.length > 1 ? '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' : '')
      + '</sheetView></sheetViews>';
    xml += '<sheetFormatPr defaultRowHeight="15"/>';
    if (columns) {
      xml += '<cols>' + widths.map((w, i) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w.toFixed(1) + '" customWidth="1"/>').join('') + '</cols>';
    }

    xml += '<sheetData>';
    rows.forEach((row, r) => {
      const cells = [];
      row.forEach((value, c) => {
        if (value === null || value === undefined || value === '') return;
        const ref = Values.columnName(c) + (r + 1);
        const bold = header && r === 0;
        if (value instanceof Date && !bold) {
          if (isNaN(value.getTime())) return;
          const serial = Math.round((value.getTime() / 86400000 + 25569) * 1e8) / 1e8;
          cells.push('<c r="' + ref + '" s="' + (value.getTime() % 86400000 === 0 ? 2 : 3) + '"><v>' + serial + '</v></c>');
        } else if (typeof value === 'number' && isFinite(value) && !bold) {
          cells.push('<c r="' + ref + '"><v>' + value + '</v></c>');
        } else if (typeof value === 'boolean' && !bold) {
          cells.push('<c r="' + ref + '" t="b"><v>' + (value ? 1 : 0) + '</v></c>');
        } else {
          const text = Values.format(value).slice(0, 32767);
          cells.push('<c r="' + ref + '" t="inlineStr"' + (bold ? ' s="1"' : '') + '><is><t xml:space="preserve">' + escapeXml(text) + '</t></is></c>');
        }
      });
      if (cells.length) xml += '<row r="' + (r + 1) + '">' + cells.join('') + '</row>';
    });
    xml += '</sheetData>';
    if (header && rows.length > 1 && columns) {
      xml += '<autoFilter ref="A1:' + Values.columnName(columns - 1) + rows.length + '"/>';
    }
    return xml + '</worksheet>';
  }

  async function write(sheets, opt) {
    opt = opt || {};
    if (!sheets.length) throw new Error('Nothing to export: the workbook has no sheets.');
    const used = new Set();
    const named = sheets.map((s) => ({ name: Values.sheetName(s.name, used), rows: s.rows }));

    const files = [
      {
        name: '[Content_Types].xml',
        data: XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          + '<Default Extension="xml" ContentType="application/xml"/>'
          + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          + named.map((s, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('')
          + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
          + '</Types>',
      },
      {
        name: '_rels/.rels',
        data: XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          + '<Relationship Id="rId1" Type="' + NS_REL + '/officeDocument" Target="xl/workbook.xml"/></Relationships>',
      },
      {
        name: 'xl/workbook.xml',
        data: XML_HEAD + '<workbook xmlns="' + NS_MAIN + '" xmlns:r="' + NS_REL + '"><bookViews><workbookView/></bookViews><sheets>'
          + named.map((s, i) => '<sheet name="' + escapeXml(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('')
          + '</sheets></workbook>',
      },
      {
        name: 'xl/_rels/workbook.xml.rels',
        data: XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          + named.map((s, i) => '<Relationship Id="rId' + (i + 1) + '" Type="' + NS_REL + '/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('')
          + '<Relationship Id="rId' + (named.length + 1) + '" Type="' + NS_REL + '/styles" Target="styles.xml"/></Relationships>',
      },
      { name: 'xl/styles.xml', data: STYLES },
    ];
    named.forEach((s, i) => files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: sheetXml(s.rows, opt) }));

    return Zip.create(files);
  }

  return { read, write };
});
