/*
 * Legacy Excel 97-2003 (.xls, BIFF8) reader and writer. No dependencies.
 *
 *   XlsLite.isXls(bytes)           -> true when the bytes start with the OLE2 signature
 *   XlsLite.read(bytes)            -> { sheets: [{ name, rows }] }   (synchronous)
 *   XlsLite.write(sheets, opt)     -> Uint8Array                      (synchronous)   opt: { header }
 *
 * `rows` uses the same cell values as the rest of the project (see values.js).
 * Reading supports BIFF8 (Excel 97 and newer): shared strings (with CONTINUE records), numbers, RK and MULRK,
 * booleans, errors, cached formula results, dates by number format, 1900/1904 date systems.
 * Not supported: Excel 5.0/95 (BIFF5) and older files, password-protected files, formatting, charts.
 * Writing is limited to 65,536 rows and 256 columns (the .xls limits).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./values.js'));
  else root.XlsLite = factory(root.Values);
})(typeof self !== 'undefined' ? self : this, function (Values) {
  'use strict';

  const SIGNATURE = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  const FIRST_SPECIAL_SECTOR = 0xFFFFFFFA; // ids from here up mean "free", "end of chain", etc.

  const ERRORS = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!', 0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A' };

  const u16 = (d, o) => d[o] | (d[o + 1] << 8);
  const u32 = (d, o) => (d[o] | (d[o + 1] << 8) | (d[o + 2] << 16) | (d[o + 3] << 24)) >>> 0;
  const f64 = (d, o) => new DataView(d.buffer, d.byteOffset, d.byteLength).getFloat64(o, true);

  function isXls(bytes) {
    return bytes.length >= 8 && SIGNATURE.every((b, i) => bytes[i] === b);
  }

  // ------------------------------------------------ OLE2 compound file (CFB)

  /** Returns { stream(name) -> Uint8Array | null } for a compound file. */
  function openCompound(u8) {
    if (!isXls(u8)) throw new Error('This is not a valid .xls file.');
    const sectorSize = 1 << u16(u8, 30);
    const miniSize = 1 << u16(u8, 32);
    const numFat = u32(u8, 44);
    const dirStart = u32(u8, 48);
    const cutoff = u32(u8, 56);
    const miniFatStart = u32(u8, 60);

    const sector = (id) => {
      const start = (id + 1) * sectorSize;
      if (start >= u8.length) throw new Error('The .xls file is truncated or corrupted.');
      return u8.subarray(start, Math.min(start + sectorSize, u8.length));
    };

    // sector numbers of the FAT: 109 in the header, the rest in DIFAT sectors
    const fatSectors = [];
    for (let i = 0; i < 109 && fatSectors.length < numFat; i++) {
      const id = u32(u8, 76 + i * 4);
      if (id < FIRST_SPECIAL_SECTOR) fatSectors.push(id);
    }
    let difat = u32(u8, 68);
    for (let n = 0, count = u32(u8, 72); n < count && difat < FIRST_SPECIAL_SECTOR; n++) {
      const s = sector(difat);
      const per = sectorSize / 4 - 1;
      for (let i = 0; i < per && fatSectors.length < numFat; i++) {
        const id = u32(s, i * 4);
        if (id < FIRST_SPECIAL_SECTOR) fatSectors.push(id);
      }
      difat = u32(s, per * 4);
    }

    const fat = [];
    fatSectors.forEach((id) => { const s = sector(id); for (let i = 0; i + 4 <= s.length; i += 4) fat.push(u32(s, i)); });

    function readChain(start, table, unit, source) {
      const parts = [];
      let total = 0;
      let id = start;
      for (let guard = 0; id < FIRST_SPECIAL_SECTOR && guard <= table.length; guard++) {
        const part = source(id);
        parts.push(part);
        total += part.length;
        id = table[id];
      }
      const out = new Uint8Array(total);
      let pos = 0;
      parts.forEach((p) => { out.set(p, pos); pos += p.length; });
      return out;
    }

    const dir = readChain(dirStart, fat, sectorSize, sector);
    const entries = [];
    for (let o = 0; o + 128 <= dir.length; o += 128) {
      const nameLen = Math.max(0, Math.min(u16(dir, o + 64), 64) - 2);
      entries.push({
        name: new TextDecoder('utf-16le').decode(dir.subarray(o, o + nameLen)),
        type: dir[o + 66],
        start: u32(dir, o + 116),
        size: u32(dir, o + 120),
      });
    }
    const root = entries[0];

    let miniFat = null;
    let miniStream = null;
    const loadMini = () => {
      if (miniFat) return;
      const raw = readChain(miniFatStart, fat, sectorSize, sector);
      miniFat = [];
      for (let i = 0; i + 4 <= raw.length; i += 4) miniFat.push(u32(raw, i));
      miniStream = readChain(root.start, fat, sectorSize, sector);
    };

    return {
      names: () => entries.filter((e) => e.type === 2).map((e) => e.name),
      stream(name) {
        const entry = entries.find((e) => e.type === 2 && e.name.toLowerCase() === name.toLowerCase());
        if (!entry) return null;
        if (entry.size < cutoff) {
          loadMini();
          const data = readChain(entry.start, miniFat, miniSize, (id) => miniStream.subarray(id * miniSize, (id + 1) * miniSize));
          return data.subarray(0, entry.size);
        }
        return readChain(entry.start, fat, sectorSize, sector).subarray(0, entry.size);
      },
    };
  }

  // -------------------------------------------------------------- BIFF reading

  function decodeChars(bytes, wide) {
    if (wide) return new TextDecoder('utf-16le').decode(bytes);
    let s = '';
    for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return s;
  }

  /** XLUnicodeString inside one record: u16 length, u8 flags, characters. */
  function readUnicode(d, o) {
    const cch = u16(d, o);
    const flags = d[o + 2];
    o += 3;
    if (flags & 8) o += 2;
    if (flags & 4) o += 4;
    const wide = !!(flags & 1);
    return decodeChars(d.subarray(o, o + cch * (wide ? 2 : 1)), wide);
  }

  /** The shared string table. It can continue over several CONTINUE records, which re-state the "wide" flag. */
  function readSharedStrings(segments) {
    const out = [];
    let si = 0;
    let seg = segments[0];
    let p = 8; // skip the total and unique counts
    const unique = u32(seg, 4);

    const nextSegment = () => {
      si++;
      seg = segments[si];
      p = 0;
      return seg !== undefined;
    };

    for (let n = 0; n < unique; n++) {
      if (p >= seg.length && !nextSegment()) break;
      const cch = u16(seg, p);
      const flags = seg[p + 2];
      p += 3;
      const runs = flags & 8 ? u16(seg, p) : 0;
      if (flags & 8) p += 2;
      const extra = flags & 4 ? u32(seg, p) : 0;
      if (flags & 4) p += 4;

      let wide = !!(flags & 1);
      let text = '';
      let remaining = cch;
      while (remaining > 0) {
        if (p >= seg.length) {
          if (!nextSegment()) break;
          wide = !!(seg[p] & 1);
          p++;
          continue;
        }
        const width = wide ? 2 : 1;
        const take = Math.min(remaining, Math.floor((seg.length - p) / width));
        if (take === 0) { // a single stray byte at the end of a segment
          if (!nextSegment()) break;
          wide = !!(seg[p] & 1);
          p++;
          continue;
        }
        text += decodeChars(seg.subarray(p, p + take * width), wide);
        p += take * width;
        remaining -= take;
      }

      let skip = runs * 4 + extra;
      while (skip > 0) {
        if (p >= seg.length && !nextSegment()) break;
        const step = Math.min(skip, seg.length - p);
        p += step;
        skip -= step;
      }
      out.push(text);
    }
    return out;
  }

  function decodeRk(v) {
    let n;
    if (v & 2) {
      n = (v | 0) >> 2;
    } else {
      const view = new DataView(new ArrayBuffer(8));
      view.setUint32(4, (v & 0xFFFFFFFC) >>> 0, true);
      n = view.getFloat64(0, true);
    }
    return v & 1 ? n / 100 : n;
  }

  function readWorkbookGlobals(wb) {
    if (wb.length < 8 || u16(wb, 0) !== 0x0809) {
      throw new Error('This .xls file is too old (Excel 4 or earlier). Open it in Excel and save it as .xlsx.');
    }
    if (u16(wb, 4) !== 0x0600) {
      throw new Error('This .xls file was saved by Excel 5.0/95 (BIFF5), which is not supported. Open it in Excel and save it as .xlsx.');
    }
    const info = { date1904: false, xf: [], formats: {}, sst: [], sheets: [] };

    let p = 0;
    while (p + 4 <= wb.length) {
      const id = u16(wb, p);
      const size = u16(wb, p + 2);
      const d = wb.subarray(p + 4, p + 4 + size);
      p += 4 + size;

      if (id === 0x000A) break; // EOF of the globals substream
      if (id === 0x002F) throw new Error('This .xls file is password-protected, which is not supported.');
      if (id === 0x0022) info.date1904 = u16(d, 0) === 1;
      else if (id === 0x00E0) info.xf.push(u16(d, 2));
      else if (id === 0x041E) info.formats[u16(d, 0)] = readUnicode(d, 2);
      else if (id === 0x0085) {
        const wide = !!(d[7] & 1);
        info.sheets.push({
          offset: u32(d, 0),
          type: d[5],
          name: decodeChars(d.subarray(8, 8 + d[6] * (wide ? 2 : 1)), wide),
        });
      } else if (id === 0x00FC) {
        const segments = [d];
        while (p + 4 <= wb.length && u16(wb, p) === 0x003C) { // CONTINUE
          const n = u16(wb, p + 2);
          segments.push(wb.subarray(p + 4, p + 4 + n));
          p += 4 + n;
        }
        info.sst = readSharedStrings(segments);
      }
    }
    info.dateXf = info.xf.map((formatId) => Values.isDateFormat(formatId, info.formats));
    return info;
  }

  function readSheet(wb, start, info) {
    const sparse = [];
    let maxRow = -1;
    let maxCol = -1;
    let pending = null; // a formula cell whose text result comes in the next STRING record

    const put = (r, c, value) => {
      if (value === null || value === undefined || value === '') return;
      (sparse[r] || (sparse[r] = []))[c] = value;
      if (r > maxRow) maxRow = r;
      if (c > maxCol) maxCol = c;
    };
    const number = (n, xf) => (info.dateXf[xf] ? Values.serialToDate(n, info.date1904) : n);

    let depth = 0;
    let p = start;
    while (p + 4 <= wb.length) {
      const id = u16(wb, p);
      const size = u16(wb, p + 2);
      const d = wb.subarray(p + 4, p + 4 + size);
      p += 4 + size;

      if (id === 0x0809) { depth++; continue; }
      if (id === 0x000A) { if (--depth <= 0) break; continue; }
      if (depth !== 1) continue; // inside an embedded chart or similar

      const row = u16(d, 0);
      const col = u16(d, 2);
      const xf = u16(d, 4);

      switch (id) {
        case 0x00FD: // LABELSST
          put(row, col, info.sst[u32(d, 6)]);
          break;
        case 0x0203: // NUMBER
          put(row, col, number(f64(d, 6), xf));
          break;
        case 0x027E: // RK
          put(row, col, number(decodeRk(u32(d, 6)), xf));
          break;
        case 0x00BD: { // MULRK
          const last = u16(d, d.length - 2);
          for (let c = col, o = 4; c <= last && o + 6 <= d.length - 2; c++, o += 6) {
            put(row, c, number(decodeRk(u32(d, o + 2)), u16(d, o)));
          }
          break;
        }
        case 0x0205: // BOOLERR
          put(row, col, d[7] ? (ERRORS[d[6]] || '#ERROR') : d[6] !== 0);
          break;
        case 0x0204: // LABEL
          put(row, col, readUnicode(d, 6));
          break;
        case 0x0006: // FORMULA: the cached result is what we show
          pending = null;
          if (d[12] === 0xFF && d[13] === 0xFF) {
            if (d[6] === 0) pending = { row, col };
            else if (d[6] === 1) put(row, col, d[8] !== 0);
            else if (d[6] === 2) put(row, col, ERRORS[d[8]] || '#ERROR');
          } else {
            put(row, col, number(f64(d, 6), xf));
          }
          break;
        case 0x0207: // STRING (text result of the formula before it)
          if (pending) put(pending.row, pending.col, readUnicode(d, 0));
          pending = null;
          break;
        default:
      }
    }

    const rows = [];
    for (let r = 0; r <= maxRow; r++) {
      const out = new Array(maxCol + 1).fill(null);
      if (sparse[r]) sparse[r].forEach((value, c) => { out[c] = value; });
      rows.push(out);
    }
    return rows;
  }

  function read(input) {
    const u8 = input instanceof Uint8Array ? input : new Uint8Array(input);
    const file = openCompound(u8);
    const wb = file.stream('Workbook');
    if (!wb) {
      if (file.stream('Book')) throw new Error('This .xls file was saved by Excel 5.0/95, which is not supported. Open it in Excel and save it as .xlsx.');
      throw new Error('This file is not an Excel workbook (no Workbook stream found).');
    }

    const info = readWorkbookGlobals(wb);
    const sheets = [];
    info.sheets.forEach((s, i) => {
      if (s.type !== 0 || u16(wb, s.offset) !== 0x0809) return; // skip chart sheets and macro sheets
      sheets.push({ name: s.name || 'Sheet' + (i + 1), rows: readSheet(wb, s.offset, info) });
    });
    if (!sheets.length) throw new Error('The workbook has no readable sheets.');
    return { sheets };
  }

  // -------------------------------------------------------------- BIFF writing

  const MAX_ROWS = 65536;
  const MAX_COLS = 256;
  const MAX_RECORD = 8224; // maximum data bytes in one BIFF8 record

  // cell XF indexes used by the writer
  const XF_GENERAL = 15;
  const XF_BOLD = 16;
  const XF_DATE = 17;
  const XF_DATETIME = 18;
  const FMT_DATE = 164;
  const FMT_DATETIME = 165;

  /** A growable little-endian byte buffer. */
  class Out {
    constructor(capacity) {
      this.buf = new Uint8Array(capacity || 1 << 16);
      this.view = new DataView(this.buf.buffer);
      this.length = 0;
    }

    ensure(n) {
      if (this.length + n <= this.buf.length) return;
      let capacity = this.buf.length * 2;
      while (capacity < this.length + n) capacity *= 2;
      const grown = new Uint8Array(capacity);
      grown.set(this.buf.subarray(0, this.length));
      this.buf = grown;
      this.view = new DataView(grown.buffer);
    }

    u8(v) { this.ensure(1); this.buf[this.length++] = v; return this; }
    u16(v) { this.ensure(2); this.view.setUint16(this.length, v, true); this.length += 2; return this; }
    u32(v) { this.ensure(4); this.view.setUint32(this.length, v, true); this.length += 4; return this; }
    f64(v) { this.ensure(8); this.view.setFloat64(this.length, v, true); this.length += 8; return this; }
    bytes(b) { this.ensure(b.length); this.buf.set(b, this.length); this.length += b.length; return this; }
    zeros(n) { this.ensure(n); this.length += n; return this; }
    result() { return this.buf.slice(0, this.length); }
  }

  /** Append one record: id, size, data (a Uint8Array or a function that fills an Out). */
  function record(out, id, data) {
    if (typeof data === 'function') { const o = new Out(64); data(o); data = o.result(); }
    out.u16(id).u16(data.length).bytes(data);
  }

  const isLatin1 = (s) => !/[^\u0000-ÿ]/.test(s);

  function utf16(s) {
    const b = new Uint8Array(s.length * 2);
    for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); b[i * 2] = c & 0xFF; b[i * 2 + 1] = c >> 8; }
    return b;
  }

  function latin1(s) {
    const b = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i);
    return b;
  }

  /** XF record: 20 bytes. Layout follows files written by LibreOffice. */
  function xfRecord(o, font, format, isStyle, used) {
    o.u16(font).u16(format);
    o.u16(isStyle ? 0xFFF5 : 0x0001); // locked + style flag + parent index
    o.u8(0x20).u8(0).u8(0).u8(used);  // alignment, rotation, indent, attributes in use
    o.u32(0).u32(0);                  // borders
    o.u16(0x20C0);                    // fill colours (automatic)
  }

  function fontRecord(o, height, weight, name) {
    o.u16(height).u16(0).u16(name === 'Arial' ? 0x7FFF : 8).u16(weight).u16(0);
    o.u8(0).u8(2).u8(1).u8(0);
    o.u8(name.length).u8(1).bytes(utf16(name));
  }

  function writeGlobals(out, names, sheetOffsets, sst) {
    const sheetCount = names.length;
    record(out, 0x0809, (o) => o.u16(0x0600).u16(0x0005).u16(0x0DBB).u16(0x07CC).u32(0).u32(6));
    record(out, 0x00E1, (o) => o.u16(0x04B0));                     // INTERFACEHDR
    record(out, 0x00C1, (o) => o.u16(0));                          // MMS
    record(out, 0x00E2, new Uint8Array(0));                        // INTERFACEEND
    record(out, 0x0042, (o) => o.u16(0x04B0));                     // CODEPAGE: UTF-16
    record(out, 0x0161, (o) => o.u16(0));                          // DSF
    record(out, 0x013D, (o) => { for (let i = 1; i <= sheetCount; i++) o.u16(i); }); // TABID
    record(out, 0x009C, (o) => o.u16(14));                         // FNGROUPCOUNT
    record(out, 0x003D, (o) => o.u16(0).u16(0).u16(0x4000).u16(0x2000).u16(0x0038).u16(0).u16(0).u16(1).u16(0x0258)); // WINDOW1
    record(out, 0x0040, (o) => o.u16(0));                          // BACKUP
    record(out, 0x008D, (o) => o.u16(0));                          // HIDEOBJ
    record(out, 0x0022, (o) => o.u16(0));                          // DATEMODE: 1900
    record(out, 0x000E, (o) => o.u16(1));                          // PRECISION
    record(out, 0x01B7, (o) => o.u16(0));                          // REFRESHALL
    record(out, 0x00DA, (o) => o.u16(0));                          // BOOKBOOL

    // five fonts (index 4 is never used by Excel): normal, three filler, bold
    record(out, 0x0031, (o) => fontRecord(o, 220, 400, 'Calibri'));
    for (let i = 0; i < 3; i++) record(out, 0x0031, (o) => fontRecord(o, 200, 400, 'Arial'));
    record(out, 0x0031, (o) => fontRecord(o, 220, 700, 'Calibri'));

    const formats = [[FMT_DATE, 'yyyy\\-mm\\-dd'], [FMT_DATETIME, 'yyyy\\-mm\\-dd\\ hh:mm:ss']];
    formats.forEach(([id, code]) => record(out, 0x041E, (o) => o.u16(id).u16(code.length).u8(0).bytes(latin1(code))));

    for (let i = 0; i < 15; i++) record(out, 0x00E0, (o) => xfRecord(o, 0, 0, true, i === 0 ? 0 : 0xF4)); // style XFs
    record(out, 0x00E0, (o) => xfRecord(o, 0, 0, false, 0));                  // 15 general
    record(out, 0x00E0, (o) => xfRecord(o, 5, 0, false, 0x08));                // 16 bold
    record(out, 0x00E0, (o) => xfRecord(o, 0, FMT_DATE, false, 0x04));         // 17 date
    record(out, 0x00E0, (o) => xfRecord(o, 0, FMT_DATETIME, false, 0x04));     // 18 date and time

    record(out, 0x0293, (o) => o.u16(0x8000).u8(0).u8(0xFF));      // STYLE: Normal
    record(out, 0x0160, (o) => o.u16(0));                          // USESELFS

    names.forEach((name, i) => record(out, 0x0085, (o) => {
      const latin = isLatin1(name);
      o.u32(sheetOffsets[i]).u8(0).u8(0).u8(name.length).u8(latin ? 0 : 1).bytes(latin ? latin1(name) : utf16(name));
    }));
    record(out, 0x008C, (o) => o.u16(1).u16(1));                   // COUNTRY

    writeSharedStrings(out, sst);
    record(out, 0x000A, new Uint8Array(0));                        // EOF
  }

  /** SST record plus CONTINUE records. Text may be split between records at a character boundary. */
  function writeSharedStrings(out, sst) {
    const records = [];
    let current = new Out(MAX_RECORD + 16);
    let room = MAX_RECORD - 8;
    const startNext = () => { records.push(current); current = new Out(MAX_RECORD + 16); room = MAX_RECORD; };

    sst.list.forEach((text) => {
      const wide = !isLatin1(text);
      const width = wide ? 2 : 1;
      if (room < 3 + width) startNext();
      current.u16(text.length).u8(wide ? 1 : 0);
      room -= 3;

      let i = 0;
      while (i < text.length) {
        const fit = Math.floor(room / width);
        if (fit === 0) { startNext(); current.u8(wide ? 1 : 0); room -= 1; continue; }
        const piece = text.slice(i, i + fit);
        current.bytes(wide ? utf16(piece) : latin1(piece));
        room -= piece.length * width;
        i += piece.length;
      }
    });
    records.push(current);

    records.forEach((rec, i) => {
      if (i === 0) {
        record(out, 0x00FC, (o) => o.u32(sst.total).u32(sst.list.length).bytes(rec.result()));
      } else {
        record(out, 0x003C, rec.result());
      }
    });
  }

  function writeSheet(rows, opt, sst, isFirst) {
    const header = opt.header !== false;
    const rowCount = rows.length;
    const colCount = rows.reduce((max, r) => Math.max(max, r.length), 0);
    if (rowCount > MAX_ROWS || colCount > MAX_COLS) {
      throw new Error('The .xls format supports at most 65,536 rows and 256 columns. Export as .xlsx instead.');
    }

    const out = new Out();
    record(out, 0x0809, (o) => o.u16(0x0600).u16(0x0010).u16(0x0DBB).u16(0x07CC).u32(0).u32(6));
    record(out, 0x0055, (o) => o.u16(8)); // DEFCOLWIDTH

    // column widths from the first rows, like the .xlsx writer
    const widths = new Array(colCount).fill(8);
    rows.slice(0, 200).forEach((row) => row.forEach((v, c) => {
      const len = Values.format(v).split('\n').reduce((max, line) => Math.max(max, line.length), 0);
      widths[c] = Math.max(widths[c], Math.min(60, len * 1.15 + 2));
    }));
    widths.forEach((w, c) => record(out, 0x007D, (o) => o.u16(c).u16(c).u16(Math.min(65280, Math.round(w * 256))).u16(XF_GENERAL).u16(0).u16(0)));

    // which rows have cells
    const filled = rows.map((row) => row.some((v) => v !== null && v !== undefined && v !== ''));
    const lastFilled = filled.lastIndexOf(true);
    record(out, 0x0200, (o) => o.u32(0).u32(lastFilled + 1).u16(0).u16(lastFilled < 0 ? 0 : colCount).u16(0)); // DIMENSIONS

    // rows are written in blocks of 32: the ROW records first, then the cells of those rows
    for (let start = 0; start <= lastFilled; start += 32) {
      const block = [];
      for (let r = start; r < Math.min(start + 32, rowCount); r++) if (filled[r]) block.push(r);

      block.forEach((r) => {
        let first = -1;
        let last = -1;
        rows[r].forEach((v, c) => { if (v !== null && v !== undefined && v !== '') { if (first < 0) first = c; last = c; } });
        record(out, 0x0208, (o) => o.u16(r).u16(first).u16(last + 1).u16(300).u16(0).u16(0).u16(0x0100).u16(XF_GENERAL));
      });

      block.forEach((r) => {
        rows[r].forEach((value, c) => {
          if (value === null || value === undefined || value === '') return;
          const bold = header && r === 0;
          if (value instanceof Date && !bold) {
            if (isNaN(value.getTime())) return;
            const serial = value.getTime() / 86400000 + 25569;
            out.u16(0x0203).u16(14).u16(r).u16(c).u16(value.getTime() % 86400000 === 0 ? XF_DATE : XF_DATETIME).f64(serial);
          } else if (typeof value === 'number' && isFinite(value) && !bold) {
            out.u16(0x0203).u16(14).u16(r).u16(c).u16(XF_GENERAL).f64(value);
          } else if (typeof value === 'boolean' && !bold) {
            out.u16(0x0205).u16(8).u16(r).u16(c).u16(XF_GENERAL).u8(value ? 1 : 0).u8(0);
          } else {
            const text = Values.format(value).slice(0, 32767);
            out.u16(0x00FD).u16(10).u16(r).u16(c).u16(bold ? XF_BOLD : XF_GENERAL).u32(sst.index(text));
          }
        });
      });
    }

    record(out, 0x023E, (o) => o.u16(isFirst ? 0x06B6 : 0x04B6).u16(0).u16(0).u32(0x40).u16(60).u16(100).u32(0)); // WINDOW2
    record(out, 0x001D, (o) => o.u8(3).u16(0).u16(0).u16(0).u16(1).u16(0).u16(0).u8(0).u8(0)); // SELECTION
    record(out, 0x000A, new Uint8Array(0));
    return out.result();
  }

  // -------------------------------------------------- OLE2 compound file writer

  const END_OF_CHAIN = 0xFFFFFFFE;
  const FREE_SECTOR = 0xFFFFFFFF;
  const FAT_SECTOR = 0xFFFFFFFD;
  const DIFAT_SECTOR = 0xFFFFFFFC;

  /** Wrap one stream named "Workbook" in a version 3 compound file (512-byte sectors). */
  function buildCompound(stream) {
    const size = Math.max(4096, Math.ceil(stream.length / 512) * 512); // streams under 4096 bytes would need a mini stream
    const dataSectors = size / 512;
    const dirSector = dataSectors;

    // how many FAT sectors (and DIFAT sectors, when more than 109 FAT sectors are needed)
    let fatCount = 1;
    let difatCount = 0;
    for (;;) {
      const total = dataSectors + 1 + fatCount + difatCount;
      const needFat = Math.ceil(total / 128);
      const needDifat = needFat > 109 ? Math.ceil((needFat - 109) / 127) : 0;
      if (needFat === fatCount && needDifat === difatCount) break;
      fatCount = needFat;
      difatCount = needDifat;
    }
    const fatStart = dirSector + 1;
    const difatStart = fatStart + fatCount;
    const totalSectors = difatStart + difatCount;

    const fat = new Uint32Array(fatCount * 128).fill(FREE_SECTOR);
    for (let i = 0; i < dataSectors - 1; i++) fat[i] = i + 1;
    fat[dataSectors - 1] = END_OF_CHAIN;
    fat[dirSector] = END_OF_CHAIN;
    for (let i = 0; i < fatCount; i++) fat[fatStart + i] = FAT_SECTOR;
    for (let i = 0; i < difatCount; i++) fat[difatStart + i] = DIFAT_SECTOR;

    const out = new Uint8Array(512 + totalSectors * 512);
    const view = new DataView(out.buffer);
    out.set(SIGNATURE, 0);
    view.setUint16(24, 0x003E, true);
    view.setUint16(26, 3, true);
    view.setUint16(28, 0xFFFE, true);
    view.setUint16(30, 9, true);   // 512-byte sectors
    view.setUint16(32, 6, true);   // 64-byte mini sectors
    view.setUint32(44, fatCount, true);
    view.setUint32(48, dirSector, true);
    view.setUint32(56, 4096, true);
    view.setUint32(60, END_OF_CHAIN, true);
    view.setUint32(68, difatCount ? difatStart : END_OF_CHAIN, true);
    view.setUint32(72, difatCount, true);
    for (let i = 0; i < 109; i++) view.setUint32(76 + i * 4, i < fatCount ? fatStart + i : FREE_SECTOR, true);

    out.set(stream, 512);

    // extra DIFAT sectors list the FAT sectors beyond the first 109
    for (let d = 0; d < difatCount; d++) {
      const base = 512 + (difatStart + d) * 512;
      for (let i = 0; i < 127; i++) {
        const index = 109 + d * 127 + i;
        view.setUint32(base + i * 4, index < fatCount ? fatStart + index : FREE_SECTOR, true);
      }
      view.setUint32(base + 127 * 4, d + 1 < difatCount ? difatStart + d + 1 : END_OF_CHAIN, true);
    }

    // directory: Root Entry, Workbook, and two unused slots
    const entry = (index, name, type, child, start, length) => {
      const o = 512 + dirSector * 512 + index * 128;
      for (let i = 0; i < name.length; i++) view.setUint16(o + i * 2, name.charCodeAt(i), true);
      view.setUint16(o + 64, name ? (name.length + 1) * 2 : 0, true);
      out[o + 66] = type;
      out[o + 67] = 1;
      view.setUint32(o + 68, FREE_SECTOR, true);
      view.setUint32(o + 72, FREE_SECTOR, true);
      view.setUint32(o + 76, child, true);
      view.setUint32(o + 116, start, true);
      view.setUint32(o + 120, length, true);
    };
    entry(0, 'Root Entry', 5, 1, END_OF_CHAIN, 0);
    entry(1, 'Workbook', 2, FREE_SECTOR, 0, size);
    entry(2, '', 0, FREE_SECTOR, 0, 0);
    entry(3, '', 0, FREE_SECTOR, 0, 0);

    for (let i = 0; i < fat.length; i++) view.setUint32(512 + fatStart * 512 + i * 4, fat[i], true);
    return out;
  }

  /** sheets: [{ name, rows }], opt: { header } -> Uint8Array of a .xls file. */
  function write(sheets, opt) {
    opt = opt || {};
    if (!sheets.length) throw new Error('Nothing to export: the workbook has no sheets.');

    const used = new Set();
    const names = sheets.map((s) => Values.sheetName(s.name, used));

    // shared strings are collected while the sheets are written
    const sst = {
      list: [],
      total: 0,
      map: new Map(),
      index(text) {
        this.total++;
        let i = this.map.get(text);
        if (i === undefined) { i = this.list.length; this.list.push(text); this.map.set(text, i); }
        return i;
      },
    };
    const streams = sheets.map((s, i) => writeSheet(s.rows, opt, sst, i === 0));

    // the globals contain the byte offset of each sheet, and the offsets depend on the globals' length
    const measure = new Out();
    writeGlobals(measure, names, names.map(() => 0), sst);
    const offsets = [];
    let position = measure.length;
    streams.forEach((s) => { offsets.push(position); position += s.length; });

    const workbook = new Out(position + 16);
    writeGlobals(workbook, names, offsets, sst);
    streams.forEach((s) => workbook.bytes(s));
    return buildCompound(workbook.result());
  }

  return { isXls, read, write, openCompound, buildCompound };
});
