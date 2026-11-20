/*
 * Minimal ZIP reader/writer (enough for .xlsx files). No dependencies.
 * Uses the browser's built-in DecompressionStream / CompressionStream ('deflate-raw').
 * Works as a plain <script> (global `Zip`) and as a CommonJS module (for the Node tests).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Zip = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const encoder = new TextEncoder();
  const decoder = new TextDecoder('utf-8');

  const CRC_TABLE = (function () {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  async function runStream(bytes, stream) {
    const piped = new Blob([bytes]).stream().pipeThrough(stream);
    return new Uint8Array(await new Response(piped).arrayBuffer());
  }

  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === 'function') {
      try { return await runStream(bytes, new DecompressionStream('deflate-raw')); } catch (e) { /* try zlib below */ }
    }
    if (typeof require === 'function') return new Uint8Array(require('zlib').inflateRawSync(bytes));
    throw new Error('This browser cannot open .xlsx files (DecompressionStream is missing). Update the browser.');
  }

  async function deflateRaw(bytes) {
    if (typeof CompressionStream === 'function') {
      try { return await runStream(bytes, new CompressionStream('deflate-raw')); } catch (e) { /* store instead */ }
    }
    if (typeof require === 'function') return new Uint8Array(require('zlib').deflateRawSync(bytes));
    return null; // caller stores the file uncompressed
  }

  // ------------------------------------------------------------------ read

  function view(u8) {
    return new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  }

  function readEntries(u8) {
    const dv = view(u8);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 22 - 65535); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('This is not a valid .xlsx file (zip directory not found).');

    const count = dv.getUint16(eocd + 10, true);
    const cdOffset = dv.getUint32(eocd + 16, true);
    if (count === 0xFFFF || cdOffset === 0xFFFFFFFF) throw new Error('Zip64 archives are not supported.');

    const entries = new Map();
    let p = cdOffset;
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('The zip directory is corrupted.');
      const nameLen = dv.getUint16(p + 28, true);
      const extraLen = dv.getUint16(p + 30, true);
      const commentLen = dv.getUint16(p + 32, true);
      const name = decoder.decode(u8.subarray(p + 46, p + 46 + nameLen));
      entries.set(name, {
        name,
        flags: dv.getUint16(p + 8, true),
        method: dv.getUint16(p + 10, true),
        compressedSize: dv.getUint32(p + 20, true),
        localOffset: dv.getUint32(p + 42, true),
      });
      p += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  async function extract(u8, entry) {
    if (entry.flags & 1) throw new Error('Password-protected files are not supported.');
    const dv = view(u8);
    const o = entry.localOffset;
    if (dv.getUint32(o, true) !== 0x04034b50) throw new Error('The zip file is corrupted.');
    const start = o + 30 + dv.getUint16(o + 26, true) + dv.getUint16(o + 28, true);
    const data = u8.subarray(start, start + entry.compressedSize);
    if (entry.method === 0) return data;
    if (entry.method === 8) return inflateRaw(data);
    throw new Error('Unsupported zip compression method ' + entry.method + '.');
  }

  /** Open a zip from an ArrayBuffer / Uint8Array. Returns { has(name), text(name), names() }. */
  function open(input) {
    const u8 = input instanceof Uint8Array ? input : new Uint8Array(input);
    const entries = readEntries(u8);
    return {
      names: () => Array.from(entries.keys()),
      has: (name) => entries.has(name),
      async bytes(name) {
        const entry = entries.get(name);
        if (!entry) throw new Error('File not found inside the archive: ' + name);
        return extract(u8, entry);
      },
      async text(name) {
        return decoder.decode(await this.bytes(name));
      },
    };
  }

  // ----------------------------------------------------------------- write

  /** Create a zip. files: [{ name, data: string | Uint8Array }]. Resolves to a Uint8Array. */
  async function create(files) {
    const now = new Date();
    const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();

    const parts = [];
    const central = [];
    let offset = 0;

    for (const file of files) {
      const nameBytes = encoder.encode(file.name);
      const raw = typeof file.data === 'string' ? encoder.encode(file.data) : file.data;
      const crc = crc32(raw);

      let method = 0;
      let data = raw;
      if (raw.length > 128) {
        const packed = await deflateRaw(raw);
        if (packed && packed.length < raw.length) { method = 8; data = packed; }
      }

      const local = new Uint8Array(30 + nameBytes.length);
      const l = view(local);
      l.setUint32(0, 0x04034b50, true);
      l.setUint16(4, 20, true);
      l.setUint16(6, 0x0800, true); // UTF-8 names
      l.setUint16(8, method, true);
      l.setUint16(10, dosTime, true);
      l.setUint16(12, dosDate, true);
      l.setUint32(14, crc, true);
      l.setUint32(18, data.length, true);
      l.setUint32(22, raw.length, true);
      l.setUint16(26, nameBytes.length, true);
      local.set(nameBytes, 30);

      const entry = new Uint8Array(46 + nameBytes.length);
      const c = view(entry);
      c.setUint32(0, 0x02014b50, true);
      c.setUint16(4, 20, true);
      c.setUint16(6, 20, true);
      c.setUint16(8, 0x0800, true);
      c.setUint16(10, method, true);
      c.setUint16(12, dosTime, true);
      c.setUint16(14, dosDate, true);
      c.setUint32(16, crc, true);
      c.setUint32(20, data.length, true);
      c.setUint32(24, raw.length, true);
      c.setUint16(28, nameBytes.length, true);
      c.setUint32(42, offset, true);
      entry.set(nameBytes, 46);

      parts.push(local, data);
      central.push(entry);
      offset += local.length + data.length;
    }

    const centralSize = central.reduce((sum, e) => sum + e.length, 0);
    const end = new Uint8Array(22);
    const e = view(end);
    e.setUint32(0, 0x06054b50, true);
    e.setUint16(8, files.length, true);
    e.setUint16(10, files.length, true);
    e.setUint32(12, centralSize, true);
    e.setUint32(16, offset, true);

    const all = parts.concat(central, [end]);
    const out = new Uint8Array(all.reduce((sum, p) => sum + p.length, 0));
    let pos = 0;
    for (const p of all) { out.set(p, pos); pos += p.length; }
    return out;
  }

  return { open, create, crc32 };
});
