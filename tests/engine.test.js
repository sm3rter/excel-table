'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const Zip = require('../js/zip.js');
const Values = require('../js/values.js');
const XlsxLite = require('../js/xlsx.js');
const Tabular = require('../js/tabular.js');

const fixture = (name) => fs.readFileSync(path.join(__dirname, 'fixtures', name));
const d = (iso) => new Date(iso);

// ---------------------------------------------------------------- values

test('Values.parse keeps ids as text and converts plain numbers', () => {
  assert.equal(Values.parse('42'), 42);
  assert.equal(Values.parse('-3.5'), -3.5);
  assert.equal(Values.parse('00123'), '00123');
  assert.equal(Values.parse('+201001234567'), '+201001234567');
  assert.equal(Values.parse('1234567890123456789'), '1234567890123456789');
  assert.equal(Values.parse(''), null);
  assert.equal(Values.parse('true'), 'true');
  assert.equal(Values.parse('false', true), false);
  assert.deepEqual(Values.parse('2024-03-05', new Date(0)), d('2024-03-05T00:00:00Z'));
  assert.equal(Values.parse('2024-03-05'), '2024-03-05');
});

test('Values.format and columnName', () => {
  assert.equal(Values.format(d('2024-03-05T00:00:00Z')), '2024-03-05');
  assert.equal(Values.format(d('2024-03-05T13:45:00Z')), '2024-03-05 13:45');
  assert.equal(Values.format(d('2024-03-05T13:45:30Z')), '2024-03-05 13:45:30');
  assert.equal(Values.format(true), 'TRUE');
  assert.equal(Values.format(null), '');
  assert.deepEqual([0, 25, 26, 27, 701, 702].map(Values.columnName), ['A', 'Z', 'AA', 'AB', 'ZZ', 'AAA']);
});

test('Values.compare sorts numbers numerically, text naturally, empty last', () => {
  const sorted = [10, 'b', null, 2, 'a10', 'a2', ''].sort(Values.compare);
  assert.deepEqual(sorted, [2, 10, 'a2', 'a10', 'b', null, '']);
});

// ------------------------------------------------------------------- zip

test('Zip create/open round trip (compressed and stored files)', async () => {
  const big = 'مرحبا '.repeat(500);
  const bytes = await Zip.create([
    { name: 'a.txt', data: 'hello' },
    { name: 'dir/big.txt', data: big },
    { name: 'bin.dat', data: new Uint8Array([0, 1, 2, 255]) },
  ]);
  const zip = Zip.open(bytes);
  assert.deepEqual(zip.names(), ['a.txt', 'dir/big.txt', 'bin.dat']);
  assert.equal(await zip.text('a.txt'), 'hello');
  assert.equal(await zip.text('dir/big.txt'), big);
  assert.deepEqual(Array.from(await zip.bytes('bin.dat')), [0, 1, 2, 255]);
  assert.throws(() => Zip.open(new Uint8Array([1, 2, 3])), /not a valid/);
});

// ------------------------------------------------- reading real-world files

test('reads a file written by openpyxl (strings, numbers, booleans, dates, empty cells)', async () => {
  const { sheets } = await XlsxLite.read(fixture('openpyxl.xlsx'));
  assert.deepEqual(sheets.map((s) => s.name), ['Cars', 'ملاحظات']);
  const rows = sheets[0].rows;
  assert.deepEqual(rows[0], ['الماركة', 'الموديل', 'Year', 'Price', 'Available', 'Added', 'Updated']);
  assert.deepEqual(rows[1], ['هيونداي', 'كريتا', 2021, 15999.5, true, d('2024-03-05T00:00:00Z'), d('2024-03-05T13:45:30Z')]);
  assert.deepEqual(rows[2], ['Toyota', 'Corolla & Co <x> "q"', 2020, null, false, null, null]);
  assert.deepEqual(rows[3], new Array(7).fill(null));
  assert.deepEqual(rows[4], ['Zero', '00123', -7, 0.1, null, null, null]);
  assert.equal(rows[5][1], 'line1\nline2');
});

test('keeps the position of cells that do not start at A1', async () => {
  const { sheets } = await XlsxLite.read(fixture('openpyxl.xlsx'));
  const rows = sheets[1].rows;
  assert.equal(rows.length, 5);
  assert.equal(rows[2][1], 'offset cell');
  assert.equal(rows[4][3], 42);
});

test('reads a file written by xlsxwriter (rich text, formulas, number formats, dates)', async () => {
  const { sheets } = await XlsxLite.read(fixture('xlsxwriter.xlsx'));
  assert.deepEqual(sheets[0].rows, [
    ['Rich text', 0.25, 0.5, d('2000-01-31T00:00:00Z'), '  spaces  '],
    ['plain', 1234.5, 'ab', true, null],
  ]);
});

test('reads a file written by LibreOffice', async () => {
  const { sheets } = await XlsxLite.read(fixture('libreoffice.xlsx'));
  assert.deepEqual(sheets[0].rows, [['name', 'qty', 'note'], ['شاي', 3, 'a, b'], ['قهوة', 12.5, 'x']]);
});

test('rejects files that are not workbooks', async () => {
  await assert.rejects(XlsxLite.read(new Uint8Array([80, 75, 1, 2])), /not a valid/);
  const zipOnly = await Zip.create([{ name: 'hello.txt', data: 'x' }]);
  await assert.rejects(XlsxLite.read(zipOnly), /not an Excel workbook/);
});

// --------------------------------------------------------------- writing

test('write then read gives back the same values', async () => {
  const sheets = [
    {
      name: 'Data',
      rows: [
        ['name', 'qty', 'ok', 'when', 'note'],
        ['شاي & قهوة', 3, true, d('2024-03-05T00:00:00Z'), '<b>"x"</b>'],
        ['00123', 12.5, false, d('2024-03-05T13:45:30Z'), 'line1\nline2'],
        [null, null, null, null, ' padded '],
        ['x'.repeat(40000), -0.1, null, null, null],
      ],
    },
    { name: 'Empty', rows: [] },
  ];
  const back = (await XlsxLite.read(await XlsxLite.write(sheets, { rtl: true }))).sheets;

  assert.deepEqual(back.map((s) => s.name), ['Data', 'Empty']);
  const rows = back[0].rows;
  assert.deepEqual(rows.slice(0, 4), sheets[0].rows.slice(0, 4));
  assert.equal(rows[4][0].length, 32767); // Excel's cell limit
  assert.equal(rows[4][1], -0.1);
  assert.deepEqual(back[1].rows, []);
});

test('write cleans sheet names and keeps them unique', async () => {
  const rows = [['a']];
  const out = await XlsxLite.write([
    { name: 'Bad/Name:*?', rows },
    { name: 'bad name   ', rows },
    { name: 'x'.repeat(50), rows },
    { name: '', rows },
  ]);
  const names = (await XlsxLite.read(out)).sheets.map((s) => s.name);
  assert.equal(new Set(names.map((n) => n.toLowerCase())).size, 4);
  names.forEach((n) => { assert.ok(n.length <= 31); assert.doesNotMatch(n, /[\[\]:*?\/\\]/); });
});

test('write strips characters that are illegal in XML', async () => {
  const out = await XlsxLite.write([{ name: 'S', rows: [['a\u0000b\u0008c']] }]);
  assert.equal((await XlsxLite.read(out)).sheets[0].rows[0][0], 'abc');
});

// ------------------------------------------------------------------- csv

test('CSV parse: quotes, delimiters, newlines in cells, typed values', () => {
  const rows = Tabular.parseCsv('name,qty,note\r\n"a, ""b""",5,"line1\nline2"\r\nشاي,00123,\r\n');
  assert.deepEqual(rows, [
    ['name', 'qty', 'note'],
    ['a, "b"', 5, 'line1\nline2'],
    ['شاي', '00123', null],
  ]);
});

test('CSV delimiter detection (semicolon and tab)', () => {
  assert.deepEqual(Tabular.parseCsv('a;b;c\n1;2;3'), [['a', 'b', 'c'], [1, 2, 3]]);
  assert.deepEqual(Tabular.parseCsv('a\tb\n1\t2'), [['a', 'b'], [1, 2]]);
  assert.deepEqual(Tabular.parseCsv('"a;b",c\n1,2'), [['a;b', 'c'], [1, 2]]);
});

test('CSV round trip with special characters', () => {
  const rows = [['a', 'b'], ['x,y', 'say "hi"'], ['two\nlines', null], [1.5, true]];
  const back = Tabular.parseCsv(Tabular.toCsv(rows));
  assert.deepEqual(back, [['a', 'b'], ['x,y', 'say "hi"'], ['two\nlines', null], [1.5, 'TRUE']]);
});

test('CSV text decoding: BOM, UTF-8, and legacy Arabic windows-1256', () => {
  const utf8 = new TextEncoder().encode('﻿شاي');
  assert.equal(Tabular.decodeText(utf8), 'شاي');
  assert.equal(Tabular.decodeText(new TextEncoder().encode('قهوة')), 'قهوة');
  const cp1256 = new Uint8Array([0xD4, 0xC7, 0xED]); // "شاي" in windows-1256
  assert.equal(Tabular.decodeText(cp1256), 'شاي');
});

// ------------------------------------------------------------------ json

test('JSON import: array of objects, array of arrays, object of sheets', () => {
  assert.deepEqual(Tabular.parseJson('[{"a":1,"b":"x"},{"a":2,"c":true}]'),
    [{ name: 'Sheet1', rows: [['a', 'b', 'c'], [1, 'x', null], [2, null, true]] }]);
  assert.deepEqual(Tabular.parseJson('[[1,2],[3,4]]')[0].rows, [[1, 2], [3, 4]]);
  assert.deepEqual(Tabular.parseJson('{"One":[[1]],"Two":[{"k":"v"}]}').map((s) => s.name), ['One', 'Two']);
  assert.deepEqual(Tabular.parseJson('[{"a":{"n":1}}]')[0].rows[1], ['{"n":1}']);
  assert.throws(() => Tabular.parseJson('[1,2]'), /object, or every item/);
  assert.throws(() => Tabular.parseJson('{bad'), /Invalid JSON/);
});

test('JSON export uses the first row as keys and de-duplicates them', () => {
  const rows = [['name', 'name', ''], ['a', 'b', d('2024-03-05T00:00:00Z')], ['c']];
  assert.deepEqual(JSON.parse(Tabular.toJson(rows, true)), [
    { name: 'a', name_2: 'b', column_3: '2024-03-05' },
    { name: 'c', name_2: null, column_3: null },
  ]);
  assert.deepEqual(JSON.parse(Tabular.toJson([[1, 2]], false)), [[1, 2]]);
});

// ------------------------------------------------------------------ .xls

const XlsLite = require('../js/xls.js');

test('reads .xls files written by LibreOffice', () => {
  const book = XlsLite.read(fs.readFileSync(path.join(__dirname, 'fixtures/libreoffice-from-openpyxl.xls')));
  assert.equal(book.sheets.length, 2);
  const rows = book.sheets[0].rows;
  assert.deepEqual(rows[0], ['الماركة', 'الموديل', 'Year', 'Price', 'Available', 'Added', 'Updated']);
  assert.equal(rows[1][3], 15999.5);
  assert.equal(rows[1][5].toISOString(), '2024-03-05T00:00:00.000Z');
  assert.equal(rows[2][1], 'Corolla & Co <x> "q"');
  assert.equal(book.sheets[1].rows[2][1], 'offset cell');
});

test('reads the second .xls fixture', () => {
  const book = XlsLite.read(fs.readFileSync(path.join(__dirname, 'fixtures/libreoffice-from-xlsxwriter.xls')));
  assert.ok(book.sheets.length >= 1);
  assert.ok(book.sheets[0].rows.length > 1);
});

test('.xls write/read round trip, including long text split across SST records', () => {
  const long = 'سلام '.repeat(4000);
  const wide = 'x'.repeat(30000);
  const rows = [
    ['Name', 'Qty', 'When', 'Ok', 'Note'],
    ['علي', 1.5, d('2024-03-05T00:00:00Z'), true, 'é'],
    ['b', -2, d('2024-03-05T13:45:00Z'), false, wide],
    [null, '', null, null, long],
  ];
  for (let i = 0; i < 2000; i++) rows.push(['row' + i, i, null, null, 't' + (i % 7)]);
  const bytes = XlsLite.write([{ name: 'One', rows }, { name: 'One', rows: [['a'], [1]] }], {});
  assert.ok(XlsLite.isXls(bytes));
  const back = XlsLite.read(bytes);
  assert.deepEqual(back.sheets.map((s) => s.name), ['One', 'One (2)']);
  const r = back.sheets[0].rows;
  assert.equal(r.length, rows.length);
  assert.deepEqual(r[1], rows[1]);
  assert.equal(r[2][4], wide);
  assert.equal(r[3][4], long);
  assert.equal(r[2][2].toISOString(), '2024-03-05T13:45:00.000Z');
  assert.deepEqual(r[r.length - 1], rows[rows.length - 1]);
});

test('.xls writer rejects sheets beyond the format limits and readers reject junk', () => {
  assert.throws(() => XlsLite.write([{ name: 'S', rows: new Array(65537).fill(['a']) }]), /65,536/);
  assert.throws(() => XlsLite.write([{ name: 'S', rows: [new Array(257).fill('a')] }]), /256/);
  assert.equal(XlsLite.isXls(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8])), false);
  assert.throws(() => XlsLite.read(new Uint8Array([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0])));
});
