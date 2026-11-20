/* Excel Table: user interface. Plain JavaScript, no build step, works from file:// and on GitHub Pages. */
(function () {
  'use strict';

  const { t } = I18n;
  const PAGE_SIZE = 100;
  const MAX_BYTES = 100 * 1024 * 1024;
  const $ = (id) => document.getElementById(id);

  const els = {
    title: $('title'), openBtn: $('open-btn'), langBtn: $('lang-btn'),
    welcome: $('welcome'), workspace: $('workspace'), drop: $('drop'), sampleBtn: $('sample-btn'),
    cellRef: $('cell-ref'), cellInput: $('cell-input'),
    search: $('search'), addRow: $('add-row'), exportBtn: $('export-btn'), headerToggle: $('header-toggle'),
    tabs: $('tabs'), gridWrap: $('grid-wrap'), grid: $('grid'),
    prev: $('prev'), next: $('next'), pagerInfo: $('pager-info'),
    sheet: $('sheet'), sheetTitle: $('sheet-title'), sheetItems: $('sheet-items'), sheetClose: $('sheet-close'),
    toast: $('toast'), toastText: $('toast-text'), toastAction: $('toast-action'),
    busy: $('busy'), fileInput: $('file-input'),
  };

  const state = {
    sheets: [],        // [{ name, rows, sort: null | { col, dir } }]
    active: 0,
    fileName: '',
    headerRow: true,
    query: '',
    page: 0,
    sel: null,         // { r, c } in sheet.rows coordinates
    dirty: false,
    version: 0,        // bumped when the order or the filter result can change
  };

  // ------------------------------------------------------------- utilities

  const current = () => state.sheets[state.active];
  const columnCount = (sheet) => sheet.rows[0].length;

  function esc(s) {
    return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  }

  function display(value) {
    const text = Values.format(value);
    return text.length > 300 ? text.slice(0, 300) + '…' : text;
  }

  function store(key, value) {
    try { if (value === undefined) return localStorage.getItem(key); localStorage.setItem(key, value); } catch (e) { /* private mode */ }
    return null;
  }

  function looksRtl(rows) {
    let letters = 0;
    let arabic = 0;
    for (let r = 0; r < Math.min(rows.length, 50); r++) {
      for (const v of rows[r]) {
        if (typeof v !== 'string') continue;
        const l = (v.match(/\p{L}/gu) || []).length;
        letters += l;
        arabic += (v.match(/[֐-ࣿ]/g) || []).length;
      }
    }
    return letters > 0 && arabic / letters > 0.5;
  }

  // ----------------------------------------------------------------- toast

  let toastTimer = 0;
  function toast(message, options) {
    options = options || {};
    clearTimeout(toastTimer);
    els.toastText.textContent = message;
    els.toast.classList.toggle('error', !!options.error);
    els.toastAction.hidden = !options.action;
    if (options.action) {
      els.toastAction.textContent = options.action.label;
      els.toastAction.onclick = () => { hideToast(); options.action.run(); };
    }
    els.toast.hidden = false;
    toastTimer = setTimeout(hideToast, options.action ? 7000 : options.error ? 7000 : 3500);
  }
  function hideToast() { clearTimeout(toastTimer); els.toast.hidden = true; }

  // ------------------------------------------------------------ view model

  let viewCache = null;

  function rowMatches(row, q) {
    for (const v of row) if (v !== null && Values.format(v).toLocaleLowerCase().includes(q)) return true;
    return false;
  }

  /** Row indexes (into sheet.rows) in the order they are shown: filtered by search, then sorted. Excludes the header row. */
  function getView() {
    const sheet = current();
    const key = [state.active, state.version, state.query, state.headerRow, sheet.sort ? sheet.sort.col + ':' + sheet.sort.dir : ''].join('|');
    if (viewCache && viewCache.key === key) return viewCache.list;

    const q = state.query.trim().toLocaleLowerCase();
    const list = [];
    for (let r = state.headerRow ? 1 : 0; r < sheet.rows.length; r++) {
      if (!q || rowMatches(sheet.rows[r], q)) list.push(r);
    }
    if (sheet.sort) {
      const { col, dir } = sheet.sort;
      list.sort((a, b) => dir * Values.compare(sheet.rows[a][col], sheet.rows[b][col]) || a - b);
    }
    viewCache = { key, list };
    return list;
  }

  function changed(structural) {
    state.dirty = true;
    if (structural || state.query || current().sort) state.version++;
  }

  // ------------------------------------------------------------- rendering

  function cellHtml(value, r, c) {
    return '<td data-r="' + r + '" data-c="' + c + '" dir="auto"' + (typeof value === 'number' ? ' class="num"' : '') + '>' + esc(display(value)) + '</td>';
  }

  function renderBar() {
    const open = state.sheets.length > 0;
    els.openBtn.hidden = !open;
    els.langBtn.textContent = I18n.lang === 'ar' ? 'EN' : 'ع';
    els.title.textContent = open ? state.fileName : t('app.title');
    document.title = open ? state.fileName + ' – ' + t('app.title') : t('app.title');
    els.welcome.hidden = open;
    els.workspace.hidden = !open;
  }

  function renderTabs() {
    els.tabs.hidden = state.sheets.length < 2;
    els.tabs.innerHTML = '';
    state.sheets.forEach((sheet, i) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'tab';
      b.dir = 'auto';
      b.textContent = sheet.name;
      b.setAttribute('aria-current', i === state.active ? 'true' : 'false');
      b.addEventListener('click', () => switchSheet(i));
      els.tabs.append(b);
    });
  }

  function renderGrid() {
    const sheet = current();
    const cols = columnCount(sheet);
    const view = getView();
    const pages = Math.max(1, Math.ceil(view.length / PAGE_SIZE));
    state.page = Math.min(state.page, pages - 1);
    const slice = view.slice(state.page * PAGE_SIZE, (state.page + 1) * PAGE_SIZE);

    let html = '<thead><tr><th class="corner" scope="col"></th>';
    for (let c = 0; c < cols; c++) {
      const name = Values.columnName(c);
      const arrow = sheet.sort && sheet.sort.col === c ? (sheet.sort.dir === 1 ? '▲' : '▼') : '';
      html += '<th class="col" scope="col" data-col="' + c + '"><button type="button" class="hbtn" aria-label="' + esc(t('col.aria', { name })) + '">'
        + name + (arrow ? '<span class="sort" aria-hidden="true">' + arrow + '</span>' : '') + '</button></th>';
    }
    html += '</tr>';
    if (state.headerRow) {
      html += '<tr class="hdr"><th class="rowno" scope="row" data-row="0"><button type="button" class="hbtn" aria-label="' + esc(t('row.aria', { n: 1 })) + '">1</button></th>';
      for (let c = 0; c < cols; c++) html += cellHtml(sheet.rows[0][c], 0, c);
      html += '</tr>';
    }
    html += '</thead><tbody>';
    for (const r of slice) {
      html += '<tr><th class="rowno" scope="row" data-row="' + r + '"><button type="button" class="hbtn" aria-label="' + esc(t('row.aria', { n: r + 1 })) + '">' + (r + 1) + '</button></th>';
      for (let c = 0; c < cols; c++) html += cellHtml(sheet.rows[r][c], r, c);
      html += '</tr>';
    }
    html += '</tbody>';
    els.grid.innerHTML = html;
    paintSelection();
  }

  function renderPager() {
    const total = getView().length;
    const pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
    if (total === 0) {
      els.pagerInfo.textContent = t('pager.none');
    } else {
      const from = state.page * PAGE_SIZE + 1;
      const to = Math.min(total, (state.page + 1) * PAGE_SIZE);
      els.pagerInfo.textContent = t('pager.info', { from: from.toLocaleString('en'), to: to.toLocaleString('en'), total: total.toLocaleString('en') });
    }
    els.prev.disabled = state.page <= 0;
    els.next.disabled = state.page >= pages - 1;
  }

  function renderAll() {
    renderBar();
    if (!state.sheets.length) return;
    renderTabs();
    renderGrid();
    renderPager();
    renderFormula();
  }

  // ------------------------------------------------------------- selection

  function cellElement(r, c) {
    return els.grid.querySelector('td[data-r="' + r + '"][data-c="' + c + '"]');
  }

  function paintSelection() {
    els.grid.querySelectorAll('.sel, .active').forEach((el) => el.classList.remove('sel', 'active'));
    if (!state.sel) return;
    const { r, c } = state.sel;
    const td = cellElement(r, c);
    if (td) td.classList.add('sel');
    const col = els.grid.querySelector('th.col[data-col="' + c + '"]');
    const row = els.grid.querySelector('th.rowno[data-row="' + r + '"]');
    if (col) col.classList.add('active');
    if (row) row.classList.add('active');
  }

  function renderFormula() {
    if (!state.sel) {
      els.cellRef.textContent = '';
      els.cellInput.value = '';
      els.cellInput.disabled = true;
      return;
    }
    const { r, c } = state.sel;
    els.cellRef.textContent = Values.columnName(c) + (r + 1);
    els.cellInput.disabled = false;
    els.cellInput.value = Values.format(current().rows[r][c]);
  }

  /** Save what is typed in the formula bar into the selected cell. */
  function commit() {
    if (!state.sel) return;
    const { r, c } = state.sel;
    const sheet = current();
    const old = sheet.rows[r][c];
    const text = els.cellInput.value;
    if (text === Values.format(old)) return;

    const value = Values.parse(text, old);
    sheet.rows[r][c] = value;
    changed(false);
    const td = cellElement(r, c);
    if (td) {
      td.textContent = display(value);
      td.classList.toggle('num', typeof value === 'number');
    }
  }

  function select(r, c, focusInput) {
    commit();
    state.sel = { r, c };
    paintSelection();
    renderFormula();
    const td = cellElement(r, c);
    if (td) td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    if (focusInput !== false) {
      els.cellInput.focus({ preventScroll: true });
      const end = els.cellInput.value.length;
      els.cellInput.setSelectionRange(end, end);
    }
  }

  function move(dr, dc) {
    if (!state.sel) return;
    commit();
    let { r, c } = state.sel;
    c += dc;
    if (c < 0 || c >= columnCount(current())) return;
    if (dr) {
      const view = getView();
      const order = state.headerRow ? [0].concat(view) : view;
      const i = order.indexOf(r) + dr;
      if (order.indexOf(r) < 0 || i < 0 || i >= order.length) return;
      r = order[i];
      const pos = view.indexOf(r);
      if (pos >= 0 && Math.floor(pos / PAGE_SIZE) !== state.page) {
        state.page = Math.floor(pos / PAGE_SIZE);
        renderGrid();
        renderPager();
      }
    }
    select(r, c);
  }

  // ------------------------------------------------------------ sheet menu

  function openMenu(title, items) {
    els.sheetTitle.textContent = title;
    els.sheetItems.innerHTML = '';
    items.forEach((item) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'sheet-item' + (item.danger ? ' danger' : '');
      const label = document.createElement('span');
      label.className = 't';
      label.textContent = item.label;
      b.append(label);
      if (item.hint) {
        const hint = document.createElement('span');
        hint.className = 'h';
        hint.textContent = item.hint;
        b.append(hint);
      }
      b.disabled = !!item.disabled;
      b.addEventListener('click', () => { els.sheet.close(); item.run(); });
      els.sheetItems.append(b);
    });
    els.sheet.showModal();
  }

  function columnMenu(c) {
    const sheet = current();
    const sorted = sheet.sort && sheet.sort.col === c;
    openMenu(t('menu.col.title', { name: Values.columnName(c) }), [
      { label: t('menu.col.asc'), run: () => sortBy(c, 1) },
      { label: t('menu.col.desc'), run: () => sortBy(c, -1) },
      { label: t('menu.col.clear'), run: () => sortBy(c, 0), disabled: !sorted },
      { label: t('menu.col.insert'), run: () => insertColumn(c) },
      { label: t('menu.col.delete'), run: () => deleteColumn(c), danger: true, disabled: columnCount(sheet) <= 1 },
    ]);
  }

  function rowMenu(r) {
    openMenu(t('menu.row.title', { n: r + 1 }), [
      { label: t('menu.row.insert'), run: () => insertRow(r) },
      { label: t('menu.row.delete'), run: () => deleteRow(r), danger: true, disabled: current().rows.length <= 1 },
    ]);
  }

  // ------------------------------------------------------- data operations

  function emptyRow() {
    return new Array(columnCount(current())).fill(null);
  }

  /** Clear search and sort so a new row is visible, then show the page that contains it. */
  function revealRow(r) {
    state.query = '';
    els.search.value = '';
    current().sort = null;
    state.version++;
    const pos = r - (state.headerRow ? 1 : 0);
    state.page = pos < 0 ? 0 : Math.floor(pos / PAGE_SIZE);
    state.sel = null;
    renderGrid();
    renderPager();
    renderFormula();
    select(r, 0);
  }

  function addRow() {
    commit();
    const sheet = current();
    sheet.rows.push(emptyRow());
    changed(true);
    revealRow(sheet.rows.length - 1);
  }

  function insertRow(r) {
    commit();
    const sheet = current();
    const at = state.headerRow && r === 0 ? 1 : r + 1;
    sheet.rows.splice(at, 0, emptyRow());
    changed(true);
    revealRow(at);
  }

  function deleteRow(r) {
    commit();
    const sheet = current();
    const removed = sheet.rows.splice(r, 1)[0];
    state.sel = null;
    changed(true);
    renderAll();
    toast(t('toast.rowDeleted'), {
      action: {
        label: t('toast.undo'),
        run: () => {
          sheet.rows.splice(Math.min(r, sheet.rows.length), 0, removed);
          changed(true);
          if (sheet === current()) renderAll();
        },
      },
    });
  }

  function insertColumn(c) {
    commit();
    const sheet = current();
    sheet.rows.forEach((row) => row.splice(c + 1, 0, null));
    if (sheet.sort && sheet.sort.col > c) sheet.sort.col++;
    state.sel = null;
    changed(true);
    renderAll();
  }

  function deleteColumn(c) {
    commit();
    const sheet = current();
    const removed = sheet.rows.map((row) => row.splice(c, 1)[0]);
    const oldSort = sheet.sort ? { col: sheet.sort.col, dir: sheet.sort.dir } : null;
    if (sheet.sort) {
      if (sheet.sort.col === c) sheet.sort = null;
      else if (sheet.sort.col > c) sheet.sort.col--;
    }
    state.sel = null;
    changed(true);
    renderAll();
    toast(t('toast.colDeleted'), {
      action: {
        label: t('toast.undo'),
        run: () => {
          sheet.rows.forEach((row, i) => row.splice(c, 0, removed[i]));
          sheet.sort = oldSort;
          changed(true);
          if (sheet === current()) renderAll();
        },
      },
    });
  }

  function sortBy(col, dir) {
    commit();
    current().sort = dir === 0 ? null : { col, dir };
    state.version++;
    state.page = 0;
    renderGrid();
    renderPager();
  }

  function switchSheet(index) {
    commit();
    state.active = index;
    state.page = 0;
    state.sel = null;
    viewCache = null;
    renderAll();
  }

  // ------------------------------------------------------------ open files

  function normalize(sheets) {
    return sheets.map((sheet, i) => {
      const width = Math.max(1, sheet.rows.reduce((max, row) => Math.max(max, row.length), 0));
      const rows = sheet.rows.map((row) => {
        const out = new Array(width).fill(null);
        row.forEach((v, c) => { out[c] = v === undefined || v === '' ? null : v; });
        return out;
      });
      if (!rows.length) rows.push(new Array(width).fill(null));
      return { name: sheet.name || t('sheet.empty') + ' ' + (i + 1), rows, sort: null };
    });
  }

  function load(sheets, fileName) {
    state.sheets = normalize(sheets);
    state.active = 0;
    state.fileName = fileName;
    state.headerRow = true;
    els.headerToggle.checked = true;
    state.query = '';
    els.search.value = '';
    state.page = 0;
    state.sel = null;
    state.dirty = false;
    state.version++;
    viewCache = null;
    hideToast();
    renderAll();
    els.gridWrap.scrollTo(0, 0);
  }

  function baseName(name) {
    return name.replace(/\.[^.]+$/, '').replace(/[\\/:*?"<>|]+/g, '_').trim() || 'spreadsheet';
  }

  /** An error whose message is already written for the user (not a technical reason). */
  function friendly(message) {
    const error = new Error(message);
    error.friendly = true;
    return error;
  }

  async function openFile(file) {
    if (!file) return;
    if (file.size > MAX_BYTES) { toast(t('err.big'), { error: true }); return; }

    els.busy.hidden = false;
    try {
      await new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0))); // let the busy overlay paint
      const bytes = new Uint8Array(await file.arrayBuffer());
      const ext = (file.name.match(/\.([^.]+)$/) || [, ''])[1].toLowerCase();
      const isZip = bytes[0] === 0x50 && bytes[1] === 0x4B;
      const isOle = bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0;

      let sheets;
      if (isOle) throw friendly(t('err.xls'));
      if (isZip || ext === 'xlsx' || ext === 'xlsm') {
        sheets = (await XlsxLite.read(bytes)).sheets;
      } else if (['csv', 'tsv', 'txt', 'json'].includes(ext)) {
        const text = Tabular.decodeText(bytes);
        sheets = ext === 'json' ? Tabular.parseJson(text) : [{ name: baseName(file.name), rows: Tabular.parseCsv(text, ext === 'tsv' ? '\t' : undefined) }];
      } else {
        throw friendly(t('err.type'));
      }
      load(sheets, baseName(file.name));
    } catch (error) {
      toast(error.friendly ? error.message : t('err.open', { reason: error.message }), { error: true });
    } finally {
      els.busy.hidden = true;
      els.fileInput.value = '';
    }
  }

  // ---------------------------------------------------------------- export

  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    toast(t('toast.exported', { name }));
  }

  async function exportAs(kind) {
    commit();
    const sheet = current();
    const base = state.fileName;
    try {
      if (kind === 'xlsx') {
        const bytes = await XlsxLite.write(state.sheets.map((s) => ({ name: s.name, rows: s.rows })), { rtl: looksRtl(sheet.rows), header: state.headerRow });
        download(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), base + '.xlsx');
        return;
      }
      const suffix = state.sheets.length > 1 ? '-' + baseName(sheet.name) : '';
      if (kind === 'csv') {
        download(new Blob(['﻿' + Tabular.toCsv(sheet.rows)], { type: 'text/csv;charset=utf-8' }), base + suffix + '.csv');
      } else {
        download(new Blob([Tabular.toJson(sheet.rows, state.headerRow)], { type: 'application/json' }), base + suffix + '.json');
      }
    } catch (error) {
      toast(t('err.export', { reason: error.message }), { error: true });
    }
  }

  function exportMenu() {
    commit();
    const n = state.sheets.length;
    openMenu(t('export.title'), [
      { label: t('export.xlsx'), hint: n > 1 ? t('export.xlsx.hint', { n }) : t('export.xlsx.hint1'), run: () => exportAs('xlsx') },
      { label: t('export.csv'), hint: t('export.csv.hint'), run: () => exportAs('csv') },
      { label: t('export.json'), hint: state.headerRow ? t('export.json.hintHeader') : t('export.json.hintPlain'), run: () => exportAs('json') },
    ]);
  }

  // ---------------------------------------------------------------- events

  els.grid.addEventListener('click', (e) => {
    const td = e.target.closest('td[data-r]');
    if (td) { select(+td.dataset.r, +td.dataset.c); return; }
    const col = e.target.closest('th.col');
    if (col) { columnMenu(+col.dataset.col); return; }
    const row = e.target.closest('th.rowno');
    if (row) rowMenu(+row.dataset.row);
  });

  els.cellInput.addEventListener('change', commit);
  els.cellInput.addEventListener('keydown', (e) => {
    if (e.isComposing) return;
    if (e.key === 'Enter') { e.preventDefault(); move(e.shiftKey ? -1 : 1, 0); }
    else if (e.key === 'Tab') { e.preventDefault(); move(0, e.shiftKey ? -1 : 1); }
    else if (e.key === 'ArrowDown') { e.preventDefault(); move(1, 0); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); move(-1, 0); }
    else if (e.key === 'Escape' && state.sel) { renderFormula(); }
  });

  let searchTimer = 0;
  els.search.addEventListener('input', () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      commit();
      state.query = els.search.value;
      state.page = 0;
      state.version++;
      renderGrid();
      renderPager();
    }, 150);
  });

  els.headerToggle.addEventListener('change', () => {
    commit();
    state.headerRow = els.headerToggle.checked;
    state.page = 0;
    state.sel = null;
    state.version++;
    renderGrid();
    renderPager();
    renderFormula();
  });

  els.prev.addEventListener('click', () => { state.page--; renderGrid(); renderPager(); els.gridWrap.scrollTo(0, 0); });
  els.next.addEventListener('click', () => { state.page++; renderGrid(); renderPager(); els.gridWrap.scrollTo(0, 0); });
  els.addRow.addEventListener('click', addRow);
  els.exportBtn.addEventListener('click', exportMenu);

  els.drop.addEventListener('click', () => els.fileInput.click());
  els.openBtn.addEventListener('click', () => els.fileInput.click());
  els.fileInput.addEventListener('change', () => openFile(els.fileInput.files[0]));
  els.sampleBtn.addEventListener('click', () => {
    load([{ name: window.SAMPLE_SHEET.name, rows: window.SAMPLE_SHEET.rows.map((r) => r.slice()) }], 'car-brands');
  });

  els.sheetClose.addEventListener('click', () => els.sheet.close());
  els.sheet.addEventListener('click', (e) => { if (e.target === els.sheet) els.sheet.close(); });

  els.langBtn.addEventListener('click', () => {
    setLanguage(I18n.lang === 'ar' ? 'en' : 'ar');
  });

  // drag and drop a file anywhere on the page
  let dragDepth = 0;
  const hasFiles = (e) => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');
  window.addEventListener('dragenter', (e) => { if (hasFiles(e)) { dragDepth++; document.body.classList.add('dragging'); } });
  window.addEventListener('dragleave', (e) => { if (hasFiles(e) && --dragDepth <= 0) { dragDepth = 0; document.body.classList.remove('dragging'); } });
  window.addEventListener('dragover', (e) => { if (hasFiles(e)) e.preventDefault(); });
  window.addEventListener('drop', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    dragDepth = 0;
    document.body.classList.remove('dragging');
    openFile(e.dataTransfer.files[0]);
  });

  window.addEventListener('beforeunload', (e) => {
    if (state.dirty) { e.preventDefault(); e.returnValue = ''; }
  });

  // ------------------------------------------------------------------ boot

  function setLanguage(lang) {
    I18n.setLang(lang);
    store('excel-table-lang', I18n.lang);
    I18n.apply();
    renderAll();
  }

  const saved = store('excel-table-lang');
  setLanguage(saved === 'ar' || saved === 'en' ? saved : (navigator.language || '').toLowerCase().startsWith('ar') ? 'ar' : 'en');
})();
