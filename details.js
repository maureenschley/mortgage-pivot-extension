'use strict';

// Loan Details: click a pivot cell, see the loans behind it, choose which columns to show.
// The column choice is saved in the string parameter "Selected Columns" (per user via custom views on Server/Cloud).

const PIVOT_SHEETS = ['Pivot (with totals)', 'Pivot (no totals)'];
const PARAM_NAME = 'Selected Columns';
const PARAM_VERSION = 1;
const PAGE_ROWS = 100;
// Row limiter: the workbook parameter below sets how many rows can be shown and exported (steps are defined on the parameter).
const ROWS_PARAM_NAME = 'Number of Rows Shown';
const DEFAULT_ROW_LIMIT = 10000;
const ROW_CEILING = 100000;

// Shown until the user picks their own columns.
const DEFAULT_COLUMNS = [
  'Loan ID', 'Loan Amount', 'Loan Type', 'Branch', 'Pipeline Status',
  'Loan Officer', 'Property State', 'Application Date', 'Scheduled Closing Date',
];
// Offered in the picker but never on by default.
const SENSITIVE = /first name|last name|email|phone|street address/i;
// Workbook helper fields that come back with the data but are not loan columns.
const NOT_LOAN_FIELDS = /^(Cell:|Eff |Row Attribute|Column Attribute|Stage Sort|AGG\(|Header|Report|Help|Today|Loan Count$|Value$|Show Totals Sheet$|Show No-Totals Sheet$)/i;

const $ = (id) => document.getElementById(id);
let dashboard = null;
let columnsParam = null;
let rowsParam = null;
let rowLimit = DEFAULT_ROW_LIMIT;
let showRowNumber = true;
let reader = null;
let token = 0;            // newest selection wins
let allColumns = [];      // loan column names, in data order
let chosen = [...DEFAULT_COLUMNS];   // columns to show, in display order
let page = { index: 0, data: [], columns: [] };
let total = 0;
let saveTimer = null;
let currentWs = null;     // worksheet the current selection came from
let exporting = false;
let cancelExport = false;

function log(msg) {
  // Diagnostics are disabled in the hosted build.
}

function setStatus(msg) {
  $('status').textContent = msg;
}

// ---- saved column choice ----------------------------------------------------

function loadSaved() {
  if (!columnsParam) return;
  try {
    const saved = JSON.parse(columnsParam.currentValue.value);
    if (saved && saved.v === PARAM_VERSION && Array.isArray(saved.cols) && saved.cols.length) {
      chosen = [...saved.cols];
      if (typeof saved.rowNum === 'boolean') showRowNumber = saved.rowNum;
    }
  } catch (e) { /* empty or old-format value: keep defaults */ }
}

function scheduleSave() {
  if (!columnsParam) return;
  clearTimeout(saveTimer);
  saveTimer = setTimeout(async () => {
    try {
      await columnsParam.changeValueAsync(JSON.stringify({ v: PARAM_VERSION, cols: chosen, rowNum: showRowNumber }));
    } catch (e) {
      log(`ERROR saving column choice: ${e.message || e}`);
    }
  }, 400);
}

// ---- column picker ----------------------------------------------------------

function renderPicker() {
  const filter = $('col-filter').value.trim().toLowerCase();
  const list = $('col-list');
  list.replaceChildren();
  allColumns.filter((c) => c.toLowerCase().includes(filter)).forEach((name) => {
    const label = document.createElement('label');
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.checked = chosen.includes(name);
    box.addEventListener('change', () => {
      if (box.checked) chosen.push(name); else chosen = chosen.filter((c) => c !== name);
      renderOrder();
      onColumnsChanged();
    });
    label.append(box, ' ', name);
    if (SENSITIVE.test(name)) {
      const tag = document.createElement('span');
      tag.className = 'tag';
      tag.textContent = 'personal data';
      label.append(' ', tag);
    }
    list.appendChild(label);
  });
  renderOrder();
  $('col-count').textContent = `${visibleColumns().length} of ${allColumns.length} columns shown`;
}

// Shown columns as pills: drag to reorder (or focus a pill and press Alt+Left / Alt+Right), x to remove.
// Dragging is done with pointer events, not the browser's drag-and-drop: in Tableau Desktop a browser drag that
// is dropped anywhere the page does not claim is handed to Tableau, which tries to load it as a data file.
let drag = null;   // { name, ghost, startX, startY, active, target, after }

function moveColumn(name, targetName, after) {
  if (name === targetName) return;
  chosen = chosen.filter((c) => c !== name);
  const at = chosen.indexOf(targetName);
  chosen.splice(at < 0 ? chosen.length : at + (after ? 1 : 0), 0, name);
}

function clearDropMarks() {
  document.querySelectorAll('.pill.drop-before, .pill.drop-after, .pill.dragging').forEach((el) => el.classList.remove('drop-before', 'drop-after', 'dragging'));
}

function endDrag(commit) {
  if (!drag) return;
  const d = drag;
  drag = null;
  if (d.ghost) d.ghost.remove();
  clearDropMarks();
  document.body.classList.remove('is-dragging');
  if (!d.active || !commit) return;
  if (d.target) moveColumn(d.name, d.target, d.after);
  else if (d.overList) moveColumn(d.name, visibleColumns().slice(-1)[0], true);
  else return;
  renderOrder();
  onColumnsChanged();
}

function onDragMove(e) {
  if (!drag) return;
  if (e.buttons === 0 && e.pointerType === 'mouse') { endDrag(false); return; }   // button released outside the page
  if (!drag.active) {
    if (Math.hypot(e.clientX - drag.startX, e.clientY - drag.startY) < 5) return;
    drag.active = true;
    document.body.classList.add('is-dragging');
    const ghost = document.createElement('div');
    ghost.className = 'pill ghost';
    ghost.textContent = drag.name;
    document.body.appendChild(ghost);
    drag.ghost = ghost;
    const src = [...document.querySelectorAll('#col-order .pill')].find((p) => p.dataset.name === drag.name);
    if (src) src.classList.add('dragging');
  }
  e.preventDefault();
  drag.ghost.style.left = `${e.clientX + 8}px`;
  drag.ghost.style.top = `${e.clientY + 8}px`;
  clearDropMarks();
  const src = [...document.querySelectorAll('#col-order .pill')].find((p) => p.dataset.name === drag.name);
  if (src) src.classList.add('dragging');
  const hit = document.elementFromPoint(e.clientX, e.clientY);
  const pill = hit && hit.closest ? hit.closest('#col-order .pill') : null;
  drag.target = null;
  drag.overList = !!(hit && hit.closest && hit.closest('#col-order')) && !pill;
  if (pill && pill.dataset.name !== drag.name) {
    const r = pill.getBoundingClientRect();
    drag.target = pill.dataset.name;
    drag.after = e.clientX > r.left + r.width / 2;
    pill.classList.add(drag.after ? 'drop-after' : 'drop-before');
  }
}

function renderOrder(focusName) {
  const list = $('col-order');
  list.replaceChildren();
  const shown = visibleColumns();
  if (!shown.length) {
    list.textContent = 'No columns chosen.';
    return;
  }
  shown.forEach((name, k) => {
    const pill = document.createElement('span');
    pill.className = 'pill';
    pill.tabIndex = 0;
    pill.dataset.name = name;
    pill.title = 'Drag to reorder';
    const label = document.createElement('span');
    label.textContent = name;
    const x = document.createElement('button');
    x.textContent = '✕';
    x.title = 'Remove';
    x.addEventListener('click', () => {
      chosen = chosen.filter((c) => c !== name);
      renderPicker();
      onColumnsChanged();
    });
    pill.append(label, x);

    pill.addEventListener('pointerdown', (e) => {
      if (e.button !== 0 || e.target.closest('button')) return;
      drag = { name, ghost: null, startX: e.clientX, startY: e.clientY, active: false, target: null, after: false, overList: false };
    });
    pill.addEventListener('keydown', (e) => {
      if (!e.altKey || (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight')) return;
      const target = shown[k + (e.key === 'ArrowLeft' ? -1 : 1)];
      if (!target) return;
      e.preventDefault();
      moveColumn(name, target, e.key === 'ArrowRight');
      renderOrder(name);
      onColumnsChanged();
    });
    list.appendChild(pill);
  });
  if (focusName) {
    const el = [...list.querySelectorAll('.pill')].find((p) => p.dataset.name === focusName);
    if (el) el.focus();
  }
}

function setupOrderDrop() {
  document.addEventListener('pointermove', onDragMove);
  document.addEventListener('pointerup', () => endDrag(true));
  document.addEventListener('pointercancel', () => endDrag(false));
  // Belt and braces: swallow any browser drag-and-drop dropped on this page so Tableau never receives it.
  ['dragover', 'drop'].forEach((t) => document.addEventListener(t, (e) => e.preventDefault()));
}

function onColumnsChanged() {
  $('col-count').textContent = `${visibleColumns().length} of ${allColumns.length} columns shown`;
  renderTable();
  scheduleSave();
}

function setChosen(names) {
  chosen = [...names];
  renderPicker();
  onColumnsChanged();
}

function visibleColumns() {
  return chosen.filter((c) => allColumns.includes(c));
}

// ---- table ------------------------------------------------------------------

// Rows that may be shown and exported for the current selection: the smaller of the selection, the parameter, and the ceiling.
function allowedRows() {
  return Math.min(total, rowLimit, ROW_CEILING);
}

function renderTable() {
  const grid = $('grid');
  grid.replaceChildren();
  if (!page.columns.length) return;
  const allowed = allowedRows();
  const base = page.index * PAGE_ROWS;
  const rows = page.data.slice(0, Math.max(0, allowed - base));
  const idx = visibleColumns().map((c) => page.columns.findIndex((pc) => pc.fieldName === c)).filter((i) => i >= 0);
  const head = grid.insertRow();
  const addHead = (text) => {
    const th = document.createElement('th');
    th.textContent = text;
    head.appendChild(th);
  };
  if (showRowNumber) addHead('Row Number');
  idx.forEach((i) => addHead(page.columns[i].fieldName));
  rows.forEach((row, r) => {
    const tr = grid.insertRow();
    if (showRowNumber) tr.insertCell().textContent = (base + r + 1).toLocaleString();
    idx.forEach((i) => { tr.insertCell().textContent = row[i].formattedValue; });
  });
  const first = base + 1;
  const last = base + rows.length;
  // One line states the count; when the selection is cut off it also says what to do about it.
  const advice = allowed >= ROW_CEILING ? ' Please contact us to discuss your needs.' : ' Choose a higher Number of Rows Shown to see more.';
  $('summary').textContent = total > allowed
    ? `Showing ${first.toLocaleString()}-${last.toLocaleString()} of ${allowed.toLocaleString()} (selection has ${total.toLocaleString()}).${advice}`
    : `${total.toLocaleString()} loans. Showing ${first.toLocaleString()}-${last.toLocaleString()}.`;
  $('prev').disabled = page.index === 0;
  $('next').disabled = last >= allowed;
}

async function goToPage(i) {
  const mine = token;
  const p = await reader.getPageAsync(i);
  if (mine !== token) return;
  page = { index: i, data: p.data, columns: p.columns };
  renderTable();
}

// ---- selection --------------------------------------------------------------

async function releaseReader() {
  const old = reader;
  reader = null;
  if (old) { try { await old.releaseAsync(); } catch (e) { /* already released */ } }
}

function clearDetails(message) {
  page = { index: 0, data: [], columns: [] };
  total = 0;
  $('grid').replaceChildren();
  $('summary').textContent = message;
  $('prev').disabled = true;
  $('next').disabled = true;
  currentWs = null;
  $('export-toggle').disabled = true;
  $('export-menu').hidden = true;
}

// ---- export -----------------------------------------------------------------

async function runExport(kind) {
  if (exporting || !currentWs) return;
  exporting = true;
  cancelExport = false;
  const buttons = document.querySelectorAll('.choice');
  buttons.forEach((b) => { b.disabled = true; });
  $('export-progress').hidden = false;
  $('copy-box').hidden = true;
  $('export-cancel').hidden = false;
  $('export-text').textContent = 'Starting…';
  const started = performance.now();
  try {
    const result = await window.LoanExport.run(kind, {
      ws: currentWs,
      columns: visibleColumns(),
      rowNumber: showRowNumber,
      limit: Math.min(rowLimit, ROW_CEILING),
      isCancelled: () => cancelExport,
      onProgress: (done, all) => { $('export-text').textContent = `Reading loans… ${done.toLocaleString()} of ${all.toLocaleString()}`; },
      log,
      showManualCopy: (text) => {
        $('copy-box').hidden = false;
        $('copy-box').value = text;
        $('copy-box').focus();
        $('copy-box').select();
      },
    });
    const rows = result.rows.toLocaleString();
    const cut = result.total > result.rows;
    const ofTotal = cut ? ` of ${result.total.toLocaleString()}` : '';
    const capNote = cut && result.rows >= ROW_CEILING ? ' Please contact us to discuss your needs.' : '';
    const partial = kind === 'copy' && cut
      ? ` Only the first ${rows}${ofTotal} loans were copied (limited by Number of Rows Shown and the copy limit).${capNote}` : '';
    $('export-text').textContent = kind === 'copy'
      ? (result.manual ? `The browser blocked automatic copying. The text is selected below: press Ctrl+C (Cmd+C on Mac), then paste into Excel.${partial}`
        : `Copied ${rows} loans. Paste into Excel.${partial}`)
      : `Saved ${result.filename} (${rows}${ofTotal} loans${cut ? ', limited by Number of Rows Shown' : ''}).${capNote} If no download appeared, use Copy to clipboard.`;
    log(`Export ${kind}${result.manual ? ' (manual copy)' : ''} took ${((performance.now() - started) / 1000).toFixed(1)} s: ${result.rows} rows, ${visibleColumns().length} columns${result.filename ? ', ' + result.filename : ''}`);
  } catch (e) {
    $('export-text').textContent = e.message === 'cancelled' ? 'Export cancelled.' : `Export failed: ${e.message || e}`;
    if (e.message !== 'cancelled') log(`ERROR exporting ${kind}: ${e.message || e}`);
  } finally {
    exporting = false;
    buttons.forEach((b) => { b.disabled = false; });
    $('export-cancel').hidden = true;
  }
}

async function onSelection(ws) {
  const mine = ++token;
  try {
    const marks = await ws.getSelectedMarksAsync();
    if (mine !== token) return;
    const markCount = marks.data.reduce((n, t) => n + t.data.length, 0);
    await releaseReader();
    if (markCount === 0) {
      // Nothing selected: fetch nothing (the data call would otherwise cover every loan).
      clearDetails('Click a number in the pivot to see its loans.');
      return;
    }
    $('summary').textContent = 'Loading…';
    const tables = await ws.getUnderlyingTablesAsync();
    if (mine !== token) return;
    const r = await ws.getUnderlyingTableDataReaderAsync(tables[0].id, PAGE_ROWS, { ignoreSelection: false, includeAllColumns: true, maxRows: 0 });
    if (mine !== token) { await r.releaseAsync(); return; }
    reader = r;
    total = r.totalRowCount;
    const first = await r.getPageAsync(0);
    if (mine !== token) return;
    if (!allColumns.length) {
      allColumns = first.columns.map((c) => c.fieldName).filter((n) => !NOT_LOAN_FIELDS.test(n));
      log(`Loan columns offered (${allColumns.length}): ${allColumns.join(' | ')}`);
      log(`Helper fields filtered out: ${first.columns.map((c) => c.fieldName).filter((n) => NOT_LOAN_FIELDS.test(n)).join(' | ')}`);
      renderPicker();
    }
    page = { index: 0, data: first.data, columns: first.columns };
    currentWs = ws;
    $('export-toggle').disabled = false;
    renderTable();
  } catch (e) {
    log(`ERROR loading selection: ${e.message || e}`);
    clearDetails('Could not load the loans for this selection.');
  }
}

// ---- row limit parameter ----------------------------------------------------

function readRowLimit() {
  if (!rowsParam) return;
  const raw = rowsParam.currentValue.value;
  const v = Number(String(raw).replace(/[,\s]/g, ''));
  log(`Number of Rows Shown raw value: ${JSON.stringify(raw)} (${typeof raw}) -> ${v}`);
  if (Number.isFinite(v) && v > 0) rowLimit = v;
}

// The parameter object read at startup keeps its old value, so fetch a fresh copy from the change event.
async function onRowLimitChanged(evt) {
  try {
    rowsParam = await evt.getParameterAsync();
  } catch (e) {
    log(`Could not re-read Number of Rows Shown: ${e.message || e}`);
  }
  readRowLimit();
  log(`Number of Rows Shown is now ${rowLimit}`);
  if (reader && page.index * PAGE_ROWS >= allowedRows()) await goToPage(0); else renderTable();
}

// ---- startup ----------------------------------------------------------------

async function start() {
  try {
    await tableau.extensions.initializeAsync();
  } catch (e) {
    setStatus('Could not start (is this running inside Tableau?)');
    return;
  }
  dashboard = tableau.extensions.dashboardContent.dashboard;
  const pivots = dashboard.worksheets.filter((w) => PIVOT_SHEETS.includes(w.name));
  if (!pivots.length) { setStatus('No pivot worksheets found on this dashboard'); return; }
  pivots.forEach((ws) => ws.addEventListener(tableau.TableauEventType.MarkSelectionChanged, () => onSelection(ws)));

  const params = await dashboard.getParametersAsync();
  columnsParam = params.find((p) => p.name === PARAM_NAME) || null;
  rowsParam = params.find((p) => p.name === ROWS_PARAM_NAME) || null;
  loadSaved();
  $('row-number').checked = showRowNumber;
  readRowLimit();
  if (rowsParam) rowsParam.addEventListener(tableau.TableauEventType.ParameterChanged, onRowLimitChanged);
  setStatus([
    columnsParam ? '' : `Column choices will not be saved (no "${PARAM_NAME}" parameter).`,
    rowsParam ? '' : `No "${ROWS_PARAM_NAME}" parameter found: showing up to ${DEFAULT_ROW_LIMIT.toLocaleString()} rows.`,
  ].filter(Boolean).join(' '));

  setupOrderDrop();
  $('col-filter').addEventListener('input', renderPicker);
  $('col-defaults').addEventListener('click', () => setChosen(DEFAULT_COLUMNS));
  $('col-all').addEventListener('click', () => setChosen(allColumns.filter((c) => !SENSITIVE.test(c))));
  $('col-none').addEventListener('click', () => setChosen([]));
  $('row-number').addEventListener('change', () => { showRowNumber = $('row-number').checked; renderTable(); scheduleSave(); });
  $('col-toggle').addEventListener('click', () => { $('picker').hidden = !$('picker').hidden; });
  $('export-toggle').addEventListener('click', () => { $('export-menu').hidden = !$('export-menu').hidden; });
  document.querySelectorAll('.choice').forEach((b) => b.addEventListener('click', () => runExport(b.dataset.kind)));
  $('export-cancel').addEventListener('click', () => { cancelExport = true; });
  $('prev').addEventListener('click', () => goToPage(page.index - 1));
  $('next').addEventListener('click', () => goToPage(page.index + 1));
  clearDetails('Click a number in the pivot to see its loans.');
}

document.addEventListener('DOMContentLoaded', start);
