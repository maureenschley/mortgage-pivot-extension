'use strict';

// Export of every loan behind the current pivot selection, in the columns the user picked.
//   xlsx : real Excel workbook (true dates, numbers with number formats, IDs/ZIPs as text, header styling)
//   csv  : plain values (dates as ISO yyyy-mm-dd so Excel cannot flip day and month), UTF-8 with BOM
//   copy : tab-separated values on the clipboard, same values as csv
// window.LoanExport.run(kind, { ws, columns, isCancelled, onProgress, log }) -> { rows, filename? }

(function () {
  const EXPORT_PAGE_ROWS = 2000;
  // Clipboard copies are limited: pasting hundreds of thousands of rows is impractical and the browser clipboard is unreliable at that size.
  const COPY_ROW_LIMIT = 10000;

  // Numeric fields that are identifiers, not quantities: exported as text, padded to a fixed width (0 = no padding).
  const IDENTIFIER_COLUMNS = { 'Property ZIP': 5, 'Loan Officer NMLS ID': 0 };

  const NUMERIC_TYPES = ['int', 'float'];
  const DATE_TYPES = ['date', 'date-time'];

  // ---- value helpers --------------------------------------------------------

  function isBlank(cell) {
    // Tableau sends empty cells as the text "%null%" (formatted "Null").
    const v = cell.value;
    return v === null || v === undefined || v === '' || v === '%null%' || cell.formattedValue === '%null%';
  }

  // Returns { y, m, d, hh, mm, ss } for a date cell, or null if it cannot be read as a date.
  function dateParts(cell) {
    const v = cell.value;
    if (typeof v === 'string') {
      const m = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}):(\d{2}))?/.exec(v);
      if (m) return { y: +m[1], m: +m[2], d: +m[3], hh: +(m[4] || 0), mm: +(m[5] || 0), ss: +(m[6] || 0) };
    }
    const n = cell.nativeValue;
    if (n instanceof Date && !isNaN(n)) {
      return { y: n.getFullYear(), m: n.getMonth() + 1, d: n.getDate(), hh: n.getHours(), mm: n.getMinutes(), ss: n.getSeconds() };
    }
    return null;
  }

  const pad2 = (n) => String(n).padStart(2, '0');
  const isoDate = (p) => `${p.y}-${pad2(p.m)}-${pad2(p.d)}`;
  const isoDateTime = (p) => `${isoDate(p)} ${pad2(p.hh)}:${pad2(p.mm)}:${pad2(p.ss)}`;

  function excelSerial(p) {
    const days = Math.floor(Date.UTC(p.y, p.m - 1, p.d) / 86400000) + 25569;
    return days + (p.hh * 3600 + p.mm * 60 + p.ss) / 86400;
  }

  function identifierText(name, cell) {
    const n = Number(cell.value);
    const s = Number.isFinite(n) ? String(Math.round(n)) : String(cell.value);
    return s.padStart(IDENTIFIER_COLUMNS[name] || 0, '0');
  }

  // Looks at the display text of one page to choose an Excel number format that matches the dashboard.
  function inferNumberFormat(formattedValues) {
    let currency = false, percent = false, thousands = false, decimals = 0;
    formattedValues.forEach((f) => {
      if (!f || f === '%null%') return;
      if (/^[-(]?\$/.test(f)) currency = true;
      if (/%/.test(f)) percent = true;
      if (/\d,\d{3}/.test(f)) thousands = true;
      const d = /\.(\d+)/.exec(f);
      if (d) decimals = Math.max(decimals, d[1].length);
    });
    return (currency ? '"$"' : '') + (thousands ? '#,##0' : '0') + (decimals ? '.' + '0'.repeat(decimals) : '') + (percent ? '%' : '');
  }

  function inferDateFormat(formattedValues, withTime) {
    const sample = formattedValues.find((f) => f && f !== '%null%') || '';
    const base = /^\d{4}-/.test(sample) ? 'yyyy-mm-dd' : /^\d{1,2}\/\d{1,2}\/\d{4}/.test(sample) ? (/^\d{2}\//.test(sample) ? 'mm/dd/yyyy' : 'm/d/yyyy') : 'mm/dd/yyyy';
    return withTime ? base + ' h:mm' : base;
  }

  // ---- reading every page ---------------------------------------------------

  // Calls consume(rows, firstPageColumns, pageNumber) for each page, in order. Stops after maxRows rows if given.
  // Returns { total: rows behind the selection, read: rows passed to consume }.
  async function readAll(ctx, consume, maxRows) {
    const tables = await ctx.ws.getUnderlyingTablesAsync();
    const reader = await ctx.ws.getUnderlyingTableDataReaderAsync(tables[0].id, EXPORT_PAGE_ROWS, { ignoreSelection: false, includeAllColumns: true, maxRows: 0 });
    try {
      const total = reader.totalRowCount;
      let done = 0;
      let columns = null;
      for (let i = 0; i < reader.pageCount; i++) {
        if (ctx.isCancelled()) throw new Error('cancelled');
        const page = await reader.getPageAsync(i);
        if (!columns) columns = page.columns;
        const left = maxRows ? maxRows - done : Infinity;
        const data = left < page.data.length ? page.data.slice(0, left) : page.data;
        await consume(data, columns, i);
        done += data.length;
        ctx.onProgress(done, maxRows ? Math.min(total, maxRows) : total);
        if (maxRows && done >= maxRows) break;
      }
      return { total, read: done };
    } finally {
      try { await reader.releaseAsync(); } catch (e) { /* already released */ }
    }
  }

  // With ctx.rowNumber a "Row Number" column (1, 2, 3, ...) is added at the far left; it has no source column (i = -1).
  function columnIndexes(ctx, columns) {
    const cols = ctx.columns.map((name) => {
      const i = columns.findIndex((c) => c.fieldName === name);
      if (i < 0) throw new Error(`Column "${name}" is not in the data`);
      return { name, i, type: columns[i].dataType };
    });
    if (ctx.rowNumber) cols.unshift({ name: 'Row Number', i: -1, type: 'rownum' });
    return cols;
  }

  // ---- plain values (csv, copy) --------------------------------------------

  function plainValue(col, cell) {
    if (isBlank(cell)) return '';
    if (col.name in IDENTIFIER_COLUMNS) return identifierText(col.name, cell);
    if (DATE_TYPES.includes(col.type)) {
      const p = dateParts(cell);
      if (p) return col.type === 'date-time' ? isoDateTime(p) : isoDate(p);
      return cell.formattedValue;
    }
    if (NUMERIC_TYPES.includes(col.type)) return String(cell.value);
    return String(cell.formattedValue);
  }

  async function runPlain(ctx, separator, quote, maxRows) {
    const out = [];
    let cols = null;
    const clean = quote
      ? (s) => (/[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s)
      : (s) => s.replace(/[\t\r\n]+/g, ' ');
    const eol = quote ? '\r\n' : '\n';
    let seen = 0;
    const counts = await readAll(ctx, (data, columns) => {
      if (!cols) {
        cols = columnIndexes(ctx, columns);
        out.push(cols.map((c) => clean(c.name)).join(separator) + eol);
      }
      const lines = new Array(data.length);
      for (let r = 0; r < data.length; r++) {
        lines[r] = cols.map((c) => (c.i < 0 ? String(seen + r + 1) : clean(plainValue(c, data[r][c.i])))).join(separator) + eol;
      }
      seen += data.length;
      out.push(lines.join(''));
    }, maxRows);
    return { parts: out, rows: counts.read, total: counts.total };
  }

  // ---- xlsx -----------------------------------------------------------------

  const enc = new TextEncoder();
  const xmlEscape = (s) => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  function colLetter(n) {
    let s = '';
    for (n++; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    return s;
  }

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crcUpdate(c, bytes) {
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return c;
  }
  const crc32 = (bytes) => (crcUpdate(0xffffffff, bytes) ^ 0xffffffff) >>> 0;

  // Compresses a stream of byte chunks to raw deflate as it arrives (so the full sheet is never held in memory).
  // Falls back to storing the bytes uncompressed where CompressionStream is unavailable.
  function makeDeflater() {
    const canCompress = typeof CompressionStream === 'function';
    const out = [];
    let crc = 0xffffffff;
    let size = 0;
    let writer = null;
    let pump = null;
    if (canCompress) {
      const cs = new CompressionStream('deflate-raw');
      writer = cs.writable.getWriter();
      const reader = cs.readable.getReader();
      pump = (async () => {
        for (;;) {
          const { done, value } = await reader.read();
          if (done) return;
          out.push(value);
        }
      })();
    }
    return {
      async write(bytes) {
        crc = crcUpdate(crc, bytes);
        size += bytes.length;
        if (writer) await writer.write(bytes); else out.push(bytes);
      },
      async finish(name) {
        if (writer) { await writer.close(); await pump; }
        const total = out.reduce((n, c) => n + c.length, 0);
        const data = new Uint8Array(total);
        let at = 0;
        out.forEach((c) => { data.set(c, at); at += c.length; });
        return { name, data, crc: (crc ^ 0xffffffff) >>> 0, size, method: writer ? 8 : 0 };
      },
    };
  }

  // Minimal zip writer (stored, no compression). files: [{ name, data: Uint8Array }]
  function zipStore(files) {
    const now = new Date();
    const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    const parts = [];
    const central = [];
    let offset = 0;
    files.forEach((f) => {
      const name = enc.encode(f.name);
      const crc = f.crc !== undefined ? f.crc : crc32(f.data);
      const rawSize = f.size !== undefined ? f.size : f.data.length;
      const method = f.method || 0;
      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true); local.setUint16(4, 20, true); local.setUint16(6, 0x0800, true);
      local.setUint16(8, method, true); local.setUint16(10, dosTime, true); local.setUint16(12, dosDate, true);
      local.setUint32(14, crc, true); local.setUint32(18, f.data.length, true); local.setUint32(22, rawSize, true);
      local.setUint16(26, name.length, true); local.setUint16(28, 0, true);
      parts.push(local.buffer, name, f.data);
      const cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, 0x02014b50, true); cd.setUint16(4, 20, true); cd.setUint16(6, 20, true); cd.setUint16(8, 0x0800, true);
      cd.setUint16(10, method, true); cd.setUint16(12, dosTime, true); cd.setUint16(14, dosDate, true);
      cd.setUint32(16, crc, true); cd.setUint32(20, f.data.length, true); cd.setUint32(24, rawSize, true);
      cd.setUint16(28, name.length, true); cd.setUint32(42, offset, true);
      central.push(cd.buffer, name);
      offset += 30 + name.length + f.data.length;
    });
    const centralSize = central.reduce((n, p) => n + (p.byteLength !== undefined ? p.byteLength : p.length), 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, centralSize, true); end.setUint32(16, offset, true);
    return new Blob([...parts, ...central, end.buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  async function runXlsx(ctx) {
    const formats = new Map();          // number format code -> style index
    const styleFor = (code) => {
      if (!formats.has(code)) formats.set(code, formats.size + 2); // 0 = normal, 1 = header
      return formats.get(code);
    };
    let cols = null;
    const meta = [];                    // per column: { kind, style }
    const widths = [];
    const deflater = makeDeflater();
    let rowNo = 1;
    let started = false;
    let probed = false;

    const counts = await readAll(ctx, async (data, columns) => {
      let headRow = '';
      if (!cols) {
        cols = columnIndexes(ctx, columns);
        cols.forEach((c, k) => {
          if (c.i < 0) { meta[k] = { kind: 'rownum' }; widths[k] = 12; return; }
          const samples = data.map((r) => r[c.i].formattedValue);
          if (c.name in IDENTIFIER_COLUMNS) meta[k] = { kind: 'text' };
          else if (DATE_TYPES.includes(c.type)) {
            const withTime = c.type === 'date-time';
            meta[k] = { kind: 'date', style: styleFor(inferDateFormat(samples, withTime)) };
            if (!probed) {
              const sample = data.find((r) => !isBlank(r[c.i]));
              if (sample) {
                probed = true;
                ctx.log(`Date probe: "${c.name}" type=${c.type} value=${JSON.stringify(sample[c.i].value)} (${typeof sample[c.i].value}) nativeValue=${Object.prototype.toString.call(sample[c.i].nativeValue)} formatted=${JSON.stringify(sample[c.i].formattedValue)}`);
              }
            }
          } else if (NUMERIC_TYPES.includes(c.type)) meta[k] = { kind: 'number', style: styleFor(inferNumberFormat(samples)) };
          else meta[k] = { kind: 'text' };
          widths[k] = Math.max(8, Math.min(40, c.name.length + 2));
        });
        const head = cols.map((c, k) => `<c r="${colLetter(k)}1" s="1" t="inlineStr"><is><t>${xmlEscape(c.name)}</t></is></c>`).join('');
        headRow = `<row r="1">${head}</row>`;
      }
      const xml = new Array(data.length);
      for (let r = 0; r < data.length; r++) {
        rowNo++;
        let cells = '';
        for (let k = 0; k < cols.length; k++) {
          const c = cols[k];
          const ref = colLetter(k) + rowNo;
          const m = meta[k];
          if (m.kind === 'rownum') { cells += `<c r="${ref}"><v>${rowNo - 1}</v></c>`; continue; }
          const cell = data[r][c.i];
          if (isBlank(cell)) continue;
          if (m.kind === 'number' && Number.isFinite(Number(cell.value))) {
            cells += `<c r="${ref}" s="${m.style}"><v>${Number(cell.value)}</v></c>`;
          } else if (m.kind === 'date') {
            const p = dateParts(cell);
            if (p) { cells += `<c r="${ref}" s="${m.style}"><v>${excelSerial(p)}</v></c>`; continue; }
            cells += `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(cell.formattedValue)}</t></is></c>`;
          } else {
            const text = c.name in IDENTIFIER_COLUMNS ? identifierText(c.name, cell) : String(cell.formattedValue);
            if (!started && text.length + 2 > widths[k] && widths[k] < 40) widths[k] = Math.min(40, text.length + 2);
            cells += `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(text)}</t></is></c>`;
          }
        }
        xml[r] = `<row r="${rowNo}">${cells}</row>`;
      }
      if (!started) {
        // Column widths come from the header names and the first page; the sheet head is written once they are known.
        started = true;
        const head =
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
          '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
          '<cols>' + widths.map((w, k) => `<col min="${k + 1}" max="${k + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>' +
          '<sheetData>';
        await deflater.write(enc.encode(head + headRow));
      }
      await deflater.write(enc.encode(xml.join('')));
    }, ctx.limit);
    const rows = counts.read;

    const lastCol = colLetter(cols.length - 1);
    await deflater.write(enc.encode(`</sheetData><autoFilter ref="A1:${lastCol}${rowNo}"/></worksheet>`));
    const sheetFile = await deflater.finish('xl/worksheets/sheet1.xml');

    const numFmts = [...formats.keys()].map((code, i) => `<numFmt numFmtId="${164 + i}" formatCode="${xmlEscape(code)}"/>`).join('');
    const xfs = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
      [...formats.keys()].map((code, i) => `<xf numFmtId="${164 + i}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`).join('');
    const styles =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (formats.size ? `<numFmts count="${formats.size}">${numFmts}</numFmts>` : '') +
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFEDEDED"/></patternFill></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${2 + formats.size}">${xfs}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

    const files = [
      { name: '[Content_Types].xml', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
      { name: '_rels/.rels', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Loans" sheetId="1" r:id="rId1"/></sheets></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
      { name: 'xl/styles.xml', text: styles },
    ].map((f) => ({ name: f.name, data: enc.encode(f.text) }));
    files.push(sheetFile);

    return { blob: zipStore(files), rows, total: counts.total };
  }

  // ---- delivery -------------------------------------------------------------

  function download(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  // Returns true if the text reached the clipboard, false if the caller must offer manual copy.
  // In Tableau Desktop the Clipboard API call hangs (frame not focused), so the older copy command goes first.
  async function copyText(text, log) {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    let ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    ta.remove();
    if (ok) return true;
    log('execCommand copy did not work; trying the Clipboard API');
    if (navigator.clipboard && navigator.clipboard.writeText) {
      try {
        await Promise.race([navigator.clipboard.writeText(text), new Promise((_, rej) => setTimeout(() => rej(new Error('timed out')), 3000))]);
        return true;
      } catch (e) { log(`Clipboard API did not work: ${e.message || e}`); }
    }
    return false;
  }

  const stamp = () => { const d = new Date(); return `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}-${pad2(d.getHours())}${pad2(d.getMinutes())}`; };

  async function run(kind, ctx) {
    if (!ctx.columns.length) throw new Error('Choose at least one column first');
    if (kind === 'xlsx') {
      const { blob, rows, total } = await runXlsx(ctx);
      const filename = `Loan-Details-${stamp()}.xlsx`;
      download(blob, filename);
      return { rows, total, filename };
    }
    if (kind === 'csv') {
      const { parts, rows, total } = await runPlain(ctx, ',', true, ctx.limit);
      const filename = `Loan-Details-${stamp()}.csv`;
      download(new Blob(['\uFEFF', ...parts], { type: 'text/csv;charset=utf-8' }), filename);
      return { rows, total, filename };
    }
    if (kind === 'copy') {
      const { parts, rows, total } = await runPlain(ctx, '\t', false, Math.min(ctx.limit || COPY_ROW_LIMIT, COPY_ROW_LIMIT));
      const text = parts.join('');
      const copied = await copyText(text, ctx.log);
      if (!copied) ctx.showManualCopy(text);
      return { rows, total, manual: !copied, limit: COPY_ROW_LIMIT };
    }
    throw new Error(`Unknown export type: ${kind}`);
  }

  window.LoanExport = { run };
})();
