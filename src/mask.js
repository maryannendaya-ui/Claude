(() => {
'use strict';

// ---------------------------------------------------------------------------
// Helpers (shared conventions with the Report Builder)
// ---------------------------------------------------------------------------
const $ = (s, el = document) => el.querySelector(s);
const esc = (x) => String(x ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = (x) => String(x ?? '').toLowerCase().replace(/[^\p{L}\p{N}]+/gu, '');
const isDate = (v) => v instanceof Date && !isNaN(v);
const isBlank = (v) => v === null || v === undefined || v === '';
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const fmtDate = (d) => `${String(d.getUTCDate()).padStart(2, '0')}-${MONTHS[d.getUTCMonth()]}-${d.getUTCFullYear()}`;
const ID_RE = /(\b(id|code|no|num|number|ref|reference|requisition)\.?\s*$)|#|\b(material|vendor|supplier|contractor|plant)\s*$/i;
// A mask token: PREFIX_0042 — distinctive enough to find again inside any text.
const TOKEN_RE = /\b([A-Z][A-Z0-9]{0,11}_\d{4,6})\b/g;

function toast(msg, bad) {
  const t = $('#toast');
  t.textContent = msg;
  t.className = 'show' + (bad ? ' bad' : '');
  clearTimeout(toast.t);
  toast.t = setTimeout(() => { t.className = ''; }, bad ? 7000 : 3500);
}
function busy(on, text) { $('#busy').hidden = !on; if (text) $('#busyText').textContent = text; }
const nextFrame = () => new Promise((r) => setTimeout(r, 20));
function download(name, blob) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 5000);
}

function serialToDate(n) { return new Date(Math.round((n - 25569) * 864e5)); }
function cellValue(cell) {
  if (!cell) return null;
  switch (cell.t) {
    case 'e': case 'z': return null;
    case 'd': return isDate(cell.v) ? new Date(Date.UTC(cell.v.getFullYear(), cell.v.getMonth(), cell.v.getDate())) : null;
    case 'b': return cell.v ? 'TRUE' : 'FALSE';
    case 'n': return cell.z && XLSX.SSF.is_date(cell.z) ? serialToDate(cell.v) : cell.v;
    default: {
      const s = String(cell.v ?? '').replace(/ /g, ' ').trim();
      return s === '' ? null : s;
    }
  }
}
function sheetToGrid(ws) {
  let maxR = -1, maxC = -1;
  for (const k of Object.keys(ws)) {
    if (k[0] === '!') continue;
    const a = XLSX.utils.decode_cell(k);
    if (a.r > maxR) maxR = a.r;
    if (a.c > maxC) maxC = a.c;
  }
  const grid = [];
  for (let R = 0; R <= maxR; R++) {
    const row = new Array(maxC + 1).fill(null);
    for (let C = 0; C <= maxC; C++) row[C] = cellValue(ws[XLSX.utils.encode_cell({ r: R, c: C })]);
    grid.push(row);
  }
  return grid;
}
function detectHeaderRow(grid) {
  let best = 0, bestScore = 0;
  for (let i = 0; i < Math.min(grid.length, 30); i++) {
    const strs = grid[i].filter((v) => typeof v === 'string' && !/^[\d.,\s-]+$/.test(v));
    const distinct = new Set(strs.map((x) => x.toLowerCase())).size;
    if (distinct > bestScore) { bestScore = distinct; best = i; }
  }
  return best;
}
async function readWorkbook(file) {
  const buf = await file.arrayBuffer();
  if (/\.(csv|txt|tsv)$/i.test(file.name)) {
    let text = new TextDecoder('utf-8').decode(buf);
    if (text.includes('�')) text = new TextDecoder('windows-1252').decode(buf);
    return XLSX.read(text.replace(/^﻿/, ''), { type: 'string', raw: true });
  }
  return XLSX.read(buf, { type: 'array', cellNF: true });
}

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
const state = {
  files: [],      // { name, sheets: [{ name, grid, headerIdx }] }
  columns: [],    // { norm, name, mode, count, samples: [] } aggregated by header
  key: null,      // loaded key for unmasking: { nameRev: Map token->orig, codeRev: Map digits->orig }
  keyRaw: null,
  maskResults: [],    // { name, blob, note, key? } ready to save
  unmaskResults: [],
  maps: null          // the session's mapping — reused on every mask run so repeated runs stay consistent
};
function getMaps() {
  return state.maps || (state.maps = { names: new Map(), codes: new Map(), counters: {}, usedNums: new Set() });
}
/** Seeds the session mapping from an earlier key file, so files masked now stay
 *  consistent with files masked in a previous session. */
function seedMaps(k) {
  const maps = getMaps();
  for (const [t, o] of k.names || []) {
    const key = String(o).trim().replace(/\s+/g, ' ').toLowerCase();
    if (!maps.names.has(key)) maps.names.set(key, { token: t, original: String(o).trim() });
    const m = t.match(/^([A-Z][A-Z0-9]*)_(\d+)$/);
    if (m) maps.counters[m[1]] = Math.max(maps.counters[m[1]] || 0, +m[2]);
  }
  for (const [mk, o] of k.codes || []) {
    const key = String(o).replace(/^0+(?=\d)/, '');
    if (!maps.codes.has(key)) maps.codes.set(key, { masked: mk, original: String(o) });
    maps.usedNums.add(mk);
    maps.usedNums.add(key);
  }
}

/** Explicit save buttons — browsers often block automatic multi-file downloads,
 *  so the results of each step are always listed with their own Save button. */
function renderResults(elId, results, which) {
  const el = $(elId);
  if (!results.length) { el.innerHTML = ''; return; }
  el.innerHTML = `<div class="tbl-wrap" style="margin-top:14px"><table class="ui"><thead><tr><th>Ready to save</th><th></th><th></th></tr></thead><tbody>` +
    results.map((r, i) => `<tr>
      <td>${r.key ? '🔑 ' : ''}${esc(r.name)}${r.key ? ' — <strong>keep this private; needed to unmask</strong>' : ''}</td>
      <td class="muted">${esc(r.note || '')}</td>
      <td><button class="btn primary sm" data-dl="${which}:${i}">Save file</button></td>
    </tr>`).join('') +
    `</tbody></table></div>
    <div class="row gap" style="margin-top:10px"><button class="btn" data-dl="${which}:all">Save all (${results.length})</button>
    <span class="muted">If a download does not appear, your browser blocked it — use the individual Save buttons.</span></div>`;
}
document.addEventListener('click', async (e) => {
  const t = e.target.closest('[data-dl]');
  if (!t) return;
  const [which, what] = t.dataset.dl.split(':');
  const results = which === 'm' ? state.maskResults : state.unmaskResults;
  if (what === 'all') {
    for (const r of results) { download(r.name, r.blob); await new Promise((res) => setTimeout(res, 500)); }
  } else {
    const r = results[+what];
    if (r) download(r.name, r.blob);
  }
});

// ---------------------------------------------------------------------------
// Mask side
// ---------------------------------------------------------------------------
const MODES = [['token', 'Mask text (token)'], ['renumber', 'Mask code (new number)'], ['keep', 'Keep as is'], ['blank', 'Blank out']];

function suggestMode(col) {
  let nums = 0, texts = 0, dates = 0;
  for (const v of col.samples) {
    if (isDate(v)) dates++;
    else if (typeof v === 'number') nums++;
    else texts++;
  }
  const total = nums + texts + dates;
  if (!total) return 'keep';
  if (dates / total > 0.6) return 'keep';
  if (nums / total > 0.6) return ID_RE.test(col.name) ? 'renumber' : 'keep';  // amounts stay, IDs get renumbered
  const strs = col.samples.filter((v) => typeof v === 'string');
  // quantities with a unit ("1.000 AU") and small harmless enums (Open/Closed, USD/AED) stay readable
  if (strs.length && strs.every((v) => /^[\d.,\s]+[A-Za-z%]{1,5}$/.test(v.trim()))) return 'keep';
  const distinct = new Set(strs.map((v) => v.trim().toLowerCase()));
  if (strs.length >= 5 && distinct.size <= 6 && [...distinct].every((v) => v.length <= 12)) return 'keep';
  return 'token';
}

async function addMaskFiles(fileList) {
  const files = [...fileList];
  if (!files.length) return;
  busy(true, 'Reading files…');
  await nextFrame();
  for (const f of files) {
    try {
      const wb = await readWorkbook(f);
      const entry = { name: f.name, sheets: [] };
      for (const sn of wb.SheetNames) {
        const grid = sheetToGrid(wb.Sheets[sn]);
        if (!grid.length) continue;
        entry.sheets.push({ name: sn, grid, headerIdx: detectHeaderRow(grid) });
      }
      if (entry.sheets.length) state.files.push(entry);
    } catch (e) {
      console.error(e);
      toast(`Could not read ${f.name}: ${e.message || e}`, true);
    }
  }
  rebuildColumns();
  busy(false);
  renderMask();
}

function rebuildColumns() {
  const byNorm = new Map(state.columns.map((c) => [c.norm, c]));
  const seen = new Map();
  for (const f of state.files) {
    for (const sh of f.sheets) {
      const hdr = sh.grid[sh.headerIdx] || [];
      hdr.forEach((h, ci) => {
        if (isBlank(h)) return;
        const name = String(h).replace(/\s+/g, ' ').trim();
        const n = norm(name) || 'col' + ci;
        let col = seen.get(n);
        if (!col) {
          col = byNorm.get(n) || { norm: n, name, mode: null };
          col.count = 0;
          col.samples = [];
          seen.set(n, col);
        }
        col.count++;
        for (let r = sh.headerIdx + 1; r < sh.grid.length && col.samples.length < 60; r += Math.max(1, Math.floor((sh.grid.length - sh.headerIdx) / 25))) {
          const v = sh.grid[r] && sh.grid[r][ci];
          if (!isBlank(v)) col.samples.push(v);
        }
      });
    }
  }
  state.columns = [...seen.values()];
  for (const c of state.columns) if (!c.mode) c.mode = suggestMode(c);
}

function sampleText(c) {
  const v = c.samples.find((x) => !isBlank(x));
  if (v === undefined) return '';
  return isDate(v) ? fmtDate(v) : String(v).slice(0, 60);
}

function renderMask() {
  const el = $('#maskCols');
  $('#maskOpts').hidden = !state.files.length;
  $('#btnMask').disabled = !state.files.length;
  if (!state.files.length) { el.innerHTML = ''; $('#maskStats').textContent = ''; return; }
  const nSheets = state.files.reduce((a, f) => a + f.sheets.length, 0);
  $('#maskStats').textContent = `${state.files.length} file(s), ${nSheets} sheet(s), ${state.columns.length} distinct columns`;
  const rows = state.columns.map((c, i) => `<tr>
    <td>${esc(c.name)}</td>
    <td class="num">${c.count}</td>
    <td class="muted" style="max-width:320px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(sampleText(c))}</td>
    <td><select data-i="${i}">${MODES.map(([v, l]) => `<option value="${v}" ${c.mode === v ? 'selected' : ''}>${l}</option>`).join('')}</select></td>
    <td class="muted">${esc(maskPreview(c))}</td>
  </tr>`).join('');
  el.innerHTML = `<div class="tbl-wrap"><table class="ui"><thead><tr><th>Column</th><th class="num">Sheets</th><th>Sample (real)</th><th>Masking</th><th>Becomes</th></tr></thead><tbody>${rows}</tbody></table></div>`;
}

function maskPreview(c) {
  const v = c.samples.find((x) => !isBlank(x));
  if (v === undefined) return '';
  const trial = { names: new Map(), codes: new Map(), counters: {}, usedNums: new Set() };
  const out = maskValue(v, c, trial);
  return isBlank(out) ? '(blank)' : isDate(out) ? fmtDate(out) : String(out).slice(0, 60);
}

// --- the mapping machinery ---------------------------------------------------
function tokenPrefix(colName) {
  const w = String(colName).match(/[A-Za-z]{2,}/);
  return (w ? w[0].toUpperCase().slice(0, 10) : 'VAL');
}
function nameToken(raw, col, maps) {
  const key = String(raw).trim().replace(/\s+/g, ' ').toLowerCase();
  if (!key) return null;
  let t = maps.names.get(key);
  if (!t) {
    const p = tokenPrefix(col.name);
    maps.counters[p] = (maps.counters[p] || 0) + 1;
    t = { token: `${p}_${String(maps.counters[p]).padStart(4, '0')}`, original: String(raw).trim() };
    maps.names.set(key, t);
  }
  return t.token;
}
function codeNumber(digits, maps) {
  const key = digits.replace(/^0+(?=\d)/, '');
  let m = maps.codes.get(key);
  if (!m) {
    maps.usedNums.add(key);  // no future mask may equal this original
    const len = Math.max(digits.length, 4);
    let cand;
    do {
      cand = String(Math.floor(Math.random() * 9) + 1);
      while (cand.length < len) cand += String(Math.floor(Math.random() * 10));
    } while (maps.usedNums.has(cand));
    maps.usedNums.add(cand);
    m = { masked: cand, original: digits };
    maps.codes.set(key, m);
  }
  // keep the original's leading zeros shape
  return m.masked.padStart(digits.length, '9');
}
/** Remaps every 4+ digit run inside a string consistently. */
function remapDigits(s, maps) {
  return String(s).replace(/(\d)[,.\s](?=\d)/g, '$1').replace(/\d{4,}/g, (d) => codeNumber(d, maps));
}
function maskValue(v, col, maps) {
  if (isBlank(v)) return null;
  switch (col.mode) {
    case 'keep': return v;
    case 'blank': return null;
    case 'renumber':
      if (isDate(v)) return v;
      if (typeof v === 'number') return Number.isInteger(v) ? Number(codeNumber(String(Math.abs(v)), maps)) * Math.sign(v || 1) : v;
      if (/\d{4,}/.test(String(v).replace(/(\d)[,.\s](?=\d)/g, '$1'))) return remapDigits(v, maps);
      return nameToken(v, col, maps);
    default: {  // token
      if (isDate(v) || typeof v === 'number') return v;
      const s = String(v).trim();
      // "NAME (CODE)" keeps its shape: text part -> token, code part -> remapped number
      const m = s.match(/^(.*?)[\s]*\(\s*([A-Za-z0-9 \-\/.]{2,15})\s*\)\s*$/);
      if (m && /\d{3,}/.test(m[2])) {
        const text = m[1].trim();
        return (text ? nameToken(text, col, maps) + ' ' : '') + '(' + remapDigits(m[2], maps) + ')';
      }
      return nameToken(s, col, maps);
    }
  }
}

async function doMask() {
  if (typeof ExcelJS === 'undefined') { toast('Excel library failed to load.', true); return; }
  busy(true, 'Masking…');
  await nextFrame();
  try {
    const maps = getMaps();   // persists for the session: masking twice gives identical results
    const colByNorm = new Map(state.columns.map((c) => [c.norm, c]));
    const blankTop = $('#optBlankTop').checked;
    let cells = 0;
    const outputs = [];
    for (const f of state.files) {
      const wb = new ExcelJS.Workbook();
      for (const sh of f.sheets) {
        const ws = wb.addWorksheet(sh.name.replace(/[[\]:*?/\\]/g, ' ').slice(0, 31) || 'Sheet');
        const hdr = sh.grid[sh.headerIdx] || [];
        const cols = hdr.map((h) => (isBlank(h) ? null : colByNorm.get(norm(String(h).replace(/\s+/g, ' ').trim())) || null));
        sh.grid.forEach((row, r) => {
          if (r < sh.headerIdx && blankTop) return;              // titles above the header: dropped
          const xr = ws.getRow(r + 1 - (blankTop ? sh.headerIdx : 0));
          if (r === sh.headerIdx) {
            row.forEach((h, ci) => { if (!isBlank(h)) { const cell = xr.getCell(ci + 1); cell.value = String(h); cell.font = { bold: true }; } });
            return;
          }
          if (r < sh.headerIdx) { row.forEach((v, ci) => { if (!isBlank(v)) xr.getCell(ci + 1).value = isDate(v) ? v : v; }); return; }
          row.forEach((v, ci) => {
            if (isBlank(v)) return;
            const col = cols[ci];
            const out = col ? maskValue(v, col, maps) : v;
            cells++;
            if (!isBlank(out)) xr.getCell(ci + 1).value = out;
          });
        });
        ws.columns.forEach((c) => { c.width = 18; });
      }
      const buf = await wb.xlsx.writeBuffer();
      outputs.push({ name: f.name.replace(/\.[^.]+$/, '') + ' - MASKED.xlsx', blob: new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }) });
    }
    // key file
    const key = {
      app: 'maskkit', version: 1, created: new Date().toISOString(),
      columns: state.columns.map((c) => ({ name: c.name, mode: c.mode })),
      names: [...maps.names.values()].map((t) => [t.token, t.original]),
      codes: [...maps.codes.values()].map((m) => [m.masked, m.original])
    };
    let keyBlob;
    const pass = $('#passphrase').value;
    if (pass) {
      const enc = await encryptJSON(key, pass);
      keyBlob = new Blob([JSON.stringify(enc)], { type: 'application/json' });
    } else {
      keyBlob = new Blob([JSON.stringify(key, null, 1)], { type: 'application/json' });
    }
    busy(false);
    state.maskResults = outputs.map((o) => ({ ...o, note: 'safe to share' }));
    state.maskResults.push({ name: 'UNMASK KEY - keep private.maskkey.json', blob: keyBlob, key: true, note: pass ? 'encrypted' : 'plain — consider a passphrase' });
    renderResults('#maskResults', state.maskResults, 'm');
    $('#maskResults').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    download(state.maskResults[0].name, state.maskResults[0].blob);  // first one is a direct user action, never blocked
    toast(`Masked ${cells.toLocaleString('en-US')} cells (${maps.names.size.toLocaleString('en-US')} text values, ${maps.codes.size.toLocaleString('en-US')} codes). Save each file below — including the key.`);
  } catch (e) {
    console.error(e);
    busy(false);
    toast('Masking failed: ' + (e.message || e), true);
  }
}

// ---------------------------------------------------------------------------
// Key encryption (optional, AES-GCM with PBKDF2)
// ---------------------------------------------------------------------------
const te = new TextEncoder(), td = new TextDecoder();
async function deriveKey(pass, salt) {
  const base = await crypto.subtle.importKey('raw', te.encode(pass), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations: 200000, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
}
const b64 = (u8) => btoa(String.fromCharCode(...u8));
const unb64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
async function encryptJSON(obj, pass) {
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const key = await deriveKey(pass, salt);
  const data = te.encode(JSON.stringify(obj));
  const ct = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, data));
  return { app: 'maskkit', version: 1, encrypted: true, salt: b64(salt), iv: b64(iv), data: b64(ct) };
}
async function decryptJSON(obj, pass) {
  const key = await deriveKey(pass, unb64(obj.salt));
  const pt = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(obj.iv) }, key, unb64(obj.data));
  return JSON.parse(td.decode(new Uint8Array(pt)));
}

// ---------------------------------------------------------------------------
// Unmask side
// ---------------------------------------------------------------------------
async function loadKey(file) {
  try {
    state.keyRaw = JSON.parse(await file.text());
    if (state.keyRaw.app !== 'maskkit') throw new Error('not a Mask Kit key file');
    if (state.keyRaw.encrypted) {
      $('#keyName').textContent = `🔒 ${file.name} — enter the passphrase, then load the processed files`;
      state.key = null;
    } else {
      buildKey(state.keyRaw);
      $('#keyName').textContent = `✓ ${file.name} (${state.key.nameRev.size + state.key.codeRev.size} mappings)`;
    }
    toast('Key file loaded.');
  } catch (e) {
    state.key = state.keyRaw = null;
    $('#keyName').textContent = '';
    toast(`Could not load key: ${e.message || e}`, true);
  }
}
function buildKey(k) {
  state.key = {
    nameRev: new Map(k.names.map(([t, o]) => [t, o])),
    codeRev: new Map(k.codes.map(([m, o]) => [m.replace(/^0+(?=\d)/, ''), o]))
  };
}
async function ensureKey() {
  if (state.key) return true;
  if (!state.keyRaw) { toast('Load the key file first.', true); return false; }
  if (state.keyRaw.encrypted) {
    const pass = $('#passphrase2').value;
    if (!pass) { toast('This key file is encrypted — enter the passphrase.', true); return false; }
    try {
      buildKey(await decryptJSON(state.keyRaw, pass));
      $('#keyName').textContent = '✓ key decrypted';
      return true;
    } catch (e) {
      toast('Wrong passphrase (or corrupted key file).', true);
      return false;
    }
  }
  return false;
}

function unmaskString(s, stats) {
  let out = String(s).replace(TOKEN_RE, (t) => {
    const o = state.key.nameRev.get(t);
    if (o !== undefined) { stats.names++; return o; }
    stats.unknown++;
    return t;
  });
  // codes written with thousand separators ("968,168" / "968 168") first…
  out = out.replace(/\d{1,3}(?:[,\s]\d{3})+/g, (g) => {
    const o = state.key.codeRev.get(g.replace(/[,\s]/g, '').replace(/^0+(?=\d)/, ''));
    if (o !== undefined) { stats.codes++; return o; }
    return g;
  });
  // …then plain digit runs
  out = out.replace(/\d{4,}/g, (d) => {
    const o = state.key.codeRev.get(d.replace(/^0+(?=\d)/, ''));
    if (o !== undefined) { stats.codes++; return o; }
    return d;
  });
  return out;
}
function unmaskCellValue(v, stats) {
  if (v === null || v === undefined) return v;
  if (typeof v === 'string') return unmaskString(v, stats);
  if (typeof v === 'number' && Number.isInteger(v)) {
    const o = state.key.codeRev.get(String(Math.abs(v)).replace(/^0+(?=\d)/, ''));
    if (o !== undefined) { stats.codes++; const n = Number(o); return isNaN(n) ? o : n * Math.sign(v || 1); }
  }
  return v;
}

async function unmaskFiles(fileList) {
  const files = [...fileList];
  if (!files.length) return;
  if (!state.key && !state.keyRaw) { $('#unmaskList').innerHTML = '<div class="empty" style="margin-top:12px">Load the key file (left box) before dropping the processed files.</div>'; toast('Load the key file first.', true); return; }
  if (!(await ensureKey())) return;
  state.unmaskResults = [];
  busy(true, 'Unmasking…');
  await nextFrame();
  const report = [];
  for (const f of files) {
    const stats = { names: 0, codes: 0, unknown: 0 };
    try {
      let blob, outName;
      if (/\.(csv|txt|tsv)$/i.test(f.name)) {
        let text = new TextDecoder('utf-8').decode(await f.arrayBuffer());
        text = unmaskString(text, stats);
        outName = f.name.replace(/\.[^.]+$/, '') + ' - UNMASKED.csv';
        blob = new Blob(['﻿' + text], { type: 'text/csv' });
      } else {
        // ExcelJS round-trip keeps the formatting of the processed report
        const wb = new ExcelJS.Workbook();
        let legacy = false;
        const buf = await f.arrayBuffer();
        try {
          await wb.xlsx.load(buf);
        } catch (err) {
          // old .xls (or anything ExcelJS cannot open): values-only fallback via SheetJS
          legacy = true;
          const xwb = XLSX.read(buf, { type: 'array', cellNF: true });
          for (const sn of xwb.SheetNames) {
            const ws = wb.addWorksheet(unmaskString(sn, stats).replace(/[[\]:*?/\\]/g, ' ').slice(0, 31) || 'Sheet');
            sheetToGrid(xwb.Sheets[sn]).forEach((row, r) => row.forEach((v, c) => {
              if (!isBlank(v)) ws.getRow(r + 1).getCell(c + 1).value = unmaskCellValue(v, stats);
            }));
          }
        }
        if (!legacy) wb.worksheets.forEach((ws) => {
          const nn = unmaskString(ws.name, stats).replace(/[[\]:*?/\\]/g, ' ').slice(0, 31);
          if (nn && nn !== ws.name) { try { ws.name = nn; } catch (err) { /* duplicate name — keep */ } }
        });
        if (!legacy) wb.eachSheet((ws) => {
          ws.eachRow({ includeEmpty: false }, (row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
              const v = cell.value;
              if (v === null || v === undefined) return;
              if (typeof v === 'string' || typeof v === 'number') {
                const nv = unmaskCellValue(v, stats);
                if (nv !== v) cell.value = nv;
              } else if (v.richText) {
                let changed = false;
                const rt = v.richText.map((p) => {
                  const nt = unmaskString(p.text, stats);
                  if (nt !== p.text) changed = true;
                  return { ...p, text: nt };
                });
                if (changed) cell.value = { richText: rt };
              } else if (v.text !== undefined && v.hyperlink !== undefined) {
                const nt = unmaskString(v.text, stats);
                if (nt !== v.text) cell.value = { ...v, text: nt };
              } else if (v.formula !== undefined && typeof v.result === 'string') {
                const nr = unmaskString(v.result, stats);
                if (nr !== v.result) cell.value = { ...v, result: nr };
              }
            });
          });
        });
        const out = await wb.xlsx.writeBuffer();
        outName = f.name.replace(/\.[^.]+$/, '') + ' - UNMASKED.xlsx';
        blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      }
      state.unmaskResults.push({ name: outName, blob, note: `${(stats.names + stats.codes).toLocaleString('en-US')} values restored (${stats.names.toLocaleString('en-US')} names, ${stats.codes.toLocaleString('en-US')} codes)` + (stats.unknown ? ` — \u26a0 ${stats.unknown.toLocaleString('en-US')} token(s) NOT in this key: they were masked in a different run/session, load that run's key and unmask again` : '') });
      if (stats.unknown) toast(`${f.name}: ${stats.unknown} masked token(s) are not in this key — they came from a different masking run. Unmask the result again with that run's key, or re-mask everything in one session.`, true);
    } catch (e) {
      console.error(e);
      report.push(`${esc(f.name)}: <span class="bad-text">failed — ${esc(e.message || e)}</span>`);
    }
  }
  busy(false);
  renderResults('#unmaskList', state.unmaskResults, 'u');
  if (report.length) $('#unmaskList').insertAdjacentHTML('beforeend', `<ul style="margin:12px 0 0;padding-left:20px;line-height:1.8">${report.map((r) => `<li>${r}</li>`).join('')}</ul>`);
  if (state.unmaskResults.length) {
    $('#unmaskStats').textContent = `${state.unmaskResults.length} file(s) unmasked — save them below.`;
    download(state.unmaskResults[state.unmaskResults.length - 1].name, state.unmaskResults[state.unmaskResults.length - 1].blob);
  }
}

// ---------------------------------------------------------------------------
// Events
// ---------------------------------------------------------------------------
function wireDrop(id, handler) {
  const el = $(id);
  el.addEventListener('dragover', (e) => { e.preventDefault(); el.classList.add('over'); });
  el.addEventListener('dragleave', () => el.classList.remove('over'));
  el.addEventListener('drop', (e) => { e.preventDefault(); el.classList.remove('over'); handler(e.dataTransfer.files); });
}

$('#maskInput').addEventListener('change', (e) => { addMaskFiles(e.target.files); e.target.value = ''; });
wireDrop('#dropMask', addMaskFiles);
$('#maskCols').addEventListener('change', (e) => {
  const i = e.target.dataset.i;
  if (i === undefined) return;
  state.columns[+i].mode = e.target.value;
  renderMask();
});
$('#btnMask').addEventListener('click', doMask);
$('#seedKeyInput').addEventListener('change', async (e) => {
  const f = e.target.files[0];
  e.target.value = '';
  if (!f) return;
  try {
    let k = JSON.parse(await f.text());
    if (k.app !== 'maskkit') throw new Error('not a Mask Kit key file');
    if (k.encrypted) {
      const pass = $('#passphrase').value;
      if (!pass) throw new Error('key is encrypted — enter its passphrase in the field above first');
      k = await decryptJSON(k, pass);
    }
    seedMaps(k);
    $('#seedKeyName').textContent = `✓ continuing from ${f.name} (${(k.names || []).length + (k.codes || []).length} mappings kept)`;
    toast('Previous key loaded — new masking stays consistent with files masked under that key.');
  } catch (err) {
    toast('Could not load previous key: ' + (err.message || err), true);
  }
});
$('#keyInput').addEventListener('change', (e) => { if (e.target.files[0]) loadKey(e.target.files[0]); e.target.value = ''; });
wireDrop('#dropKey', (f) => f[0] && loadKey(f[0]));
$('#unmaskInput').addEventListener('change', (e) => { unmaskFiles(e.target.files); e.target.value = ''; });
wireDrop('#dropUnmask', unmaskFiles);
$('#passphrase2').addEventListener('change', () => { if (state.keyRaw && state.keyRaw.encrypted) state.key = null; });
$('#btnReset').addEventListener('click', () => {
  if (state.files.length && !confirm('Clear all loaded files and column settings?')) return;
  state.files = []; state.columns = []; state.key = null; state.keyRaw = null;
  state.maskResults = []; state.unmaskResults = []; state.maps = null;
  $('#seedKeyName').textContent = '';
  $('#keyName').textContent = ''; $('#unmaskList').innerHTML = ''; $('#maskResults').innerHTML = ''; $('#unmaskStats').textContent = '';
  renderMask();
});

window.MaskKit = { state, maskValue, unmaskString };
renderMask();
})();
