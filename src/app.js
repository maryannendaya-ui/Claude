(() => {
'use strict';

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------
const SETTINGS_VERSION = 1;
const LS_KEY = 'reportBuilder.settings.v1';
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MONTH_LOOKUP = {};
['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
  .forEach((m, i) => { MONTH_LOOKUP[m] = i + 1; MONTH_LOOKUP[m.slice(0, 3)] = i + 1; });
MONTH_LOOKUP.sept = 9;

const TYPE_OPTIONS = [['auto', 'Auto-detect'], ['text', 'Text'], ['int', 'Number (whole)'], ['dec', 'Number (2 decimals)'], ['pct', 'Percent'], ['date', 'Date']];
const TYPE_NAME = { text: 'Text', int: 'Number', dec: 'Number (0.00)', pct: 'Percent', date: 'Date' };
const NUMFMT = { text: 'General', int: '#,##0', dec: '#,##0.00', pct: '0.00%', date: 'dd-mmm-yyyy' };
const AGG_OPTIONS = [['count', 'Count of records'], ['counta', 'Count of'], ['distinct', 'Unique count of'], ['list', 'List unique of'], ['sum', 'Sum of'], ['avg', 'Average of'], ['min', 'Minimum of'], ['max', 'Maximum of']];
const LIST_CAP = 15;
const NUMERIC = new Set(['int', 'dec', 'pct']);

const TOTAL_RE = /^\s*(grand\s*|sub\s*-?\s*)?totals?\b/i;
const PLACEHOLDER_RE = /^\s*(-+|—|–|n\/?a|nil|none|null|tbd|tba)\s*$/i;
const ID_RE = /(\b(id|code|no|num|number|ref|reference|item)\.?\s*$)|#|\b(phone|mobile|telephone|tel|fax|zip|postal|iban|passport|s\/n|sn|serial)\b/i;
const SERIAL_RE = /^\s*(s\/?\s*n\.?|sr\.?\s*(no\.?)?|sl\.?\s*(no\.?)?|no\.?|#|serial(\s*(no\.?|number))?|item\s*no\.?)\s*$/i;
const CAT_RE = /(fmsn|uom|abc|department|dept|division|region|area|zone|status|category|type|class|branch|location|site|project|client|customer|vendor|supplier|nationality|country|city|gender|team|grade|level|designation|position|role|product|service|segment|channel|stage|priority|company|entity|unit|section|group|currency|method|mode)/i;

// Excel styling palette (ARGB)
const XC = {
  title: 'FF1F3864', sub: 'FF595959', head: 'FF1F4E79', headFont: 'FFFFFFFF',
  band: 'FFF3F6FB', total: 'FFDCE6F2', border: 'FFA6B4C8', section: 'FF1F3864'
};

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
const defaultOpts = () => ({
  addSN: true, addSource: true, dropTotals: true, fillMerged: true, dedupe: false,
  masterTotals: true, sortBy: '', sortDir: 'asc', align: 'center', dateOrder: 'DMY', splitBy: ''
});

const state = {
  sources: [],          // one entry per sheet
  template: null,       // { fileName, headers: [{ name, norm }] }
  colPrefs: {},         // key -> { name, include, type }
  colOrder: [],         // keys in display order
  aliasOverrides: {},   // header norm -> column key
  columns: [],          // derived each reconcile
  normMap: new Map(),   // header norm -> column key
  opts: defaultOpts(),
  breakdowns: [],
  bdAuto: true,
  title: 'Master Report',
  fileName: '',
  master: null,
  summary: null,
  tab: 'summary',
  peek: null,
  sheetPrefs: {},       // sheetKey -> { include, headerRow, role } — reapplied when the same file/sheet is loaded again
  connections: [],      // one per included reference sheet
  connPrefs: {},        // sheetKey -> { name, refKeyNorm, masterKey, addCount, exportSheet }
  verification: null
};
let nextId = 1;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
const $ = (s, el = document) => el.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = (s) => String(s ?? '').toLowerCase().replace(/&/g, ' and ').replace(/[^\p{L}\p{N}]+/gu, '');
const digits = (s) => s.replace(/\D+/g, '');
const isDate = (v) => v instanceof Date && !isNaN(v);
const isBlank = (v) => v === null || v === undefined || v === '';
const sheetKey = (s) => s.fileName + '|' + s.sheetName;
/** Normalises a key for matching: trims, ignores case and leading zeros ("000123" matches 123). */
function keyNorm(v) {
  if (isBlank(v)) return null;
  if (isDate(v)) return 'd' + v.getTime();
  let s = String(v).trim().toLowerCase().replace(/\s+/g, ' ').replace(/\.0+$/, '');
  if (/^\d+$/.test(s)) s = s.replace(/^0+(?=\d)/, '');
  return s || null;
}
/** Keys a master value can match on: the whole value, plus any embedded number of 4+ digits. */
/** Pulls the code out of values like "INSPECTION, OVERHAUL (321110)" -> "321110".
 *  Prefers a code in trailing parentheses; falls back to the last embedded number. */
function extractCode(v) {
  if (isBlank(v)) return null;
  const s = String(isDate(v) ? fmtDate(v) : v);
  const m = s.match(/\(\s*([A-Za-z0-9][A-Za-z0-9 \-\/.]{0,14})\s*\)\s*$/);
  if (m && /\d/.test(m[1])) return m[1].replace(/\s+/g, '');
  const collapsed = s.replace(/(\d)[,.\s](?=\d)/g, '$1');
  const nums = collapsed.match(/\d{4,}/g);
  return nums ? nums[nums.length - 1] : null;
}

function matchCandidates(v) {
  const out = [];
  const k = keyNorm(v);
  if (!k) return out;
  out.push(k);
  if (typeof v === 'string') {
    const collapsed = v.replace(/(\d)[,.\s](?=\d)/g, '$1');
    const nums = collapsed.match(/\d{4,}/g);
    if (nums) for (const d of nums) { const n = d.replace(/^0+(?=\d)/, ''); if (n !== k && !out.includes(n)) out.push(n); }
  }
  return out;
}
const fmtInt = (n) => Math.round(n).toLocaleString('en-US');

function colLetter(c) {
  let s = '';
  c += 1;
  while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}
function fmtDate(d) {
  return `${String(d.getUTCDate()).padStart(2, '0')}-${MONTHS[d.getUTCMonth()]}-${d.getUTCFullYear()}`;
}
function fmtDateTime(d) {
  return `${String(d.getDate()).padStart(2, '0')}-${MONTHS[d.getMonth()]}-${d.getFullYear()} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
}
function makeDate(y, m, d) {
  if (m < 1 || m > 12 || d < 1 || d > 31 || y < 1900 || y > 2200) return null;
  const dt = new Date(Date.UTC(y, m - 1, d));
  return dt.getUTCDate() === d ? dt : null;
}
function serialToDate(n, date1904) {
  if (date1904) n += 1462;
  return new Date(Math.round((n - 25569) * 864e5));
}
function similarity(a, b) {
  if (a === b) return 1;
  const m = a.length, n = b.length;
  if (!m || !n) return 0;
  let prev = Array.from({ length: n + 1 }, (_, j) => j);
  for (let i = 1; i <= m; i++) {
    const cur = [i];
    for (let j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    prev = cur;
  }
  return 1 - prev[n] / Math.max(m, n);
}
function cmpVals(a, b) {
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'number') return -1;
  if (typeof b === 'number') return 1;
  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: 'base' });
}
function sortVal(v) {
  if (isDate(v)) return v.getTime();
  return v;
}

/** Parses "1,234.50", "AED 1,200", "(500)", "12.5%" etc. Returns { n, pct } or null. */
function toNumber(v) {
  if (typeof v === 'number') return isFinite(v) ? { n: v, pct: false } : null;
  if (typeof v !== 'string') return null;
  let s = v.trim();
  if (!s || /^0\d/.test(s.replace(/^[-+]/, ''))) return null; // leading zeros => code, not a number
  let neg = false, pct = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1).trim(); }
  if (s.endsWith('%')) { pct = true; s = s.slice(0, -1).trim(); }
  s = s.replace(/^(?:[A-Za-z]{1,4}\.?\s+|[$€£¥₱₹﷼]\s*)/, '').replace(/(?:\s+[A-Za-z]{1,4}|\s*[$€£¥₱₹﷼])$/, '');
  s = s.replace(/[\s,']/g, '');
  if (!/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(s)) return null;
  let n = parseFloat(s);
  if (!isFinite(n)) return null;
  if (neg) n = -n;
  return { n, pct };
}

/** Parses common text dates. Returns a UTC-midnight Date or null. */
function parseDateStr(s) {
  if (typeof s !== 'string') return null;
  s = s.trim();
  let m;
  if ((m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[ T]\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/))) {
    return makeDate(+m[1], +m[2], +m[3]);
  }
  if ((m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4}|\d{2})(?:\s+\d{1,2}:\d{2}(?::\d{2})?(?:\s*[ap]m)?)?$/i))) {
    let a = +m[1], b = +m[2], y = +m[3];
    if (y < 100) y += y < 50 ? 2000 : 1900;
    const mdy = state.opts.dateOrder === 'MDY';
    let d = mdy ? makeDate(y, a, b) : makeDate(y, b, a);
    if (!d) d = mdy ? makeDate(y, b, a) : makeDate(y, a, b); // unambiguous fallback
    return d;
  }
  if ((m = s.match(/^(\d{1,2})(?:st|nd|rd|th)?[-\s/.]+([A-Za-z]{3,9})\.?[-\s/.,]+(\d{4}|\d{2})$/))) {
    const mo = MONTH_LOOKUP[m[2].toLowerCase()];
    let y = +m[3]; if (y < 100) y += y < 50 ? 2000 : 1900;
    return mo ? makeDate(y, mo, +m[1]) : null;
  }
  if ((m = s.match(/^([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})$/))) {
    const mo = MONTH_LOOKUP[m[1].toLowerCase()];
    return mo ? makeDate(+m[3], mo, +m[2]) : null;
  }
  return null;
}

function displayValue(v, type) {
  if (isBlank(v)) return '';
  if (isDate(v)) return fmtDate(v);
  if (typeof v === 'number') {
    if (type === 'pct') return (v * 100).toFixed(2) + '%';
    if (type === 'dec') return v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    if (type === 'int') return fmtInt(v);
    return Number.isInteger(v) ? v.toLocaleString('en-US') : v.toLocaleString('en-US', { maximumFractionDigits: 4 });
  }
  return String(v);
}
function autoType(v, type) {
  if (type && type !== 'auto') return type;
  if (isDate(v)) return 'date';
  if (typeof v === 'number') return Number.isInteger(v) ? 'int' : 'dec';
  return 'text';
}

function toast(msg, bad) {
  const t = $('#toast');
  t.textContent = msg;
  t.className = 'show' + (bad ? ' bad' : '');
  clearTimeout(toast.t);
  toast.t = setTimeout(() => { t.className = ''; }, bad ? 6000 : 3200);
}
function busy(on, text) {
  $('#busy').hidden = !on;
  if (text) $('#busyText').textContent = text;
}
const nextFrame = () => new Promise((r) => setTimeout(r, 20));

// ---------------------------------------------------------------------------
// Reading files
// ---------------------------------------------------------------------------
function cellValue(cell, date1904) {
  if (!cell) return null;
  switch (cell.t) {
    case 'e': case 'z': return null;
    case 'd': return isDate(cell.v) ? new Date(Date.UTC(cell.v.getFullYear(), cell.v.getMonth(), cell.v.getDate())) : null;
    case 'b': return cell.v ? 'TRUE' : 'FALSE';
    case 'n':
      if (cell.z && XLSX.SSF.is_date(cell.z)) return serialToDate(cell.v, date1904);
      return cell.v;
    default: {
      const s = String(cell.v ?? '').replace(/ /g, ' ').trim();
      return s === '' ? null : s;
    }
  }
}

function sheetToGrid(ws, date1904) {
  let maxR = -1, maxC = -1, minR = Infinity, minC = Infinity;
  for (const k of Object.keys(ws)) {
    if (k[0] === '!') continue;
    const a = XLSX.utils.decode_cell(k);
    const c = ws[k];
    if (!c || c.t === 'z' || (c.v === undefined && c.f === undefined)) continue;
    if (a.r > maxR) maxR = a.r; if (a.c > maxC) maxC = a.c;
    if (a.r < minR) minR = a.r; if (a.c < minC) minC = a.c;
  }
  if (maxR < 0) return { grid: [], filled: [], r0: 0, c0: 0 };
  const grid = [];
  for (let R = minR; R <= maxR; R++) {
    const row = new Array(maxC - minC + 1);
    for (let C = minC; C <= maxC; C++) row[C - minC] = cellValue(ws[XLSX.utils.encode_cell({ r: R, c: C })], date1904);
    grid.push(row);
  }
  let filled = grid;
  const merges = ws['!merges'] || [];
  if (merges.length) {
    filled = grid.map((r) => r.slice());
    for (const m of merges) {
      const r0 = m.s.r - minR, c0 = m.s.c - minC;
      if (r0 < 0 || c0 < 0 || r0 >= grid.length) continue;
      const v = grid[r0][c0];
      for (let r = r0; r <= Math.min(m.e.r - minR, grid.length - 1); r++) {
        for (let c = c0; c <= Math.min(m.e.c - minC, filled[r].length - 1); c++) filled[r][c] = v;
      }
    }
  }
  return { grid, filled, r0: minR, c0: minC };
}

async function readWorkbook(file) {
  const isText = /\.(csv|txt|tsv)$/i.test(file.name);
  const buf = await file.arrayBuffer();
  if (isText) {
    let text = new TextDecoder('utf-8').decode(buf);
    if (text.includes('�')) text = new TextDecoder('windows-1252').decode(buf);
    return XLSX.read(text.replace(/^﻿/, ''), { type: 'string', raw: true });
  }
  return XLSX.read(buf, { type: 'array', cellNF: true, cellDates: false });
}

function detectHeaderRow(grid) {
  let best = 0, bestScore = 0;
  const lim = Math.min(grid.length, 30);
  for (let i = 0; i < lim; i++) {
    const strs = grid[i].filter((v) => typeof v === 'string' && !toNumber(v));
    const distinct = new Set(strs.map((s) => s.toLowerCase())).size;
    if (distinct > bestScore) { bestScore = distinct; best = i; }
  }
  return best;
}

/** Recognizes reference lists (AVL/ACL and similar) so they are mapped to the
 *  master instead of being merged into it as rows — wherever they were dropped. */
function looksLikeRefList(src) {
  const names = src.fileName + ' | ' + src.sheetName;
  if (/(^|[^a-z0-9])(avl|acl)([^a-z0-9]|$)/i.test(names)) return true;
  if (/approved\s+(vendor|contractor|supplier)|vendor\s+list|contractor\s+list|supplier\s+list/i.test(names)) return true;
  for (const row of src.grid.slice(0, 3)) {
    for (const v of row) {
      if (typeof v === 'string' && /approved\s+(vendor|contractor|supplier)\s+list/i.test(v)) return true;
    }
  }
  return false;
}

async function addFiles(fileList, role) {
  const files = [...fileList];
  if (!files.length) return;
  busy(true, 'Reading files…');
  await nextFrame();
  let added = 0;
  const errors = [];
  const newSources = [];
  const autoRef = [];
  for (const f of files) {
    try {
      const wb = await readWorkbook(f);
      const date1904 = !!(wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904);
      for (const name of wb.SheetNames) {
        const ws = wb.Sheets[name];
        if (!ws) continue;
        const g = sheetToGrid(ws, date1904);
        const hidden = !!(wb.Workbook && wb.Workbook.Sheets && (wb.Workbook.Sheets.find((s) => s.name === name) || {}).Hidden);
        const src = {
          id: nextId++, fileName: f.name, sheetName: /\.(csv|txt|tsv)$/i.test(f.name) ? '(CSV)' : name,
          grid: g.grid, filled: g.filled, r0: g.r0, c0: g.c0,
          headerIdx: 0, include: g.grid.length > 0 && !hidden, hidden, role: role || 'data'
        };
        src.headerIdx = detectHeaderRow(state.opts.fillMerged ? src.filled : src.grid);
        if (!role && src.role !== 'ref' && looksLikeRefList(src)) { src.role = 'ref'; autoRef.push(src.fileName + ' › ' + src.sheetName); }
        const pref = state.sheetPrefs[sheetKey(src)];
        if (pref) {
          if (pref.include !== undefined) src.include = pref.include;
          if (pref.role) src.role = pref.role;
          if (pref.headerRow) src.headerIdx = Math.max(0, Math.min(src.grid.length - 1, pref.headerRow - src.r0 - 1));
        }
        state.sources.push(src);
        newSources.push(src);
        added++;
      }
    } catch (e) {
      console.error(e);
      errors.push(`${f.name}: ${e.message || e}`);
    }
  }
  busy(false);
  markDuplicateCopies(newSources);
  if (errors.length) toast('Could not read ' + errors.join('; '), true);
  else if (autoRef.length) toast(`Loaded ${added} sheet(s). Recognized as connected list(s), mapped to the master instead of merged into it: ${autoRef.join(', ')}. Change under “Use as” if this is wrong.`);
  else if (added) toast(`Loaded ${added} sheet${added > 1 ? 's' : ''} from ${files.length} file${files.length > 1 ? 's' : ''}`);
  refresh('sources');
}

/** Sheets whose headers and rows exactly match an earlier sheet (e.g. the same data sorted differently) are excluded. */
function markDuplicateCopies(newSources) {
  const sig = (s) => {
    analyseSource(s);
    if (!s.dataRows.length) return null;
    let acc = 0;
    for (const r of s.dataRows) {
      const str = JSON.stringify(r);
      let h = 2166136261;
      for (let i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 16777619); }
      acc = (acc + (h >>> 0)) % 4294967296;
    }
    return s.headers.map((h) => h.norm).sort().join(',') + '#' + s.dataRows.length + '#' + acc;
  };
  const seen = new Map();
  for (const s of state.sources) {
    const k = sig(s);
    if (!k) { s.dupOf = null; continue; }
    const first = seen.get(k);
    if (first && first !== s) {
      s.dupOf = `${first.fileName} › ${first.sheetName}`;
      const pref = state.sheetPrefs[sheetKey(s)] || {};
      if (newSources.includes(s) && pref.include === undefined) s.include = false;
    } else {
      s.dupOf = null;
      seen.set(k, s);
    }
  }
}

async function loadTemplate(file) {
  try {
    const wb = await readWorkbook(file);
    let headers = null;
    for (const name of wb.SheetNames) {
      const g = sheetToGrid(wb.Sheets[name], false);
      if (!g.grid.length) continue;
      const hi = detectHeaderRow(g.filled);
      const seen = new Set();
      headers = g.filled[hi].filter((v) => !isBlank(v)).map((v) => String(v).replace(/\s+/g, ' ').trim())
        .filter((h) => { const n = norm(h); if (!n || seen.has(n)) return false; seen.add(n); return true; })
        .map((h) => ({ name: h, norm: norm(h) }));
      if (headers.length) break;
    }
    if (!headers || !headers.length) throw new Error('no column headers found');
    state.template = { fileName: file.name, headers };
    // Template order and names take priority.
    state.colOrder = [...headers.map((h) => h.norm), ...state.colOrder.filter((k) => !headers.some((h) => h.norm === k))];
    for (const h of headers) state.colPrefs[h.norm] = { ...(state.colPrefs[h.norm] || {}), name: h.name, include: true };
    for (const k of Object.keys(state.colPrefs)) if (!headers.some((h) => h.norm === k) && state.colPrefs[k].include === undefined) state.colPrefs[k].include = false;
    toast(`Master layout loaded: ${headers.length} columns`);
    refresh('sources');
  } catch (e) {
    toast(`Could not read layout file ${file.name}: ${e.message || e}`, true);
  }
}

// ---------------------------------------------------------------------------
// Analysis: headers, data rows, columns
// ---------------------------------------------------------------------------
const gridOf = (s) => (state.opts.fillMerged ? s.filled : s.grid);
const included = () => state.sources.filter((s) => s.include && s.role !== 'ref');
const refSources = () => state.sources.filter((s) => s.include && s.role === 'ref');

function analyseSource(src) {
  const g = gridOf(src);
  src.headers = [];
  src.dataRows = [];
  src.stats = { blank: 0, totals: 0, repeats: 0 };
  const hi = src.headerIdx;
  if (hi < 0 || hi >= g.length) return;
  const hrow = g[hi];
  const width = g.reduce((w, r) => Math.max(w, r.length), 0);
  const used = {};
  for (let c = 0; c < width; c++) {
    let v = hrow[c];
    let name = isBlank(v) ? '' : (isDate(v) ? fmtDate(v) : String(v)).replace(/\s+/g, ' ').trim();
    if (!name) {
      let hasData = false;
      for (let r = hi + 1; r < g.length && !hasData; r++) if (!isBlank(g[r][c])) hasData = true;
      if (!hasData) continue;
      name = 'Column ' + colLetter(src.c0 + c);
    }
    let n = norm(name) || 'col' + colLetter(src.c0 + c).toLowerCase();
    // Header cells merged across columns repeat the same text; keep only the first.
    if (used[n] && state.opts.fillMerged && src.grid[hi][c] === null && g[hi][c - 1] === v) continue;
    if (used[n]) { used[n]++; name += ` (${used[n]})`; n += String(used[n]); } else used[n] = 1;
    src.headers.push({ name, norm: n, c });
  }
  const H = src.headers;
  for (let r = hi + 1; r < g.length; r++) {
    const row = g[r];
    const cells = H.map((h) => (row[h.c] === undefined ? null : row[h.c]));
    let first = null, count = 0, match = 0;
    for (let i = 0; i < cells.length; i++) {
      if (cells[i] === null) continue;
      if (first === null) first = cells[i];
      count++;
      if (norm(cells[i]) === H[i].norm) match++;
    }
    if (!count) { src.stats.blank++; continue; }
    if (match >= Math.max(2, H.length * 0.6)) { src.stats.repeats++; continue; }
    if (state.opts.dropTotals && typeof first === 'string' && TOTAL_RE.test(first)) { src.stats.totals++; continue; }
    src.dataRows.push(cells);
  }
}

function reconcileColumns() {
  const observed = new Map();
  for (const s of included()) {
    for (const h of s.headers) {
      const o = observed.get(h.norm);
      if (o) { o.count++; } else observed.set(h.norm, { name: h.name, count: 1 });
    }
  }
  const cols = [];
  const byKey = new Map();
  const make = (key, name, fromTemplate) => {
    const c = { key, defaultName: name, fromTemplate, aliases: [] };
    cols.push(c); byKey.set(key, c);
    return c;
  };
  if (state.template) for (const h of state.template.headers) make(h.norm, h.name, true);
  // Also keep columns that exist in saved order (from loaded settings) so overrides can target them.
  const normMap = new Map();
  const attach = (c, n, o, auto) => { c.aliases.push({ norm: n, name: o.name, count: o.count, auto }); normMap.set(n, c.key); };
  const deferred = [];
  for (const [n, o] of observed) {
    const ov = state.aliasOverrides[n];
    if (ov && ov !== n) { deferred.push([n, o, ov]); continue; }
    if (byKey.has(n)) { attach(byKey.get(n), n, o, false); continue; }
    let best = null, bestScore = 0;
    if (!ov && n.length >= 5) {
      for (const c of cols) {
        if (c.key.length < 5 || digits(c.key) !== digits(n)) continue;
        const th = c.fromTemplate ? 0.8 : 0.86;
        const sc = similarity(n, c.key);
        if (sc >= th && sc > bestScore) { best = c; bestScore = sc; }
      }
    }
    if (best) attach(best, n, o, true);
    else attach(make(n, o.name, false), n, o, false);
  }
  for (const [n, o, ov] of deferred) {
    const c = byKey.get(ov) || make(ov, o.name, false);
    attach(c, n, o, false);
  }
  // Order: saved order first, then new columns in discovery order.
  const orderIdx = new Map(state.colOrder.map((k, i) => [k, i]));
  const discovered = cols.map((c) => c.key);
  cols.sort((a, b) => {
    const ia = orderIdx.has(a.key) ? orderIdx.get(a.key) : 1e6 + discovered.indexOf(a.key);
    const ib = orderIdx.has(b.key) ? orderIdx.get(b.key) : 1e6 + discovered.indexOf(b.key);
    return ia - ib;
  });
  for (const c of cols) {
    const p = state.colPrefs[c.key] || {};
    c.name = p.name || c.defaultName;
    const isSerial = SERIAL_RE.test(c.defaultName) && state.opts.addSN;
    c.include = p.include !== undefined ? p.include : (state.template ? c.fromTemplate : !isSerial);
    c.type = p.type || 'auto';
  }
  state.colOrder = cols.map((c) => c.key).concat(state.colOrder.filter((k) => !byKey.has(k)));
  state.columns = cols;
  state.normMap = normMap;
}

const MEASURE_RE = /(value|amount|price|cost|spend|revenue|sales|budget|qty|quantity|stock|consumption|balance|total|count|hours|days|weight|volume|rate)/i;
function detectType(c, vals) {
  let total = 0, dates = 0, nums = 0, ints = 0, pcts = 0, big = 0;
  const distinct = new Set();
  for (const v of vals) {
    if (isBlank(v) || (typeof v === 'string' && PLACEHOLDER_RE.test(v))) continue;
    total++;
    if (isDate(v) || parseDateStr(v)) { dates++; continue; }
    const x = toNumber(v);
    if (x) { nums++; if (Number.isInteger(x.n)) ints++; if (x.pct) pcts++; if (x.n >= 1000) big++; distinct.add(x.n); }
  }
  if (!total) return 'text';
  if (dates / total >= 0.8) return 'date';
  if (nums / total >= 0.8 && !ID_RE.test(c.name) && !ID_RE.test(c.defaultName)) {
    if (pcts / nums > 0.5) return 'pct';
    const idLike = ints === nums && nums >= 10 && big === nums && distinct.size / nums > 0.95
      && !MEASURE_RE.test(c.name) && !MEASURE_RE.test(c.defaultName);
    if (idLike) return 'text';
    return ints === nums ? 'int' : 'dec';
  }
  return 'text';
}

function coerce(v, type) {
  if (isBlank(v)) return { v: null, ok: true };
  if (type === 'text') return { v: isDate(v) ? fmtDate(v) : v, ok: true };
  if (typeof v === 'string' && PLACEHOLDER_RE.test(v)) return { v, ok: true };
  if (type === 'date') {
    if (isDate(v)) return { v, ok: true };
    if (typeof v === 'number' && v > 0 && v < 2958466) return { v: serialToDate(v), ok: true };
    const d = parseDateStr(String(v));
    return d ? { v: d, ok: true } : { v, ok: false };
  }
  if (isDate(v)) return { v: fmtDate(v), ok: false };
  const x = toNumber(v);
  if (!x) return { v, ok: false };
  let n = x.n;
  if (type === 'pct' && (x.pct || Math.abs(n) > 1)) n /= 100;
  return { v: n, ok: true };
}

function defaultConnName(src) {
  const stem = src.fileName.replace(/\.[^.]+$/, '');
  const m = stem.match(/[A-Za-z]{2,}[A-Za-z&]*/);  // first real word, skipping upload prefixes
  return (m ? m[0].toUpperCase() : stem.slice(0, 20)) || 'List';
}
const BRING_CAP = 20;   // distinct values carried into one master cell
function syncConnections() {
  const prev = new Map(state.connections.map((c) => [c.key, c]));
  state.connections = refSources().map((src) => {
    const key = sheetKey(src);
    const old = prev.get(key) || state.connPrefs[key] || {};
    const has = (n) => n && src.headers.some((h) => h.norm === n);
    const notKey = (h) => !/(number|no\.?$|\bid\b|code)/i.test(h.name);
    const nameCol = src.headers.find((h) => /\bname\b/i.test(h.name) && notKey(h)) || src.headers.find((h) => /desc/i.test(h.name) && notKey(h));
    return {
      key, srcId: src.id,
      name: old.name || defaultConnName(src),
      refKeyNorm: has(old.refKeyNorm) ? old.refKeyNorm
        : (src.headers.find((h) => /(number|no\.?$|\bid\b|code)/i.test(h.name)) || src.headers[0] || {}).norm,
      masterKey: old.masterKey || '',
      refKeyNorm2: has(old.refKeyNorm2) ? old.refKeyNorm2 : '',
      masterKey2: old.masterKey2 || '',
      bring: old.bring !== undefined ? (has(old.bring) ? old.bring : '') : (nameCol ? nameCol.norm : ''),
      addCount: old.addCount !== false,
      exportSheet: !!old.exportSheet,
      stats: null
    };
  });
}
/** Builds keyNorm -> { count, vals:Set of bring-column values } for one ref key column. */
function connIndex(conn, keyColNorm) {
  const src = state.sources.find((x) => x.id === conn.srcId);
  const hi = src ? src.headers.findIndex((h) => h.norm === keyColNorm) : -1;
  if (!src || hi < 0) return null;
  const bi = conn.bring ? src.headers.findIndex((h) => h.norm === conn.bring) : -1;
  const idx = new Map();
  for (const row of src.dataRows) {
    const k = keyNorm(row[hi]);
    if (!k) continue;
    let e = idx.get(k);
    if (!e) { e = { count: 0, vals: bi >= 0 ? new Set() : null }; idx.set(k, e); }
    e.count++;
    if (bi >= 0 && !isBlank(row[bi])) e.vals.add(String(isDate(row[bi]) ? fmtDate(row[bi]) : row[bi]).trim());
  }
  return idx;
}
/** Scores every (master column, list column) pairing by sampled hit rate.
 *  Returns pairs sorted best-first — the engine picks from these, and the
 *  connection card shows them so a failed match is never a mystery. */
function pairScan(src, defs, records, index) {
  const step = Math.max(1, Math.floor(records.length / 300));
  const sample = [];
  for (let i = 0; i < records.length && sample.length < 300; i += step) sample.push(records[i]);
  // candidate keys per master column, computed once
  const candsByDef = new Map();
  for (const d of defs) {
    if (d.source || d.conn) continue;
    const mi = index[d.key];
    let nonblank = 0;
    const cands = sample.map((r) => {
      if (isBlank(r[mi])) return null;
      nonblank++;
      return matchCandidates(r[mi]);
    });
    if (nonblank >= 3) candsByDef.set(d, cands);
  }
  const pairs = [];
  for (const h of src.headers) {
    const hi = src.headers.indexOf(h);
    const keys = new Set();
    for (const row of src.dataRows) { const k = keyNorm(row[hi]); if (k) keys.add(k); }
    if (!keys.size) continue;
    for (const [d, cands] of candsByDef) {
      let hits = 0;
      for (const c of cands) if (c && c.some((k) => keys.has(k))) hits++;
      if (hits >= 3) pairs.push({ refNorm: h.norm, refName: h.name, masterKey: d.key, masterName: d.name, hits, pct: hits / sample.length });
    }
  }
  pairs.sort((a, b) => b.hits - a.hits);
  return pairs;
}
function joinVals(set) {
  if (!set || !set.size) return null;
  const arr = [...set].sort((a, b) => a.localeCompare(b));
  return arr.length > BRING_CAP ? arr.slice(0, BRING_CAP).join('; ') + `; (+${arr.length - BRING_CAP} more)` : arr.join('; ');
}
/** Appends match-count and brought-value columns per active connection; records gain values before __file. */
function applyConnections(defs, index, records) {
  syncConnections();
  const fi = defs.findIndex((d) => d.key === '__file');
  let at = fi;
  for (const conn of state.connections) {
    const src = state.sources.find((x) => x.id === conn.srcId);
    const userPref = state.connPrefs[conn.key] || {};
    const userPinned = userPref.masterKey && index[userPref.masterKey] !== undefined && !(defs[index[userPref.masterKey]] || {}).conn;
    conn.diag = null;
    if (!userPinned && src && records.length) {
      // Full scan: every master column against every list column; take the best
      // pairing, and a second one as the fallback rule when it uses different columns.
      const pairs = pairScan(src, defs, records, index);
      conn.diag = pairs.slice(0, 4);
      if (pairs.length) {
        conn.refKeyNorm = pairs[0].refNorm;
        conn.masterKey = pairs[0].masterKey;
        if (!userPref.refKeyNorm2) {
          const second = pairs.find((p) => p.masterKey !== pairs[0].masterKey && p.refNorm !== pairs[0].refNorm);
          if (second) { conn.refKeyNorm2 = second.refNorm; conn.masterKey2 = second.masterKey; }
        }
      } else {
        conn.masterKey = '';
      }
    }
    const idx = connIndex(conn, conn.refKeyNorm);
    if (!idx) { conn.stats = { error: 'key column not found' }; continue; }
    const refRows = [...idx.values()].reduce((a, b) => a + b.count, 0);
    const mi = index[conn.masterKey];
    const idx2 = conn.refKeyNorm2 ? connIndex(conn, conn.refKeyNorm2) : null;
    const mi2 = conn.masterKey2 ? index[conn.masterKey2] : undefined;
    const rule2 = idx2 && mi2 !== undefined && !(defs[mi2] || {}).conn;
    if (mi === undefined && !rule2) { conn.stats = { refRows, keys: idx.size, nomatch: true }; continue; }
    let matched = 0, via1 = 0, via2 = 0;
    const counts = [], brings = [];
    for (const r of records) {
      let hit = null, rule = 0;
      if (mi !== undefined) for (const k of matchCandidates(r[mi])) { const e = idx.get(k); if (e) { hit = e; rule = 1; break; } }
      if (!hit && rule2) for (const k of matchCandidates(r[mi2])) { const e = idx2.get(k); if (e) { hit = e; rule = 2; break; } }
      if (hit) { matched++; if (rule === 1) via1++; else via2++; }
      counts.push(hit ? hit.count : 0);
      brings.push(hit ? joinVals(hit.vals) : null);
    }
    conn.stats = { refRows, keys: idx.size, matched, unmatched: records.length - matched, via1, via2: rule2 ? via2 : null };
    if (!conn.addCount) continue;
    defs.splice(at, 0, { key: 'conn:' + conn.key, name: conn.name + ' Matches', type: 'int', conn: true });
    records.forEach((r, i) => r.splice(at, 0, counts[i]));
    at++;
    if (conn.bring) {
      const src = state.sources.find((x) => x.id === conn.srcId);
      const bh = src && src.headers.find((h) => h.norm === conn.bring);
      defs.splice(at, 0, { key: 'connv:' + conn.key, name: conn.name + ' ' + (bh ? bh.name : 'Values'), type: 'text', conn: true, bring: true });
      records.forEach((r, i) => r.splice(at, 0, brings[i]));
      at++;
    }
  }
  Object.keys(index).forEach((k) => delete index[k]);
  defs.forEach((d, i) => { index[d.key] = i; });
}

function buildMaster() {
  const cols = state.columns.filter((c) => c.include);
  const keyIdx = new Map(cols.map((c, i) => [c.key, i]));
  const records = [];
  let onlyExcluded = 0;
  const srcCounts = new Map();   // sheetKey -> { onlyExcluded, dupes }
  for (const s of included()) {
    const sc = { onlyExcluded: 0, dupes: 0 };
    srcCounts.set(sheetKey(s), sc);
    const map = s.headers.map((h) => keyIdx.get(state.normMap.get(h.norm)));
    for (const cells of s.dataRows) {
      const vals = new Array(cols.length + 2).fill(null);
      let any = false;
      for (let i = 0; i < cells.length; i++) {
        const k = map[i];
        if (k === undefined || cells[i] === null) continue;
        if (vals[k] === null) vals[k] = cells[i];
        any = true;
      }
      if (!any) { onlyExcluded++; sc.onlyExcluded++; continue; }
      vals[cols.length] = s.fileName;
      vals[cols.length + 1] = s.sheetName;
      records.push(vals);
    }
  }
  const issues = [];
  cols.forEach((c, i) => {
    c.resolved = c.type === 'auto' ? detectType(c, records.map((r) => r[i])) : c.type;
    let bad = 0; const examples = [];
    for (const r of records) {
      const res = coerce(r[i], c.resolved);
      r[i] = res.v;
      if (!res.ok) { bad++; if (examples.length < 3) examples.push(String(res.v)); }
    }
    if (bad) issues.push({ col: c.name, count: bad, examples });
  });

  // Snapshot of every column's total before de-duplication — the verification layer
  // reconciles these against the final master.
  const rawSums = cols.map((c, i) => {
    if (c.resolved !== 'int' && c.resolved !== 'dec') return null;
    let t = 0; for (const r of records) if (typeof r[i] === 'number') t += r[i];
    return t;
  });
  const removedSums = cols.map(() => 0);
  let dupes = 0;
  let rows = records;
  if (state.opts.dedupe) {
    const seen = new Set();
    rows = records.filter((r) => {
      const k = JSON.stringify(r.slice(0, cols.length).map((v) => (isDate(v) ? 'd' + v.getTime() : typeof v === 'string' ? v.toLowerCase() : v)));
      if (seen.has(k)) {
        dupes++;
        srcCounts.get(sheetKey({ fileName: r[cols.length], sheetName: r[cols.length + 1] })).dupes++;
        cols.forEach((c, i) => { if (rawSums[i] !== null && typeof r[i] === 'number') removedSums[i] += r[i]; });
        return false;
      }
      seen.add(k); return true;
    });
  }

  const defs = cols.map((c) => ({ key: c.key, name: c.name, type: c.resolved }));
  defs.push({ key: '__file', name: 'Source File', type: 'text', source: true });
  defs.push({ key: '__sheet', name: 'Source Sheet', type: 'text', source: true });

  // Codes embedded in text (e.g. "INSPECTION, OVERHAUL (321110)") get their own
  // column — auto-enabled when at least half of a column's values end in a
  // numeric code in parentheses; always switchable per column in step 2.
  for (let j = cols.length - 1; j >= 0; j--) {
    const c = cols[j];
    if (c.resolved !== 'text') { c.xcode = false; continue; }
    let on = (state.colPrefs[c.key] || {}).xcode;
    if (on === undefined) {
      let hits = 0, tot = 0;
      for (let k = 0; k < rows.length && tot < 200; k++) {
        const v = rows[k][j];
        if (isBlank(v)) continue;
        tot++;
        if (/\(\s*\d{3,}\s*\)\s*$/.test(String(v))) hits++;
      }
      on = tot >= 5 && hits / tot >= 0.5;
    }
    c.xcode = !!on;
    if (!on) continue;
    defs.splice(j + 1, 0, { key: c.key + ':code', name: c.name.replace(/\s+code$/i, '') + ' Code', type: 'text', derived: true });
    for (const r of rows) r.splice(j + 1, 0, extractCode(r[j]));
  }
  const index = {};
  defs.forEach((d, i) => { index[d.key] = i; });

  const si = index[state.opts.sortBy];
  if (si !== undefined) {
    const dir = state.opts.sortDir === 'desc' ? -1 : 1;
    rows = rows.map((r, i) => [r, i]).sort((a, b) => {
      const x = a[0][si], y = b[0][si];
      if (isBlank(x) && isBlank(y)) return a[1] - b[1];
      if (isBlank(x)) return 1;
      if (isBlank(y)) return -1;
      return (cmpVals(sortVal(x), sortVal(y)) * dir) || (a[1] - b[1]);
    }).map((p) => p[0]);
  }

  applyConnections(defs, index, rows);

  const stats = { blank: 0, totals: 0, repeats: 0, onlyExcluded, dupes };
  for (const s of included()) { stats.blank += s.stats.blank; stats.totals += s.stats.totals; stats.repeats += s.stats.repeats; }
  const sums = {};
  cols.forEach((c, i) => { if (rawSums[i] !== null) sums[c.key] = { name: c.name, type: c.resolved, raw: rawSums[i], removed: removedSums[i] }; });
  state.master = { defs, index, records: rows, issues, stats, srcCounts, rawSums: sums };
}

/** Columns as they appear in the exported master sheet. */
function masterView(records) {
  const M = state.master;
  const out = [];
  if (state.opts.addSN) out.push({ name: 'S/N', type: 'int', get: (r, i) => i + 1, sn: true });
  M.defs.forEach((d, i) => {
    if (d.source && !state.opts.addSource) return;
    out.push({ name: d.name, type: d.type, get: (r) => r[i] });
  });
  return { defs: out, rows: records.map((r, i) => out.map((d) => d.get(r, i))) };
}

// ---------------------------------------------------------------------------
// Summary
// ---------------------------------------------------------------------------
function measureLabel(m, M) {
  if (m.agg === 'count') return 'No. of Records';
  const name = M.defs[M.index[m.col]].name;
  return { counta: `Count of ${name}`, distinct: `Unique ${name}`, list: `${name} (distinct list)`, sum: `Total ${name}`, avg: `Average ${name}`, min: `Minimum ${name}`, max: `Maximum ${name}` }[m.agg];
}
function measureType(m, M) {
  if (m.agg === 'list') return 'text';
  if (m.agg === 'count' || m.agg === 'counta' || m.agg === 'distinct') return 'int';
  const t = M.defs[M.index[m.col]].type;
  if (m.agg === 'avg' && t === 'int') return 'dec';
  return t;
}
function aggregate(rows, m, M) {
  if (m.agg === 'count') return rows.length;
  const ci = M.index[m.col];
  if (m.agg === 'counta') { let n = 0; for (const r of rows) if (!isBlank(r[ci])) n++; return n; }
  if (m.agg === 'distinct') {
    const s = new Set();
    for (const r of rows) { const v = r[ci]; if (!isBlank(v)) s.add(isDate(v) ? v.getTime() : typeof v === 'string' ? v.trim().toLowerCase() : v); }
    return s.size;
  }
  if (m.agg === 'list') {
    // Values carried over from a connected list are already "a; b; c" — split them
    // so the result is a true union of distinct items.
    const seen = new Map();
    for (const r of rows) {
      const v = r[ci];
      if (isBlank(v)) continue;
      for (const part of String(isDate(v) ? fmtDate(v) : v).split('; ')) {
        const t = part.trim();
        if (t && !/^\(\+\d+ more\)$/.test(t)) seen.set(t.toLowerCase(), t);
      }
    }
    const arr = [...seen.values()].sort((a, b) => a.localeCompare(b));
    if (!arr.length) return null;
    return arr.length > LIST_CAP ? arr.slice(0, LIST_CAP).join('; ') + ` (+${arr.length - LIST_CAP} more)` : arr.join('; ');
  }
  let sum = 0, n = 0, mn = Infinity, mx = -Infinity, dateCol = false;
  for (const r of rows) {
    const v = r[ci];
    let x = null;
    if (typeof v === 'number') x = v; else if (isDate(v)) { x = v.getTime(); dateCol = true; }
    if (x === null) continue;
    sum += x; n++; if (x < mn) mn = x; if (x > mx) mx = x;
  }
  if (m.agg === 'sum') return sum;
  if (!n) return null;
  if (m.agg === 'avg') return dateCol ? null : sum / n;
  const val = m.agg === 'min' ? mn : mx;
  return dateCol ? new Date(val) : val;
}

function groupKey(v, mode) {
  if (isBlank(v)) return { k: '￿', label: '(Blank)', sort: null, blank: true };
  if (mode && mode !== 'value') {
    if (!isDate(v)) return { k: '￾', label: '(No date)', sort: null, blank: true };
    const y = v.getUTCFullYear(), m = v.getUTCMonth();
    if (mode === 'month') return { k: `${y}-${m}`, label: `${MONTHS[m]} ${y}`, sort: y * 12 + m };
    if (mode === 'quarter') { const q = Math.floor(m / 3) + 1; return { k: `${y}-Q${q}`, label: `Q${q} ${y}`, sort: y * 4 + q }; }
    return { k: String(y), label: String(y), sort: y };
  }
  if (isDate(v)) return { k: 'd' + v.getTime(), label: v, sort: v.getTime() };
  if (typeof v === 'number') return { k: 'n' + v, label: v, sort: v };
  const s = String(v).trim();
  return { k: 's' + s.toLowerCase(), label: s, sort: s };
}
function groupRows(records, ci, mode) {
  const groups = new Map();
  for (const r of records) {
    const g = groupKey(r[ci], mode);
    let e = groups.get(g.k);
    if (!e) { e = { ...g, rows: [] }; groups.set(g.k, e); }
    e.rows.push(r);
  }
  return [...groups.values()];
}
function sortGroups(list, sort, keyFn) {
  return list.sort((a, b) => {
    if (a.blank !== b.blank) return a.blank ? 1 : -1;
    if (sort === 'desc') {
      const d = (keyFn(b) ?? -Infinity) - (keyFn(a) ?? -Infinity);
      if (d) return d;
    }
    return cmpVals(a.sort ?? '', b.sort ?? '');
  });
}

const MODE_LABEL = { month: 'Month', quarter: 'Quarter', year: 'Year' };

function breakdownBlock(bd, M) {
  const gi = M.index[bd.groupBy];
  if (gi === undefined) return null;
  const measures = (bd.measures || []).filter((m) => m.agg === 'count' || M.index[m.col] !== undefined);
  if (!measures.length) measures.push({ agg: 'count' });
  const mode = M.defs[gi].type === 'date' ? (bd.groupMode || 'value') : 'value';
  const gName = M.defs[gi].name + (MODE_LABEL[mode] ? ` (${MODE_LABEL[mode]})` : '');
  const sort = bd.sort || (mode !== 'value' || M.defs[gi].type === 'date' ? 'alpha' : 'desc');
  const groups = groupRows(M.records, gi, mode);
  const gType = mode === 'value' ? M.defs[gi].type : 'text';
  const si = bd.splitBy ? M.index[bd.splitBy] : undefined;

  if (si !== undefined && si !== gi) {
    const m = measures[0];
    const sName = M.defs[si].name;
    let splits = sortGroups(groupRows(M.records, si, 'value'), 'alpha', () => 0);
    let other = null;
    if (splits.length > 25) {
      const top = [...splits].sort((a, b) => b.rows.length - a.rows.length).slice(0, 24);
      const keep = new Set(top.map((s) => s.k));
      other = { k: '\u0000other', label: 'Other', rows: [], members: new Set() };
      for (const s of splits) if (!keep.has(s.k)) { other.rows.push(...s.rows); other.members.add(s.k); }
      splits = splits.filter((s) => keep.has(s.k)).concat(other);
    }
    const splitOf = (r) => { const k = groupKey(r[si], 'value').k; return other && other.members.has(k) ? other.k : k; };
    sortGroups(groups, sort, (g) => aggregate(g.rows, m, M));
    const t = measureType(m, M);
    const rows = groups.map((g) => {
      const bySplit = new Map();
      for (const r of g.rows) { const k = splitOf(r); if (!bySplit.has(k)) bySplit.set(k, []); bySplit.get(k).push(r); }
      return [g.label, ...splits.map((s) => (bySplit.has(s.k) ? aggregate(bySplit.get(s.k), m, M) : (m.agg === 'count' || m.agg === 'sum' || m.agg === 'counta' || m.agg === 'distinct' ? 0 : null))), aggregate(g.rows, m, M)];
    });
    return {
      title: `${measureLabel(m, M)} by ${gName} and ${sName}`,
      headers: [gName, ...splits.map((s) => (isDate(s.label) ? fmtDate(s.label) : String(s.label))), 'Total'],
      types: [gType, ...splits.map(() => t), t],
      rows,
      total: ['Total', ...splits.map((s) => aggregate(s.rows, m, M)), aggregate(M.records, m, M)]
    };
  }

  let gs = bd.hideBlank ? groups.filter((g) => !g.blank) : groups;
  const universe = bd.hideBlank ? [].concat(...gs.map((g) => g.rows)) : M.records;
  const vals = new Map(gs.map((g) => [g, measures.map((m) => aggregate(g.rows, m, M))]));
  sortGroups(gs, sort, (g) => { const v = vals.get(g)[0]; return typeof v === 'number' ? v : null; });
  const topN = Math.max(0, +bd.topN || 0);
  let note = '';
  if (topN > 0 && gs.length > topN) {
    const rest = gs.slice(topN);
    gs = gs.slice(0, topN);
    const other = { label: `Other (${rest.length.toLocaleString('en-US')} values)`, rows: [].concat(...rest.map((g) => g.rows)) };
    vals.set(other, measures.map((m) => aggregate(other.rows, m, M)));
    gs = gs.concat(other);
    note = ` — Top ${topN}`;
  }
  const totals = measures.map((m) => aggregate(universe, m, M));
  const pct = bd.pct !== false && ['count', 'sum', 'counta'].includes(measures[0].agg) && totals[0];
  const rows = gs.map((g) => {
    const v = vals.get(g);
    return [g.label, ...v, ...(pct ? [typeof v[0] === 'number' ? v[0] / totals[0] : null] : [])];
  });
  return {
    title: `By ${gName}${note}${bd.hideBlank ? ' (blank ignored)' : ''}`,
    headers: [gName, ...measures.map((m) => measureLabel(m, M)), ...(pct ? ['% of Total'] : [])],
    types: [gType, ...measures.map((m) => measureType(m, M)), ...(pct ? ['pct'] : [])],
    rows,
    total: [bd.hideBlank ? 'Total (excl. blank)' : 'Total', ...totals, ...(pct ? [1] : [])]
  };
}

function suggestBreakdowns() {
  const M = state.master;
  if (!M || !M.records.length) return [];
  const n = M.records.length;
  const dataDefs = M.defs.filter((d) => !d.source);
  const nums = dataDefs.filter((d) => (d.type === 'int' || d.type === 'dec') && !d.conn && !ID_RE.test(d.name)).slice(0, 3);
  const measures = [{ agg: 'count' }, ...nums.map((d) => ({ agg: 'sum', col: d.key }))];
  const out0 = [];
  // The headline report: which materials/items have the most purchase records,
  // and (when a connected list brings them in) who their vendors are.
  const matCol = dataDefs.filter((d) => d.type === 'text' && /(material|item|product|service)/i.test(d.name))
    .map((d) => ({ d, distinct: new Set(M.records.map((r) => groupKey(r[M.index[d.key]], 'value').k)).size,
      score: (/desc/i.test(d.name) ? 2 : 0) + (/material/i.test(d.name) ? 1 : 0) }))
    .filter((x) => x.distinct >= 3)
    .sort((a, b) => b.score - a.score || b.distinct - a.distinct)[0];
  if (matCol) {
    const bringDefs = M.defs.filter((d) => d.bring).slice(0, 2);
    const moneyDef = nums.find((d) => /value|amount|cost|price|spend/i.test(d.name)) || nums[0];
    out0.push({ id: nextId++, groupBy: matCol.d.key, groupMode: 'value', splitBy: '',
      measures: [{ agg: 'count' }, ...(moneyDef ? [{ agg: 'sum', col: moneyDef.key }] : []), ...bringDefs.map((d) => ({ agg: 'list', col: d.key }))],
      pct: true, sort: 'desc', topN: 15, hideBlank: true });
  }
  const cands = [];
  for (const d of dataDefs) {
    if (d.type !== 'text' || ID_RE.test(d.name)) continue;
    const distinct = new Set(M.records.map((r) => groupKey(r[M.index[d.key]], 'value').k)).size;
    if (distinct < 2 || distinct > 40 || distinct > Math.max(3, n * 0.5)) continue;
    let score = CAT_RE.test(d.name) ? 10 : 0;
    if (distinct <= 15) score += 3;
    cands.push({ d, score: score - distinct / 100 });
  }
  cands.sort((a, b) => b.score - a.score);
  const out = out0.concat(cands.slice(0, 3).map(({ d }) => ({ id: nextId++, groupBy: d.key, groupMode: 'value', splitBy: '', measures: measures.map((m) => ({ ...m })), pct: true, sort: 'desc' })));
  const dateDef = dataDefs.find((d) => d.type === 'date');
  if (dateDef) {
    const months = new Set(M.records.map((r) => groupKey(r[M.index[dateDef.key]], 'month').k));
    if (months.size > 1) out.push({ id: nextId++, groupBy: dateDef.key, groupMode: months.size > 24 ? 'year' : 'month', splitBy: '', measures: measures.map((m) => ({ ...m })), pct: true, sort: 'alpha' });
  }
  if (!out.length && included().length > 1) {
    out.push({ id: nextId++, groupBy: '__file', groupMode: 'value', splitBy: '', measures: measures.map((m) => ({ ...m })), pct: true, sort: 'alpha' });
  }
  return out;
}

/** Offline "analyst": deterministic plain-English findings written from the computed data. */
function analystFindings() {
  const M = state.master;
  const out = [];
  const n = M.records.length;
  if (!n) return out;
  const pc = (x) => (x / n * 100).toFixed(1) + '%';

  // 1. Top materials by number of purchase records, with their vendors.
  const topBd = state.breakdowns.map((bd) => breakdownBlock(bd, M)).find((b) => b && /material|item|product|service/i.test(b.headers[0]));
  if (topBd && topBd.rows.length) {
    const listCols = topBd.headers.map((h, i) => (topBd.types[i] === 'text' && i > 0 ? i : -1)).filter((i) => i > 0);
    topBd.rows.slice(0, 3).forEach((r, i) => {
      if (String(r[0]).startsWith('Other (')) return;
      const vendors = listCols.map((c) => r[c]).filter(Boolean).join('; ');
      out.push(`#${i + 1} ${topBd.headers[0]} by purchase records: "${r[0]}" with ${displayValue(r[1], 'int')} records (${pc(r[1])})` +
        (vendors ? `. Approved vendors: ${vendors}` : ''));
    });
  }

  // 2. Connected-list coverage and gaps.
  for (const conn of state.connections) {
    const st = conn.stats;
    if (!st || st.error || st.nomatch || st.matched === undefined) continue;
    out.push(`${conn.name}: ${st.matched.toLocaleString('en-US')} of ${n.toLocaleString('en-US')} records (${pc(st.matched)}) have a match in the list` +
      (st.unmatched ? `; ${st.unmatched.toLocaleString('en-US')} records (${pc(st.unmatched)}) have NO match — review whether these need an approved source.` : '.'));
    const ci = M.index['connv:' + conn.key];
    if (ci !== undefined && st.matched) {
      const freq = new Map();
      for (const r of M.records) {
        if (isBlank(r[ci])) continue;
        for (const part of String(r[ci]).split('; ')) {
          const t = part.trim();
          if (t && !/^\(\+\d+ more\)$/.test(t)) freq.set(t, (freq.get(t) || 0) + 1);
        }
      }
      const top = [...freq.entries()].sort((a, b) => b[1] - a[1]).slice(0, 3);
      if (top.length) out.push(`${conn.name}: most frequently available vendors across matched records — ` +
        top.map(([v, c]) => `${v} (available for ${c.toLocaleString('en-US')} records)`).join('; ') + '.');
    }
  }

  // 3. Concentration: biggest value column by group of the first breakdown.
  const money = M.defs.find((d) => !d.source && !d.conn && d.type === 'dec' && /value|amount|cost|price/i.test(d.name));
  if (money) {
    const mi = M.index[money.key];
    let tot = 0; for (const r of M.records) if (typeof r[mi] === 'number') tot += r[mi];
    if (tot > 0) {
      const sorted = M.records.map((r) => (typeof r[mi] === 'number' ? r[mi] : 0)).sort((a, b) => b - a);
      let acc = 0, k = 0;
      while (k < sorted.length && acc < tot * 0.8) acc += sorted[k++];
      out.push(`${money.name}: ${k.toLocaleString('en-US')} records (${pc(k)}) account for 80% of the total ${displayValue(tot, 'dec')}.`);
    }
  }

  // 4. Data quality worth knowing.
  for (const is of M.issues.slice(0, 2)) out.push(`Data quality: ${is.count.toLocaleString('en-US')} value(s) in "${is.col}" are not in the expected format (e.g. ${is.examples.join(', ')}).`);
  return out.slice(0, 8);
}

function computeSummary() {
  const M = state.master;
  const blocks = [];
  const srcs = included();
  const files = new Set(srcs.map((s) => s.fileName));
  const now = new Date();
  blocks.push({
    title: 'Report Overview', kv: true,
    headers: ['Item', 'Details'], types: ['text', 'auto'],
    rows: [
      ['Report', state.title || 'Master Report'],
      ['Generated on', fmtDateTime(now)],
      ['Files processed', files.size],
      ['Sheets processed', srcs.length],
      ['Total records', M.records.length],
      ['Columns in master sheet', M.defs.filter((d) => !d.source).length]
    ]
  });

  const findings = analystFindings();
  if (findings.length) {
    blocks.push({
      title: 'Analyst Highlights (generated automatically — verify before acting)',
      headers: ['#', 'Finding'], types: ['int', 'text'],
      rows: findings.map((f, i) => [i + 1, f])
    });
  }

  // Records by source
  const counts = new Map();
  for (const r of M.records) { const k = r[M.index.__file] + '\u0000' + r[M.index.__sheet]; counts.set(k, (counts.get(k) || 0) + 1); }
  const total = M.records.length;
  blocks.push({
    title: 'Records by Source',
    headers: ['Source File', 'Sheet', 'Records', '% of Total'], types: ['text', 'text', 'int', 'pct'],
    rows: srcs.map((s) => { const c = counts.get(s.fileName + '\u0000' + s.sheetName) || 0; return [s.fileName, s.sheetName, c, total ? c / total : 0]; }),
    total: ['Total', '', total, total ? 1 : 0]
  });

  // Key figures for numeric columns
  const numDefs = M.defs.filter((d) => !d.source && !d.conn && NUMERIC.has(d.type));
  if (numDefs.length) {
    blocks.push({
      title: 'Key Figures',
      headers: ['Column', 'Total', 'Average', 'Minimum', 'Maximum', 'Records with a Value'],
      types: ['text', 'auto', 'auto', 'auto', 'auto', 'int'],
      rows: numDefs.map((d) => {
        const t = d.type, avgT = t === 'int' ? 'dec' : t;
        const sum = aggregate(M.records, { agg: 'sum', col: d.key }, M);
        return [d.name,
          t === 'pct' ? '—' : { v: sum, t },
          { v: aggregate(M.records, { agg: 'avg', col: d.key }, M), t: avgT },
          { v: aggregate(M.records, { agg: 'min', col: d.key }, M), t },
          { v: aggregate(M.records, { agg: 'max', col: d.key }, M), t },
          aggregate(M.records, { agg: 'counta', col: d.key }, M)];
      })
    });
  }

  const conns = state.connections.filter((c) => c.stats && !c.stats.error);
  if (conns.length) {
    blocks.push({
      title: 'Connected Lists — Match Coverage',
      headers: ['Connected List', 'List Rows', 'Unique Keys', 'Master Rows Matched', 'Master Rows Not Matched', 'Coverage'],
      types: ['text', 'int', 'int', 'int', 'int', 'pct'],
      rows: conns.map((c) => c.stats.nomatch
        ? [c.name + ' (no matching master column found)', c.stats.refRows, c.stats.keys, 0, M.records.length, 0]
        : [c.name, c.stats.refRows, c.stats.keys, c.stats.matched, c.stats.unmatched, M.records.length ? c.stats.matched / M.records.length : 0])
    });
  }

  for (const bd of state.breakdowns) {
    const b = breakdownBlock(bd, M);
    if (b) blocks.push(b);
  }

  const S = M.stats;
  const q = [
    ['Blank rows skipped', S.blank],
    ['Repeated header rows skipped', S.repeats]
  ];
  if (state.opts.dropTotals) q.push(['Total / subtotal rows excluded', S.totals]);
  if (state.opts.dedupe) q.push(['Duplicate rows removed', S.dupes]);
  if (S.onlyExcluded) q.push(['Rows with data only in excluded columns', S.onlyExcluded]);
  for (const is of M.issues) q.push([`Values in "${is.col}" not in expected format`, is.count]);
  blocks.push({ title: 'Data Checks', headers: ['Check', 'Rows / Values'], types: ['text', 'int'], rows: q });
  state.summary = blocks;
}

// ---------------------------------------------------------------------------
// Verification: independent reconciliation of the output against the raw input
// ---------------------------------------------------------------------------
const OK = '\u2713 Tally';
const tallyNum = (d, type) => (d === 0 ? OK : '\u2717 Off by ' + displayValue(d, type || 'int'));
const tallyFloat = (d, scale) => (Math.abs(d) <= Math.max(1e-6, Math.abs(scale) * 1e-9) ? OK : '\u2717 Off by ' + displayValue(d, 'dec'));

function computeVerification() {
  const M = state.master;
  if (!M) { state.verification = null; return; }
  const blocks = [];

  // 1. Row reconciliation — master-side counts are recounted from the final records,
  //    raw-side counts straight from each sheet's grid.
  const fi = M.index.__file, shi = M.index.__sheet;
  const perSrc = new Map();
  for (const r of M.records) { const k = r[fi] + '|' + r[shi]; perSrc.set(k, (perSrc.get(k) || 0) + 1); }
  const rowRows = included().map((src) => {
    const g = gridOf(src);
    const raw = Math.max(0, g.length - src.headerIdx - 1);
    const sc = M.srcCounts.get(sheetKey(src)) || { onlyExcluded: 0, dupes: 0 };
    const skipped = src.stats.blank + src.stats.repeats + src.stats.totals + sc.onlyExcluded + sc.dupes;
    const inMaster = perSrc.get(src.fileName + '|' + src.sheetName) || 0;
    return [src.fileName, src.sheetName, raw, src.stats.blank, src.stats.repeats, src.stats.totals, sc.onlyExcluded + sc.dupes, inMaster, tallyNum(raw - skipped - inMaster)];
  });
  const tot = (i) => rowRows.reduce((a, r) => a + r[i], 0);
  blocks.push({
    title: 'Row Count Reconciliation — every raw row accounted for',
    headers: ['Source File', 'Sheet', 'Raw Rows', 'Blank', 'Repeated Headers', 'Total Rows Skipped', 'Duplicates / Unmapped', 'In Master', 'Check'],
    types: ['text', 'text', 'int', 'int', 'int', 'int', 'int', 'int', 'status'],
    rows: rowRows,
    total: ['Total', '', tot(2), tot(3), tot(4), tot(5), tot(6), tot(7), tallyNum(tot(2) - tot(3) - tot(4) - tot(5) - tot(6) - tot(7))]
  });
  const masterRecount = [...perSrc.values()].reduce((a, b) => a + b, 0);
  blocks.push({
    title: 'Master Sheet Check',
    headers: ['Check', 'Expected', 'Actual', 'Status'], types: ['text', 'int', 'int', 'status'],
    rows: [['Master records = sum of rows kept per source', tot(7), M.records.length, tallyNum(tot(7) - M.records.length)],
           ['Every master row carries its source', M.records.length, masterRecount, tallyNum(M.records.length - masterRecount)]]
  });

  // 2. Value reconciliation — totals of every numeric column, raw vs master.
  const valRows = [];
  for (const [key, info] of Object.entries(M.rawSums)) {
    const ci = M.index[key];
    if (ci === undefined) continue;
    let fin = 0;
    for (const r of M.records) if (typeof r[ci] === 'number') fin += r[ci];
    const d = info.raw - info.removed - fin;
    valRows.push([info.name, { v: info.raw, t: info.type }, { v: info.removed, t: info.type }, { v: fin, t: info.type }, tallyFloat(d, info.raw)]);
  }
  if (valRows.length) {
    blocks.push({
      title: 'Value Reconciliation — column totals, raw data vs master sheet',
      headers: ['Column', 'Raw Data Total', 'Removed Duplicates', 'Master Sheet Total', 'Status'],
      types: ['text', 'auto', 'auto', 'auto', 'status'],
      rows: valRows
    });
  }

  // 3. Summary cross-check — every breakdown total re-added from its own rows
  //    and compared with a direct recount of the master.
  const bdRows = [];
  for (const bd of state.breakdowns) {
    const b = breakdownBlock(bd, M);
    if (!b || !b.total) continue;
    b.headers.forEach((h, i) => {
      if (i === 0 || !['int', 'dec'].includes(b.types[i])) return;
      const colSum = b.rows.reduce((a, r) => a + (typeof r[i] === 'number' ? r[i] : 0), 0);
      const totalCell = typeof b.total[i] === 'number' ? b.total[i] : 0;
      bdRows.push([b.title, h, { v: colSum, t: b.types[i] }, { v: totalCell, t: b.types[i] }, tallyFloat(colSum - totalCell, colSum)]);
    });
  }
  if (bdRows.length) {
    blocks.push({
      title: 'Summary Cross-check — breakdown rows add up to their totals',
      headers: ['Summary Table', 'Column', 'Sum of Rows', 'Total Row', 'Status'],
      types: ['text', 'text', 'auto', 'auto', 'status'],
      rows: bdRows
    });
  }

  // 4. Extra sheets partition the master.
  const si = M.index[state.opts.splitBy];
  if (si !== undefined) {
    const groups = groupRows(M.records, si, M.defs[si].type === 'date' ? 'month' : 'value');
    const n = groups.reduce((a, g) => a + g.rows.length, 0);
    blocks.push({
      title: 'Extra Sheets Check',
      headers: ['Check', 'Expected', 'Actual', 'Status'], types: ['text', 'int', 'int', 'status'],
      rows: [[`Rows across the ${Math.min(groups.length, 100)} "${M.defs[si].name}" sheets = master records${groups.length > 100 ? ' (over 100 values — not all exported!)' : ''}`,
        M.records.length, n, groups.length > 100 ? '\u2717 Review' : tallyNum(M.records.length - n)]]
    });
  }

  // 5. Connected lists — coverage recounted from the match columns.
  const connRows = [];
  for (const conn of state.connections) {
    if (!conn.stats || conn.stats.error || conn.stats.nomatch) continue;
    const ci = M.index['conn:' + conn.key];
    if (ci === undefined) {
      connRows.push([conn.name, 'matched rows (column not added)', conn.stats.matched, conn.stats.matched, OK]);
      continue;
    }
    let m = 0;
    for (const r of M.records) if (typeof r[ci] === 'number' && r[ci] > 0) m++;
    connRows.push([conn.name, 'master rows with a match', conn.stats.matched, m, tallyNum(conn.stats.matched - m)]);
  }
  if (connRows.length) {
    blocks.push({
      title: 'Connected Lists Check',
      headers: ['Connected List', 'Check', 'Reported', 'Recounted', 'Status'],
      types: ['text', 'text', 'int', 'int', 'status'],
      rows: connRows
    });
  }

  const bad = blocks.reduce((a, b) => a + b.rows.concat(b.total ? [b.total] : []).filter((r) => String(r[r.length - 1]).startsWith('\u2717')).length, 0);
  state.verification = { blocks, bad };
}

// ---------------------------------------------------------------------------
// Refresh pipeline
// ---------------------------------------------------------------------------
function refresh(level) {
  try {
    if (level === 'sources') state.sources.forEach(analyseSource);
    if (level === 'sources' || level === 'columns') reconcileColumns();
    if (level !== 'summary' && level !== 'preview') {
      if (included().length) buildMaster(); else state.master = null;
      if (state.bdAuto) state.breakdowns = suggestBreakdowns();
    }
    if (state.master && level !== 'preview') { computeSummary(); computeVerification(); }
  } catch (e) {
    console.error(e);
    toast('Something went wrong: ' + (e.message || e), true);
  }
  render(level);
}

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------
function render(level) {
  const has = state.sources.length > 0;
  const ready = !!state.master;
  $('#step2').hidden = !has;
  $('#step3').hidden = !ready;
  $('#step4').hidden = !ready;
  $('#templateName').textContent = state.template ? `✓ ${state.template.fileName} (${state.template.headers.length} columns)` : '';
  if (level === 'sources' || level === 'columns' || level === 'all') { renderSources(); renderMasterOpts(); renderColumns(); }
  if (level !== 'preview' && level !== 'summary') renderBreakdowns();
  if (level === 'summary') renderBreakdowns();
  renderConnections();
  renderPreview();
  $('#title').value = state.title;
  $('#fileName').value = state.fileName;
}

function renderSources() {
  const el = $('#sources');
  if (!state.sources.length) { el.innerHTML = ''; return; }
  const rows = state.sources.map((s) => {
    const hdrRow = s.r0 + s.headerIdx + 1;
    const dropped = s.stats.totals + s.stats.repeats;
    return `<tr class="${s.include ? '' : 'off'}" data-id="${s.id}">
      <td><input type="checkbox" data-act="inc" ${s.include ? 'checked' : ''} aria-label="Include sheet"></td>
      <td>${esc(s.fileName)}</td>
      <td>${esc(s.sheetName)}${s.hidden ? ' <span class="badge">hidden</span>' : ''}</td>
      <td><select data-act="role" aria-label="Use as"><option value="data" ${s.role !== 'ref' ? 'selected' : ''}>Data</option><option value="ref" ${s.role === 'ref' ? 'selected' : ''}>Connected list</option></select></td>
      <td><input type="number" min="${s.r0 + 1}" max="${s.r0 + Math.max(s.grid.length, 1)}" value="${hdrRow}" data-act="hdr" aria-label="Header row"></td>
      <td class="num">${s.headers.length}</td>
      <td class="num">${s.dataRows.length.toLocaleString('en-US')}</td>
      <td>${dropped ? `<span class="badge warn" title="Total/subtotal rows: ${s.stats.totals}; repeated header rows: ${s.stats.repeats}">${dropped} skipped</span> ` : ''}${s.dupOf ? `<span class="badge" title="Same columns and identical rows as ${esc(s.dupOf)}">same data as ${esc(s.dupOf)}</span>` : ''}${!s.grid.length ? '<span class="badge">empty</span>' : ''}</td>
      <td><button class="btn sm" data-act="peek">${state.peek === s.id ? 'Hide' : 'Preview'}</button>
          <button class="icon-btn" data-act="del" title="Remove sheet" aria-label="Remove sheet">✕</button></td>
    </tr>${state.peek === s.id ? `<tr data-id="${s.id}"><td colspan="9">${peekHtml(s)}</td></tr>` : ''}`;
  }).join('');
  el.innerHTML = `<div class="tbl-wrap"><table class="ui"><thead><tr><th></th><th>File</th><th>Sheet</th><th>Use as</th><th>Header row</th><th class="num">Columns</th><th class="num">Data rows</th><th>Notes</th><th></th></tr></thead><tbody>${rows}</tbody></table></div>`;
}

function peekHtml(s) {
  const g = gridOf(s);
  const lim = Math.min(g.length, Math.max(s.headerIdx + 8, 12));
  const width = Math.min(g.reduce((w, r) => Math.max(w, r.length), 0), 15);
  let h = '';
  for (let r = 0; r < lim; r++) {
    h += `<tr class="${r === s.headerIdx ? 'hdr' : ''}"><td class="rn" data-act="sethdr" data-r="${r}" title="Use as header row">${s.r0 + r + 1}</td>`;
    for (let c = 0; c < width; c++) h += `<td>${esc(displayValue(g[r][c]))}</td>`;
    h += '</tr>';
  }
  return `<div class="peek-hint">Highlighted row is used as the column headers. Click a row number to change it.</div><div class="peek"><table>${h}</table></div>`;
}

function renderMasterOpts() {
  const o = state.opts;
  const cols = state.columns.filter((c) => c.include);
  const colOpts = (sel, extra) => extra + cols.map((c) => `<option value="${esc(c.key)}" ${sel === c.key ? 'selected' : ''}>${esc(c.name)}</option>`).join('');
  const cb = (k, label, title) => `<label title="${esc(title || '')}"><input type="checkbox" data-opt="${k}" ${o[k] ? 'checked' : ''}> ${label}</label>`;
  $('#masterOpts').innerHTML = `
    ${cb('addSN', 'Add S/N column')}
    ${cb('addSource', 'Add Source File &amp; Sheet columns')}
    ${cb('dropTotals', 'Skip total / subtotal rows', 'Rows whose first value starts with “Total”, “Sub Total” or “Grand Total”')}
    ${cb('fillMerged', 'Fill merged cells', 'Copy the value of a merged cell into every row/column it covers')}
    ${cb('dedupe', 'Remove duplicate rows')}
    ${cb('masterTotals', 'Totals row under master')}
    <label>Sort by <select data-opt="sortBy">${colOpts(o.sortBy, '<option value="">Original order</option>')}</select>
      <select data-opt="sortDir"><option value="asc" ${o.sortDir === 'asc' ? 'selected' : ''}>A → Z / smallest first</option><option value="desc" ${o.sortDir === 'desc' ? 'selected' : ''}>Z → A / largest first</option></select></label>
    <label title="How to read text dates like 03/04/2025">Text dates <select data-opt="dateOrder"><option value="DMY" ${o.dateOrder === 'DMY' ? 'selected' : ''}>DD/MM/YYYY</option><option value="MDY" ${o.dateOrder === 'MDY' ? 'selected' : ''}>MM/DD/YYYY</option></select></label>
    <label>Alignment <select data-opt="align"><option value="center" ${o.align === 'center' ? 'selected' : ''}>Center everything</option><option value="smart" ${o.align === 'smart' ? 'selected' : ''}>Text left, numbers right</option></select></label>
    <label title="Also create one extra sheet per value of this column">Extra sheet per <select data-opt="splitBy">${colOpts(o.splitBy, '<option value="">(none)</option><option value="__file"' + (o.splitBy === '__file' ? ' selected' : '') + '>Source File</option>')}</select></label>`;
}

function renderColumns() {
  const el = $('#columns');
  const cols = state.columns;
  if (!cols.length) { el.innerHTML = '<div class="empty" style="margin-top:14px">No columns found yet. Include at least one sheet with a header row.</div>'; return; }
  const nSheets = included().length;
  const resolved = new Map((state.master ? state.master.defs : []).map((d) => [d.key, d.type]));
  const issues = new Map((state.master ? state.master.issues : []).map((i) => [i.col, i]));
  const rows = cols.map((c, i) => {
    const seen = c.aliases.reduce((a, x) => a + x.count, 0);
    const chips = c.aliases.map((a) => `<span class="chip ${a.auto ? 'auto' : ''}" title="${a.auto ? 'Matched automatically by similar name' : ''}">${esc(a.name)}${a.count > 1 ? ` ×${a.count}` : ''}${(a.norm !== c.key || a.auto) ? `<button data-act="split" data-norm="${esc(a.norm)}" title="Separate into its own column" aria-label="Separate">×</button>` : ''}</span>`).join('');
    const det = resolved.get(c.key);
    const typeSel = TYPE_OPTIONS.map(([v, l]) => `<option value="${v}" ${c.type === v ? 'selected' : ''}>${v === 'auto' && det && c.type === 'auto' ? `Auto (${TYPE_NAME[det]})` : l}</option>`).join('');
    const mergeSel = `<option value="">—</option>` + cols.filter((o) => o !== c).map((o) => `<option value="${esc(o.key)}">${esc(o.name)}</option>`).join('');
    const issue = issues.get(c.name);
    let note = '';
    if (!c.aliases.length) note = '<span class="warn-text">Not found in any sheet</span>';
    else if (seen < nSheets) note = `<span class="muted">${seen} of ${nSheets} sheets</span>`;
    if (issue && c.include) note += `<div class="warn-text" title="Examples: ${esc(issue.examples.join(', '))}">${issue.count} value(s) not ${TYPE_NAME[det] || ''}</div>`;
    return `<tr class="${c.include ? '' : 'off'}" data-key="${esc(c.key)}">
      <td><input type="checkbox" data-act="inc" ${c.include ? 'checked' : ''} aria-label="Include column"></td>
      <td style="white-space:nowrap"><button class="icon-btn" data-act="up" ${i === 0 ? 'disabled' : ''} aria-label="Move up">▲</button><button class="icon-btn" data-act="down" ${i === cols.length - 1 ? 'disabled' : ''} aria-label="Move down">▼</button></td>
      <td><input class="name" type="text" value="${esc(c.name)}" data-act="name" aria-label="Column name">${resolved.get(c.key) === 'text' && c.include ? `<label class="xc" title="Pull the code out of values like “INSPECTION, OVERHAUL (321110)” into a separate “${esc(c.name)} Code” column"><input type="checkbox" data-act="xcode" ${c.xcode ? 'checked' : ''}> code → own column</label>` : ''}</td>
      <td><div class="chips">${chips || '<span class="muted">—</span>'}</div></td>
      <td><select data-act="type">${typeSel}</select></td>
      <td><select data-act="merge" aria-label="Combine into">${mergeSel}</select></td>
      <td>${note}</td>
    </tr>`;
  }).join('');
  const nInc = cols.filter((c) => c.include).length;
  el.innerHTML = `<div class="row gap" style="margin-top:14px"><span class="muted">${nInc} of ${cols.length} columns included</span>
      <button class="btn sm" data-act="all">Include all</button><button class="btn sm" data-act="none">Exclude all</button></div>
    <div class="tbl-wrap"><table class="ui"><thead><tr><th></th><th>Order</th><th>Master column name</th><th>Source headers</th><th>Type</th><th>Combine into</th><th>Notes</th></tr></thead><tbody>${rows}</tbody></table></div>`;
}

function renderConnections() {
  const el = $('#connections');
  const M = state.master;
  $('#stepConn').hidden = !state.connections.length || !M;
  if (!state.connections.length || !M) { el.innerHTML = ''; return; }
  el.innerHTML = state.connections.map((conn) => {
    const src = state.sources.find((x) => x.id === conn.srcId);
    const refOpts = (sel, extra) => extra + (src ? src.headers : []).map((h) => `<option value="${esc(h.norm)}" ${h.norm === sel ? 'selected' : ''}>${esc(h.name)}</option>`).join('');
    const masterOpts = (sel) => M.defs.filter((d) => !d.source && !d.conn).map((d) => `<option value="${esc(d.key)}" ${d.key === sel ? 'selected' : ''}>${esc(d.name)}</option>`).join('');
    const st = conn.stats || {};
    let stats = '';
    if (st.error) stats = `<span class="bad-text">${esc(st.error)}</span>`;
    else if (st.nomatch) stats = '<span class="warn-text">Not connected yet — see the scan result below and pick the columns manually.</span>';
    else if (st.matched !== undefined) {
      const pc = M.records.length ? (st.matched / M.records.length * 100).toFixed(1) : '0';
      const via = st.via2 !== null ? ` (rule 1: ${st.via1.toLocaleString('en-US')}, rule 2: ${st.via2.toLocaleString('en-US')})` : '';
      stats = `<span class="${st.matched ? 'ok-text' : 'bad-text'}">${st.matched.toLocaleString('en-US')} of ${M.records.length.toLocaleString('en-US')} master rows matched (${pc}%)${via}</span> · ${st.refRows.toLocaleString('en-US')} list rows, ${st.keys.toLocaleString('en-US')} unique keys`;
    }
    let diag = '';
    if (conn.diag && conn.diag.length) {
      diag = 'Pairings found by the automatic scan: ' + conn.diag.map((p) => `${esc(p.masterName)} ↔ ${esc(p.refName)} (${(p.pct * 100).toFixed(0)}%)`).join(' · ');
    } else if (conn.diag && !conn.diag.length) {
      const mi2 = M.index[conn.masterKey];
      const mSample = mi2 !== undefined ? M.records.map((r) => r[mi2]).filter((v) => !isBlank(v)).slice(0, 2).map((v) => `"${esc(displayValue(v))}"`).join(', ') : '';
      const hi2 = src ? src.headers.findIndex((h) => h.norm === conn.refKeyNorm) : -1;
      const rSample = hi2 >= 0 ? src.dataRows.map((r) => r[hi2]).filter((v) => !isBlank(v)).slice(0, 3).map((v) => `"${esc(displayValue(v))}"`).join(', ') : '';
      diag = `<span class="bad-text">The scan compared every master column with every list column and found no values in common.</span> ` +
        `List keys look like ${rSample}${mSample ? `; e.g. master “${esc((M.defs[mi2] || {}).name || '')}” holds ${mSample}` : ''}. ` +
        `Pick the columns manually above — or the connecting value may genuinely not exist in the raw files (e.g. the list keys on a group/vendor number that no raw column contains).`;
    }
    return `<div class="bd" data-conn="${esc(conn.key)}">
      <div class="bd-head"><h3>${esc(src ? src.fileName + ' › ' + src.sheetName : conn.key)}</h3></div>
      <div class="bd-grid">
        <label class="field"><span>Name (used for the added columns)</span><input type="text" data-cn="name" value="${esc(conn.name)}"></label>
        <label class="field"><span>Rule 1 — list key column</span><select data-cn="refKeyNorm">${refOpts(conn.refKeyNorm, '')}</select></label>
        <label class="field"><span>Rule 1 — master column</span><select data-cn="masterKey"><option value="">(auto)</option>${masterOpts(conn.masterKey)}</select></label>
        <label class="field"><span>Rule 2 (fallback) — list key column</span><select data-cn="refKeyNorm2">${refOpts(conn.refKeyNorm2, '<option value="">(none)</option>')}</select></label>
        <label class="field"><span>Rule 2 (fallback) — master column</span><select data-cn="masterKey2"><option value="">(none)</option>${masterOpts(conn.masterKey2)}</select></label>
        <label class="field"><span>Bring across (distinct values per match)</span><select data-cn="bring"><option value="">(count only)</option>${refOpts(conn.bring, '')}</select></label>
      </div>
      <div class="bd-grid" style="margin-top:8px">
        <label class="field" style="flex-direction:row;align-items:center;gap:6px"><input type="checkbox" data-cn="addCount" ${conn.addCount ? 'checked' : ''}> <span>Add “${esc(conn.name)} Matches”${conn.bring ? ' and brought-value' : ''} column(s) to master</span></label>
        <label class="field" style="flex-direction:row;align-items:center;gap:6px"><input type="checkbox" data-cn="exportSheet" ${conn.exportSheet ? 'checked' : ''}> <span>Include full list as its own sheet in the export</span></label>
      </div>
      <div class="conn-stats">${stats}</div>
      ${diag ? `<div class="conn-stats muted">${diag}</div>` : ''}
    </div>`;
  }).join('');
}

function renderBreakdowns() {
  const el = $('#breakdowns');
  const M = state.master;
  if (!M) { el.innerHTML = ''; return; }
  const defs = M.defs;
  const opt = (d, sel) => `<option value="${esc(d.key)}" ${sel === d.key ? 'selected' : ''}>${esc(d.name)}</option>`;
  if (!state.breakdowns.length) { el.innerHTML = '<div class="empty">No breakdowns. Add one to group the data, for example by department or by month.</div>'; return; }
  el.innerHTML = state.breakdowns.map((bd, idx) => {
    const g = defs[M.index[bd.groupBy]];
    const isDateG = g && g.type === 'date';
    const meas = (bd.measures || []).map((m, mi) => {
      const allowed = ['sum', 'avg'].includes(m.agg) ? defs.filter((d) => NUMERIC.has(d.type)) : ['min', 'max'].includes(m.agg) ? defs.filter((d) => NUMERIC.has(d.type) || d.type === 'date') : defs;
      return `<div class="measure" data-mi="${mi}">
        <select data-bd="agg">${AGG_OPTIONS.map(([v, l]) => `<option value="${v}" ${m.agg === v ? 'selected' : ''}>${l}</option>`).join('')}</select>
        ${m.agg === 'count' ? '' : `<select data-bd="mcol">${allowed.length ? '' : '<option value="">(no suitable column)</option>'}${allowed.map((d) => opt(d, m.col)).join('')}</select>`}
        <button class="icon-btn" data-bd="mdel" aria-label="Remove measure" ${bd.measures.length < 2 ? 'disabled' : ''}>✕</button></div>`;
    }).join('');
    return `<div class="bd" data-bd-id="${bd.id}">
      <div class="bd-head"><h3>Breakdown ${idx + 1}${g ? ': by ' + esc(g.name) : ' <span class="warn-text">(column no longer in master)</span>'}</h3>
        <button class="icon-btn" data-bd="del" aria-label="Remove breakdown">✕ Remove</button></div>
      <div class="bd-grid">
        <label class="field"><span>Group rows by</span><select data-bd="groupBy">${g ? '' : '<option value="">(choose)</option>'}${defs.map((d) => opt(d, bd.groupBy)).join('')}</select></label>
        ${isDateG ? `<label class="field"><span>Dates as</span><select data-bd="groupMode">${[['value', 'Each date'], ['month', 'Month'], ['quarter', 'Quarter'], ['year', 'Year']].map(([v, l]) => `<option value="${v}" ${(bd.groupMode || 'value') === v ? 'selected' : ''}>${l}</option>`).join('')}</select></label>` : ''}
        <label class="field"><span>Spread across columns by</span><select data-bd="splitBy"><option value="">(none)</option>${defs.filter((d) => d.key !== bd.groupBy).map((d) => opt(d, bd.splitBy)).join('')}</select></label>
        <label class="field"><span>Order</span><select data-bd="sort"><option value="desc" ${bd.sort === 'desc' ? 'selected' : ''}>Largest first</option><option value="alpha" ${bd.sort === 'alpha' ? 'selected' : ''}>A → Z / by date</option></select></label>
        <label class="field" title="Show only the top N groups; the rest are rolled into an “Other” row. Not applied when spreading across columns."><span>Show top</span><input type="number" min="0" data-bd="topN" value="${+bd.topN || ''}" placeholder="all" style="width:70px"></label>
        <label class="field" style="flex-direction:row;align-items:center;gap:6px"><input type="checkbox" data-bd="hideBlank" ${bd.hideBlank ? 'checked' : ''}> <span>Ignore blank</span></label>
        <label class="field" style="flex-direction:row;align-items:center;gap:6px"><input type="checkbox" data-bd="pct" ${bd.pct !== false ? 'checked' : ''}> <span>Show % of total</span></label>
      </div>
      <div class="field" style="margin-top:10px"><span>Values${bd.splitBy ? ' (only the first is used when spreading across columns)' : ''}</span>
        <div class="measures">${meas}</div>
        <div><button class="btn sm" data-bd="madd" style="margin-top:4px">+ Add value</button></div></div>
    </div>`;
  }).join('');
}

function blockHtml(b) {
  const left = state.opts.align === 'smart';
  const cell = (v, t) => {
    let val = v, type = t;
    if (v && typeof v === 'object' && !isDate(v)) { val = v.v; type = v.t; }
    if (type === 'status') {
      const ok = String(val).startsWith('\u2713');
      return `<td class="${ok ? 'ok-text' : 'bad-text'}">${esc(val)}</td>`;
    }
    type = autoType(val, type);
    const cls = left ? (NUMERIC.has(type) ? 'n' : 't') : '';
    return `<td class="${cls}">${esc(displayValue(val, type))}</td>`;
  };
  let h = `<div class="xl-title">${esc(b.title)}</div><table class="xl${left ? ' left' : ''}"><thead><tr>${b.headers.map((x) => `<th>${esc(x)}</th>`).join('')}</tr></thead><tbody>`;
  for (const r of b.rows) h += `<tr>${r.map((v, i) => cell(v, b.types[i])).join('')}</tr>`;
  if (b.total) h += `<tr class="total">${b.total.map((v, i) => cell(v, i === 0 ? 'text' : b.types[i])).join('')}</tr>`;
  return h + '</tbody></table>';
}

function renderPreview() {
  const el = $('#preview');
  const M = state.master;
  document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('active', t.dataset.tab === state.tab));
  if (!M) { el.innerHTML = ''; return; }
  const n = M.records.length;
  const V = state.verification;
  $('#exportInfo').innerHTML = `${n.toLocaleString('en-US')} records · ${M.defs.filter((d) => !d.source && !d.conn).length} columns · ${included().length} sheet(s)` +
    (V ? (V.bad ? ` · <span class="bad-text">\u2717 ${V.bad} verification check(s) failed</span>` : ' · <span class="ok-text">\u2713 verified against raw data</span>') : '');
  if (state.tab === 'summary') {
    el.innerHTML = (state.summary || []).map(blockHtml).join('');
    return;
  }
  if (state.tab === 'verify') {
    const V = state.verification;
    if (!V) { el.innerHTML = ''; return; }
    const head = V.bad
      ? `<div class="bad-text" style="margin-bottom:12px">\u2717 ${V.bad} check(s) did not tally — review below before using the report.</div>`
      : `<div class="ok-text" style="margin-bottom:12px">\u2713 All checks tally: every raw row is accounted for and all totals match the raw data.</div>`;
    el.innerHTML = head + V.blocks.map(blockHtml).join('');
    return;
  }
  const lim = 200;
  const view = masterView(M.records.slice(0, lim));
  const totals = state.opts.masterTotals ? masterTotals(view.defs, M.records) : null;
  el.innerHTML = blockHtml({
    title: `${state.title || 'Master Report'}${n > lim ? ` — showing first ${lim} of ${n.toLocaleString('en-US')} rows` : ''}`,
    headers: view.defs.map((d) => d.name), types: view.defs.map((d) => d.type), rows: view.rows,
    total: totals && n <= lim ? totals : null
  });
}

function masterTotals(defs, records) {
  if (!defs.some((d) => (d.type === 'int' || d.type === 'dec') && !d.sn)) return null;
  const full = masterView(records);
  return full.defs.map((d, i) => {
    if (i === 0) return 'TOTAL';
    if ((d.type !== 'int' && d.type !== 'dec') || d.sn) return '';
    let s = 0; for (const r of full.rows) if (typeof r[i] === 'number') s += r[i];
    return s;
  });
}

// ---------------------------------------------------------------------------
// Excel export (ExcelJS)
// ---------------------------------------------------------------------------
const styleCache = new Map();
function xStyle(role, type, band) {
  const key = `${role}|${type}|${band ? 1 : 0}|${state.opts.align}`;
  if (styleCache.has(key)) return styleCache.get(key);
  const thin = { style: 'thin', color: { argb: XC.border } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };
  const smart = state.opts.align === 'smart' && role !== 'header';
  const horizontal = !smart ? 'center' : NUMERIC.has(type) ? 'right' : type === 'date' ? 'center' : 'left';
  const s = {
    font: { name: 'Calibri', size: 11, bold: role !== 'body', color: role === 'header' ? { argb: XC.headFont } : { argb: 'FF1B2433' } },
    alignment: { horizontal, vertical: 'middle', wrapText: true, indent: smart && horizontal !== 'center' ? 1 : 0 },
    border: role === 'total' ? { ...border, top: { style: 'double', color: { argb: XC.title } } } : border,
    numFmt: NUMFMT[type] || 'General'
  };
  const fill = role === 'header' ? XC.head : role === 'total' ? XC.total : band ? XC.band : null;
  if (fill) s.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
  styleCache.set(key, s);
  return s;
}

function textLen(v, type) {
  const s = displayValue(v, type);
  return s.split('\n').reduce((m, l) => Math.max(m, l.length), 0);
}
function headerLen(h) {
  const words = String(h).split(/\s+/);
  return Math.max(Math.min(String(h).length, 18), ...words.map((w) => w.length));
}

function writeTitle(ws, row, c1, c2, text, size, opts = {}) {
  if (c2 > c1) ws.mergeCells(row, c1, row, c2);
  const cell = ws.getCell(row, c1);
  cell.value = text;
  cell.style = {
    font: { name: 'Calibri', size, bold: !opts.italic, italic: !!opts.italic, color: { argb: opts.color || XC.title } },
    alignment: { horizontal: opts.left ? 'left' : 'center', vertical: 'middle' }
  };
}

/** Writes a styled table; returns { nextRow, widths } */
function writeTable(ws, r0, c0, headers, types, rows, total, widths, opts = {}) {
  const hr = ws.getRow(r0);
  hr.height = 30;
  headers.forEach((h, i) => {
    const cell = hr.getCell(c0 + i);
    cell.value = h;
    cell.style = xStyle('header', 'text', false);
    widths[c0 + i] = Math.max(widths[c0 + i] || 0, headerLen(h));
  });
  let r = r0 + 1;
  const sample = rows.length > 3000 ? 3000 : rows.length;
  rows.forEach((row, ri) => {
    const xr = ws.getRow(r);
    for (let i = 0; i < headers.length; i++) {
      let v = row[i], t = types[i];
      if (v && typeof v === 'object' && !isDate(v)) { t = v.t; v = v.v; }
      const cell = xr.getCell(c0 + i);
      if (t === 'status') {
        const ok = String(v).startsWith('\u2713');
        cell.value = v;
        cell.style = { ...xStyle('body', 'text', ri % 2 === 1), font: { name: 'Calibri', size: 11, bold: true, color: { argb: ok ? 'FF18794E' : 'FFB42318' } } };
        if (ri < sample) widths[c0 + i] = Math.max(widths[c0 + i] || 0, String(v).length);
        continue;
      }
      t = autoType(v, t);
      cell.value = isBlank(v) ? null : v;
      cell.style = xStyle('body', t, ri % 2 === 1);
      if (ri < sample) widths[c0 + i] = Math.max(widths[c0 + i] || 0, textLen(v, t));
    }
    r++;
  });
  if (total) {
    const xr = ws.getRow(r);
    xr.height = 20;
    for (let i = 0; i < headers.length; i++) {
      let v = total[i], t = i === 0 ? 'text' : types[i];
      const cell = xr.getCell(c0 + i);
      if (t === 'status') {
        const ok = String(v).startsWith('\u2713');
        cell.value = v;
        cell.style = { ...xStyle('total', 'text', false), font: { name: 'Calibri', size: 11, bold: true, color: { argb: ok ? 'FF18794E' : 'FFB42318' } } };
        widths[c0 + i] = Math.max(widths[c0 + i] || 0, String(v).length);
        continue;
      }
      if (opts.formulaTotals && typeof v === 'number' && rows.length) {
        const col = colLetter(c0 + i - 1);
        cell.value = { formula: `SUBTOTAL(109,${col}${r0 + 1}:${col}${r - 1})`, result: v };
      } else cell.value = isBlank(v) ? null : v;
      cell.style = xStyle('total', autoType(v, t), false);
      widths[c0 + i] = Math.max(widths[c0 + i] || 0, textLen(v, autoType(v, t)));
    }
    r++;
  }
  return r;
}

function applyWidths(ws, widths, min = 9, max = 50) {
  for (const k of Object.keys(widths)) ws.getColumn(+k).width = Math.min(max, Math.max(min, Math.ceil(widths[k] * 1.12) + 3));
}

function pageSetup(ws, landscape, title, printRows, nCols) {
  // Fitting a very wide sheet to one page would make it unreadable in print.
  const fit = !nCols || nCols <= 14;
  ws.pageSetup = {
    paperSize: 9, orientation: landscape ? 'landscape' : 'portrait',
    fitToPage: fit, ...(fit ? { fitToWidth: 1, fitToHeight: 0 } : { scale: 70 }),
    horizontalCentered: true, margins: { left: 0.4, right: 0.4, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 }
  };
  if (printRows) ws.pageSetup.printTitlesRow = printRows;
  const safe = String(title).replace(/&/g, '&&');
  ws.headerFooter = { oddFooter: `&L&8${safe}&C&8Page &P of &N&R&8&D` };
}

function writeDataSheet(wb, name, title, records, generated) {
  const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 4, showGridLines: false }] });
  const view = masterView(records);
  const n = view.defs.length;
  writeTitle(ws, 1, 1, n, title, 16);
  ws.getRow(1).height = 30;
  writeTitle(ws, 2, 1, n, `Generated ${generated}  ·  ${records.length.toLocaleString('en-US')} records`, 10, { italic: true, color: XC.sub });
  ws.getRow(3).height = 6;
  const widths = {};
  const totals = state.opts.masterTotals ? masterTotals(view.defs, records) : null;
  const end = writeTable(ws, 4, 1, view.defs.map((d) => d.name), view.defs.map((d) => d.type), view.rows, totals, widths, { formulaTotals: true });
  applyWidths(ws, widths);
  if (state.opts.addSN) ws.getColumn(1).width = Math.max(7, ws.getColumn(1).width - 2);
  if (records.length) ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4 + records.length, column: n } };
  pageSetup(ws, n > 6, title, '4:4', n);
  return end;
}

function writeBlockSheet(wb, sheetName, heading, sub, blocks, tabColor) {
  const ws = wb.addWorksheet(sheetName, { views: [{ showGridLines: false }], properties: { tabColor: { argb: tabColor } } });
  ws.getColumn(1).width = 3;
  const maxW = Math.max(4, ...blocks.map((b) => b.headers.length));
  writeTitle(ws, 1, 2, 1 + maxW, heading, 16);
  ws.getRow(1).height = 30;
  writeTitle(ws, 2, 2, 1 + maxW, sub, 10, { italic: true, color: XC.sub });
  const widths = {};
  let r = 4;
  for (const b of blocks) {
    const w = b.headers.length;
    writeTitle(ws, r, 2, 1 + w, b.title, 12, { left: state.opts.align === 'smart' });
    for (let c = 2; c <= 1 + w; c++) ws.getCell(r, c).border = { bottom: { style: 'medium', color: { argb: XC.section } } };
    ws.getRow(r).height = 22;
    r = writeTable(ws, r + 1, 2, b.headers, b.types, b.rows, b.total || null, widths);
    r += 2;
  }
  applyWidths(ws, widths, 12, 45);
  pageSetup(ws, maxW > 6, heading);
  return ws;
}

function writeSummarySheet(wb, title, generated) {
  const ws = wb.addWorksheet('Summary', { views: [{ showGridLines: false }], properties: { tabColor: { argb: XC.head } } });
  ws.getColumn(1).width = 3;
  const blocks = state.summary;
  const maxW = Math.max(4, ...blocks.map((b) => b.headers.length));
  writeTitle(ws, 1, 2, 1 + maxW, `${title} — Summary`, 16);
  ws.getRow(1).height = 30;
  writeTitle(ws, 2, 2, 1 + maxW, `Generated ${generated}`, 10, { italic: true, color: XC.sub });
  const widths = {};
  let r = 4;
  for (const b of blocks) {
    const w = b.headers.length;
    writeTitle(ws, r, 2, 1 + w, b.title, 12, { left: state.opts.align === 'smart' });
    for (let c = 2; c <= 1 + w; c++) ws.getCell(r, c).border = { bottom: { style: 'medium', color: { argb: XC.section } } };
    ws.getRow(r).height = 22;
    r = writeTable(ws, r + 1, 2, b.headers, b.types, b.rows, b.total || null, widths);
    r += 2;
  }
  applyWidths(ws, widths, 12, 45);
  pageSetup(ws, maxW > 6, title);
}

function safeSheetName(name, used) {
  let s = String(name ?? '').replace(/[[\]:*?/\\]/g, ' ').replace(/\s+/g, ' ').trim().replace(/^'+|'+$/g, '') || 'Blank';
  if (/^history$/i.test(s)) s = 'History data';
  s = s.slice(0, 31);
  let out = s, i = 2;
  while (used.has(out.toLowerCase())) { const suf = ` (${i++})`; out = s.slice(0, 31 - suf.length) + suf; }
  used.add(out.toLowerCase());
  return out;
}

async function exportExcel() {
  const M = state.master;
  if (!M) return;
  if (typeof ExcelJS === 'undefined') { toast('Excel library failed to load.', true); return; }
  busy(true, 'Building Excel report…');
  await nextFrame();
  try {
    styleCache.clear();
    computeSummary();
    const title = state.title || 'Master Report';
    const generated = fmtDateTime(new Date());
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Report Builder';
    wb.created = new Date();
    wb.calcProperties.fullCalcOnLoad = true;
    computeVerification();
    const used = new Set(['summary', 'master data', 'verification']);
    writeSummarySheet(wb, title, generated);
    writeDataSheet(wb, 'Master Data', title, M.records, generated);
    if (state.verification) {
      const V = state.verification;
      const vws = writeBlockSheet(wb, 'Verification', `${title} — Verification`,
        `Generated ${generated}  ·  ${V.bad ? V.bad + ' check(s) DID NOT tally — review before use' : 'All checks tally with the raw data'}`,
        V.blocks, V.bad ? 'FFB42318' : 'FF18794E');
      vws.getCell(2, 2).style = { font: { name: 'Calibri', size: 11, bold: true, color: { argb: V.bad ? 'FFB42318' : 'FF18794E' } }, alignment: { horizontal: 'center', vertical: 'middle' } };
    }
    // Connected lists as their own formatted sheets (optional)
    for (const conn of state.connections) {
      if (!conn.exportSheet) continue;
      const src = state.sources.find((x) => x.id === conn.srcId);
      if (!src) continue;
      const ws = wb.addWorksheet(safeSheetName(conn.name + ' (list)', used), { views: [{ state: 'frozen', ySplit: 4, showGridLines: false }] });
      const n = src.headers.length;
      writeTitle(ws, 1, 1, n, `${conn.name} — connected list (kept separate from the master)`, 14);
      ws.getRow(1).height = 26;
      writeTitle(ws, 2, 1, n, `Source: ${src.fileName} › ${src.sheetName}  ·  ${src.dataRows.length.toLocaleString('en-US')} rows`, 10, { italic: true, color: XC.sub });
      const widths = {};
      writeTable(ws, 4, 1, src.headers.map((h) => h.name), src.headers.map(() => 'auto'), src.dataRows, null, widths);
      applyWidths(ws, widths);
      if (src.dataRows.length) ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: 4 + src.dataRows.length, column: n } };
      pageSetup(ws, n > 6, conn.name, '4:4', n);
    }

    const si = M.index[state.opts.splitBy];
    if (si !== undefined) {
      const groups = sortGroups(groupRows(M.records, si, M.defs[si].type === 'date' ? 'month' : 'value'), 'alpha', () => 0);
      if (groups.length > 100) toast(`"${M.defs[si].name}" has ${groups.length} values; only the first 100 get their own sheet.`, true);
      for (const g of groups.slice(0, 100)) {
        const label = isDate(g.label) ? fmtDate(g.label) : String(g.label);
        writeDataSheet(wb, safeSheetName(label, used), `${title} — ${M.defs[si].name}: ${label}`, g.rows, generated);
      }
    }
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const fname = (state.fileName || title).replace(/[\\/:*?"<>|]+/g, ' ').trim() || 'Report';
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = /\.xlsx$/i.test(fname) ? fname : fname + '.xlsx';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
    saveLocal();
    toast('Report downloaded: ' + a.download);
  } catch (e) {
    console.error(e);
    toast('Export failed: ' + (e.message || e), true);
  } finally {
    busy(false);
  }
}

// ---------------------------------------------------------------------------
// Settings save / load
// ---------------------------------------------------------------------------
function settingsObj() {
  for (const c of state.columns) state.colPrefs[c.key] = { ...(state.colPrefs[c.key] || {}), name: c.name, include: c.include, type: c.type };
  return {
    app: 'report-builder', version: SETTINGS_VERSION,
    title: state.title, fileName: state.fileName, opts: state.opts,
    template: state.template, colPrefs: state.colPrefs, colOrder: state.colOrder, aliasOverrides: state.aliasOverrides, sheetPrefs: state.sheetPrefs, connPrefs: state.connPrefs,
    bdAuto: state.bdAuto, breakdowns: state.breakdowns.map(({ id, ...b }) => b)
  };
}
function applySettings(s) {
  if (!s || s.app !== 'report-builder') throw new Error('not a Report Builder settings file');
  state.title = s.title || 'Master Report';
  state.fileName = s.fileName || '';
  state.opts = { ...defaultOpts(), ...(s.opts || {}) };
  state.template = s.template || null;
  state.colPrefs = s.colPrefs || {};
  state.colOrder = s.colOrder || [];
  state.aliasOverrides = s.aliasOverrides || {};
  state.sheetPrefs = s.sheetPrefs || {};
  state.connPrefs = s.connPrefs || {};
  state.connections = [];
  for (const src of state.sources) {
    const pref = state.sheetPrefs[sheetKey(src)];
    if (!pref) continue;
    if (pref.include !== undefined) src.include = pref.include;
    if (pref.role) src.role = pref.role;
    if (pref.headerRow) src.headerIdx = Math.max(0, Math.min(src.grid.length - 1, pref.headerRow - src.r0 - 1));
  }
  state.bdAuto = s.bdAuto !== false;
  state.breakdowns = (s.breakdowns || []).map((b) => ({ ...b, id: nextId++ }));
  refresh('sources');
}
function saveLocal() {
  try { localStorage.setItem(LS_KEY, JSON.stringify(settingsObj())); } catch (e) { /* storage unavailable */ }
}
function hasLocal() {
  try { return !!localStorage.getItem(LS_KEY); } catch (e) { return false; }
}

// ---------------------------------------------------------------------------
// Events
// ---------------------------------------------------------------------------
function setPref(key, patch) {
  state.colPrefs[key] = { ...(state.colPrefs[key] || {}), ...patch };
}
function moveCol(key, dir) {
  const keys = state.columns.map((c) => c.key);
  const i = keys.indexOf(key), j = i + dir;
  if (i < 0 || j < 0 || j >= keys.length) return;
  [keys[i], keys[j]] = [keys[j], keys[i]];
  state.colOrder = keys.concat(state.colOrder.filter((k) => !keys.includes(k)));
}
const userBd = () => { state.bdAuto = false; };
const findBd = (el) => state.breakdowns.find((b) => b.id === +el.closest('[data-bd-id]').dataset.bdId);

function wireDrop(id, handler) {
  const el = $(id);
  el.addEventListener('dragover', (e) => { e.preventDefault(); el.classList.add('over'); });
  el.addEventListener('dragleave', () => el.classList.remove('over'));
  el.addEventListener('drop', (e) => { e.preventDefault(); el.classList.remove('over'); handler(e.dataTransfer.files); });
}

function init() {
  $('#fileInput').addEventListener('change', (e) => { addFiles(e.target.files); e.target.value = ''; });
  $('#refInput').addEventListener('change', (e) => { addFiles(e.target.files, 'ref'); e.target.value = ''; });
  $('#templateInput').addEventListener('change', (e) => { if (e.target.files[0]) loadTemplate(e.target.files[0]); e.target.value = ''; });
  wireDrop('#dropData', addFiles);
  wireDrop('#dropRef', (f) => addFiles(f, 'ref'));
  wireDrop('#dropTemplate', (f) => f[0] && loadTemplate(f[0]));

  $('#sources').addEventListener('click', (e) => {
    const t = e.target.closest('[data-act]');
    if (!t) return;
    const tr = t.closest('tr[data-id]');
    const src = state.sources.find((s) => s.id === +(tr && tr.dataset.id));
    if (!src) return;
    const act = t.dataset.act;
    if (act === 'peek') { state.peek = state.peek === src.id ? null : src.id; renderSources(); }
    else if (act === 'del') { state.sources = state.sources.filter((s) => s !== src); refresh('sources'); }
    else if (act === 'sethdr') { src.headerIdx = +t.dataset.r; state.sheetPrefs[sheetKey(src)] = { include: src.include, headerRow: src.r0 + src.headerIdx + 1, role: src.role }; refresh('sources'); }
  });
  $('#sources').addEventListener('change', (e) => {
    const t = e.target;
    const tr = t.closest('tr[data-id]');
    const src = state.sources.find((s) => s.id === +(tr && tr.dataset.id));
    if (!src) return;
    if (t.dataset.act === 'inc') src.include = t.checked;
    if (t.dataset.act === 'role') src.role = t.value;
    if (t.dataset.act === 'hdr') src.headerIdx = Math.max(0, Math.min(src.grid.length - 1, (+t.value || 1) - src.r0 - 1));
    state.sheetPrefs[sheetKey(src)] = { include: src.include, headerRow: src.r0 + src.headerIdx + 1, role: src.role };
    refresh('sources');
  });

  $('#masterOpts').addEventListener('change', (e) => {
    const k = e.target.dataset.opt;
    if (!k) return;
    state.opts[k] = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
    refresh(['fillMerged', 'dropTotals'].includes(k) ? 'sources' : k === 'addSN' ? 'columns' : 'master');
  });

  $('#columns').addEventListener('click', (e) => {
    const t = e.target.closest('[data-act]');
    if (!t) return;
    const act = t.dataset.act;
    if (act === 'all' || act === 'none') { state.columns.forEach((c) => setPref(c.key, { include: act === 'all' })); refresh('columns'); return; }
    const key = t.closest('tr') && t.closest('tr').dataset.key;
    if (act === 'up' || act === 'down') { moveCol(key, act === 'up' ? -1 : 1); refresh('columns'); }
    else if (act === 'split') { state.aliasOverrides[t.dataset.norm] = t.dataset.norm; setPref(t.dataset.norm, { include: true }); refresh('columns'); }
  });
  $('#columns').addEventListener('change', (e) => {
    const t = e.target;
    const key = t.closest('tr') && t.closest('tr').dataset.key;
    if (!key) return;
    const act = t.dataset.act;
    if (act === 'inc') setPref(key, { include: t.checked });
    else if (act === 'name') setPref(key, { name: t.value.trim() || undefined });
    else if (act === 'xcode') setPref(key, { xcode: t.checked });
    else if (act === 'type') setPref(key, { type: t.value });
    else if (act === 'merge' && t.value) {
      const from = state.columns.find((c) => c.key === key);
      for (const a of from.aliases) state.aliasOverrides[a.norm] = t.value;
      for (const [n, k] of Object.entries(state.aliasOverrides)) if (k === key) state.aliasOverrides[n] = t.value;
      setPref(t.value, { include: true });
    }
    refresh('columns');
  });

  $('#breakdowns').addEventListener('change', (e) => {
    const t = e.target, f = t.dataset.bd;
    if (!f) return;
    const bd = findBd(t);
    if (!bd) return;
    userBd();
    const mi = t.closest('[data-mi]') ? +t.closest('[data-mi]').dataset.mi : -1;
    if (f === 'agg') {
      const m = bd.measures[mi];
      m.agg = t.value;
      const M = state.master;
      const okType = (d) => (['sum', 'avg'].includes(m.agg) ? NUMERIC.has(d.type) : ['min', 'max'].includes(m.agg) ? NUMERIC.has(d.type) || d.type === 'date' : true);
      if (m.agg !== 'count' && (!m.col || !okType(M.defs[M.index[m.col]] || {}))) { const d = M.defs.find(okType); m.col = d ? d.key : ''; }
    } else if (f === 'mcol') bd.measures[mi].col = t.value;
    else if (f === 'pct' || f === 'hideBlank') bd[f] = t.checked;
    else if (f === 'topN') bd.topN = +t.value || 0;
    else bd[f] = t.value;
    if (f === 'groupBy') { bd.groupMode = 'value'; if (bd.splitBy === bd.groupBy) bd.splitBy = ''; }
    computeSummary();
    render('summary');
  });
  $('#breakdowns').addEventListener('click', (e) => {
    const t = e.target.closest('[data-bd]');
    if (!t || t.tagName === 'SELECT' || t.tagName === 'INPUT') return;
    const bd = findBd(t);
    if (!bd) return;
    userBd();
    const f = t.dataset.bd;
    if (f === 'del') state.breakdowns = state.breakdowns.filter((b) => b !== bd);
    else if (f === 'madd') bd.measures.push({ agg: 'count' });
    else if (f === 'mdel') bd.measures.splice(+t.closest('[data-mi]').dataset.mi, 1);
    computeSummary(); computeVerification();
    render('summary');
  });
  $('#connections').addEventListener('change', (e) => {
    const t = e.target, f = t.dataset.cn;
    if (!f) return;
    const conn = state.connections.find((c) => c.key === t.closest('[data-conn]').dataset.conn);
    if (!conn) return;
    if (f === 'name') conn.name = t.value.trim() || defaultConnName(state.sources.find((x) => x.id === conn.srcId) || { fileName: 'List' });
    else if (f === 'addCount' || f === 'exportSheet') conn[f] = t.checked;
    else conn[f] = t.value;
    state.connPrefs[conn.key] = { name: conn.name, refKeyNorm: conn.refKeyNorm, masterKey: conn.masterKey, refKeyNorm2: conn.refKeyNorm2, masterKey2: conn.masterKey2, bring: conn.bring, addCount: conn.addCount, exportSheet: conn.exportSheet };
    refresh('master');
  });

  $('#btnAddBd').addEventListener('click', () => {
    const M = state.master;
    const d = M.defs.find((x) => x.type === 'text' && !x.source) || M.defs[0];
    userBd();
    state.breakdowns.push({ id: nextId++, groupBy: d.key, groupMode: 'value', splitBy: '', measures: [{ agg: 'count' }], pct: true, sort: 'desc' });
    computeSummary(); computeVerification();
    render('summary');
  });
  $('#btnResetBd').addEventListener('click', () => { state.bdAuto = true; state.breakdowns = suggestBreakdowns(); computeSummary(); computeVerification(); render('summary'); });

  $('#title').addEventListener('change', (e) => { state.title = e.target.value.trim(); refresh('master'); });
  $('#fileName').addEventListener('change', (e) => { state.fileName = e.target.value.trim(); });
  document.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => { state.tab = t.dataset.tab; renderPreview(); }));
  $('#btnExport').addEventListener('click', exportExcel);

  $('#btnSaveSettings').addEventListener('click', () => {
    const blob = new Blob([JSON.stringify(settingsObj(), null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = ((state.fileName || state.title || 'report').replace(/[\\/:*?"<>|]+/g, ' ').trim() || 'report') + ' - settings.json';
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
    saveLocal();
    toast('Settings saved. Load them next time to reuse the same layout.');
  });
  $('#btnLoadSettings').addEventListener('click', () => $('#settingsInput').click());
  $('#settingsInput').addEventListener('change', async (e) => {
    const f = e.target.files[0];
    e.target.value = '';
    if (!f) return;
    try { applySettings(JSON.parse(await f.text())); toast('Settings loaded'); } catch (err) { toast('Could not load settings: ' + err.message, true); }
  });
  $('#btnRestore').hidden = !hasLocal();
  $('#btnRestore').addEventListener('click', () => {
    try { applySettings(JSON.parse(localStorage.getItem(LS_KEY))); toast('Last used settings restored'); } catch (err) { toast('Could not restore settings', true); }
  });
  $('#btnReset').addEventListener('click', () => {
    if (state.sources.length && !confirm('Remove all loaded files and settings?')) return;
    Object.assign(state, { sources: [], template: null, colPrefs: {}, colOrder: [], aliasOverrides: {}, sheetPrefs: {}, connPrefs: {}, connections: [], verification: null, columns: [], opts: defaultOpts(), breakdowns: [], bdAuto: true, title: 'Master Report', fileName: '', master: null, summary: null, peek: null });
    refresh('sources');
  });
  if (typeof XLSX === 'undefined' || typeof ExcelJS === 'undefined') toast('A required library did not load. Please use the built ReportBuilder.html file.', true);
  render('all');
}

// Exposed for automated tests.
window.ReportBuilder = { state, refresh };
init();
})();
