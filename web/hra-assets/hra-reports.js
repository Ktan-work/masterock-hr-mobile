/* Central HRA web - Phase 3 reports. Port of hrapp/forms/cpcec.py and hrapp/forms/dashboard.py (same rules, same output). */
(function () {
'use strict';
const ASSETS = (window.HRA_ASSETS || 'hra-assets/');
const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MON3 = MON.map(m => m.slice(0, 3));
const up = v => String(v ?? '').replace(/\s+/g, ' ').trim().toUpperCase();
const p2 = n => String(n).padStart(2, '0');
const escRe = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const pyRound = x => { const f = Math.floor(x), d = x - f; return d > 0.5 ? f + 1 : d < 0.5 ? f : (f % 2 === 0 ? f : f + 1); };
const round2 = (x, n) => { const k = 10 ** n; return pyRound(x * k) / k; };                                        // round(x, n), half-even
const pyFloat = x => Number.isInteger(x) ? x.toFixed(1) : String(x);                                             // repr(float)
const NUMRE = /^\s*[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?\s*$/;

/* ---------- dates: strptime-like, tried in the same order as the Python code ---------- */
function ymd(y, m, d) { const t = new Date(y, m - 1, d); return (t.getFullYear() === y && t.getMonth() === m - 1 && t.getDate() === d) ? t : null; }
const monIdx = (s, full) => { const i = (full ? MON : MON3).findIndex(x => x.toLowerCase() === s.toLowerCase()); return i < 0 ? 0 : i + 1; };
const FMT = {
  '%d/%m/%Y': [/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, m => ymd(+m[3], +m[2], +m[1])],
  '%d/%m/%y': [/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/, m => ymd(+m[3] < 69 ? 2000 + +m[3] : 1900 + +m[3], +m[2], +m[1])],
  '%Y-%m-%d': [/^(\d{4})-(\d{1,2})-(\d{1,2})$/, m => ymd(+m[1], +m[2], +m[3])],
  '%d-%b-%Y': [/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/, m => monIdx(m[2]) ? ymd(+m[3], monIdx(m[2]), +m[1]) : null],
  '%d %b %Y': [/^(\d{1,2}) ([A-Za-z]{3}) (\d{4})$/, m => monIdx(m[2]) ? ymd(+m[3], monIdx(m[2]), +m[1]) : null],
  '%d %B %Y': [/^(\d{1,2}) ([A-Za-z]+) (\d{4})$/, m => monIdx(m[2], true) ? ymd(+m[3], monIdx(m[2], true), +m[1]) : null],
  '%B %d, %Y': [/^([A-Za-z]+) (\d{1,2}), (\d{4})$/, m => monIdx(m[1], true) ? ymd(+m[3], monIdx(m[1], true), +m[2]) : null],
  '%m/%d/%Y': [/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/, m => ymd(+m[3], +m[1], +m[2])],
};
function parseWith(v, order) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return (v > 20000 && v < 80000) ? new Date(Math.round((v - 25569) * 864e5) + new Date().getTimezoneOffset() * 6e4) : null;
  const s = String(v).trim();
  for (const f of order) { const [re, fn] = FMT[f], m = s.match(re); if (m) { const d = fn(m); if (d) return d; } }
  return null;
}
const toDateCommon = v => parseWith(v, ['%d/%m/%Y', '%Y-%m-%d', '%d/%m/%y', '%d-%b-%Y', '%d %b %Y', '%d %B %Y', '%m/%d/%Y']);
const toDateDash = v => parseWith(v, ['%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d', '%d-%b-%Y', '%d %b %Y', '%B %d, %Y', '%m/%d/%Y']);

/* =========================================================== CPCEC (port of cpcec.py) */
const FIRST = 4, NCOL = 22;
const C = { NO: 1, NAME: 3, PP: 4, PPEXP: 5, NAT: 6, DOM: 7, POS: 8, ARR: 12, RV: 13, VTYPE: 14, VEXP: 15, STATUS: 16, ADDR: 17, EPF: 18, KEY: 22 };
const mdy = v => { const d = toDateCommon(v); return d ? `${p2(d.getMonth() + 1)}/${p2(d.getDate())}/${d.getFullYear()}` : String(v ?? '').trim(); };

/* values = tab values (rows of strings), employees = row objects keyed by UPPER header -> {updates:[[row,col,value]], newRows, log} */
function cpcecPlan(values, employees) {
  const cell = (r, c) => (r - 1 < values.length && c - 1 < (values[r - 1] || []).length) ? (values[r - 1][c - 1] ?? '') : '';
  const have = {}, pprow = {}, dom = {}; let maxNo = 0;
  for (let r = FIRST; r <= values.length; r++) {
    const k = up(cell(r, C.KEY)); if (k) have[k] = r;
    const pp = up(cell(r, C.PP)).replace(/[^A-Z0-9]/g, ''); if (pp) pprow[pp] = r;
    const n = String(cell(r, C.NO)); if (NUMRE.test(n)) maxNo = Math.max(maxNo, Math.trunc(parseFloat(n)));
    const nat = up(cell(r, C.NAT)), dc = String(cell(r, C.DOM)).trim();
    if (nat && dc && !(nat in dom)) dom[nat] = dc;
  }
  const updates = [], newRows = [], log = [];
  for (const e of employees) {
    const g = k => e[k] || '';
    const emp = up(g('SL EMP ID'));
    if (!emp || up(g('COMPANY')) !== 'MROCK' || up(g('EXPAT / LOCAL')) !== 'EXPAT') continue;
    if (up(g('VISA STATUS')).includes('BUSINESS')) continue;
    const inactive = up(g('ACTIVE/ INACTIVE')) === 'INACTIVE';
    if (emp in have) {
      const r = have[emp];
      if (inactive && r > 0 && up(cell(r, C.STATUS)) !== 'CANCELLED') { updates.push([r, C.STATUS, 'Cancelled']); log.push(`Cancelled  ${emp}  ${g('FULL NAME')}`); }
      continue;
    }
    const pp = up(g('PASSPORT NUMBER')).replace(/[^A-Z0-9]/g, '');
    if (pp && pp in pprow && !up(cell(pprow[pp], C.KEY))) {
      updates.push([pprow[pp], C.KEY, emp]); have[emp] = pprow[pp]; log.push(`Linked     ${emp}  ${g('FULL NAME')}`); continue;
    }
    maxNo += 1;
    const row = new Array(NCOL).fill('');
    const vals = { [C.NO]: maxNo, [C.NAME]: g('FULL NAME'), [C.PP]: g('PASSPORT NUMBER'), [C.PPEXP]: mdy(g('DATE OF EXPIRY (MM/DD/YYYY)')),
      [C.NAT]: g('NATIONALITY'), [C.DOM]: dom[up(g('NATIONALITY'))] || '', [C.POS]: g('POSITION (PC /MR DESIGNATION)'),
      [C.ARR]: mdy(g('MROCK /REAL JOIN DATE (MM/DD/YY)')), [C.RV]: g('RESIDENCE VISA NUMBER'), [C.VTYPE]: g('VISA STATUS'),
      [C.VEXP]: mdy(g('MOST RECENT EXPIRY DATE')), [C.STATUS]: inactive ? 'Cancelled' : 'Active',
      [C.ADDR]: g('RESIDENCE ADDRESS IN SRI LANKA'), [C.EPF]: g('EPF MEMBER NO'), [C.KEY]: emp };
    for (const [c, v] of Object.entries(vals)) row[+c - 1] = v;
    newRows.push(row); have[emp] = -1;
    log.push(`Added      ${emp}  ${g('FULL NAME')}` + (inactive ? '  (resigned / inactive)' : ''));
  }
  return { updates, newRows, log };
}
function cpcecStatus(row) {
  const s = [13, 14, 15].map(i => up(row[i])).join(' ');
  return s.includes('CANCEL') ? 'Cancelled' : (s.includes('PENDING') ? 'Pending arrival' : 'Active');
}

/* Excel in the original format (template hra-assets/cpcec_template.xlsx), needs ExcelJS */
async function cpcecExcel(values) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await (await fetch(ASSETS + 'cpcec_template.xlsx')).arrayBuffer());
  const ws = wb.worksheets[0], ncol = 21;
  const style = []; for (let c = 1; c <= ncol; c++) style.push(JSON.parse(JSON.stringify(ws.getCell(4, c).style || {})));
  if (ws.rowCount >= 4) ws.spliceRows(4, ws.rowCount - 3);
  try { ws.conditionalFormattings = []; } catch (e) {}
  ws.getCell('C1').value = null;
  const fills = { 'Cancelled': 'FFCFE2F3', 'Pending arrival': 'FFFCE4D6' };
  const rows = values.slice(FIRST - 1).filter(r => r.some(x => String(x ?? '').trim()));
  rows.forEach((r, i) => {
    const rr = FIRST + i, st = cpcecStatus(r);
    for (let c = 1; c <= ncol; c++) {
      const v = c - 1 < r.length ? String(r[c - 1] ?? '').trim() : '';
      const cell = ws.getCell(rr, c); cell.style = JSON.parse(JSON.stringify(style[c - 1]));
      const m = v.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      const d = m ? ymd(+m[3], +m[1], +m[2]) : null;
      if (d) { cell.value = new Date(Date.UTC(+m[3], +m[1] - 1, +m[2])); cell.numFmt = 'mm/dd/yyyy'; }
      else if (c === C.NO && /^\d+$/.test(v)) cell.value = parseInt(v, 10);
      else cell.value = v || null;
      if (fills[st]) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fills[st] } };
    }
  });
  ws.views = [{ state: 'frozen', xSplit: 3, ySplit: 3, topLeftCell: 'D4' }];
  return { bytes: new Uint8Array(await wb.xlsx.writeBuffer()), rows: rows.length };
}

/* =========================================================== HR Dashboard (port of dashboard.py) */
const NAT = {
  'VIETNAMESE': ['Vietnamese', '🇻🇳', 14, 108], 'FILIPINO': ['Filipino', '🇵🇭', 12.8, 121.8], 'FILIPINOS': ['Filipino', '🇵🇭', 12.8, 121.8],
  'THAI': ['Thai', '🇹🇭', 13, 100.5], 'CAMBODIAN': ['Cambodian', '🇰🇭', 12.6, 104.9], 'INDONESIAN': ['Indonesian', '🇮🇩', -2.5, 118],
  'SRI LANKAN': ['Sri Lankan', '🇱🇰', 7.9, 80.7], 'CHINESE': ['Chinese', '🇨🇳', 35.9, 104.2], 'MALAYSIAN': ['Malaysian', '🇲🇾', 4.2, 108],
  'INDIAN': ['Indian', '🇮🇳', 20.6, 78.9], 'NEPAL': ['Nepalese', '🇳🇵', 28.4, 84.1], 'NEPALESE': ['Nepalese', '🇳🇵', 28.4, 84.1], 'NEPALI': ['Nepalese', '🇳🇵', 28.4, 84.1],
  'BANGLADESHI': ['Bangladeshi', '🇧🇩', 23.7, 90.4], 'ITALIANA': ['Italiana', '🇮🇹', 42.5, 12.6], 'ITALIAN': ['Italiana', '🇮🇹', 42.5, 12.6],
  'SAINT LUCIA': ['Saint Lucia', '🇱🇨', 13.9, -60.9], 'SAINT LUCIAN': ['Saint Lucia', '🇱🇨', 13.9, -60.9], 'MYANMAR': ['Myanmar', '🇲🇲', 21.9, 95.9],
  'BURMESE': ['Myanmar', '🇲🇲', 21.9, 95.9], 'PAKISTANI': ['Pakistani', '🇵🇰', 30.4, 69.3], 'PAKISTAN': ['Pakistani', '🇵🇰', 30.4, 69.3],
  'KOREAN': ['Korean', '🇰🇷', 36.5, 127.8], 'JAPANESE': ['Japanese', '🇯🇵', 36.2, 138.3], 'TAIWANESE': ['Taiwanese', '🇹🇼', 23.7, 121],
  'SINGAPOREAN': ['Singaporean', '🇸🇬', 1.35, 103.8], 'LAOTIAN': ['Laotian', '🇱🇦', 19.9, 102.5], 'BRITISH': ['British', '🇬🇧', 54, -2],
  'AUSTRALIAN': ['Australian', '🇦🇺', -25, 134], 'AMERICAN': ['American', '🇺🇸', 39, -98], 'BRAZILIAN': ['Brazilian', '🇧🇷', -14, -51],
};
const pyTitle = s => s.toLowerCase().replace(/(^|[^A-Za-z])([a-z])/g, (m, a, b) => a + b.toUpperCase());
function natInfo(s) {
  const k = String(s ?? '').trim().toUpperCase().replace(/\s+/g, ' ');
  if (!k || k === 'NONE') return ['Unknown', '🏳️', null, null];
  return NAT[k] || [pyTitle(k), '🏳️', null, null];
}
const KEEP_UPPER = new Set(['VIP', 'IT', 'THB', 'VND', 'IDR', 'CNY', 'SL', 'GPI', 'P2P', 'QA', 'HR', 'MR', 'PC', 'BPO']);
function smartTitle(s) {
  return String(s).trim().split(' ').map(w => {
    const core = w.replace(/[^A-Za-z0-9]/g, '');
    if (!core || /\d/.test(core) || KEEP_UPPER.has(core.toUpperCase())) return w.toUpperCase();
    return !w.startsWith('(') ? w.slice(0, 1).toUpperCase() + w.slice(1).toLowerCase() : '(' + w.slice(1, 2).toUpperCase() + w.slice(2).toLowerCase();
  }).join(' ');
}
const ordinal = n => `${n}${(n % 100 >= 11 && n % 100 <= 13) ? 'th' : ({ 1: 'st', 2: 'nd', 3: 'rd' }[n % 10] || 'th')}`;
const isUpperKey = s => /[A-Z]/.test(s) && !/[a-z]/.test(s);
function pyJson(v) {                                  // json.dumps(..., ensure_ascii=False) look-alike
  if (Array.isArray(v)) return '[' + v.map(pyJson).join(', ') + ']';
  if (typeof v === 'string') return JSON.stringify(v);
  if (v && v.__float !== undefined) return pyFloat(v.__float);
  return String(v);
}
const DEFAULT_INPUTS = { MNL_TOTAL: 143, MNL_EXPAT: 8, MNL_LOCAL: 135, MNL_MALE: 64, MNL_FEMALE: 79, MNL_REGULAR: 131, MNL_PROB: 12,
  SL_PAID: 13, SL_ABS: 0, SL_TARDY: 15, MNL_PAID: 3, MNL_ABS: 3, MNL_TARDY: 7, MNL_VISA_EXP: 0, MNL_PP_EXP: 0, SSS: 143, PHILHEALTH: 143, PAGIBIG: 143 };
const num = (v, d = 0) => NUMRE.test(String(v ?? '')) ? pyRound(parseFloat(v)) : d;

/* 'Dashboard Inputs' tab values -> {KEY: value} (rows whose column A is an UPPER-CASE key and column C is filled) */
function dashInputs(vals) {
  const width = Math.max(0, ...vals.map(r => r.length)), out = {};
  for (const r0 of vals) { const r = r0.concat(new Array(width - r0.length).fill(''));
    if (r.length > 2 && String(r[0]).trim() && String(r[2]).trim() !== '' && isUpperKey(String(r[0]).trim())) out[String(r[0]).trim()] = r[2]; }
  return out;
}

function dashCompute(emps, inputs, today) {
  const g = (r, k) => r[k] ?? null;
  const comp = r => ({ MROCK: 'MR', PCL: 'PC' })[up(g(r, 'COMPANY'))];
  const active = emps.filter(r => up(g(r, 'ACTIVE/ INACTIVE')) === 'ACTIVE' && comp(r));
  const D = {};
  const sameMonth = d => d !== null && d.getFullYear() === today.getFullYear() && d.getMonth() === today.getMonth();
  for (const c of ['MR', 'PC']) {
    const R = active.filter(r => comp(r) === c), cnt = f => R.filter(f).length;
    D[c] = { total: R.length, expat: cnt(r => up(g(r, 'EXPAT / LOCAL')) === 'EXPAT'), local: cnt(r => up(g(r, 'EXPAT / LOCAL')) === 'LOCAL'),
      male: cnt(r => up(g(r, 'GENDER (F / M)')).startsWith('M')), female: cnt(r => up(g(r, 'GENDER (F / M)')).startsWith('F')),
      regular: cnt(r => up(g(r, 'EMPLOYMENT STATUS')) === 'REGULAR'), prob: cnt(r => up(g(r, 'EMPLOYMENT STATUS')).startsWith('PROB')) };
    D[c].visa_exp = cnt(r => sameMonth(toDateDash(g(r, 'MOST RECENT EXPIRY DATE'))));
    D[c].pp_exp = cnt(r => sameMonth(toDateDash(g(r, 'DATE OF EXPIRY (MM/DD/YYYY)'))));
    const bc = c === 'MR' ? "B'CARD STATUS (MR)" : "B'CARD STATUS (PC)";
    D[c].bcard = cnt(r => up(g(r, bc)) === 'RECEIVED');
    D[c].tin = cnt(r => up(g(r, 'TIN NUMBER')) === 'RECEIVED' || /^[\d\-\s]{6,}$/.test(up(g(r, 'TIN NUMBER'))));
  }
  D.SL = D.MR.total + D.PC.total;
  const I = {}; for (const [k, v] of Object.entries(DEFAULT_INPUTS)) I[k] = num(k in inputs ? inputs[k] : v, v);
  D.I = I; D.ALL = D.SL + I.MNL_TOTAL;

  const nat = new Map();
  for (const r of active) {
    const [name, flag, lat, lon] = natInfo(g(r, 'NATIONALITY'));
    if (!nat.has(name)) nat.set(name, { name, flag, lat, lon, MR: 0, PC: 0, male: 0, female: 0, expat: 0, local: 0 });
    const e = nat.get(name); e[comp(r)] += 1;
    const sx = up(g(r, 'GENDER (F / M)'));
    e[sx.startsWith('M') ? 'male' : 'female'] += ['M', 'F'].includes(sx.slice(0, 1)) ? 1 : 0;
    e[up(g(r, 'EXPAT / LOCAL')) === 'EXPAT' ? 'expat' : 'local'] += 1;
  }
  D.nat = [...nat.values()].sort((a, b) => (b.MR + b.PC) - (a.MR + a.PC));
  const dept = new Map();
  for (const r of active) {
    const k = up(g(r, 'DEPARTMENT (INTERNAL)')) || 'UNASSIGNED';
    if (!dept.has(k)) dept.set(k, { name: smartTitle(k), MR: 0, PC: 0 });
    dept.get(k)[comp(r)] += 1;
  }
  D.dept = [...dept.values()].sort((a, b) => (b.MR + b.PC) - (a.MR + a.PC));

  const first = new Date(today.getFullYear(), today.getMonth(), 1);
  const lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const weeks = [[1, 7], [8, 14], [15, 21], [22, lastDay]];
  const wk = d => { for (let i = 0; i < weeks.length; i++) if (weeks[i][0] <= d.getDate() && d.getDate() <= weeks[i][1]) return i; return weeks.length - 1; };
  const on = { MR: [], PC: [] }, off = { MR: [], PC: [] };
  for (const r of emps) {
    const c = comp(r); if (!c) continue;
    const jd = toDateDash(g(r, c === 'MR' ? 'MROCK /REAL JOIN DATE (MM/DD/YY)' : 'PC JOIN DATE (MM/DD/YYYY)'));
    if (jd && first <= jd && jd <= today) on[c].push([jd, r]);
    const rd = toDateDash(g(r, c === 'MR' ? 'MROCK RESIGNATION DATE (MM/DD/YYYY)' : 'FOR PC RESIGNATION DATE (MM/DD/YYYY)'));
    if (rd && first <= rd && rd <= today) off[c].push([rd, r]);
  }
  D.weeks = weeks.map(([a, b], i) => `Week ${i + 1} (${a}–${b})`);
  D.on_w = weeks.map((_, i) => on.MR.concat(on.PC).filter(([d]) => wk(d) === i).length);
  D.off_w = weeks.map((_, i) => off.MR.concat(off.PC).filter(([d]) => wk(d) === i).length);
  D.on = { MR: on.MR.length, PC: on.PC.length }; D.off = { MR: off.MR.length, PC: off.PC.length };
  const sl = Math.max(D.SL, 1);
  D.attr = round2((D.off.MR + D.off.PC) / sl * 100, 2);
  D.attr_w = D.off_w.map(n => round2(n / sl * 100, 2));
  for (const c of ['MR', 'PC']) {
    const agg = new Map();
    for (const [d, r] of off[c]) {
      const [name, flag] = natInfo(g(r, 'NATIONALITY'));
      if (!agg.has(name)) agg.set(name, { name, flag, count: 0, people: [] });
      const e = agg.get(name); e.count += 1;
      e.people.push(`${g(r, 'FULL NAME')} (${g(r, 'SL EMP ID')}) — ${p2(d.getDate())} ${MON3[d.getMonth()]} ${d.getFullYear()}`);
    }
    D['offnat_' + c] = [...agg.values()].sort((a, b) => b.count - a.count);
  }
  return D;
}

/* ---- template fill helpers (warnings when the template changed) */
let WARN = [];
function sub1(html, re, repl, label) {             // first match only
  let n = 0; const out = html.replace(re, (...a) => { n++; return typeof repl === 'function' ? repl(...a) : repl; });
  if (!n) WARN.push('not found in template: ' + label); return out;
}
const subAll = (html, re, repl) => html.replace(new RegExp(re.source, re.flags.includes('g') ? re.flags : re.flags + 'g'), typeof repl === 'function' ? repl : () => repl);
function inBlock(html, start, endPattern, fn, label) {
  const i = html.indexOf(start); if (i < 0) { WARN.push('block not found: ' + label); return html; }
  const re = new RegExp(endPattern, 'gs'); re.lastIndex = i + start.length; const m = re.exec(html);
  if (!m) { WARN.push('block end not found: ' + label); return html; }
  const j = m.index + m[0].length; return html.slice(0, i) + fn(html.slice(i, j)) + html.slice(j);
}
function attRows(block, items, basis) {
  const rates = items.map(([, n]) => basis ? round2(n / basis * 100, 2) : 0);
  const mx = Math.max(...rates) || 1;
  items.forEach(([label, n], k) => {
    const rate = rates[k], w = rate ? pyRound(rate / mx * 100) : 0;
    const re = new RegExp('(<tr><td>' + escRe(label) + '</td><td>)[^<]*(</td><td><div class="rbar"><div class="rb"><div class="rf" style="width:)[\\d.]+(%;[^"]*"></div></div>)[^<]*(</div>)');
    let hit = 0; block = block.replace(re, (m, a, b, c, d) => { hit++; return `${a}${n}${b}${w}${c}${rate.toFixed(2)}%${d}`; });
    if (!hit) WARN.push('attendance row not found: ' + label);
  });
  return subAll(block, /Basis: \d+ employees/, `Basis: ${basis} employees`);
}
function expRows(block, visa, pp) {
  const mx = Math.max(visa, pp, 1);
  for (const [label, n] of [['Visa Expiring', visa], ['Passport Expiring', pp]]) {
    const w = n ? pyRound(n / mx * 100) : 0;
    block = block.replace(new RegExp('(<tr><td>' + label + '</td><td>)[^<]*(</td><td><div class="rbar"><div class="rb"><div class="rf" style="width:)[\\d.]+%'), (m, a, b) => `${a}${n}${b}${w}%`);
  }
  return block;
}
function cpItem(block, label, done, total, check = false) {
  const pct = total ? round2(done / total * 100, 1) : 0;
  const pctS = (pct === 0 || pct === 100 ? pct.toFixed(0) : pct.toFixed(1)) + '%';
  const width = pyFloat(pct);                         // Python: float either way (0.0 when total is 0)
  const re = new RegExp('(<span>' + escRe(label) + '</span><span class="cp-pct"[^>]*>)[^<]*(</span></div><div class="cp-bg"><div class="cp-fill" style="width:)[\\d.]+%(.*?<div class="cp-note">)[^<]*(</div>)', 's');
  const note = check ? `${done}/${total} ` + (done >= total && total ? '✅' : '⏳') : `Done: ${done} · Pending: ${total - done} · Total: ${total}`;
  let hit = 0; const out = block.replace(re, (m, a, b, c, d) => { hit++; return `${a}${pctS}${b}${width}%${c}${note}${d}`; });
  if (!hit) WARN.push('compliance item not found: ' + label); return out;
}
function headcount(block, d) {
  block = block.replace(/(class="bc-num"[^>]*>)\d+(<)/, (m, a, b) => `${a}${d.total}${b}`);
  for (const [lab, key] of [['Expats', 'expat'], ['Locals', 'local'], ['♂ Male', 'male'], ['♀ Female', 'female'], ['Regular', 'regular'], ['Probationary', 'prob']])
    block = block.replace(new RegExp('(<div class="mrow"><span>' + escRe(lab) + '</span><span>)\\d+(</span>)'), (m, a, b) => `${a}${d[key]}${b}`);
  return block;
}

function dashFill(html, D, today) {
  const I = D.I, SL = D.SL, MNL = I.MNL_TOTAL;
  const longDate = `${ordinal(today.getDate())} ${MON[today.getMonth()]} ${today.getFullYear()}`;
  const shortDate = `${today.getDate()} ${MON3[today.getMonth()]} ${today.getFullYear()}`;
  const monYear = `${MON[today.getMonth()]} ${today.getFullYear()}`, mon = MON[today.getMonth()];

  html = sub1(html, /(<title>[^<]*— )[A-Za-z]+ \d{4}(<\/title>)/, (m, a, b) => `${a}${MON3[today.getMonth()]} ${today.getFullYear()}${b}`, 'title');
  html = sub1(html, /Data as of [^<]+/, `Data as of ${longDate}`, 'header date');
  html = subAll(html, /([Aa])s of \d{1,2}(?:st|nd|rd|th) [A-Z][a-z]+ \d{4}/, (m, a) => `${a}s of ${longDate}`);
  html = subAll(html, /([Aa])s of \d{1,2} [A-Z][a-z]{2} \d{4}/, (m, a) => `${a}s of ${shortDate}`);

  html = sub1(html, /(<div class="total-num">)\d+(<\/div>)/, (m, a, b) => `${a}${D.ALL}${b}`, 'total headcount');
  html = sub1(html, /(region-badge" style="color:var\(--mr\);[^>]*>)\d+/, (m, a) => `${a}${SL}`, 'SL badge');
  html = sub1(html, /(region-badge" style="color:var\(--mnl\);[^>]*>)\d+/, (m, a) => `${a}${MNL}`, 'Manila badge');
  const probEnd = '<span>Probationary</span><span>\\d+</span></div>';
  html = inBlock(html, 'MR Branch — Masterock</div>', probEnd, b => headcount(b, D.MR), 'MR headcount');
  html = inBlock(html, 'PCBPO Branch — Port City BPO</div>', probEnd, b => headcount(b, D.PC), 'PC headcount');
  const mnl = { total: MNL, expat: I.MNL_EXPAT, local: I.MNL_LOCAL, male: I.MNL_MALE, female: I.MNL_FEMALE, regular: I.MNL_REGULAR, prob: I.MNL_PROB };
  html = inBlock(html, 'Manila Office — Philippines</div>', probEnd, b => headcount(b, mnl), 'Manila headcount');
  html = inBlock(html, 'Sri Lanka Combined (MR + PCBPO)</div>', '</tbody>', b => attRows(b, [['Paid Leaves', I.SL_PAID], ['Absenteeism', I.SL_ABS], ['Tardiness', I.SL_TARDY]], SL), 'SL attendance');
  html = inBlock(html, 'Manila Metrics</div>', '</tbody>', b => attRows(b, [['Paid Leaves', I.MNL_PAID], ['Absenteeism', I.MNL_ABS], ['Tardiness', I.MNL_TARDY]], MNL), 'Manila attendance');

  const visaSeg = seg => {
    seg = subAll(seg, /— [A-Z][a-z]+ \d{4} Only/, `— ${monYear} Only`);
    seg = subAll(seg, /— [A-Z][a-z]+ \d{4}<\/div>/, `— ${monYear}</div>`);
    seg = subAll(seg, /[A-Z][a-z]+ \d{4} expiries — year \d{4}/, `${monYear} expiries — year ${today.getFullYear()}`);
    seg = subAll(seg, /[A-Z][a-z]+ \d{4} expiries</, `${monYear} expiries<`);
    seg = inBlock(seg, 'color:var(--mr);">MR Branch</div>', '</tbody>', b => expRows(b, D.MR.visa_exp, D.MR.pp_exp), 'MR expiry');
    seg = inBlock(seg, 'color:var(--pc);">PCBPO Branch</div>', '</tbody>', b => expRows(b, D.PC.visa_exp, D.PC.pp_exp), 'PC expiry');
    seg = inBlock(seg, 'color:var(--mnl);">Manila Office</div>', '</tbody>', b => expRows(b, I.MNL_VISA_EXP, I.MNL_PP_EXP), 'Manila expiry');
    return seg;
  };
  html = inBlock(html, '<!-- ── ROW 3: Visa', '<!-- ── ROW 4', visaSeg, 'visa section');
  const bcSeg = seg => {
    const mr = b => cpItem(cpItem(b, 'B-Card', D.MR.bcard, D.MR.total), 'TIN No.', D.MR.tin, D.MR.total);
    const pc = b => cpItem(cpItem(b, 'B-Card', D.PC.bcard, D.PC.total), 'TIN No.', D.PC.tin, D.PC.total);
    seg = inBlock(seg, 'margin-bottom:14px;">MR Branch</div>', 'TIN No\\..*?cp-note">[^<]*</div>', mr, 'MR compliance');
    seg = inBlock(seg, 'margin-bottom:14px;">PCBPO Branch</div>', 'TIN No\\..*?cp-note">[^<]*</div>', pc, 'PC compliance');
    for (const [lab, key] of [['SSS', 'SSS'], ['PhilHealth', 'PHILHEALTH'], ['Pag-IBIG', 'PAGIBIG']]) seg = cpItem(seg, lab, I[key], MNL, true);
    return seg;
  };
  html = inBlock(html, '<!-- ── ROW 4', '</div><!-- /grid -->', bcSeg, 'compliance section');

  html = sub1(html, /\d+ employees across Sri Lanka/, `${D.ALL} employees across Sri Lanka`, 'globe caption');
  html = sub1(html, /(font-size:20px;color:var\(--mr\);margin-left:auto;">)\d+/, (m, a) => `${a}${SL}`, 'map SL badge');
  html = sub1(html, /(font-size:20px;color:var\(--mnl\);margin-left:auto;">)\d+/, (m, a) => `${a}${MNL}`, 'map Manila badge');
  html = subAll(html, /(Breakdown — Sri Lanka \()\d+(\))/, (m, a, b) => `${a}${SL}${b}`);
  html = subAll(html, /(Breakdown — Manila \()\d+(\))/, (m, a, b) => `${a}${MNL}${b}`);

  const deptJs = 'const deptData=[\n' + D.dept.map(d => `  {name:${pyJson(d.name)}, MR:${d.MR}, PC:${d.PC}}`).join(',\n') + '\n];';
  html = sub1(html, /const deptData=\[.*?\];/s, () => deptJs, 'deptData');
  html = sub1(html, /const maxDept=\d+;/, `const maxDept=${Math.max(...(D.dept.length ? D.dept.map(d => d.MR + d.PC) : [1]))};`, 'maxDept');
  const natJs = 'const natData=[\n' + D.nat.map(n => `  {name:${pyJson(n.name)},flag:${pyJson(n.flag)},MR:${n.MR},PC:${n.PC}}`).join(',\n') + '\n];';
  html = sub1(html, /const natData=\[.*?\];/s, () => natJs, 'natData');
  html = sub1(html, /const maxNat=\d+;/, `const maxNat=${Math.max(...(D.nat.length ? D.nat.map(n => n.MR + n.PC) : [1]))};`, 'maxNat');
  html = subAll(html, /\(\(total\/\d+\)\*100\)/, `((total/${SL})*100)`);
  html = subAll(html, /\(\(d\.count\/144\)\*100\)/, `((d.count/${MNL})*100)`);
  const gl = [];
  for (const n of D.nat) {
    if (n.lat === null) continue;
    gl.push(`  ${pyJson(n.name.toUpperCase())}:{total:${n.MR + n.PC},MR:${n.MR},PC:${n.PC},male:${n.male},female:${n.female},expat:${n.expat},local:${n.local},lat:${n.lat},lon:${n.lon},flag:${pyJson(n.flag)}}`);
  }
  gl.push(`  MANILA:{total:${MNL},MR:0,PC:0,male:${I.MNL_MALE},female:${I.MNL_FEMALE},expat:${I.MNL_EXPAT},local:${I.MNL_LOCAL},lat:14.6,lon:121,flag:'🇵🇭',isMnl:true}`);
  html = sub1(html, /const globeNatData=\{.*?\n\};/s, () => 'const globeNatData={\n' + gl.join(',\n') + ',\n};', 'globeNatData');

  const onT = D.on.MR + D.on.PC, offT = D.off.MR + D.off.PC;
  html = sub1(html, /MR: \d+ \(MROCK Join Date\) · PCBPO: \d+ \(PC Join Date\)<\/div>\s*<div style="([^"]*)">\d+</,
    (m, a) => `MR: ${D.on.MR} (MROCK Join Date) · PCBPO: ${D.on.PC} (PC Join Date)</div>\n      <div style="${a}">${onT}<`, 'onboarding total');
  html = sub1(html, /MR: \d+ \(MROCK Resignation Date\) · PCBPO: \d+ \(PC Resignation Date\)<\/div>\s*<div style="([^"]*)">\d+</,
    (m, a) => `MR: ${D.off.MR} (MROCK Resignation Date) · PCBPO: ${D.off.PC} (PC Resignation Date)</div>\n      <div style="${a}">${offT}<`, 'offboarding total');
  html = subAll(html, /(Weekly (?:Onboarding|Offboarding|Attrition Rate) — )[A-Z][a-z]+/, (m, a) => a + mon);
  html = subAll(html, /(Weekly (?:Onboarding|Offboarding|Attrition Rate) Trend \()[A-Z][a-z]+(\))/, (m, a, b) => a + mon + b);
  html = sub1(html, /Current Active Headcount \(\d+\)/, `Current Active Headcount (${SL})`, 'attrition basis');
  html = sub1(html, /(color:var\(--a4\);margin:8px 0;">)[\d.]+%/, (m, a) => `${a}${D.attr.toFixed(2)}%`, 'attrition rate');
  html = sub1(html, /\d+ offboarded ÷ \d+ active = [\d.]+%/, `${offT} offboarded ÷ ${SL} active = ${D.attr.toFixed(2)}%`, 'attrition note');
  const labels = 'labels:' + pyJson(D.weeks);
  html = subAll(html, /labels:\['Week 1[^\]]*\]/, () => labels);
  html = sub1(html, /(label:'New Joiners',\s*data:)\[[^\]]*\]/, (m, a) => a + pyJson(D.on_w), 'onboard chart');
  html = sub1(html, /(label:'Departures',\s*data:)\[[^\]]*\]/, (m, a) => a + pyJson(D.off_w), 'offboard chart');
  html = sub1(html, /(label:'Attrition Rate',\s*data:)\[[^\]]*\]/, (m, a) => a + pyJson(D.attr_w.map(x => ({ __float: x }))), 'attrition chart');
  for (const c of ['MR', 'PC']) {
    const arr = D['offnat_' + c];
    const body = '[' + arr.map(a => `{name:${pyJson(a.name)}, flag:${pyJson(a.flag)}, count:${a.count}, people:${pyJson(a.people)}}`).join(',') + ']';
    html = sub1(html, new RegExp('const offboard' + c + '=\\[.*?\\];', 's'), () => `const offboard${c}=${body};`, 'offboard' + c);
    html = sub1(html, new RegExp('const maxOff' + c + '=\\d+;'), `const maxOff${c}=${Math.max(...(arr.length ? arr.map(a => a.count) : [1]))};`, 'maxOff' + c);
    html = sub1(html, new RegExp('(' + c + ' — Offboarding by Nationality</div>\\s*<div class="cc-sub">[^<]*· Total: )\\d+'), (m, a) => `${a}${D.off[c]}`, c + ' offboard total');
  }
  html = sub1(html, /(offboardMR\.sort.*?\(\(d\.count\/)\d+(\)\*100\))/s, (m, a, b) => `${a}${Math.max(D.off.MR, 1)}${b}`, 'MR offboard %');
  html = sub1(html, /(offboardPC\.sort.*?\(\(d\.count\/)\d+(\)\*100\))/s, (m, a, b) => `${a}${Math.max(D.off.PC, 1)}${b}`, 'PC offboard %');
  html = html.split("offPcTbl.innerHTML+=`<tr>\n      <td>${d.flag} ${d.name}</td>").join("offPcTbl.innerHTML+=`<tr title=\"${(d.people||[]).join('\\n')}\">\n      <td>${d.flag} ${d.name}</td>");
  return html;
}

async function dashboard(emps, inputVals, today = new Date()) {
  WARN = [];
  today = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const D = dashCompute(emps, dashInputs(inputVals), today);
  const tpl = await (await fetch(ASSETS + 'hr-dashboard-template.html?v=1')).text();
  const html = dashFill(tpl, D, today);
  const name = `HR_Dashboard_${p2(today.getDate())}_${MON[today.getMonth()]}_${today.getFullYear()}.html`;
  return { html, name, summary: { SL: D.SL, MR: D.MR.total, PC: D.PC.total, MNL: D.I.MNL_TOTAL, on: D.on, off: D.off, attr: D.attr }, warn: WARN.slice() };
}

window.HRAReports = { cpcecPlan, cpcecStatus, cpcecExcel, dashboard, dashInputs, isUpperKey, FIRST, NCOL };
})();
