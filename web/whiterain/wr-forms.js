/* Whiterain HRA web - Phase 2 forms. Word letters: {TOKENS} replaced inside the original .docx (formatting kept).
   Excel forms: original .xlsx filled with ExcelJS. ERD: the Masterock form engine (hra-forms.js). */
(function () {
'use strict';
const BASE = 'forms/';
const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const xmlEsc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const up = s => String(s ?? '').replace(/\s+/g, ' ').trim().toUpperCase();
const ordinal = n => `${n}${(n % 100 >= 11 && n % 100 <= 13) ? 'th' : ({ 1: 'st', 2: 'nd', 3: 'rd' }[n % 10] || 'th')}`;
function parseIso(s) { const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; }   // <input type=date> value
const longDate = s => { const d = s instanceof Date ? s : parseIso(s); return d ? `${MON[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}` : ''; };
const money = v => { const n = parseFloat(String(v ?? '').replace(/,/g, '')); return isNaN(n) ? String(v ?? '').trim() : n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); };
const num = v => { const n = parseFloat(String(v ?? '').replace(/,/g, '')); return isNaN(n) ? null : n; };
const safe = s => String(s).replace(/[\\/:*?"<>|,]/g, '').replace(/\s+/g, ' ').trim();
const fetchBuf = async p => { const r = await fetch(BASE + p); if (!r.ok) throw new Error('Could not load ' + p + ' (' + r.status + ')'); return r.arrayBuffer(); };

/* names from a Whiterain row */
function names(r) {
  const g = k => String(r[k] ?? '').replace(/\s+/g, ' ').trim();
  const full = g('FULL NAME'), last = g('LAST NAME') || full.split(',')[0].trim();
  const first = g('FIRST NAME') || (full.includes(',') ? full.slice(full.indexOf(',') + 1).trim() : full), middle = g('MIDDLE NAME');
  return { full, last, first, middle, given: [first, middle].filter(Boolean).join(' '), natural: [first, middle, last].filter(Boolean).join(' ') };
}

/* ---------- Word: replace {TOKENS} in every XML part that can hold text */
async function docx(template, tokens) {
  const zip = await JSZip.loadAsync(await fetchBuf(template));
  const parts = Object.keys(zip.files).filter(f => /^word\/(document|header\d*|footer\d*)\.xml$/.test(f));
  const left = new Set();
  for (const f of parts) {
    let x = await zip.file(f).async('string');
    x = x.replace(/\{([A-Z_]+)\}/g, (m, k) => k in tokens ? xmlEsc(tokens[k]) : (left.add(k), m));
    zip.file(f, x);
  }
  if (left.size) throw new Error('Template field(s) not filled: ' + [...left].join(', '));
  return zip.generateAsync({ type: 'uint8array', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

/* ---------- Excel: write only the filled cells straight into the template's XML (logos, styles, formulas untouched).
   fn(wb) calls setv(wb.getWorksheet(name) | wb.worksheets[i], 'B7', value) like before. */
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const colNum = c => [...c].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
async function xlsx(template, fn) {
  const zip = await JSZip.loadAsync(await fetchBuf(template)), P = new DOMParser(), S = new XMLSerializer();
  const wbx = P.parseFromString(await zip.file('xl/workbook.xml').async('string'), 'application/xml');
  const rels = P.parseFromString(await zip.file('xl/_rels/workbook.xml.rels').async('string'), 'application/xml');
  const target = id => [...rels.getElementsByTagName('Relationship')].find(r => r.getAttribute('Id') === id).getAttribute('Target').replace(/^\/?(xl\/)?/, 'xl/');
  const sheets = [...wbx.getElementsByTagNameNS(NS, 'sheet')].map(s => ({ name: s.getAttribute('name'), path: target(s.getAttribute('r:id') || s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')), edits: {} }));
  const wb = { worksheets: sheets, getWorksheet: n => sheets.find(s => s.name === n) };
  fn(wb);
  for (const sh of sheets) {
    if (!Object.keys(sh.edits).length) continue;
    const doc = P.parseFromString(await zip.file(sh.path).async('string'), 'application/xml'), data = doc.getElementsByTagNameNS(NS, 'sheetData')[0];
    for (const [addr, v] of Object.entries(sh.edits)) {
      const [, col, rs] = addr.match(/^([A-Z]+)(\d+)$/), rn = +rs;
      let row = [...data.getElementsByTagNameNS(NS, 'row')].find(r => +r.getAttribute('r') === rn);
      if (!row) { row = doc.createElementNS(NS, 'row'); row.setAttribute('r', rn); const after = [...data.getElementsByTagNameNS(NS, 'row')].find(r => +r.getAttribute('r') > rn); data.insertBefore(row, after || null); }
      let c = [...row.getElementsByTagNameNS(NS, 'c')].find(x => x.getAttribute('r') === addr);
      if (!c) { c = doc.createElementNS(NS, 'c'); c.setAttribute('r', addr); const after = [...row.getElementsByTagNameNS(NS, 'c')].find(x => colNum(x.getAttribute('r').replace(/\d+/g, '')) > colNum(col)); row.insertBefore(c, after || null); }
      while (c.firstChild) c.removeChild(c.firstChild); c.removeAttribute('t');
      if (v === null || v === undefined || v === '') continue;                         // empty value, cell style kept
      if (typeof v === 'number') { const e = doc.createElementNS(NS, 'v'); e.textContent = String(v); c.appendChild(e); }
      else { c.setAttribute('t', 'inlineStr'); const is = doc.createElementNS(NS, 'is'), t = doc.createElementNS(NS, 't'); t.setAttribute('xml:space', 'preserve'); t.textContent = String(v); is.appendChild(t); c.appendChild(is); }
    }
    zip.file(sh.path, S.serializeToString(doc));
  }
  // recalculate formulas (totals, weighted scores) when the file is opened
  let w = await zip.file('xl/workbook.xml').async('string');
  w = /<calcPr\b/.test(w) ? w.replace(/<calcPr\b([^>]*?)(\/?)>/, (m, a, sl) => `<calcPr${a.replace(/\s*fullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"${sl}>`) : w.replace(/(<(oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|<\/workbook>)/, '<calcPr fullCalcOnLoad="1"/>$1');
  zip.file('xl/workbook.xml', w);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
const setv = (ws, addr, v) => { ws.edits[addr] = (v === '' || v === undefined) ? null : v; };

/* ---------- the forms (o = options typed on the page) */
async function coe(r, o) {
  const n = names(r), ms = o.title === 'Ms.', d = parseIso(o.issue) || new Date();
  const bytes = await docx('COE.docx', { TITLE: o.title, NAME: up(n.natural), POSITION: o.position || r['POSITION'] || '', JOIN_DATE: longDate(o.join),
    PRONOUN: ms ? 'She' : 'He', POSSESSIVE: ms ? 'her' : 'his', SALARY: money(o.salary), DEMINIMIS: money(o.deminimis),
    PURPOSE: o.purpose || 'whatever legal purpose it may serve', ISSUE_DAY: ordinal(d.getDate()), ISSUE_MONTH_YEAR: `${MON[d.getMonth()]} ${d.getFullYear()}`, SIGNATORY: o.signatory || '' });
  return { name: safe(`COE ${r['EMPLOYEE ID']} ${n.natural}`) + '.docx', bytes };
}
async function confirmation(r, o) {
  const n = names(r);
  const bytes = await docx('WR_CONFIRMATION.docx', { LETTER_DATE: longDate(o.date), NAME_UPPER: up(n.natural), PROBATION_START: longDate(o.probation),
    EFFECTIVE_DATE: longDate(o.effective), TITLE: o.title });
  return { name: safe(`Confirmation ${r['EMPLOYEE ID']} ${n.natural}`) + '.docx', bytes };
}
async function promotion(r, o) {
  const n = names(r);
  const bytes = await docx('WR_PROMOTION.docx', { LETTER_DATE: longDate(o.date), NAME_UPPER: up(n.natural), NEW_POSITION_UPPER: up(o.newPosition),
    DEPARTMENT_UPPER: up(o.department), TITLE: o.title, FULL_NAME: up(n.full), NEW_POSITION: o.newPosition, EFFECTIVE_DATE: longDate(o.effective),
    NEW_SALARY: money(o.newSalary), SIGNATORY: o.signatory || '', SIGNATORY_TITLE: o.signatoryTitle || '' });
  return { name: safe(`Promotion ${r['EMPLOYEE ID']} ${n.natural}`) + '.docx', bytes };
}
const NPA_TICK = { 'Conversion to Regular': 'D15', 'Reclassification': 'H15', 'Transfer': 'K15', 'Merit Increase': 'D17', 'Promotion': 'H17', 'Others': 'K17' };
async function npa(r, o) {
  const n = names(r);
  const bytes = await xlsx('NPA.xlsx', wb => {
    const ws = wb.getWorksheet('FORM');
    setv(ws, 'B7', up(n.full)); setv(ws, 'G7', r['NICKNAME'] || ''); setv(ws, 'K7', r['EMPLOYEE ID']);
    setv(ws, 'B10', longDate(o.join)); setv(ws, 'G10', longDate(o.effective));
    for (const a of o.actions) if (NPA_TICK[a]) setv(ws, NPA_TICK[a], '✔');
    if (o.actions.includes('Others') && o.otherText) setv(ws, 'L18', o.otherText);
    setv(ws, 'J21', 'PROPOSED');
    setv(ws, 'F22', o.curPosition); setv(ws, 'F23', o.curDept); setv(ws, 'F24', num(o.curBasic) ?? o.curBasic); setv(ws, 'F25', num(o.curDem) ?? o.curDem);
    setv(ws, 'J22', o.newPosition); setv(ws, 'J23', o.newDept); setv(ws, 'J24', num(o.newBasic) ?? o.newBasic); setv(ws, 'J25', num(o.newDem) ?? o.newDem);
    const tot = (a, b) => (num(a) !== null || num(b) !== null) ? (num(a) || 0) + (num(b) || 0) : null;
    setv(ws, 'F28', tot(o.curBasic, o.curDem)); setv(ws, 'J28', tot(o.newBasic, o.newDem));
    setv(ws, 'C30', o.remarks);
  });
  return { name: safe(`NPA ${r['EMPLOYEE ID']} ${n.natural}`) + '.xlsx', bytes };
}
async function appraisal(r, o) {
  const n = names(r);
  const bytes = await xlsx('APPRAISAL.xlsx', wb => {
    const ws = wb.worksheets[0];
    setv(ws, 'C5', up(n.full)); setv(ws, 'C6', o.position || r['POSITION'] || ''); setv(ws, 'C7', o.department || r['DEPARTMENT'] || '');
    setv(ws, 'F6', o.reviewType); setv(ws, 'F7', o.period);
  });
  return { name: safe(`Performance Appraisal ${r['EMPLOYEE ID']} ${n.natural}`) + '.xlsx', bytes };
}
async function teamBuilding(r, o) {
  const n = names(r), dept = o.department || r['DEPARTMENT'] || '', members = o.members.split('\n').map(s => s.trim()).filter(Boolean);
  const bytes = await xlsx('TEAM_BUILDING.xlsx', wb => {
    const ca = wb.getWorksheet('CA'), f = wb.getWorksheet('Form');
    setv(ca, 'C7', n.natural); setv(ca, 'C8', dept); setv(ca, 'C9', r['EMPLOYEE ID']);
    setv(ca, 'D12', longDate(o.from)); setv(ca, 'G12', longDate(o.to)); setv(ca, 'D14', longDate(o.liquidate));
    setv(ca, 'B17', `Team Building Fund - ${o.quarter}`); setv(ca, 'B18', dept ? `${dept} Dept.` : ''); setv(ca, 'B19', members.length ? `${members.length} employees` : '');
    setv(ca, 'I17', num(o.amount)); setv(ca, 'B42', n.natural);
    setv(f, 'E9', n.natural); setv(f, 'E10', dept);
    const q = { Q1: 'E12', Q2: 'H12', Q3: 'E14', Q4: 'H14' }[o.quarter], ql = { E12: '1st', H12: '2nd', E14: '3rd', H14: '4th' };
    if (q) setv(f, q, '✔ ' + ql[q]);
    setv(f, 'E15', o.amount ? 'Php ' + money(o.amount) : ''); setv(f, 'E16', longDate(o.received)); setv(f, 'E17', longDate(o.liquidate));
    setv(f, 'D20', dept ? `${dept} Department` : ''); setv(f, 'H21', dept); setv(f, 'D22', o.amount ? 'Php ' + money(o.amount) : '');
    members.slice(0, 20).forEach((m, i) => setv(f, (i % 2 ? 'H' : 'B') + (32 + Math.floor(i / 2)), m));
  });
  return { name: safe(`Team Building ${o.quarter} ${dept || r['EMPLOYEE ID']}`) + '.xlsx', bytes };
}
/* ERD: Masterock's form engine with this employee mapped to its column names */
async function erd(r, o) {
  const n = names(r), d = parseIso(o.join);
  const row = { 'SL EMP ID': r['EMPLOYEE ID'], 'FULL NAME': n.full, 'FIRST NAME': n.first, 'MIDDLE NAME': n.middle, 'LAST NAME / SURNAME': n.last,
    'POSITION (PC /MR DESIGNATION)': o.position || r['POSITION'] || '', 'DEPARTMENT (INTERNAL)': o.department || r['DEPARTMENT'] || '', 'COMPANY': 'WR',
    'MROCK /REAL JOIN DATE (MM/DD/YY)': d ? `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}` : '' };
  const res = await HRAForms.fill(row, ['ERD']);
  if (res.errs.length) throw new Error(res.errs.join('\n'));
  return { name: safe(`ERD ${r['EMPLOYEE ID']} ${n.natural}`) + '.pdf', bytes: res.files[0].bytes };
}
window.WRForms = { names, coe, confirmation, promotion, npa, appraisal, teamBuilding, erd, longDate, NPA_ACTIONS: Object.keys(NPA_TICK) };
})();
