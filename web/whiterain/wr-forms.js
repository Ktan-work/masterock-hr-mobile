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
  const zip = await JSZip.loadAsync(typeof template === 'string' ? await fetchBuf(template) : template), P = new DOMParser(), S = new XMLSerializer();
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
/* ---------- PRF / CA / ATD (template comes from the private Drive folder, passed in as bytes) */
const COMPANIES = { WR: ['WHITERAIN DIGITAL TECHNOLOGIES INC.', '20/F, Zuellig Building Makati Ave. cor. Paseo de Roxas,', 'Makati City, Philippines'],
  MQ: ['MARQUEE HOLDINGS, LTD.', '28/F, Zuellig Building Makati Ave. cor. Paseo de Roxas,', 'Makati City, Philippines'],
  AQ: ['AQUOZ SOLUTIONS INC.', '28/F, Zuellig Building Makati Ave. cor. Paseo de Roxas,', 'Makati City, Philippines 1226'],
  GS: ['GLOW SPARK TECHNOLOGIES INC.', '37/F LKG Tower 6801 Ayala Avenue Bell Air 1209', 'Makati NCR, Fourth District Philippines'],
  GP: ['GAMEPLAY INTEACTIVE LIMITED', '28th floor, Zuellig building, Paseo De Roxas cor. Makati Ave', 'Bel-Air, Makati City'],
  AP: ['APRICUS TECHNOLOGY INC.', '8/F Aguirre Bldg, 107 Legaspi Vill., San Lorenzo', 'Makati City 1223'] };
function words(n) {                                   // 26847 -> "Twenty-Six Thousand Eight Hundred Forty-Seven"
  const a = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const t = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  const h = x => [x >= 100 ? a[Math.floor(x / 100)] + ' Hundred' : '', x % 100 < 20 ? a[x % 100] : t[Math.floor(x % 100 / 10)] + (x % 10 ? '-' + a[x % 10] : '')].filter(Boolean).join(' ');
  if (!n) return 'Zero'; const out = []; [[1e9, 'Billion'], [1e6, 'Million'], [1e3, 'Thousand'], [1, '']].forEach(([v, w]) => { const k = Math.floor(n / v) % 1000; if (k) out.push(h(k) + (w ? ' ' + w : '')); });
  return out.join(' ');
}
function pesoWords(v) { const n = num(v) || 0, w = Math.floor(n), c = Math.round((n - w) * 100); return `${words(w)} Pesos${c ? ' and ' + words(c) + ' Centavos' : ''} Only`; }
async function prf(templateBytes, o) {
  const [cn, a1, a2] = COMPANIES[o.company] || COMPANIES.WR;
  const bytes = await xlsx(templateBytes, wb => {
    const p = wb.getWorksheet('PRF'), c = wb.getWorksheet('CA'), t = wb.getWorksheet('ATD');
    setv(p, 'L1', o.prfNo); setv(p, 'M1', (o.company || 'WR').toLowerCase()); setv(p, 'I3', cn); setv(p, 'I4', a1); setv(p, 'I5', a2);
    setv(p, 'H7', longDate(o.date)); setv(p, 'B8', o.payee); setv(p, 'H8', o.pprNo);
    [11, 13, 14, 15, 16, 17, 18].forEach((row, i) => { const it = o.items[i] || {}; setv(p, 'B' + row, longDate(it.date) || it.date || ''); setv(p, 'C' + row, it.no || ''); setv(p, 'D' + row, it.text || ''); setv(p, 'I' + row, num(it.amount)); });
    setv(p, 'D21', num(o.previous) ?? o.previous); setv(p, 'B26', o.comment); setv(p, 'B31', o.preparedBy); setv(p, 'H31', o.approvedBy); setv(p, 'C33', o.department);
    if (o.ca) {
      setv(c, 'H3', cn); setv(c, 'H4', a1); setv(c, 'H5', a2);
      setv(c, 'B7', o.caPayee); setv(c, 'B8', o.caDept); setv(c, 'B9', o.caId); setv(c, 'B11', o.caPurpose);
      setv(c, 'C12', longDate(o.caFrom)); setv(c, 'F12', longDate(o.caTo)); setv(c, 'D14', longDate(o.caLiquidate));
      setv(c, 'A17', o.caParticulars); setv(c, 'H17', num(o.caAmount)); setv(c, 'D28', o.caPrevious || 'N/A'); setv(c, 'A33', o.caComment);
      setv(c, 'A41', o.preparedBy); setv(c, 'D41', o.caCheckedBy); setv(c, 'G41', o.approvedBy); setv(c, 'B42', o.caDept);
    }
    if (o.atd && o.caPayee) {
      setv(t, 'A7', `          I, ${o.caPayee}, authorizes ${titleCo(cn)} to deduct the advances I claimed in case of failure to liquidate within the period specified in the amount of ${pesoWords(o.caAmount)}  (Php ${money(o.caAmount)})\n`);
      setv(t, 'A8', `          It is understood that the authority shall continue to be in effect for as I am employed with ${titleCo(cn)} or in the event of my resignation, the remaining amount shall be deducted from my last pay.\n`);
      setv(t, 'A10', `EMPLOYEE CONFORME:                    ${o.caPayee}        .           \n`);
    }
  });
  return { name: safe(`PRF ${o.prfNo || ''} ${o.payee || o.caPayee || ''} ${o.date || ''}`) + '.xlsx', bytes };
}
const titleCo = s => s.toLowerCase().replace(/\b([a-z])/g, m => m.toUpperCase()).replace(/\bInc\b/, 'Inc').replace(/\bLtd\b/, 'Ltd');
/* ---------- CTBC bank account opening (scanned pages: text placed at measured spots, y measured from the top, A4 595 x 842) */
async function ctbc(r, o) {
  const n = names(r), { PDFDocument, StandardFonts, rgb } = PDFLib;
  const doc = await PDFDocument.load(await fetchBuf('CTBC.pdf')), f = await doc.embedFont(StandardFonts.Helvetica), fb = await doc.embedFont(StandardFonts.HelveticaBold);
  const ink = rgb(0.05, 0.1, 0.45);
  const put = (p, x, y, s, size = 9, font = f, maxw = 0) => { s = String(s || '').trim().replace(/[^\x20-\x7E -ÿ]/g, ''); if (!s) return; const pg = doc.getPage(p - 1); let sz = size;
    if (maxw) while (sz > 5 && font.widthOfTextAtSize(s, sz) > maxw) sz -= 0.25;
    if (maxw && font.widthOfTextAtSize(s, sz) > maxw) {               // still too long: two lines
      const words = s.split(' '); let a = ''; while (words.length && font.widthOfTextAtSize((a + ' ' + words[0]).trim(), sz) <= maxw) a = (a + ' ' + words.shift()).trim();
      pg.drawText(a, { x, y: 842 - y + sz * 0.55, size: sz, font, color: ink }); pg.drawText(words.join(' '), { x, y: 842 - y - sz * 0.6, size: sz, font, color: ink }); return; }
    pg.drawText(s, { x, y: 842 - y, size: sz, font, color: ink }); };
  const tick = (p, x, y) => put(p, x, y, 'X', 9, fb);
  const [cn, a1, a2] = COMPANIES[o.company] || COMPANIES.WR, addr = `${a1} ${a2}`.replace(/\s+/g, ' ');
  const lastFirstMid = `${up(n.last)}, ${up(n.first)}${n.middle ? ' ' + up(n.middle) : ''}`, ms = o.title === 'Ms.';
  // page 1 - specimen signature card
  put(1, 14, 318, lastFirstMid, 10); put(1, 24, 390, up(n.natural), 9);
  // page 2 - client information file
  tick(2, o.title === 'Mr.' ? 33.5 : o.title === 'Mrs.' ? 91.5 : 61.5, 225);
  put(2, 168, 229, lastFirstMid, 9.5);
  tick(2, ms || o.title === 'Mrs.' ? 72.5 : 33.5, 251);
  put(2, 457, 251, up(o.nationality), 8.5);
  put(2, 30, 664, cn, 9); put(2, 30, 690, addr, 8.5); put(2, 30, 717, up(o.position || r['POSITION']), 9);
  // page 4 - card application
  tick(4, o.title === 'Mr.' ? 97.5 : o.title === 'Mrs.' ? 159 : 128.5, 127);
  put(4, 207, 128, up(n.last), 9); put(4, 311, 128, up(n.first), 9); put(4, 420, 128, up(n.middle), 9);
  put(4, 352, 147, up(o.nationality), 7.5);
  put(4, 90, 249, up(o.position || r['POSITION']), 7.5); put(4, 197, 239, `${cn}, ${addr}`, 7, f, 305);
  return { name: safe(`CTBC account opening ${r['EMPLOYEE ID']} ${n.natural}`) + '.pdf', bytes: await doc.save() };
}
window.WRForms = { ctbc, prf, COMPANIES, pesoWords, names, coe, confirmation, promotion, npa, appraisal, teamBuilding, erd, longDate, NPA_ACTIONS: Object.keys(NPA_TICK) };
})();
