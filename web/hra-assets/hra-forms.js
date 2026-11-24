/* Central HRA web - forms. Line-for-line port of the desktop app's hrapp/forms (pcforms.py, coe.py, epfetf.py, idphotos.py)
   onto pdf-lib. Character-box positions on the PC form templates come from geometry.json (measured once by the desktop code). */
(function () {
'use strict';
const ASSETS = (window.HRA_ASSETS || 'hra-assets/');
const { PDFDocument, StandardFonts, rgb } = PDFLib;
const K = 0.72, INK = [0.05, 0.12, 0.40];
const FONT_NAMES = { helv: StandardFonts.Helvetica, hebo: StandardFonts.HelveticaBold, heit: StandardFonts.HelveticaOblique,
                     tiro: StandardFonts.TimesRoman, tibo: StandardFonts.TimesRomanBold };
let GEO = null, MEAS = null;

async function fetchBytes(path) { const r = await fetch(ASSETS + path); if (!r.ok) throw new Error('Could not load ' + path + ' (' + r.status + ')'); return new Uint8Array(await r.arrayBuffer()); }
async function init() {
  if (!GEO) GEO = await (await fetch(ASSETS + 'geometry.json')).json();
  if (!MEAS) { const d = await PDFDocument.create(); MEAS = {}; for (const [k, v] of Object.entries(FONT_NAMES)) MEAS[k] = await d.embedFont(v); }
}

/* ---------- text helpers (same metrics as PyMuPDF's base-14 fonts) ---------- */
const okChar = {};
function clean(s, font) {           // keep only characters the standard PDF fonts can show
  let out = '';
  for (const ch of String(s)) {
    const key = font + ch;
    if (okChar[key] === undefined) { try { MEAS[font].widthOfTextAtSize(ch, 10); okChar[key] = true; } catch (e) { okChar[key] = false; } }
    out += okChar[key] ? ch : (ch === '…' ? '...' : '?');
  }
  return out;
}
// text width exactly as the desktop app (PyMuPDF) measures it: plain glyph widths, no kerning
const L = (t, size, font = 'helv') => {
  const s = clean(t, font), W = GEO.widths[font]; let w = 0;
  for (const ch of s) w += W[ch] !== undefined ? W[ch] : MEAS[font].widthOfTextAtSize(ch, 1000);
  return w * size / 1000;
};
const pyRound = x => { const f = Math.floor(x), d = x - f; return d > 0.5 ? f + 1 : d < 0.5 ? f : (f % 2 === 0 ? f : f + 1); };   // Python's round()
const median = a => { const s = [...a].sort((x, y) => x - y), n = s.length; return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2; };
const up = v => String(v ?? '').replace(/\s+/g, ' ').trim().toUpperCase();

/* ---------- dates (sheet dates are dd/mm/yyyy text) ---------- */
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
function mk(y, m, d) { const t = new Date(y, m - 1, d); return t.getFullYear() === y && t.getMonth() === m - 1 && t.getDate() === d ? { y, m, d } : null; }
function monthIdx(s) { const u = s.toLowerCase(); const i = MONTHS.findIndex(n => n.toLowerCase() === u || n.slice(0, 3).toLowerCase() === u); return i < 0 ? null : i + 1; }
function toDate(v) {
  if (v === null || v === undefined || v === '') return null;
  const s = String(v).trim(); let m;
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) { const r = mk(+m[3], +m[2], +m[1]); if (r) return r; }
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) { const r = mk(+m[1], +m[2], +m[3]); if (r) return r; }
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/))) { const yy = +m[3]; const r = mk(yy < 69 ? 2000 + yy : 1900 + yy, +m[2], +m[1]); if (r) return r; }
  if ((m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/))) { const mo = monthIdx(m[2]); const r = mo && mk(+m[3], mo, +m[1]); if (r) return r; }
  if ((m = s.match(/^(\d{1,2}) ([A-Za-z]+) (\d{4})$/))) { const mo = monthIdx(m[2]); const r = mo && mk(+m[3], mo, +m[1]); if (r) return r; }
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))) { const r = mk(+m[3], +m[1], +m[2]); if (r) return r; }
  return null;
}
const p2 = n => String(n).padStart(2, '0');
const dmy = d => d ? `${p2(d.d)}/${p2(d.m)}/${d.y}` : '';
const d8 = d => d ? `${p2(d.d)}${p2(d.m)}${d.y}` : '';
const longDate = d => d ? `${d.d} ${MONTHS[d.m - 1]} ${d.y}` : '';
const today = () => { const t = new Date(); return { y: t.getFullYear(), m: t.getMonth() + 1, d: t.getDate() }; };
const ageOn = (dob, t) => dob ? t.y - dob.y - ((t.m < dob.m || (t.m === dob.m && t.d < dob.d)) ? 1 : 0) : null;
const fmt2 = x => x.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const usd = x => x !== null && x !== undefined ? 'USD ' + fmt2(x) : '';
const safe = s => s.replace(/[\\/:*?"<>|,]/g, '').trim();

/* ---------- employee values (port of pcforms.values) ---------- */
const NAT_COUNTRY = { 'FILIPINO': 'PHILIPPINES', 'FILIPINOS': 'PHILIPPINES', 'THAI': 'THAILAND', 'VIETNAMESE': 'VIETNAM', 'INDONESIAN': 'INDONESIA',
  'CAMBODIAN': 'CAMBODIA', 'MALAYSIAN': 'MALAYSIA', 'CHINESE': 'CHINA', 'INDIAN': 'INDIA', 'NEPAL': 'NEPAL', 'NEPALESE': 'NEPAL',
  'MYANMAR': 'MYANMAR', 'SRI LANKAN': 'SRI LANKA', 'PAKISTANI': 'PAKISTAN', 'BANGLADESHI': 'BANGLADESH', 'KOREAN': 'SOUTH KOREA' };
function values(e) {
  const g = k => e[k];
  const s = k => { const v = up(g(k)); return ['NONE', 'N/A', 'NA', '-'].includes(v) ? '' : v; };
  const comp = s('COMPANY'), pcl = comp === 'PCL', fullDb = s('FULL NAME');
  const last = s('LAST NAME / SURNAME') || fullDb.split(',')[0].trim();
  const given = [s('FIRST NAME'), s('MIDDLE NAME')].filter(Boolean).join(' ') || (fullDb.includes(',') ? fullDb.slice(fullDb.indexOf(',') + 1).trim() : '');
  const passportName = s('NAME AS IN PASSPORT') || (given + ' ' + last).trim() || fullDb.replace(/,/g, ' ');
  const initials = given.split(/\s+/).filter(Boolean).map(t => t[0] + '.').join(' ');
  const local = s('EXPAT / LOCAL') === 'LOCAL', nic = s('NATIONAL ID NUMBER (LOCALS ONLY)'), pp = s('PASSPORT NUMBER');
  let join = pcl ? toDate(g('PC JOIN DATE (MM/DD/YYYY)')) : null;
  join = join || toDate(g('MROCK /REAL JOIN DATE (MM/DD/YY)')) || toDate(g('PC JOIN DATE (MM/DD/YYYY)'));
  const resign = toDate(g('REAL RESIGNATION DATE (MM/DD/YYYY)')) || toDate(g(pcl ? 'FOR PC RESIGNATION DATE (MM/DD/YYYY)' : 'MROCK RESIGNATION DATE (MM/DD/YYYY)'));
  const dob = toDate(g('BIRTH DATE (MM/DD/YYYY)'));
  const num = k => { const v = String(g(k) ?? '').replace(/,/g, '').replace(/USD/g, '').trim(); return /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(v) ? parseFloat(v) : null; };
  let tin = s('TIN NUMBER'); tin = /^[\d\- ]{5,}$/.test(tin) ? tin : '';
  const nat = s('NATIONALITY');
  const [employer, empAddr] = pcl ? ['PORT CITY BPO (PVT) LTD', 'AITKEN SPENCE TOWER 2, NO 315 VAUXHALL STREET, COLOMBO 02']
                                  : ['MASTEROCK LIMITED', '28TH FLOOR ONE GALLE FACE TOWER, 18 2 CENTRE ROAD, COLOMBO 02'];
  return {
    first: s('FIRST NAME'), middle: s('MIDDLE NAME'), id: s('SL EMP ID'), pcl_id: s('PCL EMP ID'), mr_id: s('MROCK EMP ID'), company: comp, pcl, local,
    active: s('ACTIVE/ INACTIVE') === 'ACTIVE', name: passportName, initials: (initials + ' ' + last).trim(), last,
    nickname: s('NICKNAME'), gender: s('GENDER (F / M)').slice(0, 1), nationality: nat, country: NAT_COUNTRY[nat] || s('COUNTRY OF BIRTH') || nat,
    dob, age: ageOn(dob, today()), birth_place: s('PLACE OF BIRTH'), birth_country: s('COUNTRY OF BIRTH') || NAT_COUNTRY[nat] || '',
    mobile: s('LOCAL CONTACT NUMBER'), email: String(g('PERSONAL EMAIL') ?? '').trim(), addr_sl: s('RESIDENCE ADDRESS IN SRI LANKA'),
    addr_home: s('PERMANENT ADDRESS'), civil: s('CIVIL STATUS'), nic, pp, id_no: (local && nic ? nic : pp), id_kind: (local && nic ? 'NIC' : 'PASSPORT'),
    pp_issue: toDate(g('DATE OF ISSUE (MM/DD/YYYY)')), pp_expiry: toDate(g('DATE OF EXPIRY (MM/DD/YYYY)')), pp_place: s('PASSPORT PLACE OF ISSUE'),
    visa_status: s('VISA STATUS'), visa_no: s('RESIDENCE VISA NUMBER'), visa_issue: toDate(g('MOST RECENT ISSUED DATE')), visa_expiry: toDate(g('MOST RECENT EXPIRY DATE')),
    join, resign, probation_end: toDate(g(pcl ? 'PROBATION ENDING DATE (PC)' : 'PROBATION ENDING DATE (MROCK)')),
    status: s('EMPLOYMENT STATUS'), rehire: s('ELIGIBILITY TO REHIRE'), supervisor: s('HEAD/LINE MANAGER/SUPERVISOR/TL'),
    position: s('POSITION (PC /MR DESIGNATION)') || s('DESIGNATION (INTERNAL)'), designation: s('DESIGNATION (INTERNAL)') || s('POSITION (PC /MR DESIGNATION)'),
    department: s('DEPARTMENT (INTERNAL)'), salary: num('BASIC SALARY (USD)'), fixed_allow: num('FIXED ALLOWANCE (USD)'), tin,
    bank: s('BANK NAME'), bank_acc: s('BANK ACCOUNT NUMBER'), bank_name: s('BENEFICIARY NAME'),
    spouse: s('SPOUSE FULL NAME'), spouse_nat: s('SPOUSE NATIONALITY'), spouse_pp: s('SPOUSE PASSPORT NUMBER'), spouse_addr: s('SPOUSE POSTAL ADDRESS'),
    spouse_exp: toDate(g('SPOUSE PASSPORT EXPIRY DATE')), emer_name: s('EMERGENCY CONTACT NAME'), emer_no: s('EMERGENCY CONTACT NUMBER'),
    epf: String(g('EPF MEMBER NO') ?? '').replace(/\D/g, ''), employer, emp_addr: empAddr,
  };
}

/* ---------- Filler: same drawing helpers as the desktop app's pcforms.Filler ---------- */
class Filler {
  static async open(file) {
    const f = new Filler(); f.file = file; f.ver = {}; f.fonts = {};
    f.doc = await PDFDocument.load(await fetchBytes('pcforms/' + file));
    for (const [k, v] of Object.entries(FONT_NAMES)) f.fonts[k] = await f.doc.embedFont(v);
    return f;
  }
  page(p) { return this.doc.getPage(p - 1); }
  box(p) { const pg = GEO.pages[this.file] && GEO.pages[this.file][p - 1]; return pg ? pg.cb : [0, 0, this.page(p).getWidth(), this.page(p).getHeight()]; }
  X(p, x) { return this.box(p)[0] + x; }
  Y(p, y) { return this.box(p)[3] - y; }
  insert(p, x, y, s, font = 'helv', size = 8.5, color = INK) {
    this.page(p).drawText(clean(s, font), { x: this.X(p, x), y: this.Y(p, y), size, font: this.fonts[font], color: rgb(...color) });
  }
  text(p, x, y, s, size = 8.5, maxw = null, font = 'helv', minsize = 6.5) {
    s = String(s ?? '').trim(); if (!s) return;
    const X = x * K, Y = y * K, Lf = (t, sz) => L(t, sz, font);
    let lines = [s];
    if (maxw) {
      const W = maxw * K;
      while (size > minsize && Lf(s, size) > W) size -= 0.25;
      if (Lf(s, size) > W) {
        for (const n of [2, 3]) {
          const words = s.split(' '); lines = []; let cur = ''; const target = s.length / n;
          for (const w of words) { if (cur && (cur.length + 1 + w.length > target * 1.15)) { lines.push(cur); cur = w; } else cur = (cur + ' ' + w).trim(); }
          lines.push(cur);
          let sz = size; while (sz > 5 && Math.max(...lines.map(l => Lf(l, sz))) > W) sz -= 0.25;
          if (Math.max(...lines.map(l => Lf(l, sz))) <= W || n === 3) { size = sz; break; }
        }
      }
    }
    const lh = size * 1.12, top = Y - (lines.length - 1) * lh / 2;
    lines.forEach((l, i) => this.insert(p, X, top + i * lh + size * 0.35, l, font, size));
  }
  rectPt(p, x0, y0, x1, y1, opts) { this.page(p).drawRectangle({ x: this.X(p, x0), y: this.Y(p, y1), width: x1 - x0, height: y1 - y0, ...opts }); }
  cover(p, x0, y0, x1, y1) { this.rectPt(p, x0 * K, y0 * K, x1 * K, y1 * K, { color: rgb(1, 1, 1) }); }
  poly(p, pts, width) {
    for (let i = 1; i < pts.length; i++) this.page(p).drawLine({ start: { x: this.X(p, pts[i - 1][0]), y: this.Y(p, pts[i - 1][1]) }, end: { x: this.X(p, pts[i][0]), y: this.Y(p, pts[i][1]) }, thickness: width, color: rgb(...INK), lineCap: 1 });
  }
  check_pt(p, X, Y, s = 0.9) { const d = 4.2 * s; this.poly(p, [[X - d, Y], [X - d * 0.3, Y + d * 0.75], [X + d, Y - d]], 1.4); }
  check(p, x, y, s = 1.0) { const X = x * K, Y = y * K, d = 4.2 * s; this.poly(p, [[X - d, Y], [X - d * 0.3, Y + d * 0.75], [X + d, Y - d]], 1.5); }
  circle(p, x0, y0, x1, y1) {
    const X0 = x0 * K, Y0 = y0 * K, X1 = x1 * K, Y1 = y1 * K;
    this.page(p).drawEllipse({ x: this.X(p, (X0 + X1) / 2), y: this.Y(p, (Y0 + Y1) / 2), xScale: (X1 - X0) / 2, yScale: (Y1 - Y0) / 2, borderColor: rgb(...INK), borderWidth: 1.1 });
  }
  strike(p, x0, y, x1) { this.page(p).drawLine({ start: { x: this.X(p, x0 * K), y: this.Y(p, y * K) }, end: { x: this.X(p, x1 * K), y: this.Y(p, y * K) }, thickness: 1.1, color: rgb(...INK) }); }
  cells(p, y, x0, x1) {
    const key = `${this.file}|${p}|${y}|${x0}|${x1}|${this.ver[p] || 0}`, c = GEO.cells[key];
    if (!c) { console.warn('no box geometry for', key); return []; }
    return c;
  }
  boxes(p, rows, text, wrap = true) {
    text = String(text ?? '').toUpperCase().trim(); if (!text) return;
    const all = rows.map(r => this.cells(p, ...r)).filter(c => c.length); if (!all.length) return;
    const words = wrap ? text.split(' ') : [text]; let ri = 0, ci = 0;
    for (const w of words) {
      const wl = [...w].length;
      while (ri < all.length) { const left = all[ri].length - ci; if (wl <= left || (ci === 0 && wl > all[ri].length)) break; ri++; ci = 0; }
      for (const ch of w) {
        if (ri >= all.length) return;
        if (ci >= all[ri].length) { ri++; ci = 0; if (ri >= all.length) return; }
        const [cx, cy, cw, chh] = all[ri][ci], size = Math.min(chh * 0.62, cw * 0.95, 9.5), tw = L(ch, size);
        this.insert(p, cx - tw / 2, cy + size * 0.36, ch, 'helv', size); ci++;
      }
      ci++; if (ri < all.length && ci >= all[ri].length) { ri++; ci = 0; }
    }
  }
  rebox(p, y, x0, x1, ref = null) {
    const own = this.cells(p, y, x0, x1), rc = ref ? this.cells(p, ...ref) : own;
    if (own.length && rc.length >= 2) {
      const xs = rc.map(c => c[0]), pitch = median(xs.slice(1).map((v, i) => v - xs[i])), bw = median(rc.map(c => c[2]));
      const cy = own[0][1], ch = own[0][3];
      let left = Math.min(...own.map(c => c[0] - c[2] / 2)), right = Math.max(...own.map(c => c[0] + c[2] / 2));
      if (ref && Math.abs(ref[1] - x0) < 3 && Math.abs(ref[2] - x1) < 3) { left = Math.min(...rc.map(c => c[0] - c[2] / 2)); right = Math.max(...rc.map(c => c[0] + c[2] / 2)); }
      const n = Math.max(1, pyRound((right - left - bw) / pitch) + 1);
      this.rectPt(p, Math.min(left, x0 * K) - 1.5, cy - ch / 2 - 1.5, Math.max(right, x1 * K) + 1.5, cy + ch / 2 + 1.5, { color: rgb(1, 1, 1) });
      for (let i = 0; i < n; i++) { const x = left + i * pitch; this.rectPt(p, x, cy - ch / 2, x + bw, cy + ch / 2, { borderColor: rgb(0, 0, 0), borderWidth: 1.0 }); }
    }
    this.ver[p] = (this.ver[p] || 0) + 1;
  }
  async save() { return await this.doc.save(); }
}

/* ---------- the forms (ports of pcforms.f_*) ---------- */
function f_info(V, F) {
  const t = (p, x, y, s, o = {}) => F.text(p, x, y, s, o.size ?? 8.5, o.maxw ?? null);
  t(1, 275, 172, V.position, { maxw: 158 }); t(1, 617, 172, V.id, { maxw: 95 });
  t(1, 272, 265, V.name, { maxw: 335 }); t(1, 272, 302, V.initials, { maxw: 335 }); t(1, 272, 338, V.nickname, { maxw: 335 });
  t(1, 272, 388, [(V.local ? V.nic : ''), V.pp].filter(Boolean).join(' / '), { maxw: 335 });
  t(1, 272, 440, V.addr_home, { maxw: 440 }); t(1, 272, 476, V.addr_sl, { maxw: 440 });
  t(1, 272, 512, dmy(V.dob), { maxw: 110 }); t(1, 566, 512, V.birth_place, { maxw: 148 });
  if (V.gender === 'M') F.circle(1, 266, 540, 303, 557);
  if (V.gender === 'F') F.circle(1, 304, 540, 349, 557);
  const c = V.civil;
  if (c.startsWith('SING')) F.circle(1, 266, 570, 306, 586); else if (c.startsWith('MARR')) F.circle(1, 307, 570, 351, 586); else if (c.startsWith('DIV')) F.circle(1, 266, 584, 315, 600);
  t(1, 566, 548, V.mobile, { maxw: 148 }); t(1, 272, 621, V.nationality, { maxw: 110 }); t(1, 566, 628, V.email, { maxw: 148, size: 8 });
  t(1, 272, 648, 'SRI LANKA', { maxw: 110 });
  if (V.local) { F.cover(1, 563, 650, 713, 667); t(1, 566, 658, 'LOCAL'); }
  t(1, 272, 729, V.pp, { maxw: 110 }); t(1, 566, 729, V.pp_place, { maxw: 148 });
  t(1, 272, 766, dmy(V.pp_issue), { maxw: 110 }); t(1, 566, 766, dmy(V.pp_expiry), { maxw: 148 });
  t(1, 272, 803, V.nationality, { maxw: 110 });
  t(1, 272, 876, V.emer_name, { maxw: 110 }); t(1, 566, 950, V.emer_no, { maxw: 148 });
  t(2, 566, 483, V.supervisor, { maxw: 148 });
  t(2, 230, 519, dmy(V.join) + (V.active ? ' - PRESENT' : (V.resign ? ' - ' + dmy(V.resign) : '')), { maxw: 150 });
  if (V.position) { F.cover(2, 227, 540, 381, 571); t(2, 230, 555, V.position, { maxw: 150, size: 8 }); }
  if (!V.pcl) { F.cover(2, 227, 467, 381, 501); t(2, 230, 488, V.employer, { maxw: 150 }); }
  t(3, 592, 694, V.bank_acc, { maxw: 122 }); t(3, 592, 735, V.bank_name, { maxw: 122, size: 7.5 });
  t(4, 228, 573, V.position, { maxw: 262, size: 8 }); t(4, 600, 573, V.status, { maxw: 112 });
  t(4, 228, 608, dmy(V.join), { maxw: 262 });
  if (V.probation_end) t(4, 228, 644, 'UNTIL ' + dmy(V.probation_end), { maxw: 262 });
  t(4, 600, 644, usd(V.fixed_allow), { maxw: 112 }); t(4, 228, 679, usd(V.salary), { maxw: 262 });
}
function f_checklist(V, F) {
  const t = (x, y, s) => F.text(1, x, y, s, 6.8, 192, 'helv', 4.6);
  const keys = [[146, 'id'], [161, 'name'], [176, 'nickname'], [191, 'nationality'], [206, 'position'], [222, 'department']];
  keys.forEach(([y, k]) => t(224, y, V[k])); t(224, 237, dmy(V.join));
  F.text(1, 95, 995, V.name, 7, 110);
  if (!V.active) { keys.forEach(([y, k]) => t(604, y, V[k])); t(604, 237, dmy(V.resign)); t(604, 305, V.rehire); F.text(1, 500, 995, V.name, 7, 110); }
}
function f_erd(V, F) {
  F.text(1, 265, 229, V.name, 10, 455); F.text(1, 265, 261, V.position, 10, 455); F.text(1, 265, 294, V.department, 10, 455);
  if (V.join) { F.cover(1, 262, 314, 500, 335); F.text(1, 268, 324, dmy(V.join), 10); }
}
function f_fatca(V, F) {
  F.text(1, 312, 108, V.id, 6.5, 55);
  if (V.local) F.check(1, 272, 331); else F.check(1, 90, 331);
  F.text(1, 75, 709, V.name, 11, 700);
  if (!V.local) {
    F.text(2, 432, 212, 'SRI LANKA', 9, 120);
    const addr = V.addr_home; let first = '';
    for (const w of addr.split(' ')) { if (L((first + ' ' + w).trim(), 8) <= 222 * K) first = (first + ' ' + w).trim(); else break; }
    F.text(2, 560, 231, first, 8, 225, 'helv', 6); F.text(2, 85, 251, addr.slice(first.length).trim(), 8, 700, 'helv', 6);
    F.text(2, 500, 267, V.nationality, 9, 285); F.text(2, 432, 290, V.country, 9, 350);
    F.text(2, 472, 377, V.pp, 9, 310); F.text(2, 530, 395, V.pp_place, 9, 255);
    F.text(2, 472, 412, V.visa_status, 9, 310); F.text(2, 528, 430, dmy(V.visa_expiry), 9, 255);
    F.text(2, 702, 448, V.tin, 9, 90);
  }
}
function f_edd(V, F) {
  const t = (y, s, size = 9) => F.text(1, 437, y, s, size, 278, 'helv', 5.2);
  t(304, V.name); t(323, (V.id_kind === 'NIC' ? 'NIC ' : 'PASSPORT ') + V.id_no);
  t(395, 'INDIVIDUAL'); t(431, 'SALARY ACCOUNT'); t(467, 'SRI LANKA'); t(503, V.position, 8);
  F.check(1, 442, 520);
  if (V.salary !== null) F.text(1, 462, 747, usd(V.salary) + ' (BASIC SALARY)', 9, 250);
  if (!V.local) F.check(2, 402, 228);
}
function f_immigration(V, F) {
  const B = (p, rows, s, wrap = true) => F.boxes(p, rows, s, wrap), t = (p, x, y, s, size = 8.5, maxw = null) => F.text(p, x, y, s, size, maxw);
  B(1, [[331, 247, 420]], V.pp, false); t(1, 541, 331, V.pp_place, 9, 135);
  B(1, [[373, 281, 420]], d8(V.pp_issue), false); B(1, [[373, 540, 678]], d8(V.pp_expiry), false);
  B(1, [[495, 92, 678], [515, 92, 678]], V.name);
  B(1, [[546, 177, 403]], V.nationality);
  if (V.gender === 'M') F.check(1, 238, 654);
  if (V.gender === 'F') F.check(1, 410, 654);
  B(1, [[693, 180, 317]], d8(V.dob), false);
  t(1, 302, 724, V.birth_place, 9, 115); B(1, [[724, 505, 678]], V.birth_country);
  const c = V.civil;
  for (const [key, x] of [['SING', 256], ['MARR', 376], ['WID', 513], ['DIV', 651]]) if (c.startsWith(key)) F.check(1, x, 760);
  if (c.startsWith('MARR')) {
    B(1, [[805, 179, 678], [825, 179, 678]], V.spouse); B(1, [[865, 179, 433]], V.spouse_nat);
    B(1, [[902, 179, 433]], V.spouse_pp, false); B(1, [[865, 453, 678], [883, 453, 678], [902, 453, 678]], V.spouse_addr);
    B(1, [[942, 316, 454]], d8(V.spouse_exp), false);
  }
  const home = V.addr_home;
  B(1, [[1057, 263, 678]], home);
  const cells1 = F.cells(1, 1057, 263, 678).length;
  if (home.length > cells1) { const cut = home.slice(0, cells1).lastIndexOf(' '); const rest = cut > 0 ? home.slice(cut + 1) : home.slice(cells1); B(2, [[25, 263, 678]], rest); }
  B(2, [[55, 263, 678], [73, 263, 678]], V.addr_sl);
  B(2, [[103, 263, 678]], V.position);
  B(2, [[138, 520, 678]], V.mobile.replace(/\D/g, ''), false);
  if (V.email) {
    const cells = F.cells(2, 163, 520, 678), pitch = cells.length > 1 ? cells[1][0] - cells[0][0] : 10;
    [...V.email].forEach((ch, i) => {
      let cx, cy; if (i < cells.length) [cx, cy] = cells[i]; else { cx = cells[cells.length - 1][0] + pitch * (i - cells.length + 1); cy = cells[cells.length - 1][1]; }
      const size = 8, tw = L(ch, size); F.insert(2, cx - tw / 2, cy + size * 0.36, ch, 'helv', size);
    });
  }
  if (V.visa_no) F.check(2, 531, 791); else F.check(2, 272, 791);
  B(2, [[1008, 368, 404]], '01', false); F.strike(2, 419, 1009, 456);
  if (V.visa_no) {
    B(2, [[1056, 263, 453]], V.visa_no.replace(/\//g, ''), false);
    B(3, [[44, 263, 402]], d8(V.visa_issue), false);
    B(3, [[70, 263, 678]], 'EMPLOYMENT');
  }
  B(3, [[183, 246, 403]], V.pp, false); B(3, [[183, 540, 678]], d8(V.pp_issue), false);
  B(3, [[222, 248, 386]], d8(V.pp_expiry), false); t(3, 541, 222, V.pp_place, 9, 135);
  t(5, 240, 575, V.name, 11, 560); t(5, 240, 633, V.pp, 11); t(5, 240, 685, V.nationality, 11);
  t(5, 215, 742, ({ M: 'MALE', F: 'FEMALE' })[V.gender] || '', 11);
  B(5, [[495, 568, 792]], V.pp, false);
}
function f_account(V, F) {
  const B = (p, rows, s, wrap = true) => F.boxes(p, rows, s, wrap);
  F.check(1, V.gender === 'M' ? 47 : 183, 305);
  B(1, [[329, 38, 788], [347, 38, 788]], V.name);
  B(1, [[383, 38, 788]], V.initials);
  F.check(1, V.id_kind === 'NIC' ? 47 : 223, 422);
  B(1, [[423, 340, 556]], V.id_no, false);
  B(1, [[423, 643, 788]], d8(V.dob), false);
  B(1, [[476, 38, 411], [494, 38, 411], [512, 38, 411]], V.addr_home);
  B(1, [[476, 430, 788], [494, 430, 788], [512, 430, 788]], V.addr_sl);
  B(1, [[540, 143, 788]], V.email.toUpperCase(), false);
  B(1, [[582, 38, 218]], V.mobile.replace(/\D/g, ''), false);
  if (V.tin) B(1, [[582, 588, 788]], V.tin, false);
  B(1, [[631, 38, 418]], V.position); B(1, [[631, 430, 788]], V.employer);
  B(1, [[677, 38, 788]], V.emp_addr);
  if (V.nationality === 'SRI LANKAN') { F.check(1, 47, 726); F.check(1, 438, 728); }
  else {
    F.check(1, 152, 726); F.text(1, 205, 723, V.nationality, 8, 160);
    if (V.visa_no || V.visa_status === 'RESIDENCE VISA') F.check(1, 556, 752); else F.check(1, 438, 752);
    F.text(1, 512, 797, V.visa_status.toLowerCase().replace(/(^|[^a-z])([a-z])/g, (m, a, b) => a + b.toUpperCase()), 7, 72);
    B(1, [[808, 30, 178]], d8(V.visa_expiry), false);
  }
  B(1, [[768, 38, 326]], 'SRI LANKA');
  F.text(1, 44, 876, 'SALARY ACCOUNT', 9, 300);
  F.check(1, 169, 901);
}
function f_abh(V, F) {
  const vals = {}, put = (names, text) => { [...String(text ?? '').toUpperCase()].forEach((ch, i) => { if (i < names.length) vals[names[i]] = ch; }); };
  const tb = (...nums) => nums.map(n => n === 1 ? 'Text Box 6' : `Text Box 6_${n}`), range = (a, b) => Array.from({ length: b - a }, (_, i) => a + i);
  put(tb(1, ...range(2, 13)), V.id_no.replace(/ /g, ''));
  put(tb(...range(21, 28)), V.epf); put(tb(...range(153, 160)), V.epf);
  put(tb(128, 129, 130, 131, 132, 133, 134, 135), d8(V.join));
  vals['Text Box 6_28'] = V.position; vals['Text Box 7_2'] = V.name;
  const ini = V.initials; put(tb(...range(29, 51)).concat(tb(...range(51, 73))), ini.padEnd(22).slice(0, 22) + ini.slice(22));
  vals['Text Box 31'] = V.addr_home;
  put(tb(73, 74, 75, 76, 77, 78, 79, 80), d8(V.dob));
  if (V.age !== null) put(tb(81, 82), p2(V.age));
  vals['Text Box 6_145'] = [V.birth_place, V.birth_country].filter(Boolean).join(', ');
  vals['Text Box 2'] = V.nationality;
  if (V.civil.startsWith('MARR') && V.spouse) {
    const sp = V.spouse.split(/\s+/).filter(Boolean);
    const spIni = sp.length > 1 ? sp.slice(0, -1).map(x => x[0] + '.').join(' ') + ' ' + sp[sp.length - 1] : V.spouse;
    put(['Text Box 11', 'Text Box 1', 'Text Box 8', 'Text Box 10'].concat(range(12, 30).map(n => `Text Box ${n}`)), spIni.trim());
  }
  put(tb(127, 136, 137, 138, 139, 140, 141, 142, 143, 144), V.mobile.replace(/\D/g, ''));
  vals['Text Box 9'] = V.email; vals['Text Box 4'] = V.name;
  const form = F.doc.getForm();
  for (const r of GEO.abh_radios) {
    const x = r.x0;
    const on = (r.name === 'Option Button 1_2' && ((x < 500 && V.gender === 'F') || (x > 500 && V.gender === 'M'))) ||
               (r.name === 'Option Button 1' && ((x < 300 && V.civil.startsWith('MARR')) || (x > 300 && V.civil.startsWith('SING'))));
    if (on) { try { form.getRadioGroup(r.name).select(r.on); } catch (e) { console.warn('radio', r, e); } }
  }
  for (const [name, v] of Object.entries(vals)) {
    if (!v) continue;
    try { const f = form.getTextField(name), max = f.getMaxLength(); f.setText(clean(max ? String(v).slice(0, max) : String(v), 'helv')); } catch (e) { console.warn('field', name, e.message); }
  }
}

/* Employee Info Form header: company drop-down instead of the Masterock logo; company name in the declaration */
const INFO_COMPANIES = ['Masterock', 'Whiterain', 'Imperium', 'Blueline', 'Netease', 'Nexzen'];
const INFO_LEGAL = { Masterock: 'MASTEROCK LIMITED' };
function infoCompanyFor(rowCompany) { const c = up(rowCompany); for (const n of INFO_COMPANIES) if (c.includes(n.toUpperCase()) || (n === 'Masterock' && ['MROCK', 'MR'].includes(c))) return n; return 'Masterock'; }
function infoHeader(F, company) {
  const legal = INFO_LEGAL[company] || company.toUpperCase();
  F.rectPt(1, 20, 18, 580, 84, { color: rgb(1, 1, 1) });
  for (const [x0, y0, x1, y1, bg] of GEO.mrinfo_hits) F.rectPt(1, x0 - 0.3, y0, x1 + 0.3, y1, { color: rgb(...bg) });
  for (const [x0, y0, x1, y1] of GEO.mrinfo_hits) {
    const w = x1 - x0, h = y1 - y0; let size = h * 0.78;
    while (size > 4 && L(legal, size, 'hebo') > w + 25) size -= 0.25;
    const tw = L(legal, size, 'hebo'); F.insert(1, x0 + Math.max(0, (w - tw) / 2), y1 - h * 0.22, legal, 'hebo', size, [0, 0, 0]);
  }
  F.insert(1, 30, 38, 'COMPANY', 'hebo', 7, [0.42, 0.48, 0.58]);
  const dd = F.doc.getForm().createDropdown('Company'); dd.addOptions(INFO_COMPANIES); dd.select(company);
  dd.addToPage(F.page(1), { x: F.X(1, 28), y: F.Y(1, 80), width: 272, height: 38, textColor: rgb(0.03, 0.07, 0.22), borderWidth: 0, font: F.fonts.hebo });
  dd.setFontSize(24); dd.defaultUpdateAppearances(F.fonts.hebo);
}
function f_mr_info(V, F) {
  infoHeader(F, V.form_company || infoCompanyFor(V.company));
  const put = (x, y, s, maxx, size = 8.0, twoLines = false) => {
    s = String(s ?? '').trim(); if (!s) return;
    const W = maxx - x - 4;
    if (twoLines && L(s, size) > W) {
      const words = s.split(' '); let first = '';
      while (words.length && L((first + ' ' + words[0]).trim(), size) <= W) first = (first + ' ' + words.shift()).trim();
      const rest = words.join(' ');
      F.insert(1, x, y, first, 'helv', size);
      const x2 = 30.5; let z = size; while (z > 5.5 && L(rest, z) > maxx - x2 - 4) z -= 0.25;
      F.insert(1, x2, y + 12, rest, 'helv', z); return;
    }
    while (size > 5.5 && L(s, size) > W) size -= 0.25;
    F.insert(1, x, y, s, 'helv', size);
  };
  const tick = r => F.check_pt(1, (r[0] + r[2]) / 2, (r[1] + r[3]) / 2), Y = top => top + 8.0;
  const first = V.first || V.name.split(' ')[0], last = V.last, middle = V.first ? V.middle : '';
  put(82, Y(183.6), first, 212); put(265, Y(183.6), middle, 395); put(440, Y(183.6), last, 568);
  put(93, Y(209.1), V.nickname, 212); put(272, Y(209.1), dmy(V.dob), 395);
  put(463, Y(209.1), [V.birth_place, V.birth_country].filter(Boolean).join(', '), 568);
  put(128, Y(240.3), V.addr_home, 568, 8.0, true);
  put(108, Y(271.5), V.addr_sl, 568, 8.0, true);
  if (V.gender === 'M') tick([28.7, 319.1, 37.7, 328.2]);
  if (V.gender === 'F') tick([102.4, 319.1, 111.4, 328.2]);
  for (const [key, box] of [['SING', [303.6, 319.1, 312.7, 328.2]], ['MARR', [371.7, 319.1, 380.7, 328.2]], ['DIV', [439.7, 319.1, 448.8, 328.2]], ['WID', [507.7, 319.1, 516.8, 328.2]]])
    if (V.civil.startsWith(key)) tick(box);
  put(143, Y(335.3), V.mobile, 303); put(371, Y(335.3), V.email.toLowerCase(), 568);
  put(362, Y(360.8), V.nationality, 568);
  put(111, Y(407.6), V.pp, 303); put(373, Y(407.6), V.pp_place, 568);
  put(95, Y(433.1), dmy(V.pp_issue), 303); put(374, Y(433.1), dmy(V.pp_expiry), 568);
  put(104, Y(485.5), V.emer_name, 303);
  put(286, Y(542.2), V.emer_no, 395);
}

const FORM_FILES = { MR_INFO: 'Masterock_Employment_Information_Form.pdf', INFO: 'PC_001_Info Form.pdf', IMMIGRATION: 'PC_002_Immigration form.pdf', ABH: 'PC_003_ABH Form.pdf',
  ACCOUNT: 'Account Opening Application Form.pdf', EDD: 'EDD Form.pdf', FATCA: 'FATCA Declaration.pdf', ERD: 'Employee_Relationship_Disclosure_Form_Final_2026 (ERD).pdf', CHECKLIST: null };
const FORMS_RUN = { MR_INFO: f_mr_info, INFO: f_info, IMMIGRATION: f_immigration, ABH: f_abh, ACCOUNT: f_account, EDD: f_edd, FATCA: f_fatca, ERD: f_erd, CHECKLIST: f_checklist };
const FORM_LABELS = { INFO: 'PC 001 - Info Form', IMMIGRATION: 'PC 002 - Immigration (expats only)', ABH: 'PC 003 - ABH Form', ACCOUNT: 'Account Opening Application',
  EDD: 'EDD Form', FATCA: 'FATCA Declaration', ERD: 'Employee Relationship Disclosure (ERD)', CHECKLIST: 'Employee Checklist (MR / PCL)' };
const RV_CATEGORIES = { 'Private Sector': [667, 214], 'Board of Investment (BOI)': [289, 214], 'Public Sector': [462, 214], 'Other - Port City (as printed)': null };
const MROCK_EMPLOYER = ['MASTEROCK LIMITED', '28TH FLR ONE GALLE FACE TOWER 18,2 CENTRE RD COLOMBO 02', '762271349'];

async function fill(row, keys, company = null) {
  await init();
  keys = keys.filter(k => FORMS_RUN[k]); if (!keys.length) keys = Object.keys(FORMS_RUN);
  const V = values(row); V.form_company = company;
  const files = [], notes = [], errs = [];
  for (const k of keys) {
    let fn = FORM_FILES[k]; if (k === 'CHECKLIST') fn = V.pcl ? 'Employee Checklist_PCL.pdf' : 'Employee Checklist_MR.pdf';
    if (!fn) continue;
    if (k === 'IMMIGRATION' && V.local) { notes.push('IMMIGRATION skipped - the residence visa form is for expats only'); continue; }
    try {
      const F = await Filler.open(fn); FORMS_RUN[k](V, F);
      const base = k === 'MR_INFO' ? `Employee Information Form (${V.form_company || infoCompanyFor(V.company)})` : fn.replace(/\.pdf$/i, '');
      files.push({ name: `${base} - ${V.id}.pdf`, bytes: await F.save() });
    } catch (e) { console.error(e); errs.push(`${k}: ${e.message}`); }
  }
  return { folder: safe(`${V.id} ${V.name}`), files, notes, errs, V };
}
async function residenceVisa(row, category = 'Private Sector') {
  await init();
  const V = values(row), F = await Filler.open(FORM_FILES.IMMIGRATION); f_immigration(V, F);
  const pos = RV_CATEGORIES[category];
  if (!V.pcl || pos) { F.rebox(2, 320, 228, 385, [381, 330, 490]); if (pos) F.check(2, ...pos); }
  if (!V.pcl) {
    const [name, addr, mob] = MROCK_EMPLOYER;
    for (const y of [427, 478, 497]) F.rebox(2, y, 160, 678, [447, 160, 678]);
    F.cover(2, 437, 524, 466, 544); F.rebox(2, 534, 280, 436, [566, 280, 436]);
    F.cover(2, 679, 556, 805, 577); F.rebox(2, 566, 522, 678, [534, 522, 678]);
    F.boxes(2, [[427, 160, 678], [447, 160, 678]], name);
    F.boxes(2, [[478, 160, 678], [497, 160, 678]], addr);
    F.boxes(2, [[534, 522, 678]], mob, false);
  }
  return { name: `Residence Visa Form (IM 35) - ${V.id}.pdf`, bytes: await F.save(), V };
}

/* ---------- Certificate of Employment (port of coe.py) ---------- */
const EMPLOYERS = { MROCK: ['MASTEROCK LIMITED', '28TH FLOOR ONE GALLE FACE TOWER 18,2 CENTRE ROAD COLOMBO 2, SRI LANKA', '0762271349'],
                    PCL: ['MASTEROCK LIMITED', '28TH FLOOR ONE GALLE FACE TOWER 18,2 CENTRE ROAD COLOMBO 2, SRI LANKA', '0762271349'] };
const titleCase = s => s.toLowerCase().replace(/(^|[^a-z])([a-z])/g, (m, a, b) => a + b.toUpperCase());
function coeFacts(r) {
  const g = k => up(r[k]), company = g('COMPANY'), [empName, empAddr, empMob] = EMPLOYERS[company] || EMPLOYERS.MROCK;
  const active = g('ACTIVE/ INACTIVE') === 'ACTIVE';
  const full = (g('NAME AS IN PASSPORT') || g('FULL NAME').replace(/,/g, ' ')).split(/\s+/).filter(Boolean).join(' ');
  const sal = ({ F: 'Ms.', M: 'Mr.' })[g('GENDER (F / M)').slice(0, 1)] || '';
  const surname = titleCase(g('FULL NAME').split(',')[0].trim());
  const nic = g('NATIONAL ID NUMBER (LOCALS ONLY)'), pp = g('PASSPORT NUMBER');
  const idText = g('EXPAT / LOCAL') === 'LOCAL' && nic ? `National ID No. ${nic}` : (pp ? `Passport No. ${pp}` : '');
  const pcl = company === 'PCL';
  const join = (pcl ? toDate(r['PC JOIN DATE (MM/DD/YYYY)']) : null) || toDate(r['MROCK /REAL JOIN DATE (MM/DD/YY)']) || toDate(r['PC JOIN DATE (MM/DD/YYYY)']);
  const end = toDate(r['REAL RESIGNATION DATE (MM/DD/YYYY)']) || (pcl ? toDate(r['FOR PC RESIGNATION DATE (MM/DD/YYYY)']) : null)
    || toDate(r['MROCK RESIGNATION DATE (MM/DD/YYYY)']) || toDate(r['FOR PC RESIGNATION DATE (MM/DD/YYYY)']);
  return { company, emp_name: empName, emp_addr: empAddr, emp_mob: empMob, active, full, sal, surname, id_text: idText, join, end, occ: g('POSITION (PC /MR DESIGNATION)'), emp_id: g('SL EMP ID') };
}
function coeWording(r) {
  const f = coeFacts(r);
  const body = `This is to certify that ${(f.sal + ' ' + f.full).trim()}` + (f.id_text ? `, holder of ${f.id_text},` : '') + (f.active ? ' has been employed with ' : ' was employed with ') + f.emp_name
    + (f.occ ? ` as ${f.occ}` : '') + ' from ' + (longDate(f.join) || '[join date not in the database]')
    + (f.active ? ' up to the present.' : ' until ' + (longDate(f.end) || '[end date not in the database]') + '.');
  const req = `This certification is issued upon the request of ${(f.sal + ' ' + f.surname).trim()} for whatever legal purpose it may serve.`;
  const ver = `For verification, please contact the Human Resources Department of ${f.emp_name}` + (f.emp_mob ? ` at ${f.emp_mob}.` : '.');
  const missing = [['join date', f.join], ['end date', f.active || f.end], ['ID / passport number', f.id_text]].filter(x => !x[1]).map(x => x[0]);
  return { f, paras: [body, req, ver], missing };
}
async function coePdf(r) {
  await init();
  const { f, paras } = coeWording(r), doc = await PDFDocument.create(), pg = doc.addPage([595, 842]), fonts = {};
  for (const [k, v] of Object.entries(FONT_NAMES)) fonts[k] = await doc.embedFont(v);
  const NAVY = [0.12, 0.22, 0.39], GOLD = [0.79, 0.64, 0.15], CINK = [0.1, 0.1, 0.1], H = 842;
  const put = (x, y, s, font, size, color) => pg.drawText(clean(s, font), { x, y: H - y, size, font: fonts[font], color: rgb(...color) });
  const L_ = 72, R = 523;
  put(L_, 80, f.emp_name, 'hebo', 16, NAVY);
  // insert_textbox (left aligned, 9 pt): first baseline = top + ascender
  const [ASC, DESC] = GEO.fontm.helv;          // PyMuPDF insert_textbox: baseline = top + size*asc, next line + size*asc*lineheight
  wrapWords(f.emp_addr, 'helv', 9, R - L_).forEach((ln, i) => put(L_, 88 + 9 * ASC + i * 9 * ASC * (ASC - DESC), ln, 'helv', 9, CINK));
  put(L_, 128, `Mobile: ${f.emp_mob}`, 'helv', 9, CINK);
  pg.drawLine({ start: { x: L_, y: H - 140 }, end: { x: R, y: H - 140 }, thickness: 1.5, color: rgb(...GOLD) });
  const td = longDate(today()); put(R - L(td, 11), 172, td, 'helv', 11, CINK);
  const title = 'CERTIFICATE OF EMPLOYMENT'; put((595 - L(title, 16, 'hebo')) / 2, 228, title, 'hebo', 16, NAVY);
  put(L_, 280, 'TO WHOM IT MAY CONCERN:', 'hebo', 11, CINK);
  let y = 305;
  for (const p of paras) {
    const lines = wrapWords(p, 'helv', 11, R - L_), lh = 11 * ASC * 1.5;
    lines.forEach((ln, i) => {
      const base = y + 11 * ASC + i * lh;
      if (i < lines.length - 1) {                       // justified: spread the words to the full width
        const words = ln.split(' '), wsum = words.reduce((a, w) => a + L(w, 11), 0), gap = words.length > 1 ? (R - L_ - wsum) / (words.length - 1) : 0;
        let x = L_; words.forEach(w => { put(x, base, w, 'helv', 11, CINK); x += L(w, 11) + gap; });
      } else put(L_, base, ln, 'helv', 11, CINK);
    });
    y += lines.length * 11 * 1.5 - DESC * 11 + 16;    // PyMuPDF's "used height" of the box
  }
  y = Math.max(y + 60, 600);
  pg.drawLine({ start: { x: L_, y: H - y }, end: { x: L_ + 190, y: H - y }, thickness: 0.8, color: rgb(...CINK) });
  ['Authorized Signatory', 'Human Resources Department', f.emp_name].forEach((s, i) => put(L_, y + 16 + i * 15, s, i === 0 ? 'hebo' : 'helv', 10, CINK));
  const note = 'Not valid without the company stamp and authorized signature.';
  put((595 - L(note, 8, 'heit')) / 2, 790, note, 'heit', 8, [0.4, 0.4, 0.4]);
  const t = today();
  return { name: safe(`COE ${f.emp_id} ${f.full} ${t.y}-${p2(t.m)}-${p2(t.d)}`) + '.pdf', bytes: await doc.save() };
}
function wrapWords(s, font, size, width) {
  const out = []; let cur = '';
  for (const w of String(s).split(' ')) { const t = cur ? cur + ' ' + w : w; if (!cur || L(t, size, font) <= width) cur = t; else { out.push(cur); cur = w; } }
  if (cur) out.push(cur); return out;
}

/* ---------- EPF / ETF claim forms (port of epfetf.py) ---------- */
const EMONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const numV = v => { const x = parseFloat(String(v ?? '').replace(/,/g, '').trim()); return isNaN(x) ? 0 : x; };
const mem = v => String(v ?? '').replace(/\D/g, '').replace(/^0+/, '');
function contributions(vals, memberNo) {
  const m = mem(memberNo), epf = {}, etf = {}; let epfName = '', etfName = '';
  if (!m) return { epf, etf, epfName, etfName };
  for (let r of vals.slice(3)) {
    r = [...r, ...Array(20).fill('')];
    if (mem(r[1]) === m && /^\d+$/.test(String(r[0]).trim())) {
      const k = +String(r[0]).trim(), [e, t] = epf[k] || [0, 0]; epf[k] = [e + numV(r[7]), t + numV(r[4])];
      if (k === Math.max(...Object.keys(epf).map(Number))) epfName = String(r[2]).trim();
    }
    if (mem(r[11]) === m && /^\d+$/.test(String(r[10]).trim())) {
      const k = +String(r[10]).trim(); etf[k] = (etf[k] || 0) + numV(r[15]);
      if (k === Math.max(...Object.keys(etf).map(Number))) etfName = `${String(r[12]).trim()} ${String(r[13]).trim()}`.trim();
    }
  }
  return { epf, etf, epfName, etfName };
}
function epfFacts(r, vals) {
  const g = k => String(r[k] ?? '').trim(), company = up(g('COMPANY')) || 'MROCK', [empName, empAddr, empMob] = EMPLOYERS[company] || EMPLOYERS.MROCK;
  const empNo = (vals.length > 3 && vals[3].length > 19 ? String(vals[3][19]).trim() : '') || 'A060108';
  const full = up(g('NAME AS IN PASSPORT')) || up(g('FULL NAME')).replace(/,/g, ' ').split(/\s+/).filter(Boolean).join(' ');
  const dob = toDate(r['BIRTH DATE (MM/DD/YYYY)']), age = ageOn(dob, today());
  const join = toDate(r['MROCK /REAL JOIN DATE (MM/DD/YY)']) || toDate(r['PC JOIN DATE (MM/DD/YYYY)']);
  const end = toDate(r['REAL RESIGNATION DATE (MM/DD/YYYY)']) || toDate(r['MROCK RESIGNATION DATE (MM/DD/YYYY)']);
  const nicLocal = g('NATIONAL ID NUMBER (LOCALS ONLY)'), pp = up(g('PASSPORT NUMBER'));
  const nic = up(g('EXPAT / LOCAL')) === 'LOCAL' && nicLocal ? nicLocal : (pp || nicLocal);
  const member = g('EPF MEMBER NO'), c = contributions(vals, member);
  const gender = up(g('GENDER (F / M)')).slice(0, 1), civil = up(g('CIVIL STATUS'));
  return { company, emp_name: empName, emp_addr: empAddr, emp_mob: empMob, emp_no: empNo, full, addr: g('PERMANENT ADDRESS') || g('RESIDENCE ADDRESS IN SRI LANKA'),
    spouse: g('SPOUSE FULL NAME'), dob, age, join, end, active: up(g('ACTIVE/ INACTIVE')) === 'ACTIVE', reason: g('REASON TO LEAVE'), nic,
    bank: g('BANK NAME'), acc_type: g('BANK ACCOUNT TYPE'), acc: g('BANK ACCOUNT NUMBER'), branch: g('BRANCH CODE'), mobile: g('LOCAL CONTACT NUMBER'),
    position: g('POSITION (PC /MR DESIGNATION)'), member, epf: c.epf, etf: c.etf, epf_name: c.epfName, etf_name: c.etfName,
    title: gender === 'M' ? 'Mr. ' : (civil.startsWith('MAR') ? 'Mrs. ' : 'Miss '), emp_id: up(g('SL EMP ID')) };
}
const NUM = x => ({ __num: x });
function valuesK(f) {
  const V = { F10: f.full, F11: f.addr, F14: f.spouse, F18: f.emp_no.slice(1), G18: f.emp_no.slice(0, 1), H18: f.member, I18: dmy(f.end),
    F22: f.dob ? `${f.age} years   -   ${dmy(f.dob)}` : '', F23: dmy(f.end), F24: f.reason, F25: `${f.emp_name}, ${f.emp_addr}`, F28: f.full, F29: f.nic,
    F33: f.acc_type, F34: f.bank, F35: f.acc, F36: f.branch, F43: f.full, F44: f.epf_name, F45: f.epf_name, F53: f.emp_no, E60: f.emp_no, I60: f.member };
  const keys = Object.keys(f.epf).map(Number).sort((a, b) => a - b), last = keys.length ? keys[keys.length - 1] : 0;
  V.B51 = `I / we ....................................................  Manager / Superintendent / Proprietor of  ${f.emp_name}  situated at  `
    + `${f.emp_addr}  do hereby certify that  ${f.full}  (Membership Number  ${f.member})  was employed as  ${f.position}`
    + '  in the above estate / establishment and that he / she is a member of the Employees Provident Fund. He / She is retiring / '
    + "leaving employment for the reason stated in paragraph 10 of his / her claim. His / Her contributions have been remitted in "
    + "full to the Central Bank and the last month's contributions were included in the return of the contribution for the month of  "
    + (last ? `${EMONTHS[last % 100 - 1]} ${Math.floor(last / 100)}` : '....................') + '.';
  const months = keys.slice(-6);
  months.forEach((k, i) => { const col = 'DEFGHI'[i]; V[col + '63'] = `${EMONTHS[k % 100 - 1].slice(0, 3)} ${Math.floor(k / 100)}`; V[col + '64'] = NUM(f.epf[k][0]); V[col + '65'] = NUM(f.epf[k][1]); });
  if (keys.length) { V.J64 = NUM(months.reduce((a, k) => a + f.epf[k][0], 0)); V.J65 = NUM(months.reduce((a, k) => a + f.epf[k][1], 0)); }
  return V;
}
function valuesVI(f) {
  const dd = d => d ? `${p2(d.d)} / ${p2(d.m)} / ${d.y}` : '';
  const V = { F11: f.title + (f.etf_name || f.epf_name), F12: f.full, F13: f.addr, F14: `Residence: ..................     Mobile: ${f.mobile}`,
    F15: f.dob ? `${dd(f.dob)}          Age: ${f.age}` : '', F16: f.nic, B20: dmy(f.join), D20: dmy(f.end), E20: f.emp_name, G20: f.emp_no, H20: f.member, I20: f.position,
    F24: f.reason, F28: f.bank, F29: f.branch, F30: f.acc, F31: f.acc_type, F36: f.etf_name || f.epf_name, F37: f.full, F38: dd(f.end), F39: f.nic,
    F40: f.reason, F41: `${f.emp_no}                    Member No:  ${f.member}`, F43: f.join ? `${f.join.y}   /   ${p2(f.join.m)}` : '',
    F64: f.emp_name, F65: f.emp_addr, F66: f.emp_mob };
  const etf = f.etf, keys = Object.keys(etf).map(Number);
  if (keys.length) {
    const last = Math.max(...keys), ly = Math.floor(last / 100), lm = last % 100; V.F44 = `${ly}   /   ${p2(lm)}`;
    const tot = [0, 0, 0];
    for (let i = 0; i < 6; i++) {
      const y = ly - i;
      if (keys.some(k => Math.floor(k / 100) === y)) {
        const h1 = keys.filter(k => Math.floor(k / 100) === y && k % 100 <= 6).reduce((a, k) => a + etf[k], 0);
        const h2 = keys.filter(k => Math.floor(k / 100) === y && k % 100 > 6).reduce((a, k) => a + etf[k], 0);
        V['C' + (48 + i)] = String(y); V['D' + (48 + i)] = NUM(h1); V['E' + (48 + i)] = NUM(h2); V['F' + (48 + i)] = NUM(h1 + h2);
        tot[0] += h1; tot[1] += h2; tot[2] += h1 + h2;
      }
    }
    V.D60 = NUM(tot[0]); V.E60 = NUM(tot[1]); V.F60 = NUM(tot[2]);
    let t12 = 0;
    for (let m = 1; m <= 12; m++) { const k = (m <= lm ? ly : ly - 1) * 100 + m; if (k in etf) { V['I' + (47 + m)] = NUM(etf[k]); t12 += etf[k]; } }
    V.I60 = NUM(t12);
  }
  return V;
}
async function renderClaim(form, vals) {
  await init();
  const spec = await (await fetch(ASSETS + `claimforms/${form}.json`)).json(), doc = await PDFDocument.load(await fetchBytes(`claimforms/${form}.pdf`)), fonts = {};
  for (const [k, v] of Object.entries(FONT_NAMES)) fonts[k] = await doc.embedFont(v);
  const geo = GEO.pages[form + '.pdf'];
  for (const [addr, raw] of Object.entries(vals)) {
    const fd = spec.fields[addr]; if (!fd || raw === null || raw === undefined || raw === '') continue;
    const num = typeof raw === 'object' && raw.__num !== undefined, s = num ? fmt2(raw.__num) : String(raw);
    const font = fd.serif ? (fd.bold ? 'tibo' : 'tiro') : (fd.bold ? 'hebo' : 'helv');
    let size = fd.size * (fd.serif ? 1.0 : 0.92);
    const color = [16, 8, 0].map(sh => ((fd.color >> sh) & 255) / 255);
    const [x0, , x1] = fd.bbox, width = fd.right - fd.left;
    const align = fd.halign !== 'general' ? fd.halign : (num ? 'right' : 'left');
    const avail = align !== 'left' ? width - 2 : fd.right - x0;
    const lines = fd.wrap ? wrapWords(s, font, size, avail) : [s];
    if (!fd.wrap) while (size > 5 && L(s, size, font) > avail) size -= 0.25;
    const lh = size * 1.18; let base = fd.baseline;
    if (lines.length > 1) { if (fd.valign === 'center') base -= (lines.length - 1) * lh / 2; else if (fd.valign === 'bottom') base -= (lines.length - 1) * lh; }
    const pg = doc.getPage(fd.page), cb = geo ? geo[fd.page].cb : [0, 0, pg.getWidth(), pg.getHeight()];
    lines.forEach((ln, i) => {
      const w = L(ln, size, font), x = align === 'right' ? fd.right - w : align === 'center' ? (x0 + x1) / 2 - w / 2 : x0;
      pg.drawText(clean(ln, font), { x: cb[0] + x, y: cb[3] - (base + i * lh), size, font: fonts[font], color: rgb(...color) });
    });
  }
  return await doc.save();
}
async function claim(which, r, vals) {
  await init();
  const f = epfFacts(r, vals), t = today(), stamp = `${t.y}-${p2(t.m)}-${p2(t.d)}`;
  if (which === 'K') return { name: safe(`EPF Form K ${f.emp_id} ${f.full} ${stamp}`) + '.pdf', bytes: await renderClaim('form_k', valuesK(f)), f };
  return { name: safe(`ETF Form VI ${f.emp_id} ${f.full} ${stamp}`) + '.pdf', bytes: await renderClaim('form_vi', valuesVI(f)), f };
}

/* ---------- ID photo sheet (port of idphotos.py): A4 at 300 dpi, 4 copies of 35 x 45 mm per employee ---------- */
async function idPhotos(people, photoBlobFor) {
  const DPI = 300, mm = v => Math.round(v / 25.4 * DPI);
  const PAGE_W = mm(210), PAGE_H = mm(297), PW = mm(35), PH = mm(45), HGAP = mm(8), VGAP = mm(8), BLOCK_GAP = mm(18);
  const TOP = mm(14), BOTTOM = mm(14), CAP_H = mm(7), ROW_GAP = mm(12), BLOCK_W = 2 * PW + HGAP, BLOCK_H = CAP_H + 2 * PH + VGAP;
  const pages = [], missing = [], left0 = Math.floor((PAGE_W - (2 * BLOCK_W + BLOCK_GAP)) / 2);
  const newPage = () => { const c = document.createElement('canvas'); c.width = PAGE_W; c.height = PAGE_H; const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, PAGE_W, PAGE_H); pages.push(c); return g; };
  const short = (s, n = 34) => s.length <= n ? s : s.slice(0, n - 1) + '…';
  let g = newPage(), y = TOP, slot = 0;
  for (const [emp, name] of people) {
    const blob = await photoBlobFor(emp); if (!blob) { missing.push(emp); continue; }
    if (slot === 0 && y + BLOCK_H > PAGE_H - BOTTOM) { g = newPage(); y = TOP; }
    const left = left0 + slot * (BLOCK_W + BLOCK_GAP);
    g.fillStyle = 'rgb(31,56,100)'; g.font = `${mm(3.0)}px Calibri, Arial, sans-serif`; g.textBaseline = 'alphabetic';
    g.fillText(`${emp}  ${short(name)}`, left, y + CAP_H - mm(1.5));
    const bmp = await createImageBitmap(blob, { imageOrientation: 'from-image' });
    let sx = 0, sy = 0, sw = bmp.width, sh = bmp.height; const target = PW / PH;
    if (sw / sh > target) { const nw = Math.floor(sh * target); sx = Math.floor((sw - nw) / 2); sw = nw; } else { const nh = Math.floor(sw / target); sy = Math.floor((sh - nh) * 0.3); sh = nh; }
    const top = y + CAP_H; g.imageSmoothingQuality = 'high';
    for (let k = 0; k < 4; k++) {
      const x = left + (k % 2) * (PW + HGAP), yy = top + Math.floor(k / 2) * (PH + VGAP);
      g.drawImage(bmp, sx, sy, sw, sh, x, yy, PW, PH);
      g.strokeStyle = 'rgb(170,170,170)'; g.lineWidth = 1; g.strokeRect(x - 0.5, yy - 0.5, PW + 1, PH + 1);
    }
    if (++slot === 2) { slot = 0; y += BLOCK_H + ROW_GAP; }
  }
  if (slot === 1) y += BLOCK_H + ROW_GAP;
  if (missing.length) {
    if (y + mm(8) > PAGE_H - BOTTOM) { g = newPage(); y = TOP; }
    g.fillStyle = 'rgb(192,0,0)'; g.font = `${mm(2.4)}px Calibri, Arial, sans-serif`; g.fillText('No photo on file for: ' + missing.join(', '), left0, y + mm(5));
  }
  const doc = await PDFDocument.create();
  for (const c of pages) {
    const jpg = await new Promise(res => c.toBlob(b => res(b), 'image/jpeg', 0.92)), img = await doc.embedJpg(new Uint8Array(await jpg.arrayBuffer()));
    const pg = doc.addPage([210 / 25.4 * 72, 297 / 25.4 * 72]); pg.drawImage(img, { x: 0, y: 0, width: pg.getWidth(), height: pg.getHeight() });
  }
  const t = today(), ids = people.slice(0, 4).map(p => p[0]).join(' ') + (people.length > 4 ? ' +more' : '');
  return { name: `ID Photos ${t.y}-${p2(t.m)}-${p2(t.d)} ${ids}.pdf`, bytes: await doc.save(), missing };
}

window.HRAForms = { init, values, fill, residenceVisa, coeWording, coePdf, claim, epfFacts, idPhotos, infoCompanyFor,
  INFO_COMPANIES, FORM_LABELS, RV_CATEGORIES };
})();
