/**
 * PORMODiT Asset – ระบบบริหารจัดการครุภัณฑ์และพัสดุคลังย่อย (เทศบาล / อบต.)
 * ไฟล์: Code.gs  (คู่กับ index.html)
 * เริ่มต้น: รันฟังก์ชัน setupSystem() หนึ่งครั้งเพื่ออนุญาตสิทธิ์ แล้ว Deploy เป็น Web App
 * ผู้ใช้เริ่มต้น: admin / dome1234  (เปลี่ยนรหัสผ่านทันที)
 */
const SHEET_ID = '1zrKwJHWD-Dk7iQKnguMcldAVMhkW5jZp1ZCUvC-X818'; // ไอดีชีต
const FOLDER_ID = '1Dtef5Z3voEZ-1c6Q6-dPOrwFVnp4gMhJ'; // ไอดีโฟลเดอร์
const TZ = 'Asia/Bangkok';
const ST = { ok: 'ใช้งานได้', loan: 'ถูกยืม', rep: 'รอซ่อม', bad: 'ชำรุด', pend: 'รอจำหน่าย', gone: 'จำหน่ายแล้ว' };
const AUD = ['CreatedAt', 'CreatedBy', 'UpdatedAt', 'UpdatedBy'];

const SCHEMA = {
  Users: ['UserID', 'Username', 'FullName', 'Role', 'DeptID', 'PasswordHash', 'Salt', 'TelegramChatID', 'PhotoUrl', 'Status'],
  Roles: ['Role', 'RoleName', 'Description'],
  Departments: ['DeptID', 'DeptName', 'Status'],
  Locations: ['LocationID', 'LocationName', 'DeptID', 'Status'],
  AssetCategories: ['CategoryID', 'CategoryName', 'Type', 'Status'],
  Assets: ['AssetID', 'RegNo', 'Name', 'CategoryID', 'Brand', 'Model', 'SerialNo', 'Qty', 'Unit', 'Cost', 'ReceivedDate', 'InspectedDate', 'PurchaseDocNo', 'FundSource', 'DeptID', 'LocationID', 'Custodian', 'LifeYears', 'DepStartDate', 'SalvageValue', 'Status', 'ImageId', 'Notes'],
  AssetDocuments: ['DocID', 'AssetID', 'DocType', 'FileName', 'FileUrl', 'FileId'],
  AssetHistory: ['HistID', 'AssetID', 'Action', 'Detail', 'ActionAt', 'ActionBy'],
  DepreciationPolicies: ['PolicyID', 'CategoryID', 'LifeYears', 'SalvageType', 'SalvageValue', 'Method', 'Status'],
  DepreciationRecords: ['RecordID', 'AssetID', 'FiscalYear', 'AsOfDate', 'Depreciation', 'Accumulated', 'BookValue'],
  Loans: ['LoanID', 'DocNo', 'Borrower', 'DeptID', 'Purpose', 'LoanDate', 'DueDate', 'Status', 'RequestedBy', 'ApprovedBy', 'ApprovedAt', 'DecisionNote', 'ReturnDate', 'ReturnNote'],
  LoanDetails: ['DetailID', 'LoanID', 'AssetID', 'ConditionOut', 'ConditionIn'],
  Transfers: ['TransferID', 'DocNo', 'AssetID', 'FromDeptID', 'ToDeptID', 'ToLocationID', 'ToCustodian', 'Reason', 'Status', 'RequestedBy', 'ApprovedBy', 'ApprovedAt', 'ReceivedBy', 'ReceivedAt'],
  Materials: ['MaterialID', 'Code', 'Name', 'Unit', 'CategoryID', 'ReorderPoint', 'Status'],
  StockLots: ['LotID', 'MaterialID', 'ReceiptDocNo', 'ReceivedDate', 'QtyIn', 'QtyRemain', 'UnitCost', 'Supplier'],
  StockTransactions: ['TxID', 'MaterialID', 'LotID', 'Type', 'Qty', 'UnitCost', 'RefDoc', 'TxDate', 'ByUser', 'DeptID'],
  Requisitions: ['ReqID', 'DocNo', 'DeptID', 'RequestedBy', 'Purpose', 'ReqDate', 'Status', 'ApprovedBy', 'ApprovedAt', 'DecisionNote', 'IssuedBy', 'IssuedAt'],
  RequisitionDetails: ['DetailID', 'ReqID', 'MaterialID', 'Qty', 'IssuedQty'],
  Maintenance: ['MaintID', 'DocNo', 'AssetID', 'Problem', 'ReportedBy', 'ReportDate', 'Status', 'ApprovedBy', 'ApprovedAt', 'ResultDate', 'ResultNote', 'Cost', 'Technician'],
  MaintenanceHistory: ['HistID', 'MaintID', 'Status', 'Note', 'At', 'By'],
  AnnualInspections: ['RoundID', 'FiscalYear', 'Name', 'StartDate', 'EndDate', 'Committee', 'Status'],
  InspectionAssignments: ['AssignID', 'RoundID', 'InspectorUser', 'DeptID'],
  InspectionResults: ['ResultID', 'RoundID', 'AssetID', 'Result', 'FoundLocationID', 'Note', 'PhotoId', 'Lat', 'Lng', 'InspectedBy', 'InspectedAt', 'FollowStatus', 'FollowNote'],
  Disposals: ['DisposalID', 'DocNo', 'AssetID', 'Reason', 'Status', 'ReportDate', 'CommitteeOrder', 'FindingNote', 'ApprovedBy', 'ApprovedAt', 'Method', 'Evidence', 'CompletedAt'],
  Documents: ['DocNo', 'DocType', 'RefID', 'IssuedAt'],
  Notifications: ['NotifID', 'Type', 'Message', 'RefID', 'Key', 'Sent', 'SentAt'],
  AuditLogs: ['LogID', 'At', 'UserID', 'Username', 'Action', 'Entity', 'EntityID', 'Before', 'After'],
  Settings: ['Key', 'Value'],
  FiscalYears: ['FiscalYear', 'StartDate', 'EndDate', 'Status'],
  NumberSequences: ['SeqKey', 'Prefix', 'FiscalYear', 'Last']
};

/* ============================ Utilities ============================ */
let _SS = null, _M = {};
const ss = () => _SS || (_SS = SpreadsheetApp.openById(SHEET_ID));
const headers = n => SCHEMA[n].concat(AUD);
const now = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
const today = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
const addDays = n => { const d = new Date(); d.setDate(d.getDate() + n); return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); };
const uid = p => p + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
const fmtv = v => v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss').replace(' 00:00:00', '') : v;
const pD = s => { const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
const fyOf = s => { const d = pD(s) || new Date(); return d.getFullYear() + (d.getMonth() >= 9 ? 1 : 0) + 543; };

function sh(n) {
  let s = ss().getSheetByName(n);
  const h = headers(n);
  if (!s) {
    s = ss().insertSheet(n);
    const sd = ss().getSheetByName('Sheet1') || ss().getSheetByName('ชีต1');
    if (sd && ss().getSheets().length > 1 && sd.getLastRow() === 0) ss().deleteSheet(sd);
  }
  if (s.getLastRow() === 0 || String(s.getRange(1, 1).getValue()) === '') {
    if (s.getMaxColumns() < h.length) s.insertColumnsAfter(s.getMaxColumns(), h.length - s.getMaxColumns());
    s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#0B1F4B').setFontColor('#FFFFFF');
    s.setFrozenRows(1);
    s.getRange(2, 1, Math.max(s.getMaxRows() - 1, 1), h.length).setNumberFormat('@');
  }
  return s;
}
function rows(n) {
  if (_M[n]) return _M[n];
  const v = sh(n).getDataRange().getValues();
  if (v.length < 2) return (_M[n] = []);
  const h = v[0];
  return (_M[n] = v.slice(1).filter(r => r[0] !== '').map(r => { const o = {}; h.forEach((k, i) => o[k] = fmtv(r[i])); return o; }));
}
function insMany(n, list, u) {
  if (!list.length) return [];
  const h = headers(n), t = now(), by = u ? u.username : 'SYSTEM', s = sh(n);
  const out = list.map(o => Object.assign({}, o, { CreatedAt: o.CreatedAt || t, CreatedBy: o.CreatedBy || by, UpdatedAt: t, UpdatedBy: by }));
  const vals = out.map(o => h.map(k => o[k] === undefined || o[k] === null ? '' : o[k]));
  const start = s.getLastRow() + 1;
  if (start + vals.length - 1 > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), start + vals.length - 1 - s.getMaxRows());
  s.getRange(start, 1, vals.length, h.length).setNumberFormat('@').setValues(vals);
  delete _M[n];
  return out;
}
const ins = (n, o, u) => insMany(n, [o], u)[0];
function upd(n, idk, id, patch, u) {
  const s = sh(n), v = s.getDataRange().getValues(), h = v[0], ci = h.indexOf(idk);
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][ci]) === String(id)) {
      const before = {}; h.forEach((k, j) => before[k] = fmtv(v[i][j]));
      const after = Object.assign({}, before, patch, { UpdatedAt: now(), UpdatedBy: u ? u.username : 'SYSTEM' });
      s.getRange(i + 1, 1, 1, h.length).setNumberFormat('@').setValues([h.map(k => after[k] === undefined || after[k] === null ? '' : after[k])]);
      delete _M[n];
      return { before: before, after: after };
    }
  }
  throw new Error('ไม่พบข้อมูล ' + id);
}
const find = (n, k, id) => rows(n).find(r => String(r[k]) === String(id));
const map = (n, k, v) => { const m = {}; rows(n).forEach(r => m[r[k]] = r[v]); return m; };
const must = (n, k, id, msg) => { const r = find(n, k, id); if (!r) throw new Error(msg || 'ไม่พบข้อมูล'); return r; };
const num = v => Number(v) || 0;
function diff(b, a) { const o = []; Object.keys(a).forEach(k => { if (AUD.indexOf(k) < 0 && String(b[k]) !== String(a[k])) o.push(k + ': ' + b[k] + ' → ' + a[k]); }); return o.join('; ') || '(ไม่มีการเปลี่ยนแปลง)'; }
function audit(u, action, entity, id, before, after) {
  const j = x => x === undefined || x === null ? '' : String(typeof x === 'string' ? x : JSON.stringify(x)).slice(0, 40000);
  ins('AuditLogs', { LogID: uid('LOG'), At: now(), UserID: u ? u.id : 'SYSTEM', Username: u ? u.username : 'SYSTEM', Action: action, Entity: entity, EntityID: id, Before: j(before), After: j(after) }, u);
}
const hist = (aid, act, detail, u) => ins('AssetHistory', { HistID: uid('H'), AssetID: aid, Action: act, Detail: detail || '', ActionAt: now(), ActionBy: u ? u.name : 'SYSTEM' }, u);
function nextNo(key, type) {
  const fy = fyOf(today()), sk = key + '-' + fy, r = find('NumberSequences', 'SeqKey', sk);
  let last = 1;
  if (r) { last = num(r.Last) + 1; upd('NumberSequences', 'SeqKey', sk, { Last: last }); }
  else ins('NumberSequences', { SeqKey: sk, Prefix: key, FiscalYear: fy, Last: 1 });
  const no = key + '-' + fy + '/' + ('0000' + last).slice(-4);
  ins('Documents', { DocNo: no, DocType: type || key, RefID: '', IssuedAt: now() });
  return no;
}
const getSet = k => { const r = find('Settings', 'Key', k); return r ? r.Value : ''; };
function setSet(k, v, u) { if (find('Settings', 'Key', k)) upd('Settings', 'Key', k, { Value: v }, u); else ins('Settings', { Key: k, Value: v }, u); }

/* ============================ Auth ============================ */
function hashPw(pw, salt) {
  let d = salt + pw;
  for (let i = 0; i < 1000; i++) d = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, d + salt, Utilities.Charset.UTF_8));
  return d;
}
function login(p) {
  const un = String(p.username || '').trim().toLowerCase(), c = CacheService.getScriptCache(), fk = 'F_' + un;
  if (+c.get(fk) >= 5) throw new Error('ลองผิดหลายครั้ง กรุณารอ 10 นาที');
  const r = rows('Users').find(x => String(x.Username).toLowerCase() === un && x.Status === 'ใช้งาน');
  if (!r || hashPw(String(p.password || ''), r.Salt) !== r.PasswordHash) {
    c.put(fk, String((+c.get(fk) || 0) + 1), 600);
    throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  }
  c.remove(fk);
  const tok = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  const u = { id: r.UserID, username: r.Username, name: r.FullName, role: r.Role, dept: r.DeptID, photoUrl: r.PhotoUrl || '' };
  c.put('S_' + tok, JSON.stringify(u), 21600);
  audit(u, 'LOGIN', 'Users', u.id);
  return { token: tok, user: u };
}
function sess(t) {
  if (!t) return null;
  const c = CacheService.getScriptCache(), s = c.get('S_' + t);
  if (!s) return null;
  c.put('S_' + t, s, 21600);
  return JSON.parse(s);
}
function logout(p, u) { CacheService.getScriptCache().remove('S_' + p._t); audit(u, 'LOGOUT', 'Users', u.id); return true; }
function pwChange(p, u) {
  const r = must('Users', 'UserID', u.id);
  if (hashPw(String(p.old || ''), r.Salt) !== r.PasswordHash) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
  if (String(p.pw || '').length < 8) throw new Error('รหัสผ่านใหม่ต้องยาวอย่างน้อย 8 ตัวอักษร');
  const salt = Utilities.getUuid();
  upd('Users', 'UserID', u.id, { Salt: salt, PasswordHash: hashPw(p.pw, salt) }, u);
  audit(u, 'CHANGE_PASSWORD', 'Users', u.id);
  return true;
}

/* ============================ Web entry ============================ */
function doGet(e) {
  ensureSetup();
  const t = HtmlService.createTemplateFromFile('index');
  t.pre = JSON.stringify({ a: (e && e.parameter && e.parameter.a) || '' }).replace(/</g, '\\u003c');
  return t.evaluate().setTitle('ระบบบริหารจัดการครุภัณฑ์และพัสดุ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function ensureSetup() {
  const P = PropertiesService.getScriptProperties();
  if (P.getProperty('SETUP_V') === '1') return;
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try { if (P.getProperty('SETUP_V') !== '1') { setupSystem(); P.setProperty('SETUP_V', '1'); } } finally { lock.releaseLock(); }
}

/* ============================ Boot / list ============================ */
function boot(p, u) {
  const act = r => r.filter(x => x.Status !== 'ยกเลิกใช้');
  const usr = rows('Users').filter(x => x.Status === 'ใช้งาน');
  // ดึง PhotoUrl ล่าสุดจากชีต (กรณีมีการอัปเดตรูปหลัง login)
  const me = find('Users', 'UserID', u.id);
  const user = Object.assign({}, u, { photoUrl: (me && me.PhotoUrl) ? me.PhotoUrl : (u.photoUrl || '') });
  return {
    user: user, appUrl: ScriptApp.getService().getUrl(), org: getSet('ORG_NAME') || 'เทศบาลตำบลบึงโขงหลง',
    defaultPw: u.username === 'admin' && hashPw('admin1234', must('Users', 'UserID', u.id).Salt) === must('Users', 'UserID', u.id).PasswordHash,
    depts: act(rows('Departments')), locs: act(rows('Locations')), cats: act(rows('AssetCategories')),
    users: usr.map(x => ({ id: x.UserID, name: x.FullName, role: x.Role, dept: x.DeptID, username: x.Username, PhotoUrl: x.PhotoUrl || '' }))
  };
}
const LIST = {
  assets: ['Assets'], assetDocs: ['AssetDocuments'], assetHist: ['AssetHistory'], policies: ['DepreciationPolicies'], depRecs: ['DepreciationRecords'],
  loans: ['Loans'], loanDetails: ['LoanDetails'], transfers: ['Transfers'], materials: ['Materials'], lots: ['StockLots'], tx: ['StockTransactions'],
  reqs: ['Requisitions'], reqDetails: ['RequisitionDetails'], maint: ['Maintenance'], maintHist: ['MaintenanceHistory'], rounds: ['AnnualInspections'],
  assigns: ['InspectionAssignments'], results: ['InspectionResults'], disposals: ['Disposals'], notifs: ['Notifications'],
  audit: ['AuditLogs', ['admin', 'approver']], users: ['Users', ['admin']], depts: ['Departments'], locs: ['Locations'], cats: ['AssetCategories'], fys: ['FiscalYears'], docs: ['Documents']
};
function listEntity(p, u) {
  const d = LIST[p.e]; if (!d) throw new Error('ไม่รู้จักข้อมูล');
  if (d[1] && d[1].indexOf(u.role) < 0) throw new Error('คุณไม่มีสิทธิ์ดูข้อมูลนี้');
  let r = rows(d[0]);
  if (p.e === 'users') r = r.map(x => { const o = Object.assign({}, x); delete o.PasswordHash; delete o.Salt; return o; });
  if (p.e === 'audit') r = r.slice(-800).reverse();      // เดิม 1500
  if (p.e === 'assetHist') r = r.slice(-300).reverse();
  if (p.e === 'tx') r = r.slice(-500).reverse();
  return r;
}

/* ============================ Depreciation ============================ */
function polMap() { const m = {}; rows('DepreciationPolicies').filter(x => x.Status !== 'ยกเลิกใช้').forEach(x => m[x.CategoryID] = x); return m; }
function depOf(a, date, pol) {
  const cost = num(a.Cost), pc = pol[a.CategoryID];
  const life = num(a.LifeYears) || (pc ? num(pc.LifeYears) : 0);
  let salv = 0;
  if (a.SalvageValue !== '' && a.SalvageValue !== undefined) salv = num(a.SalvageValue);
  else if (pc) salv = pc.SalvageType === 'percent' ? cost * num(pc.SalvageValue) / 100 : num(pc.SalvageValue);
  const start = pD(a.DepStartDate) || pD(a.ReceivedDate);
  if (!life || !start || cost <= 0) return { ok: false, cost: cost, acc: 0, book: cost, life: life, salv: salv, annual: 0 };
  const base = Math.max(cost - salv, 0), annual = base / life;
  const days = (date - start) / 864e5;
  const acc = days <= 0 ? 0 : Math.min(base * days / (life * 365.25), base);
  return { ok: true, cost: cost, acc: acc, book: cost - acc, life: life, salv: salv, annual: annual, base: base };
}
function depCalc(p) {
  const asOf = pD(p.asOf) || new Date(), pol = polMap(), cn = map('AssetCategories', 'CategoryID', 'CategoryName');
  const list = rows('Assets').filter(a => a.Status !== ST.gone).map(a => {
    const d = depOf(a, asOf, pol);
    return { AssetID: a.AssetID, RegNo: a.RegNo, Name: a.Name, Cat: cn[a.CategoryID] || '', DeptID: a.DeptID, Start: a.DepStartDate || a.ReceivedDate, Cost: d.cost, Life: d.life, Salv: d.salv, Annual: d.annual, Acc: d.acc, Book: d.book, Ok: d.ok };
  });
  return { asOf: p.asOf || today(), list: list };
}
function depSave(p, u) {
  const c = depCalc(p), fy = fyOf(c.asOf), ex = {};
  rows('DepreciationRecords').forEach(r => ex[r.AssetID + '|' + r.AsOfDate] = 1);
  const add = c.list.filter(x => x.Ok && !ex[x.AssetID + '|' + c.asOf]).map(x => ({ RecordID: uid('DR'), AssetID: x.AssetID, FiscalYear: fy, AsOfDate: c.asOf, Depreciation: x.Annual.toFixed(2), Accumulated: x.Acc.toFixed(2), BookValue: x.Book.toFixed(2) }));
  insMany('DepreciationRecords', add, u);
  audit(u, 'DEP_SAVE', 'DepreciationRecords', c.asOf, null, { saved: add.length });
  return { saved: add.length };
}

/* ============================ Dashboard / alerts ============================ */
function stockBal() { const m = {}; rows('StockLots').forEach(l => { const q = num(l.QtyRemain); const o = m[l.MaterialID] = m[l.MaterialID] || { q: 0, v: 0 }; o.q += q; o.v += q * num(l.UnitCost); }); return m; }
function dashboard() {
  const A = rows('Assets').filter(a => a.Status !== ST.gone), pol = polMap(), t = new Date();
  const cn = map('AssetCategories', 'CategoryID', 'CategoryName'), dn = map('Departments', 'DeptID', 'DeptName');
  let cost = 0, acc = 0; const cat = {}, dep = {}, st = {};
  A.forEach(a => { const d = depOf(a, t, pol); cost += d.cost; acc += d.acc; const c = cn[a.CategoryID] || 'ไม่ระบุ', dd = dn[a.DeptID] || 'ไม่ระบุ'; cat[c] = (cat[c] || 0) + 1; dep[dd] = (dep[dd] || 0) + d.cost; st[a.Status] = (st[a.Status] || 0) + 1; });
  const y0 = t.getFullYear() + (t.getMonth() >= 9 ? 1 : 0), dy = [];
  for (let y = y0 - 4; y <= y0; y++) { let v = 0; const e = new Date(y, 8, 30), b = new Date(y - 1, 8, 30); A.forEach(a => { v += depOf(a, e, pol).acc - depOf(a, b, pol).acc; }); dy.push([String(y + 543), Math.round(v * 100) / 100]); }
  const B = stockBal(), M = rows('Materials').filter(m => m.Status !== 'ยกเลิกใช้');
  let mv = 0, mq = 0; const low = [];
  M.forEach(m => { const b = B[m.MaterialID] || { q: 0, v: 0 }; mv += b.v; mq += b.q; if (b.q <= num(m.ReorderPoint)) low.push({ Code: m.Code, Name: m.Name, Unit: m.Unit, Qty: b.q, Reorder: num(m.ReorderPoint) }); });
  const L = rows('Loans').filter(l => l.Status === 'ยืมอยู่'), td = today();
  const over = L.filter(l => l.DueDate < td).map(l => ({ DocNo: l.DocNo, Borrower: l.Borrower, DueDate: l.DueDate, Days: Math.round((pD(td) - pD(l.DueDate)) / 864e5) }));
  const months = []; for (let i = 5; i >= 0; i--) { const d = new Date(); d.setMonth(d.getMonth() - i, 1); months.push(Utilities.formatDate(d, TZ, 'yyyy-MM')); }
  const out = months.map(() => 0), inn = months.map(() => 0);
  rows('StockTransactions').forEach(x => { const i = months.indexOf(String(x.TxDate).slice(0, 7)); if (i < 0) return; const v = num(x.Qty) * num(x.UnitCost); if (x.Type === 'OUT') out[i] += v; else inn[i] += v; });
  const rd = rows('AnnualInspections').filter(r => r.Status === 'เปิด').pop();
  const ir = {}; let pend = 0, rinfo = null;
  if (rd) { const R = rows('InspectionResults').filter(x => x.RoundID === rd.RoundID); R.forEach(x => ir[x.Result] = (ir[x.Result] || 0) + 1); pend = Math.max(A.length - R.length, 0); ir['ยังไม่ตรวจ'] = pend; rinfo = { name: rd.Name, end: rd.EndDate }; }
  return {
    k: { assets: A.length, cost: cost, acc: acc, book: cost - acc, matCount: M.length, matQty: mq, matValue: mv, openLoans: L.length, overdue: over.length, repair: rows('Maintenance').filter(m => m.Status === 'รออนุมัติ' || m.Status === 'กำลังซ่อม').length, inspPending: pend, low: low.length },
    ch: { cat: Object.entries(cat), dept: Object.entries(dep), status: Object.entries(st), dep: dy, months: months, out: out, inn: inn, insp: Object.entries(ir) },
    overdue: over, low: low, round: rinfo, sample: rows('Assets').some(a => a.CreatedBy === 'SEED')
  };
}
function buildAlerts() {
  const t = today(), out = [], B = stockBal();
  rows('Loans').filter(l => l.Status === 'ยืมอยู่' && l.DueDate).forEach(l => {
    if (l.DueDate < t) out.push({ type: 'loan', level: 'danger', ref: l.LoanID, date: l.DueDate, msg: 'ยืมเกินกำหนด ' + l.DocNo + ' ผู้ยืม ' + l.Borrower + ' ครบกำหนด ' + l.DueDate });
    else if (l.DueDate <= addDays(7)) out.push({ type: 'loandue', level: 'warning', ref: l.LoanID, date: l.DueDate, msg: 'ใกล้ครบกำหนดคืน ' + l.DocNo + ' ผู้ยืม ' + l.Borrower + ' วันที่ ' + l.DueDate });
  });
  rows('Materials').filter(m => m.Status !== 'ยกเลิกใช้').forEach(m => { const q = (B[m.MaterialID] || { q: 0 }).q; if (q <= num(m.ReorderPoint)) out.push({ type: 'stock', level: 'warning', ref: m.MaterialID, date: t, msg: 'วัสดุใกล้หมด ' + m.Name + ' คงเหลือ ' + q + ' ' + m.Unit + ' (จุดสั่งซื้อ ' + num(m.ReorderPoint) + ')' }); });
  const an = map('Assets', 'AssetID', 'Name');
  rows('Maintenance').filter(m => m.Status === 'รออนุมัติ' || m.Status === 'กำลังซ่อม').forEach(m => out.push({ type: 'repair', level: 'info', ref: m.MaintID, date: m.ReportDate, msg: 'ครุภัณฑ์รอซ่อม ' + (an[m.AssetID] || m.AssetID) + ' (' + m.Status + ') ' + m.Problem }));
  rows('AnnualInspections').filter(r => r.Status === 'เปิด' && r.EndDate).forEach(r => { if (r.EndDate <= addDays(14)) out.push({ type: 'round', level: r.EndDate < t ? 'danger' : 'warning', ref: r.RoundID, date: r.EndDate, msg: 'กำหนดสิ้นสุดตรวจนับ "' + r.Name + '" วันที่ ' + r.EndDate }); });
  return out.sort((a, b) => String(a.date).localeCompare(String(b.date)));
}

function alertsAct() {
  return buildAlerts();
}

function sendAlerts(p, u) {
  const r = notifyAlerts(true);
  audit(u, 'SEND_ALERTS', 'Notifications', today(), null, r);
  return r;
}

const OFF = ['admin', 'officer'], APR = ['admin', 'approver'], ADM = ['admin'];
const ACT = {
  boot: { f: boot }, dashboard: { f: dashboard }, alerts: { f: alertsAct }, list: { f: listEntity }, logout: { f: logout },
  assetGet: { f: assetGet }, assetSave: { f: assetSave, w: 1, r: OFF }, assetImport: { f: assetImport, w: 1, r: OFF },
  fileUpload: { f: fileUpload, w: 1, r: ['admin', 'officer', 'inspector'] }, fileData: { f: fileData },
  depCalc: { f: depCalc }, depSave: { f: depSave, w: 1, r: OFF },
  loanCreate: { f: loanCreate, w: 1 }, loanDecide: { f: loanDecide, w: 1, r: APR }, loanReturn: { f: loanReturn, w: 1, r: OFF },
  trCreate: { f: trCreate, w: 1, r: OFF }, trDecide: { f: trDecide, w: 1, r: APR }, trReceive: { f: trReceive, w: 1, r: OFF },
  matSave: { f: matSave, w: 1, r: OFF }, stockReceive: { f: stockReceive, w: 1, r: OFF },
  reqCreate: { f: reqCreate, w: 1 }, reqDecide: { f: reqDecide, w: 1, r: APR }, reqIssue: { f: reqIssue, w: 1, r: OFF },
  mtCreate: { f: mtCreate, w: 1 }, mtDecide: { f: mtDecide, w: 1, r: APR }, mtResult: { f: mtResult, w: 1, r: OFF },
  roundSave: { f: roundSave, w: 1, r: OFF }, assignSave: { f: assignSave, w: 1, r: OFF }, roundClose: { f: roundClose, w: 1, r: OFF },
  inspLookup: { f: inspLookup, r: ['admin', 'officer', 'inspector'] }, inspSave: { f: inspSave, w: 1, r: ['admin', 'officer', 'inspector'] },
  inspSummary: { f: inspSummary }, inspFollow: { f: inspFollow, w: 1, r: OFF },
  dpCreate: { f: dpCreate, w: 1, r: OFF }, dpStep: { f: dpStep, w: 1, r: ['admin', 'officer', 'approver'] },
  masterSave: { f: masterSave, w: 1, r: ADM }, userSave: { f: userSave, w: 1, r: ADM }, pwChange: { f: pwChange, w: 1 },
  settingsGet: { f: settingsGet }, settingsSave: { f: settingsSave, w: 1, r: ADM }, tgTest: { f: tgTest, r: ADM },
  triggersInstall: { f: triggersInstall, r: ADM }, backupNow: { f: backupNow, r: ADM }, sendAlerts: { f: sendAlerts, r: OFF },
  clearSample: { f: clearSample, w: 1, r: ADM }
};

function api(token, action, p) {
  // ลบ _M = {}; ออก
  try {
    p = p || {};
    if (action === 'login') return { ok: true, data: login(p) };
    const u = sess(token);
    if (!u) return { ok: false, code: 'AUTH', error: 'หมดเวลาการใช้งาน กรุณาเข้าสู่ระบบใหม่' };
    const h = ACT[action];
    if (!h) throw new Error('ไม่รู้จักคำสั่ง ' + action);
    if (h.r && h.r.indexOf(u.role) < 0) throw new Error('คุณไม่มีสิทธิ์ดำเนินการนี้');
    p._t = token;
    let lock = null;
    if (h.w) { 
      lock = LockService.getScriptLock(); 
      lock.waitLock(25000); 
      _M = {};   // เคลียร์ cache เฉพาะตอนเขียนข้อมูล
    }
    try { 
      const r = h.f(p, u); 
      return { ok: true, data: r === undefined ? true : r }; 
    }
    finally { if (lock) lock.releaseLock(); }
  } catch (e) { return { ok: false, error: String(e.message || e) }; }
}

function getCached(key, loader, sec) {
  const c = CacheService.getScriptCache();
  const v = c.get(key);
  if (v) return JSON.parse(v);
  const data = loader();
  try { c.put(key, JSON.stringify(data), sec || 300); } catch(e) {}
  return data;
}



function tg(text) {
  const tk = PropertiesService.getScriptProperties().getProperty('TG_TOKEN'), ch = getSet('TG_CHAT');
  if (!tk || !ch) return false;
  try {
    const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + tk + '/sendMessage', { method: 'post', contentType: 'application/json', payload: JSON.stringify({ chat_id: ch, text: String(text).slice(0, 4000) }), muteHttpExceptions: true });
    return r.getResponseCode() === 200;
  } catch (e) { return false; }
}
function notifyAlerts(force) {
  const a = buildAlerts(), t = today(), seen = {};
  rows('Notifications').forEach(n => seen[n.Key] = 1);
  const fresh = a.filter(x => force || !seen['D-' + t + '-' + x.type + '-' + x.ref]);
  if (!fresh.length) return { total: a.length, sent: 0, tg: false };
  const org = getSet('ORG_NAME') || 'ระบบครุภัณฑ์';
  const ok = tg('📌 ' + org + ' – แจ้งเตือนประจำวัน ' + t + '\n' + fresh.slice(0, 40).map(x => (x.level === 'danger' ? '🔴 ' : x.level === 'warning' ? '🟡 ' : '🔵 ') + x.msg).join('\n'));
  insMany('Notifications', fresh.map(x => ({ NotifID: uid('NT'), Type: x.type, Message: x.msg, RefID: x.ref, Key: 'D-' + t + '-' + x.type + '-' + x.ref + (force ? '-' + Date.now() : ''), Sent: ok ? 'Y' : 'N', SentAt: now() })));
  return { total: a.length, sent: fresh.length, tg: ok };
}
function dailyNotify() { notifyAlerts(false); }
function weeklyBackup() { backupNow({}, null); }
const tgSafe = t => { try { tg(t); } catch (e) { } };

/* ============================ Assets ============================ */
function assetSave(p, u) {
  const a = Object.assign({}, p.asset || {});
  [['RegNo', 'เลขทะเบียน'], ['Name', 'ชื่อครุภัณฑ์'], ['CategoryID', 'ประเภท'], ['DeptID', 'กองงาน']].forEach(k => { if (!String(a[k[0]] || '').trim()) throw new Error('กรุณากรอก ' + k[1]); });
  a.RegNo = String(a.RegNo).trim();
  if (num(a.Cost) < 0 || num(a.Qty) < 0) throw new Error('ราคาและจำนวนต้องไม่ติดลบ');
  const dup = rows('Assets').find(x => String(x.RegNo).toLowerCase() === a.RegNo.toLowerCase() && x.AssetID !== a.AssetID);
  if (dup) throw new Error('เลขทะเบียนซ้ำกับ "' + dup.Name + '"');
  delete a.Status; delete a.ImageId;
  if (a.AssetID) {
    const cur = must('Assets', 'AssetID', a.AssetID);
    if (cur.Status === ST.gone) throw new Error('ครุภัณฑ์ที่จำหน่ายแล้วแก้ไขไม่ได้');
    const r = upd('Assets', 'AssetID', a.AssetID, a, u);
    hist(a.AssetID, 'แก้ไขข้อมูล', diff(r.before, r.after), u);
    audit(u, 'UPDATE', 'Assets', a.AssetID, r.before, r.after);
    return a.AssetID;
  }
  a.AssetID = uid('AST'); a.Status = ST.ok; a.Qty = a.Qty || 1;
  ins('Assets', a, u); hist(a.AssetID, 'เพิ่มครุภัณฑ์', 'เลขทะเบียน ' + a.RegNo, u); audit(u, 'CREATE', 'Assets', a.AssetID, null, a);
  return a.AssetID;
}
function assetImport(p, u) {
  const ex = {}; rows('Assets').forEach(a => ex[String(a.RegNo).toLowerCase()] = 1);
  const good = [], errs = [], seen = {};
  (p.rows || []).forEach((a, i) => {
    const ln = i + 2;
    if (!a.RegNo || !a.Name || !a.CategoryID || !a.DeptID) { errs.push('แถว ' + ln + ': ข้อมูลจำเป็นไม่ครบ (เลขทะเบียน/ชื่อ/ประเภท/กองงาน)'); return; }
    const k = String(a.RegNo).trim().toLowerCase();
    if (ex[k] || seen[k]) { errs.push('แถว ' + ln + ': เลขทะเบียนซ้ำ ' + a.RegNo); return; }
    if (num(a.Cost) < 0) { errs.push('แถว ' + ln + ': ราคาติดลบ'); return; }
    seen[k] = 1; a.RegNo = String(a.RegNo).trim(); a.AssetID = uid('AST'); a.Status = ST.ok; a.Qty = a.Qty || 1; good.push(a);
  });
  insMany('Assets', good, u);
  insMany('AssetHistory', good.map(a => ({ HistID: uid('H'), AssetID: a.AssetID, Action: 'นำเข้าจาก Excel', Detail: a.RegNo, ActionAt: now(), ActionBy: u.name })), u);
  audit(u, 'IMPORT', 'Assets', today(), null, { ok: good.length, errors: errs.length });
  return { ok: good.length, errors: errs };
}
function assetGet(p) {
  const a = must('Assets', 'AssetID', p.id);
  return { asset: a, docs: rows('AssetDocuments').filter(d => d.AssetID === a.AssetID), hist: rows('AssetHistory').filter(h => h.AssetID === a.AssetID).reverse(), results: rows('InspectionResults').filter(r => r.AssetID === a.AssetID).reverse() };
}
function saveToFolder(sub, name, mime, b64) {
  if (!b64) throw new Error('ไม่พบข้อมูลไฟล์');
  if (b64.length > 14e6) throw new Error('ไฟล์ใหญ่เกิน 10 MB');
  const root = DriveApp.getFolderById(FOLDER_ID), nm = String(sub).replace(/[\\/:*?"<>|]/g, '_'), it = root.getFoldersByName(nm), fo = it.hasNext() ? it.next() : root.createFolder(nm);
  return fo.createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime || 'application/octet-stream', name || ('file-' + Date.now())));
}
function fileUpload(p, u) {
  const a = must('Assets', 'AssetID', p.assetId);
  const f = saveToFolder('ASSET-' + a.RegNo, p.name, p.mime, p.data);
  const d = ins('AssetDocuments', { DocID: uid('DOC'), AssetID: a.AssetID, DocType: p.docType || 'อื่นๆ', FileName: p.name, FileUrl: f.getUrl(), FileId: f.getId() }, u);
  if (p.docType === 'ภาพถ่าย' && !a.ImageId && /^image\//.test(p.mime || '')) upd('Assets', 'AssetID', a.AssetID, { ImageId: f.getId() }, u);
  hist(a.AssetID, 'แนบเอกสาร', (p.docType || '') + ': ' + p.name, u);
  audit(u, 'UPLOAD', 'AssetDocuments', d.DocID, null, { asset: a.RegNo, file: p.name, type: p.docType });
  return d;
}
function fileData(p) {
  const id = String(p.id || '');
  const ok = rows('AssetDocuments').some(d => d.FileId === id) || rows('Assets').some(a => a.ImageId === id) || rows('InspectionResults').some(r => r.PhotoId === id);
  if (!ok) throw new Error('ไม่พบไฟล์ในทะเบียน');
  const f = DriveApp.getFileById(id);
  if (f.getSize() > 4 * 1024 * 1024) throw new Error('ไฟล์ใหญ่เกินกว่าจะแสดงตัวอย่าง กรุณาเปิดจาก Drive');
  return { mime: f.getMimeType(), data: Utilities.base64Encode(f.getBlob().getBytes()) };
}

/* ============================ Loans / Transfers ============================ */
function openAssetSet() {
  const open = {}, L = {}; rows('Loans').forEach(l => L[l.LoanID] = l);
  rows('LoanDetails').forEach(d => { const l = L[d.LoanID]; if (l && (l.Status === 'รออนุมัติ' || l.Status === 'ยืมอยู่')) open[d.AssetID] = 'ยื่นยืมอยู่'; });
  rows('Transfers').forEach(t => { if (t.Status === 'รออนุมัติ' || t.Status === 'รอรับโอน') open[t.AssetID] = 'อยู่ระหว่างโอนย้าย'; });
  return open;
}
function loanCreate(p, u) {
  const ids = p.assetIds || [];
  if (!ids.length) throw new Error('กรุณาเลือกครุภัณฑ์ที่ต้องการยืม');
  if (!p.borrower || !p.dueDate) throw new Error('กรุณากรอกชื่อผู้ยืมและวันครบกำหนดคืน');
  const ld = p.loanDate || today();
  if (p.dueDate < ld) throw new Error('วันครบกำหนดต้องไม่ก่อนวันยืม');
  const open = openAssetSet();
  ids.forEach(id => { const a = must('Assets', 'AssetID', id); if (a.Status !== ST.ok) throw new Error(a.RegNo + ' ไม่ว่าง (' + a.Status + ')'); if (open[id]) throw new Error(a.RegNo + ' ' + open[id]); });
  const loan = ins('Loans', { LoanID: uid('LN'), DocNo: nextNo('LN', 'ใบยืมพัสดุ'), Borrower: p.borrower, DeptID: p.deptId || '', Purpose: p.purpose || '', LoanDate: ld, DueDate: p.dueDate, Status: 'รออนุมัติ', RequestedBy: u.name }, u);
  insMany('LoanDetails', ids.map(id => ({ DetailID: uid('LD'), LoanID: loan.LoanID, AssetID: id, ConditionOut: p.conditionOut || 'ปกติ' })), u);
  audit(u, 'LOAN_REQUEST', 'Loans', loan.LoanID, null, loan);
  return loan.LoanID;
}
function loanDecide(p, u) {
  const l = must('Loans', 'LoanID', p.id);
  if (l.Status !== 'รออนุมัติ') throw new Error('รายการนี้ไม่อยู่ในสถานะรออนุมัติ');
  const det = rows('LoanDetails').filter(d => d.LoanID === l.LoanID);
  if (p.approve) {
    det.forEach(d => { const a = must('Assets', 'AssetID', d.AssetID); if (a.Status !== ST.ok) throw new Error(a.RegNo + ' ไม่ว่างแล้ว (' + a.Status + ')'); });
    det.forEach(d => { upd('Assets', 'AssetID', d.AssetID, { Status: ST.loan }, u); hist(d.AssetID, 'อนุมัติการยืม', l.DocNo + ' ผู้ยืม ' + l.Borrower, u); });
  }
  const r = upd('Loans', 'LoanID', l.LoanID, { Status: p.approve ? 'ยืมอยู่' : 'ไม่อนุมัติ', ApprovedBy: u.name, ApprovedAt: now(), DecisionNote: p.note || '' }, u);
  audit(u, p.approve ? 'LOAN_APPROVE' : 'LOAN_REJECT', 'Loans', l.LoanID, r.before, r.after);
  return true;
}
function loanReturn(p, u) {
  const l = must('Loans', 'LoanID', p.id);
  if (l.Status !== 'ยืมอยู่') throw new Error('รายการนี้ไม่อยู่ในสถานะยืมอยู่');
  const det = rows('LoanDetails').filter(d => d.LoanID === l.LoanID), cond = {};
  (p.items || []).forEach(i => cond[i.assetId] = i.condition || 'ปกติ');
  det.forEach(d => {
    const c = cond[d.AssetID] || 'ปกติ';
    upd('LoanDetails', 'DetailID', d.DetailID, { ConditionIn: c }, u);
    upd('Assets', 'AssetID', d.AssetID, { Status: c === 'ชำรุด' ? ST.bad : ST.ok }, u);
    hist(d.AssetID, 'รับคืนครุภัณฑ์', l.DocNo + ' สภาพ: ' + c, u);
  });
  const r = upd('Loans', 'LoanID', l.LoanID, { Status: 'คืนแล้ว', ReturnDate: today(), ReturnNote: p.note || '' }, u);
  audit(u, 'LOAN_RETURN', 'Loans', l.LoanID, r.before, r.after);
  return true;
}
function trCreate(p, u) {
  const a = must('Assets', 'AssetID', p.assetId);
  if (a.Status !== ST.ok) throw new Error('โอนย้ายได้เฉพาะครุภัณฑ์สถานะ "ใช้งานได้" (ปัจจุบัน: ' + a.Status + ')');
  if (openAssetSet()[a.AssetID]) throw new Error('ครุภัณฑ์นี้มีรายการยืมหรือโอนย้ายค้างอยู่');
  if (!p.toDeptId) throw new Error('กรุณาเลือกกองงานปลายทาง');
  const t = ins('Transfers', { TransferID: uid('TR'), DocNo: nextNo('TR', 'ใบโอนย้าย'), AssetID: a.AssetID, FromDeptID: a.DeptID, ToDeptID: p.toDeptId, ToLocationID: p.toLocationId || '', ToCustodian: p.toCustodian || '', Reason: p.reason || '', Status: 'รออนุมัติ', RequestedBy: u.name }, u);
  hist(a.AssetID, 'ขอโอนย้าย', t.DocNo, u); audit(u, 'TRANSFER_REQUEST', 'Transfers', t.TransferID, null, t);
  return t.TransferID;
}
function trDecide(p, u) {
  const t = must('Transfers', 'TransferID', p.id);
  if (t.Status !== 'รออนุมัติ') throw new Error('สถานะไม่ถูกต้อง');
  const r = upd('Transfers', 'TransferID', t.TransferID, { Status: p.approve ? 'รอรับโอน' : 'ไม่อนุมัติ', ApprovedBy: u.name, ApprovedAt: now() }, u);
  hist(t.AssetID, p.approve ? 'อนุมัติโอนย้าย' : 'ไม่อนุมัติโอนย้าย', t.DocNo, u); audit(u, p.approve ? 'TRANSFER_APPROVE' : 'TRANSFER_REJECT', 'Transfers', t.TransferID, r.before, r.after);
  return true;
}
function trReceive(p, u) {
  const t = must('Transfers', 'TransferID', p.id);
  if (t.Status !== 'รอรับโอน') throw new Error('สถานะไม่ถูกต้อง');
  const a = must('Assets', 'AssetID', t.AssetID);
  if (a.Status !== ST.ok) throw new Error('ครุภัณฑ์ไม่อยู่ในสถานะที่โอนได้ (' + a.Status + ')');
  const ra = upd('Assets', 'AssetID', a.AssetID, { DeptID: t.ToDeptID, LocationID: t.ToLocationID || a.LocationID, Custodian: t.ToCustodian || a.Custodian }, u);
  const r = upd('Transfers', 'TransferID', t.TransferID, { Status: 'โอนแล้ว', ReceivedBy: u.name, ReceivedAt: now() }, u);
  hist(a.AssetID, 'รับโอนครุภัณฑ์', t.DocNo + ' | ' + diff(ra.before, ra.after), u);
  audit(u, 'TRANSFER_RECEIVE', 'Transfers', t.TransferID, ra.before, ra.after);
  return true;
}

/* ============================ Materials / FIFO ============================ */
function matSave(p, u) {
  const m = Object.assign({}, p.mat || {});
  if (!m.Code || !m.Name || !m.Unit) throw new Error('กรุณากรอก รหัส ชื่อ และหน่วยนับ');
  if (num(m.ReorderPoint) < 0) throw new Error('จุดสั่งซื้อต้องไม่ติดลบ');
  const dup = rows('Materials').find(x => String(x.Code).toLowerCase() === String(m.Code).toLowerCase() && x.MaterialID !== m.MaterialID);
  if (dup) throw new Error('รหัสวัสดุซ้ำกับ ' + dup.Name);
  if (m.MaterialID) { const r = upd('Materials', 'MaterialID', m.MaterialID, m, u); audit(u, 'UPDATE', 'Materials', m.MaterialID, r.before, r.after); return m.MaterialID; }
  m.MaterialID = uid('MAT'); m.Status = 'ใช้งาน'; ins('Materials', m, u); audit(u, 'CREATE', 'Materials', m.MaterialID, null, m);
  return m.MaterialID;
}
function stockReceive(p, u) {
  const m = must('Materials', 'MaterialID', p.materialId), q = num(p.qty), c = num(p.unitCost);
  if (q <= 0) throw new Error('จำนวนรับต้องมากกว่า 0');
  if (c < 0) throw new Error('ราคาต่อหน่วยต้องไม่ติดลบ');
  const doc = p.docNo || nextNo('RC', 'ใบรับวัสดุ');
  const lot = ins('StockLots', { LotID: uid('LOT'), MaterialID: m.MaterialID, ReceiptDocNo: doc, ReceivedDate: p.date || today(), QtyIn: q, QtyRemain: q, UnitCost: c, Supplier: p.supplier || '' }, u);
  ins('StockTransactions', { TxID: uid('TX'), MaterialID: m.MaterialID, LotID: lot.LotID, Type: 'IN', Qty: q, UnitCost: c, RefDoc: doc, TxDate: p.date || today(), ByUser: u.name, DeptID: '' }, u);
  audit(u, 'STOCK_IN', 'StockLots', lot.LotID, null, lot);
  return lot.LotID;
}
function reqCreate(p, u) {
  const items = {};
  (p.items || []).forEach(i => { if (num(i.qty) > 0) items[i.materialId] = (items[i.materialId] || 0) + num(i.qty); });
  if (!Object.keys(items).length) throw new Error('กรุณาเพิ่มรายการวัสดุ');
  if (!p.deptId || !p.purpose) throw new Error('กรุณาเลือกกองงานและระบุวัตถุประสงค์');
  Object.keys(items).forEach(id => must('Materials', 'MaterialID', id, 'ไม่พบวัสดุ'));
  const rq = ins('Requisitions', { ReqID: uid('RQ'), DocNo: nextNo('RQ', 'ใบเบิกวัสดุ'), DeptID: p.deptId, RequestedBy: u.name, Purpose: p.purpose, ReqDate: today(), Status: 'รออนุมัติ' }, u);
  insMany('RequisitionDetails', Object.keys(items).map(id => ({ DetailID: uid('RD'), ReqID: rq.ReqID, MaterialID: id, Qty: items[id], IssuedQty: 0 })), u);
  audit(u, 'REQ_CREATE', 'Requisitions', rq.ReqID, null, { doc: rq.DocNo, items: items });
  return rq.ReqID;
}
function reqDecide(p, u) {
  const q = must('Requisitions', 'ReqID', p.id);
  if (q.Status !== 'รออนุมัติ') throw new Error('สถานะไม่ถูกต้อง');
  const r = upd('Requisitions', 'ReqID', q.ReqID, { Status: p.approve ? 'อนุมัติแล้ว' : 'ไม่อนุมัติ', ApprovedBy: u.name, ApprovedAt: now(), DecisionNote: p.note || '' }, u);
  audit(u, p.approve ? 'REQ_APPROVE' : 'REQ_REJECT', 'Requisitions', q.ReqID, r.before, r.after);
  return true;
}
function reqIssue(p, u) {
  const q = must('Requisitions', 'ReqID', p.id);
  if (q.Status !== 'อนุมัติแล้ว') throw new Error('ต้องอนุมัติก่อนจ่ายวัสดุ');
  const det = rows('RequisitionDetails').filter(d => d.ReqID === q.ReqID), mn = map('Materials', 'MaterialID', 'Name');
  const lots = rows('StockLots').filter(l => num(l.QtyRemain) > 0).sort((a, b) => (String(a.ReceivedDate) + String(a.CreatedAt) + a.LotID).localeCompare(String(b.ReceivedDate) + String(b.CreatedAt) + b.LotID));
  const rem = {}, plan = [], touched = {}; lots.forEach(l => rem[l.LotID] = num(l.QtyRemain));
  det.forEach(d => {
    let need = num(d.Qty);
    lots.filter(l => l.MaterialID === d.MaterialID).forEach(l => { if (need <= 0) return; const t = Math.min(need, rem[l.LotID]); if (t > 0) { rem[l.LotID] -= t; need -= t; touched[l.LotID] = 1; plan.push({ d: d, l: l, t: t }); } });
    if (need > 0) throw new Error('วัสดุไม่เพียงพอ: ' + (mn[d.MaterialID] || d.MaterialID) + ' ขาด ' + need + ' (ไม่อนุญาตให้ยอดติดลบ)');
  });
  Object.keys(touched).forEach(id => upd('StockLots', 'LotID', id, { QtyRemain: rem[id] }, u));
  insMany('StockTransactions', plan.map(x => ({ TxID: uid('TX'), MaterialID: x.d.MaterialID, LotID: x.l.LotID, Type: 'OUT', Qty: x.t, UnitCost: x.l.UnitCost, RefDoc: q.DocNo, TxDate: today(), ByUser: u.name, DeptID: q.DeptID })), u);
  det.forEach(d => upd('RequisitionDetails', 'DetailID', d.DetailID, { IssuedQty: d.Qty }, u));
  const r = upd('Requisitions', 'ReqID', q.ReqID, { Status: 'จ่ายแล้ว', IssuedBy: u.name, IssuedAt: now() }, u);
  audit(u, 'REQ_ISSUE', 'Requisitions', q.ReqID, r.before, { status: 'จ่ายแล้ว', fifo: plan.map(x => ({ lot: x.l.LotID, doc: x.l.ReceiptDocNo, qty: x.t })) });
  return plan.length;
}

/* ============================ Maintenance ============================ */
const mh = (id, st, note, u) => ins('MaintenanceHistory', { HistID: uid('MH'), MaintID: id, Status: st, Note: note || '', At: now(), By: u.name }, u);
function mtCreate(p, u) {
  const a = must('Assets', 'AssetID', p.assetId);
  if ([ST.gone, ST.pend, ST.rep].indexOf(a.Status) >= 0) throw new Error('ครุภัณฑ์อยู่ในสถานะ "' + a.Status + '" แจ้งซ่อมไม่ได้');
  if (a.Status === ST.loan) throw new Error('ครุภัณฑ์ถูกยืมอยู่ กรุณารับคืนก่อน');
  if (!p.problem) throw new Error('กรุณาระบุอาการ/ปัญหา');
  const m = ins('Maintenance', { MaintID: uid('MT'), DocNo: nextNo('MT', 'ใบแจ้งซ่อม'), AssetID: a.AssetID, Problem: p.problem, ReportedBy: u.name, ReportDate: today(), Status: 'รออนุมัติ' }, u);
  upd('Assets', 'AssetID', a.AssetID, { Status: ST.rep }, u);
  mh(m.MaintID, 'รออนุมัติ', p.problem, u); hist(a.AssetID, 'แจ้งซ่อม', m.DocNo + ': ' + p.problem, u); audit(u, 'MAINT_REQUEST', 'Maintenance', m.MaintID, null, m);
  tgSafe('🔧 แจ้งซ่อมใหม่ ' + m.DocNo + '\n' + a.Name + ' (' + a.RegNo + ')\nอาการ: ' + p.problem);
  return m.MaintID;
}
function mtDecide(p, u) {
  const m = must('Maintenance', 'MaintID', p.id);
  if (m.Status !== 'รออนุมัติ') throw new Error('สถานะไม่ถูกต้อง');
  const r = upd('Maintenance', 'MaintID', m.MaintID, { Status: p.approve ? 'กำลังซ่อม' : 'ไม่อนุมัติ', ApprovedBy: u.name, ApprovedAt: now() }, u);
  if (!p.approve) upd('Assets', 'AssetID', m.AssetID, { Status: ST.ok }, u);
  mh(m.MaintID, r.after.Status, p.note, u); hist(m.AssetID, p.approve ? 'อนุมัติซ่อม' : 'ไม่อนุมัติซ่อม', m.DocNo, u);
  audit(u, p.approve ? 'MAINT_APPROVE' : 'MAINT_REJECT', 'Maintenance', m.MaintID, r.before, r.after);
  return true;
}
function mtResult(p, u) {
  const m = must('Maintenance', 'MaintID', p.id);
  if (m.Status !== 'กำลังซ่อม') throw new Error('ต้องอนุมัติและอยู่ในสถานะ "กำลังซ่อม" ก่อนบันทึกผล');
  if (num(p.cost) < 0) throw new Error('ค่าซ่อมต้องไม่ติดลบ');
  const ok = !!p.repaired;
  const r = upd('Maintenance', 'MaintID', m.MaintID, { Status: ok ? 'ซ่อมเสร็จ' : 'ซ่อมไม่ได้', ResultDate: p.date || today(), ResultNote: p.note || '', Cost: num(p.cost), Technician: p.technician || '' }, u);
  upd('Assets', 'AssetID', m.AssetID, { Status: ok ? ST.ok : ST.bad }, u);
  mh(m.MaintID, r.after.Status, p.note, u); hist(m.AssetID, 'บันทึกผลซ่อม', m.DocNo + ' ' + r.after.Status + ' ค่าซ่อม ' + num(p.cost), u);
  audit(u, 'MAINT_RESULT', 'Maintenance', m.MaintID, r.before, r.after);
  return true;
}

/* ============================ Annual inspection ============================ */
function roundSave(p, u) {
  const r = Object.assign({}, p.round || {});
  if (!r.Name || !r.StartDate || !r.EndDate) throw new Error('กรุณากรอกชื่อรอบ วันเริ่ม และวันสิ้นสุด');
  if (r.EndDate < r.StartDate) throw new Error('วันสิ้นสุดต้องไม่ก่อนวันเริ่ม');
  if (r.RoundID) { const x = upd('AnnualInspections', 'RoundID', r.RoundID, r, u); audit(u, 'UPDATE', 'AnnualInspections', r.RoundID, x.before, x.after); return r.RoundID; }
  r.RoundID = uid('RD'); r.Status = 'เปิด'; r.FiscalYear = r.FiscalYear || fyOf(r.StartDate); ins('AnnualInspections', r, u); audit(u, 'CREATE', 'AnnualInspections', r.RoundID, null, r);
  return r.RoundID;
}
function assignSave(p, u) {
  must('AnnualInspections', 'RoundID', p.roundId); must('Users', 'UserID', p.userId, 'ไม่พบผู้ใช้');
  if (rows('InspectionAssignments').some(a => a.RoundID === p.roundId && a.InspectorUser === p.userId && String(a.DeptID) === String(p.deptId || ''))) throw new Error('มอบหมายซ้ำ');
  const a = ins('InspectionAssignments', { AssignID: uid('AS'), RoundID: p.roundId, InspectorUser: p.userId, DeptID: p.deptId || '' }, u);
  audit(u, 'ASSIGN', 'InspectionAssignments', a.AssignID, null, a);
  return a.AssignID;
}
function roundClose(p, u) {
  const r = upd('AnnualInspections', 'RoundID', p.id, { Status: 'ปิด' }, u);
  audit(u, 'ROUND_CLOSE', 'AnnualInspections', p.id, r.before, r.after);
  return true;
}
function inspLookup(p, u) {
  const rd = must('AnnualInspections', 'RoundID', p.roundId);
  if (rd.Status !== 'เปิด') throw new Error('รอบตรวจนับนี้ปิดแล้ว');
  let c = String(p.code || '').trim(); const m = c.match(/[?&]a=([^&#]+)/); if (m) c = decodeURIComponent(m[1]);
  const lc = c.toLowerCase();
  const a = rows('Assets').find(x => String(x.RegNo).toLowerCase() === lc || String(x.AssetID).toLowerCase() === lc);
  if (!a) throw new Error('ไม่พบครุภัณฑ์รหัส "' + c + '" ในทะเบียน');
  return { asset: a, result: rows('InspectionResults').find(r => r.RoundID === rd.RoundID && r.AssetID === a.AssetID) || null, conflict: String(a.Custodian).trim() === String(u.name).trim() };
}
function inspSave(p, u) {
  const rd = must('AnnualInspections', 'RoundID', p.roundId);
  if (rd.Status !== 'เปิด') throw new Error('รอบตรวจนับนี้ปิดแล้ว');
  const a = must('Assets', 'AssetID', p.assetId);
  if (['พบ', 'ไม่พบ', 'ชำรุด', 'ย้ายสถานที่'].indexOf(p.result) < 0) throw new Error('ผลการตรวจไม่ถูกต้อง');
  if (String(a.Custodian).trim() && String(a.Custodian).trim() === String(u.name).trim()) throw new Error('ผู้ตรวจนับต้องไม่ใช่ผู้รับผิดชอบครุภัณฑ์รายการนี้');
  if (u.role === 'inspector') {
    const as = rows('InspectionAssignments').filter(x => x.RoundID === rd.RoundID);
    if (as.length) { const mine = as.filter(x => x.InspectorUser === u.id); if (!mine.length || !mine.some(x => !x.DeptID || x.DeptID === a.DeptID)) throw new Error('ครุภัณฑ์นี้ไม่อยู่ในกองงานที่ได้รับมอบหมาย'); }
  }
  let photo = '';
  if (p.photo && p.photo.data) photo = saveToFolder('INSPECTION-' + rd.RoundID, a.RegNo + '-' + Date.now() + '.jpg', p.photo.mime || 'image/jpeg', p.photo.data).getId();
  const rec = { Result: p.result, FoundLocationID: p.foundLocationId || '', Note: p.note || '', Lat: p.lat || '', Lng: p.lng || '', InspectedBy: u.name, InspectedAt: now() };
  if (photo) rec.PhotoId = photo;
  const ex = rows('InspectionResults').find(r => r.RoundID === rd.RoundID && r.AssetID === a.AssetID);
  if (ex) { const r = upd('InspectionResults', 'ResultID', ex.ResultID, rec, u); audit(u, 'INSPECT_UPDATE', 'InspectionResults', ex.ResultID, r.before, r.after); }
  else { const r = ins('InspectionResults', Object.assign({ ResultID: uid('IR'), RoundID: rd.RoundID, AssetID: a.AssetID }, rec), u); audit(u, 'INSPECT', 'InspectionResults', r.ResultID, null, r); }
  hist(a.AssetID, 'ตรวจนับประจำปี', rd.Name + ': ' + p.result, u);
  return true;
}
function inspSummary(p) {
  const R = {}; rows('InspectionResults').filter(r => r.RoundID === p.roundId).forEach(r => R[r.AssetID] = r);
  return rows('Assets').filter(a => a.Status !== ST.gone).map(a => { const r = R[a.AssetID]; return { AssetID: a.AssetID, RegNo: a.RegNo, Name: a.Name, DeptID: a.DeptID, LocationID: a.LocationID, Custodian: a.Custodian, Result: r ? r.Result : 'ยังไม่ตรวจ', ResultID: r ? r.ResultID : '', FoundLocationID: r ? r.FoundLocationID : '', Note: r ? r.Note : '', By: r ? r.InspectedBy : '', At: r ? r.InspectedAt : '', PhotoId: r ? r.PhotoId : '', FollowStatus: r ? r.FollowStatus : '', FollowNote: r ? r.FollowNote : '' }; });
}
function inspFollow(p, u) {
  const r = upd('InspectionResults', 'ResultID', p.id, { FollowStatus: p.status, FollowNote: p.note || '' }, u);
  audit(u, 'INSPECT_FOLLOW', 'InspectionResults', p.id, r.before, r.after);
  return true;
}

/* ============================ Disposal ============================ */
function dpCreate(p, u) {
  const a = must('Assets', 'AssetID', p.assetId);
  if ([ST.gone, ST.pend, ST.loan].indexOf(a.Status) >= 0) throw new Error('ครุภัณฑ์อยู่ในสถานะ "' + a.Status + '" เสนอจำหน่ายไม่ได้');
  if (!p.reason) throw new Error('กรุณาระบุเหตุผลการจำหน่าย');
  const d = ins('Disposals', { DisposalID: uid('DP'), DocNo: nextNo('DP', 'รายงานการจำหน่าย'), AssetID: a.AssetID, Reason: p.reason, Status: 'เสนอรายงาน', ReportDate: p.date || today() }, u);
  upd('Assets', 'AssetID', a.AssetID, { Status: ST.pend }, u);
  hist(a.AssetID, 'เสนอจำหน่าย', d.DocNo + ': ' + p.reason, u); audit(u, 'DISPOSAL_REQUEST', 'Disposals', d.DisposalID, null, d);
  return d.DisposalID;
}
function dpStep(p, u) {
  const d = must('Disposals', 'DisposalID', p.id), s = p.step;
  const need = (st, roles) => { if (d.Status !== st) throw new Error('ขั้นตอนไม่ถูกต้อง (ปัจจุบัน: ' + d.Status + ')'); if (roles.indexOf(u.role) < 0) throw new Error('คุณไม่มีสิทธิ์ในขั้นตอนนี้'); };
  let patch, act;
  if (s === 'order') { need('เสนอรายงาน', OFF); if (!p.text) throw new Error('กรุณาระบุเลขที่/รายละเอียดคำสั่งแต่งตั้งคณะกรรมการ'); patch = { Status: 'ตั้งคณะกรรมการ', CommitteeOrder: p.text }; act = 'บันทึกคำสั่งคณะกรรมการ'; }
  else if (s === 'finding') { need('ตั้งคณะกรรมการ', OFF); if (!p.text) throw new Error('กรุณาบันทึกผลสอบข้อเท็จจริง'); patch = { Status: 'สอบข้อเท็จจริงแล้ว', FindingNote: p.text }; act = 'บันทึกผลสอบข้อเท็จจริง'; }
  else if (s === 'approve') { need('สอบข้อเท็จจริงแล้ว', APR); patch = { Status: 'อนุมัติแล้ว', ApprovedBy: u.name, ApprovedAt: now() }; act = 'อนุมัติจำหน่าย'; }
  else if (s === 'complete') { need('อนุมัติแล้ว', OFF); if (!p.method || !p.text) throw new Error('กรุณาระบุวิธีจำหน่ายและหลักฐาน'); patch = { Status: 'จำหน่ายแล้ว', Method: p.method, Evidence: p.text, CompletedAt: now() }; act = 'จำหน่ายแล้ว'; }
  else if (s === 'cancel') { if (['admin', 'officer', 'approver'].indexOf(u.role) < 0) throw new Error('ไม่มีสิทธิ์'); if (d.Status === 'จำหน่ายแล้ว' || d.Status === 'ยกเลิก') throw new Error('ยกเลิกไม่ได้'); patch = { Status: 'ยกเลิก' }; act = 'ยกเลิกการจำหน่าย'; }
  else throw new Error('ไม่รู้จักขั้นตอน');
  const r = upd('Disposals', 'DisposalID', d.DisposalID, patch, u);
  if (s === 'complete') upd('Assets', 'AssetID', d.AssetID, { Status: ST.gone }, u);
  if (s === 'cancel') upd('Assets', 'AssetID', d.AssetID, { Status: ST.ok }, u);
  hist(d.AssetID, act, d.DocNo, u); audit(u, 'DISPOSAL_' + s.toUpperCase(), 'Disposals', d.DisposalID, r.before, r.after);
  return true;
}

/* ============================ Admin ============================ */
const MASTER = { depts: ['Departments', 'DeptID', 'DEP'], locs: ['Locations', 'LocationID', 'LOC'], cats: ['AssetCategories', 'CategoryID', 'CAT'], policies: ['DepreciationPolicies', 'PolicyID', 'POL'], fys: ['FiscalYears', 'FiscalYear', ''] };
function masterSave(p, u) {
  const d = MASTER[p.e]; if (!d) throw new Error('ไม่รู้จักข้อมูล');
  const r = Object.assign({}, p.row || {});
  if (p.e === 'depts' && !r.DeptName) throw new Error('กรุณากรอกชื่อกองงาน');
  if (p.e === 'locs' && !r.LocationName) throw new Error('กรุณากรอกชื่อสถานที่');
  if (p.e === 'cats' && !r.CategoryName) throw new Error('กรุณากรอกชื่อประเภท');
  if (p.e === 'policies') { if (!r.CategoryID || num(r.LifeYears) <= 0) throw new Error('กรุณาเลือกประเภทและกำหนดอายุการใช้งาน (ปี)'); if (r.SalvageValue === '' || r.SalvageValue === undefined) throw new Error('กรุณากำหนดมูลค่าคงเหลือ (ระบบไม่กำหนดให้อัตโนมัติ)'); r.Method = r.Method || 'เส้นตรง'; }
  if (p.e === 'fys' && !r.FiscalYear) throw new Error('กรุณากรอกปีงบประมาณ');
  r.Status = r.Status || 'ใช้งาน';
  const exists = r[d[1]] && find(d[0], d[1], r[d[1]]);
  if (exists) { const x = upd(d[0], d[1], r[d[1]], r, u); audit(u, 'UPDATE', d[0], r[d[1]], x.before, x.after); }
  else { if (d[2]) r[d[1]] = uid(d[2]); ins(d[0], r, u); audit(u, 'CREATE', d[0], r[d[1]], null, r); }
  return r[d[1]];
}
function userSave(p, u) {
  const x = Object.assign({}, p.user || {}), pw = x.password; delete x.password;
  if (!x.Username || !x.FullName || !x.Role) throw new Error('กรุณากรอกชื่อผู้ใช้ ชื่อ-สกุล และสิทธิ์');
  if (['admin', 'officer', 'approver', 'inspector', 'staff'].indexOf(x.Role) < 0) throw new Error('สิทธิ์ไม่ถูกต้อง');
  if (rows('Users').some(r => String(r.Username).toLowerCase() === String(x.Username).toLowerCase() && r.UserID !== x.UserID)) throw new Error('ชื่อผู้ใช้ซ้ำ');
  if (pw) { if (String(pw).length < 8) throw new Error('รหัสผ่านต้องยาวอย่างน้อย 8 ตัวอักษร'); x.Salt = Utilities.getUuid(); x.PasswordHash = hashPw(pw, x.Salt); }
  x.Status = x.Status || 'ใช้งาน';
  if (x.UserID) {
    const before = must('Users', 'UserID', x.UserID);
    if (before.Role === 'admin' && (x.Role !== 'admin' || x.Status !== 'ใช้งาน') && rows('Users').filter(r => r.Role === 'admin' && r.Status === 'ใช้งาน').length < 2) throw new Error('ต้องมีผู้ดูแลระบบที่ใช้งานอย่างน้อย 1 คน');
    const r = upd('Users', 'UserID', x.UserID, x, u);
    audit(u, 'UPDATE', 'Users', x.UserID, { Role: r.before.Role, Status: r.before.Status, FullName: r.before.FullName }, { Role: r.after.Role, Status: r.after.Status, FullName: r.after.FullName, pwChanged: !!pw });
    return x.UserID;
  }
  if (!pw) throw new Error('กรุณากำหนดรหัสผ่านเริ่มต้น');
  x.UserID = uid('USR');
  ins('Users', x, u);
  audit(u, 'CREATE', 'Users', x.UserID, null, { Username: x.Username, Role: x.Role });
  return x.UserID;
}
function settingsGet() {
  return { ORG_NAME: getSet('ORG_NAME'), TG_CHAT: getSet('TG_CHAT'), tgTokenSet: !!PropertiesService.getScriptProperties().getProperty('TG_TOKEN'), triggers: ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction()) };
}
function settingsSave(p, u) {
  if (p.ORG_NAME !== undefined) setSet('ORG_NAME', p.ORG_NAME, u);
  if (p.TG_CHAT !== undefined) setSet('TG_CHAT', p.TG_CHAT, u);
  if (p.TG_TOKEN) PropertiesService.getScriptProperties().setProperty('TG_TOKEN', p.TG_TOKEN.trim());
  audit(u, 'SETTINGS', 'Settings', 'general', null, { org: p.ORG_NAME, chat: p.TG_CHAT, tokenChanged: !!p.TG_TOKEN });
  return true;
}
function tgTest() { if (!tg('✅ ทดสอบการแจ้งเตือนจากระบบครุภัณฑ์และพัสดุ ' + now())) throw new Error('ส่ง Telegram ไม่สำเร็จ ตรวจสอบ Bot Token และ Chat ID (ต้องกด Start กับบอทก่อน)'); return true; }
function triggersInstall(p, u) {
  ScriptApp.getProjectTriggers().forEach(t => { if (['dailyNotify', 'weeklyBackup'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyNotify').timeBased().everyDays(1).atHour(8).inTimezone(TZ).create();
  ScriptApp.newTrigger('weeklyBackup').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(2).inTimezone(TZ).create();
  audit(u, 'TRIGGERS', 'Settings', 'triggers');
  return true;
}
function backupNow(p, u) {
  const f = DriveApp.getFileById(SHEET_ID).makeCopy('BACKUP-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmm'), DriveApp.getFolderById(FOLDER_ID));
  audit(u, 'BACKUP', 'Spreadsheet', f.getId());
  return { url: f.getUrl() };
}
function clearSample(p, u) {
  let n = 0;
  Object.keys(SCHEMA).forEach(name => {
    if (name === 'AuditLogs') return;
    const s = sh(name), v = s.getDataRange().getValues(); if (v.length < 2) return;
    const ci = v[0].indexOf('CreatedBy');
    for (let i = v.length - 1; i >= 1; i--) if (v[i][ci] === 'SEED') { s.deleteRow(i + 1); n++; }
  });
  _M = {};
  audit(u, 'CLEAR_SAMPLE', 'ALL', 'SEED', null, { deleted: n });
  return { deleted: n };
}

/* ============================ Setup & sample data ============================ */
function setupSystem() {
  Object.keys(SCHEMA).forEach(n => sh(n));
  if (!rows('Roles').length) insMany('Roles', [
    { Role: 'admin', RoleName: 'ผู้ดูแลระบบ', Description: 'จัดการผู้ใช้ สิทธิ์ ตั้งค่า และข้อมูลทั้งหมด' },
    { Role: 'officer', RoleName: 'เจ้าหน้าที่พัสดุ', Description: 'บันทึกทะเบียน รับ-จ่าย โอนย้าย ซ่อม จำหน่าย' },
    { Role: 'approver', RoleName: 'ผู้อนุมัติ', Description: 'อนุมัติยืม โอน เบิก ซ่อม จำหน่าย' },
    { Role: 'inspector', RoleName: 'ผู้ตรวจนับ', Description: 'ตรวจนับผ่านมือถือ/สแกน QR ตามที่ได้รับมอบหมาย' },
    { Role: 'staff', RoleName: 'ผู้ใช้ทั่วไป', Description: 'ดูทะเบียน ขอยืม ขอเบิก แจ้งซ่อม' }]);
  if (!rows('Users').length) { const s = Utilities.getUuid(); ins('Users', { UserID: uid('USR'), Username: 'admin', FullName: 'ผู้ดูแลระบบ', Role: 'admin', DeptID: '', PasswordHash: hashPw('admin1234', s), Salt: s, TelegramChatID: '', Status: 'ใช้งาน' }); }
  if (!find('Settings', 'Key', 'ORG_NAME')) ins('Settings', { Key: 'ORG_NAME', Value: 'เทศบาลตำบลบึงโขงหลง' });
  if (!rows('FiscalYears').length) { const y = fyOf(today()); ins('FiscalYears', { FiscalYear: y, StartDate: (y - 544) + '-10-01', EndDate: (y - 543) + '-09-30', Status: 'ใช้งาน' }); }
  if (!find('Settings', 'Key', 'SEEDED')) { seedSample(); ins('Settings', { Key: 'SEEDED', Value: now() }); }
  return 'ตั้งค่าระบบเรียบร้อย: ' + Object.keys(SCHEMA).length + ' ชีต | ผู้ใช้เริ่มต้น admin / admin1234';
}
function seedSample() {
  const S = 'SEED', sys = { username: 'SYSTEM' };
  const sg = o => Object.assign({ CreatedBy: S }, o);
  if (!rows('Departments').length) {
    insMany('Departments', [['DEP-01', 'กองคลัง'], ['DEP-02', 'สำนักปลัด'], ['DEP-03', 'กองช่าง'], ['DEP-04', 'กองสาธารณสุขและสิ่งแวดล้อม'], ['DEP-05', 'กองการศึกษา']].map(x => ({ DeptID: x[0], DeptName: x[1], Status: 'ใช้งาน' })));
    insMany('Locations', [['LOC-01', 'ห้องกองคลัง', 'DEP-01'], ['LOC-02', 'ห้องสำนักปลัด', 'DEP-02'], ['LOC-03', 'โรงพัสดุกองช่าง', 'DEP-03'], ['LOC-04', 'ห้องประชุมใหญ่', 'DEP-02'], ['LOC-05', 'ห้องกองสาธารณสุข', 'DEP-04']].map(x => ({ LocationID: x[0], LocationName: x[1], DeptID: x[2], Status: 'ใช้งาน' })));
    insMany('AssetCategories', [['CAT-01', 'ครุภัณฑ์คอมพิวเตอร์', 'ครุภัณฑ์'], ['CAT-02', 'ครุภัณฑ์สำนักงาน/เฟอร์นิเจอร์', 'ครุภัณฑ์'], ['CAT-03', 'ครุภัณฑ์ไฟฟ้าและวิทยุ', 'ครุภัณฑ์'], ['CAT-04', 'ครุภัณฑ์ยานพาหนะและขนส่ง', 'ครุภัณฑ์'], ['CAT-05', 'ครุภัณฑ์การเกษตร/เครื่องจักรกล', 'ครุภัณฑ์'], ['CAT-06', 'ครุภัณฑ์การแพทย์', 'ครุภัณฑ์'], ['CAT-07', 'วัสดุสำนักงาน', 'วัสดุ'], ['CAT-08', 'วัสดุสาธารณสุข', 'วัสดุ']].map(x => ({ CategoryID: x[0], CategoryName: x[1], Type: x[2], Status: 'ใช้งาน' })));
    insMany('DepreciationPolicies', [['CAT-01', 3], ['CAT-02', 8], ['CAT-03', 5], ['CAT-04', 8], ['CAT-05', 5], ['CAT-06', 5]].map(x => ({ PolicyID: 'POL-' + x[0].slice(-2), CategoryID: x[0], LifeYears: x[1], SalvageType: 'amount', SalvageValue: 0, Method: 'เส้นตรง (ค่าตัวอย่าง – ต้องตรวจสอบตามหลักเกณฑ์หน่วยงาน)', Status: 'ใช้งาน' })));
  }
  const nm = 'นางสาวสมหญิง ใจดี', nm2 = 'นายสมชาย รักงาน', nm3 = 'นางมานี มีสุข';
  const A = [
    ['ตย.001/2565', 'โน้ตบุ๊ก', 'CAT-01', 'Lenovo', 'ThinkPad E14', 'SN-LN-0001', 28900, '2022-03-15', 'DEP-01', 'LOC-01', nm],
    ['ตย.002/2565', 'เครื่องพิมพ์เลเซอร์', 'CAT-01', 'HP', 'LaserJet M404dn', 'SN-HP-0002', 12500, '2022-03-15', 'DEP-01', 'LOC-01', nm],
    ['ตย.003/2564', 'เครื่องปรับอากาศแบบแยกส่วน 24,000 BTU', 'CAT-03', 'Daikin', 'FTM24', 'SN-DK-0003', 38500, '2021-06-01', 'DEP-02', 'LOC-04', nm2],
    ['ตย.004/2563', 'โต๊ะทำงานเหล็ก', 'CAT-02', '-', '-', '-', 6800, '2020-10-20', 'DEP-02', 'LOC-02', nm2],
    ['ตย.005/2563', 'ตู้เหล็กเก็บเอกสาร 4 ลิ้นชัก', 'CAT-02', '-', '-', '-', 4900, '2020-10-20', 'DEP-01', 'LOC-01', nm],
    ['ตย.006/2566', 'รถกระบะบรรทุก 4 ประตู', 'CAT-04', 'Toyota', 'Hilux Revo', 'SN-TY-0006', 799000, '2023-02-10', 'DEP-03', 'LOC-03', nm3],
    ['ตย.007/2565', 'เครื่องตัดหญ้าสะพายบ่า', 'CAT-05', 'Honda', 'GX35', 'SN-HD-0007', 14500, '2022-08-05', 'DEP-03', 'LOC-03', nm3],
    ['ตย.008/2566', 'เครื่องวัดความดันโลหิตดิจิทัล', 'CAT-06', 'Omron', 'HEM-7156', 'SN-OM-0008', 8900, '2023-05-12', 'DEP-04', 'LOC-05', 'นางสาวใจใส ใฝ่ดี'],
    ['ตย.009/2564', 'เครื่องโปรเจคเตอร์', 'CAT-03', 'Epson', 'EB-X06', 'SN-EP-0009', 24500, '2021-11-18', 'DEP-02', 'LOC-04', nm2],
    ['ตย.010/2565', 'เครื่องสำรองไฟ UPS 1 kVA', 'CAT-01', 'APC', 'BX1100', 'SN-APC-0010', 6900, '2022-03-15', 'DEP-01', 'LOC-01', nm]
  ].map((x, i) => sg({ AssetID: 'AST-SAMPLE' + ('0' + (i + 1)).slice(-2), RegNo: x[0], Name: x[1], CategoryID: x[2], Brand: x[3], Model: x[4], SerialNo: x[5], Qty: 1, Unit: 'เครื่อง', Cost: x[6], ReceivedDate: x[7], InspectedDate: x[7], PurchaseDocNo: 'PO-ตย./' + (i + 1), FundSource: 'เงินรายได้', DeptID: x[8], LocationID: x[9], Custodian: x[10], Status: i === 8 ? ST.loan : i === 9 ? ST.rep : ST.ok, Notes: 'ข้อมูลตัวอย่าง – ลบได้ที่ ตั้งค่า > ข้อมูลตัวอย่าง' }));
  insMany('Assets', A);
  insMany('AssetHistory', A.map(a => sg({ HistID: uid('H'), AssetID: a.AssetID, Action: 'เพิ่มครุภัณฑ์ (ข้อมูลตัวอย่าง)', Detail: a.RegNo, ActionAt: now(), ActionBy: 'SYSTEM' })));
  insMany('Materials', [['MAT-S1', 'M-001', 'กระดาษ A4 70 แกรม', 'รีม', 'CAT-07', 30], ['MAT-S2', 'M-002', 'ปากกาลูกลื่นสีน้ำเงิน', 'ด้าม', 'CAT-07', 60], ['MAT-S3', 'M-003', 'แฟ้มสันกว้าง 3 นิ้ว', 'เล่ม', 'CAT-07', 10], ['MAT-S4', 'M-004', 'ตลับหมึกพิมพ์เลเซอร์', 'ตลับ', 'CAT-07', 3], ['MAT-S5', 'M-005', 'แอลกอฮอล์ 70% ขนาด 450 มล.', 'ขวด', 'CAT-08', 10]].map(x => sg({ MaterialID: x[0], Code: x[1], Name: x[2], Unit: x[3], CategoryID: x[4], ReorderPoint: x[5], Status: 'ใช้งาน' })));
  const d60 = addDays(-60), d20 = addDays(-20), d10 = addDays(-10);
  insMany('StockLots', [['LOT-S1', 'MAT-S1', 'RC-ตย./001', d60, 100, 70, 105], ['LOT-S2', 'MAT-S1', 'RC-ตย./002', d20, 50, 50, 110], ['LOT-S3', 'MAT-S2', 'RC-ตย./003', d60, 200, 50, 5], ['LOT-S4', 'MAT-S3', 'RC-ตย./004', d60, 40, 40, 38], ['LOT-S5', 'MAT-S4', 'RC-ตย./005', d60, 6, 4, 1650], ['LOT-S6', 'MAT-S5', 'RC-ตย./006', d20, 24, 24, 45]].map(x => sg({ LotID: x[0], MaterialID: x[1], ReceiptDocNo: x[2], ReceivedDate: x[3], QtyIn: x[4], QtyRemain: x[5], UnitCost: x[6], Supplier: 'ร้านตัวอย่าง' })));
  const tx = [['MAT-S1', 'LOT-S1', 'IN', 100, 105, 'RC-ตย./001', d60, ''], ['MAT-S1', 'LOT-S2', 'IN', 50, 110, 'RC-ตย./002', d20, ''], ['MAT-S2', 'LOT-S3', 'IN', 200, 5, 'RC-ตย./003', d60, ''], ['MAT-S3', 'LOT-S4', 'IN', 40, 38, 'RC-ตย./004', d60, ''], ['MAT-S4', 'LOT-S5', 'IN', 6, 1650, 'RC-ตย./005', d60, ''], ['MAT-S5', 'LOT-S6', 'IN', 24, 45, 'RC-ตย./006', d20, ''],
  ['MAT-S1', 'LOT-S1', 'OUT', 30, 105, 'RQ-ตย./001', d10, 'DEP-03'], ['MAT-S2', 'LOT-S3', 'OUT', 150, 5, 'RQ-ตย./001', d10, 'DEP-03'], ['MAT-S4', 'LOT-S5', 'OUT', 2, 1650, 'RQ-ตย./001', d10, 'DEP-03']];
  insMany('StockTransactions', tx.map((x, i) => sg({ TxID: 'TX-S' + (i + 1), MaterialID: x[0], LotID: x[1], Type: x[2], Qty: x[3], UnitCost: x[4], RefDoc: x[5], TxDate: x[6], ByUser: 'SYSTEM', DeptID: x[7] })));
  insMany('Requisitions', [sg({ ReqID: 'RQ-S1', DocNo: 'RQ-ตย./001', DeptID: 'DEP-03', RequestedBy: nm3, Purpose: 'ใช้ในงานธุรการกองช่าง', ReqDate: d10, Status: 'จ่ายแล้ว', ApprovedBy: 'SYSTEM', ApprovedAt: d10, IssuedBy: 'SYSTEM', IssuedAt: d10 }), sg({ ReqID: 'RQ-S2', DocNo: 'RQ-ตย./002', DeptID: 'DEP-01', RequestedBy: nm, Purpose: 'ใช้พิมพ์เอกสารงบประมาณ', ReqDate: today(), Status: 'รออนุมัติ' })]);
  insMany('RequisitionDetails', [['RD-S1', 'RQ-S1', 'MAT-S1', 30, 30], ['RD-S2', 'RQ-S1', 'MAT-S2', 150, 150], ['RD-S3', 'RQ-S1', 'MAT-S4', 2, 2], ['RD-S4', 'RQ-S2', 'MAT-S1', 20, 0]].map(x => sg({ DetailID: x[0], ReqID: x[1], MaterialID: x[2], Qty: x[3], IssuedQty: x[4] })));
  insMany('Loans', [sg({ LoanID: 'LN-S1', DocNo: 'LN-ตย./001', Borrower: 'นายสมศักดิ์ ตั้งใจ', DeptID: 'DEP-05', Purpose: 'ใช้จัดอบรมโรงเรียน', LoanDate: addDays(-20), DueDate: addDays(-5), Status: 'ยืมอยู่', RequestedBy: 'SYSTEM', ApprovedBy: 'SYSTEM', ApprovedAt: addDays(-20) })]);
  insMany('LoanDetails', [sg({ DetailID: 'LD-S1', LoanID: 'LN-S1', AssetID: 'AST-SAMPLE09', ConditionOut: 'ปกติ' })]);
  insMany('Maintenance', [sg({ MaintID: 'MT-S1', DocNo: 'MT-ตย./001', AssetID: 'AST-SAMPLE10', Problem: 'เปิดเครื่องไม่ติด แบตเตอรี่เสื่อม', ReportedBy: nm, ReportDate: addDays(-3), Status: 'รออนุมัติ' })]);
  insMany('MaintenanceHistory', [sg({ HistID: 'MH-S1', MaintID: 'MT-S1', Status: 'รออนุมัติ', Note: 'แจ้งซ่อมตัวอย่าง', At: now(), By: nm })]);
  const rid = 'RD-S1';
  insMany('AnnualInspections', [sg({ RoundID: rid, FiscalYear: fyOf(today()), Name: 'ตรวจนับประจำปี (ตัวอย่าง)', StartDate: addDays(-10), EndDate: addDays(20), Committee: 'คณะกรรมการตัวอย่าง', Status: 'เปิด' })]);
  insMany('InspectionResults', [['AST-SAMPLE01', 'พบ'], ['AST-SAMPLE02', 'พบ'], ['AST-SAMPLE04', 'พบ'], ['AST-SAMPLE06', 'ชำรุด']].map((x, i) => sg({ ResultID: 'IR-S' + (i + 1), RoundID: rid, AssetID: x[0], Result: x[1], Note: x[1] === 'ชำรุด' ? 'กระจกมองข้างแตก (ตัวอย่าง)' : '', InspectedBy: 'SYSTEM', InspectedAt: now() })));
}

// เพิ่ม 2 ฟังก์ชันนี้ลงในไฟล์ .gs (เช่น Code.gs) ต่อจากโค้ดเดิมที่มีอยู่

/**
 * รับ Request จาก GitHub Pages (Frontend)
 */
function doPost(e) {
  try {
    // 1. รับและแปลงข้อมูลที่ส่งมาจาก fetch() 
    var req = JSON.parse(e.postData.contents);
    
    // 2. เรียกฟังก์ชัน api() ที่คุณมีอยู่แล้วทำงาน
    var result = api(req.token, req.action, req.payload);
    
    // 3. ส่งผลลัพธ์กลับไปยัง Frontend เป็นรูปแบบ JSON
    return ContentService.createTextOutput(JSON.stringify({ ok: true, data: result }))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch (err) {
    // จัดการ Error เช่น Token หมดอายุ หรือข้อผิดพลาดอื่นๆ
    var code = (err.message && err.message.indexOf('AUTH') > -1) ? 'AUTH' : 'ERROR';
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message, code: code }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ฟังก์ชันสำหรับการตั้งค่า CORS preflight (เผื่อเบราว์เซอร์ร้องขอ OPTIONS/GET)
 * และป้องกัน Error กรณีเปิด URL Web App ด้วยเบราว์เซอร์ตรงๆ
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "API is running", timestamp: new Date() }))
                       .setMimeType(ContentService.MimeType.JSON);
}
