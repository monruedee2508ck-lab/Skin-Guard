/**
 * Skin Guard – ระบบกลางบน Google Sheets + Google Drive
 * วางโค้ดนี้ใน Apps Script (จากเมนู Extensions ของ Google Sheet หรือสร้าง New project ที่ script.google.com ก็ได้) แล้ว Deploy เป็น Web app
 * 1) เปลี่ยน KEY ด้านล่างเป็นรหัสลับของหน่วยงาน (ยาวอย่างน้อย 16 ตัว) และห้ามเผยแพร่
 * 2) Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone
 */
const KEY = 'CHANGE-ME-TO-A-LONG-SECRET';
const PHOTO_FOLDER = 'SkinGuard-Photos';
const KEEP_DAYS = 7;            // ลบภาพอัตโนมัติเมื่อครบ 7 วันนับจากวันที่ถ่าย
const VIEW_TAB = 'ภาพรายเวร';    // แท็บสำหรับเปิดดูภาพ วันที่ เวลา และเวร
const TZ = 'Asia/Bangkok';

const SHEETS = {
  patients: ['hn', 'updated', 'ot', 'deleted', 'json'],
  history:  ['id', 'hn', 'date', 'updated', 'ot', 'deleted', 'json'],
  turns:    ['hn', 'updated', 'ot', 't'],
  photos:   ['id', 'store', 'hn', 'date', 'site', 'updated', 'ot', 'deleted', 'fileId', 'meta', 'time', 'shift', 'kind', 'expires'],
  log:      ['time', 'ot', 'action', 'key'],
  admins:   ['email', 'name', 'salt', 'hash', 'status', 'created', 'approvedBy', 'token', 'tokenExp']
};

const BOOK_NAMES = ['Skin guard-ข้อมูลกลาง', 'Skin Guard – ข้อมูลกลาง', 'Skin Guard - ข้อมูลกลาง'];
let _book = null;
/** ใช้ได้ทั้งแบบผูกกับ Google Sheet และแบบสร้างโปรเจกต์ใหม่ที่ script.google.com (จะหา/สร้างสเปรดชีตให้เอง) */
function book() {
  if (_book) return _book;
  try { const a = SpreadsheetApp.getActiveSpreadsheet(); if (a) return (_book = a); } catch (e) {}
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SS_ID');
  if (id) { try { return (_book = SpreadsheetApp.openById(id)); } catch (e) {} }
  for (let i = 0; i < BOOK_NAMES.length; i++) {
    const it = DriveApp.getFilesByName(BOOK_NAMES[i]);
    if (it.hasNext()) { _book = SpreadsheetApp.openById(it.next().getId()); break; }
  }
  if (!_book) _book = SpreadsheetApp.create(BOOK_NAMES[1]);
  props.setProperty('SS_ID', _book.getId());
  return _book;
}

/**
 * กด ▶ Run ฟังก์ชันนี้ 1 ครั้งหลังวางโค้ด เพื่อสร้างแท็บข้อมูลทั้งหมด โฟลเดอร์เก็บภาพ และขออนุญาตใช้งาน
 */
function setup() {
  if (KEY.indexOf('CHANGE-ME') === 0) throw new Error('ยังไม่ได้เปลี่ยน KEY ด้านบนเป็นรหัสลับของหน่วยงาน');
  const ss = book();
  Object.keys(SHEETS).forEach(function (name) {
    const sh = sheet(name);
    const head = sh.getRange(1, 1, 1, SHEETS[name].length);
    head.setValues([SHEETS[name]]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#132b4f');
    sh.setFrozenRows(1);
  });
  const first = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  const f = folder();
  viewSheet(); dailySheet(); woundSheet(); KPI_TABS.forEach(kpiSheet);
  ensureCleanupTrigger();
  Logger.log('ตั้งเวลาลบภาพอัตโนมัติทุกวัน 01:00 น. (ภาพครบ ' + KEEP_DAYS + ' วัน หรือผู้ป่วยจำหน่ายแล้ว)');
  Logger.log('พร้อมใช้งาน: สร้างแท็บ ' + Object.keys(SHEETS).join(', ') + ' และโฟลเดอร์ภาพ ' + f.getName() + ' แล้ว');
  Logger.log('สเปรดชีต: ' + ss.getName() + ' ' + ss.getUrl());
  Logger.log('ขั้นต่อไป: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone');
}

function doGet() { return out({ ok: true, app: 'SkinGuard', time: Date.now() }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'bad_json' }); }
  if (KEY.indexOf('CHANGE-ME') === 0) return out({ ok: false, error: 'key_not_set' });
  if (!req || req.key !== KEY) return out({ ok: false, error: 'unauthorized' });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return out({ ok: false, error: 'busy' });
  try {
    if (req.action === 'ping') return out({ ok: true, time: Date.now() });
    if (req.action === 'push') { const r = push(req); try { maybeCleanup(); } catch (e) {} return out(r); }
    if (req.action === 'pull') return out(pull(req));
    if (req.action === 'photo') return out(getPhoto(req.id));
    if (req.action === 'adminRegister') return out(adminRegister(req));
    if (req.action === 'adminLogin') return out(adminLogin(req));
    if (req.action === 'adminList') return out(adminList(req));
    if (req.action === 'adminSet') return out(adminSet(req));
    return out({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet(name) {
  const ss = book();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(SHEETS[name]); sh.setFrozenRows(1); }
  else if (sh.getLastColumn() < SHEETS[name].length) sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#132b4f');
  return sh;
}

function rowIndex(sh) {
  const last = sh.getLastRow(), map = {};
  if (last < 2) return map;
  sh.getRange(2, 1, last - 1, 1).getValues().forEach((r, i) => { map[String(r[0])] = i + 2; });
  return map;
}

function upsert(sh, idx, key, row) {
  const r = idx[String(key)];
  if (r) sh.getRange(r, 1, 1, row.length).setValues([row]);
  else { sh.appendRow(row); idx[String(key)] = sh.getLastRow(); }
}

function folder() {
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}

function push(req) {
  const now = Date.now(), items = req.items || [], done = [];
  const P = sheet('patients'), H = sheet('history'), T = sheet('turns'), F = sheet('photos'), L = sheet('log');
  const pi = rowIndex(P), hi = rowIndex(H), ti = rowIndex(T), fi = rowIndex(F);
  items.forEach(it => {
    const ot = String(it.ot || '');
    if (it.type === 'patient') upsert(P, pi, it.key, [it.key, now, ot, '', JSON.stringify(it.data)]);
    else if (it.type === 'delpatient') upsert(P, pi, it.key, [it.key, now, ot, 'y', '']);
    else if (it.type === 'hist') upsert(H, hi, it.key, [it.key, it.data.hn || '', it.data.date || '', now, ot, '', JSON.stringify(it.data)]);
    else if (it.type === 'delhist') upsert(H, hi, it.key, [it.key, '', '', now, ot, 'y', '']);
    else if (it.type === 'turn') upsert(T, ti, it.key, [it.key, now, ot, Number(it.data.t) || 0]);
    else if (it.type === 'photo') {
      const rec = it.data || {}, img = String(rec.img || '');
      const m = img.match(/^data:([^;]+);base64,(.*)$/);
      let fileId = '';
      if (m) {
        const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], it.key + '.jpg');
        fileId = folder().createFile(blob).getId();
      }
      const view = rec._view || {}, thumb = view.thumb || '';
      const meta = Object.assign({}, rec); delete meta.img; delete meta._view;
      const time = String(rec.time || ''), sh = shiftOf(time), exp = addDays(dstr(rec.date), KEEP_DAYS);
      upsert(F, fi, it.key, ["'" + it.key, it.store || '', "'" + (rec.hn || ''), "'" + dstr(rec.date), rec.site || '', now, ot, '', fileId, JSON.stringify(meta), "'" + time, sh, view.kind || rec.kind || it.store || '', "'" + exp]);
      try { viewAdd(it.key, rec, view, thumb, sh, exp, ot, fileId); } catch (e) {}
    }
    else if (it.type === 'kpi') { kpiWrite(it.key, it.data || {}, ot); }
    else if (it.type === 'daily') { dailyWrite(it.key, it.data || {}, ot); }
    else if (it.type === 'wound') { woundWrite(it.key, it.data || {}, ot); }
    else if (it.type === 'delphoto') {
      const r = fi[String(it.key)];
      if (r) photoGone(F, r, now, ot);
      else upsert(F, fi, it.key, ["'" + it.key, it.store || '', '', '', '', now, ot, 'y', '', '', '', '', '', '']);
    }
    else return;
    L.appendRow([new Date(now), ot, it.type, String(it.key)]);
    done.push(it.type + '|' + it.key);
  });
  return { ok: true, time: now, done: done };
}

function rowsSince(sh, since, updCol) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return sh.getRange(2, 1, last - 1, head.length).getValues()
    .filter(r => Number(r[updCol]) > since)
    .map(r => { const o = {}; head.forEach((h, i) => { o[h] = r[i]; }); return o; });
}

function pull(req) {
  const since = Number(req.since) || 0, now = Date.now();
  return {
    ok: true, time: now,
    patients: rowsSince(sheet('patients'), since, 1),
    history: rowsSince(sheet('history'), since, 3),
    turns: rowsSince(sheet('turns'), since, 1),
    photos: rowsSince(sheet('photos'), since, 5).map(r => { delete r.fileId; return r; })
  };
}

function getPhoto(id) {
  const F = sheet('photos'), r = rowIndex(F)[String(id)];
  if (!r) return { ok: false, error: 'not_found' };
  const row = F.getRange(r, 1, 1, 10).getValues()[0], fid = row[8];
  if (!fid || row[7] === 'y') return { ok: false, error: 'deleted' };
  const file = DriveApp.getFileById(fid), blob = file.getBlob();
  const meta = JSON.parse(row[9] || '{}');
  meta.img = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  return { ok: true, store: row[1], rec: meta };
}

/* ---------- Admin accounts ---------- */
function sha(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function admins() {
  const sh = sheet('admins'), last = sh.getLastRow();
  if (last < 2) return { sh, rows: [] };
  const rows = sh.getRange(2, 1, last - 1, 9).getValues().map((r, i) => ({
    row: i + 2, email: String(r[0]).toLowerCase(), name: r[1], salt: r[2], hash: r[3], status: r[4], token: r[7], tokenExp: Number(r[8]) || 0 }));
  return { sh, rows };
}
function adminRegister(req) {
  const email = String(req.email || '').trim().toLowerCase(), name = String(req.name || '').trim(), pass = String(req.pass || '');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || name.length < 2 || pass.length < 6) return { ok: false, error: 'invalid_input' };
  const A = admins();
  if (A.rows.some(r => r.email === email)) return { ok: false, error: 'exists' };
  const salt = Utilities.getUuid(), first = !A.rows.some(r => r.status === 'approved');
  const status = first ? 'approved' : 'pending';
  A.sh.appendRow([email, name, salt, sha('skin-guard|' + salt + '|' + pass), status, new Date(), first ? 'first-account' : '', '', '']);
  sheet('log').appendRow([new Date(), email, 'adminRegister', status]);
  return { ok: true, status };
}
function adminLogin(req) {
  const email = String(req.email || '').trim().toLowerCase(), pass = String(req.pass || '');
  const A = admins(), u = A.rows.find(r => r.email === email);
  if (!u || sha('skin-guard|' + u.salt + '|' + pass) !== u.hash) return { ok: false, error: 'invalid' };
  if (u.status !== 'approved') return { ok: false, error: 'pending' };
  const token = Utilities.getUuid(), exp = Date.now() + 12 * 3600 * 1000;
  A.sh.getRange(u.row, 8, 1, 2).setValues([[token, exp]]);
  sheet('log').appendRow([new Date(), email, 'adminLogin', '']);
  return { ok: true, name: u.name, email, token };
}
function adminAuth(token) {
  if (!token) return null;
  const A = admins(), u = A.rows.find(r => r.token === token && r.tokenExp > Date.now() && r.status === 'approved');
  return u ? { A, u } : null;
}
function adminList(req) {
  const a = adminAuth(req.token); if (!a) return { ok: false, error: 'session_expired' };
  return { ok: true, admins: a.A.rows.map(r => ({ email: r.email, name: r.name, status: r.status })) };
}
function adminSet(req) {
  const a = adminAuth(req.token); if (!a) return { ok: false, error: 'session_expired' };
  const email = String(req.email || '').toLowerCase(), t = a.A.rows.find(r => r.email === email);
  if (!t) return { ok: false, error: 'not_found' };
  if (req.op === 'approve') { a.A.sh.getRange(t.row, 5).setValue('approved'); a.A.sh.getRange(t.row, 7).setValue(a.u.email); }
  else if (req.op === 'remove') { if (email === a.u.email) return { ok: false, error: 'self' }; a.A.sh.deleteRow(t.row); }
  else return { ok: false, error: 'bad_op' };
  sheet('log').appendRow([new Date(), a.u.email, 'admin:' + req.op, email]);
  return { ok: true };
}

/* ---------- ภาพรายเวร + ลบภาพอัตโนมัติ ---------- */
const VIEW_HEAD = ['วันที่', 'เวลา', 'เวร', 'HN', 'ประเภทภาพ', 'ตำแหน่ง', 'สภาพผิว / ขนาดแผล', 'บันทึกเพิ่มเติม', 'ผู้บันทึก (OT)', 'ภาพ', 'ลิงก์ภาพ (Google Drive)', 'ลบอัตโนมัติวันที่', 'id'];
function dstr(d) {
  if (d instanceof Date) return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  const m = String(d || '').match(/\d{4}-\d{2}-\d{2}/); return m ? m[0] : '';
}
function addDays(d, n) {
  if (!d) return '';
  const x = new Date(d + 'T00:00:00Z'); x.setUTCDate(x.getUTCDate() + n); return x.toISOString().slice(0, 10);
}
function today() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function shiftOf(t) {
  const h = parseInt(String(t || '').slice(0, 2), 10);
  if (isNaN(h)) return '';
  return h >= 8 && h < 16 ? 'เวรเช้า' : h >= 16 ? 'เวรบ่าย' : 'เวรดึก';
}
function viewSheet() {
  const ss = book();
  let sh = ss.getSheetByName(VIEW_TAB);
  if (!sh) {
    sh = ss.insertSheet(VIEW_TAB, 0);
    sh.getRange(1, 1, 1, VIEW_HEAD.length).setValues([VIEW_HEAD]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#0f766e');
    sh.setFrozenRows(1);
    sh.setColumnWidth(10, 130); sh.setColumnWidth(7, 220); sh.setColumnWidth(11, 160); sh.hideColumns(13);
  }
  return sh;
}
function viewFind(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  const c = sh.getRange(2, 13, sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return c ? c.getRow() : 0;
}
function viewAdd(id, rec, view, thumb, sh, exp, ot, fileId) {
  const V = viewSheet();
  let r = viewFind(V, id);
  const row = ["'" + dstr(rec.date), "'" + String(rec.time || ''), sh, "'" + (rec.hn || ''), view.kind || '', view.site || rec.site || '', view.what || '', view.note || '', "'" + ot, '',
    fileId ? 'https://drive.google.com/file/d/' + fileId + '/view' : '', "'" + exp, "'" + id];
  if (!r) { V.appendRow(row); r = V.getLastRow(); } else V.getRange(r, 1, 1, row.length).setValues([row]);
  V.setRowHeight(r, 96);
  if (thumb && /^data:image\//.test(thumb)) {
    try { V.getRange(r, 10).setValue(SpreadsheetApp.newCellImage().setSourceUrl(thumb).setAltTextTitle(view.site || 'ภาพ').build()); }
    catch (e) { V.getRange(r, 10).setValue('ดูที่ลิงก์'); }
  }
  // เรียงใหม่: ล่าสุดอยู่บน
  if (V.getLastRow() > 2) V.getRange(2, 1, V.getLastRow() - 1, VIEW_HEAD.length).sort([{ column: 1, ascending: false }, { column: 2, ascending: false }]);
}
function viewRemove(id) {
  const V = book().getSheetByName(VIEW_TAB); if (!V) return;
  const r = viewFind(V, id); if (r) V.deleteRow(r);
}
function photoGone(F, r, now, ot) {
  const row = F.getRange(r, 1, 1, 10).getValues()[0], fid = row[8];
  if (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} }
  F.getRange(r, 6, 1, 5).setValues([[now, ot, 'y', '', '']]);
  try { viewRemove(row[0]); } catch (e) {}
}
/** ลบภาพที่ครบ KEEP_DAYS วัน และภาพของผู้ป่วยที่จำหน่ายแล้ว (ทำงานอัตโนมัติทุกวัน และทุกครั้งที่มีการส่งข้อมูล ไม่เกินชั่วโมงละครั้ง) */
function cleanupPhotos() {
  const F = sheet('photos'), last = F.getLastRow();
  if (last < 2) return 0;
  const gone = {}, P = sheet('patients');
  if (P.getLastRow() > 1) P.getRange(2, 1, P.getLastRow() - 1, 4).getValues().forEach(r => { if (r[3] === 'y') gone[String(r[0])] = 1; });
  const rows = F.getRange(2, 1, last - 1, 10).getValues(), t = today(), now = Date.now();
  let n = 0;
  rows.forEach((r, i) => {
    if (r[7] === 'y') return;
    const d = dstr(r[3]), exp = addDays(d, KEEP_DAYS);
    if ((exp && exp <= t) || gone[String(r[2])]) { photoGone(F, i + 2, now, 'auto'); n++; }
  });
  // แถวในแท็บภาพรายเวรที่ไม่มีภาพแล้ว
  const V = book().getSheetByName(VIEW_TAB);
  if (V && V.getLastRow() > 1) {
    const vr = V.getRange(2, 1, V.getLastRow() - 1, VIEW_HEAD.length).getValues();
    for (let i = vr.length - 1; i >= 0; i--) { const exp = dstr(vr[i][11]); if ((exp && exp <= t) || gone[String(vr[i][3])]) V.deleteRow(i + 2); }
  }
  if (n) sheet('log').appendRow([new Date(now), 'auto', 'cleanupPhotos', n + ' ภาพ']);
  return n;
}
function maybeCleanup() {
  const p = PropertiesService.getScriptProperties(), last = Number(p.getProperty('LAST_CLEAN') || 0);
  if (Date.now() - last < 3600e3) return;
  p.setProperty('LAST_CLEAN', String(Date.now()));
  cleanupPhotos();
}
function ensureCleanupTrigger() {
  const has = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'cleanupPhotos');
  if (!has) ScriptApp.newTrigger('cleanupPhotos').timeBased().everyDays(1).atHour(1).inTimezone(TZ).create();
}

/* ---------- ตัวชี้วัดคุณภาพ (Clinical Outcomes / Nursing Performance / QA & Research) ---------- */
const KPI_TABS = ['Clinical Outcomes', 'Nursing Performance', 'QA & Research'];
const KPI_HEAD = ['เดือน', 'ตัวชี้วัด', 'ประเภท', 'ค่า', 'รายละเอียด', 'สถานะ', 'เป้าหมาย', 'อัปเดตล่าสุด'];
function kpiSheet(name) {
  const ss = book();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, KPI_HEAD.length).setValues([KPI_HEAD]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1d4ed8');
    sh.setFrozenRows(1); sh.setColumnWidth(2, 260); sh.setColumnWidth(5, 320);
  }
  return sh;
}
function kpiWrite(month, data, ot) {
  const secs = data.sections || {}, now = new Date();
  KPI_TABS.forEach(name => {
    const rows = secs[name]; if (!rows || !rows.length) return;
    const sh = kpiSheet(name), last = sh.getLastRow();
    if (last > 1) {
      const ms = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
      for (let i = ms.length - 1; i >= 0; i--) if (String(ms[i][0]) === String(month)) sh.deleteRow(i + 2);
    }
    const out = rows.map(r => ["'" + month, r[0], r[1], "'" + r[2], r[3], r[4], r[5], now]);
    sh.getRange(sh.getLastRow() + 1, 1, out.length, KPI_HEAD.length).setValues(out);
  });
}
/* ---------- Braden รายวัน: HN ชื่อ คะแนน Braden ของแต่ละวัน ---------- */
const DAILY_TAB = 'Braden รายวัน';
const DAILY_HEAD = ['วันที่', 'เวลา', 'HN', 'ชื่อ-สกุล', 'Braden', 'ระดับความเสี่ยง', 'การรับความรู้สึก', 'ความเปียกชื้น', 'การทำกิจกรรม', 'การเคลื่อนไหว', 'โภชนาการ', 'แรงเสียดทาน', 'จำนวนแผล', 'PUSH สูงสุด', 'ผู้ประเมิน', 'OT', 'key'];
function dailySheet() {
  const ss = book();
  let sh = ss.getSheetByName(DAILY_TAB);
  if (!sh) {
    sh = ss.insertSheet(DAILY_TAB, 1);
    sh.getRange(1, 1, 1, DAILY_HEAD.length).setValues([DAILY_HEAD]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#0f766e');
    sh.setFrozenRows(1); sh.hideColumns(DAILY_HEAD.length);
  }
  return sh;
}
function dailyWrite(key, d, ot) {
  const sh = dailySheet(), b = d.sub || {}, n = v => (v === undefined || v === null || v === '') ? '' : v;
  const row = ["'" + (d.date || ''), "'" + (d.time || ''), "'" + (d.hn || ''), d.name || '', n(d.braden), d.level || '', n(b.sens), n(b.mois), n(b.act), n(b.mob), n(b.nut), n(b.fric), n(d.wounds), n(d.push), d.nurse || '', "'" + (d.ot || ot || ''), "'" + key];
  let r = 0;
  if (sh.getLastRow() > 1) { const c = sh.getRange(2, DAILY_HEAD.length, sh.getLastRow() - 1, 1).createTextFinder(String(key)).matchEntireCell(true).findNext(); if (c) r = c.getRow(); }
  if (r) sh.getRange(r, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, DAILY_HEAD.length).sort([{ column: 1, ascending: false }, { column: 3, ascending: true }]);
}

/* ---------- แผลกดทับ: ตำแหน่ง ขนาด ระดับ PUSH และวัสดุปิดแผล (1 แถวต่อแผลต่อวัน) ---------- */
const WOUND_TAB = 'แผลกดทับ';
const WOUND_HEAD = ['วันที่', 'เวลา', 'HN', 'ชื่อ-สกุล', 'แผลที่', 'ตำแหน่งแผล', 'ชนิดแผล', 'ระดับแผล', 'ยาว (cm)', 'กว้าง (cm)', 'ลึก (cm)', 'พื้นที่ (cm²)', 'สิ่งคัดหลั่ง (คะแนน)', 'เนื้อเยื่อ (คะแนน)', 'คะแนน PUSH', 'วัสดุที่ใช้ปิดแผล', 'เปลี่ยนวัสดุล่าสุด', 'ผู้ประเมิน', 'OT', 'key'];
function woundSheet() {
  const ss = book();
  let sh = ss.getSheetByName(WOUND_TAB);
  if (!sh) {
    sh = ss.insertSheet(WOUND_TAB, 2);
    sh.getRange(1, 1, 1, WOUND_HEAD.length).setValues([WOUND_HEAD]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#b4234a');
    sh.setFrozenRows(1); sh.setColumnWidth(16, 260); sh.hideColumns(WOUND_HEAD.length);
  }
  return sh;
}
function woundWrite(key, d, ot) {
  const sh = woundSheet(), n = v => (v === undefined || v === null || v === '') ? '' : v;
  const row = ["'" + (d.date || ''), "'" + (d.time || ''), "'" + (d.hn || ''), d.name || '', n(d.no), d.site || '', d.wtype || '', d.stage || '', n(d.L), n(d.W), n(d.D), n(d.area), n(d.exu), n(d.tis), n(d.push), d.dress || '', d.dressAt || '', d.nurse || '', "'" + (d.ot || ot || ''), "'" + key];
  let r = 0;
  if (sh.getLastRow() > 1) { const c = sh.getRange(2, WOUND_HEAD.length, sh.getLastRow() - 1, 1).createTextFinder(String(key)).matchEntireCell(true).findNext(); if (c) r = c.getRow(); }
  if (r) sh.getRange(r, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, WOUND_HEAD.length).sort([{ column: 1, ascending: false }, { column: 3, ascending: true }, { column: 5, ascending: true }]);
}
