/**
 * Skin Guard – ระบบกลางบน Google Sheets + Google Drive
 * วางโค้ดนี้ใน Apps Script (จากเมนู Extensions ของ Google Sheet หรือสร้าง New project ที่ script.google.com ก็ได้) แล้ว Deploy เป็น Web app
 * 1) เปลี่ยน KEY ด้านล่างเป็นรหัสลับของหน่วยงาน (ยาวอย่างน้อย 16 ตัว) และห้ามเผยแพร่
 * 2) Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone
 */
const KEY = 'CHANGE-ME-TO-A-LONG-SECRET';
const PHOTO_FOLDER = 'SkinGuard-Photos';

const SHEETS = {
  patients: ['hn', 'updated', 'ot', 'deleted', 'json'],
  history:  ['id', 'hn', 'date', 'updated', 'ot', 'deleted', 'json'],
  turns:    ['hn', 'updated', 'ot', 't'],
  photos:   ['id', 'store', 'hn', 'date', 'site', 'updated', 'ot', 'deleted', 'fileId', 'meta'],
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
    if (req.action === 'push') return out(push(req));
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
      const meta = Object.assign({}, rec); delete meta.img;
      upsert(F, fi, it.key, [it.key, it.store || '', rec.hn || '', rec.date || '', rec.site || '', now, ot, '', fileId, JSON.stringify(meta)]);
    }
    else if (it.type === 'delphoto') {
      const r = fi[String(it.key)];
      if (r) { const fid = F.getRange(r, 9).getValue(); if (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} } }
      upsert(F, fi, it.key, [it.key, it.store || '', '', '', '', now, ot, 'y', '', '']);
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
