/**
 * Skin Guard – ระบบกลางบน Google Sheets + Google Drive
 * วางโค้ดนี้ใน Apps Script ของ Google Sheet ที่สร้างไว้ แล้ว Deploy เป็น Web app
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
  log:      ['time', 'ot', 'action', 'key']
};

function doGet() { return out({ ok: true, app: 'SkinGuard', time: Date.now() }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'bad_json' }); }
  if (!req || req.key !== KEY) return out({ ok: false, error: 'unauthorized' });
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (req.action === 'ping') return out({ ok: true, time: Date.now() });
    if (req.action === 'push') return out(push(req));
    if (req.action === 'pull') return out(pull(req));
    if (req.action === 'photo') return out(getPhoto(req.id));
    return out({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
