/**
 * 生涯排排樂：Google Apps Script 後端
 * 把這整份程式貼到 Google 試算表的「擴充功能 → Apps Script」，
 * 再部署為「網頁應用程式」（執行身分：我；存取權：所有人）。
 */

const PASSWORD = '0214';          // 講者密碼，要和網頁一致
const SHEET_NAME = 'responses';
const HEADER = ['clientId', 'nickname', 'updatedAt', 'data'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 避免暱稱被試算表當成公式
function safeText_(s, max) {
  s = String(s || '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'results') {
    if (p.pw !== PASSWORD) return out_({ ok: false, error: 'bad_password' });
    const values = sheet_().getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < values.length; i++) {
      const r = values[i];
      if (!r[0]) continue;
      try {
        rows.push({
          clientId: String(r[0]),
          nickname: String(r[1] || '').replace(/^'/, ''),
          updatedAt: r[2] instanceof Date ? r[2].toISOString() : String(r[2]),
          data: JSON.parse(r[3])
        });
      } catch (err) { /* 略過損壞的資料列 */ }
    }
    return out_({ ok: true, rows: rows });
  }
  return out_({ ok: true, service: 'career-sort' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: 'bad_json' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();

    if (body.action === 'submit') {
      const id = String(body.clientId || '').slice(0, 64);
      if (!id || !body.data || !Array.isArray(body.data.categories)) {
        return out_({ ok: false, error: 'bad_data' });
      }
      const data = JSON.stringify(body.data);
      if (data.length > 20000) return out_({ ok: false, error: 'too_large' });

      const row = [id, safeText_(body.nickname, 20), new Date(), data];
      const last = sh.getLastRow();
      const ids = last > 1
        ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); })
        : [];
      const idx = ids.indexOf(id);
      if (idx >= 0) {
        sh.getRange(idx + 2, 1, 1, HEADER.length).setValues([row]);  // 重新送出：覆蓋舊資料
      } else {
        sh.appendRow(row);
      }
      return out_({ ok: true });
    }

    if (body.action === 'clear') {
      if (body.pw !== PASSWORD) return out_({ ok: false, error: 'bad_password' });
      const last = sh.getLastRow();
      if (last > 1) sh.deleteRows(2, last - 1);
      return out_({ ok: true });
    }

    return out_({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}
