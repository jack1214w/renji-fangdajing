/**
 * 我的使用說明書 — 後端
 * 部署：Apps Script 專案（建議從 Google 試算表的「擴充功能 → Apps Script」建立，
 * 這樣會自動綁定該試算表，SHEET_ID 可留空）。
 */

const CONFIG = {
  SHEET_ID: '',               // 留空＝使用綁定的試算表；否則填入試算表網址中的那串 ID
  SHEET_NAME: '使用說明書',
  LOG_NAME: '修改紀錄',
  TEACHER_PASSWORD: '1940',   // 老師檢視頁的密碼
  MAX_LEN: 1000               // 單一欄位字數上限
};

const HEADERS = [
  '更新時間', '班級', '座號', '姓名', '產品型號',
  '使用方法', '使用方法開放',
  '充電', '耗電', '電力開放',
  '注意事項', '注意事項開放',
  '故障排除', '故障排除開放',
  '保固條款', '保固開放',
  '兩週後回顧', '原始資料'
];

/* ===== 網頁入口 ===== */

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'form';
  const file = page === 'teacher' ? 'Teacher' : 'Index';
  return HtmlService.createTemplateFromFile(file)
    .evaluate()
    .setTitle('我的使用說明書')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* ===== 試算表 ===== */

function getBook_() {
  return CONFIG.SHEET_ID ? SpreadsheetApp.openById(CONFIG.SHEET_ID)
                         : SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_() {
  const book = getBook_();
  if (!book) throw new Error('找不到試算表。請從試算表的「擴充功能 → Apps Script」建立專案，或在 CONFIG.SHEET_ID 填入試算表 ID。');
  let sheet = book.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = book.insertSheet(CONFIG.SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#E2EAF3');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    [6, 8, 9, 11, 13, 15, 17].forEach(c => sheet.setColumnWidth(c, 320));
    sheet.hideColumns(HEADERS.indexOf('原始資料') + 1);
  }
  return sheet;
}

/** 建立／修復試算表結構。可從編輯器直接執行一次。 */
function setup() {
  const sheet = getSheet_();
  SpreadsheetApp.flush();
  return '已就緒：' + sheet.getParent().getName() + ' / ' + sheet.getName();
}

/* ===== 工具 ===== */

function text_(v) {
  const s = String(v == null ? '' : v).trim();
  return s.length > CONFIG.MAX_LEN ? s.slice(0, CONFIG.MAX_LEN) : s;
}

/** 避免以 = + - @ 開頭的內容被試算表當成公式 */
function safe_(v) {
  const s = text_(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function openLevel_(v) {
  const allow = ['公開', '小組', '只給老師'];
  return allow.indexOf(String(v)) >= 0 ? String(v) : '只給老師';
}

function keyOf_(cls, no) {
  return text_(cls) + '-' + text_(no);
}

function troubleToText_(list) {
  if (!Array.isArray(list)) return '';
  return list
    .filter(r => r && (r.s || r.c || r.a))
    .map(r => [text_(r.s), text_(r.c), text_(r.a)].join('｜'))
    .join('\n');
}

function findRow_(sheet, cls, no) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const values = sheet.getRange(2, 2, last - 1, 2).getValues(); // 班級、座號
  for (let i = 0; i < values.length; i++) {
    if (keyOf_(values[i][0], values[i][1]) === keyOf_(cls, no)) return i + 2;
  }
  return 0;
}

/* ===== 學生端 API ===== */

/**
 * 儲存（同一班級＋座號會覆蓋前一次，並在「修改紀錄」留存舊版）
 * @param {Object} data 前端送出的表單資料
 */
function saveManual(data) {
  try {
    const d = data || {};
    if (!text_(d.cls) || !text_(d.no) || !text_(d.name)) {
      return { ok: false, msg: '請先填寫班級、座號、姓名。' };
    }
    if (!text_(d.model) && !text_(d.useText)) {
      return { ok: false, msg: '至少要填「產品型號」和「使用方法」。' };
    }

    const sheet = getSheet_();
    const now = new Date();
    const row = [
      now, safe_(d.cls), safe_(d.no), safe_(d.name), safe_(d.model),
      safe_(d.useText), openLevel_(d.useOpen),
      safe_(d.chargeText), safe_(d.drainText), openLevel_(d.powerOpen),
      safe_(d.warnText), openLevel_(d.warnOpen),
      safe_(troubleToText_(d.trouble)), openLevel_(d.troubleOpen),
      safe_(d.warrantyText), openLevel_(d.warrantyOpen),
      safe_(d.reviewText),
      JSON.stringify(d).slice(0, 40000)
    ];

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const at = findRow_(sheet, d.cls, d.no);
      if (at) {
        archive_(sheet, at);
        sheet.getRange(at, 1, 1, HEADERS.length).setValues([row]);
      } else {
        sheet.appendRow(row);
      }
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    return { ok: true, savedAt: Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') };
  } catch (err) {
    return { ok: false, msg: '儲存失敗：' + err.message };
  }
}

/** 覆蓋前先把舊版留在「修改紀錄」工作表 */
function archive_(sheet, rowIndex) {
  const book = sheet.getParent();
  let log = book.getSheetByName(CONFIG.LOG_NAME);
  if (!log) {
    log = book.insertSheet(CONFIG.LOG_NAME);
    log.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    log.setFrozenRows(1);
  }
  const old = sheet.getRange(rowIndex, 1, 1, HEADERS.length).getValues();
  log.appendRow(old[0]);
}

/** 讀回自己之前寫的內容（班級＋座號＋姓名三者都要對） */
function loadManual(query) {
  try {
    const q = query || {};
    const sheet = getSheet_();
    const at = findRow_(sheet, q.cls, q.no);
    if (!at) return { ok: false, msg: '找不到資料。可能還沒存過，或班級座號不對。' };

    const row = sheet.getRange(at, 1, 1, HEADERS.length).getValues()[0];
    if (text_(row[3]) !== text_(q.name)) {
      return { ok: false, msg: '姓名和這個座號的紀錄不符。' };
    }
    let data = {};
    try { data = JSON.parse(row[HEADERS.indexOf('原始資料')] || '{}'); } catch (e) { data = {}; }
    return {
      ok: true,
      data: data,
      savedAt: row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') : ''
    };
  } catch (err) {
    return { ok: false, msg: '讀取失敗：' + err.message };
  }
}

/* ===== 老師端 API ===== */

function listManuals(password, cls) {
  if (String(password) !== CONFIG.TEACHER_PASSWORD) {
    return { ok: false, msg: '密碼不正確。' };
  }
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return { ok: true, rows: [], headers: HEADERS.slice(0, 17) };

  const values = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const tz = Session.getScriptTimeZone();
  const rows = values
    .filter(r => !cls || String(r[1]) === String(cls))
    .map(r => {
      const o = r.slice(0, 17);
      o[0] = r[0] ? Utilities.formatDate(new Date(r[0]), tz, 'MM/dd HH:mm') : '';
      return o;
    })
    .sort((a, b) => String(a[1] + a[2]).localeCompare(String(b[1] + b[2]), 'zh-Hant', { numeric: true }));

  return { ok: true, rows: rows, headers: HEADERS.slice(0, 17) };
}

function listClasses(password) {
  if (String(password) !== CONFIG.TEACHER_PASSWORD) return { ok: false, msg: '密碼不正確。' };
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return { ok: true, classes: [] };
  const values = sheet.getRange(2, 2, last - 1, 1).getValues().map(r => String(r[0]));
  const uniq = values.filter((v, i) => v && values.indexOf(v) === i).sort();
  return { ok: true, classes: uniq };
}

/** 回傳試算表網址，方便老師直接開啟 */
function sheetUrl(password) {
  if (String(password) !== CONFIG.TEACHER_PASSWORD) return { ok: false, msg: '密碼不正確。' };
  return { ok: true, url: getSheet_().getParent().getUrl() };
}
