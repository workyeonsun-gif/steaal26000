/**
 * 도서 반품 키오스크 — 서버 코드 (Google Apps Script)
 *
 * 시트 구성
 *  - 도서목록   : ISBN | 제목 | 판정보(초판, 개정 2판 등) | 발간일 | 출판사   (도서 마스터, 관리자가 채움)
 *  - 반품기록   : 반품일시 | 연수과정 | 기수 | 개강일 | ISBN | 제목 | 판정보 | 권수 | 입력방식 | QR원문
 *  - 미등록반품 : 반품일시 | 연수과정 | 기수 | 개강일 | ISBN | 권수 | 입력방식 | QR원문
 *               (도서목록에 없는 ISBN을 스캔한 반품은 이 시트에 따로 기록됩니다)
 */

var SHEET_BOOKS = '도서목록';
var SHEET_RETURNS = '반품기록';
var SHEET_UNREGISTERED = '미등록반품';

/** 웹앱 진입점 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('도서 반품 키오스크')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 최초 1회 실행: 필요한 시트와 헤더를 생성합니다.
 * Apps Script 편집기에서 이 함수를 선택하고 ▶ 실행하세요.
 */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var books = ss.getSheetByName(SHEET_BOOKS);
  if (!books) {
    books = ss.insertSheet(SHEET_BOOKS);
    books.getRange(1, 1, 1, 5)
      .setValues([['ISBN', '제목', '판정보', '발간일', '출판사']])
      .setFontWeight('bold');
    books.setFrozenRows(1);
    // ISBN이 숫자로 변환되어 앞자리가 깨지지 않도록 텍스트 형식 지정
    books.getRange('A:A').setNumberFormat('@');
    books.setColumnWidth(1, 140);
    books.setColumnWidth(2, 320);
  }

  var rets = ss.getSheetByName(SHEET_RETURNS);
  if (!rets) {
    rets = ss.insertSheet(SHEET_RETURNS);
    rets.getRange(1, 1, 1, 10)
      .setValues([['반품일시', '연수과정', '기수', '개강일', 'ISBN', '제목', '판정보', '권수', '입력방식', 'QR원문']])
      .setFontWeight('bold');
    rets.setFrozenRows(1);
    rets.getRange('E:E').setNumberFormat('@');
    rets.setColumnWidth(1, 150);
    rets.setColumnWidth(2, 220);
    rets.setColumnWidth(6, 320);
  }

  ensureUnregisteredSheet_(ss);
}

/** 미등록반품 시트가 없으면 만들어서 돌려줍니다. */
function ensureUnregisteredSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_UNREGISTERED);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_UNREGISTERED);
    sheet.getRange(1, 1, 1, 8)
      .setValues([['반품일시', '연수과정', '기수', '개강일', 'ISBN', '권수', '입력방식', 'QR원문']])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('E:E').setNumberFormat('@');
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(5, 140);
  }
  return sheet;
}

/** 바코드에서 읽힌 문자열을 ISBN 비교용으로 정규화 (숫자와 X만 남김) */
function normalizeIsbn_(raw) {
  return String(raw || '').toUpperCase().replace(/[^0-9X]/g, '');
}

/**
 * 도서목록 시트에서 ISBN으로 도서를 조회합니다.
 * @return {Object} {found, isbn, title, edition, pubDate, publisher}
 */
function getBookByIsbn(isbn) {
  var key = normalizeIsbn_(isbn);
  if (!key) return { found: false, isbn: '' };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BOOKS);
  if (sheet && sheet.getLastRow() > 1) {
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (normalizeIsbn_(rows[i][0]) === key) {
        var pubDate = rows[i][3];
        if (pubDate instanceof Date) {
          pubDate = Utilities.formatDate(pubDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        return {
          found: true,
          isbn: key,
          title: String(rows[i][1] || ''),
          edition: String(rows[i][2] || ''),
          pubDate: String(pubDate || ''),
          publisher: String(rows[i][4] || '')
        };
      }
    }
  }
  return { found: false, isbn: key };
}

/**
 * 도서목록 시트에서 도서명으로 검색합니다. (부분 일치, 대소문자/공백 무시)
 * @return {Object[]} [{isbn, title, edition, pubDate, publisher}] 최대 20건
 */
function searchBooks(query) {
  var q = String(query || '').replace(/\s+/g, '').toLowerCase();
  if (!q) return [];

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BOOKS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var results = [];
  for (var i = 0; i < rows.length && results.length < 20; i++) {
    var title = String(rows[i][1] || '');
    if (title.replace(/\s+/g, '').toLowerCase().indexOf(q) === -1) continue;
    var pubDate = rows[i][3];
    if (pubDate instanceof Date) {
      pubDate = Utilities.formatDate(pubDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    results.push({
      isbn: normalizeIsbn_(rows[i][0]),
      title: title,
      edition: String(rows[i][2] || ''),
      pubDate: String(pubDate || ''),
      publisher: String(rows[i][4] || '')
    });
  }
  return results;
}

/**
 * 반품 내역을 기록합니다.
 *  - 도서목록에 있는 도서   → 반품기록 시트
 *  - 도서목록에 없는 도서   → 미등록반품 시트 (ISBN + 반품정보만 기록)
 * @param {Object} payload {course, cohort, startDate, qrRaw, mode, items:[{isbn,title,edition,qty,found}]}
 * @return {Object} {ok, count, unregistered}
 */
function submitReturn(payload) {
  if (!payload || !payload.items || !payload.items.length) {
    throw new Error('반품할 도서가 없습니다.');
  }
  var course = String(payload.course || '').trim();
  var cohort = String(payload.cohort || '').trim();
  var startDate = String(payload.startDate || '').trim();
  if (!course) throw new Error('연수 과정 정보가 없습니다.');

  var modeLabel = payload.mode === 'scan' ? '바코드 반복 스캔' : '수기 입력';
  var qrRaw = String(payload.qrRaw || '');
  var now = new Date();

  var registeredRows = [];
  var unregisteredRows = [];
  payload.items.forEach(function (it) {
    var qty = Math.max(1, parseInt(it.qty, 10) || 1);
    var isbn = normalizeIsbn_(it.isbn);
    if (it.found) {
      registeredRows.push([now, course, cohort, startDate, isbn, String(it.title || ''), String(it.edition || ''), qty, modeLabel, qrRaw]);
    } else {
      unregisteredRows.push([now, course, cohort, startDate, isbn, qty, modeLabel, qrRaw]);
    }
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (registeredRows.length) {
      var sheet = ss.getSheetByName(SHEET_RETURNS);
      if (!sheet) throw new Error('"' + SHEET_RETURNS + '" 시트가 없습니다. setupSheets를 먼저 실행하세요.');
      sheet.getRange(sheet.getLastRow() + 1, 1, registeredRows.length, registeredRows[0].length).setValues(registeredRows);
    }
    if (unregisteredRows.length) {
      var unregSheet = ensureUnregisteredSheet_(ss);
      unregSheet.getRange(unregSheet.getLastRow() + 1, 1, unregisteredRows.length, unregisteredRows[0].length).setValues(unregisteredRows);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { ok: true, count: registeredRows.length + unregisteredRows.length, unregistered: unregisteredRows.length };
}
