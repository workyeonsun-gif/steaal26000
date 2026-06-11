/**
 * 도서 반품 키오스크 — 서버 코드 (Google Apps Script)
 *
 * 시트 구성
 *  - 도서목록   : ISBN | 제목 | 판정보(초판, 개정 2판 등) | 발간일 | 출판사   (도서 마스터, 관리자가 채움)
 *  - 반품기록   : 반품일시 | 연수과정 | 기수 | 개강일 | ISBN | 제목 | 판정보 | 권수 | 입력방식 | QR원문 | 부서 | 입력자
 *  - 미등록반품 : 반품일시 | 연수과정 | 기수 | 개강일 | ISBN | 제목 | 권수 | 입력방식 | QR원문 | 부서 | 입력자
 *               (도서목록에 없는 ISBN 스캔분과 사용자가 직접 입력한 제목의 반품이 이 시트에 기록됩니다)
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
    rets.getRange(1, 1, 1, 12)
      .setValues([['반품일시', '연수과정', '기수', '개강일', 'ISBN', '제목', '판정보', '권수', '입력방식', 'QR원문', '부서', '입력자']])
      .setFontWeight('bold');
    rets.setFrozenRows(1);
    rets.getRange('E:E').setNumberFormat('@');
    rets.setColumnWidth(1, 150);
    rets.setColumnWidth(2, 220);
    rets.setColumnWidth(6, 320);
  }

  ensureUnregisteredSheet_(ss);
  ensureMetaColumns_(ss);
}

/* 수정 추적용 메타 열: 수정발생 | 수정일 | 반품처리완료(체크박스) */
var COL_MODIFIED = '수정발생';
var COL_MODIFIED_AT = '수정일';
var COL_DONE = '반품처리완료';

/** 반품기록·미등록반품 시트에 메타 열이 없으면 추가합니다. ('선택' 체크박스 열 앞에 삽입) */
function ensureMetaColumns_(ss) {
  [SHEET_RETURNS, SHEET_UNREGISTERED].forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    [COL_MODIFIED, COL_MODIFIED_AT, COL_DONE].forEach(function (h) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
      if (headers.indexOf(h) !== -1) return;
      var selIdx = headers.indexOf('선택');
      var col;
      if (selIdx !== -1) {
        sheet.insertColumnBefore(selIdx + 1);
        col = selIdx + 1;
      } else {
        col = sheet.getLastColumn() + 1;
      }
      sheet.getRange(1, col).setValue(h).setFontWeight('bold');
    });
    var headers2 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    var doneCol = headers2.indexOf(COL_DONE) + 1;
    if (doneCol > 0) {
      var lastData = lastDataRow_(sheet);
      if (lastData > 1) sheet.getRange(2, doneCol, lastData - 1, 1).insertCheckboxes();
      // 데이터가 없는 아래 행에 남은 FALSE 값을 지워 getLastRow가 부풀지 않게 함
      if (sheet.getMaxRows() > lastData) {
        sheet.getRange(lastData + 1, doneCol, sheet.getMaxRows() - lastData, 1).clearContent();
      }
    }
  });
}

/** 미등록반품 시트가 없으면 만들어서 돌려줍니다. */
function ensureUnregisteredSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_UNREGISTERED);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_UNREGISTERED);
    sheet.getRange(1, 1, 1, 11)
      .setValues([['반품일시', '연수과정', '기수', '개강일', 'ISBN', '제목', '권수', '입력방식', 'QR원문', '부서', '입력자']])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('E:E').setNumberFormat('@');
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(5, 140);
    sheet.setColumnWidth(6, 320);
  }
  return sheet;
}

/** 바코드에서 읽힌 문자열을 ISBN 비교용으로 정규화 (숫자와 X만 남김) */
function normalizeIsbn_(raw) {
  return String(raw || '').toUpperCase().replace(/[^0-9X]/g, '');
}

/**
 * 실제 데이터의 마지막 행 번호.
 * 체크박스(선택/반품처리완료) 열의 FALSE 값으로 부풀려진 getLastRow() 대신,
 * 체크박스가 없는 A~I열의 공란 여부로 판단합니다. A~I 중 하나라도 값이 있으면
 * 데이터 행으로 보며, 새 기록은 항상 그 바로 다음 행에 누적 기록됩니다.
 */
function lastDataRow_(sheet) {
  var cols = Math.min(9, sheet.getMaxColumns());   // A~I열만 검사 (체크박스 열 제외)
  var vals = sheet.getRange(1, 1, sheet.getMaxRows(), cols).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    for (var j = 0; j < cols; j++) {
      var v = vals[i][j];
      if (v !== '' && v !== null && v !== false) return i + 1;
    }
  }
  return 1;
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

/** PDF 내역서용 HTML 이스케이프 */
function escapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** 확인 화면과 같은 내용의 반품 내역서 HTML 생성 */
function buildReceiptHtml_(payload) {
  var e = escapeHtml_;
  var rows = '', total = 0;
  (payload.items || []).forEach(function (it) {
    var qty = Math.max(1, parseInt(it.qty, 10) || 1);
    total += qty;
    rows += '<tr><td>' + (it.found ? '' : '[미등록] ') + e(it.title) + '</td>' +
      '<td>' + e(it.edition || '-') + '</td><td>' + e(it.isbn || '-') + '</td>' +
      '<td class="n">' + qty + '</td></tr>';
  });
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:"Malgun Gothic","Noto Sans KR",sans-serif;padding:24px;color:#111}' +
    'h1{font-size:20px;text-align:center;margin-bottom:18px}' +
    '.meta{font-size:12px;margin-bottom:14px}.meta div{margin:2px 0}' +
    'table{width:100%;border-collapse:collapse;font-size:12px}' +
    'th,td{border:1px solid #999;padding:6px 8px;text-align:left}' +
    'th{background:#eee}.n{text-align:right}' +
    '</style></head><body>' +
    '<h1>도서 반품 내역서 (연수과정)</h1>' +
    '<div class="meta">' +
    '<div>연수과정: ' + e(payload.course) + '</div>' +
    '<div>기수: ' + e(payload.cohort || '-') + '</div>' +
    '<div>개강일: ' + e(payload.startDate || '-') + '</div>' +
    '<div>부서: ' + e(payload.dept || '-') + '</div>' +
    '<div>출력일시: ' + now + '</div></div>' +
    '<table><tr><th>제목</th><th>판정보</th><th>ISBN</th><th>권수</th></tr>' +
    rows +
    '<tr><th colspan="3">합계</th><th class="n">' + total + '</th></tr>' +
    '</table></body></html>';
}

/** 내역서 HTML을 PDF Blob으로 변환 */
function createReceiptPdf_(payload) {
  if (!payload || !payload.items || !payload.items.length) throw new Error('PDF로 만들 내역이 없습니다.');
  var pdf = Utilities.newBlob(buildReceiptHtml_(payload), MimeType.HTML, 'receipt.html').getAs(MimeType.PDF);
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  var who = String(payload.course || '반품').replace(/[\\\/:*?"<>|\s]/g, '').slice(0, 20) || '반품';
  pdf.setName('반품내역_' + who + '_' + stamp + '.pdf');
  return pdf;
}

/** 클라이언트 다운로드용: PDF를 base64로 반환 */
function getReceiptPdf(payload) {
  var pdf = createReceiptPdf_(payload);
  return { name: pdf.getName(), base64: Utilities.base64Encode(pdf.getBytes()) };
}

/** 반품 내역서 PDF를 이메일로 전송 */
function emailReceipt(payload, email) {
  email = String(email || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('올바른 이메일 주소가 아닙니다.');
  var pdf = createReceiptPdf_(payload);
  MailApp.sendEmail({
    to: email,
    subject: '[도서 반품] 반품 내역서 — ' + String(payload.course || ''),
    body: '도서 반품 내역서를 첨부합니다.\n\n연수과정: ' + String(payload.course || '') +
      '\n기수: ' + String(payload.cohort || '-') +
      '\n개강일: ' + String(payload.startDate || '-'),
    attachments: [pdf]
  });
  return { ok: true };
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
  var dept = String(payload.dept || '').trim();            // 부서 (진입화면에서 선택)
  var operator = String(payload.operator || '').trim();   // 입력자 이름 (선택값)
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
      registeredRows.push([now, course, cohort, startDate, isbn, String(it.title || ''), String(it.edition || ''), qty, modeLabel, qrRaw, dept, operator]);
    } else {
      unregisteredRows.push([now, course, cohort, startDate, isbn, String(it.title || ''), qty, modeLabel, qrRaw, dept, operator]);
    }
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (registeredRows.length) {
      var sheet = ss.getSheetByName(SHEET_RETURNS);
      if (!sheet) throw new Error('"' + SHEET_RETURNS + '" 시트가 없습니다. setupSheets를 먼저 실행하세요.');
      sheet.getRange(lastDataRow_(sheet) + 1, 1, registeredRows.length, registeredRows[0].length).setValues(registeredRows);
    }
    if (unregisteredRows.length) {
      var unregSheet = ensureUnregisteredSheet_(ss);
      unregSheet.getRange(lastDataRow_(unregSheet) + 1, 1, unregisteredRows.length, unregisteredRows[0].length).setValues(unregisteredRows);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { ok: true, count: registeredRows.length + unregisteredRows.length, unregistered: unregisteredRows.length };
}

/* ─────────── 스프레드시트 상단 메뉴: 선택한 행 PDF 다운로드 ─────────── */

var CHECKBOX_HEADER = '선택';

/** 스프레드시트를 열면 상단에 커스텀 메뉴를 추가합니다. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 반품 키오스크')
    .addItem('☑️ 선택 체크박스 추가/정비', 'ensureCheckboxes')
    .addSeparator()
    .addItem('✅ 전체 선택 (현재 시트)', 'selectAllRows')
    .addItem('⬜ 전체 선택 해제 (현재 시트)', 'deselectAllRows')
    .addSeparator()
    .addItem('📄 선택한 행 PDF 다운로드', 'downloadSelectedPdf')
    .addSeparator()
    .addItem('🔎 날짜로 조회/수정 (15일 이내)', 'openManageDialog')
    .addToUi();
}

function pdfTargetSheets_() {
  return [SHEET_RETURNS, SHEET_UNREGISTERED];
}

/** 시트에서 '선택' 체크박스 열 번호를 찾습니다. 없으면 0 */
function checkboxCol_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return 0;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === CHECKBOX_HEADER) return i + 1;
  }
  return 0;
}

/** 반품기록·미등록반품 시트 맨 끝에 '선택' 체크박스 열을 만들고 정비합니다. */
function ensureCheckboxes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  pdfTargetSheets_().forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var col = checkboxCol_(sheet);
    if (!col) {
      col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(CHECKBOX_HEADER).setFontWeight('bold');
    }
    var lastData = lastDataRow_(sheet);
    if (lastData > 1) sheet.getRange(2, col, lastData - 1, 1).insertCheckboxes();
    // 데이터가 없는 아래 행에 남은 FALSE 값을 지워 getLastRow가 부풀지 않게 함
    if (sheet.getMaxRows() > lastData) {
      sheet.getRange(lastData + 1, col, sheet.getMaxRows() - lastData, 1).clearContent();
    }
  });
  ss.toast('반품기록·미등록반품 시트에 선택 체크박스를 준비했습니다.', '☑️ 완료', 5);
}

function setAllChecks_(checked) {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (pdfTargetSheets_().indexOf(sheet.getName()) === -1) {
    SpreadsheetApp.getUi().alert('⚠️ "' + SHEET_RETURNS + '" 또는 "' + SHEET_UNREGISTERED + '" 시트에서 실행해주세요.');
    return;
  }
  var col = checkboxCol_(sheet);
  if (!col) { ensureCheckboxes(); col = checkboxCol_(sheet); }
  var lastRow = lastDataRow_(sheet);
  if (lastRow > 1) {
    sheet.getRange(2, col, lastRow - 1, 1).insertCheckboxes().setValue(checked);
  }
  SpreadsheetApp.getActiveSpreadsheet()
    .toast(sheet.getName() + ' 시트 ' + (lastRow - 1) + '행을 ' + (checked ? '전체 선택' : '전체 해제') + '했습니다.',
      checked ? '✅ 전체 선택' : '⬜ 전체 해제', 5);
}
function selectAllRows() { setAllChecks_(true); }
function deselectAllRows() { setAllChecks_(false); }

/** 체크된 행의 표시값(서식 적용된 문자열) 배열을 반환합니다. */
function selectedRows_(name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) return [];
  var col = checkboxCol_(sheet);
  if (!col) return [];
  var lastRow = lastDataRow_(sheet);
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, col).getDisplayValues();
  var checks = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    if (checks[i][0] === true) out.push(data[i]);
  }
  return out;
}

/** 키오스크 확인 화면 스타일의 PDF용 HTML */
function buildSelectedPdfHtml_(rets, unreg) {
  var e = escapeHtml_;
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:"Malgun Gothic","Noto Sans KR",sans-serif;background:#f1f5f9;padding:28px;color:#0f172a}' +
    'h1{font-size:22px;text-align:center;margin:0 0 6px}' +
    '.sub{text-align:center;color:#64748b;font-size:11px;margin-bottom:18px}' +
    '.card{background:#fff;border-radius:14px;padding:18px 20px;margin-bottom:18px;border:1px solid #e2e8f0}' +
    'h2{font-size:14px;margin:0 0 10px}' +
    'table{width:100%;border-collapse:collapse;font-size:11px}' +
    'th{color:#64748b;font-weight:700;text-align:left;border-bottom:2px solid #e2e8f0;padding:7px 6px}' +
    'td{border-bottom:1px solid #e2e8f0;padding:7px 6px}' +
    '.n{text-align:right;font-weight:700}' +
    '.total td{font-weight:800;border-bottom:none}' +
    '</style></head><body>' +
    '<h1>📚 도서 반품 내역 (연수과정)</h1>' +
    '<div class="sub">생성일시: ' + now + '</div>';

  var grand = 0;
  if (rets.length) {
    var sum = 0;
    html += '<div class="card"><h2>반품기록 — ' + rets.length + '건</h2>' +
      '<table><tr><th>반품일시</th><th>연수과정</th><th>기수</th><th>개강일</th><th>제목</th><th>판정보</th><th>ISBN</th><th style="text-align:right">권수</th></tr>';
    rets.forEach(function (r) {
      var qty = parseInt(r[7], 10) || 0; sum += qty;
      html += '<tr><td>' + e(r[0]) + '</td><td>' + e(r[1]) + '</td><td>' + e(r[2]) + '</td><td>' + e(r[3]) + '</td>' +
        '<td>' + e(r[5]) + '</td><td>' + e(r[6] || '-') + '</td><td>' + e(r[4]) + '</td>' +
        '<td class="n">' + e(r[7]) + '</td></tr>';
    });
    html += '<tr class="total"><td colspan="7">합계</td><td class="n">' + sum + '</td></tr></table></div>';
    grand += sum;
  }
  if (unreg.length) {
    var sum2 = 0;
    html += '<div class="card"><h2>미등록반품 — ' + unreg.length + '건</h2>' +
      '<table><tr><th>반품일시</th><th>연수과정</th><th>기수</th><th>개강일</th><th>제목</th><th>ISBN</th><th style="text-align:right">권수</th></tr>';
    unreg.forEach(function (r) {
      var qty = parseInt(r[6], 10) || 0; sum2 += qty;
      html += '<tr><td>' + e(r[0]) + '</td><td>' + e(r[1]) + '</td><td>' + e(r[2]) + '</td><td>' + e(r[3]) + '</td>' +
        '<td>[미등록] ' + e(r[5]) + '</td><td>' + e(r[4] || '-') + '</td>' +
        '<td class="n">' + e(r[6]) + '</td></tr>';
    });
    html += '<tr class="total"><td colspan="6">합계</td><td class="n">' + sum2 + '</td></tr></table></div>';
    grand += sum2;
  }
  html += '<div class="sub">총 ' + grand + '권</div></body></html>';
  return html;
}

/** 메뉴: 체크된 행만 모아 키오스크 스타일 PDF로 다운로드 */
function downloadSelectedPdf() {
  var rets = selectedRows_(SHEET_RETURNS);
  var unreg = selectedRows_(SHEET_UNREGISTERED);
  if (!rets.length && !unreg.length) {
    SpreadsheetApp.getUi().alert('⚠️ 선택된 행이 없습니다.\n\n"☑️ 선택 체크박스 추가/정비"로 체크박스를 만든 뒤, PDF로 만들 행을 체크하고 다시 실행해주세요.');
    return;
  }
  var pdf = Utilities.newBlob(buildSelectedPdfHtml_(rets, unreg), MimeType.HTML, 'r.html').getAs(MimeType.PDF);
  var name = '반품내역_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.pdf';
  pdf.setName(name);
  var b64 = Utilities.base64Encode(pdf.getBytes());
  var dlg = HtmlService.createHtmlOutput(
    '<style>body{font-family:sans-serif;text-align:center;padding:18px}' +
    'a{display:inline-block;background:#2563eb;color:#fff;padding:12px 26px;border-radius:10px;text-decoration:none;font-weight:bold}</style>' +
    '<p>📄 ' + name + '</p>' +
    '<p style="color:#64748b;font-size:13px">자동으로 다운로드되지 않으면 버튼을 눌러주세요</p>' +
    '<a id="dl" download="' + name + '" href="data:application/pdf;base64,' + b64 + '">⬇️ 다운로드</a>' +
    '<scr' + 'ipt>document.getElementById("dl").click();</scr' + 'ipt>'
  ).setWidth(380).setHeight(190);
  SpreadsheetApp.getUi().showModalDialog(dlg, '📄 PDF 다운로드');
}

/* ─────────── 날짜로 조회/수정 (수정이력 추적, 15일 제한) ─────────── */

var SHEET_AUDIT = '수정이력';
var EDIT_LIMIT_DAYS = 15;   // 입력일로부터 이 기간이 지나면 수정 불가

/** 시트별 수정 허용 항목 (그 외 열은 화면에서 수정 불가) */
function editableFields_() {
  return {
    '반품기록': ['연수과정', '기수', '개강일', '부서', '권수', '입력자'],
    '미등록반품': ['연수과정', '기수', '개강일', '부서', '제목', '권수', '입력자']
  };
}

function fieldType_(name) {
  if (name === '권수') return 'number';
  if (name === '상태') return 'status';
  return 'text';
}

/** 수정이력 시트가 없으면 생성 */
function ensureAuditSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_AUDIT);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_AUDIT);
    sheet.getRange(1, 1, 1, 7)
      .setValues([['수정일시', '시트', '행', '항목', '변경 전', '변경 후', '수정자']])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
  }
  return sheet;
}

/** 메뉴: 조회/수정 대화상자 열기 */
function openManageDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Manage').setWidth(940).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, '🔎 반품 기록 조회/수정');
}

var SEARCH_RANGE_DAYS = 30;   // 날짜 없이 검색할 때 조회 범위 (검색일 기준)
var SEARCH_MAX_RESULTS = 200;

/**
 * 히스토리 검색: 날짜 / 연수과정명 / 도서명 / 입력자 조합으로 조회합니다.
 * 날짜를 지정하지 않으면 검색일 기준 최근 SEARCH_RANGE_DAYS일 이내로 제한됩니다.
 * @param {Object} criteria {date:'yyyy-MM-dd'|'', course:'', title:'', operator:''}
 */
function searchRecords(criteria) {
  criteria = criteria || {};
  var dateStr = String(criteria.date || '').trim();
  var hasDate = /^\d{4}-\d{2}-\d{2}$/.test(dateStr);
  var norm = function (s) { return String(s || '').replace(/\s+/g, '').toLowerCase(); };
  var qCourse = norm(criteria.course);
  var qTitle = norm(criteria.title);
  var qOper = norm(criteria.operator);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = Session.getScriptTimeZone();
  var fieldsBySheet = editableFields_();
  var minTs = Date.now() - SEARCH_RANGE_DAYS * 86400000;
  var records = [];
  var truncated = false;

  pdfTargetSheets_().forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var lastDataR = lastDataRow_(sheet);
    if (lastDataR < 2) return;
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    var courseIdx = headers.indexOf('연수과정');
    var titleIdx = headers.indexOf('제목');
    var operIdx = headers.indexOf('입력자');
    var isbnIdx = headers.indexOf('ISBN');
    var doneIdx = headers.indexOf(COL_DONE);
    var modIdx = headers.indexOf(COL_MODIFIED);
    var modAtIdx = headers.indexOf(COL_MODIFIED_AT);
    var values = sheet.getRange(2, 1, lastDataR - 1, lastCol).getValues();
    var disp = sheet.getRange(2, 1, lastDataR - 1, lastCol).getDisplayValues();
    for (var i = 0; i < values.length; i++) {
      var ts = values[i][0];
      if (!(ts instanceof Date)) continue;
      // 날짜 지정 시 그 날짜만, 미지정 시 최근 30일 이내만
      if (hasDate) {
        if (Utilities.formatDate(ts, tz, 'yyyy-MM-dd') !== dateStr) continue;
      } else if (ts.getTime() < minTs) {
        continue;
      }
      if (qCourse && (courseIdx === -1 || norm(disp[i][courseIdx]).indexOf(qCourse) === -1)) continue;
      if (qTitle && (titleIdx === -1 || norm(disp[i][titleIdx]).indexOf(qTitle) === -1)) continue;
      if (qOper && (operIdx === -1 || norm(disp[i][operIdx]).indexOf(qOper) === -1)) continue;
      if (records.length >= SEARCH_MAX_RESULTS) { truncated = true; break; }
      var within = (Date.now() - ts.getTime()) / 86400000 <= EDIT_LIMIT_DAYS;
      var done = doneIdx > -1 && values[i][doneIdx] === true;   // 반품처리완료 체크 시 기간 무관 수정 불가
      records.push({
        sheet: name,
        row: i + 2,
        ts: Utilities.formatDate(ts, tz, 'yyyy-MM-dd HH:mm'),
        summary: (titleIdx > -1 ? disp[i][titleIdx] : '') +
          (isbnIdx > -1 && disp[i][isbnIdx] ? ' · ISBN ' + disp[i][isbnIdx] : ''),
        editable: within && !done,
        done: done,
        lockReason: done ? '반품 처리 완료' : (!within ? '입력일로부터 ' + EDIT_LIMIT_DAYS + '일 경과' : ''),
        modified: modIdx > -1 ? disp[i][modIdx] : '',
        modifiedAt: modAtIdx > -1 ? disp[i][modAtIdx] : '',
        fields: (fieldsBySheet[name] || []).map(function (fn) {
          var ci = headers.indexOf(fn);
          return { name: fn, value: ci === -1 ? '' : disp[i][ci], type: fieldType_(fn) };
        })
      });
    }
  });

  var rangeNote = hasDate
    ? dateStr + ' (해당 날짜의 입력 기록)'
    : Utilities.formatDate(new Date(minTs), tz, 'yyyy-MM-dd') + ' ~ 오늘 (최근 ' + SEARCH_RANGE_DAYS + '일)';
  return {
    limitDays: EDIT_LIMIT_DAYS,
    rangeDays: SEARCH_RANGE_DAYS,
    rangeNote: rangeNote,
    truncated: truncated,
    records: records
  };
}

/**
 * 수정 사항을 저장하고 변경 전/후 값을 수정이력 시트에 남깁니다.
 * @param {Object[]} edits [{sheet, row, changes:{항목명: 새값}}]
 * @return {Object} {updated, blocked:[설명]}
 */
function saveRecordEdits(edits) {
  if (!edits || !edits.length) return { updated: 0, blocked: [] };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var user = '';
  try { user = Session.getActiveUser().getEmail() || ''; } catch (e) { /* 권한에 따라 빈 값 */ }
  var fieldsBySheet = editableFields_();
  var updated = 0, blocked = [];

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureMetaColumns_(ss);
    var audit = ensureAuditSheet_(ss);
    edits.forEach(function (ed) {
      var sheet = ss.getSheetByName(ed.sheet);
      var allowed = fieldsBySheet[ed.sheet];
      if (!sheet || !allowed) return;
      var lastCol = sheet.getLastColumn();
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
      var ts = sheet.getRange(ed.row, 1).getValue();
      // 서버에서도 잠금 조건을 다시 검증 (화면 우회 방지)
      if (!(ts instanceof Date) || (Date.now() - ts.getTime()) / 86400000 > EDIT_LIMIT_DAYS) {
        blocked.push(ed.sheet + ' ' + ed.row + '행 — 입력일로부터 ' + EDIT_LIMIT_DAYS + '일이 지나 수정할 수 없습니다');
        return;
      }
      var doneIdx = headers.indexOf(COL_DONE);
      if (doneIdx > -1 && sheet.getRange(ed.row, doneIdx + 1).getValue() === true) {
        blocked.push(ed.sheet + ' ' + ed.row + '행 — 반품 처리가 완료된 내역은 수정할 수 없습니다');
        return;
      }
      // 조회 후 시트가 정렬/삭제되어 행 번호가 어긋났으면 엉뚱한 행 수정을 차단
      if (ed.ts) {
        var tsStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
        if (ed.ts !== tsStr) {
          blocked.push(ed.sheet + ' ' + ed.row + '행 — 시트 내용이 변경되어 수정하지 않았습니다. 다시 조회한 뒤 시도해주세요');
          return;
        }
      }
      var now = new Date();
      var rowUpdated = 0;
      Object.keys(ed.changes || {}).forEach(function (fn) {
        if (allowed.indexOf(fn) === -1) return;
        var ci = headers.indexOf(fn);
        if (ci === -1) return;
        var cell = sheet.getRange(ed.row, ci + 1);
        var oldVal = String(cell.getDisplayValue());
        var newVal = String(ed.changes[fn]).trim();
        if (fn === '권수') {
          var n = parseInt(newVal, 10);
          if (!n || n < 1) { blocked.push(ed.sheet + ' ' + ed.row + '행 권수 — 1 이상이어야 합니다'); return; }
          newVal = String(n);
        }
        if (oldVal === newVal) return;
        cell.setValue(fn === '권수' ? parseInt(newVal, 10) : newVal);
        audit.appendRow([now, ed.sheet, ed.row, fn, oldVal, newVal, user]);
        rowUpdated++;
        updated++;
      });
      // 수정이 발생한 행에는 수정발생/수정일 표시
      if (rowUpdated) {
        var modIdx = headers.indexOf(COL_MODIFIED);
        var modAtIdx = headers.indexOf(COL_MODIFIED_AT);
        if (modIdx > -1) sheet.getRange(ed.row, modIdx + 1).setValue('수정됨');
        if (modAtIdx > -1) sheet.getRange(ed.row, modAtIdx + 1).setValue(now);
      }
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { updated: updated, blocked: blocked };
}

/* ─────────── 키오스크 조회 화면: 선택한 행 PDF 다운로드/이메일 ─────────── */

/** 조회 화면에서 선택한 행 참조를 시트별 표시값 행으로 변환 */
function rowsByRef_(refs) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = {};
  out[SHEET_RETURNS] = [];
  out[SHEET_UNREGISTERED] = [];
  (refs || []).forEach(function (ref) {
    var name = String(ref.sheet || '');
    if (!(name in out)) return;
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    var row = parseInt(ref.row, 10);
    if (!row || row < 2 || row > lastDataRow_(sheet)) return;
    out[name].push(sheet.getRange(row, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]);
  });
  return out;
}

function buildHistoryPdf_(refs) {
  var grouped = rowsByRef_(refs);
  var rets = grouped[SHEET_RETURNS], unreg = grouped[SHEET_UNREGISTERED];
  if (!rets.length && !unreg.length) throw new Error('PDF로 만들 내역이 없습니다. 행을 다시 선택해주세요.');
  var pdf = Utilities.newBlob(buildSelectedPdfHtml_(rets, unreg), MimeType.HTML, 'r.html').getAs(MimeType.PDF);
  pdf.setName('반품내역_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.pdf');
  return pdf;
}

/** 조회 화면: 선택한 행 PDF를 base64로 반환 (다운로드용) */
function getHistoryPdf(refs) {
  var pdf = buildHistoryPdf_(refs);
  return { name: pdf.getName(), base64: Utilities.base64Encode(pdf.getBytes()) };
}

/** 조회 화면: 선택한 행 PDF를 이메일로 전송 */
function emailHistoryPdf(refs, email) {
  email = String(email || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('올바른 이메일 주소가 아닙니다.');
  var pdf = buildHistoryPdf_(refs);
  MailApp.sendEmail({
    to: email,
    subject: '[도서 반품] 반품 내역 — ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    body: '선택한 반품 내역을 PDF로 첨부합니다.',
    attachments: [pdf]
  });
  return { ok: true };
}


/* ─────────── 키오스크 조회 화면: 미확정 내역 삭제 ─────────── */

/**
 * 반품처리완료(확정)되지 않은 내역 한 건을 삭제합니다.
 * 삭제 전 행 전체 내용을 수정이력 시트에 스냅숏으로 남깁니다.
 * @param {Object} ref {sheet, row, ts}
 * @return {Object} {ok, reason}
 */
function deleteRecord(ref) {
  if (!ref || !ref.sheet || !ref.row) throw new Error('삭제 대상이 올바르지 않습니다.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if ([SHEET_RETURNS, SHEET_UNREGISTERED].indexOf(String(ref.sheet)) === -1) {
      throw new Error('잘못된 시트입니다.');
    }
    var sheet = ss.getSheetByName(String(ref.sheet));
    if (!sheet) throw new Error('시트를 찾을 수 없습니다.');
    var row = parseInt(ref.row, 10);
    if (!row || row < 2 || row > lastDataRow_(sheet)) {
      return { ok: false, reason: '행을 찾을 수 없습니다. 다시 조회한 뒤 시도해주세요.' };
    }
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    var ts = sheet.getRange(row, 1).getValue();
    // 조회 후 시트가 정렬/변경되어 행이 어긋났으면 엉뚱한 행 삭제를 차단
    if (!(ts instanceof Date) ||
        (ref.ts && Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') !== String(ref.ts))) {
      return { ok: false, reason: '시트 내용이 변경되어 삭제하지 않았습니다. 다시 조회한 뒤 시도해주세요.' };
    }
    // 반품처리완료(확정)된 내역은 삭제 불가
    var doneIdx = headers.indexOf(COL_DONE);
    if (doneIdx > -1 && sheet.getRange(row, doneIdx + 1).getValue() === true) {
      return { ok: false, reason: '반품 처리가 완료된 내역은 삭제할 수 없습니다.' };
    }
    var user = '';
    try { user = Session.getActiveUser().getEmail() || ''; } catch (e) { /* 권한에 따라 빈 값 */ }
    // 삭제 전 행 전체 내용을 수정이력에 보존 (복구·추적용)
    var rowVals = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
    var snapshot = headers.map(function (h, i) {
      return rowVals[i] !== '' ? h + '=' + rowVals[i] : '';
    }).filter(Boolean).join(' | ');
    ensureAuditSheet_(ss).appendRow([new Date(), sheet.getName(), row, '행 삭제', snapshot, '(삭제됨)', user]);
    sheet.deleteRow(row);
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
