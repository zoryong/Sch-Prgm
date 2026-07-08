/**
 * 배드민턴 모임 참석 신청 Google Apps Script (API 전용)
 *
 * [배포 방법 - index.html 방식]
 * 1. 스프레드시트 > 확장 프로그램 > Apps Script
 * 2. 이 Code.gs 내용을 붙여넣기
 * 3. 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스: 모든 사용자)
 * 4. 배포 URL을 index.html 의 APPS_SCRIPT_URL 에 입력
 *
 * 스프레드시트 1행 헤더: 제출시간 | 성명
 */

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = (params.action || '').trim();
  var name = (params.name || '').trim();
  var callback = sanitizeCallback_(params.callback);

  if (action === 'ping') {
    return respond_(callback, {
      ok: true,
      message: 'API ready'
    });
  }

  if (!name) {
    if (callback) {
      return jsonpResponse(callback, {
        ok: false,
        message: '성명을 입력해 주세요.'
      });
    }

    return jsonResponse({
      ok: true,
      message: '배드민턴 모임 참석 신청 API'
    });
  }

  try {
    return respond_(callback, processSubmission(name));
  } catch (error) {
    return respond_(callback, {
      ok: false,
      message: '오류가 발생했습니다. 다시 시도해 주세요.'
    });
  }
}

function doPost(e) {
  try {
    var name = (e.parameter.name || '').trim();
    return jsonResponse(processSubmission(name));
  } catch (error) {
    return jsonResponse({
      ok: false,
      message: '오류가 발생했습니다. 다시 시도해 주세요.'
    });
  }
}

function respond_(callback, payload) {
  if (callback) {
    return jsonpResponse(callback, payload);
  }

  return jsonResponse(payload);
}

function processSubmission(name) {
  if (!name) {
    return {
      ok: false,
      message: '성명을 입력해 주세요.'
    };
  }

  var sheet = getOrCreateSheet();
  var rows = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    var existingName = String(rows[i][1] || '').trim();
    if (existingName === name) {
      return {
        ok: false,
        message: '이미 신청하셨습니다.'
      };
    }
  }

  sheet.appendRow([new Date(), name]);

  return {
    ok: true,
    message: '참석 신청이 완료되었습니다.'
  };
}

function getOrCreateSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('참석신청');

  if (!sheet) {
    sheet = spreadsheet.insertSheet('참석신청');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['제출시간', '성명']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:B1').setFontWeight('bold');
  }

  return sheet;
}

function sanitizeCallback_(callback) {
  var value = (callback || '').trim();
  if (/^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(value)) {
    return value;
  }
  return '';
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpResponse(callback, payload) {
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(payload) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
