/**
 * 배드민턴 모임 참석 신청 Google Apps Script
 *
 * [배포 방법]
 * 1. 스프레드시트 > 확장 프로그램 > Apps Script
 * 2. Code.gs 와 Index.html 내용 붙여넣기
 * 3. 배포 > 새 배포 > 웹 앱 (실행: 나, 액세스: 모든 사용자)
 * 4. 배포 URL을 공유 (이 URL로 접속하면 신청 페이지가 열림)
 *
 * 스프레드시트 1행 헤더: 제출시간 | 성명
 */

function doGet(e) {
  var params = (e && e.parameter) || {};
  var name = (params.name || '').trim();
  var callback = (params.callback || '').trim();

  if (name || callback) {
    return handleApiRequest_(name, callback);
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('배드민턴 모임 친목 저녁 식사 신청')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

function submitRsvp(name) {
  try {
    return processSubmission(name);
  } catch (error) {
    return {
      ok: false,
      message: '오류가 발생했습니다. 다시 시도해 주세요.'
    };
  }
}

function handleApiRequest_(name, callback) {
  if (!name) {
    if (callback) {
      return jsonpResponse(callback, {
        ok: false,
        message: '성명을 입력해 주세요.'
      });
    }

    return ContentService
      .createTextOutput('배드민턴 모임 참석 신청 API')
      .setMimeType(ContentService.MimeType.TEXT);
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
