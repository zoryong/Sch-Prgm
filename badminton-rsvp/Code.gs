/**
 * 배드민턴 모임 참석 신청 Google Apps Script
 *
 * 스프레드시트 1행 헤더: 제출시간 | 성명
 */

function doGet(e) {
  var name = (e && e.parameter && e.parameter.name || '').trim();

  if (!name) {
    return ContentService
      .createTextOutput('배드민턴 모임 참석 신청 API')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  return submitRsvp(name);
}

function doPost(e) {
  var name = (e.parameter.name || '').trim();
  return submitRsvp(name);
}

function submitRsvp(name) {
  try {
    if (!name) {
      return jsonResponse({
        ok: false,
        message: '성명을 입력해 주세요.'
      });
    }

    var sheet = getOrCreateSheet();
    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      var existingName = String(rows[i][1] || '').trim();
      if (existingName === name) {
        return jsonResponse({
          ok: false,
          message: '이미 신청하셨습니다.'
        });
      }
    }

    sheet.appendRow([new Date(), name]);

    return jsonResponse({
      ok: true,
      message: '참석 신청이 완료되었습니다.'
    });
  } catch (error) {
    return jsonResponse({
      ok: false,
      message: '오류가 발생했습니다. 다시 시도해 주세요.'
    });
  }
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
