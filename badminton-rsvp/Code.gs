/**
 * 배드민턴 모임 참석 신청 Google Apps Script
 *
 * 스프레드시트 1행 헤더: 제출시간 | 성명 | 서명이미지
 * 서명 이미지는 스프레드시트와 같은 Drive 폴더 안에, 스프레드시트와 동일한 이름의
 * 하위 폴더에 PNG로 저장되고, 시트에는 파일 링크가 기록됩니다.
 */

function doPost(e) {
  try {
    var data = parseRequest(e);
    var action = (data.action || '').trim();

    if (action === 'ping') {
      return jsonResponse({ ok: true, message: 'pong' });
    }

    var name = (data.name || '').trim();
    var signature = (data.signature || '').trim();

    if (!name) {
      return jsonResponse({ ok: false, message: '성명을 입력해 주세요.' });
    }

    if (!signature) {
      return jsonResponse({ ok: false, message: '서명을 입력해 주세요.' });
    }

    var sheet = getOrCreateSheet();
    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      var existingName = String(rows[i][1] || '').trim();
      if (existingName === name) {
        return jsonResponse({ ok: false, message: '이미 신청하셨습니다.' });
      }
    }

    var signatureUrl = saveSignature(name, signature);

    sheet.appendRow([new Date(), name, signatureUrl]);

    return jsonResponse({ ok: true, message: '참석 신청이 완료되었습니다.' });
  } catch (error) {
    return jsonResponse({ ok: false, message: '오류가 발생했습니다. 다시 시도해 주세요.' });
  }
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'ping') {
    return jsonResponse({ ok: true, message: 'pong' });
  }
  return ContentService
    .createTextOutput('배드민턴 모임 참석 신청 API')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * fetch(text/plain)로 전송된 JSON 본문 또는 폼 파라미터를 모두 처리합니다.
 */
function parseRequest(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (parseError) {
      // JSON이 아니면 파라미터로 폴백
    }
  }
  return (e && e.parameter) ? e.parameter : {};
}

/**
 * Base64 PNG(data URL 포함 가능)를 Drive에 저장하고 링크를 반환합니다.
 */
function saveSignature(name, signature) {
  var base64 = signature.replace(/^data:image\/\w+;base64,/, '');
  var bytes = Utilities.base64Decode(base64);
  var fileName = name + '_' + formatTimestamp(new Date()) + '.png';
  var blob = Utilities.newBlob(bytes, 'image/png', fileName);

  var folder = getOrCreateSignatureFolder();
  var file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (sharingError) {
    // 도메인 정책 등으로 공유 설정이 막혀도 저장은 유지
  }

  return file.getUrl();
}

/**
 * 스프레드시트가 있는 Drive 폴더 안에, 스프레드시트와 같은 이름의 하위 폴더를 찾거나 만듭니다.
 */
function getOrCreateSignatureFolder() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var folderName = spreadsheet.getName();
  var spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
  var parents = spreadsheetFile.getParents();

  if (!parents.hasNext()) {
    var rootFolders = DriveApp.getFoldersByName(folderName);
    if (rootFolders.hasNext()) {
      return rootFolders.next();
    }
    return DriveApp.createFolder(folderName);
  }

  var parentFolder = parents.next();
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

function formatTimestamp(date) {
  var tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  return Utilities.formatDate(date, tz, 'yyyyMMdd_HHmmss');
}

function getOrCreateSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('참석신청');

  if (!sheet) {
    sheet = spreadsheet.insertSheet('참석신청');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['제출시간', '성명', '서명이미지']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:C1').setFontWeight('bold');
  }

  return sheet;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
