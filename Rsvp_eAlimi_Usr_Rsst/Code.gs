/**
 * 학교 e-알리미 회원 정보 등록 Google Apps Script
 *
 * 스프레드시트 1행 헤더: 제출시간 | 개인정보동의 | 성함 | 휴대전화 | 서명이미지
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

    var consent = (data.consent || '').trim();
    var name = (data.name || '').trim();
    var phone = normalizePhone(data.phone || '');
    var signature = (data.signature || '').trim();

    if (!consent) {
      return jsonResponse({ ok: false, message: '개인정보 수집·이용 및 제3자 제공 동의 여부를 선택해 주세요.' });
    }

    if (consent !== '예' && consent !== '아니오') {
      return jsonResponse({ ok: false, message: '동의 여부를 올바르게 선택해 주세요.' });
    }

    if (!name) {
      return jsonResponse({ ok: false, message: '성함을 입력해 주세요.' });
    }

    if (!phone) {
      return jsonResponse({ ok: false, message: '휴대전화 연락처를 입력해 주세요.' });
    }

    if (!isValidPhone(phone)) {
      return jsonResponse({ ok: false, message: '휴대전화 번호 형식을 확인해 주세요.' });
    }

    if (!signature) {
      return jsonResponse({ ok: false, message: '서명을 입력해 주세요.' });
    }

    var sheet = getOrCreateSheet();
    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      var existingPhone = normalizePhone(String(rows[i][3] || ''));
      if (existingPhone === phone) {
        return jsonResponse({ ok: false, message: '이미 등록된 휴대전화 번호입니다.' });
      }
    }

    var signatureUrl = saveSignature(name, signature);

    sheet.appendRow([new Date(), consent, name, formatPhone(phone), signatureUrl]);

    return jsonResponse({ ok: true, message: '응답이 제출되었습니다. 감사합니다.' });
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
    .createTextOutput('학교 e-알리미 회원 정보 등록 API')
    .setMimeType(ContentService.MimeType.TEXT);
}

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

function normalizePhone(value) {
  return String(value || '').replace(/\D/g, '');
}

function isValidPhone(value) {
  return /^01[016789]\d{7,8}$/.test(normalizePhone(value));
}

function formatPhone(digits) {
  var phone = normalizePhone(digits);
  if (phone.length === 11) {
    return phone.slice(0, 3) + '-' + phone.slice(3, 7) + '-' + phone.slice(7);
  }
  if (phone.length === 10) {
    return phone.slice(0, 3) + '-' + phone.slice(3, 6) + '-' + phone.slice(6);
  }
  return phone;
}

function getOrCreateSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('등록응답');

  if (!sheet) {
    sheet = spreadsheet.insertSheet('등록응답');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['제출시간', '개인정보동의', '성함', '휴대전화', '서명이미지']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:E1').setFontWeight('bold');
  }

  return sheet;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
