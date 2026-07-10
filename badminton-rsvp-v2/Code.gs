/**
 * 배드민턴 모임 참석 신청 (v2 - Google 로그인 필수)
 *
 * - 프론트엔드에서 Google Identity Services로 로그인 → ID 토큰(JWT) 전송
 * - 서버가 토큰을 검증하여 이메일을 확보 (모든 Google 계정 허용)
 * - 이메일 기준으로 중복 신청 방지
 * - 서명 이미지는 Drive 폴더에 PNG로 저장하고 시트에 링크 기록
 *
 * 스프레드시트 1행 헤더: 제출시간 | 성명 | 이메일 | 서명이미지
 */

// GCP에서 발급한 OAuth 클라이언트 ID (index.html의 GOOGLE_CLIENT_ID와 동일해야 함)
var OAUTH_CLIENT_ID = 'YOUR_OAUTH_CLIENT_ID.apps.googleusercontent.com';

// 허용 도메인 (v2는 제한 없음. 빈 문자열이면 모든 Google 계정 허용)
var ALLOWED_DOMAIN = '';

var SIGNATURE_FOLDER_NAME = '배드민턴서명_v2';

function doPost(e) {
  try {
    var data = parseRequest(e);
    var action = (data.action || '').trim();

    if (action === 'ping') {
      return jsonResponse({ ok: true, message: 'pong' });
    }

    var claims = verifyIdToken(data.idToken);
    if (!claims) {
      return jsonResponse({ ok: false, message: '로그인 정보가 유효하지 않습니다. 다시 로그인해 주세요.' });
    }

    var email = String(claims.email || '').toLowerCase();
    if (!email) {
      return jsonResponse({ ok: false, message: '이메일 정보를 확인할 수 없습니다. 다시 로그인해 주세요.' });
    }

    if (ALLOWED_DOMAIN) {
      var domain = String(claims.hd || email.split('@')[1] || '').toLowerCase();
      if (domain !== ALLOWED_DOMAIN.toLowerCase()) {
        return jsonResponse({ ok: false, message: '이 설문은 @' + ALLOWED_DOMAIN + ' 계정만 응답할 수 있습니다.' });
      }
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
      var existingEmail = String(rows[i][2] || '').trim().toLowerCase();
      if (existingEmail === email) {
        return jsonResponse({ ok: false, message: '이미 신청하셨습니다. (' + email + ')' });
      }
    }

    var signatureUrl = saveSignature(name, email, signature);

    sheet.appendRow([new Date(), name, email, signatureUrl]);

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
    .createTextOutput('배드민턴 모임 참석 신청 API (v2)')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Google ID 토큰(JWT)을 검증하고 클레임(payload)을 반환합니다.
 * 검증 실패 시 null.
 */
function verifyIdToken(idToken) {
  if (!idToken) {
    return null;
  }
  try {
    var url = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken);
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      return null;
    }
    var info = JSON.parse(response.getContentText());

    // 이 토큰이 우리 앱(클라이언트 ID)을 위해 발급된 것인지 확인
    if (info.aud !== OAUTH_CLIENT_ID) {
      return null;
    }
    // 이메일 인증 여부 확인
    if (String(info.email_verified) !== 'true') {
      return null;
    }
    // 만료 시간 확인 (tokeninfo가 이미 검사하지만 이중 확인)
    if (info.exp && (Number(info.exp) * 1000) < Date.now()) {
      return null;
    }
    return info;
  } catch (err) {
    return null;
  }
}

/**
 * fetch(text/plain)로 전송된 JSON 본문 또는 폼 파라미터를 처리합니다.
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
function saveSignature(name, email, signature) {
  var base64 = signature.replace(/^data:image\/\w+;base64,/, '');
  var bytes = Utilities.base64Decode(base64);
  var safeName = name.replace(/[\\\/:*?"<>|]/g, '_');
  var fileName = safeName + '_' + formatTimestamp(new Date()) + '.png';
  var blob = Utilities.newBlob(bytes, 'image/png', fileName);

  var folder = getOrCreateFolder(SIGNATURE_FOLDER_NAME);
  var file = folder.createFile(blob);
  file.setDescription('신청자: ' + name + ' / ' + email);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (sharingError) {
    // 도메인 정책 등으로 공유 설정이 막혀도 저장은 유지
  }

  return file.getUrl();
}

function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
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
    sheet.appendRow(['제출시간', '성명', '이메일', '서명이미지']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:D1').setFontWeight('bold');
  }

  return sheet;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
