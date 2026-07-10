/**
 * 배드민턴 모임 참석 신청 (v3 - 지정 도메인 계정만 응답, HtmlService 방식)
 *
 * 동작 방식 (로그인 버튼 없음):
 * - 페이지 자체를 Apps Script 웹앱(HtmlService)으로 제공
 * - 웹앱 배포 시 "액세스 권한"을 도메인 내 사용자로 제한하면
 *   Google이 접속 시 로그인을 강제하고, 같은 도메인 사용자만 접근 가능
 * - 실행 계정 "나(Me)"이므로 내 Drive/시트에 저장됨
 * - Session.getActiveUser().getEmail() 로 접속자 이메일 확보 (같은 도메인이라 정상 반환)
 * - 이메일 기준 중복 신청 방지
 *
 * 스프레드시트 1행 헤더: 제출시간 | 성명 | 이메일 | 서명이미지
 *
 * [배포 설정]
 *  - 실행 계정(Execute as): 나(Me)
 *  - 액세스 권한(Who has access): "yihongcheon-h.goeyi.kr 도메인 내 모든 사용자"
 *    (Google Workspace 관리 도메인일 때 이 옵션이 표시됩니다)
 */

// 허용 도메인 (이 도메인 계정만 응답 가능) - 서버 측 이중 검증용
var ALLOWED_DOMAIN = 'yihongcheon-h.goeyi.kr';

var SIGNATURE_FOLDER_NAME = '배드민턴서명_v3';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('배드민턴 모임 친목 저녁 식사 신청')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * 페이지 로드 시 접속자 이메일을 반환합니다. (같은 도메인이면 정상 반환)
 */
function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail() || '';
}

/**
 * 참석 신청 처리. 클라이언트에서 google.script.run 으로 호출됩니다.
 * payload: { name: String, signature: String(dataURL PNG) }
 */
function submitRsvp(payload) {
  try {
    payload = payload || {};
    var email = String(Session.getActiveUser().getEmail() || '').toLowerCase();

    if (!email) {
      return { ok: false, message: '로그인 정보를 확인할 수 없습니다. 학교 계정으로 접속했는지 확인해 주세요.' };
    }

    if (ALLOWED_DOMAIN) {
      var domain = email.split('@')[1] || '';
      if (domain !== ALLOWED_DOMAIN.toLowerCase()) {
        return { ok: false, message: '이 설문은 @' + ALLOWED_DOMAIN + ' 계정만 응답할 수 있습니다.' };
      }
    }

    var name = String(payload.name || '').trim();
    var signature = String(payload.signature || '').trim();

    if (!name) {
      return { ok: false, message: '성명을 입력해 주세요.' };
    }
    if (!signature) {
      return { ok: false, message: '서명을 입력해 주세요.' };
    }

    var sheet = getOrCreateSheet();
    var rows = sheet.getDataRange().getValues();

    for (var i = 1; i < rows.length; i++) {
      var existingEmail = String(rows[i][2] || '').trim().toLowerCase();
      if (existingEmail === email) {
        return { ok: false, message: '이미 신청하셨습니다. (' + email + ')' };
      }
    }

    var signatureUrl = saveSignature(name, email, signature);

    sheet.appendRow([new Date(), name, email, signatureUrl]);

    return { ok: true, message: '참석 신청이 완료되었습니다.' };
  } catch (error) {
    return { ok: false, message: '오류가 발생했습니다. 다시 시도해 주세요.' };
  }
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
