function doGet(e) {
  try {
    // ⭐️ 1. 여기에 선생님의 시트 주소창에서 복사한 고유 ID를 넣으세요!
    var sheetId = '1VjF84E7kdloHBlwLLjDmLpG66s_djoSk-N7lvLFNhL8'; 
    var sheetName = '참석신청'; 
    
    // openById를 사용하여 속도를 획기적으로 높입니다.
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    
    // 초기 연결 테스트(ping) 응답
    if (e.parameter.action === 'ping') {
      return createJsonResponse(e, { ok: true, message: 'pong' });
    }
    
    var name = e.parameter.name;
    if (!name) {
      return createJsonResponse(e, { ok: false, message: '이름이 전달되지 않았습니다.' });
    }
    
    // ⭐️ 2. 중복 이름 검사 로직
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      // B열(인덱스 1)에 입력된 이름이 있는지 확인합니다.
      if (data[i][1] === name) { 
        // 중복이면 '이미 신청하셨습니다.' 메시지를 HTML로 보냅니다.
        return createJsonResponse(e, { ok: false, message: '이미 신청하셨습니다.' });
      }
    }
    
    // 중복이 아니면 시트에 기록
    sheet.appendRow([new Date(), name]);
    return createJsonResponse(e, { ok: true, message: '참석 신청이 완료되었습니다.' });
    
  } catch (error) {
    // 예상치 못한 에러 발생 시 처리
    return createJsonResponse(e, { ok: false, message: error.toString() });
  }
}

// JSONP 응답을 만들어주는 필수 함수
function createJsonResponse(e, response) {
  var callback = e.parameter.callback;
  var jsonString = JSON.stringify(response);
  var jsonp = callback ? callback + '(' + jsonString + ');' : jsonString;
  
  return ContentService.createTextOutput(jsonp)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
