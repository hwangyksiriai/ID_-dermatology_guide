/**
 * ID 피부과 × SIRIAI — 인플루언서 신청 폼 → 구글시트 연동
 *
 * 사용법: README_신청폼_연동.md 참고
 * 이 코드를 신청 데이터를 받을 구글시트의 [확장 프로그램 > Apps Script]에 붙여넣고
 * "배포 > 새 배포 > 웹 앱"으로 배포하세요. (액세스 권한: 모든 사용자)
 *
 * 시트와 헤더는 자동으로 생성/동기화됩니다.
 * (헤더가 코드와 다르면 1행이 자동으로 갱신됩니다 — 컬럼을 추가해도 시트를 비울 필요 없음)
 */

var SHEET_NAME = '신청';
var HEADERS = ['신청일시', '이름', '인스타그램', '휴대폰', '이메일', '방문가능일정', '2차활용동의'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // 헤더 보장/동기화 — 비어있거나 코드의 헤더와 다르면 1행을 갱신
    var needHeader = false;
    if (sheet.getLastRow() === 0) {
      needHeader = true;
    } else {
      var cur = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
      if (cur.join('|') !== HEADERS.join('|')) needHeader = true;
    }
    if (needHeader) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }

    var d = JSON.parse(e.postData.contents);
    sheet.appendRow([
      new Date(),
      d.name || '',
      d.instagram || '',
      d.phone || '',
      d.email || '',
      d.visit || '',
      d.consent ? '동의' : '미동의'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// 배포 동작 확인용 (브라우저로 웹앱 URL 열면 표시됨)
function doGet() {
  return ContentService
    .createTextOutput('ID 피부과 신청 폼 엔드포인트 — 정상 동작 중')
    .setMimeType(ContentService.MimeType.TEXT);
}
