/**
 * ID 피부과 × SIRIAI — 인플루언서 신청 폼 → 구글시트 연동
 *
 * 사용법: README_신청폼_연동.md 참고
 * 이 코드를 신청 데이터를 받을 구글시트의 [확장 프로그램 > Apps Script]에 붙여넣고
 * "배포 > 새 배포 > 웹 앱"으로 배포하세요. (액세스 권한: 모든 사용자)
 *
 * 시트와 헤더는 최초 신청 시 자동으로 생성됩니다.
 */

var SHEET_NAME = '신청';
var HEADERS = ['신청일시', '이름', '나이', '인스타그램', '휴대폰', '이메일', '2차활용동의'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // 헤더가 없으면 자동 생성
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }

    var d = JSON.parse(e.postData.contents);
    sheet.appendRow([
      new Date(),
      d.name || '',
      d.age || '',
      d.instagram || '',
      d.phone || '',
      d.email || '',
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
