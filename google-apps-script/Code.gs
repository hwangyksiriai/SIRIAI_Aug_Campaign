/**
 * SIRIAI 10월 통합 지원서 → 구글시트 연동 스크립트
 *
 * 사용법 (처음 설정)
 * 1) 새로 만든 스프레드시트에서 [확장 프로그램] → [Apps Script] 클릭
 * 2) 기본 생성된 코드를 모두 지우고 이 파일 내용을 붙여넣기
 * 3) 우측 상단 [배포] → [새 배포]
 *    - 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한: 모든 사용자
 * 4) 배포 후 나오는 웹 앱 URL(.../exec)을 복사
 * 5) index-A~I.html 의 SHEET_ENDPOINT 값을 그 URL로 교체
 *
 * 코드만 수정된 경우 (URL 이미 있음)
 * - Apps Script 에디터에서 코드를 이 내용으로 덮어쓴 뒤
 * - [배포] → [배포 관리] → 연필 아이콘 → 버전: "새 버전" → 배포
 * - 이 방식이면 exec URL이 그대로 유지되어 HTML 쪽은 수정할 필요 없음
 */

var SHEET_NAME = '시트1';

var HEADERS = [
  '제출일시',
  '구분',
  '고료',
  '지원유형',
  '지원브랜드',
  '리노센트랩 진행주차',
  '이름',
  '인스타그램',
  '휴대폰',
  '이메일',
  '우편번호',
  '배송지 주소',
  '요청사항',
  '동의여부'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet();
    var timestamp = new Date();

    var brandsArr = Array.isArray(data.brands) ? data.brands : (data.brands ? [data.brands] : ['']);
    var isIndividual = data.support_type === '개별 지원';

    // 개별 지원으로 여러 브랜드를 선택한 경우 브랜드마다 별도 행으로 기록.
    // 전체 지원은 브랜드를 나열한 값 그대로 한 행에 기록.
    var brandRows = isIndividual ? brandsArr : [brandsArr.join(', ')];

    brandRows.forEach(function (brandName) {
      var week = (brandName.indexOf('Re:Nnocent') > -1) ? (data.nnocent_week || '') : '';
      sheet.appendRow([
        timestamp,
        data.division || '',
        data.amount || '',
        data.support_type || '',
        brandName,
        week,
        data.name || '',
        data.instagram || '',
        data.phone || '',
        data.email || '',
        data.zonecode || '',
        data.address || '',
        data.notes || '',
        data.agree || ''
      ]);
    });

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

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
