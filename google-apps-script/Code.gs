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

// 브랜드마다 자기 컬럼을 갖는다. 지원한 브랜드 컬럼에는 'O'가 채워진다.
var BRAND_COLUMNS = [
  'Re:Nnocent (리노센트랩)',
  '루솜 LUSOM',
  'it:fu (잇퓨)',
  'Veryhour (베리아워)',
  'NURICLE (뉴리클)',
  'CUMCURA (쿰쿠라)'
];

var HEADERS = ['제출일시', '구분', '고료', '지원유형']
  .concat(BRAND_COLUMNS)
  .concat([
    '리노센트랩 진행주차',
    '이름',
    '인스타그램',
    '휴대폰',
    '이메일',
    '우편번호',
    '배송지 주소',
    '요청사항',
    '동의여부'
  ]);

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet();

    var brandsArr = Array.isArray(data.brands) ? data.brands : (data.brands ? [data.brands] : []);

    var brandFlags = BRAND_COLUMNS.map(function (col) {
      return brandsArr.indexOf(col) > -1 ? 'O' : '';
    });

    var row = [new Date(), data.division || '', data.amount || '', data.support_type || '']
      .concat(brandFlags)
      .concat([
        data.nnocent_week || '',
        data.name || '',
        data.instagram || '',
        data.phone || '',
        data.email || '',
        data.zonecode || '',
        data.address || '',
        data.notes || '',
        data.agree || ''
      ]);

    var nextRow = sheet.getLastRow() + 1;
    // 우편번호는 숫자로 인식되면 앞자리 0이 사라지므로, 값을 쓰기 전에
    // 해당 셀을 텍스트 서식으로 고정해 그대로 저장되게 한다.
    var zipColIndex = HEADERS.indexOf('우편번호') + 1;
    sheet.getRange(nextRow, zipColIndex).setNumberFormat('@');
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);

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

/**
 * 컬럼 구조를 바꾼 뒤(예: 이번처럼 지원브랜드를 컬럼으로 나눌 때) 한 번만 실행.
 * 기존 테스트 데이터를 모두 지우고 새 헤더로 다시 씁니다.
 * Apps Script 에디터에서 이 함수(resetSheet)를 선택한 뒤 [실행] 버튼으로 딱 한 번 돌리면 됩니다.
 */
function resetSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  sheet.clear();
  sheet.appendRow(HEADERS);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}
