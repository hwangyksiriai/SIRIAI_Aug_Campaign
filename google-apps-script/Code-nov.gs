/**
 * SIRIAI 11월 통합 지원서 → 구글시트 연동 스크립트 (신규인원 nov-1~4 전용, 새 시트)
 *
 * 처음 설정
 * 1) 새 스프레드시트에서 [확장 프로그램] → [Apps Script] 클릭
 * 2) 기본 코드를 모두 지우고 이 파일 내용을 붙여넣은 뒤 저장
 * 3) 우측 상단 [배포] → [새 배포]
 *    - 유형: 웹 앱 / 실행 계정: 나 / 액세스 권한: 모든 사용자
 * 4) 배포 후 나오는 웹 앱 URL(.../exec)을 nov-1~4.html 의 SHEET_ENDPOINT 값으로 교체
 *
 * 코드만 수정된 경우
 * - 코드를 덮어쓴 뒤 [배포] → [배포 관리] → 연필 아이콘 → 버전: "새 버전" → 배포
 * - exec URL이 그대로 유지되어 HTML 쪽은 수정할 필요 없음
 *
 * 시트 구조: 지원자 1명 = 1행
 * - 전체 지원: 모든 브랜드 컬럼에 체크
 * - 개별 지원: 선택한 브랜드 컬럼에만 체크
 */

var SHEET_NAME = '시트1';

// 지원 페이지(formBrands)의 브랜드 이름과 글자 하나까지 똑같아야 체크된다.
var BRAND_COLUMNS = [
  '루솜 LUSOM',
  'LOYR (로이어)',
  'it:fu (잇퓨)',
  'PURCELL (퍼셀)'
];

var HEADERS = ['제출일시', '구분', '고료', '지원유형']
  .concat(BRAND_COLUMNS)
  .concat([
    '이름',
    '인스타그램',
    '휴대폰',
    '이메일',
    '우편번호',
    '배송지 주소',
    '요청사항',
    '동의여부'
  ]);

// 앞자리 0이 사라지지 않도록 텍스트로 저장할 컬럼
var TEXT_COLUMNS = ['휴대폰', '우편번호'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet();

    var brandsArr = Array.isArray(data.brands) ? data.brands : (data.brands ? [data.brands] : []);
    var isAll = data.support_type === '전체 지원';
    var brandChecks = BRAND_COLUMNS.map(function (col) {
      return isAll || brandsArr.indexOf(col) > -1;
    });

    // 페이지에서 배송지 주소와 상세 주소를 합쳐 보내지만, 따로 오는 경우에도 한 칸에 합친다.
    var address = [data.address, data.addressDetail].filter(function (v) { return v; }).join(' ');

    var row = [new Date(), data.division || '', data.amount || '', data.support_type || '']
      .concat(brandChecks)
      .concat([
        data.name || '',
        data.instagram || '',
        data.phone || '',
        data.email || '',
        data.zonecode || '',
        address,
        data.notes || '',
        data.agree || ''
      ]);

    var nextRow = sheet.getLastRow() + 1;
    TEXT_COLUMNS.forEach(function (name) {
      sheet.getRange(nextRow, HEADERS.indexOf(name) + 1).setNumberFormat('@');
    });
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    sheet.getRange(nextRow, HEADERS.indexOf(BRAND_COLUMNS[0]) + 1, 1, BRAND_COLUMNS.length).insertCheckboxes();

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
