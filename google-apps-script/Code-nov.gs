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
  'PURCELL (퍼셀)',
  'NURICLE (뉴리클)',
  'Re:Nnocent (리노센트랩)'
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
var TEXT_COLUMNS = ['제출일시', '휴대폰', '우편번호'];

// 'username' · '@username' · 'instagram.com/username' · 프로필 링크를
// https://www.instagram.com/username/ 형태로 맞춘다. 알아볼 수 없으면 입력값 그대로 둔다.
function instagramUrl(raw) {
  var v = String(raw || '').trim().replace(/\s+/g, '');
  var m = v.match(/instagram\.com\/@?([A-Za-z0-9._]+)/i);
  var handle = m ? m[1] : v.replace(/^@+/, '');
  if (!/^[A-Za-z0-9._]{1,30}$/.test(handle) || /^(p|reel|reels|stories|explore|accounts)$/i.test(handle)) return v;
  return 'https://www.instagram.com/' + handle.toLowerCase() + '/';
}

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

    // 제출 시간: 스크립트 시간대 설정과 무관하게 한국 시간으로, 날짜와 시각을 함께 기록
    var submittedAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    var row = [submittedAt, data.division || '', data.amount || '', data.support_type || '']
      .concat(brandChecks)
      .concat([
        data.name || '',
        instagramUrl(data.instagram),
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
  } else {
    ensureBrandColumns(sheet);
  }
  return sheet;
}

// 이미 데이터가 쌓인 시트에 브랜드가 추가된 경우, 빠진 브랜드 컬럼을 앞 브랜드 바로 뒤에 끼워 넣어
// 기존 행과 새 행의 칸이 어긋나지 않게 한다. (기존 행의 새 브랜드 칸은 빈칸으로 남는다)
function ensureBrandColumns(sheet) {
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  BRAND_COLUMNS.forEach(function (col, i) {
    if (header.indexOf(col) > -1) return;
    var prev = i === 0 ? HEADERS.indexOf(col) : header.indexOf(BRAND_COLUMNS[i - 1]) + 1;
    sheet.insertColumnAfter(prev);
    sheet.getRange(1, prev + 1).setValue(col).setFontWeight('bold');
    header.splice(prev, 0, col);
  });
}
