/**
 * SIRIAI 통합 지원서 → 구글시트 연동 스크립트 (신규인원 sejgfb·eaaeyf·cpdr9t·5g7v6y.html 전용, 새 시트)
 *
 * 처음 설정
 * 1) 새 스프레드시트에서 [확장 프로그램] → [Apps Script] 클릭
 * 2) 기본 코드를 모두 지우고 이 파일 내용을 붙여넣은 뒤 저장
 * 3) 우측 상단 [배포] → [새 배포]
 *    - 유형: 웹 앱 / 실행 계정: 나 / 액세스 권한: 모든 사용자
 * 4) 배포 후 나오는 웹 앱 URL(.../exec)을 신규인원 4개 페이지(sejgfb·eaaeyf·cpdr9t·5g7v6y.html)의 SHEET_ENDPOINT 값으로 교체
 *
 * 코드만 수정된 경우
 * - 코드를 덮어쓴 뒤 [배포] → [배포 관리] → 연필 아이콘 → 버전: "새 버전" → 배포
 * - exec URL이 그대로 유지되어 HTML 쪽은 수정할 필요 없음
 *
 * 시트 구조: 지원자 1명 = 1행
 * - 페이지가 보낸 브랜드 목록(전체 지원이면 그 페이지의 모든 브랜드)에 해당하는 컬럼에만 체크
 * - 값은 1행 제목을 보고 같은 이름의 칸에 넣는다. 칸 순서가 바뀌거나 컬럼이 추가돼도 어긋나지 않는다.
 */

var SHEET_NAME = '시트1';

// 지원 페이지(formBrands)의 브랜드 이름과 글자 하나까지 똑같아야 체크된다.
// 기존 브랜드 컬럼: 지원유형 바로 뒤에 모여 있다. (뉴리클은 기존 지원 기록 보존용으로 남겨 둔다)
var BRAND_COLUMNS = [
  '루솜 LUSOM',
  'LOYR (로이어)',
  'it:fu (잇퓨)',
  'PURCELL (퍼셀)',
  'NURICLE (뉴리클)',
  'Re:Nnocent (리노센트랩)'
];

// 나중에 추가된 브랜드 컬럼: 헷갈리지 않도록 시트 맨 뒤(동의여부 다음)에 붙인다.
var EXTRA_BRAND_COLUMNS = [
  'FROM9 (프롬나인)'
];

var ALL_BRAND_COLUMNS = BRAND_COLUMNS.concat(EXTRA_BRAND_COLUMNS);

var INFO_COLUMNS = [
  '이름',
  '인스타그램',
  '휴대폰',
  '이메일',
  '우편번호',
  '배송지 주소',
  '요청사항',
  '동의여부'
];

var HEADERS = ['제출일시', '구분', '고료', '지원유형']
  .concat(BRAND_COLUMNS)
  .concat(INFO_COLUMNS)
  .concat(EXTRA_BRAND_COLUMNS);

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
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var brandsArr = Array.isArray(data.brands) ? data.brands : (data.brands ? [data.brands] : []);
    // 페이지는 전체 지원일 때도 그 페이지의 브랜드 목록을 보낸다. 목록이 비어 온 경우에만 기존 브랜드 전체로 본다.
    if (data.support_type === '전체 지원' && brandsArr.length === 0) brandsArr = BRAND_COLUMNS.slice();

    // 페이지에서 배송지 주소와 상세 주소를 합쳐 보내지만, 따로 오는 경우에도 한 칸에 합친다.
    var address = [data.address, data.addressDetail].filter(function (v) { return v; }).join(' ');

    // 제출 시간: 스크립트 시간대 설정과 무관하게 한국 시간으로, 날짜와 시각을 함께 기록
    var submittedAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    var values = {
      '제출일시': submittedAt,
      '구분': data.division || '',
      '고료': data.amount || '',
      '지원유형': data.support_type || '',
      '이름': data.name || '',
      '인스타그램': instagramUrl(data.instagram),
      '휴대폰': data.phone || '',
      '이메일': data.email || '',
      '우편번호': data.zonecode || '',
      '배송지 주소': address,
      '요청사항': data.notes || '',
      '동의여부': data.agree || ''
    };
    ALL_BRAND_COLUMNS.forEach(function (col) {
      values[col] = brandsArr.indexOf(col) > -1;
    });

    var row = header.map(function (name) {
      return Object.prototype.hasOwnProperty.call(values, name) ? values[name] : '';
    });

    var nextRow = sheet.getLastRow() + 1;
    TEXT_COLUMNS.forEach(function (name) {
      var c = header.indexOf(name);
      if (c > -1) sheet.getRange(nextRow, c + 1).setNumberFormat('@');
    });
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    ALL_BRAND_COLUMNS.forEach(function (col) {
      var c = header.indexOf(col);
      if (c > -1) sheet.getRange(nextRow, c + 1).insertCheckboxes();
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
  } else {
    ensureColumns(sheet);
  }
  return sheet;
}

// 이미 데이터가 쌓인 시트에 빠진 컬럼을 채워 넣는다. (기존 행의 새 칸은 빈칸으로 남는다)
// - 기존 브랜드(BRAND_COLUMNS): 앞 브랜드 바로 뒤에 끼워 넣는다.
// - 그 밖의 컬럼(추가 브랜드 등): 시트 맨 뒤에 붙인다.
function ensureColumns(sheet) {
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  BRAND_COLUMNS.forEach(function (col, i) {
    if (header.indexOf(col) > -1) return;
    var prev = i === 0 ? HEADERS.indexOf(col) : header.indexOf(BRAND_COLUMNS[i - 1]) + 1;
    sheet.insertColumnAfter(prev);
    sheet.getRange(1, prev + 1).setValue(col).setFontWeight('bold');
    header.splice(prev, 0, col);
  });
  HEADERS.forEach(function (col) {
    if (header.indexOf(col) > -1) return;
    header.push(col);
    if (sheet.getMaxColumns() < header.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), header.length - sheet.getMaxColumns());
    sheet.getRange(1, header.length).setValue(col).setFontWeight('bold');
  });
}
