/**
 * SIRIAI 12월 통합 지원서 → 구글시트 연동 스크립트
 * (신규인원 sejgfb·eaaeyf·cpdr9t·5g7v6y.html 전용, 새 시트)
 *
 * 처음 설정
 * 1) 새 스프레드시트에서 [확장 프로그램] → [Apps Script] 클릭
 * 2) 기본 코드를 모두 지우고 이 파일 내용을 붙여넣은 뒤 저장
 * 3) 우측 상단 [배포] → [새 배포]
 *    - 유형: 웹 앱 / 실행 계정: 나 / 액세스 권한: 모든 사용자
 * 4) 배포 후 나오는 웹 앱 URL(.../exec)을 신규인원 4개 페이지의 SHEET_ENDPOINT 값으로 교체
 *
 * 코드만 수정된 경우
 * - 코드를 덮어쓴 뒤 [배포] → [배포 관리] → 연필 아이콘 → 버전: "새 버전" → 배포
 * - exec URL이 그대로 유지되어 HTML 쪽은 수정할 필요 없음
 * - 웹 앱 URL을 브라우저로 열어 {"ok":true,...} 가 보이면 연결 정상
 *
 * 시트 규칙 (라만테아 시트 연동과 동일)
 * - 모든 지원은 한 탭(시트1, 없으면 첫 번째 탭)에 1명 = 1행으로 쌓인다
 * - 휴대폰·우편번호·제출일시는 텍스트 서식이라 앞자리 0이 사라지지 않는다
 * - 배송지 주소와 상세 주소는 '배송지 주소' 한 칸에 합친다
 * - 고료별로 행 전체는 옅게, 고료 칸은 진하게 색을 칠한다
 * - 이름·요청사항에 고료 요청으로 보이는 내용이 있으면 그 칸을 노란색 + 굵게 표시한다
 * - 브랜드 컬럼은 페이지가 보낸 브랜드 목록(전체 지원이면 그 페이지의 모든 브랜드)에만 체크
 * - 값은 1행 제목을 보고 같은 이름의 칸에 넣으므로, 컬럼이 추가되거나 순서가 바뀌어도 어긋나지 않는다
 */

var SHEET_NAME = '시트1';
var TIMEZONE = 'Asia/Seoul';

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

// 시트를 새로 만들 때 쓰는 열 너비 (없는 컬럼은 기본 너비)
var COLUMN_WIDTHS = {
  '제출일시': 150, '구분': 60, '고료': 90, '지원유형': 80,
  '이름': 100, '인스타그램': 230, '휴대폰': 130, '이메일': 190,
  '우편번호': 80, '배송지 주소': 340, '요청사항': 280, '동의여부': 70
};

// 고료별 색 — 행 전체는 옅게, 고료 칸은 진하게. 새 고료가 생기면 색을 추가한다.
var PRICE_COLORS = {
  '50,000원':  { row: '#EEF6EA', cell: '#C9E3BD' },  // 초록
  '100,000원': { row: '#EAF2FB', cell: '#BFD8F2' },  // 파랑
  '200,000원': { row: '#F4ECF8', cell: '#DCC5EA' },  // 보라
  '300,000원': { row: '#FDF3E7', cell: '#F5D3A8' },  // 주황
  '500,000원': { row: '#FCEBEC', cell: '#F2C2C6' }   // 빨강
};
var DEFAULT_COLOR = { row: '#FFFFFF', cell: '#EEEEEE' };

// 고료 요청이 의심되면 노란색으로 표시할 칸
var FEE_CHECK_COLUMNS = ['이름', '요청사항'];
var FEE_HIGHLIGHT = '#FFE14D';
var FEE_PATTERN = /(원고료|고료|단가|비용|금액|협찬비|광고비|제작비|수수료|페이(?!지)|만\s*원|\d[\d,]*\s*원|\bfee\b|\bpay\b|\bcost\b)/i;

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
  lock.waitLock(20000);
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    appendApplication_(data);
    return json_({ ok: true, result: 'success' });
  } catch (err) {
    return json_({ ok: false, result: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 브라우저로 웹 앱 URL을 열었을 때 연결 확인용
function doGet() {
  return json_({ ok: true, sheet: SHEET_NAME });
}

function appendApplication_(data) {
  var sheet = getOrCreateSheet();
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var brandsArr = Array.isArray(data.brands) ? data.brands : (data.brands ? [data.brands] : []);
  // 페이지는 전체 지원일 때도 그 페이지의 브랜드 목록을 보낸다. 목록이 비어 온 경우에만 기존 브랜드 전체로 본다.
  if (data.support_type === '전체 지원' && brandsArr.length === 0) brandsArr = BRAND_COLUMNS.slice();

  // 페이지에서 배송지 주소와 상세 주소를 합쳐 보내지만, 따로 오는 경우에도 한 칸에 합친다.
  var address = [data.address, data.addressDetail].filter(function (v) { return v; }).join(' ');

  var values = {
    '제출일시': Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
    '구분': data.division || '',
    '고료': data.amount || '',
    '지원유형': data.support_type || '',
    '이름': data.name || '',
    '인스타그램': instagramUrl(data.instagram),
    '휴대폰': String(data.phone || ''),
    '이메일': data.email || '',
    '우편번호': String(data.zonecode || ''),
    '배송지 주소': address,
    '요청사항': data.notes || '',
    '동의여부': data.agree || ''
  };
  Object.keys(values).forEach(function (k) {
    if (typeof values[k] === 'string') values[k] = values[k].trim();
  });
  ALL_BRAND_COLUMNS.forEach(function (col) {
    values[col] = brandsArr.indexOf(col) > -1;
  });

  var row = header.map(function (name) {
    return Object.prototype.hasOwnProperty.call(values, name) ? values[name] : '';
  });

  var nextRow = sheet.getLastRow() + 1;
  var range = sheet.getRange(nextRow, 1, 1, row.length);

  // 앞자리 0 보존: 값을 넣기 전에 텍스트 서식으로 지정
  TEXT_COLUMNS.forEach(function (name) {
    var c = header.indexOf(name);
    if (c > -1) sheet.getRange(nextRow, c + 1).setNumberFormat('@');
  });
  range.setValues([row]);
  range.setVerticalAlignment('middle').setWrap(true);
  ALL_BRAND_COLUMNS.forEach(function (col) {
    var c = header.indexOf(col);
    if (c > -1) sheet.getRange(nextRow, c + 1).insertCheckboxes();
  });

  // 고료별 색
  var color = PRICE_COLORS[values['고료']] || DEFAULT_COLOR;
  range.setBackground(color.row);
  var priceCol = header.indexOf('고료');
  if (priceCol > -1) {
    sheet.getRange(nextRow, priceCol + 1).setBackground(color.cell).setFontWeight('bold').setHorizontalAlignment('center');
  }

  // 이름·요청사항에 고료 요청이 있으면 노란색
  FEE_CHECK_COLUMNS.forEach(function (name) {
    var c = header.indexOf(name);
    if (c > -1 && FEE_PATTERN.test(values[name])) {
      sheet.getRange(nextRow, c + 1).setBackground(FEE_HIGHLIGHT).setFontWeight('bold');
    }
  });
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    createHeader_(sheet);
  } else {
    ensureColumns(sheet);
  }
  return sheet;
}

function createHeader_(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  styleHeader_(sheet.getRange(1, 1, 1, HEADERS.length));
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 34);
  HEADERS.forEach(function (name, i) {
    sheet.setColumnWidth(i + 1, COLUMN_WIDTHS[name] || 110);
    // 휴대폰·우편번호 열 전체를 텍스트 서식으로
    if (TEXT_COLUMNS.indexOf(name) > -1 && sheet.getMaxRows() > 1) {
      sheet.getRange(2, i + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
    }
  });
}

function styleHeader_(range) {
  range.setFontWeight('bold')
    .setBackground('#B5391D')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
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
    styleHeader_(sheet.getRange(1, prev + 1).setValue(col));
    header.splice(prev, 0, col);
  });
  HEADERS.forEach(function (col) {
    if (header.indexOf(col) > -1) return;
    header.push(col);
    if (sheet.getMaxColumns() < header.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), header.length - sheet.getMaxColumns());
    styleHeader_(sheet.getRange(1, header.length).setValue(col));
    sheet.setColumnWidth(header.length, COLUMN_WIDTHS[col] || 110);
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 테스트용 — Apps Script 편집기에서 이 함수를 선택해 [실행]하면
 * 시트에 샘플 4줄(5만·10만 고료요청·20만·50만)이 들어갑니다. 확인 후 행을 지워주세요.
 */
function testAppend() {
  var brands = ['루솜 LUSOM', 'LOYR (로이어)', 'it:fu (잇퓨)', 'PURCELL (퍼셀)', 'FROM9 (프롬나인)', 'Re:Nnocent (리노센트랩)'];
  var base = {
    division: '신규',
    instagram: '@test',
    phone: '010-0123-4567',
    email: 'test@example.com',
    zonecode: '06035',
    address: '서울 강남구 가로수길 5 101호',
    agree: '동의'
  };
  [
    Object.assign({}, base, { amount: '50,000원',  support_type: '개별 지원', brands: ['FROM9 (프롬나인)'], name: '테스트A', notes: '' }),
    Object.assign({}, base, { amount: '100,000원', support_type: '개별 지원', brands: ['루솜 LUSOM', 'PURCELL (퍼셀)'], name: '테스트B', notes: '고료 15만원 희망합니다' }),
    Object.assign({}, base, { amount: '200,000원', support_type: '전체 지원', brands: brands, name: '테스트C', notes: '' }),
    Object.assign({}, base, { amount: '500,000원', support_type: '개별 지원', brands: ['Re:Nnocent (리노센트랩)'], name: '테스트D', notes: '' })
  ].forEach(function (d) { appendApplication_(d); });
}
