/**
 * 우리 반 통장 & 용돈기입장 — Apps Script Backend
 *
 * 시트 4개 탭을 단일 진실원천으로 사용합니다.
 *  - 명단        : 이름 / 적립통장잠금 / 용돈기입장잠금 / 용돈기간시작 / 용돈기간종료
 *  - 적립_활동   : 학생이름 / 활동ID / 활동이름 / 시작일 / 종료일 / 회당포인트
 *  - 적립_기록   : 학생이름 / 기록ID / 날짜 / 활동이름 / 포인트 / 확인
 *  - 용돈_기록   : 학생이름 / 기록ID / 날짜 / 내용 / 금액 / 종류 / 수입종류 / 확인
 */

const SHEET_GUIDE     = '선생님 가이드';
const SHEET_ROSTER    = '명단';
const SHEET_SAVE_ACT  = '적립_활동';
const SHEET_SAVE_REC  = '적립_기록';
const SHEET_ALLOW_REC = '용돈_기록';
const SHEET_SETTINGS  = '설정';
const SHEET_STATS     = '통계';

const HEADERS = {
  [SHEET_ROSTER]:    ['이름','용돈기간시작','용돈기간종료'],
  [SHEET_SAVE_ACT]:  ['학생이름','활동ID','활동이름','시작일','종료일','회당포인트'],
  [SHEET_SAVE_REC]:  ['학생이름','기록ID','날짜','활동이름','포인트','확인'],
  [SHEET_ALLOW_REC]: ['학생이름','기록ID','날짜','내용','금액','종류','수입종류','확인','분류','필요도'],
  [SHEET_SETTINGS]:  ['항목','값','설명'],
};

// 설정 시트 기본 행 — 반 전체에 적용되는 글로벌 설정
const DEFAULT_SETTINGS = [
  ['은행이름',       '우리반은행', '학생 로그인 화면 통장에 표시되는 이름'],
  ['적립통장사용',   true,         '체크하면 학생이 적립통장 기능을 사용할 수 있어요'],
  ['용돈기입장사용', true,         '체크하면 학생이 용돈기입장 기능을 사용할 수 있어요'],
];

const PROP_PW    = 'teacher_password';
const DEFAULT_PW = '1234';
const PROP_BANK  = 'bank_name';
const DEFAULT_BANK = '우리반은행';
const PROP_SHEET_ID = 'spreadsheet_id';

/* ===========================================================================
 * doGet — Apps Script 웹앱 진입점
 *   URL 파라미터 ?sheet=<스프레드시트ID> 가 있으면 그 시트를 사용
 *   (다른 선생님들과 코드는 공유하지 않고 URL만 공유할 때 사용)
 * =======================================================================*/
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('우리 반 통장 & 용돈기입장')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ===========================================================================
 * 내부 유틸
 * =======================================================================*/

/**
 * 스프레드시트 핸들을 얻는다.
 * - 우선순위 1: 스크립트 속성에 저장된 SPREADSHEET ID
 * - 우선순위 2: 시트에 바인딩된 스크립트라면 getActiveSpreadsheet()
 */
function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID);
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch(e) {
      throw new Error('저장된 스프레드시트 ID 로 열 수 없습니다. (' + e.message + ')');
    }
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('연결된 스프레드시트가 없습니다. 스프레드시트 → 확장 프로그램 → Apps Script 에서 다시 만들어 주세요.');
  }
  return active;
}

function getSheet_(name) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEADERS[name].length)
      .setValues([HEADERS[name]])
      .setFontWeight('bold')
      .setBackground('#FFF3E0');
    sh.setFrozenRows(1);
    if (name === SHEET_SAVE_REC)  sh.getRange('F2:F').insertCheckboxes();
    if (name === SHEET_ALLOW_REC) sh.getRange('H2:H').insertCheckboxes();
    // 날짜 포맷
    if (name === SHEET_ROSTER) {
      sh.getRange('B2:C').setNumberFormat('yyyy-mm-dd');  // 용돈기간시작/종료
    }
    if (name === SHEET_SAVE_ACT) {
      sh.getRange('D2:E').setNumberFormat('yyyy-mm-dd');
    }
    if (name === SHEET_SAVE_REC) {
      sh.getRange('C2:C').setNumberFormat('yyyy-mm-dd');
    }
    if (name === SHEET_ALLOW_REC) {
      sh.getRange('C2:C').setNumberFormat('yyyy-mm-dd');
    }
    // 설정 시트 — 기본 행 채우고 체크박스 설정
    if (name === SHEET_SETTINGS) {
      const rows = DEFAULT_SETTINGS.map(r => [r[0], r[1], r[2]]);
      sh.getRange(2, 1, rows.length, 3).setValues(rows);
      // '적립통장사용' (3행), '용돈기입장사용' (4행) 의 값 칸을 체크박스로
      sh.getRange('B3:B4').insertCheckboxes();
      sh.setColumnWidth(1, 140);
      sh.setColumnWidth(2, 160);
      sh.setColumnWidth(3, 380);
    }
  }
  return sh;
}

function rowsAsObjects_(sh) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i].every(c => c === '' || c === null)) continue;
    const obj = { _row: i + 1 };
    headers.forEach((h, j) => obj[h] = data[i][j]);
    out.push(obj);
  }
  return out;
}

function fmtDate_(v) {
  if (!v && v !== 0) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v).trim();
  // 이미 yyyy-mm-dd 형태면 그대로
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(s)) return s;
  // 기타 파싱 시도
  const d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return s;
}

function findRow_(sh, keyCol, key, key2Col, key2) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return -1;
  const headers = data[0];
  const idx  = headers.indexOf(keyCol);
  const idx2 = key2Col ? headers.indexOf(key2Col) : -1;
  if (idx < 0) return -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idx]).trim() !== String(key).trim()) continue;
    if (idx2 >= 0 && String(data[i][idx2]).trim() !== String(key2).trim()) continue;
    return i + 1;
  }
  return -1;
}

function newId_(prefix) { return prefix + '_' + Utilities.getUuid().slice(0, 8); }

/* ===========================================================================
 * 메타 / 초기화
 * =======================================================================*/
function getMeta() {
  try {
    const ss = getSpreadsheet_();
    const allReady = Object.keys(HEADERS).every(name => !!ss.getSheetByName(name));
    const id = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID) || ss.getId();
    return {
      ok: true,
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      spreadsheetId: id,
      boundMode: !PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID),
      sheetsInitialized: allReady,
      timezone: Session.getScriptTimeZone(),
    };
  } catch(e) {
    return {
      ok: false,
      error: e.message || String(e),
      spreadsheetId: PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID) || '',
      timezone: Session.getScriptTimeZone(),
    };
  }
}

function initSheets() {
  Object.keys(HEADERS).forEach(name => getSheet_(name));
  try { ensureGuideSheet_(); } catch(_) {}  // 가이드 시트
  try { ensureStatsSheet_(); } catch(_) {}  // 통계 시트 (자동 합계)
  return getMeta();
}

/* ===========================================================================
 * 통계 시트 — 학생별 합계가 수식으로 자동 계산되는 보기 전용 시트
 * =======================================================================*/
function createStats() {
  return ensureStatsSheet_();
}

function ensureStatsSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_STATS);
  const isNew = !sh;
  if (isNew) sh = ss.insertSheet(SHEET_STATS);
  // 항상 새로 작성 (수식이 바뀔 수 있어서)
  sh.clear();
  sh.clearFormats();

  // 헤더
  const headers = ['이름','확인 포인트','대기 포인트','총 수입','총 지출','잔액','기록 수'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#E8F1FE').setFontColor('#0F2C4D');
  sh.setFrozenRows(1);

  // 100명까지 자동 계산 수식 (명단의 학생 이름이 채워지면 행이 자동으로 채워짐)
  // 적립_기록   컬럼: A=학생이름  E=포인트  F=확인
  // 용돈_기록   컬럼: A=학생이름  E=금액    F=종류(income/expense)
  const ROWS = 100;
  const formulas = [];
  for (let i = 2; i <= ROWS + 1; i++) {
    formulas.push([
      `=IFERROR(명단!A${i}, "")`,
      `=IF(A${i}="", "", SUMIFS('적립_기록'!E:E, '적립_기록'!A:A, A${i}, '적립_기록'!F:F, TRUE))`,
      `=IF(A${i}="", "", SUMIFS('적립_기록'!E:E, '적립_기록'!A:A, A${i}, '적립_기록'!F:F, FALSE))`,
      `=IF(A${i}="", "", SUMIFS('용돈_기록'!E:E, '용돈_기록'!A:A, A${i}, '용돈_기록'!F:F, "income"))`,
      `=IF(A${i}="", "", SUMIFS('용돈_기록'!E:E, '용돈_기록'!A:A, A${i}, '용돈_기록'!F:F, "expense"))`,
      `=IF(A${i}="", "", D${i}-E${i})`,
      `=IF(A${i}="", "", COUNTIF('적립_기록'!A:A, A${i}) + COUNTIF('용돈_기록'!A:A, A${i}))`,
    ]);
  }
  sh.getRange(2, 1, formulas.length, headers.length).setFormulas(formulas);

  // 컬럼 너비 / 정렬
  sh.setColumnWidth(1, 100);
  [2,3,4,5,6].forEach(c => sh.setColumnWidth(c, 120));
  sh.setColumnWidth(7, 80);
  sh.getRange(2, 2, ROWS, headers.length - 1).setHorizontalAlignment('right');

  // 헤더 옆 안내
  sh.getRange(1, 9).setValue('자동 계산되는 보기 전용 시트입니다. 명단에 학생을 추가하면 자동으로 줄이 채워져요.')
    .setFontStyle('italic').setFontColor('#5F6B7A');

  return { ok: true, name: SHEET_STATS };
}

/* ===========================================================================
 * 선생님 가이드 — 스프레드시트 맨 앞에 두고 처음 받는 선생님이 읽을 수 있게
 * =======================================================================*/
function createGuide() {
  return ensureGuideSheet_();
}

function ensureGuideSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_GUIDE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_GUIDE, 0);  // 첫 번째 위치에 신규 생성
  }
  // 이미 있다면 첫 위치로 이동
  if (sh.getIndex() > 1) {
    sh.activate();
    ss.moveActiveSheet(1);
  }
  populateGuideSheet_(sh);
  return { ok: true, name: SHEET_GUIDE };
}

function populateGuideSheet_(sh) {
  sh.clear();
  sh.clearFormats();

  let appUrl = '';
  try { appUrl = ScriptApp.getService().getUrl() || ''; } catch(_) {}

  // [본문, 스타일]  — 스타일: title / section / step / qa / null(본문)
  const L = [];
  const add = (text, style) => L.push([text, style || null]);

  add('🏦  우리반 통장 · 용돈기입장  —  사용 안내', 'title');
  add('');
  add('이 스프레드시트는 학생들의 통장 데이터를 관리하는 곳이에요.');
  add('사본을 받으신 분: 아래 순서대로 따라 하시면 30분 안에 학급용 앱이 완성됩니다.');
  add('앱은 학생 전용이고, 선생님이 할 일은 모두 이 스프레드시트에서 하세요.');
  add('');

  add('① 처음 사용 — 5분 설정', 'section');
  add('');
  add('1-1. 사본 이름 바꾸기', 'step');
  add('   좌상단 파일 이름을 클릭해 우리 반 이름으로 바꿔주세요.');
  add('   예) "2026 행복초 4학년 3반 통장"');
  add('');
  add('1-2. 학생 명단 입력', 'step');
  add('   "명단" 시트의 "이름" 칸에 우리 반 학생 이름을 한 줄씩 입력.');
  add('   용돈기간시작/종료는 비워둬도 됩니다 (학생이 앱에서 직접 설정 가능).');
  add('');
  add('1-3. "설정" 시트 채우기', 'step');
  add('   • 은행이름: 학생 로그인 화면 통장에 보일 이름 (예: "우리반은행", "행복초 4-3 통장")');
  add('   • 적립통장사용 (체크박스): 체크하면 학생이 적립통장 기능 사용 가능');
  add('   • 용돈기입장사용 (체크박스): 체크하면 학생이 용돈기입장 기능 사용 가능');
  add('');
  add('1-4. (선택) 공통 활동 약정 만들기', 'step');
  add('   모든 학생이 함께 할 활동을 선생님이 미리 정해두고 싶을 때.');
  add('   "적립_활동" 시트에서:');
  add('   • 학생이름 칸에 "공통" 입력 (꼭 이렇게 적어야 인식)');
  add('   • 활동이름, 시작일, 종료일, 회당포인트 입력 (활동ID는 비워두면 자동 생성)');
  add('   학생은 공통 활동을 수정·삭제 못 하고, 자기 활동은 5개까지 따로 만들 수 있어요.');
  add('');

  add('② 학생용 URL 만들기 — 처음 한 번만', 'section');
  add('');
  add('2-1. Apps Script 편집기 열기', 'step');
  add('   상단 메뉴: 확장 프로그램 → Apps Script');
  add('   (사본을 만들 때 코드가 자동으로 함께 따라왔어요 — 따로 복사·붙여넣기 안 해도 됩니다)');
  add('');
  add('2-2. 웹 앱으로 배포', 'step');
  add('   Apps Script 편집기 우상단 [배포] → [새 배포]');
  add('   톱니바퀴 ⚙ → 웹 앱');
  add('   • 설명: 아무거나 (예: "v1")');
  add('   • 다음 사용자 인증 정보로 실행: 나(선생님 본인)');
  add('   • 액세스 권한이 있는 사용자: 모든 사람 (또는 학교 Google 계정)');
  add('   [배포] 클릭');
  add('');
  add('2-3. 권한 승인 (처음 한 번)', 'step');
  add('   "권한 검토" 화면이 떠요 → 본인 Google 계정 클릭');
  add('   "Google에서 확인하지 않은 앱" 경고가 뜨면:');
  add('   → [고급] 클릭 → [(앱 이름)(으)로 이동(안전하지 않음)] 클릭');
  add('   본인이 직접 만든 앱이라 안전합니다.');
  add('   → 권한 허용 → 배포 완료');
  add('');
  add('2-4. 학생용 URL 받기', 'step');
  add('   "웹 앱 URL" 이 화면에 표시됩니다 (https://script.google.com/macros/s/.../exec)');
  add('   이 URL을 학생들에게 공유 (구글 클래스룸·문자·QR코드 등).');
  add('   학생들은 통장 표지에서 자기 이름을 선택해 입장합니다.');
  add('');

  add('③ 선생님이 할 일 — 모두 이 스프레드시트에서', 'section');
  add('');
  add('이 앱에는 선생님 모드가 없어요. 모든 선생님 작업은 여기 시트에서 직접 합니다.');
  add('');
  add('[도장 찍기 (확인)]', 'step');
  add('   "적립_기록" 또는 "용돈_기록" 시트의 "확인" 열 체크박스를 켜세요.');
  add('   체크하는 순간 학생 화면에도 자동으로 도장이 표시됩니다.');
  add('');
  add('[설정 변경]', 'step');
  add('   "설정" 시트에서 은행이름·기능 사용 여부를 직접 수정.');
  add('   학생이 페이지를 새로고침해야 바뀐 값이 반영됩니다.');
  add('');
  add('[공통 활동 추가/수정]', 'step');
  add('   "적립_활동" 시트에서 학생이름 칸이 "공통"인 행을 추가/수정.');
  add('   학생 개인이 추가한 활동은 학생이름이 본인 이름으로 들어가요.');
  add('');
  add('[학급 통계 보기]', 'step');
  add('   "통계" 시트에 학생별 합계가 자동으로 계산되어 보입니다.');
  add('   (이름 · 확인 포인트 · 대기 포인트 · 총 수입 · 총 지출 · 잔액 · 기록 수)');
  add('   명단에 학생을 새로 추가하면 자동으로 줄이 채워져요.');
  add('');
  add('[메뉴 도구]', 'step');
  add('   스프레드시트 상단 메뉴 "🐷 우리반 통장 앱" 에서:');
  add('   • 시트 초기화 (전체 탭 생성)');
  add('   • 📘 선생님 가이드 새로고침 (이 안내 다시 만들기)');
  add('   • 📊 통계 시트 새로고침');
  add('');

  add('④ 시트 구조 — 각 탭이 하는 일', 'section');
  add('');
  add('  선생님 가이드  ← 지금 보고 계신 이 안내');
  add('  명단         ← 학급 학생 이름과 (선택) 용돈 기간');
  add('  적립_활동     ← 활동 약정 목록 (공통 + 학생 개인)');
  add('  적립_기록     ← 학생이 신청한 적립 기록 + "확인" 체크박스');
  add('  용돈_기록     ← 수입/지출 사용 내역 + 분류 + 필요도 + "확인" 체크박스');
  add('  설정         ← 학급 설정 (은행이름, 적립통장사용, 용돈기입장사용)');
  add('  통계         ← 학생별 합계 (자동 계산, 보기 전용)');
  add('');
  add('  ※ 시트 이름은 절대 바꾸지 마세요. 코드가 이 이름으로 시트를 찾습니다.');
  add('');

  add('⑤ 자주 묻는 질문', 'section');
  add('');
  add('Q. 학생들 화면에 이름이 안 보여요.', 'qa');
  add('   → "명단" 시트의 "이름" 칸이 비어있지 않은지 확인.');
  add('   → 학생이 페이지 새로고침(F5)으로 최신 명단을 다시 받아옵니다.');
  add('');
  add('Q. 설정을 바꿨는데 학생 화면이 그대로예요.', 'qa');
  add('   → 학생이 페이지를 새로고침해야 새 설정이 반영됩니다.');
  add('');
  add('Q. 새 학년에 다시 사용하려면?', 'qa');
  add('   → 가장 깔끔한 방법: 이 스프레드시트를 또 사본 만들기 → 학년별로 따로 관리');
  add('   → 또는: 명단을 새 학생으로 교체 + 적립_기록·용돈_기록 시트 데이터를 지움');
  add('');
  add('Q. 통계 시트의 숫자가 이상해요.', 'qa');
  add('   → 메뉴 [🐷 우리반 통장 앱 → 통계 시트 새로고침] 으로 수식을 다시 만들어주세요.');
  add('');
  add('Q. 이 안내(가이드) 내용이 옛날 거예요. 새로 받고 싶어요.', 'qa');
  add('   → 메뉴 [🐷 우리반 통장 앱 → 선생님 가이드 새로고침]');
  add('');
  add('Q. 다른 선생님께도 이 시스템을 나눠주고 싶어요.', 'qa');
  add('   → 이 스프레드시트 URL 끝의 /edit 를 /copy 로 바꿔 그 링크를 보내주세요.');
  add('   → 예) https://docs.google.com/spreadsheets/d/[ID]/copy');
  add('   → 받은 분은 자기 사본을 만들고, 이 가이드 안내대로 설정하면 됩니다.');
  add('   → 이메일 공유 같은 것 필요 없어요. 각자 자기 Google 계정 안에서만 동작합니다.');
  add('');
  add('Q. 코드를 직접 수정하고 싶어요.', 'qa');
  add('   → 확장 프로그램 → Apps Script 에서 자유롭게 수정 가능합니다.');
  add('   → 수정 후에는 [배포 → 배포 관리 → ✏ → 새 버전 → 배포] 로 반영하세요.');
  add('');

  // 본문 한 번에 입력
  const values = L.map(l => [l[0]]);
  sh.getRange(1, 1, values.length, 1).setValues(values);

  // 행별 스타일 적용
  for (let i = 0; i < L.length; i++) {
    const row = i + 1;
    const t = L[i][1];
    const r = sh.getRange(row, 1);
    if (t === 'title') {
      r.setFontSize(20).setFontWeight('bold').setBackground('#3182F6').setFontColor('#FFFFFF');
      sh.setRowHeight(row, 52);
    } else if (t === 'section') {
      r.setFontSize(14).setFontWeight('bold').setBackground('#E8F1FE').setFontColor('#0F2C4D');
      sh.setRowHeight(row, 38);
    } else if (t === 'step') {
      r.setFontWeight('bold').setBackground('#FFF8E1').setFontColor('#5D4500');
    } else if (t === 'qa') {
      r.setFontWeight('bold').setFontColor('#1B64DA');
    }
    r.setWrap(true);
  }

  sh.setColumnWidth(1, 900);
  sh.setHiddenGridlines(true);
}

/* ===========================================================================
 * 스프레드시트 ID 설정 (바인딩 안 된 standalone 스크립트도 사용 가능)
 * =======================================================================*/
function getSpreadsheetId() {
  return PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID) || '';
}

/**
 * 진단 정보 — 선생님 메뉴에서 정확히 어느 시트를 보고 있고
 * 거기에 무엇이 들어있는지 한 눈에 확인.
 */
function getDiagnostic() {
  const r = { time: new Date().toISOString() };
  const props = PropertiesService.getScriptProperties();
  r.savedSheetId = props.getProperty(PROP_SHEET_ID) || null;

  // 바인딩된 활성 시트(있다면)
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    r.boundActive = active ? {
      name: active.getName(),
      id: active.getId(),
      url: active.getUrl(),
    } : null;
  } catch(e) {
    r.boundActive = null;
  }

  // 실제로 사용 중인 시트 (저장된 ID 우선)
  try {
    const ss = getSpreadsheet_();
    r.connected = true;
    r.usingMode = r.savedSheetId ? 'SAVED_ID' : 'BOUND_ACTIVE';
    r.using = {
      name: ss.getName(),
      id: ss.getId(),
      url: ss.getUrl(),
    };
    // 파일 안 모든 시트 탭
    r.allSheetsInFile = ss.getSheets().map(s => ({
      name: s.getName(),
      rows: s.getLastRow(),
      cols: s.getLastColumn(),
    }));
    // 기대하는 4개 탭 상태
    r.expectedSheets = Object.keys(HEADERS).map(name => {
      const sh = ss.getSheetByName(name);
      if (!sh) return { name: name, exists: false };
      const data = sh.getDataRange().getValues();
      const actualHeaders = data[0] || [];
      return {
        name: name,
        exists: true,
        rows: data.length,
        actualHeaders: actualHeaders,
        expectedHeaders: HEADERS[name],
        headersMatch: JSON.stringify(actualHeaders) === JSON.stringify(HEADERS[name]),
      };
    });
    // 명단 샘플 (앞 8행)
    const roster = ss.getSheetByName(SHEET_ROSTER);
    if (roster) {
      const data = roster.getDataRange().getValues();
      r.rosterSample = data.slice(0, 8);
      r.rosterDataRows = Math.max(0, data.length - 1);
    } else {
      r.rosterDataRows = 0;
      r.rosterMissing = true;
    }
  } catch(e) {
    r.connected = false;
    r.error = e.message || String(e);
  }
  return r;
}

/**
 * 시트 URL 또는 ID 를 받아 저장한다. URL 이면 ID 추출.
 * 저장 후 실제로 openById 로 열어보고 헤더가 맞는지 가볍게 검증.
 */
function setSpreadsheetId(input) {
  let id = String(input || '').trim();
  if (!id) {
    PropertiesService.getScriptProperties().deleteProperty(PROP_SHEET_ID);
    return { ok: true, cleared: true };
  }
  // URL 이면 /d/{ID}/ 패턴에서 ID 추출
  const m = id.match(/\/d\/([a-zA-Z0-9_\-]+)/);
  if (m) id = m[1];
  try {
    const ss = SpreadsheetApp.openById(id);
    PropertiesService.getScriptProperties().setProperty(PROP_SHEET_ID, id);
    return { ok: true, id: id, name: ss.getName(), url: ss.getUrl() };
  } catch(e) {
    return { ok: false, error: '열 수 없습니다: ' + (e.message || e) + ' — ID/URL 과 접근 권한을 확인해 주세요.' };
  }
}

/* ===========================================================================
 * 명단 (Roster)
 * =======================================================================*/
function getRoster() {
  return rowsAsObjects_(getSheet_(SHEET_ROSTER))
    .map(r => ({
      name: String(r['이름']||'').trim(),
      allowanceStart: fmtDate_(r['용돈기간시작']),
      allowanceEnd:   fmtDate_(r['용돈기간종료']),
    }))
    .filter(r => r.name);
}

/* ===========================================================================
 * 한 학생의 모든 데이터
 * =======================================================================*/
function getStudentData(name) {
  name = String(name).trim();
  // 학생이름이 '공통' 인 활동은 선생님이 일괄 지정한 공통 약정 — 모든 학생에게 표시
  const acts = rowsAsObjects_(getSheet_(SHEET_SAVE_ACT))
    .filter(r => {
      const s = String(r['학생이름']||'').trim();
      return s === name || s === '공통';
    })
    .map(r => ({
      id: String(r['활동ID']||''),
      name: String(r['활동이름']||''),
      startDate: fmtDate_(r['시작일']),
      endDate: fmtDate_(r['종료일']),
      points: Number(r['회당포인트']||0),
      isCommon: String(r['학생이름']||'').trim() === '공통',
    }))
    .filter(a => a.id);

  const recs = rowsAsObjects_(getSheet_(SHEET_SAVE_REC))
    .filter(r => String(r['학생이름']||'').trim() === name)
    .map(r => ({
      id: String(r['기록ID']||''),
      date: fmtDate_(r['날짜']),
      activityName: String(r['활동이름']||''),
      points: Number(r['포인트']||0),
      confirmed: !!r['확인'],
    }))
    .filter(r => r.id);

  const entries = rowsAsObjects_(getSheet_(SHEET_ALLOW_REC))
    .filter(r => String(r['학생이름']||'').trim() === name)
    .map(r => ({
      id: String(r['기록ID']||''),
      date: fmtDate_(r['날짜']),
      description: String(r['내용']||''),
      amount: Number(r['금액']||0),
      kind: String(r['종류']||'income'),
      incomeType: String(r['수입종류']||'fixed'),
      confirmed: !!r['확인'],
      category: String(r['분류']||''),
      necessity: Number(r['필요도']||0),
    }))
    .filter(e => e.id);

  const rosterRow = rowsAsObjects_(getSheet_(SHEET_ROSTER))
    .find(r => String(r['이름']||'').trim() === name);
  const period = {
    start: rosterRow ? fmtDate_(rosterRow['용돈기간시작']) : '',
    end:   rosterRow ? fmtDate_(rosterRow['용돈기간종료']) : '',
  };

  return {
    savings:   { activities: acts, records: recs },
    allowance: { period, entries },
  };
}

/* ===========================================================================
 * 적립 통장 — 활동
 * =======================================================================*/
function saveSavingsActivity(name, activity) {
  const sh = getSheet_(SHEET_SAVE_ACT);
  const id = activity.id || newId_('a');
  const row = findRow_(sh, '학생이름', name, '활동ID', id);
  const values = [
    name, id,
    activity.name || '',
    activity.startDate || '',
    activity.endDate || '',
    Number(activity.points || 0)
  ];
  if (row > 0) {
    sh.getRange(row, 1, 1, values.length).setValues([values]);
  } else {
    sh.appendRow(values);
  }
  return { ok: true, id: id };
}

function deleteSavingsActivity(name, id) {
  const sh = getSheet_(SHEET_SAVE_ACT);
  const row = findRow_(sh, '학생이름', name, '활동ID', id);
  if (row > 0) sh.deleteRow(row);
  return { ok: true };
}

/* ===========================================================================
 * 적립 통장 — 기록
 * =======================================================================*/
function saveSavingsRecord(name, record) {
  const sh = getSheet_(SHEET_SAVE_REC);
  const id = record.id || newId_('r');
  const row = findRow_(sh, '학생이름', name, '기록ID', id);
  const values = [
    name, id,
    record.date || '',
    record.activityName || '',
    Number(record.points || 0),
    !!record.confirmed
  ];
  if (row > 0) {
    sh.getRange(row, 1, 1, values.length).setValues([values]);
  } else {
    sh.appendRow(values);
    const lastRow = sh.getLastRow();
    sh.getRange(lastRow, 6).insertCheckboxes();
  }
  return { ok: true, id: id };
}

function deleteSavingsRecord(name, id) {
  const sh = getSheet_(SHEET_SAVE_REC);
  const row = findRow_(sh, '학생이름', name, '기록ID', id);
  if (row > 0) sh.deleteRow(row);
  return { ok: true };
}

/* ===========================================================================
 * 용돈 기입장 — 기간 / 항목
 * =======================================================================*/
function setAllowancePeriod(name, start, end) {
  // 헤더 기반으로 위치를 찾아 안전하게 기록 — 기존 시트에 잠금 컬럼이 남아 있어도 동작
  const sh = getSheet_(SHEET_ROSTER);
  const data = sh.getDataRange().getValues();
  const headers = (data[0] || HEADERS[SHEET_ROSTER]).map(h => String(h).trim());
  const nameIdx  = headers.indexOf('이름');
  const startIdx = headers.indexOf('용돈기간시작');
  const endIdx   = headers.indexOf('용돈기간종료');
  if (nameIdx < 0) return { ok:false, error:'명단 시트에 "이름" 헤더가 없습니다.' };

  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][nameIdx]||'').trim() === String(name).trim()) { rowIdx = i + 1; break; }
  }
  if (rowIdx < 0) {
    const newRow = new Array(headers.length).fill('');
    newRow[nameIdx] = name;
    if (startIdx >= 0) newRow[startIdx] = start || '';
    if (endIdx   >= 0) newRow[endIdx]   = end   || '';
    sh.appendRow(newRow);
  } else {
    if (startIdx >= 0) sh.getRange(rowIdx, startIdx + 1).setValue(start || '');
    if (endIdx   >= 0) sh.getRange(rowIdx, endIdx + 1).setValue(end || '');
  }
  return { ok: true };
}

function saveAllowanceEntry(name, entry) {
  const sh = getSheet_(SHEET_ALLOW_REC);
  const id = entry.id || newId_('e');
  const row = findRow_(sh, '학생이름', name, '기록ID', id);
  const values = [
    name, id,
    entry.date || '',
    entry.description || '',
    Number(entry.amount || 0),
    entry.kind || 'income',
    entry.incomeType || 'fixed',
    !!entry.confirmed,
    entry.category || '',
    Number(entry.necessity || 0),
  ];
  if (row > 0) {
    sh.getRange(row, 1, 1, values.length).setValues([values]);
  } else {
    sh.appendRow(values);
    const lastRow = sh.getLastRow();
    sh.getRange(lastRow, 8).insertCheckboxes();
  }
  return { ok: true, id: id };
}

function deleteAllowanceEntry(name, id) {
  const sh = getSheet_(SHEET_ALLOW_REC);
  const row = findRow_(sh, '학생이름', name, '기록ID', id);
  if (row > 0) sh.deleteRow(row);
  return { ok: true };
}

/* ===========================================================================
 * 도장 확인 (선생님)
 * =======================================================================*/
function confirmRecord(name, kind, id, val) {
  if (kind === 'savings') {
    const sh = getSheet_(SHEET_SAVE_REC);
    const row = findRow_(sh, '학생이름', name, '기록ID', id);
    if (row > 0) sh.getRange(row, 6).setValue(!!val);
  } else if (kind === 'allowance') {
    const sh = getSheet_(SHEET_ALLOW_REC);
    const row = findRow_(sh, '학생이름', name, '기록ID', id);
    if (row > 0) sh.getRange(row, 8).setValue(!!val);
  }
  return { ok: true };
}

/* ===========================================================================
 * 선생님 실적 확인 화면용 — 전체 기록
 * =======================================================================*/
function getAllRecords() {
  const saveRecs = rowsAsObjects_(getSheet_(SHEET_SAVE_REC))
    .map(r => ({
      name: String(r['학생이름']||'').trim(),
      kind: 'savings',
      id: String(r['기록ID']||''),
      date: fmtDate_(r['날짜']),
      description: String(r['활동이름']||''),
      amount: Number(r['포인트']||0),
      confirmed: !!r['확인'],
    }))
    .filter(r => r.id);
  const allowRecs = rowsAsObjects_(getSheet_(SHEET_ALLOW_REC))
    .map(r => ({
      name: String(r['학생이름']||'').trim(),
      kind: 'allowance',
      id: String(r['기록ID']||''),
      date: fmtDate_(r['날짜']),
      description: String(r['내용']||''),
      amount: Number(r['금액']||0),
      entryKind: String(r['종류']||''),
      incomeType: String(r['수입종류']||''),
      confirmed: !!r['확인'],
    }))
    .filter(r => r.id);
  return saveRecs.concat(allowRecs);
}

/* ===========================================================================
 * 비밀번호
 * =======================================================================*/
function getTeacherPassword() {
  return PropertiesService.getScriptProperties().getProperty(PROP_PW) || DEFAULT_PW;
}
function setTeacherPassword(pw) {
  PropertiesService.getScriptProperties().setProperty(PROP_PW, String(pw||'').trim() || DEFAULT_PW);
  return { ok: true };
}

/* ===========================================================================
 * 전역 설정 (설정 시트가 단일 진실원천)
 *   - 은행이름 (text)
 *   - 적립통장사용 (checkbox)
 *   - 용돈기입장사용 (checkbox)
 * =======================================================================*/
function getSettings() {
  const sh = getSheet_(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  const m = {};
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][0]||'').trim();
    if (k) m[k] = data[i][1];
  }
  return {
    bankName: String(m['은행이름'] || '우리반은행'),
    savingsEnabled:   m['적립통장사용']   !== false,
    allowanceEnabled: m['용돈기입장사용'] !== false,
  };
}

function setSetting(key, value) {
  const sh = getSheet_(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]||'').trim() === String(key).trim()) { rowIdx = i + 1; break; }
  }
  if (rowIdx < 0) {
    sh.appendRow([key, value, '']);
    // 추가된 행이 체크박스성 키라면 체크박스로
    if (key === '적립통장사용' || key === '용돈기입장사용') {
      sh.getRange(sh.getLastRow(), 2).insertCheckboxes();
      sh.getRange(sh.getLastRow(), 2).setValue(!!value);
    }
  } else {
    sh.getRange(rowIdx, 2).setValue(value);
  }
  return { ok: true };
}

// 학생별 지출 계획 (이번 기간 동안의 자기 다짐 / 계획 자유 텍스트)
function getExpensePlan(name) {
  return PropertiesService.getScriptProperties().getProperty('plan_' + String(name||'').trim()) || '';
}
function setExpensePlan(name, plan) {
  const key = 'plan_' + String(name||'').trim();
  if (!plan) PropertiesService.getScriptProperties().deleteProperty(key);
  else PropertiesService.getScriptProperties().setProperty(key, String(plan));
  return { ok: true };
}

// 학생별 소감 / 다음 계획 적용 메모 (한 줄 조언을 보고 적는 자기 반성)
function getReflection(name) {
  return PropertiesService.getScriptProperties().getProperty('reflect_' + String(name||'').trim()) || '';
}
function setReflection(name, text) {
  const key = 'reflect_' + String(name||'').trim();
  if (!text) PropertiesService.getScriptProperties().deleteProperty(key);
  else PropertiesService.getScriptProperties().setProperty(key, String(text));
  return { ok: true };
}

// 은행 이름 — 설정 시트로 위임 (PropertiesService 값이 있다면 1회 이관)
function getBankName() {
  const settings = getSettings();
  if (settings.bankName) return settings.bankName;
  const legacy = PropertiesService.getScriptProperties().getProperty(PROP_BANK);
  return legacy || DEFAULT_BANK;
}
function setBankName(name) {
  setSetting('은행이름', String(name||'').trim() || DEFAULT_BANK);
  return { ok: true };
}

/* ===========================================================================
 * 메뉴 — 스프레드시트 열었을 때 도구 메뉴 표시 (옵션)
 * =======================================================================*/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🐷 우리반 통장 앱')
    .addItem('시트 초기화 (전체 탭 생성)', 'initSheets')
    .addItem('📘 선생님 가이드 새로고침', 'createGuide')
    .addItem('📊 통계 시트 새로고침', 'createStats')
    .addSeparator()
    .addItem('비밀번호 초기화 (1234)', 'resetPasswordToDefault_')
    .addToUi();
}
function resetPasswordToDefault_() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_PW);
  SpreadsheetApp.getActive().toast('비밀번호를 1234로 초기화했어요');
}
