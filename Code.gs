/**
 * ============================================================
 * © 2026 GEG화성 (깊이 e끌림). All rights reserved.
 *
 * 본 코드는 「저작권법」에 보호받는 저작물입니다.
 * - 복제권(제16조)·공중송신권(제18조)·배포권(제20조)은
 *   저작권자에게 있습니다.
 * - 정당 경로로 받은 이용자라도 코드의 무단 복제·재배포·
 *   재판매·리브랜딩은 허용되지 않습니다.
 * - 무단 이용 시 「저작권법」 제136조(5년 이하 징역 또는
 *   5천만 원 이하 벌금) 및 제125조(손해배상) 적용 대상이
 *   될 수 있습니다.
 * - 이용 문의: bacusiki777@gmail.com, for2102@jimj.kr
 * ============================================================
 */

// 빌드 서명
const _BUILD_SIG = 'GEGHS-DEEPE-2026';

// 출처 확인용 함수
function getBuildInfo() {
  return {
    sig: _BUILD_SIG,
    owner: 'GEG화성 (깊이 e끌림)',
    year: 2026
  };
}

/**
 * 스마트 통장 — Apps Script Backend
 *
 * 시트 4개 탭을 단일 진실원천으로 사용합니다.
 *  - 학생 명단   : 번호 / 이름 / 용돈기간시작 / 용돈기간종료
 *  - 적립_활동   : 학생이름 / 활동ID / 활동이름 / 시작일 / 종료일 / 회당포인트
 *  - 적립_기록   : 학생이름 / 기록ID / 날짜 / 활동이름 / 포인트 / 확인
 *  - 용돈_기록   : 학생이름 / 기록ID / 날짜 / 내용 / 금액 / 종류 / 수입종류 / 확인
 */

const SHEET_GUIDE     = '사용 설명';
// 예전 안내 탭 이름들 — 있으면 정리하고 '사용 설명' 하나로 통일
const GUIDE_LEGACY_NAMES = ['선생님 가이드', '📘 사용법', '사용법', '사용 안내', '📘 사용 설명'];
const SHEET_ROSTER    = '학생 명단';
const SHEET_SAVE_ACT  = '적립_활동';
const SHEET_SAVE_REC  = '적립_기록';
const SHEET_ALLOW_REC = '용돈_기록';
const SHEET_SETTINGS  = '설정';
const SHEET_STATS     = '통계';

const HEADERS = {
  [SHEET_ROSTER]:    ['번호','이름','용돈기간시작','용돈기간종료'],
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

// 탭 표시 순서 (앞에서부터). 학생 명단 오른쪽에 통계.
// '적립_활동'은 목록에 넣지 않고 숨긴다(적립통장이 활동을 저장하는 탭이라 삭제하지 않음).
const SHEET_ORDER = ['사용 설명', '학생 명단', '통계', '용돈_기록', '적립_기록', '설정'];
// 정리(삭제) 대상 군더더기 탭 — 기본 빈 시트, 예전 이름 '명단'(현재는 '학생 명단' 사용)
const JUNK_SHEET_NAMES = ['Sheet1', 'Sheet2', 'Sheet3', '시트1', '시트2', '시트3', 'Sheet', '명단'];

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
  const p = (e && e.parameter) || {};
  // 앱 화면(공용 주소)에서 넘어온 데이터 요청이면 데이터로 응답
  if (p.action) return handleApi_(p);
  // 그 외에는 기존처럼 앱 화면을 그대로 서빙 (주소를 직접 열었을 때 대비)
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('🏦 스마트 통장')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 앱 화면에서 온 요청도 같은 방식으로 처리 (쓰기 포함)
function doPost(e) {
  const p = (e && e.parameter) || {};
  return handleApi_(p);
}

/**
 * 공용 주소에 올린 앱 화면과 이 스프레드시트를 이어 주는 창구.
 * action 이름에 맞는 함수를 찾아 실행하고 결과를 돌려준다.
 * (callback 이 있으면 그 이름으로 감싸 응답 — 다른 주소의 화면에서도 읽을 수 있게)
 */
function handleApi_(p) {
  let result;
  try {
    const fn = API_ROUTES_[p.action];
    if (!fn) throw new Error('알 수 없는 요청입니다: ' + p.action);
    const args = p.args ? JSON.parse(p.args) : [];
    result = { ok: true, data: fn.apply(null, args) };
  } catch (err) {
    result = { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
  const body = JSON.stringify(result);
  if (p.callback) {
    return ContentService
      .createTextOutput(p.callback + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

// 앱 화면이 부를 수 있는 함수 목록 (기존 함수를 그대로 연결)
const API_ROUTES_ = {
  getMeta:               getMeta,
  initSheets:            initSheets,
  getRoster:             getRoster,
  getStudentData:        getStudentData,
  saveSavingsActivity:   saveSavingsActivity,
  deleteSavingsActivity: deleteSavingsActivity,
  saveSavingsRecord:     saveSavingsRecord,
  deleteSavingsRecord:   deleteSavingsRecord,
  setAllowancePeriod:    setAllowancePeriod,
  saveAllowanceEntry:    saveAllowanceEntry,
  deleteAllowanceEntry:  deleteAllowanceEntry,
  confirmRecord:         confirmRecord,
  getAllRecords:         getAllRecords,
  getTeacherPassword:    getTeacherPassword,
  setTeacherPassword:    setTeacherPassword,
  getBankName:           getBankName,
  setBankName:           setBankName,
  getSpreadsheetId:      getSpreadsheetId,
  setSpreadsheetId:      setSpreadsheetId,
  getDiagnostic:         getDiagnostic,
  getSettings:           getSettings,
  setSetting:            setSetting,
  getExpensePlan:        getExpensePlan,
  setExpensePlan:        setExpensePlan,
  getReflection:         getReflection,
  setReflection:         setReflection,
  createGuide:           createGuide,
};

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
      .setBackground('#E8F1FE')
      .setFontColor('#0F2C4D');
    sh.setFrozenRows(1);
    if (name === SHEET_SAVE_REC)  sh.getRange('F2:F').insertCheckboxes();
    if (name === SHEET_ALLOW_REC) sh.getRange('H2:H').insertCheckboxes();
    // 날짜 포맷
    if (name === SHEET_ROSTER) {
      sh.getRange('C2:D').setNumberFormat('yyyy-mm-dd');  // 용돈기간시작/종료
      // 번호(1~30)와 기본 이름(학생1~학생30)을 미리 채운다. 선생님이 실제 이름으로 덮어쓰면 됨
      const rosterRows = [];
      for (let k = 1; k <= 30; k++) rosterRows.push([k, '학생' + k]);
      sh.getRange(2, 1, rosterRows.length, 2).setValues(rosterRows);
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
  try { ensureRosterColumns_(); } catch(_) {}  // 명단 번호 열 보정 (A=번호, B=이름)
  try { ensureGuideSheet_(); } catch(_) {}     // 사용 설명 탭
  try { ensureStatsSheet_(); } catch(_) {}     // 통계 탭 (자동 합계)
  try { arrangeSheets_(); } catch(_) {}        // 탭 순서 정리 + 군더더기 탭 삭제
  return getMeta();
}

/**
 * 명단 탭을 A=번호, B=이름 구조로 맞춘다.
 * 예전 사본(A=이름)이면 앞에 번호 열을 새로 넣어 이관한다.
 */
function ensureRosterColumns_() {
  const sh = getSheet_(SHEET_ROSTER);
  const a1 = String(sh.getRange(1, 1).getValue() || '').trim();
  const b1 = String(sh.getRange(1, 2).getValue() || '').trim();
  // 예전 구조(A열=이름)면 앞에 번호 열을 새로 넣어 이관한다
  if (a1 === '이름' && b1 !== '이름') sh.insertColumnBefore(1);
  // 헤더 보정 (A=번호, B=이름)
  sh.getRange(1, 1).setValue('번호');
  sh.getRange(1, 2).setValue('이름');
  // 이름 칸(B2:B31)이 모두 비어 있으면 학생1~학생30을 기본으로 채운다.
  // 실제 학생 이름이 하나라도 있으면 절대 건드리지 않는다.
  const names = sh.getRange(2, 2, 30, 1).getValues();
  const hasName = names.some(function(r) { return String(r[0] || '').trim() !== ''; });
  if (!hasName) {
    const rows = [];
    for (let i = 1; i <= 30; i++) rows.push([i, '학생' + i]);
    sh.getRange(2, 1, 30, 2).setValues(rows);
  }
}

/**
 * 군더더기 탭(기본 빈 시트·옛 안내 탭)을 삭제하고, 표준 탭을 정해진 순서로 정렬한다.
 * 표준 7개 탭은 삭제하지 않는다.
 */
function arrangeSheets_() {
  const ss = getSpreadsheet_();
  const removeSet = JUNK_SHEET_NAMES.concat(GUIDE_LEGACY_NAMES);
  ss.getSheets().forEach(function(s) {
    if (removeSet.indexOf(s.getName()) >= 0 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch(_) {}
    }
  });
  SHEET_ORDER.forEach(function(name, idx) {
    const s = ss.getSheetByName(name);
    if (s) { s.activate(); ss.moveActiveSheet(idx + 1); }
  });
  // '적립_활동' 탭은 화면에서 숨긴다 (적립통장 기능이 이 탭에 활동을 저장하므로 삭제하지 않음)
  const actSheet = ss.getSheetByName(SHEET_SAVE_ACT);
  if (actSheet) { try { actSheet.hideSheet(); } catch(_) {} }
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
      `=IFERROR('학생 명단'!B${i}, "")`,
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
 * 사용 설명 — 스프레드시트 맨 앞에 두고 처음 받는 선생님이 읽을 수 있게
 * =======================================================================*/
function createGuide() {
  return ensureGuideSheet_();
}

function ensureGuideSheet_() {
  const ss = getSpreadsheet_();
  // 예전 안내 탭(옛 이름 포함)과 기존 '사용 설명' 탭을 모두 정리한 뒤 새로 만든다.
  // 시트가 이것들뿐이라 마지막 하나를 못 지우는 경우를 대비해 임시 탭을 먼저 만든다.
  const removeNames = GUIDE_LEGACY_NAMES.concat([SHEET_GUIDE]);
  const tmp = ss.insertSheet('__guide_tmp__' + Utilities.getUuid().slice(0, 4));
  removeNames.forEach(function(nm) {
    const s = ss.getSheetByName(nm);
    if (s) { try { ss.deleteSheet(s); } catch(_) {} }
  });
  tmp.setName(SHEET_GUIDE);
  tmp.activate();
  ss.moveActiveSheet(1);            // 첫 번째 위치로
  populateGuideSheet_(tmp);
  return { ok: true, name: SHEET_GUIDE };
}

function populateGuideSheet_(sh) {
  sh.clear();
  sh.clearFormats();

  // 앱 대표 색 (index.html 과 같은 계열)
  const C_TITLE_TX = '#0F2C4D';
  const C_TITLE_BG = '#E8F1FE';
  const C_CHAP_TX  = '#0F2C4D';
  const C_CHAP_BG  = '#E8F1FE';
  const C_THEAD_BG = '#DCE6F2';
  const C_ACCENT   = '#1B64DA';
  const C_MUTED    = '#5F6B7A';
  const C_BORDER   = '#B7C4D6';

  // 안내 내용 — 이모티콘 없이. 챕터/단계 번호는 순서대로 자동으로 매겨진다.
  //  ['title'|'chapter'|'step'|'body'|'note'|'qa'|'thead'|'trow'|'gap', ...]
  const spec = [
    ['title', '스마트 통장 — 사용 설명'],
    ['gap'],
    ['body', '이 스프레드시트는 우리 반 학생들의 통장 데이터를 관리하는 곳입니다.'],
    ['body', '사본을 받으신 분은 아래 순서대로 따라 하시면 우리 반 학생용 앱이 완성됩니다.'],
    ['body', '앱 화면은 학생용이고, 선생님이 할 일은 모두 이 스프레드시트에서 합니다.'],
    ['gap'],

    ['chapter', '처음 사용 — 기본 설정'],
    ['step', '사본 이름 바꾸기. 좌상단 파일 이름을 우리 반 이름으로 바꿉니다. 예를 들어 2026 행복초 4학년 3반 통장처럼 적습니다.'],
    ['step', '학생 명단 입력. 학생 명단 탭의 이름 칸에 우리 반 학생 이름을 한 줄씩 입력합니다. 용돈기간시작과 용돈기간종료는 비워 두어도 됩니다.'],
    ['step', '설정 채우기. 설정 탭에서 은행이름을 정하고, 적립통장사용과 용돈기입장사용 체크박스로 학생이 사용할 기능을 켭니다.'],
    ['gap'],

    ['chapter', '학생용 입장 주소 만들기'],
    ['step', '상단 메뉴에서 확장 프로그램을 눌러 Apps Script 편집기를 엽니다. 사본을 만들 때 코드가 함께 따라오므로 따로 붙여넣지 않아도 됩니다.'],
    ['step', '편집기 오른쪽 위 배포를 눌러 새 배포를 만들고 웹 앱을 고릅니다. 실행 사용자는 나 본인으로, 액세스 권한은 모든 사람 또는 학교 계정으로 두고 배포합니다.'],
    ['step', '처음이면 권한 검토 화면에서 본인 계정으로 허용합니다. 확인되지 않은 앱 경고가 뜨면 고급을 눌러 이동을 선택합니다. 본인이 만든 앱이라 안전합니다.'],
    ['step', '배포가 끝나면 웹 앱 주소가 화면에 나옵니다. 이 주소를 복사해 둡니다. 이 주소가 우리 반 시트로 통하는 열쇠입니다.'],
    ['step', '안내받은 앱 공용 주소를 브라우저에서 열고, 통장 표지 아래 설정에서 복사한 웹 앱 주소를 붙여넣어 연결합니다. 그 아래에 학생용 입장 주소가 만들어지면 복사해 학생들에게 나눠 줍니다.'],
    ['gap'],

    ['chapter', '선생님이 할 일 — 모두 이 스프레드시트에서'],
    ['body', '이 앱에는 선생님 모드가 없습니다. 모든 선생님 작업은 이 스프레드시트에서 직접 합니다.'],
    ['step', '도장 찍기. 적립_기록이나 용돈_기록 탭의 확인 열 체크박스를 켜면 학생 화면에도 도장이 표시됩니다.'],
    ['step', '설정 변경. 설정 탭에서 은행이름과 기능 사용 여부를 바꿉니다. 학생이 화면을 새로고침하면 반영됩니다.'],
    ['step', '학급 통계 보기. 통계 탭에 학생별 합계가 자동으로 계산되어 보입니다.'],
    ['gap'],

    ['chapter', '시트 구조 — 각 탭이 하는 일'],
    ['thead', '탭 이름', '하는 일'],
    ['trow', '사용 설명', '지금 보고 있는 이 안내입니다.'],
    ['trow', '학생 명단', '우리 반 학생 이름과 선택 항목인 용돈 기간을 적습니다.'],
    ['trow', '적립_기록', '학생이 신청한 적립 기록과 확인 체크박스가 있습니다.'],
    ['trow', '용돈_기록', '수입과 지출 사용 내역, 분류, 필요도, 확인 체크박스가 있습니다.'],
    ['trow', '설정', '학급 설정입니다. 은행이름과 적립통장사용, 용돈기입장사용을 둡니다.'],
    ['trow', '통계', '학생별 합계가 자동으로 계산되는 보기 전용 탭입니다.'],
    ['gap'],
    ['note', '데이터나 설정을 바꿀 때는 앱 화면이 아니라 해당 시트 탭에서 직접 수정하세요. 탭 이름은 코드에 연결되어 있으므로 삭제하거나 변경하지 마세요.'],
    ['gap'],

    ['chapter', '자주 묻는 질문'],
    ['qa', '학생 화면에 이름이 안 보여요.', '학생 명단 탭의 이름 칸이 비어 있지 않은지 확인하고, 학생이 화면을 새로고침해 최신 명단을 다시 받아오게 하세요.'],
    ['qa', '설정을 바꿨는데 학생 화면이 그대로예요.', '학생이 화면을 새로고침해야 새 설정이 반영됩니다.'],
    ['qa', '새 학년에 다시 쓰려면요?', '이 스프레드시트를 다시 사본으로 만들어 학년별로 따로 관리하거나, 명단을 새 학생으로 바꾸고 적립_기록과 용돈_기록의 데이터를 지우세요.'],
    ['qa', '다른 선생님께도 나눠주고 싶어요.', '이 스프레드시트 주소 끝의 edit 를 copy 로 바꾼 링크를 보내세요. 받은 분은 자기 사본을 만들고 이 안내대로 설정하면 됩니다.'],
    ['qa', '이 안내 내용을 새로 받고 싶어요.', "상단 메뉴 '스마트 통장' 에서 사용 설명 새로고침을 누르면 이 안내가 다시 만들어집니다."],
  ];

  // 자동 번호를 매기고 2차원 배열로 변환한다.
  const rows = [];
  const tags = [];
  let chap = 0, step = 0;
  let tableStart = -1, tableEnd = -1;
  spec.forEach(function(item) {
    const type = item[0];
    let a = '', b = '';
    if (type === 'title' || type === 'body' || type === 'note') { a = item[1]; }
    else if (type === 'chapter') { chap++; step = 0; a = chap + '. ' + item[1]; }
    else if (type === 'step')    { step++; a = chap + '-' + step + '. ' + item[1]; }
    else if (type === 'qa')      { a = 'Q. ' + item[1] + '\n' + 'A. ' + item[2]; }
    else if (type === 'thead' || type === 'trow') { a = item[1]; b = item[2]; }
    rows.push([a, b]);
    tags.push(type);
    const rn = rows.length;
    if (type === 'thead') tableStart = rn;
    if (type === 'thead' || type === 'trow') tableEnd = rn;
  });

  // 본문을 한 번에 입력하고, 긴 내용은 줄바꿈으로만 처리한다(행 높이 지정하지 않음).
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, rows.length, 2).setWrap(true).setVerticalAlignment('top');

  // 행별 스타일
  for (let i = 0; i < tags.length; i++) {
    const row = i + 1;
    const t = tags[i];
    if (t === 'thead') {
      sh.getRange(row, 1, 1, 2).setFontWeight('bold').setBackground(C_THEAD_BG);
      continue;
    }
    if (t === 'trow') continue;              // 표 본문은 테두리만
    // 표가 아닌 줄은 두 칸을 합쳐 넓게 보여준다
    sh.getRange(row, 1, 1, 2).merge();
    const r = sh.getRange(row, 1);
    if (t === 'title')        r.setFontSize(14).setFontWeight('bold').setFontColor(C_TITLE_TX).setBackground(C_TITLE_BG);
    else if (t === 'chapter') r.setFontWeight('bold').setFontColor(C_CHAP_TX).setBackground(C_CHAP_BG);
    else if (t === 'qa')      r.setFontColor(C_ACCENT);
    else if (t === 'note')    r.setFontStyle('italic').setFontColor(C_MUTED);
  }

  // 표 전체 테두리
  if (tableStart > 0 && tableEnd >= tableStart) {
    sh.getRange(tableStart, 1, tableEnd - tableStart + 1, 2)
      .setBorder(true, true, true, true, true, true, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }

  // 내용에 맞는 열 너비
  sh.setColumnWidth(1, 300);
  sh.setColumnWidth(2, 560);
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
  try { ensureRosterColumns_(); } catch(_) {}  // A=번호, B=이름 구조 자동 보정
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
  if (nameIdx < 0) return { ok:false, error:'학생 명단 시트에 "이름" 헤더가 없습니다.' };

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
    .createMenu('🏦 스마트 통장')
    .addItem('시트 초기화 (전체 탭 생성)', 'initSheets')
    .addItem('사용 설명 새로고침', 'createGuide')
    .addItem('통계 시트 새로고침', 'createStats')
    .addSeparator()
    .addItem('비밀번호 초기화 (1234)', 'resetPasswordToDefault_')
    .addToUi();
}
function resetPasswordToDefault_() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_PW);
  SpreadsheetApp.getActive().toast('비밀번호를 1234로 초기화했어요');
}
