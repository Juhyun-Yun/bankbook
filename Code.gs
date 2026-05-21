/**
 * 우리 반 통장 & 용돈기입장 — Apps Script Backend
 *
 * 시트 4개 탭을 단일 진실원천으로 사용합니다.
 *  - 명단        : 이름 / 적립통장잠금 / 용돈기입장잠금 / 용돈기간시작 / 용돈기간종료
 *  - 적립_활동   : 학생이름 / 활동ID / 활동이름 / 시작일 / 종료일 / 회당포인트
 *  - 적립_기록   : 학생이름 / 기록ID / 날짜 / 활동이름 / 포인트 / 확인
 *  - 용돈_기록   : 학생이름 / 기록ID / 날짜 / 내용 / 금액 / 종류 / 수입종류 / 확인
 */

const SHEET_ROSTER    = '명단';
const SHEET_SAVE_ACT  = '적립_활동';
const SHEET_SAVE_REC  = '적립_기록';
const SHEET_ALLOW_REC = '용돈_기록';
const SHEET_SETTINGS  = '설정';

const HEADERS = {
  [SHEET_ROSTER]:    ['이름','용돈기간시작','용돈기간종료'],
  [SHEET_SAVE_ACT]:  ['학생이름','활동ID','활동이름','시작일','종료일','회당포인트'],
  [SHEET_SAVE_REC]:  ['학생이름','기록ID','날짜','활동이름','포인트','확인'],
  [SHEET_ALLOW_REC]: ['학생이름','기록ID','날짜','내용','금액','종류','수입종류','확인'],
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
 * - 우선순위 1: 스크립트 속성에 저장된 SPREADSHEET ID 가 있으면 그걸로 열기
 *   (선생님 메뉴에서 시트 URL 을 입력해두면 이 경로로 동작)
 * - 우선순위 2: 시트에 바인딩된 스크립트라면 getActiveSpreadsheet()
 * - 둘 다 안 되면 안내가 담긴 Error 를 던짐
 */
function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID);
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch(e) {
      throw new Error('저장된 스프레드시트 ID 로 열 수 없습니다. (' + e.message + ') — 선생님 메뉴에서 ID/URL 을 다시 확인해 주세요.');
    }
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('연결된 스프레드시트가 없습니다. 선생님 메뉴에서 스프레드시트 URL 을 입력하거나, 스프레드시트 → 확장 프로그램 → Apps Script 에서 다시 만들어 주세요.');
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
  return getMeta();
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
  const acts = rowsAsObjects_(getSheet_(SHEET_SAVE_ACT))
    .filter(r => String(r['학생이름']||'').trim() === name)
    .map(r => ({
      id: String(r['활동ID']||''),
      name: String(r['활동이름']||''),
      startDate: fmtDate_(r['시작일']),
      endDate: fmtDate_(r['종료일']),
      points: Number(r['회당포인트']||0),
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
    !!entry.confirmed
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
    .createMenu('🐷 통장 앱')
    .addItem('시트 초기화', 'initSheets')
    .addItem('비밀번호 초기화 (1234)', 'resetPasswordToDefault_')
    .addToUi();
}
function resetPasswordToDefault_() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_PW);
  SpreadsheetApp.getActive().toast('비밀번호를 1234로 초기화했어요');
}
