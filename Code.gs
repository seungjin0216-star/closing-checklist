/**
 * 백석직영점 마감 체크리스트 - Google Apps Script 백엔드 (v3)
 * ------------------------------------------------------
 * 화면(UI)은 GitHub Pages의 index.html이 담당하고, 이 스크립트는
 * 순수 데이터 API 역할만 합니다.
 *
 * - GET  (JSONP)
 *     ?action=state&date=YYYY-MM-DD&callback=xxx     → 그날 체크 상태
 *     ?action=monthly&month=YYYY-MM&callback=xxx     → 이번달 청소 상태
 * - POST (JSON body)
 *     {action:'toggle', date, itemId, itemLabel, checked}
 *     {action:'complete', date}
 *     {action:'uncomplete', date}
 *     {action:'toggleMonthly', month, itemId, checked}
 *
 * 항목을 추가/수정/삭제하려면 아래 배열들을 고치고, index.html 쪽의
 * 동일한 배열도 같이 고쳐야 합니다 (두 곳에 항목 목록이 각각 있음).
 */

// ===== 채우기 (그룹 아코디언 + 낱개 항목, 화면에 보이는 순서 그대로) =====

const FILL_SEQUENCE = [
  { type: 'group', id: 'g_sauce', label: '소스류', items: [
    { id: 'gs1', label: '국수소스' },
    { id: 'gs2', label: '파장소스' },
    { id: 'gs3', label: '들기름' },
  ]},
  { type: 'group', id: 'g_left', label: '찬냉장고 하단[좌측]', items: [
    { id: 'gl1', label: '라면용대파' },
    { id: 'gl2', label: '다진고추' },
    { id: 'gl3', label: '다진대파' },
    { id: 'gl4', label: '전골용배추' },
    { id: 'gl5', label: '전골용대파' },
    { id: 'gl6', label: '단무지' },
    { id: 'gl7', label: '다대기' },
    { id: 'gl8', label: '다진마늘' },
    { id: 'gl9', label: '묵은지' },
    { id: 'gl10', label: '미나리' },
  ]},
  { type: 'group', id: 'g_right', label: '찬냉장고 하단[우측]', items: [
    { id: 'gr1', label: '천엽' },
    { id: 'gr2', label: '대파김치' },
    { id: 'gr3', label: '볶음밥김치' },
    { id: 'gr4', label: '깍두기' },
    { id: 'gr5', label: '양파소스' },
  ]},
  { type: 'group', id: 'g_six', label: '6구냉장고', items: [
    { id: 'g6_1', label: '전골용 내장' },
    { id: 'g6_2', label: '전골용 고추' },
    { id: 'g6_3', label: '콩나물' },
    { id: 'g6_4', label: '레몬' },
    { id: 'g6_5', label: '오이' },
  ]},
  { type: 'group', id: 'g_meat', label: '고기냉장고', items: [
    { id: 'gm1', label: '고기정리' },
    { id: 'gm2', label: '양파' },
    { id: 'gm3', label: '떡' },
  ]},
  { type: 'item', id: 'f_banchan', label: '찬냉장고 반찬리필' },
  { type: 'item', id: 'f_fireshow', label: '불쇼술 채우기' },
  { type: 'item', id: 'f_drink', label: '술, 음료 채우기' },
  { type: 'item', id: 'f_grape', label: '청포도, 이쑤시개 채우기' },
  { type: 'item', id: 'f_broth', label: '육수 확인' },
  { type: 'item', id: 'f_gammi', label: '감미 채우기' },
];

// ===== 청소 (매일) =====
const CLEAN_ITEMS = [
  { id: 'c1', label: '쓴 행주 전부 삶기 (테이블행주 포함)' },
  { id: 'c2', label: '술박스 정리하기' },
  { id: 'c3', label: '화구 라인 청소' },
  { id: 'c4', label: '식기세척기 끄고 청소' },
];

// ===== 정기 청소 (주 1회, 월요일에만 표시/필수) =====
const WEEKLY_ITEMS = [
  { id: 'w1', label: '덕트 청소', weekday: 1 }, // 1 = 월요일
];

// ===== 월간 청소 (한 달에 한 번, 마감 완료 조건에는 포함 안 함) =====
const MONTHLY_ITEMS = [
  { id: 'm1', label: '냉장고 성에제거 및 청소' },
  { id: 'm2', label: '주방바닥 퐁퐁청소' },
  { id: 'm3', label: '유리 청소' },
];

// ===== 마지막 확인 =====
const FINAL_ITEMS = [
  { id: 'x1', label: '고기·식자재 발주 무조건 하기' },
  { id: 'x2', label: '저녁 먹을 거 확인 후 주문' },
  { id: 'x3', label: '에어컨 끄기 (홀3, 주방1)' },
  { id: 'x4', label: '가스 끄기' },
  { id: 'x5', label: '마감정산·시재 맞추기' },
];

function flattenFillItems_() {
  const out = [];
  FILL_SEQUENCE.forEach(function (entry) {
    if (entry.type === 'group') {
      entry.items.forEach(function (it) { out.push({ id: it.id, label: it.label }); });
    } else {
      out.push({ id: entry.id, label: entry.label });
    }
  });
  return out;
}

// 그날그날 저장/조회 대상이 되는 전체 항목 (월간 항목 제외)
const ALL_DAILY_ITEMS = flattenFillItems_().concat(CLEAN_ITEMS, WEEKLY_ITEMS, FINAL_ITEMS);

// 그 날짜에 실제로 "완료하기"에 필요한 항목 (월요일이 아니면 주간 항목 제외)
function requiredItemsForDate_(dateStr) {
  const parts = dateStr.split('-').map(Number);
  const dow = new Date(parts[0], parts[1] - 1, parts[2]).getDay(); // 0=일 ... 6=토
  const base = flattenFillItems_().concat(CLEAN_ITEMS, FINAL_ITEMS);
  if (dow === 1) return base.concat(WEEKLY_ITEMS);
  return base;
}

const LOG_SHEET_NAME = 'Log';
const DAY_SHEET_NAME = 'DayStatus';
const MONTHLY_SHEET_NAME = 'MonthlyStatus';
const TIMEZONE = 'Asia/Seoul';

// ===== 웹 진입점 =====

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = params.callback;
  const action = params.action;

  let result;
  if (action === 'state') {
    result = getState(params.date);
  } else if (action === 'monthly') {
    result = getMonthlyState(params.month);
  } else {
    return HtmlService.createHtmlOutput(
      '<meta charset="utf-8"><body style="font-family:sans-serif;padding:40px;text-align:center;">' +
      '<h2>마감체크리스트</h2><p>이 링크가 아니라 새 웹앱 주소로 접속해주세요.</p></body>'
    );
  }

  const json = JSON.stringify(result);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let data = {};
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    data = {};
  }

  let result;
  if (data.action === 'toggle') {
    result = toggleItem(data.date, data.itemId, data.itemLabel, data.checked);
  } else if (data.action === 'complete') {
    result = completeDay(data.date);
  } else if (data.action === 'uncomplete') {
    result = uncompleteDay(data.date);
  } else if (data.action === 'toggleMonthly') {
    result = toggleMonthly(data.month, data.itemId, data.checked);
  } else {
    result = { ok: false, message: 'unknown action' };
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ===== 내부 유틸 =====

function getSs_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getLogSheet_() {
  const ss = getSs_();
  let sh = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME);
    sh.appendRow(['날짜', '항목ID', '항목명', '체크여부', '체크한사람', '체크시각']);
  }
  return sh;
}

function getDaySheet_() {
  const ss = getSs_();
  let sh = ss.getSheetByName(DAY_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(DAY_SHEET_NAME);
    sh.appendRow(['날짜', '완료여부', '완료시각', '참여자']);
  }
  return sh;
}

function getMonthlySheet_() {
  const ss = getSs_();
  let sh = ss.getSheetByName(MONTHLY_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(MONTHLY_SHEET_NAME);
    sh.appendRow(['년월', '항목ID', '항목명', '체크여부', '완료일', '완료시각']);
  }
  return sh;
}

function todayStr_() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
}

function monthStr_() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM');
}

function rowDate_(cellValue) {
  return cellValue instanceof Date
    ? Utilities.formatDate(cellValue, TIMEZONE, 'yyyy-MM-dd')
    : cellValue;
}

// 시트가 "YYYY-MM" / "HH:mm" 문자열을 날짜·시간 값으로 자동 인식해
// Date 객체로 바꿔버리는 경우가 있어 읽을 때 항상 문자열로 되돌린다.
function rowMonth_(cellValue) {
  return cellValue instanceof Date
    ? Utilities.formatDate(cellValue, TIMEZONE, 'yyyy-MM')
    : cellValue;
}

function rowTime_(cellValue) {
  return cellValue instanceof Date
    ? Utilities.formatDate(cellValue, TIMEZONE, 'HH:mm')
    : cellValue;
}

// ===== 일일 체크 데이터 =====

function getState(dateStr) {
  const date = dateStr || todayStr_();
  const sh = getLogSheet_();
  const values = sh.getDataRange().getValues();
  const checks = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (rowDate_(row[0]) === date) {
      checks[row[1]] = {
        checked: row[3] === true || row[3] === 'TRUE',
        at: rowTime_(row[5]),
      };
    }
  }

  const daySh = getDaySheet_();
  const dayValues = daySh.getDataRange().getValues();
  let completed = false;
  let completedAt = '';
  for (let i = 1; i < dayValues.length; i++) {
    const row = dayValues[i];
    if (rowDate_(row[0]) === date) {
      completed = row[1] === true || row[1] === 'TRUE';
      completedAt = rowTime_(row[2]);
    }
  }

  return {
    ok: true,
    date: date,
    checks: checks,
    completed: completed,
    completedAt: completedAt,
  };
}

function toggleItem(dateStr, itemId, itemLabel, checked) {
  const date = dateStr || todayStr_();
  const sh = getLogSheet_();
  const values = sh.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (rowDate_(values[i][0]) === date && values[i][1] === itemId) {
      targetRow = i + 1;
      break;
    }
  }
  const now = Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm');

  if (targetRow === -1) {
    sh.appendRow([date, itemId, itemLabel, checked, '', now]);
  } else {
    sh.getRange(targetRow, 4, 1, 3).setValues([[checked, '', now]]);
  }
  return getState(date);
}

function completeDay(dateStr) {
  const date = dateStr || todayStr_();
  const state = getState(date);
  const required = requiredItemsForDate_(date);
  const missing = required.filter(function (it) {
    return !state.checks[it.id] || !state.checks[it.id].checked;
  });
  if (missing.length > 0) {
    return {
      ok: false,
      message: '아직 체크 안 된 항목이 ' + missing.length + '개 있어요.',
      state: state,
    };
  }

  const now = Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm');
  const daySh = getDaySheet_();
  const dayValues = daySh.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < dayValues.length; i++) {
    if (rowDate_(dayValues[i][0]) === date) { targetRow = i + 1; break; }
  }
  if (targetRow === -1) {
    daySh.appendRow([date, true, now, '']);
  } else {
    daySh.getRange(targetRow, 2, 1, 2).setValues([[true, now]]);
  }
  return { ok: true, state: getState(date) };
}

function uncompleteDay(dateStr) {
  const date = dateStr || todayStr_();
  const daySh = getDaySheet_();
  const dayValues = daySh.getDataRange().getValues();
  for (let i = 1; i < dayValues.length; i++) {
    if (rowDate_(dayValues[i][0]) === date) {
      daySh.getRange(i + 1, 2).setValue(false);
      break;
    }
  }
  return getState(date);
}

// ===== 월간 청소 데이터 =====

function getMonthlyState(monthStr) {
  const m = monthStr || monthStr_();
  const sh = getMonthlySheet_();
  const values = sh.getDataRange().getValues();
  const checks = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (rowMonth_(row[0]) === m) {
      checks[row[1]] = {
        checked: row[3] === true || row[3] === 'TRUE',
        date: rowDate_(row[4]),
        at: rowTime_(row[5]),
      };
    }
  }
  return { ok: true, month: m, checks: checks };
}

function toggleMonthly(monthStr, itemId, checked) {
  const m = monthStr || monthStr_();
  const sh = getMonthlySheet_();
  const values = sh.getDataRange().getValues();
  let targetRow = -1;
  let itemLabel = '';
  MONTHLY_ITEMS.forEach(function (it) { if (it.id === itemId) itemLabel = it.label; });
  for (let i = 1; i < values.length; i++) {
    if (rowMonth_(values[i][0]) === m && values[i][1] === itemId) {
      targetRow = i + 1;
      break;
    }
  }
  const today = todayStr_();
  const now = Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm');
  if (targetRow === -1) {
    sh.appendRow([m, itemId, itemLabel, checked, today, now]);
  } else {
    sh.getRange(targetRow, 4, 1, 3).setValues([[checked, today, now]]);
  }
  return getMonthlyState(m);
}
