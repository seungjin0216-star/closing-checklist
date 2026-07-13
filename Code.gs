/**
 * 백석직영점 마감 체크리스트 - Google Apps Script 백엔드 (v4)
 * ------------------------------------------------------
 * v4부터는 "각자 폰 안에서만 체크"하는 방식으로 바뀌었습니다.
 * 체크할 때마다 서버에 저장하지 않고, "완료하기"를 누른 순간에만
 * 딱 한 번 구글시트에 기록합니다 (여러 명이 각자 완료해도 각각 기록됨).
 * 그래서 체크/해제 자체는 이 스크립트를 거치지 않고, 완료 기록과
 * 월간청소만 이 스크립트를 거칩니다.
 *
 * - GET  (JSONP)
 *     ?action=monthly&month=YYYY-MM&callback=xxx     → 이번달 청소 상태 (공유)
 * - POST (JSON body)
 *     {action:'complete', date}                       → 완료 기록 (개별)
 *     {action:'toggleMonthly', month, itemId, checked} → 월간청소 체크 (공유)
 */

// ===== 월간 청소 (한 달에 한 번, 여러 명이 볼 수 있게 공유되는 유일한 항목) =====
const MONTHLY_ITEMS = [
  { id: 'm1', label: '냉장고 성에제거 및 청소' },
  { id: 'm2', label: '주방바닥 퐁퐁청소' },
  { id: 'm3', label: '유리 청소' },
];

const DAY_SHEET_NAME = 'DayStatus';
const MONTHLY_SHEET_NAME = 'MonthlyStatus';
const TIMEZONE = 'Asia/Seoul';

// ===== 웹 진입점 =====

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = params.callback;
  const action = params.action;

  let result;
  if (action === 'monthly') {
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
  if (data.action === 'complete') {
    result = recordCompletion(data.date);
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

// 구글시트에 동시에 여러 요청이 겹쳐서 읽고 쓰다가 서로 덮어쓰는 문제를 막기 위해,
// 시트를 실제로 고치는 모든 함수는 이 잠금(Lock)을 먼저 잡고 시작한다.
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ===== 완료 기록 (개별) =====
// 체크 자체는 폰 안에서만 처리되고, "완료하기"를 누른 순간에만 한 번 호출된다.
// 여러 명이 각자 완료해도 그냥 각각 새 줄로 쌓인다 (덮어쓰기/충돌 없음).
function recordCompletion(dateStr) {
  return withLock_(function () {
    const date = dateStr || todayStr_();
    const now = Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm');
    const daySh = getDaySheet_();
    daySh.appendRow([date, true, now, '']);
    return { ok: true, date: date, completedAt: now };
  });
}

// ===== 월간 청소 데이터 (공유) =====

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
  return withLock_(function () {
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
  });
}
