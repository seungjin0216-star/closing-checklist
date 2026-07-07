/**
 * 백석직영점 마감 체크리스트 - Google Apps Script 백엔드
 * ------------------------------------------------------
 * 이 파일은 Google 스프레드시트에 바인딩된 Apps Script 프로젝트의
 * Code.gs 파일입니다. (배포 방법은 README.md 참고)
 *
 * 항목을 추가/수정/삭제하고 싶으면 아래 BASE_ITEMS / EXTRA_ITEMS /
 * INGREDIENT_ITEMS 배열만 고치면 전체 화면에 자동 반영됩니다.
 */

// ===== 체크리스트 항목 정의 =====

// 필수 마감 (모두 확인)
const BASE_ITEMS = [
  { id: 'b1', label: '들기름 등 소스류 리필' },
  { id: 'b2', label: '찬냉장고 반찬리필' },
  { id: 'b3', label: '불쇼술 채우기' },
  { id: 'b4', label: '쓴 행주 전부 삶기 (테이블행주 포함)' },
  { id: 'b5', label: '술, 음료 채우기' },
  { id: 'b6', label: '술박스 정리하기' },
  { id: 'b7', label: '육수 확인' },
  { id: 'b8', label: '에어컨 끄기 (홀3, 주방1)' },
  { id: 'b9', label: '청포도, 이쑤시개 채우기' },
  { id: 'b10', label: '화구 라인 청소' },
];

// 마감 후 추가로 확인 (직원 위주지만 누구나 체크 가능)
const EXTRA_ITEMS = [
  { id: 'e1', label: '감미 채우기' },
  { id: 'e2', label: '식기세척기 끄고 청소' },
  { id: 'e3', label: '고기정리 (부족한 거 바로 채우기)' },
  { id: 'e4', label: '고기·식자재 발주 무조건 하기' },
  { id: 'e5', label: '가스 끄기' },
  { id: 'e6', label: '마감 시재 맞추기' },
  { id: 'e7', label: '저녁 먹을 거 주문하기' },
];

// 재료준비 (추가 확인 섹션 안의 별도 체크 그룹)
const INGREDIENT_ITEMS = [
  { id: 'i1', label: '양파' },
  { id: 'i2', label: '떡' },
  { id: 'i3', label: '전골재료 - 대파' },
  { id: 'i4', label: '전골재료 - 배추' },
  { id: 'i5', label: '전골재료 - 내장' },
  { id: 'i6', label: '전골재료 - 고추' },
  { id: 'i7', label: '라면대파' },
  { id: 'i8', label: '다대기' },
  { id: 'i9', label: '다진마늘' },
  { id: 'i10', label: '파장소스재료 - 다진대파' },
  { id: 'i11', label: '파장소스재료 - 다진고추' },
  { id: 'i12', label: '콩나물' },
  { id: 'i13', label: '레몬' },
  { id: 'i14', label: '비빔국수재료 - 묵은지' },
  { id: 'i15', label: '비빔국수재료 - 오이' },
  { id: 'i16', label: '간' },
  { id: 'i17', label: '천엽' },
];

const ALL_ITEMS = [].concat(BASE_ITEMS, EXTRA_ITEMS, INGREDIENT_ITEMS);

const LOG_SHEET_NAME = 'Log';
const DAY_SHEET_NAME = 'DayStatus';
const TIMEZONE = 'Asia/Seoul';

// ===== 웹앱 진입점 =====

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('백석직영점 마감 체크리스트')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

function todayStr_() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
}

function rowDate_(cellValue) {
  return cellValue instanceof Date
    ? Utilities.formatDate(cellValue, TIMEZONE, 'yyyy-MM-dd')
    : cellValue;
}

// ===== 클라이언트(Index.html)에서 google.script.run 으로 호출하는 함수들 =====

function getItemDefinitions() {
  return { base: BASE_ITEMS, extra: EXTRA_ITEMS, ingredient: INGREDIENT_ITEMS };
}

function getSheetUrl() {
  return getSs_().getUrl();
}

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
        by: row[4],
        at: row[5],
      };
    }
  }

  const daySh = getDaySheet_();
  const dayValues = daySh.getDataRange().getValues();
  let completed = false;
  let completedAt = '';
  let participants = '';
  for (let i = 1; i < dayValues.length; i++) {
    const row = dayValues[i];
    if (rowDate_(row[0]) === date) {
      completed = row[1] === true || row[1] === 'TRUE';
      completedAt = row[2];
      participants = row[3];
    }
  }

  return {
    date: date,
    checks: checks,
    completed: completed,
    completedAt: completedAt,
    participants: participants,
  };
}

function toggleItem(dateStr, itemId, itemLabel, checked, byName) {
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
  const name = (byName || '').trim() || '이름미입력';

  if (targetRow === -1) {
    sh.appendRow([date, itemId, itemLabel, checked, name, now]);
  } else {
    sh.getRange(targetRow, 4, 1, 3).setValues([[checked, name, now]]);
  }
  return getState(date);
}

function completeDay(dateStr, byName) {
  const date = dateStr || todayStr_();
  const state = getState(date);
  const missing = ALL_ITEMS.filter(function (it) {
    return !state.checks[it.id] || !state.checks[it.id].checked;
  });
  if (missing.length > 0) {
    return {
      ok: false,
      message: '아직 체크 안 된 항목이 ' + missing.length + '개 있어요.',
      state: state,
    };
  }

  const namesSet = {};
  ALL_ITEMS.forEach(function (it) {
    const c = state.checks[it.id];
    if (c && c.by) namesSet[c.by] = true;
  });
  if (byName && byName.trim()) namesSet[byName.trim()] = true;
  const participants = Object.keys(namesSet).join(', ');
  const now = Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm');

  const daySh = getDaySheet_();
  const dayValues = daySh.getDataRange().getValues();
  let targetRow = -1;
  for (let i = 1; i < dayValues.length; i++) {
    if (rowDate_(dayValues[i][0]) === date) {
      targetRow = i + 1;
      break;
    }
  }
  if (targetRow === -1) {
    daySh.appendRow([date, true, now, participants]);
  } else {
    daySh.getRange(targetRow, 2, 1, 3).setValues([[true, now, participants]]);
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
