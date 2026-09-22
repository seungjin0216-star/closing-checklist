/**
 * 마감 체크리스트 - Google Apps Script 백엔드 (v5)
 *
 * ⚠️ 2026-09-08 — 원당본점이 붙었습니다. 배포는 하나만 씁니다.
 *
 *    화면이 branch 값을 보내면 그 지점 시트에 기록합니다.
 *        보내지 않으면(또는 'baekseok')  →  DayStatus      (백석. 기존 그대로)
 *        branch=wondang               →  DayStatus_원당
 *
 *    ⚠️ 백석 경로는 한 글자도 안 바꿨습니다. branch 가 없으면 v4 와 똑같이 돕니다.
 *       백석 마감은 매일 돌아가는 실전이라 건드리면 안 됩니다.
 *
 * (아래는 v4 설명)
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
const CARRY_SHEET_NAME = 'Carryover';   // 마감 → 오픈 이월 항목
const TIMEZONE = 'Asia/Seoul';

// ⚠️ 지점별 시트 이름
//    백석은 접미사가 없습니다. 예전에 쌓아둔 기록을 그대로 이어 쓰기 위해서입니다.
//    이름을 바꾸면 지난 기록이 통째로 안 보이게 됩니다.
const BRANCH_SUFFIX = {
  baekseok: '',          // DayStatus       (기존)
  wondang : '_원당',      // DayStatus_원당   (새로 생김)
};

function sheetName_(base, branch) {
  var suffix = BRANCH_SUFFIX[String(branch || 'baekseok')];
  if (suffix === undefined) suffix = '';   // 모르는 값이 오면 백석으로 (안전)
  return base + suffix;
}

// ===== 웹 진입점 =====

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = params.callback;
  const action = params.action;

  let result;
  if (action === 'monthly') {
    result = getMonthlyState(params.month, params.branch);
  } else if (action === 'carryover') {
    result = getCarryover(params.date, params.branch);
  } else if (action === 'checks') {
    result = getChecks(params.date, params.mode, params.branch);
  } else if (action === 'devices') {
    result = getDevices();
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
    result = recordCompletion(data.date, data.branch);
  } else if (data.action === 'completeOpen') {
    result = recordOpenCompletion(data.date, data.carried, data.branch);
  } else if (data.action === 'carryover') {
    result = saveCarryover(data.fromDate, data.targetDate, data.items, data.branch);
  } else if (data.action === 'toggleMonthly') {
    result = toggleMonthly(data.month, data.itemId, data.checked, data.branch);
  } else if (data.action === 'checks') {
    result = saveChecks(data);
  } else if (data.action === 'nameDevice') {
    result = nameDevice(data.device, data.name);
  } else if (data.action === 'cash') {
    result = saveCash(data);
  } else {
    result = { ok: false, message: 'unknown action' };
  }

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ===== 내부 유틸 =====

function getSs_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getDaySheet_(branch) {
  const ss = getSs_();
  const name = sheetName_(DAY_SHEET_NAME, branch);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(['날짜', '완료여부', '완료시각', '참여자']);
  }
  // 오픈/마감 구분 컬럼이 없으면 추가 (기존 시트 호환)
  if (sh.getLastColumn() < 5) {
    sh.getRange(1, 5).setValue('구분');
  }
  return sh;
}

function getCarrySheet_(branch) {
  const ss = getSs_();
  const name = sheetName_(CARRY_SHEET_NAME, branch);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(['마감일', '오픈일', '항목ID', '항목명', '그룹', '기록시각']);
  }
  return sh;
}

function getMonthlySheet_(branch) {
  const ss = getSs_();
  const name = sheetName_(MONTHLY_SHEET_NAME, branch);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
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

// ══════════════════════════════════════════════════════════
//  체크 공유 (2026-09-22 추가)
//
//  왜 넣었나 — 사장님 말:
//    「각각의 핸드폰에서 움직여지니 머가 문젠지 추적하기가 불편함」
//
//  그전까지 체크는 그 폰 안에만 있었습니다. 둘이 나눠 마감하면 서로 뭘 했는지
//  몰랐고, 사장님도 나중에 볼 방법이 없었습니다.
//
//  ⚠️ 예전에 한 번 공유를 넣었다가 뺀 적이 있습니다 (index.html v4 주석).
//     체크할 때마다 서버 답을 기다리게 만들어서 주춤거렸기 때문입니다.
//     이번에는 화면을 먼저 바꾸고 올리는 건 뒤에서 합니다 — 기다리지 않습니다.
//
//  ── 쌓기만 하고 고치지 않습니다 ────────────────────────
//     같은 항목을 껐다 켜도 줄을 고치지 않고 새 줄을 답니다.
//     ① 두 폰이 동시에 눌러도 서로 덮어쓸 일이 없습니다
//     ② 「누가 언제 뭘 눌렀나」가 통째로 남습니다 — 추적하려고 만든 것이니까요
//     읽을 때 항목마다 마지막 줄만 취하면 지금 상태가 됩니다.
// ══════════════════════════════════════════════════════════

const CHECK_SHEET_NAME  = 'Checks';
const DEVICE_SHEET_NAME = 'Devices';   // ⚠️ 지점 구분 없음 — 폰은 지점을 오갑니다

function getCheckSheet_(branch) {
  const ss   = getSs_();
  const name = sheetName_(CHECK_SHEET_NAME, branch);
  let sheet  = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(['시각', '날짜', '모드', '항목', '켬/끔', '폰', '찍힌시각']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e0e7ff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getDeviceSheet_() {
  const ss  = getSs_();
  let sheet = ss.getSheetByName(DEVICE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DEVICE_SHEET_NAME);
    sheet.appendRow(['폰', '이름', '기종', '첫 접속', '마지막 접속', '누른 횟수']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#fef3c7');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── 체크 저장 ────────────────────────────────────────────
//    앱이 1.5초쯤 모았다가 한 번에 보냅니다. 낱개로 오면 줄이 너무 많아집니다.
function saveChecks(data) {
  return withLock_(function () {
    const items  = data.items || [];
    if (!items.length) return { ok: true, saved: 0 };

    const date   = data.date || todayStr_();
    const mode   = data.mode || 'close';
    const device = String(data.device || '?');
    const now    = new Date();

    const rows = items.map(function (it) {
      return [
        now, date, mode, String(it.id),
        it.checked ? '켬' : '끔',
        device,
        it.at || Utilities.formatDate(now, TIMEZONE, 'HH:mm'),
      ];
    });

    const sheet = getCheckSheet_(data.branch);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);

    touchDevice_(device, data.ua, items.length);
    return { ok: true, saved: rows.length };
  });
}

// ── 폰 기록 ──────────────────────────────────────────────
//    ⚠️ 이름은 건드리지 않습니다. 사장님이 붙인 이름을 덮어쓰면 안 됩니다.
function touchDevice_(device, ua, count) {
  try {
    const sheet = getDeviceSheet_();
    const last  = sheet.getLastRow();
    const now   = new Date();

    if (last >= 2) {
      const keys = sheet.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) {
        if (String(keys[i][0]) === device) {
          const row = i + 2;
          sheet.getRange(row, 5).setValue(now);
          const prev = Number(sheet.getRange(row, 6).getValue()) || 0;
          sheet.getRange(row, 6).setValue(prev + (count || 1));
          return;
        }
      }
    }
    sheet.appendRow([device, '', String(ua || ''), now, now, count || 1]);
  } catch (err) {
    // 폰 기록이 실패해도 체크는 이미 저장됐습니다. 여기서 멈추면 안 됩니다.
    console.log('폰 기록 실패: ' + err.message);
  }
}

// ── 체크 읽기 ────────────────────────────────────────────
//    ⚠️ 뒤에서부터 훑어 항목마다 처음 만난 줄(= 가장 최근)만 취합니다.
function getChecks(dateStr, mode, branch) {
  const date  = dateStr || todayStr_();
  const want  = mode || 'close';
  const sheet = getCheckSheet_(branch);
  const last  = sheet.getLastRow();
  if (last < 2) return { ok: true, checks: {} };

  // 하루치만 보면 되므로 끝에서 800줄만 읽습니다 (전체를 읽으면 점점 느려집니다)
  const from  = Math.max(2, last - 800 + 1);
  const rows  = sheet.getRange(from, 1, last - from + 1, 7).getValues();

  const checks = {};
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    if (rowDate_(r[1]) !== date) continue;
    if (String(r[2]) !== want)   continue;
    const id = String(r[3]);
    if (checks[id]) continue;                  // 이미 더 최근 줄을 잡았습니다
    checks[id] = {
      checked: String(r[4]) === '켬',
      at     : rowTime_(r[6]) || '',
      device : String(r[5] || ''),
    };
  }
  return { ok: true, checks: checks, names: deviceNames_() };
}

// ── 폰 → 이름 표 ─────────────────────────────────────────
function deviceNames_() {
  const out   = {};
  try {
    const sheet = getDeviceSheet_();
    const last  = sheet.getLastRow();
    if (last < 2) return out;
    sheet.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      if (r[1]) out[String(r[0])] = String(r[1]);
    });
  } catch (err) {}
  return out;
}

// ── 폰 목록 (사장님 화면용) ──────────────────────────────
function getDevices() {
  const sheet = getDeviceSheet_();
  const last  = sheet.getLastRow();
  if (last < 2) return { ok: true, devices: [] };

  const rows = sheet.getRange(2, 1, last - 1, 6).getValues();
  const list = rows.map(function (r) {
    return {
      device: String(r[0]),
      name  : String(r[1] || ''),
      ua    : String(r[2] || ''),
      first : r[3] instanceof Date ? Utilities.formatDate(r[3], TIMEZONE, 'M/d HH:mm') : String(r[3] || ''),
      last  : r[4] instanceof Date ? Utilities.formatDate(r[4], TIMEZONE, 'M/d HH:mm') : String(r[4] || ''),
      count : Number(r[5]) || 0,
      hours : deviceHours_(String(r[0])),
    };
  });
  return { ok: true, devices: list };
}

// ── 그 폰이 주로 몇 시에 쓰나 ────────────────────────────
//    ⚠️ 이름을 붙일 때 제일 좋은 단서입니다. 오픈조는 아침, 마감조는 밤에 찍힙니다.
function deviceHours_(device) {
  try {
    const sheet = getCheckSheet_('baekseok');
    const last  = sheet.getLastRow();
    if (last < 2) return '';
    const from = Math.max(2, last - 1500 + 1);
    const rows = sheet.getRange(from, 1, last - from + 1, 6).getValues();

    const tally = {};
    rows.forEach(function (r) {
      if (String(r[5]) !== device) return;
      if (!(r[0] instanceof Date)) return;
      const h = Number(Utilities.formatDate(r[0], TIMEZONE, 'H'));
      tally[h] = (tally[h] || 0) + 1;
    });

    const hours = Object.keys(tally);
    if (!hours.length) return '';
    hours.sort(function (a, b) { return tally[b] - tally[a]; });
    return hours.slice(0, 2).map(function (h) { return h + '시'; }).join('·');
  } catch (err) {
    return '';
  }
}

// ── 폰에 이름 붙이기 (사장님만) ──────────────────────────
function nameDevice(device, name) {
  return withLock_(function () {
    const sheet = getDeviceSheet_();
    const last  = sheet.getLastRow();
    if (last < 2) return { ok: false, message: '폰 기록이 없습니다' };

    const keys = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === String(device)) {
        sheet.getRange(i + 2, 2).setValue(String(name || ''));
        return { ok: true };
      }
    }
    return { ok: false, message: '그런 폰이 없습니다: ' + device };
  });
}


// ══════════════════════════════════════════════════════════
//  💰 마감정산 · 시재  (2026-09-22)
//
//  사장님 말: 「시재 부족시 나한테 연락이든 알림이든 오게 해주면 내가 시재 챙기게」
//
//  ⚠️ 부족할 때만 메일이 갑니다. 맞으면 시트에 기록만 남습니다.
//     매일 「맞음」 메일이 오면 사흘 만에 안 읽게 됩니다.
//     그러면 정작 부족한 날의 메일도 안 읽게 됩니다.
//
//  ⚠️ 메일인 이유 — 솔라피 문자는 돈이 들고, 키가 막히면 알림 자체가 안 갑니다.
//     MailApp 은 구글 내장이라 키도 필요 없고 솔라피 상태와 무관합니다.
// ══════════════════════════════════════════════════════════

const CASH_SHEET_NAME = 'Cash';

function getCashSheet_(branch) {
  const ss   = getSs_();
  const name = sheetName_(CASH_SHEET_NAME, branch);
  let sheet  = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(['날짜', '시각', '결과', '부족액', '폰']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#fee2e2');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveCash(data) {
  return withLock_(function () {
    const date   = data.date || todayStr_();
    const 부족    = Number(data.short) || 0;
    const now    = new Date();
    const 시각    = Utilities.formatDate(now, TIMEZONE, 'HH:mm');
    const 지점    = (String(data.branch || 'baekseok') === 'wondang') ? '원당' : '백석';

    getCashSheet_(data.branch).appendRow([
      date, 시각, 부족 > 0 ? '부족' : '맞음', 부족, String(data.device || ''),
    ]);

    if (부족 > 0) {
      try {
        const 사람 = 폰이름_(String(data.device || ''));
        MailApp.sendEmail({
          to      : Session.getEffectiveUser().getEmail(),
          subject : '💰 [' + 지점 + '] 시재 부족 ' + 부족.toLocaleString() + '원 — ' + date,
          body    : [
            지점 + '점 마감 시재가 부족합니다.',
            '',
            '날짜    ' + date,
            '시각    ' + 시각,
            '부족액  ' + 부족.toLocaleString() + '원',
            '올린 폰  ' + (사람 || data.device || '알 수 없음'),
            '',
            '— 마감체크리스트',
          ].join('\n'),
        });
      } catch (err) {
        // ⚠️ 메일이 실패해도 기록은 이미 남았습니다. 여기서 멈추면 안 됩니다.
        console.log('시재 메일 실패: ' + err.message);
      }
    }
    return { ok: true, short: 부족 };
  });
}

// 폰 번호 → 사장님이 붙인 이름
function 폰이름_(device) {
  try {
    const t = deviceNames_();
    return t[device] || '';
  } catch (err) { return ''; }
}


// ===== 완료 기록 (개별) =====
// 체크 자체는 폰 안에서만 처리되고, "완료하기"를 누른 순간에만 한 번 호출된다.
// 여러 명이 각자 완료해도 그냥 각각 새 줄로 쌓인다 (덮어쓰기/충돌 없음).
function recordCompletion(dateStr, branch) {
  return withLock_(function () {
    const date = dateStr || todayStr_();
    const now = Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm');
    const daySh = getDaySheet_(branch);
    daySh.appendRow([date, true, now, '', '마감']);
    return { ok: true, date: date, completedAt: now };
  });
}

// ===== 오픈 완료 기록 =====
function recordOpenCompletion(dateStr, carriedLabels, branch) {
  return withLock_(function () {
    const date = dateStr || todayStr_();
    const now = Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm');
    const daySh = getDaySheet_(branch);
    daySh.appendRow([date, true, now, carriedLabels || '', '오픈']);
    return { ok: true, date: date, completedAt: now };
  });
}

// ===== 마감 → 오픈 이월 =====
// 마감 때 재료 손질 등을 못 하면 다음 오픈으로 넘긴다.
// 어떤 재료가 자주 밀리는지 쌓이면 발주·손질 시점을 조정하는 근거가 된다.
function saveCarryover(fromDate, targetDate, items, branch) {
  return withLock_(function () {
    if (!items || !items.length) return { ok: true, saved: 0 };
    const sh = getCarrySheet_(branch);
    const now = Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm');
    const rows = items.map(function (it) {
      return [fromDate || todayStr_(), targetDate || '', it.id || '', it.label || '', it.group || '', now];
    });
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    return { ok: true, saved: rows.length };
  });
}

function getCarryover(dateStr, branch) {
  const date = dateStr || todayStr_();
  const sh = getCarrySheet_(branch);
  const values = sh.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (rowDate_(row[1]) === date) {
      items.push({ id: row[2], label: row[3], group: row[4] });
    }
  }
  return { ok: true, date: date, items: items };
}

// ===== 월간 청소 데이터 (공유) =====

function getMonthlyState(monthStr, branch) {
  const m = monthStr || monthStr_();
  const sh = getMonthlySheet_(branch);
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

function toggleMonthly(monthStr, itemId, checked, branch) {
  return withLock_(function () {
    const m = monthStr || monthStr_();
    const sh = getMonthlySheet_(branch);
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
