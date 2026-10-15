// 자연속 캠핑 배치도 — 예약 저장 서버 (Google Apps Script)
// 예약 데이터는 이 스크립트가 연결된 스프레드시트의 '예약' 시트에 저장됩니다.

// ▼ 직원 비밀번호: 꼭 원하는 값으로 바꾼 뒤 배포하세요.
const STAFF_PIN = '1234';

const SHEET_NAME = '예약';
// 열 순서: 날짜 | 사이트 | 메모 | 수정시각 | 키
const COL = { date: 0, label: 1, memo: 2, updated: 3, key: 4 };

// GitHub Pages 화면이 이 주소로 요청을 보냅니다.
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fns = { getDay: getDay, checkPin: checkPin, setReservation: setReservation };
    if (!fns[req.fn]) throw new Error('알 수 없는 요청입니다.');
    out = { ok: true, result: fns[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// 브라우저로 주소를 열었을 때 작동 여부만 보여줍니다.
function doGet() {
  return ContentService.createTextOutput('자연속 캠핑 예약 서버가 작동 중입니다.');
}

function siteLabel_(key) {
  const m = /^(s|b)(\d+)$/.exec(String(key));
  if (!m) return null;
  const n = Number(m[2]);
  if (m[1] === 's' && n >= 1 && n <= 58) return n + '번';
  if (m[1] === 'b' && n >= 1 && n <= 6) return '방갈로 ' + n;
  return null;
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['날짜', '사이트', '메모', '수정시각', '키']);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('E:E').setNumberFormat('@');
  }
  return sh;
}

function isStaff_(pin) {
  return String(pin || '') === String(STAFF_PIN);
}

function validDate_(d) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(d));
}

// 직원 비밀번호 확인 (틀리면 잠깐 기다리게 해서 마구 대입하기 어렵게 함)
function checkPin(pin) {
  const ok = isStaff_(pin);
  if (!ok) Utilities.sleep(800);
  return ok;
}

// 하루치 예약 목록. 메모는 직원에게만 보냄.
function getDay(date, pin) {
  if (!validDate_(date)) throw new Error('날짜 형식이 올바르지 않습니다.');
  const staff = isStaff_(pin);
  const rows = sheet_().getDataRange().getDisplayValues().slice(1);
  const sites = rows
    .filter(r => r[COL.date] === date && siteLabel_(r[COL.key]))
    .map(r => staff ? { site: r[COL.key], memo: r[COL.memo] } : { site: r[COL.key] });
  return { sites: sites, staff: staff };
}

// 예약 체크/해제 (직원만)
function setReservation(date, key, on, memo, pin) {
  if (!isStaff_(pin)) throw new Error('직원 비밀번호가 맞지 않습니다. 다시 로그인해 주세요.');
  const label = siteLabel_(key);
  if (!validDate_(date) || !label) throw new Error('잘못된 요청입니다.');
  memo = String(memo || '').slice(0, 80);
  if (/^[=+\-@]/.test(memo)) memo = "'" + memo; // 수식으로 해석되지 않게

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const vals = sh.getDataRange().getDisplayValues();
    let row = -1;
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][COL.date] === date && vals[i][COL.key] === key) { row = i + 1; break; }
    }
    if (on) {
      const data = [[date, label, memo, new Date(), key]];
      const target = row > 0 ? row : sh.getLastRow() + 1;
      sh.getRange(target, 1, 1, 5).setValues(data);
    } else if (row > 0) {
      sh.deleteRow(row);
    }
  } finally {
    lock.releaseLock();
  }
  return true;
}
