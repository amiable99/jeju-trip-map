/**
 * 제주 가족여행 지도 - "가족이 추가한 곳" 저장 API
 *
 * 시트 `places` 열: id | createdAt | name | kind | link | memo | addedBy
 * 링크를 아는 가족 누구나 추가·삭제할 수 있다 (비밀번호 없음).
 * 배포: 웹 앱, 실행 계정 "나", 액세스 "모든 사용자".
 */

const SHEET_NAME = 'places';
const HEADERS = ['id', 'createdAt', 'name', 'kind', 'link', 'memo', 'addedBy'];
const KINDS = ['식당', '카페', '가볼 곳'];

function doGet() {
  try {
    return json_({ ok: true, places: readAll_() });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server' });
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: 'busy' });
  try {
    if (body.action === 'add') return add_(body);
    if (body.action === 'delete') return remove_(body);
    return json_({ ok: false, error: 'bad_action' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

function add_(body) {
  const name = str_(body.name);
  const kind = str_(body.kind);
  const link = str_(body.link);
  const memo = str_(body.memo);
  const addedBy = str_(body.addedBy);

  if (!name) return json_({ ok: false, error: 'name_required' });
  if (name.length > 60) return json_({ ok: false, error: 'name_too_long' });
  if (KINDS.indexOf(kind) < 0) return json_({ ok: false, error: 'bad_kind' });
  if (link && !/^https:\/\/\S+$/.test(link)) return json_({ ok: false, error: 'bad_link' });
  if (memo.length > 200) return json_({ ok: false, error: 'memo_too_long' });
  if (addedBy.length > 20) return json_({ ok: false, error: 'by_too_long' });

  const place = {
    id: Utilities.getUuid(),
    createdAt: Date.now(),
    name: name,
    kind: kind,
    link: link,
    memo: memo,
    addedBy: addedBy,
  };
  sheet_().appendRow(HEADERS.map(function (h) { return cell_(place[h]); }));
  return json_({ ok: true, place: place });
}

function remove_(body) {
  const id = str_(body.id);
  if (!id) return json_({ ok: false, error: 'not_found' });
  const sheet = sheet_();
  const ids = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  for (let i = ids.length - 1; i >= 1; i--) {
    if (String(ids[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return json_({ ok: true });
    }
  }
  return json_({ ok: false, error: 'not_found' });
}

function readAll_() {
  const sheet = sheet_();
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  return rows
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      const p = {};
      HEADERS.forEach(function (h, i) { p[h] = h === 'createdAt' ? Number(r[i]) || 0 : String(r[i]); });
      return p;
    })
    .sort(function (a, b) { return a.createdAt - b.createdAt; });
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function str_(v) {
  return v == null ? '' : String(v).trim();
}

// =, +, -, @ 로 시작하는 값은 시트가 수식으로 읽으므로 앞에 ' 를 붙여 글자로 저장한다.
function cell_(v) {
  return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
