/**
 * 제주 가족여행 지도 - "가족이 추가한 곳" 저장 API
 *
 * 시트 `places` 열: id | createdAt | name | kind | link | memo | addedBy | lat | lon
 * lat/lon 은 쓰지 않는다 (예전 버전이 쓰던 열). 좌표는 사이트가 sportsmagicnum.com 에서 찾는다.
 *   Google 서버에서는 네이버가 접속을 막아서, 여기서 찾으려 하면 시간 초과까지 기다리느라 느려졌다.
 * 목록은 CacheService 에 담아 두고, 추가·삭제할 때 지운다.
 * 링크를 아는 가족 누구나 추가·삭제할 수 있다 (비밀번호 없음).
 * 종류(kind)는 사이트(index.html)에서 정한다. 여기서는 길이만 본다.
 * 배포: 웹 앱, 실행 계정 "나", 액세스 "모든 사용자".
 */

const SHEET_NAME = 'places';
const HEADERS = ['id', 'createdAt', 'name', 'kind', 'link', 'memo', 'addedBy', 'lat', 'lon'];

const CACHE_KEY = 'places';

function doGet() {
  try {
    const cache = CacheService.getScriptCache();
    const hit = cache.get(CACHE_KEY);
    if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
    const text = JSON.stringify({ ok: true, places: readAll_() });
    if (text.length < 90000) cache.put(CACHE_KEY, text, 21600);
    return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
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
    CacheService.getScriptCache().remove(CACHE_KEY);
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
  if (!kind || kind.length > 10) return json_({ ok: false, error: 'bad_kind' });
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
    lat: '',
    lon: '',
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
      p.lat = Number(p.lat) || null;
      p.lon = Number(p.lon) || null;
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
  } else if (sheet.getLastColumn() < HEADERS.length) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
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
