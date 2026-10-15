/**
 * 제주 가족여행 지도 - "가족이 추가한 곳" 저장 API
 *
 * 시트 `places` 열: id | createdAt | name | kind | link | memo | addedBy | lat | lon
 * lat/lon 은 네이버 지도 링크에서 찾는다. 못 찾으면 lat 에 '-' 를 적어 다시 찾지 않는다.
 * 링크를 아는 가족 누구나 추가·삭제할 수 있다 (비밀번호 없음).
 * 배포: 웹 앱, 실행 계정 "나", 액세스 "모든 사용자".
 */

const SHEET_NAME = 'places';
const HEADERS = ['id', 'createdAt', 'name', 'kind', 'link', 'memo', 'addedBy', 'lat', 'lon'];
const KINDS = ['식당', '카페', '가볼 곳'];

function doGet() {
  try {
    fillMissingCoords_();
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
    lat: '-',
    lon: '',
  };
  const c = link ? coordFromLink_(link) : null;
  if (c) { place.lat = c.lat; place.lon = c.lon; }
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

// 좌표가 아직 없는 줄(lat 칸이 빈 줄)은 목록을 불러올 때 링크에서 찾아 채운다. 한 번에 5줄까지.
function fillMissingCoords_() {
  const sheet = sheet_();
  if (sheet.getLastRow() < 2) return;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  const linkCol = HEADERS.indexOf('link'), latCol = HEADERS.indexOf('lat');
  const todo = [];
  rows.forEach(function (r, i) { if (r[0] !== '' && r[latCol] === '' && todo.length < 5) todo.push(i); });
  if (!todo.length) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    todo.forEach(function (i) {
      const link = String(rows[i][linkCol]);
      const c = link ? coordFromLink_(link) : null;
      sheet.getRange(i + 2, latCol + 1, 1, 2).setValues([c ? [c.lat, c.lon] : ['-', '']]);
    });
  } finally {
    lock.releaseLock();
  }
}

// 네이버 지도 링크에서 좌표를 찾는다.
// naver.me 짧은 링크는 map.naver.com/?lat=..&lng=.. 으로 넘어가므로 그 주소에서 읽고,
// 장소 번호만 있는 링크(map.naver.com/p/entry/place/123 등)는 네이버 장소 요약 API를 시도한다.
// log 를 넘기면 단계마다 무엇을 받았는지 적는다 (checkCoords 에서 씀).
function coordFromLink_(link, log) {
  log = log || function () {};
  try {
    let url = link;
    for (let hop = 0; hop < 3; hop++) {
      const c = coordFromParams_(url);
      if (c) return c;
      if (!/^https:\/\/naver\.me\//.test(url)) break;
      const res = UrlFetchApp.fetch(url, { followRedirects: false, muteHttpExceptions: true });
      const headers = res.getHeaders();
      const next = headers.Location || headers.location;
      log('fetch ' + url + ' -> ' + res.getResponseCode() + ' Location=' + next);
      if (!next) {
        const body = res.getContentText();
        log('body: ' + body.slice(0, 300));
        const c2 = coordFromParams_(body) || coordFromText_(body);
        if (c2) return c2;
        break;
      }
      url = next;
    }
    const id = (url.match(/place\/(\d+)/) || url.match(/place\.naver\.com\/[a-z]+\/(\d+)/) || url.match(/pinId=(\d+)/) || [])[1];
    log('place id: ' + id);
    if (!id) return null;
    const res = UrlFetchApp.fetch('https://map.naver.com/p/api/place/summary/' + id, {
      muteHttpExceptions: true,
      headers: { Referer: 'https://map.naver.com/' },
    });
    log('summary -> ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
    if (res.getResponseCode() !== 200) return null;
    const co = JSON.parse(res.getContentText()).data.placeDetail.coordinate;
    return inJeju_(Number(co.latitude), Number(co.longitude));
  } catch (err) {
    log('error: ' + err);
    console.warn('coordFromLink_', link, err);
    return null;
  }
}

// 편집기에서 직접 실행하는 점검용 함수: 위치를 못 찾은 줄을 다시 찾고, 단계별 결과를 실행 로그에 남긴다.
function checkCoords() {
  const sheet = sheet_();
  if (sheet.getLastRow() < 2) { console.log('시트에 장소가 없어요.'); return; }
  const latCol = HEADERS.indexOf('lat') + 1;
  const range = sheet.getRange(2, latCol, sheet.getLastRow() - 1, 1);
  range.setValues(range.getValues().map(function (r) { return [r[0] === '-' ? '' : r[0]]; }));
  console.log('시험: ' + JSON.stringify(coordFromLink_('https://naver.me/GDQRfCbl', function (m) { console.log(m); })));
  fillMissingCoords_();
  readAll_().forEach(function (p) { console.log(p.name + ': ' + p.lat + ', ' + p.lon); });
}

// 글 안에 "y":"33.3","x":"126.8" 처럼 좌표가 들어 있으면 읽는다.
function coordFromText_(text) {
  const m = text.match(/"y"\s*:\s*"?(3[3-4]\.\d+)"?\s*,\s*"x"\s*:\s*"?(12[6-7]\.\d+)/) ||
    text.match(/"x"\s*:\s*"?(12[6-7]\.\d+)"?\s*,\s*"y"\s*:\s*"?(3[3-4]\.\d+)/);
  if (!m) return null;
  return m[1].indexOf('12') === 0 ? inJeju_(Number(m[2]), Number(m[1])) : inJeju_(Number(m[1]), Number(m[2]));
}

function coordFromParams_(url) {
  const lat = (url.match(/[?&]lat=([\d.]+)/) || [])[1];
  const lon = (url.match(/[?&]lng=([\d.]+)/) || [])[1];
  return lat && lon ? inJeju_(Number(lat), Number(lon)) : null;
}

function inJeju_(lat, lon) {
  return lat > 33 && lat < 34 && lon > 126 && lon < 127.1 ? { lat: lat, lon: lon } : null;
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
