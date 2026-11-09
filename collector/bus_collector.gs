/**
 * 버스 통과시각 수집기 (여러 정류장) — Google Apps Script
 *
 * 수집 정류장
 *  - 20675 덕이중학교 (하이파크시티2·3단지 방면)
 *  - 20086 대화역 (예비군훈련장·성저3·4단지 방면) — 대화역 → 덕이동 귀가 방향
 *
 * 사용법
 *  1) 기존 Code.gs 내용을 전부 지우고 이 코드 붙여넣기
 *  2) SERVICE_KEY에 data.go.kr 일반 인증키 입력
 *  3) 함수 선택에서 setup 실행 (정류소ID 자동 조회 + 시트 열 추가 + 트리거 재설치)
 *  4) 웹 앱을 이미 배포했다면: 배포 → 배포 관리 → 연필 → 버전 '새 버전' → 배포
 *
 * 사용하는 data.go.kr API
 *  - 경기도_버스도착정보 조회 (수집, 실시간 웹 앱)
 *  - 경기도_버스노선 조회 (정류소ID 찾기, setup 때 1회)
 */

const CONFIG = {
  SERVICE_KEY: '여기에_data.go.kr_인증키_입력',
  STATIONS: [
    { no: '20675', name: '덕이중학교', routes: ['074', '071', 'N002', '10'] },
    { no: '20086', name: '대화역',     routes: ['10', '090', 'N002'] },
  ],
  ACTIVE_START: 5 * 60,       // 05:00부터
  ACTIVE_END: 24 * 60 + 30,   // 다음날 00:30까지
  EVERY_N_MIN: 3,             // 정류장 2곳 × 3분 간격 = 하루 약 780회 (한도 1,000회)
};

const BASE = 'https://apis.data.go.kr/6410000';
const RAW = 'raw';       // 매 호출 원본 기록
const PASS = 'passes';   // 버스 통과 추정 시각 → 시간표의 재료

// ───────── 최초 1회 / 정류장 추가 시 실행 ─────────
function setup() {
  const props = PropertiesService.getScriptProperties();

  // 예전 버전(덕이중학교만 수집)에서 저장한 값 이어받기
  const legacy = props.getProperty('STATION_ID');
  if (legacy && !props.getProperty('SID_20675')) props.setProperty('SID_20675', legacy);
  const legacyState = props.getProperty('STATE');
  if (legacyState && !props.getProperty('STATE_20675')) props.setProperty('STATE_20675', legacyState);

  CONFIG.STATIONS.forEach(st => {
    if (!props.getProperty('SID_' + st.no)) {
      props.setProperty('SID_' + st.no, findStationIdByRoute_(st));
    } else {
      Logger.log(`${st.name}(${st.no}) stationId=${props.getProperty('SID_' + st.no)} (저장된 값 사용)`);
    }
  });

  const ss = SpreadsheetApp.getActive();
  sheet_(ss, RAW, ['시각', '요일', '노선', '차량1', '도착예정1(분)', '정류장전1',
                   '차량2', '도착예정2(분)', '정류장전2', '정류장']);
  sheet_(ss, PASS, ['날짜', '요일', '노선', '차량번호', '통과추정시각', '마지막관측시각', '마지막예정(분)', '정류장']);
  // 예전 시트에 '정류장' 열 제목 추가 (기존 행은 비어 있음 = 덕이중학교)
  [[RAW, 10], [PASS, 8]].forEach(([name, col]) => {
    const sh = ss.getSheetByName(name);
    if (!sh.getRange(1, col).getValue()) sh.getRange(1, col).setValue('정류장');
  });

  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('poll').timeBased().everyMinutes(1).create();
  Logger.log('설치 완료. 정류장 ' + CONFIG.STATIONS.length + '곳을 ' + CONFIG.EVERY_N_MIN + '분 간격으로 수집합니다.');
  poll(true);
}

// ───────── 트리거가 1분마다 호출 ─────────
function poll(force) {
  const now = new Date();
  const mins = now.getHours() * 60 + now.getMinutes();
  const m = mins < 3 * 60 ? mins + 24 * 60 : mins;            // 자정 넘김 처리
  if (force !== true) {
    if (m < CONFIG.ACTIVE_START || m > CONFIG.ACTIVE_END) return;
    if (mins % CONFIG.EVERY_N_MIN !== 0) return;
  }
  CONFIG.STATIONS.forEach(st => {
    try { pollStation_(st, now); } catch (e) { Logger.log(`${st.name} 오류: ${e}`); }
  });
}

function pollStation_(st, now) {
  const props = PropertiesService.getScriptProperties();
  const stationId = props.getProperty('SID_' + st.no);
  if (!stationId) { Logger.log(`${st.name}: stationId 없음. setup을 실행하세요.`); return; }

  const list = fetchArrivals_(stationId);
  if (list === null) return;

  const ss = SpreadsheetApp.getActive();
  const day = '일월화수목금토'[now.getDay()];
  const ts = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

  // 1) 원본 기록 (노선명은 문자열로 저장해 '074'의 앞자리 0 유지)
  const rows = list.map(b => [ts, day, "'" + String(b.routeName),
    b.plateNo1 || '', num_(b.predictTime1), num_(b.locationNo1),
    b.plateNo2 || '', num_(b.predictTime2), num_(b.locationNo2), st.no]);
  if (rows.length) {
    const sh = ss.getSheetByName(RAW);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  // 2) 통과 감지: 직전에 '첫 번째 버스'였던 차량이 사라지면 → 정류장을 지나간 것
  const key = 'STATE_' + st.no;
  const state = JSON.parse(props.getProperty(key) || '{}');
  const seenRoutes = {};
  const passRows = [];

  list.forEach(b => {
    const route = String(b.routeName);
    seenRoutes[route] = true;
    const prev = state[route];
    const cur = { plate: b.plateNo1 || '', predict: num_(b.predictTime1), t: now.getTime() };
    if (prev && prev.plate && prev.plate !== cur.plate && prev.plate !== (b.plateNo2 || '')) {
      passRows.push(passRow_(route, prev, now, st.no));
    }
    state[route] = cur;
  });

  // 목록에서 노선 자체가 사라진 경우(도착정보 없음)도 통과로 처리
  Object.keys(state).forEach(route => {
    if (!seenRoutes[route] && state[route].plate) {
      passRows.push(passRow_(route, state[route], now, st.no));
      state[route] = { plate: '', predict: '', t: now.getTime() };
    }
  });

  if (passRows.length) {
    const sh = ss.getSheetByName(PASS);
    sh.getRange(sh.getLastRow() + 1, 1, passRows.length, passRows[0].length).setValues(passRows);
  }
  props.setProperty(key, JSON.stringify(state));
}

// ───────── 실시간 도착정보 웹 앱 ─────────
// 배포 → 새 배포 → 유형: 웹 앱 → 실행: 나 / 액세스: 모든 사용자
// 주소 뒤에 ?stop=20086 을 붙이면 대화역, 없으면 덕이중학교.
// 정류장별 30초 캐시로 API 호출량을 제한합니다.
function doGet(e) {
  const no = (e && e.parameter && e.parameter.stop) || CONFIG.STATIONS[0].no;
  const st = CONFIG.STATIONS.find(s => s.no === no);
  if (!st) return json_({ ok: false, error: 'unknown stop' });

  const cache = CacheService.getScriptCache();
  let json = cache.get('live_' + no);
  if (!json) {
    const stationId = PropertiesService.getScriptProperties().getProperty('SID_' + no);
    const list = stationId ? fetchArrivals_(stationId) : null;
    json = JSON.stringify({
      stop: no,
      name: st.name,
      updated: Utilities.formatDate(new Date(), 'Asia/Seoul', "yyyy-MM-dd'T'HH:mm:ssXXX"),
      ok: list !== null,
      buses: (list || []).map(b => ({
        route: String(b.routeName),
        predict1: num_(b.predictTime1), stops1: num_(b.locationNo1),
        predict2: num_(b.predictTime2), stops2: num_(b.locationNo2),
      })),
    });
    if (list !== null) cache.put('live_' + no, json, 30);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ───────── 정류소ID 찾기: '버스노선 조회' API 사용 ─────────
// 이 정류장을 지나는 노선의 정류소 목록에서 표시번호가 일치하는 곳을 찾음
function findStationIdByRoute_(st) {
  const key = encodeURIComponent(CONFIG.SERVICE_KEY);
  for (const routeNo of st.routes) {
    const listUrl = `${BASE}/busrouteservice/v2/getBusRouteListv2` +
      `?serviceKey=${key}&keyword=${encodeURIComponent(routeNo)}&format=json`;
    const lb = fetchJson_(listUrl);
    const routes = asArray_(lb && lb.response && lb.response.msgBody &&
                            lb.response.msgBody.busRouteList)
      .filter(r => String(r.routeName).trim() === routeNo);
    Logger.log(`${st.name} / 노선 ${routeNo}: 후보 ${routes.length}개`);

    for (const r of routes) {
      const stUrl = `${BASE}/busrouteservice/v2/getBusRouteStationListv2` +
        `?serviceKey=${key}&routeId=${r.routeId}&format=json`;
      const sb = fetchJson_(stUrl);
      const stations = asArray_(sb && sb.response && sb.response.msgBody &&
                                sb.response.msgBody.busRouteStationList);
      const hit = stations.find(s => String(s.mobileNo).trim() === st.no);
      if (hit) {
        Logger.log(`정류소: ${hit.stationName} / stationId=${hit.stationId} (노선 ${routeNo}, ${r.regionName || ''})`);
        return String(hit.stationId);
      }
    }
  }
  throw new Error(`${st.name}(${st.no}) 정류소를 찾지 못했습니다. 위 로그를 캡처해서 보내주세요.`);
}

// ───────── 보조 함수 ─────────
function fetchArrivals_(stationId) {
  const url = `${BASE}/busarrivalservice/v2/getBusArrivalListv2` +
    `?serviceKey=${encodeURIComponent(CONFIG.SERVICE_KEY)}` +
    `&stationId=${stationId}&format=json`;
  const body = fetchJson_(url);
  if (!body) return null;
  return asArray_(body.response && body.response.msgBody && body.response.msgBody.busArrivalList);
}

function passRow_(route, prev, now, stopNo) {
  // 통과시각 ≈ 마지막 관측시각 + 그때의 도착예정(분), 단 현재 시각을 넘지 않게
  const est = Math.min(prev.t + (Number(prev.predict) || 0) * 60000, now.getTime());
  const d = new Date(est);
  return [
    Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd'),
    '일월화수목금토'[d.getDay()],
    "'" + route,
    prev.plate,
    Utilities.formatDate(d, 'Asia/Seoul', 'HH:mm'),
    Utilities.formatDate(new Date(prev.t), 'Asia/Seoul', 'HH:mm'),
    prev.predict,
    stopNo,
  ];
}

function fetchJson_(url) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code !== 200) { Logger.log('HTTP ' + code + ': ' + text.slice(0, 300)); return null; }
  try { return JSON.parse(text); }
  catch (e) { Logger.log('JSON 아님(인증키 확인): ' + text.slice(0, 300)); return null; }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function asArray_(x) { return !x ? [] : Array.isArray(x) ? x : [x]; }
function num_(x) { return (x === undefined || x === null || x === '') ? '' : Number(x); }

function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
