/**
 * 덕이중학교(20675) 버스 통과시각 수집기 — Google Apps Script
 *
 * 사용법
 *  1) 새 Google 시트 → 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
 *  2) SERVICE_KEY에 data.go.kr 일반 인증키 입력
 *  3) 함수 선택에서 setup 실행 → 권한 승인 (정류소ID 자동 조회 + 시트 생성 + 트리거 설치)
 *  4) 끝. 이후 구글 서버가 자동으로 수집합니다.
 *
 * 사용하는 data.go.kr API
 *  - 경기도_버스도착정보 조회 (수집)
 *  - 경기도_버스노선 조회 (정류소ID 찾기, setup 때 1회)
 */

const CONFIG = {
  SERVICE_KEY: '여기에_data.go.kr_인증키_입력',
  STATION_NO: '20675',        // 정류장 표시번호(모바일 번호)
  ROUTES: ['074', '071', 'N002', '10'],  // 정류소ID를 찾을 때 확인할 노선
  STATION_ID: '',             // 비워두면 setup()이 자동으로 찾음
  ACTIVE_START: 5 * 60,       // 05:00부터
  ACTIVE_END: 24 * 60 + 30,   // 다음날 00:30까지
  EVERY_N_MIN: 2,             // 개발계정 하루 1,000회 한도 → 2분 간격(하루 약 590회)
};

const BASE = 'https://apis.data.go.kr/6410000';
const RAW = 'raw';       // 매 호출 원본 기록
const PASS = 'passes';   // 버스 통과(출발) 추정 시각 → 시간표의 재료

// ───────── 최초 1회 실행 ─────────
function setup() {
  const props = PropertiesService.getScriptProperties();
  let stationId = CONFIG.STATION_ID;

  if (!stationId) stationId = findStationIdByRoute_();
  props.setProperty('STATION_ID', stationId);

  const ss = SpreadsheetApp.getActive();
  sheet_(ss, RAW, ['시각', '요일', '노선', '차량1', '도착예정1(분)', '정류장전1',
                   '차량2', '도착예정2(분)', '정류장전2']);
  sheet_(ss, PASS, ['날짜', '요일', '노선', '차량번호', '통과추정시각', '마지막관측시각', '마지막예정(분)']);

  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('poll').timeBased().everyMinutes(1).create();
  Logger.log('설치 완료. 1분마다 깨어나서 ' + CONFIG.EVERY_N_MIN + '분 간격으로 수집합니다.');
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

  const props = PropertiesService.getScriptProperties();
  const stationId = props.getProperty('STATION_ID');
  const url = `${BASE}/busarrivalservice/v2/getBusArrivalListv2` +
    `?serviceKey=${encodeURIComponent(CONFIG.SERVICE_KEY)}` +
    `&stationId=${stationId}&format=json`;
  const body = fetchJson_(url);
  if (!body) return;
  const list = asArray_(body.response && body.response.msgBody &&
                        body.response.msgBody.busArrivalList);

  const ss = SpreadsheetApp.getActive();
  const day = '일월화수목금토'[now.getDay()];
  const ts = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

  // 1) 원본 기록
  const rows = list.map(b => [ts, day, String(b.routeName),
    b.plateNo1 || '', num_(b.predictTime1), num_(b.locationNo1),
    b.plateNo2 || '', num_(b.predictTime2), num_(b.locationNo2)]);
  if (rows.length) {
    const sh = ss.getSheetByName(RAW);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  // 2) 통과 감지: 직전에 '첫 번째 버스'였던 차량이 사라지면 → 정류장을 지나간 것
  const state = JSON.parse(props.getProperty('STATE') || '{}');
  const seenRoutes = {};
  const passRows = [];

  list.forEach(b => {
    const route = String(b.routeName);
    seenRoutes[route] = true;
    const prev = state[route];
    const cur = { plate: b.plateNo1 || '', predict: num_(b.predictTime1), t: now.getTime() };
    if (prev && prev.plate && prev.plate !== cur.plate && prev.plate !== (b.plateNo2 || '')) {
      passRows.push(passRow_(route, prev, now));
    }
    state[route] = cur;
  });

  // 목록에서 노선 자체가 사라진 경우(도착정보 없음)도 통과로 처리
  Object.keys(state).forEach(route => {
    if (!seenRoutes[route] && state[route].plate) {
      passRows.push(passRow_(route, state[route], now));
      state[route] = { plate: '', predict: '', t: now.getTime() };
    }
  });

  if (passRows.length) {
    const sh = ss.getSheetByName(PASS);
    sh.getRange(sh.getLastRow() + 1, 1, passRows.length, passRows[0].length).setValues(passRows);
  }
  props.setProperty('STATE', JSON.stringify(state));
}

// ───────── 정류소ID 찾기: 이미 승인된 '버스노선 조회' API 사용 ─────────
// 이 정류장을 지나는 노선의 정류소 목록에서 표시번호(20675)가 일치하는 곳을 찾음
function findStationIdByRoute_() {
  const key = encodeURIComponent(CONFIG.SERVICE_KEY);
  for (const routeNo of CONFIG.ROUTES) {
    const listUrl = `${BASE}/busrouteservice/v2/getBusRouteListv2` +
      `?serviceKey=${key}&keyword=${encodeURIComponent(routeNo)}&format=json`;
    const lb = fetchJson_(listUrl);
    const routes = asArray_(lb && lb.response && lb.response.msgBody &&
                            lb.response.msgBody.busRouteList)
      .filter(r => String(r.routeName).trim() === routeNo);
    Logger.log(`노선 ${routeNo}: 후보 ${routes.length}개`);

    for (const r of routes) {
      const stUrl = `${BASE}/busrouteservice/v2/getBusRouteStationListv2` +
        `?serviceKey=${key}&routeId=${r.routeId}&format=json`;
      const sb = fetchJson_(stUrl);
      const stations = asArray_(sb && sb.response && sb.response.msgBody &&
                                sb.response.msgBody.busRouteStationList);
      const hit = stations.find(s => String(s.mobileNo).trim() === CONFIG.STATION_NO);
      if (hit) {
        Logger.log(`정류소: ${hit.stationName} / stationId=${hit.stationId} (노선 ${routeNo}, ${r.regionName || ''})`);
        return String(hit.stationId);
      }
    }
  }
  throw new Error('정류소를 찾지 못했습니다. 위 로그를 캡처해서 보내주세요.');
}

// ───────── 보조 함수 ─────────
function passRow_(route, prev, now) {
  // 통과시각 ≈ 마지막 관측시각 + 그때의 도착예정(분), 단 현재 시각을 넘지 않게
  const est = Math.min(prev.t + (Number(prev.predict) || 0) * 60000, now.getTime());
  const d = new Date(est);
  return [
    Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd'),
    '일월화수목금토'[d.getDay()],
    route,
    prev.plate,
    Utilities.formatDate(d, 'Asia/Seoul', 'HH:mm'),
    Utilities.formatDate(new Date(prev.t), 'Asia/Seoul', 'HH:mm'),
    prev.predict,
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

function asArray_(x) { return !x ? [] : Array.isArray(x) ? x : [x]; }
function num_(x) { return (x === undefined || x === null || x === '') ? '' : Number(x); }

function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}
