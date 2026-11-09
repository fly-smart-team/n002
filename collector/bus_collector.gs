/**
 * 버스 통과시각 수집기 (여러 정류장) — Google Apps Script
 *
 * 수집 구간 (TRIPS)
 *  - 덕이중학교 → 대화역·탄현역 방면 : 20675 (정류장 번호 고정)
 *  - 대화역 → 덕이동   : 10, N002, 67
 *  - 탄현역 → 덕이동   : 074, 071
 *  - 대화역 → 광성교회 : 80, 600, 900
 *  번호를 모르는 정류장은 setup이 노선 경로에서 자동으로 찾습니다
 *  ("출발지 이름이 들어간 정류장 → 몇 정거장 뒤 목적지 이름" 순서인 쪽).
 *
 * 사용법
 *  1) 기존 Code.gs 내용을 전부 지우고 이 코드 붙여넣기
 *  2) SERVICE_KEY에 data.go.kr 일반 인증키 입력
 *  3) 함수 선택에서 setup 실행 → 로그에서 찾은 정류장 이름·번호 확인
 *  4) 웹 앱을 이미 배포했다면: 배포 → 배포 관리 → 연필 → 버전 '새 버전' → 배포
 *
 * 하루 호출 한도(개발계정 1,000회)에 맞춰 정류장 수에 따라 수집 간격을 자동으로 늘립니다.
 * 운영계정으로 한도를 올리면 DAILY_BUDGET만 바꾸면 됩니다.
 */

const CONFIG = {
  SERVICE_KEY: '여기에_data.go.kr_인증키_입력',
  TRIPS: [
    { group: 'deogi',      label: '덕이중학교',        stop: '20675', routes: ['074', '071', 'N002', '10'] },
    { group: 'daehwa',     label: '대화역 → 덕이동',   from: ['대화역'], to: '덕이', routes: ['10', 'N002', '67'] },
    { group: 'tanhyeon',   label: '탄현역 → 덕이동',   from: ['탄현역'], to: '덕이', routes: ['074', '071'] },
    { group: 'gwangseong', label: '대화역 → 광성교회', from: ['대화역'], to: '광성교회', routes: ['80', '600', '900'], maxAhead: 40 },
  ],
  MAX_AHEAD: 25,              // 출발 정류장에서 목적지까지 최대 정거장 수
  ACTIVE_START: 5 * 60,       // 05:00부터
  ACTIVE_END: 24 * 60 + 30,   // 다음날 00:30까지
  DAILY_BUDGET: 880,          // 수집에 쓸 하루 호출 수 (나머지는 실시간 웹 앱용)
};

const BASE = 'https://apis.data.go.kr/6410000';
const RAW = 'raw';       // 매 호출 원본 기록
const PASS = 'passes';   // 버스 통과 추정 시각 → 시간표의 재료

// ───────── 최초 1회 / 정류장 바꿀 때 실행 ─────────
function setup() {
  const props = PropertiesService.getScriptProperties();
  const legacyId = props.getProperty('STATION_ID') || props.getProperty('SID_20675');
  const legacyState = props.getProperty('STATE');
  if (legacyState && !props.getProperty('STATE_20675')) props.setProperty('STATE_20675', legacyState);

  // 정류장 찾기 → 같은 정류장은 하나로 합침
  const stations = {};
  CONFIG.TRIPS.forEach(trip => {
    if (trip.stop) {
      const id = trip.stop === '20675' && legacyId ? legacyId : findStopById_(trip.stop, trip.routes);
      addStation_(stations, { no: trip.stop, id, name: trip.label }, trip, trip.routes);
      return;
    }
    trip.routes.forEach(routeNo => {
      const hit = findStopOnRoute_(routeNo, trip);
      if (hit) addStation_(stations, hit, trip, [routeNo]);
    });
  });
  const list = Object.keys(stations).map(k => stations[k]);
  props.setProperty('STATIONS', JSON.stringify(list));

  const interval = intervalFor_(list.length);
  Logger.log('──────── 수집할 정류장 ────────');
  list.forEach(st => Logger.log(`${st.no} ${st.name} (id ${st.id}) · ${st.groups.join(', ')} · 노선 ${st.routes.join(', ')}`));
  Logger.log(`정류장 ${list.length}곳, 각 ${interval}분 간격 → 하루 약 ${Math.round(list.length * (CONFIG.ACTIVE_END - CONFIG.ACTIVE_START) / interval)}회`);

  const ss = SpreadsheetApp.getActive();
  sheet_(ss, RAW, ['시각', '요일', '노선', '차량1', '도착예정1(분)', '정류장전1',
                   '차량2', '도착예정2(분)', '정류장전2', '정류장']);
  sheet_(ss, PASS, ['날짜', '요일', '노선', '차량번호', '통과추정시각', '마지막관측시각', '마지막예정(분)', '정류장']);
  [[RAW, 10], [PASS, 8]].forEach(([name, col]) => {
    const sh = ss.getSheetByName(name);
    if (!sh.getRange(1, col).getValue()) sh.getRange(1, col).setValue('정류장');
  });

  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('poll').timeBased().everyMinutes(1).create();
  Logger.log('설치 완료.');
  poll(true);
}

function addStation_(stations, hit, trip, routes) {
  const st = stations[hit.id] || (stations[hit.id] = { no: hit.no, id: hit.id, name: hit.name, groups: [], routes: [] });
  if (st.groups.indexOf(trip.group) < 0) st.groups.push(trip.group);
  routes.forEach(r => { if (st.routes.indexOf(r) < 0) st.routes.push(r); });
}

function intervalFor_(n) {
  return Math.max(2, Math.ceil(n * (CONFIG.ACTIVE_END - CONFIG.ACTIVE_START) / CONFIG.DAILY_BUDGET));
}

// ───────── 트리거가 1분마다 호출 ─────────
// 정류장마다 시작 분을 엇갈려서 한 번에 몰리지 않게 함
function poll(force) {
  const now = new Date();
  const mins = now.getHours() * 60 + now.getMinutes();
  const m = mins < 3 * 60 ? mins + 24 * 60 : mins;
  const list = JSON.parse(PropertiesService.getScriptProperties().getProperty('STATIONS') || '[]');
  if (!list.length) { Logger.log('정류장 목록이 없습니다. setup을 실행하세요.'); return; }
  const interval = intervalFor_(list.length);
  if (force !== true && (m < CONFIG.ACTIVE_START || m > CONFIG.ACTIVE_END)) return;
  list.forEach((st, k) => {
    if (force !== true && mins % interval !== k % interval) return;
    try { pollStation_(st, now); } catch (e) { Logger.log(`${st.name} 오류: ${e}`); }
  });
}

function pollStation_(st, now) {
  const props = PropertiesService.getScriptProperties();
  const list = fetchArrivals_(st.id);
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
  const all = JSON.parse(PropertiesService.getScriptProperties().getProperty('STATIONS') || '[]');
  const no = (e && e.parameter && e.parameter.stop) || '20675';
  if (no === 'list') return json_({ ok: true, stations: all.map(s => ({ no: s.no, name: s.name, groups: s.groups, routes: s.routes })) });
  const st = all.find(s => s.no === no);
  if (!st) return json_({ ok: false, error: 'unknown stop' });

  const cache = CacheService.getScriptCache();
  let json = cache.get('live_' + no);
  if (!json) {
    const list = fetchArrivals_(st.id);
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

// ───────── 정류장 찾기: '버스노선 조회' API 사용 ─────────
function routesNamed_(routeNo) {
  const url = `${BASE}/busrouteservice/v2/getBusRouteListv2` +
    `?serviceKey=${encodeURIComponent(CONFIG.SERVICE_KEY)}&keyword=${encodeURIComponent(routeNo)}&format=json`;
  const lb = fetchJson_(url);
  const all = asArray_(lb && lb.response && lb.response.msgBody && lb.response.msgBody.busRouteList)
    .filter(r => String(r.routeName).trim() === routeNo);
  const goyang = all.filter(r => String(r.regionName || '').indexOf('고양') >= 0);
  return goyang.length ? goyang : all;
}

function routeStations_(routeId) {
  const url = `${BASE}/busrouteservice/v2/getBusRouteStationListv2` +
    `?serviceKey=${encodeURIComponent(CONFIG.SERVICE_KEY)}&routeId=${routeId}&format=json`;
  const sb = fetchJson_(url);
  return asArray_(sb && sb.response && sb.response.msgBody && sb.response.msgBody.busRouteStationList)
    .sort((a, b) => Number(a.stationSeq) - Number(b.stationSeq));
}

// 정류장 번호를 알 때: 그 번호가 있는 노선에서 stationId를 찾음
function findStopById_(stopNo, routes) {
  for (const routeNo of routes) {
    for (const r of routesNamed_(routeNo)) {
      const hit = routeStations_(r.routeId).find(s => String(s.mobileNo).trim() === stopNo);
      if (hit) { Logger.log(`정류장 ${stopNo}: ${hit.stationName} (노선 ${routeNo})`); return String(hit.stationId); }
    }
  }
  throw new Error(`정류장 ${stopNo}을(를) 찾지 못했습니다.`);
}

// 정류장 번호를 모를 때: 노선 경로에서 'from' 이름 정류장 → 'to' 이름 정류장이 가장 빨리 나오는 쪽
// (같은 이름 정류장이 양방향·회차 전후로 여러 번 나오므로 목적지까지 정거장 수가 가장 적은 곳을 고름)
function findStopOnRoute_(routeNo, trip) {
  const routes = routesNamed_(routeNo);
  if (!routes.length) { Logger.log(`[${trip.label}] 노선 ${routeNo}: 노선을 찾지 못함`); return null; }
  const maxAhead = trip.maxAhead || CONFIG.MAX_AHEAD;
  const seen = [];
  for (const name of trip.from) {
    let best = null;
    for (const r of routes) {
      const sts = routeStations_(r.routeId);
      seen.push(sts);
      for (let i = 0; i < sts.length; i++) {
        if (String(sts[i].stationName).indexOf(name) < 0) continue;
        for (let j = i + 1; j < Math.min(sts.length, i + 1 + maxAhead); j++) {
          if (String(sts[j].stationName).indexOf(trip.to) >= 0) {
            if (!best || j - i < best.n) best = { s: sts[i], t: sts[j], n: j - i };
            break;
          }
        }
      }
    }
    if (best) {
      Logger.log(`[${trip.label}] 노선 ${routeNo}: ${best.s.stationName} (${best.s.mobileNo}) → ${best.n}정거장 뒤 ${best.t.stationName}`);
      return { no: String(best.s.mobileNo).trim(), id: String(best.s.stationId), name: best.s.stationName };
    }
  }
  // 못 찾았을 때: 경로에 출발지·목적지 이름이 있는지 보여줌
  const names = [].concat.apply([], seen).map(s => String(s.stationName));
  const uniq = (a) => a.filter((x, i) => a.indexOf(x) === i);
  const fromHits = uniq(names.filter(n => trip.from.some(f => n.indexOf(f) >= 0)));
  const toHits = uniq(names.filter(n => n.indexOf(trip.to) >= 0));
  Logger.log(`[${trip.label}] 노선 ${routeNo}: 찾지 못함 (노선 ${routes.map(r => r.routeName + '/' + (r.regionName || '')).join(', ')}) ` +
    `출발지 후보: ${fromHits.join(', ') || '없음'} / 목적지 후보: ${toHits.join(', ') || '없음'}`);
  return null;
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
