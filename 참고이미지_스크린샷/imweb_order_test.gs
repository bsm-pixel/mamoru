/**
 * 아임웹 V2 주문 API 실측 테스트 (Step 0 ~ Step 5)
 * ------------------------------------------------------------------
 * 사용법
 *  1) script.google.com → 새 프로젝트 → 이 코드 전체 붙여넣기
 *  2) 아래 CONFIG 5개 값만 채우기
 *  3) 상단 함수 선택창에서 [RUN_ALL] 선택 → ▶실행
 *  4) 하단 [실행 로그] 전체를 복사해서 그대로 전달
 * ------------------------------------------------------------------
 * 안전장치
 *  - STEP4_송장입력 / STEP5_발송처리 는 기본값 false (실행 안 함)
 *  - 조회(Step1) / 배송대기전환(Step2/3) 까지만 먼저 돌려보고,
 *    결과 확인 후 true 로 바꿔서 다시 실행하세요.
 *  - Step5(발송처리)는 고객에게 발송 알림이 나갈 수 있습니다.
 */

// ================== CONFIG ==================
var CONFIG = {
  KEY:        'API_KEY_여기에',       // 아임웹 관리자 > 환경설정 > 외부서비스연동(API) > Rest API V2
  SECRET:     'API_SECRET_여기에',
  ORDER_NO:   '주문번호_여기에',       // 결제완료 상태 주문 1건
  PON_OVERRIDE: '',                    // 특정 품목주문번호만 테스트 시 입력(예 '...-001'). 비우면 첫 품목 자동
  ORDER_VERSION: 'v2',                // 주문관리 v2면 'v2', 구주문이면 'v1'

  STEP4_송장입력: false,               // true 로 바꾸면 송장번호 입력 시도
  PARCEL_CODE: 'LOTTE',               // 택배사 코드 — 프로덕션 updateInvoice가 'LOTTE'(대문자)로 실동작 확인됨. 소문자 금지
  INVOICE_NO:  '0000000000000',

  STEP5_발송처리: false,               // true 면 발송처리(배송중 전환) 시도 — 고객 알림 발송 주의

  STEP6_강제배송완료: false             // true 면 강제 배송완료(complete) 시도 — 0원 테스트주문에서만!
};
// ============================================

var BASE = 'https://api.imweb.me';
var TOKEN = '';

function RUN_ALL() {
  log('==================================================');
  log('아임웹 V2 주문 API 실측 시작  /  order_version=' + CONFIG.ORDER_VERSION);
  log('주문번호: ' + CONFIG.ORDER_NO);
  log('==================================================');

  // ---------- STEP 0 : 토큰 발급 ----------
  section('STEP 0 — 토큰 발급  POST /v2/auth');
  var auth = call('post', '/v2/auth', { key: String(CONFIG.KEY).trim(), secret: String(CONFIG.SECRET).trim() }, false);
  TOKEN = pick(auth.json, 'access_token');
  if (!TOKEN) { log('!! access_token 을 찾지 못했습니다. 위 응답 원문을 확인하세요. 중단.'); return; }
  log('>> access_token 확보 (길이 ' + TOKEN.length + ')');

  // ---------- STEP 1 : 주문 + 품목주문 조회 ----------
  section('STEP 1-A — 주문 조회  GET /v2/shop/orders/' + CONFIG.ORDER_NO);
  call('get', '/v2/shop/orders/' + CONFIG.ORDER_NO + qs());

  section('STEP 1-B — 품목주문 조회  GET /v2/shop/orders/' + CONFIG.ORDER_NO + '/prod-orders');
  var po = call('get', '/v2/shop/orders/' + CONFIG.ORDER_NO + '/prod-orders' + qs());
  var prodOrderNos = extractProdOrderNos(po.json);
  log('>> 추출된 품목주문번호: ' + JSON.stringify(prodOrderNos));
  if (!prodOrderNos.length) { log('!! 품목주문번호를 찾지 못했습니다. 중단.'); return; }
  // PON_OVERRIDE 가 있으면 그 품목주문번호를 사용(특정 품목만 테스트), 없으면 첫 번째
  var PON = (CONFIG.PON_OVERRIDE && String(CONFIG.PON_OVERRIDE).trim()) ? String(CONFIG.PON_OVERRIDE).trim() : prodOrderNos[0];
  log('>> 이번 테스트 대상 품목주문번호: ' + PON);

  // ---------- STEP 2 : 배송대기 전환 ----------
  section('STEP 2 — 배송대기 전환  PATCH /v2/shop/prod-orders/' + PON + '/place' + qs());
  var r2 = call('patch', '/v2/shop/prod-orders/' + PON + '/place' + qs(), {});

  if (!isOk(r2)) {
    section('STEP 2-a — 재시도: 쿼리 없이 body 에만 order_version');
    call('patch', '/v2/shop/prod-orders/' + PON + '/place', { order_version: CONFIG.ORDER_VERSION });

    section('STEP 2-b — 재시도: 쿼리+body 동시');
    call('patch', '/v2/shop/prod-orders/' + PON + '/place' + qs(), { order_version: CONFIG.ORDER_VERSION });
  }

  // ---------- STEP 3 : 전환 결과 재확인 ----------
  section('STEP 3 — 전환 결과 재확인  GET .../prod-orders');
  var po2 = call('get', '/v2/shop/orders/' + CONFIG.ORDER_NO + '/prod-orders' + qs());
  log('>> 현재 status: ' + JSON.stringify(extractStatuses(po2.json)));

  // ---------- STEP 4 : 송장번호 입력 ----------
  if (CONFIG.STEP4_송장입력) {
    section('STEP 4 — 송장 입력  PATCH /v2/shop/prod-orders/' + PON + '/invoice');
    call('patch', '/v2/shop/prod-orders/' + PON + '/invoice' + qs(), {
      parcel_code: CONFIG.PARCEL_CODE,
      invoice_no:  CONFIG.INVOICE_NO,
      order_version: CONFIG.ORDER_VERSION
    });

    section('STEP 4-B — 송장 조회  GET /v2/shop/prod-orders/' + PON + '/invoice');
    call('get', '/v2/shop/prod-orders/' + PON + '/invoice' + qs());

    section('STEP 4-C — 송장등록 直後 상태확인 (invoice만으로 배송중 가는지)');
    var po4c = call('get', '/v2/shop/orders/' + CONFIG.ORDER_NO + '/prod-orders' + qs());
    log('>> invoice 直後 status: ' + JSON.stringify(extractStatuses(po4c.json)));
  } else {
    section('STEP 4 — 건너뜀 (CONFIG.STEP4_송장입력 = false)');
  }

  // ---------- STEP 5 : 발송처리 ----------
  if (CONFIG.STEP5_발송처리) {
    section('STEP 5 — 발송처리  PATCH /v2/shop/prod-orders/' + PON + '/send');
    call('patch', '/v2/shop/prod-orders/' + PON + '/send' + qs(), {});

    section('STEP 5-B — 최종 재확인');
    var po3 = call('get', '/v2/shop/orders/' + CONFIG.ORDER_NO + '/prod-orders' + qs());
    log('>> 최종 status: ' + JSON.stringify(extractStatuses(po3.json)));
  } else {
    section('STEP 5 — 건너뜀 (CONFIG.STEP5_발송처리 = false)');
  }

  // ---------- STEP 6 : 강제 배송완료 (complete) ----------
  if (CONFIG.STEP6_강제배송완료) {
    section('STEP 6 — 강제 배송완료  PATCH /v2/shop/prod-orders/' + PON + '/complete');
    call('patch', '/v2/shop/prod-orders/' + PON + '/complete' + qs(), {});

    section('STEP 6-B — 최종 재확인');
    var po4 = call('get', '/v2/shop/orders/' + CONFIG.ORDER_NO + '/prod-orders' + qs());
    log('>> 최종 status: ' + JSON.stringify(extractStatuses(po4.json)));
  } else {
    section('STEP 6 — 건너뜀 (CONFIG.STEP6_강제배송완료 = false)');
  }

  log('\n==================================================');
  log('테스트 종료. 위 로그 전체를 복사해서 전달하세요.');
  log('==================================================');
}

// ---------------- helpers ----------------
function qs() { return '?order_version=' + CONFIG.ORDER_VERSION; }

function call(method, path, body, useToken) {
  if (useToken === undefined) useToken = true;
  var opt = {
    method: method,
    muteHttpExceptions: true,
    contentType: 'application/json',
    headers: {}
  };
  if (useToken) opt.headers['access-token'] = TOKEN;
  if (body !== undefined && body !== null) opt.payload = JSON.stringify(body);

  log('REQUEST  ' + method.toUpperCase() + ' ' + BASE + path);
  if (body !== undefined && body !== null) log('BODY     ' + JSON.stringify(body).replace(CONFIG.SECRET, '***SECRET***').replace(CONFIG.KEY, '***KEY***'));

  var res, text;
  try {
    res = UrlFetchApp.fetch(BASE + path, opt);
    text = res.getContentText();
    log('HTTP     ' + res.getResponseCode());
  } catch (e) {
    log('EXCEPTION ' + e);
    return { json: null, text: '' };
  }
  log('RESPONSE ' + text);

  var json = null;
  try { json = JSON.parse(text); } catch (e) { log('(JSON 파싱 실패)'); }
  return { json: json, text: text, code: res.getResponseCode() };
}

function isOk(r) {
  if (!r || !r.json) return false;
  var c = r.json.code;
  return (c === 200 || c === 0 || c === '200');
}

function pick(obj, key) {
  if (!obj || typeof obj !== 'object') return null;
  if (obj[key]) return obj[key];
  for (var k in obj) {
    var v = pick(obj[k], key);
    if (v) return v;
  }
  return null;
}

function extractProdOrderNos(json) {
  var out = [];
  (function walk(o) {
    if (!o || typeof o !== 'object') return;
    if (Array.isArray(o)) { o.forEach(walk); return; }
    if (o.order_no && (o.status !== undefined || o.items !== undefined)) out.push(String(o.order_no));
    if (o.prod_order_no) out.push(String(o.prod_order_no));
    for (var k in o) walk(o[k]);
  })(json && json.data !== undefined ? json.data : json);
  return out.filter(function (v, i, a) { return a.indexOf(v) === i; });
}

function extractStatuses(json) {
  var out = [];
  (function walk(o) {
    if (!o || typeof o !== 'object') return;
    if (Array.isArray(o)) { o.forEach(walk); return; }
    if (o.status && o.order_no) out.push(o.order_no + ' → ' + o.status);
    for (var k in o) walk(o[k]);
  })(json && json.data !== undefined ? json.data : json);
  return out;
}

function section(t) { log('\n---------- ' + t + ' ----------'); }
function log(m) { Logger.log(m); console.log(m); }
