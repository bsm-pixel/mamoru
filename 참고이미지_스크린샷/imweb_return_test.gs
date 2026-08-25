/**
 * 아임웹 V2 "취소·반품·클레임" API 실측 테스트
 * ------------------------------------------------------------------
 * 목적: TMS→아임웹으로 주문 취소/반품/교환(클레임)을 자동 접수할 수 있는지,
 *       어떤 엔드포인트가 실제로 받아주는지를 실호출로 판정한다.
 *       (place/send는 이미 검증됨 — 이건 그 다음 미검증 영역)
 *
 * 사용법
 *  1) script.google.com → 새 프로젝트 → 이 코드 전체 붙여넣기
 *  2) 아래 CONFIG 채우기 — 반드시 "버려도 되는 0원 테스트 주문"으로!
 *  3) 함수 선택창 [RUN_ALL] → ▶실행
 *  4) 하단 [실행 로그] 전체 복사해서 그대로 전달
 * ------------------------------------------------------------------
 * ⚠️ 파괴적 테스트: 취소/반품이 실제로 걸릴 수 있음 → 실제 고객 주문 금지, 0원 테스트 주문만.
 *   - STEP D(조회/탐색)는 안전(읽기). 먼저 D만 돌려 결과 보고,
 *   - 실제 취소 시도(STEP E~)는 TRY_* 플래그를 하나씩 true 로 바꿔 재실행.
 */

// ================== CONFIG ==================
var CONFIG = {
  KEY:        'API_KEY_여기에',        // 아임웹 관리자 > 환경설정 > 외부서비스연동(API) > Rest API V2
  SECRET:     'API_SECRET_여기에',
  ORDER_NO:   '테스트_주문번호_여기에',  // 버려도 되는 0원/저가 테스트 주문 1건 (결제완료 or 배송대기)
  PON_OVERRIDE: '',                     // 특정 품목주문번호만. 비우면 첫 품목 자동
  ORDER_VERSION: 'v2',

  // --- 파괴적 시도 스위치 (기본 전부 false = STEP D 탐색만 실행) ---
  TRY_PRODORDER_CANCEL: false,   // PATCH /v2/shop/prod-orders/{PON}/cancel
  TRY_ORDER_CANCEL:     false,   // PATCH /v2/shop/orders/{ORDER_NO}/cancel
  TRY_STATUS_CANCEL:    false,   // PATCH /v2/shop/prod-orders/{PON}  body {status:'CANCEL'} (구 방식)
  TRY_PRODORDER_RETURN: false,   // PATCH /v2/shop/prod-orders/{PON}/return  (반품)
  TRY_PRODORDER_REFUND: false,   // PATCH /v2/shop/prod-orders/{PON}/refund  (환불)

  CANCEL_REASON: '테스트 취소'    // 취소/반품 사유 (필요 시 body 에 실림)
};
// ============================================

var BASE = 'https://api.imweb.me';
var TOKEN = '';

function RUN_ALL() {
  log('==================================================');
  log('아임웹 취소/반품/클레임 실측  /  order_version=' + CONFIG.ORDER_VERSION);
  log('주문번호: ' + CONFIG.ORDER_NO + '  (⚠️ 버려도 되는 테스트 주문이어야 함)');
  log('==================================================');

  // STEP 0 토큰
  section('STEP 0 — 토큰 발급  POST /v2/auth');
  var auth = call('post', '/v2/auth', { key: String(CONFIG.KEY).trim(), secret: String(CONFIG.SECRET).trim() }, false);
  TOKEN = pick(auth.json, 'access_token');
  if (!TOKEN) { log('!! access_token 실패. 중단.'); return; }
  log('>> access_token 확보 (길이 ' + TOKEN.length + ')');

  // STEP 1 주문/품목 조회 → PON
  section('STEP 1 — 품목주문 조회');
  var po = call('get', '/v2/shop/orders/' + CONFIG.ORDER_NO + '/prod-orders' + qs());
  var pons = extractProdOrderNos(po.json);
  log('>> 품목주문번호: ' + JSON.stringify(pons));
  if (!pons.length) { log('!! 품목주문번호 없음. 중단.'); return; }
  var PON = (CONFIG.PON_OVERRIDE && String(CONFIG.PON_OVERRIDE).trim()) ? String(CONFIG.PON_OVERRIDE).trim() : pons[0];
  log('>> 대상 품목주문번호: ' + PON);
  log('>> 현재 status: ' + JSON.stringify(extractStatuses(po.json)));

  // ============ STEP D — 안전 탐색 (읽기): 클레임/취소 리소스가 있는지 ============
  section('STEP D-1 — 클레임 목록 조회 시도  GET /v2/shop/claims');
  call('get', '/v2/shop/claims' + qs());

  section('STEP D-2 — 주문의 클레임 조회 시도  GET /v2/shop/orders/{no}/claims');
  call('get', '/v2/shop/orders/' + CONFIG.ORDER_NO + '/claims' + qs());

  section('STEP D-3 — 품목주문 취소가능여부/상세  GET /v2/shop/prod-orders/{PON}');
  call('get', '/v2/shop/prod-orders/' + PON + qs());

  // ============ STEP E — 파괴적 시도 (플래그로 하나씩) ============
  if (CONFIG.TRY_PRODORDER_CANCEL) {
    section('STEP E1 — 품목주문 취소  PATCH /v2/shop/prod-orders/' + PON + '/cancel');
    call('patch', '/v2/shop/prod-orders/' + PON + '/cancel' + qs(), { reason: CONFIG.CANCEL_REASON, order_version: CONFIG.ORDER_VERSION });
    recheck(PON);
  } else { section('STEP E1 — 건너뜀 (TRY_PRODORDER_CANCEL=false)'); }

  if (CONFIG.TRY_ORDER_CANCEL) {
    section('STEP E2 — 주문 전체 취소  PATCH /v2/shop/orders/' + CONFIG.ORDER_NO + '/cancel');
    call('patch', '/v2/shop/orders/' + CONFIG.ORDER_NO + '/cancel' + qs(), { reason: CONFIG.CANCEL_REASON, order_version: CONFIG.ORDER_VERSION });
    recheck(PON);
  } else { section('STEP E2 — 건너뜀 (TRY_ORDER_CANCEL=false)'); }

  if (CONFIG.TRY_STATUS_CANCEL) {
    section('STEP E3 — status=CANCEL 직접 PATCH  /v2/shop/prod-orders/' + PON);
    call('patch', '/v2/shop/prod-orders/' + PON + qs(), { status: 'CANCEL', order_version: CONFIG.ORDER_VERSION });
    recheck(PON);
  } else { section('STEP E3 — 건너뜀 (TRY_STATUS_CANCEL=false)'); }

  if (CONFIG.TRY_PRODORDER_RETURN) {
    section('STEP E4 — 반품  PATCH /v2/shop/prod-orders/' + PON + '/return');
    call('patch', '/v2/shop/prod-orders/' + PON + '/return' + qs(), { reason: CONFIG.CANCEL_REASON, order_version: CONFIG.ORDER_VERSION });
    recheck(PON);
  } else { section('STEP E4 — 건너뜀 (TRY_PRODORDER_RETURN=false)'); }

  if (CONFIG.TRY_PRODORDER_REFUND) {
    section('STEP E5 — 환불  PATCH /v2/shop/prod-orders/' + PON + '/refund');
    call('patch', '/v2/shop/prod-orders/' + PON + '/refund' + qs(), { reason: CONFIG.CANCEL_REASON, order_version: CONFIG.ORDER_VERSION });
    recheck(PON);
  } else { section('STEP E5 — 건너뜀 (TRY_PRODORDER_REFUND=false)'); }

  log('\n==================================================');
  log('종료. 위 로그 전체 복사해서 전달.');
  log('판정 힌트: HTTP 200 + code 200/0 + data.success = 그 엔드포인트 "됨".');
  log('           HTTP 404 = 엔드포인트 없음.  code -11/실패메시지 = 존재하나 현재 상태선 불가.');
  log('==================================================');
}

function recheck(PON) {
  section('   ↳ 재확인  GET .../prod-orders');
  var po2 = call('get', '/v2/shop/orders/' + CONFIG.ORDER_NO + '/prod-orders' + qs());
  log('   >> status: ' + JSON.stringify(extractStatuses(po2.json)));
}

// ---------------- helpers (imweb_order_test.gs 동일) ----------------
function qs() { return '?order_version=' + CONFIG.ORDER_VERSION; }
function call(method, path, body, useToken) {
  if (useToken === undefined) useToken = true;
  var opt = { method: method, muteHttpExceptions: true, contentType: 'application/json', headers: {} };
  if (useToken) opt.headers['access-token'] = TOKEN;
  if (body !== undefined && body !== null) opt.payload = JSON.stringify(body);
  log('REQUEST  ' + method.toUpperCase() + ' ' + BASE + path);
  if (body !== undefined && body !== null) log('BODY     ' + JSON.stringify(body).replace(CONFIG.SECRET, '***').replace(CONFIG.KEY, '***'));
  var res, text;
  try { res = UrlFetchApp.fetch(BASE + path, opt); text = res.getContentText(); log('HTTP     ' + res.getResponseCode()); }
  catch (e) { log('EXCEPTION ' + e); return { json: null, text: '' }; }
  log('RESPONSE ' + text);
  var json = null; try { json = JSON.parse(text); } catch (e) { log('(JSON 파싱 실패)'); }
  return { json: json, text: text, code: res.getResponseCode() };
}
function pick(obj, key) {
  if (!obj || typeof obj !== 'object') return null;
  if (obj[key]) return obj[key];
  for (var k in obj) { var v = pick(obj[k], key); if (v) return v; }
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
