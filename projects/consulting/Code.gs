/***** MAMORU 상담접수 — GAS Web App (v2.2.0-final) *****/

const SPREADSHEET_ID = '1EfiXtuJ7bQDnWHfX-XwOxujOssYKkMvtDO5CxPGYWj8';
const SHEET_NAME = '상담접수';
const DEALER_SHEET_NAME = '딜러';
const VERSION = 'v2.2.0-final';

// GitHub Pages URL (index.html 이전용)
const GITHUB_PAGES_BASE = 'https://bsm-pixel.github.io/mamoru/projects/01_consulting';
// GitHub Pages — 고객 대면 페이지 (Suggest/Reschedule/DealerConfirm/Result)
const GITHUB_PAGES_CONSULT = 'bsm-pixel.github.io/mamoru/projects/consulting';

const MAKE_WEBHOOK_URL = 'https://hook.eu2.make.com/x70jrebxr82odam3eu4oosuar9mm2fg5';
const CAL_NAME = '마모루 방문예약';

const BUSINESS = { startHour: 10, endHour: 20, durMin: 60 };
const SLOT_STEP_MIN_DEFAULT = 10;
const HOLD_PREFIX = '[HOLD]';
/**
 * 캘린더 이벤트 제목을 만드는 공통 함수
 * - row: 상담접수 행 객체 (name, phone, address 등을 포함)
 * - baseLabel: 필요하면 타입/메모 등을 뒤에 붙이고 싶을 때 사용하는 부제목(없으면 빈 문자열)
 * - isHold: HOLD 이벤트 여부 (true면 HOLD_PREFIX를 붙임)
 *
 * 최종 형태 예:
 *   [HOLD] 홍길동 / 010-1234-5678 / 서울시 강남구 어딘가 — 출장요청
 *   홍길동 / 010-1234-5678 / 마모루 방문 — 매장 방문
 */
function buildCalTitle_(row, baseLabel, isHold) {
  row = row || {};

  // 기본 값 정리
  var name = row.name ? String(row.name).trim() : '';

  // 전화번호 숫자만 추출 후 포맷팅
  var rawPhone = row.phone ? String(row.phone) : '';
  var digits   = rawPhone.replace(/\D/g, '');
  var phone    = '';

  if (digits.length === 11) {
    // 010-1234-5678 형태
    phone = digits.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3');
  } else if (digits.length === 10) {
    // 02-1234-5678 또는 031-123-4567 등 대응
    phone = digits.replace(/(\d{2,3})(\d{3,4})(\d{4})/, '$1-$2-$3');
  } else if (rawPhone) {
    // 길이가 애매하면 그대로 표시
    phone = rawPhone.trim();
  }

  var address = row.address ? String(row.address).trim() : '';

  // "이름 / 전화 / 주소" 형태로 조합
  var customerParts = [];
  if (name)    customerParts.push(name);
  if (phone)   customerParts.push(phone);
  if (address) customerParts.push(address);

  var customerLabel = customerParts.join(' / ');

  var pieces = [];

  // HOLD 이벤트면 앞에 [HOLD] 붙이기
  if (isHold) {
    pieces.push(HOLD_PREFIX);
  }

  // baseLabel(예: "매장 방문", "출장 방문" 등)과 고객 정보를 조합
  if (baseLabel && customerLabel) {
    // 예: [HOLD] 출장 방문 — 홍길동 / 010-XXXX-XXXX / 서울 강남구 …
    pieces.push(baseLabel);
    pieces.push('— ' + customerLabel);
  } else if (baseLabel) {
    // 예: [HOLD] 출장 방문
    pieces.push(baseLabel);
  } else if (customerLabel) {
    // 예: 홍길동 / 010-XXXX-XXXX / 서울 강남구 …
    pieces.push(customerLabel);
  } else {
    // 이름/전화/주소, baseLabel이 모두 비어있는 예외 상황 대비
    pieces.push('예약');
  }

  return pieces.join(' ').replace(/\s+/g, ' ').trim();
}


// 방문/출장 상호 차단 버퍼(분)
const BLOCK_BUFFER_MIN = 60;
const FIELD_BLOCK_BEFORE_MIN = 90;  // 출장 시작 1시간 30분 전
const FIELD_BLOCK_AFTER_MIN  = 90; // 출장 시작 2시간 후
const DEFAULT_DUR_MIN = (typeof BUSINESS === 'object' && BUSINESS.durMin) ? BUSINESS.durMin : 60;

// 캘린더명
const CAL_STORE_NAME = '마모루 방문예약';    // 매장
const CAL_FIELD_NAME = '마모루 출장방문';    // 출장

const TIMEZONE = 'Asia/Seoul';
function tzDate_(dateStr, timeStr){
  const [Y,M,D] = String(dateStr).split('-').map(n=>+n);
  const [h,m]   = String(timeStr||'00:00').split(':').map(n=>+n);
  return new Date(new Date(Date.UTC(Y, M-1, D, h, m)).toLocaleString('en-US', { timeZone: TIMEZONE }));
}
function addMin_(d, min){ return new Date(d.getTime() + min*60000); }
function overlap_(a0,a1,b0,b1){ return a0 < b1 && b0 < a1; }

function getCal_(name){
  const cals = CalendarApp.getCalendarsByName(name||'');
  if (!cals || !cals.length) throw new Error('Calendar not found: '+name);
  return cals[0];
}
function fetchEvents_(calName, start, end){
  return getCal_(calName).getEvents(start, end);
}

/** 버퍼 포함 차단 판정 */
function isBlockedRange_(start, end){
  const s = addMin_(start, -BLOCK_BUFFER_MIN);
  const e = addMin_(end,   BLOCK_BUFFER_MIN);
  const lists = [
    fetchEvents_(CAL_STORE_NAME, s, e),
    fetchEvents_(CAL_FIELD_NAME, s, e)
  ];
  for (const ev of [].concat(...lists)){
    const evS = ev.getStartTime(), evE = ev.getEndTime();
    if (overlap_(s,e, evS,evE)) return true;
  }
  return false;
}

const HOLD_EXPIRE_HOURS = 6;
const CLOSED_WEEKDAYS = [0];
const CLOSED_DATES = [];
const REMIND_HOURS = [24, 2];
const KEEP_BUFFER_HOURS = 12;
const ADMIN_KEY = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
const SLOTS_CACHE_PREFIX = 'slots:';
const SLOTS_CACHE_TTL_SEC = 900;

/****************************** 공통 ******************************/
function json(o){
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
function generateShortToken_() {
  return Math.random().toString(36).substr(2, 6).toUpperCase();
}

/****************************** 시트 ******************************/
function sh_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

  const baseHeaders = ['접수시각','성함','연락처','상담방식','방문일','방문시간','주소','가능요일','비고','원본'];
  const extraHeaders = [
    'UniqueID','Status','ConfirmationToken','CalendarEventID','Remind24','Remind2','선호시간대','제안내용','단축토큰',
    'holdEventIds','holdExpireAt',
    'addressZip','addressRoad','addressDetail','addressPlaceId','addressLat','addressLng'
  ];

  if (sh.getLastRow() === 0) {
    sh.appendRow([...baseHeaders, ...extraHeaders]);
    return sh;
  }
  const range = sh.getRange(1, 1, 1, sh.getLastColumn());
  const headers = range.getValues()[0].map(String);
  const need = extraHeaders.filter(h => !headers.includes(h));
  if (need.length) {
    sh.getRange(1, headers.length + 1, 1, need.length).setValues([need]);
  }
  return sh;
}

function getClosedDatesFromSheet_(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const s  = ss.getSheetByName('휴무설정');
  if (!s) return [];
  const n = Math.max(s.getLastRow() - 1, 0);
  if (n === 0) return [];
  return s.getRange(2, 1, n, 1).getValues()
    .map(r => r[0] instanceof Date ? Utilities.formatDate(r[0], TIMEZONE, 'yyyy-MM-dd') : String(r[0]).trim())
    .filter(v => /^\d{4}-\d{2}-\d{2}$/.test(v));
}

function headerMap_(){
  const sh = sh_();
  const heads = sh.getRange(1,1,1, sh.getLastColumn()).getValues()[0].map(String);
  const map = {}; heads.forEach((h,i)=> map[h]=i);
  return map;
}

/****************************** 운영설정 로드 ******************************/
const SETTINGS_CACHE_KEY = 'settings:v1';
const SETTINGS_CACHE_TTL = 300;
function settingsCacheGet_(){
  try{ const raw = CacheService.getScriptCache().get(SETTINGS_CACHE_KEY); return raw ? JSON.parse(raw) : null; }catch(_){ return null; }
}
function settingsCachePut_(obj){
  try{ CacheService.getScriptCache().put(SETTINGS_CACHE_KEY, JSON.stringify(obj||{}), SETTINGS_CACHE_TTL); }catch(_){}
}
function readSettings_(){
  try{
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName('Settings') || ss.getSheetByName('운영설정');
    if (!sh) return {};
    const rows = sh.getRange(1,1, sh.getLastRow(), 2).getValues();
    const map = {};
    for (let i=1;i<rows.length;i++){
      const k = String(rows[i][0]||'').trim().toUpperCase();
      const v = String(rows[i][1]||'').trim();
      if (k) map[k] = v;
    }
    const toInt = s => { const n = parseInt(String(s).trim(), 10); return isNaN(n) ? null : n; };
    const out = {};
    if (map.CLOSED_WEEKDAYS) out.CLOSED_WEEKDAYS = map.CLOSED_WEEKDAYS.split(',').map(s=>+s.trim()).filter(n=>!isNaN(n));
    if (map.CLOSED_DATES)    out.CLOSED_DATES    = map.CLOSED_DATES.split(',').map(s=>s.trim()).filter(Boolean);
    const startH = toInt(map.BUSINESS_START_HOUR ?? map.START_HOUR);
    const endH   = toInt(map.BUSINESS_END_HOUR   ?? map.END_HOUR);
    const durM   = toInt(map.BUSINESS_DUR_MIN    ?? map.DUR_MIN);
    const stepM  = toInt(map.SLOT_STEP_MIN       ?? map.SLOT_STEP);
    if (startH!=null && startH>=0 && startH<=23) out.BUSINESS_START_HOUR = startH;
    if (endH!=null   && endH>=1 && endH<=24)     out.BUSINESS_END_HOUR   = endH;
    if (durM!=null   && durM>=5 && durM<=240)    out.BUSINESS_DUR_MIN    = durM;
    if (stepM!=null  && stepM>=5 && stepM<=60)   out.SLOT_STEP_MIN       = stepM;
    return out;
  }catch(e){ Logger.log('readSettings_ err: '+e); return {}; }
}
function getClosedWeekdays_(){ const s = readSettings_(); return s.CLOSED_WEEKDAYS ?? CLOSED_WEEKDAYS; }
function getClosedDates_(){    const s = readSettings_(); return s.CLOSED_DATES    ?? CLOSED_DATES; }
function getBusiness_(){
  const s = readSettings_();
  return {
    startHour: s.BUSINESS_START_HOUR ?? BUSINESS.startHour,
    endHour:   s.BUSINESS_END_HOUR   ?? BUSINESS.endHour,
    durMin:    s.BUSINESS_DUR_MIN    ?? BUSINESS.durMin,
    stepMin:   s.SLOT_STEP_MIN       ?? SLOT_STEP_MIN_DEFAULT
  };
}
function getSettings(){
  const cached = settingsCacheGet_();
  if (cached && cached.CLOSED_WEEKDAYS && cached.CLOSED_DATES && cached.BUSINESS) return cached;

  const w = getClosedWeekdays_() || [];
  const a = getClosedDates_() || [];
  const b = getClosedDatesFromSheet_() || [];
  const d = Array.from(new Set([...(a||[]), ...(b||[])]));
  const biz = getBusiness_();
  const out = { CLOSED_WEEKDAYS: w, CLOSED_DATES: d, BUSINESS: biz };
  settingsCachePut_(out);
  return out;
}

/****************************** 슬롯 캐시 ******************************/
function slotsCacheKey_(dateStr){ return SLOTS_CACHE_PREFIX + String(dateStr); }
function slotsCacheFingerprint_(){
  const biz = getBusiness_();
  return [ensureCalendar_(), biz.startHour, biz.endHour, biz.durMin, biz.stepMin, FIELD_BLOCK_BEFORE_MIN, FIELD_BLOCK_AFTER_MIN].join('|');
}
function slotsCacheGet_(dateStr){
  try{
    const c = CacheService.getScriptCache();
    const raw = c.get(slotsCacheKey_(dateStr));
    if (!raw) return null;
    const j = JSON.parse(raw);
    if (j && j.fp === slotsCacheFingerprint_() && Array.isArray(j.slots)) return j.slots;
  }catch(e){}
  return null;
}
function slotsCachePut_(dateStr, slots){
  try{
    const c = CacheService.getScriptCache();
    const payload = JSON.stringify({ fp: slotsCacheFingerprint_(), slots: slots||[] });
    c.put(slotsCacheKey_(dateStr), payload, SLOTS_CACHE_TTL_SEC);
  }catch(e){}
}
function slotsCacheInvalidate_(dateStr){
  try{
    const c = CacheService.getScriptCache();
    c.remove(slotsCacheKey_(dateStr));
    const m = /^(\d{4})-(\d{2})-\d{2}$/.exec(String(dateStr||''));
    if (m){
      const year = m[1], mm = m[2];
      const fp = slotsCacheFingerprint_();
      const MONTH_CACHE_KEY = `monthslots:${fp}:${year}-${mm}`;
      c.remove(MONTH_CACHE_KEY);
    }
  }catch(e){}
}

/****************************** 캘린더 ******************************/
function ensureCalendar_(){
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty('CAL_ID');
  if (id) return id;
  const cal = CalendarApp.createCalendar(CAL_NAME, { timeZone: Session.getScriptTimeZone() });
  id = cal.getId(); p.setProperty('CAL_ID', id); return id;
}
function makeDateTimes_(visitDate, visitTime){
  const biz = getBusiness_();
  const [yy,mm,dd] = String(visitDate).split('-').map(v=>+v);
  const [HH,MM]    = String(visitTime||'10:00').split(':').map(v=>+v);
  const start = new Date(yy, (mm||1)-1, dd||1, HH||10, MM||0, 0);
  const end   = new Date(start.getTime() + biz.durMin*60000);
  return {start,end};
}

function getAvailableTimeSlots_(yyyymmdd, type){
  return getAvailableTimeSlots_SheetBased_(yyyymmdd, type);
}


function getSlots(dateStr){
  const cal = CalendarApp.getCalendarsByName(CAL_NAME)[0];
  if (!cal) throw new Error('Calendar not found: '+CAL_NAME);

  const [y,m,d] = String(dateStr).split('-').map(Number);
  const dayStart = new Date(y, m-1, d, BUSINESS.startHour, 0, 0);
  const dayEnd   = new Date(y, m-1, d, BUSINESS.endHour, 0, 0);
  const stepMs   = (SLOT_STEP_MIN_DEFAULT||10) * 60 * 1000;
  const durMs    = (DEFAULT_DUR_MIN||60) * 60 * 1000;
  const bufMs    = (BLOCK_BUFFER_MIN||0) * 60 * 1000;
  const tz       = Session.getScriptTimeZone() || 'Asia/Seoul';

  // 하루 범위 이벤트 수집
  const events = cal.getEvents(dayStart, dayEnd);

  // 바쁜구간 리스트 구성: HOLD는 ±버퍼 확장
  const busy = events.map(ev=>{
    const t     = ev.getTitle() || '';
    const hold  = t.includes(HOLD_PREFIX);
    const s     = ev.getStartTime();
    const e     = ev.getEndTime();
    const start = new Date(s.getTime() - (hold ? bufMs : 0));
    const end   = new Date(e.getTime() + (hold ? bufMs : 0));
    return [start, end];
  });

  // 슬롯 생성
  const out = [];
  for (let t = new Date(dayStart.getTime()); t.getTime()+durMs <= dayEnd.getTime(); t = new Date(t.getTime()+stepMs)){
    const s = new Date(t.getTime());
    const e = new Date(t.getTime()+durMs);
    const overlap = busy.some(([bs,be]) => e > bs && s < be);
    if (!overlap){
      out.push(Utilities.formatDate(s, tz, 'HH:mm'));
    }
  }
  return out;
}



/****************************** 라우팅 ******************************/
function getSlotsRange(dateList){
  if (!Array.isArray(dateList) || dateList.length === 0) return {};
  const out = {};
  const groups = {};
  for (let d of dateList){
    const ds = String(d||'').slice(0,10);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(ds)) continue;
    const [y,m] = ds.split('-');
    const key = `${y}-${m}`;
    if (!groups[key]) groups[key] = [];
    groups[key].push(ds);
  }
  Object.keys(groups).forEach(key=>{
    const [yy,mm] = key.split('-').map(v=>+v);
    const monthMap = getSlotsMonth_(yy, mm);
    groups[key].forEach(ds=>{ out[ds] = monthMap[ds] || []; });
  });
  return out;
}


function getSlotsMonth_(year, month1to12, type){
  return getSlotsMonth_SheetBased_(year, month1to12, type);
}

const ADMIN_CREDENTIALS = {
  username: 'mamoru',      // 원하는 아이디
  password: 'wnlehddl05'   // 원하는 비밀번호
};

const SESSION_TIMEOUT_HOURS = 24;

function adminLogin(username, password) {
  try {
    // 1) 입력값 정규화 (앞뒤 공백 제거 + 문자열 변환)
    const inputUsername = String(username || '').trim();
    const inputPassword = String(password || '').trim();

    const refUsername = String(ADMIN_CREDENTIALS.username || '').trim();
    const refPassword = String(ADMIN_CREDENTIALS.password || '').trim();

    // 2) 디버깅용 로그 (실제 뭘 입력했는지 Apps Script 로그에서 확인 가능)
    Logger.log(JSON.stringify({
      tag: 'ADMIN_LOGIN_ATTEMPT',
      inputUsername: inputUsername,
      inputPasswordLength: inputPassword.length
    }));

    // 3) 아이디 비교 규칙
    //    - 'mamoru' (설정값)
    //    - 'admin' 도 임시 허용 → 둘 다 로그인 가능
    const usernameOk =
      inputUsername === refUsername ||
      inputUsername === 'admin';

    // 4) 비밀번호 비교 (공백 제거 후 정확히 일치해야 함)
    const passwordOk = (inputPassword === refPassword);

    if (usernameOk && passwordOk) {
      const sessionToken = Utilities.getUuid();
      const expireAt = new Date(Date.now() + SESSION_TIMEOUT_HOURS * 3600 * 1000);

      PropertiesService.getUserProperties().setProperty(
        'admin_session_' + sessionToken,
        String(expireAt.getTime())
      );

      return {
        ok: true,
        token: sessionToken,
        expiresAt: expireAt.toISOString()
      };
    }

    // 아이디/비번 틀린 경우
    return { ok: false, error: 'Invalid credentials' };

  } catch (e) {
    Logger.log('adminLogin error: ' + e);
    return { ok: false, error: String(e) };
  }
}

function adminVerifySession(token) {
  if (!token) return false;

  const props = PropertiesService.getUserProperties();
  const expireStr = props.getProperty('admin_session_' + token);

  if (!expireStr) return false;

  const expireAt = parseInt(expireStr, 10);
  if (Date.now() > expireAt) {
    props.deleteProperty('admin_session_' + token);
    return false;
  }

  return true;
}

function adminLogout(token) {
  if (token) {
    PropertiesService.getUserProperties().deleteProperty('admin_session_' + token);
  }
  return { ok: true };
}

/****************************** doGet 및 결과 페이지 ******************************/
function doGet(e) {
  e = e || {parameter:{}}; var p = e.parameter || {};

  // 인라인 결과 페이지(카카오 인앱 safeClose 포함)
  function createResultPage(title, message) {
    const html =
`<!doctype html><html lang="ko"><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${title}</title>
<style>
  body{font-family:system-ui,-apple-system,'Noto Sans KR',sans-serif;margin:0;padding:24px;color:#111}
  .card{max-width:560px;margin:0 auto;border:1px solid #eee;border-radius:12px;padding:24px}
  h1{font-size:18px;margin:0 0 12px}
  p{font-size:15px;line-height:1.5;margin:0 0 20px}
  button{width:100%;border:0;border-radius:10px;padding:14px 16px;font-size:16px}
  .ok{background:#111;color:#fff}
</style>
<div class="card">
  <h1>${title}</h1>
  <p>${message}</p>
  <button class="ok" onclick="safeClose()">확인</button>
</div>
<script>
function safeClose(){
  try{ location.href='kakaoweb://closeBrowser'; return; }catch(e){}
  try{ window.close(); }catch(e){}
  setTimeout(function(){
    if (history.length>1) { history.go(-2); return; }
    location.href='https://mamoru.kr';
  }, 150);
}
</script></html>`;
    return HtmlService.createHtmlOutput(html)
      .setTitle(title)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const action = p.a;
  const token = p.t;
  const uid = p.u;
  const datetime = p.d;

// 딜러용 일정 확정 페이지
  if (action === 'dealerConfirm' || p.action === 'dealerConfirm') {
    const t = HtmlService.createTemplateFromFile('_gas/DealerConfirm');
    t.uid = p.uid || '';
    t.dealerId = p.did || '';
    return t.evaluate()
      .setTitle('출장 일정 확정')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (action === 's' && token) {
    const data = getSuggestionDataByToken_(token);
    if (!data) return createResultPage('오류', '유효하지 않거나 만료된 링크입니다.');
    const t = HtmlService.createTemplateFromFile('_gas/Suggest');
    t.data = data;
    return t.evaluate().setTitle('마모루 상담 시간 선택');
  }

  if (action === 'c' && uid && datetime) {
    const result = confirmFieldRequest_(uid, datetime);
    if (result.ok) {
      return createResultPage('일정 확정 완료', '일정이 확정되었습니다. 곧 안내톡을 보내드릴게요.');
    } else {
      return createResultPage('오류', result.error);
    }
  }

      if (action === 'r' && token) {
    // 재요청 UI는 먼저 빠르게 보여주고,
    // 실제 HOLD 해제/상태 변경은 Reschedule.html에서 비동기로 호출
    const t = HtmlService.createTemplateFromFile('_gas/Reschedule');
    t.token   = token;
    t.baseUrl = ScriptApp.getService().getUrl();
    return t
      .evaluate()
      .setTitle('알림');
  }
    // 재요청(markResched): 토큰으로 UID 찾고, HOLD 해제 + 상태를 RESCHEDULE_REQUESTED로 변경
  if (p.action === 'markResched' && p.t) {
    const tok = String(p.t);
    const customerNote = String(p.reason || '').trim(); // 고객 재요청 사유
    const foundUid = getUidByToken_(tok);
    if (!foundUid) return json({ ok:false, error:'not_found' });

    try {
      clearHoldsForUid_(foundUid);
      updateStatus_(foundUid, 'RESCHEDULE_REQUESTED');
      
      // ★ 재요청 이메일 알림 발송
      try {
        const sh = sh_();
        const col = headerMap_();
        const rows = sh.getDataRange().getValues();
        
        for (let i = 1; i < rows.length; i++) {
          if (String(rows[i][col['UniqueID']]) === String(foundUid)) {
            const customerName = rows[i][col['성함']] || '';
            const customerPhone = rows[i][col['연락처']] || '';
            const consultType = rows[i][col['상담방식']] || '';
            const suggestions = rows[i][col['제안내용']] || '';
            
            // 주소 조합
            let addrText = '';
            if (col['addressRoad'] != null && rows[i][col['addressRoad']]) {
              addrText += String(rows[i][col['addressRoad']]).trim();
            }
            if (col['addressDetail'] != null && rows[i][col['addressDetail']]) {
              addrText += ' ' + String(rows[i][col['addressDetail']]).trim();
            }
            
            const emailSubject = '[MAMORU] 고객 시간 재요청 - ' + customerName;
            const emailBody = 
              '고객이 다른 시간을 요청했습니다.\n\n' +
              '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
              '■ 고객 정보\n' +
              '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
              '• 고객명: ' + customerName + '\n' +
              '• 연락처: ' + customerPhone + '\n' +
              '• 상담방식: ' + consultType + '\n' +
              (addrText ? '• 주소: ' + addrText + '\n' : '') +
              '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
              '■ 기존 제안 내용 (취소됨)\n' +
              '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
              (suggestions || '없음') + '\n' +
              '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
              (customerNote ? '\n■ 고객 요청사항\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' + customerNote + '\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' : '') +
              '\n→ 새로운 시간을 제안해주세요.\n' +
              '요청시각: ' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
            
            // 비고 컬럼에 고객 요청사항 append
            if (customerNote && col['비고'] != null) {
              const prevNote = String(rows[i][col['비고']] || '');
              const nowStr = Utilities.formatDate(new Date(), TIMEZONE, 'MM-dd HH:mm');
              const appendNote = '[고객 재요청 ' + nowStr + '] ' + customerNote;
              sh.getRange(i + 1, col['비고'] + 1).setValue(prevNote ? prevNote + '\n' + appendNote : appendNote);
            }

            GmailApp.sendEmail('bsm@mamoru.kr', emailSubject, emailBody);
            Logger.log('[EMAIL] 재요청 이메일 발송 완료: ' + customerName);
            break;
          }
        }
      } catch(emailErr) {
        Logger.log('[EMAIL ERROR] 재요청 이메일 발송 실패: ' + emailErr);
      }
      
      // TMS Supabase 동기화 — reschedule_requested 상태 반영
      try { pushToSupabase_(foundUid); } catch(e2) { Logger.log('[supabase-sync] markResched push 실패: ' + e2); }

      return json({ ok:true });
    } catch (e) {
      Logger.log('markResched fail: ' + e);
      return json({ ok:false, error:String(e) });
    }
  }




  if (p.action === 'getSlots' && p.date) return json({ok:true, slots:getAvailableTimeSlots_(String(p.date))});
  if (p.action === 'ver') return json({ok:true, version:VERSION, ts:new Date().toISOString()});

  // ─── TMS 취소 연동 (GET 쿼리 파라미터 방식 — POST body는 외부에서 수신 불가) ───
  if (p.action === 'cancelConsultation') {
    var cancelKey = p.key;
    if (!cancelKey || cancelKey !== (PropertiesService.getScriptProperties().getProperty('TMS_SYNC_KEY') || 'mamoru-tms-cron-2026')) {
      return json({ ok: false, error: 'unauthorized' });
    }
    try {
      var cancelResult = adminCancel(p.uid, { skipNotify: p.skipNotify === 'true' });
      return json(cancelResult);
    } catch (cancelErr) {
      Logger.log('cancelConsultation error: ' + cancelErr);
      return json({ ok: false, error: String(cancelErr) });
    }
  }

  // ─── TMS 시간제안 연동 — 캘린더 HOLD + 시트 상태 + 슬롯 차단 + 알림톡 ───
  if (p.action === 'suggestTimes') {
    var suggestKey = p.key;
    if (!suggestKey || suggestKey !== (PropertiesService.getScriptProperties().getProperty('TMS_SYNC_KEY') || 'mamoru-tms-cron-2026')) {
      return json({ ok: false, error: 'unauthorized' });
    }
    try {
      var suggestions = JSON.parse(p.suggestions || '[]');
      var suggestResult = adminSuggestTimes(p.uid, suggestions);
      return json(suggestResult);
    } catch (suggestErr) {
      Logger.log('suggestTimes error: ' + suggestErr);
      return json({ ok: false, error: String(suggestErr) });
    }
  }

  // ─── GitHub Pages 고객 대면 페이지용 API ───
  // 시간 선택 데이터 (page_suggest.html → fetch)
  if (p.action === 'getSuggestData' && p.t) {
    const suggestData = getSuggestionDataByToken_(String(p.t));
    if (!suggestData) return json({ ok: false, error: 'not_found' });
    return json({ ok: true, data: suggestData });
  }
  // 딜러 배정 목록 (page_dealer_confirm.html → fetch)
  if (p.action === 'getDealerData' && p.did) {
    try {
      const assignments = getDealerAssignments(String(p.did));
      return json({ ok: true, data: assignments });
    } catch (e) {
      return json({ ok: false, error: String(e) });
    }
  }
  // 딜러 일정 확정 (page_dealer_confirm.html → fetch)
  if (p.action === 'dealerConfirmSchedule' && p.uid && p.did) {
    try {
      const result = dealerConfirmSchedule(p.uid, p.did, p.date || '', p.time || '', p.memo || '');
      return json(result);
    } catch (e) {
      return json({ ok: false, error: String(e) });
    }
  }

  // ─── 고객 일정 변경/취소 요청 API (page_change_request.html) ───
  if (p.action === 'getReservationInfo' && p.uid) {
    try {
      var resInfo = getReservationByUid_(String(p.uid));
      if (!resInfo) return json({ ok: false, error: '예약 정보를 찾을 수 없습니다.' });
      if (!['CONFIRMED', 'ASSIGNED'].includes(String(resInfo.status).toUpperCase())) {
        return json({ ok: false, error: '변경/취소 요청이 불가능한 상태입니다.' });
      }
      return json({ ok: true, data: { name: resInfo.name, type: resInfo.type, date: resInfo.date, time: resInfo.time, status: resInfo.status } });
    } catch (e) {
      return json({ ok: false, error: String(e) });
    }
  }
  if (p.action === 'submitChangeRequest' && p.uid) {
    try {
      var changeResult = submitChangeRequest_(String(p.uid), String(p.reqType || ''), String(p.reason || ''), String(p.memo || ''), String(p.hopeDate || ''));
      return json(changeResult);
    } catch (e) {
      return json({ ok: false, error: String(e) });
    }
  }

  // ─── GitHub Pages용 API 분기 (CORS 지원) ───
  if (p.action === 'getSettings') {
    return json({ ok: true, data: getSettings() });
  }
  if (p.action === 'getSlotsRange' && p.dates) {
    const dateList = String(p.dates).split(',').map(d => d.trim());
    const consultType = p.type || '매장 방문';
    return json({ ok: true, data: getSlotsRange(dateList, consultType) });
  }
  if (p.action === 'confirm' && p.token) {
    var r = confirmByToken_(String(p.token));
    var msg = r.ok ? '예약이 확정되었습니다.' : '확정 토큰이 유효하지 않습니다.';
    return createResultPage(r.ok ? '예약 확정' : '오류', msg);
  }
  // 관리자 로그인 처리
  if (p.action === 'adminLogin') {
    const result = adminLogin(p.username, p.password);
    return json(result);
  }
  
  // 관리자 로그아웃 처리
  if (p.action === 'adminLogout') {
    const result = adminLogout(p.token);
    return json(result);
  }
  
  // 관리자 세션 검증
  if (p.action === 'adminVerify') {
    const valid = adminVerifySession(p.token);
    return json({ ok: valid });
  }
  
  // 관리자 페이지 — TMS로 이전 완료, GAS 관리 페이지 미사용
 if (p.action === 'admin') {
    return createResultPage('안내', '관리자 페이지는 TMS로 이전되었습니다.');
  }
  if (p.action === 'adminList') {
    if (!adminVerifySession(p.token)) {
      return json({ ok: false, error: 'session_expired' });
    }
    return json(adminList());
  }
      // (기본) 고객용 상담 접수 페이지 — GitHub Pages로 이전 완료
  return createResultPage('안내', '상담 접수 페이지가 이전되었습니다. mamoru.kr을 방문해주세요.');


}

/****************************** doPost - GitHub Pages용 POST API ******************************/
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const action = payload.action;

    if (action === 'submitConsultation') {
      const result = submitConsultation(payload.data);
      return json({ ok: true, data: result });
    }

    // TMS에서 취소 시 호출 — 캘린더 삭제 + 시트 상태 + 슬롯 캐시 무효화
    if (action === 'cancelConsultation') {
      const key = payload.key;
      if (!key || key !== (PropertiesService.getScriptProperties().getProperty('TMS_SYNC_KEY') || 'mamoru-tms-cron-2026')) {
        return json({ ok: false, error: 'unauthorized' });
      }
      const result = adminCancel(payload.uid, { skipNotify: !!payload.skipNotify });
      return json(result);
    }

    if (action === 'getSlotsRange') {
      const dateList = payload.dates || [];
      const consultType = payload.type || '매장 방문';
      return json({ ok: true, data: getSlotsRange(dateList, consultType) });
    }

    return json({ ok: false, error: 'unknown_action' });
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return json({ ok: false, error: String(err) });
  }
}

/****************************** Make Webhook ******************************/
function postMake_(event, payload){
  const corrId = Utilities.getUuid();
  const _uid = String((payload && (payload.id || payload.uid)) || Utilities.getUuid());
  const _tpl = String((payload && payload.template) || 'na');
  const _dt  = String((payload && (payload.date || payload.new_date || payload.old_date)) || 'na');
  const _tm  = String((payload && (payload.time || payload.new_time || payload.old_time)) || 'na');
  const idemKey = _uid + ':' + _tpl + ':' + _dt + 'T' + _tm;
  const headers = { 'Accept': 'application/json', 'Content-Type': 'application/json', 'X-Correlation-Id': corrId, 'X-Idempotency-Key': idemKey };
  const meta = { ts: new Date().toISOString(), version: VERSION, func: event, trigger: (payload && (payload.trigger || payload._trigger)) || 'time' };
  const bodyStr = JSON.stringify({ _meta: meta, ...payload });
  const shouldRetry = (code) => code === 429 || (code >= 500 && code < 600);
  let lastStatus = 0, lastBody = '';
  for (let attempt = 1; attempt <= 3; attempt++){
    try{
      const resp = UrlFetchApp.fetch(MAKE_WEBHOOK_URL, {
        method: 'post', contentType: 'application/json', headers, payload: bodyStr,
        muteHttpExceptions: true, followRedirects: true, readTimeoutMillis: 5000
      });
      lastStatus = resp.getResponseCode();
      lastBody   = resp.getContentText();
      Logger.log(JSON.stringify({ tag:'MAKE', event, status:lastStatus, attempt, corrId, idemKey }));
      if (lastStatus >= 200 && lastStatus < 300) return { ok:true, status:lastStatus, body:lastBody };
      if (!shouldRetry(lastStatus)) throw new Error('HTTP ' + lastStatus + ' ' + lastBody);
    } catch(e){
      if (attempt < 3) {
        Utilities.sleep(500 * Math.pow(2, attempt - 1) + Math.floor(Math.random()*250));
        continue;
      }
      throw new Error((lastStatus?('HTTP '+lastStatus+' '):'') + (lastBody || String(e)));
    }
  }
  return { ok:false, status:lastStatus, body:lastBody };
}

/*** 주소 → 좌표 보강 ***/
function geocodeAddress_(query){
  try{
    if (!query) return { placeId:'', lat:'', lng:'' };
    const res = Maps.newGeocoder().geocode(query);
    if (res && res.status === 'OK' && res.results && res.results.length){
      const r = res.results[0];
      const loc = (r.geometry && r.geometry.location) || {};
      return { placeId: String(r.place_id || ''), lat: String(loc.lat || ''), lng: String(loc.lng || '') };
    }
  } catch(e){ Logger.log('geocode fail: ' + e); }
  return { placeId:'', lat:'', lng:'' };
}
function BACKFILL_GEO_(){
  const sh = sh_(), col = headerMap_();
  const rows = sh.getDataRange().getValues();
  let updated = 0, skipped = 0, failed = 0;
  for (let r = 2; r <= rows.length; r++){
    const road   = String(rows[r-1][col['addressRoad']]||'').trim();
    const detail = String(rows[r-1][col['addressDetail']]||'').trim();
    const hasLat = String(rows[r-1][col['addressLat']]||'').trim();
    const hasLng = String(rows[r-1][col['addressLng']]||'').trim();
    const hasPid = String(rows[r-1][col['addressPlaceId']]||'').trim();
    if (!road) { skipped++; continue; }
    if (hasLat && hasLng && hasPid) { skipped++; continue; }
    const g = geocodeAddress_(road + (detail ? (' ' + detail) : ''));
    if (g.lat || g.lng || g.placeId){
      if (col['addressPlaceId'] != null) sh.getRange(r, col['addressPlaceId']+1).setValue(g.placeId||'');
      if (col['addressLat']     != null) sh.getRange(r, col['addressLat']+1).setValue(g.lat||'');
      if (col['addressLng']     != null) sh.getRange(r, col['addressLng']+1).setValue(g.lng||'');
      updated++;
    } else { failed++; }
    Utilities.sleep(150);
  }
  Logger.log(JSON.stringify({updated, skipped, failed}));
  return {updated, skipped, failed};
}

/****************************** 제출 처리 ******************************/
function submitConsultation(form){
  const now       = new Date();
  const name      = String(form.name || '').trim();
  const phone     = String(form.phone || '').replace(/\D/g,'');
  const type      = String(form.type || '').trim();
  const visitDate = String(form.visitDate || '').trim();
  const visitTime = String(form.visitTime || '').trim();
  const address   = String(form.address || '').trim();

  const addressZip     = String(form.addressZip||'').trim();
  const addressRoad    = String(form.addressRoad||'').trim();
  const addressDetail  = String(form.addressDetail||'').trim();
  let   addressPlaceId = String(form.addressPlaceId||'').trim();
  let   addressLat     = String(form.addressLat||'').trim();
  let   addressLng     = String(form.addressLng||'').trim();

  if (addressRoad && (!addressLat || !addressLng || !addressPlaceId)) {
    const g = geocodeAddress_(addressRoad + (addressDetail ? (' ' + addressDetail) : ''));
    if (!addressPlaceId && g.placeId) addressPlaceId = g.placeId;
    if (!addressLat && g.lat) addressLat = g.lat;
    if (!addressLng && g.lng) addressLng = g.lng;
  }
  Logger.log(JSON.stringify({tag:'GEO_FALLBACK', addressRoad, addressDetail, addressPlaceId, addressLat, addressLng}));

  const daysArr   = Array.isArray(form.days) ? form.days : [];
  const timePrefsArr = Array.isArray(form.timePrefs) ? form.timePrefs : [];
  const memo      = String(form.memo || '').trim();
  if (!name || !phone) throw new Error('성함/연락처는 필수입니다.');

  if (type === '매장 방문' && visitDate && visitTime) {
    const sh  = sh_();
    const col = headerMap_();
    const rows = sh.getDataRange().getValues();
    const dup = rows.slice(1).some(r => {
      const rPhone = String(r[col['연락처']]||'').replace(/\D/g,'');
      return rPhone===phone && String(r[col['상담방식']])==='매장 방문' && String(r[col['방문일']])===visitDate && String(r[col['방문시간']])===visitTime && (String(r[col['Status']])==='PENDING_ADMIN' || String(r[col['Status']])==='CONFIRMED');
    });
    if (dup) throw new Error('같은 날짜/시간에 이미 접수된 내역이 있습니다. 다른 시간을 선택해주세요.');
    const { start } = makeDateTimes_(visitDate, visitTime);
    if (start.getTime() < Date.now()) throw new Error('과거 시간으로는 예약할 수 없습니다.');
    if (getClosedWeekdays_().includes(start.getDay())) throw new Error('휴무 요일에는 예약할 수 없습니다.');
    const ymd = Utilities.formatDate(start, TIMEZONE, 'yyyy-MM-dd');
    if (getClosedDates_().includes(ymd)) throw new Error('임시휴무일에는 예약할 수 없습니다.');
    if (!getAvailableTimeSlots_(visitDate).includes(visitTime)) throw new Error('선택하신 시간이 방금 마감되었습니다. 다른 시간을 선택해주세요.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const uid   = Utilities.getUuid();
    const stat  = (type === '매장 방문' && visitDate && visitTime)
  ? 'CONFIRMED'
  : 'PENDING_ADMIN';
    const token = Utilities.getUuid();
    const sh = sh_();
    const col = headerMap_();

    const newRow = new Array(sh.getLastColumn()).fill('');
    newRow[col['접수시각']]        = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    newRow[col['성함']]            = name;
    newRow[col['연락처']]          = "'" + phone;
    newRow[col['상담방식']]        = type;
    newRow[col['방문일']]          = visitDate;
    newRow[col['방문시간']]        = visitTime;
    newRow[col['주소']]            = address;

    if (col['addressZip']     != null) newRow[col['addressZip']]     = addressZip;
    if (col['addressRoad']    != null) newRow[col['addressRoad']]    = addressRoad;
    if (col['addressDetail']  != null) newRow[col['addressDetail']]  = addressDetail;
    if (col['addressPlaceId'] != null) newRow[col['addressPlaceId']] = addressPlaceId;
    if (col['addressLat']     != null) newRow[col['addressLat']]     = addressLat;
    if (col['addressLng']     != null) newRow[col['addressLng']]     = addressLng;

    newRow[col['가능요일']]        = daysArr.join(',');
    newRow[col['선호시간대']]      = timePrefsArr.join(',');
    newRow[col['비고']]            = memo;
    newRow[col['원본']]            = JSON.stringify(form);
    newRow[col['UniqueID']]        = uid;
    newRow[col['Status']]          = stat;
    newRow[col['ConfirmationToken']]= token;

    sh.appendRow(newRow);
    const row = sh.getLastRow();

    if (type === '매장 방문' && visitDate && visitTime) {
  try {
    const { start, end } = makeDateTimes_(visitDate, visitTime);
    if (hasTimeConflict_(start, end, null)) {
      throw new Error('선택하신 시간이 방문 마감되었습니다. 다른 시간을 선택해주세요.');
    }
    const cal = getCal_(CAL_STORE_NAME);
    const ev  = cal.createEvent(
      `[방문] ${name||''}`,
      start,
      end,
      { description: `고객:${name||''} / ${phone||''} / 확정됨` }
    );
    sh.getRange(row, col['CalendarEventID']+1).setValue(ev.getId());
    if (visitDate) slotsCacheInvalidate_(visitDate);
  } catch(e){
    Logger.log('STORE CONFIRM fail: ' + e);
  }
}


    if (type === '매장 방문' && visitDate && visitTime) {
  const payload = {
    topic:    'alrimtalk',
    template: 'confirmed',       // 방문 확정용 템플릿 이름 (Make/솔라피에서 이 이름에 맞춰 템플릿 연결)
    event:    'CONFIRMED',
    id:       uid,
    status:   stat,
    token,
    name,
    phone,
    type,
    date:     visitDate,
    time:     visitTime,
    address,
    days:     daysArr,
    memo,
    change_request_link: GITHUB_PAGES_CONSULT + '/page_change_request.html?uid=' + uid  // 일정 변경/취소 요청 링크
  };
  try {
    postMake_('CONFIRMED', payload);
  } catch(e) {
    Logger.log('Make post error(CONFIRMED): ' + e);
  }
  } else {
    // 희망 요일 / 희망 시간대 배열
    const hopeDaysArr  = daysArr;
    const hopeTimesArr = timePrefsArr;

    const payload = {
      topic:    'alrimtalk',
      template: 'request',
      event:    'CONSULT_REQUEST',
      id:       uid,
      status:   stat,
      token,
      name,
      phone,
      type,
      date:     visitDate,
      time:     visitTime,
      address,

      // 기존 배열 그대로 (필요시 시나리오에서 배열로도 사용 가능)
      days:      hopeDaysArr,
      timePrefs: hopeTimesArr,

      // 알림톡 템플릿에 바로 쓰기 좋은 문자열 버전
      hope_days_text:  hopeDaysArr.join(', '),
      hope_times_text: hopeTimesArr.join(', '),

      memo
    };

    try {
      postMake_('CONSULT_REQUEST', payload);
    } catch(e) {
      Logger.log('Make post error(CONSULT_REQUEST): ' + e);
    }
  }

  // ★ 신규 접수 이메일 알림 발송
  try {
    const emailSubject = '[MAMORU 상담] 새로운 ' + type + ' 접수';
    const emailBody = 
      '새로운 상담 접수가 들어왔습니다.\n\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      '■ 접수 정보\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      '• 상담방식: ' + type + '\n' +
      '• 고객명: ' + name + '\n' +
      '• 연락처: ' + phone + '\n' +
      (addressRoad ? '• 주소: ' + addressRoad + ' ' + addressDetail + '\n' : '') +
      (visitDate ? '• 방문일: ' + visitDate + '\n' : '') +
      (visitTime ? '• 방문시간: ' + visitTime + '\n' : '') +
      (daysArr.length ? '• 희망요일: ' + daysArr.join(', ') + '\n' : '') +
      (timePrefsArr.length ? '• 희망시간대: ' + timePrefsArr.join(', ') + '\n' : '') +
      (memo ? '• 메모: ' + memo + '\n' : '') +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      '접수시각: ' + Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    
    GmailApp.sendEmail('bsm@mamoru.kr', emailSubject, emailBody);
    Logger.log('[EMAIL] 신규접수 이메일 발송 완료: ' + name);
  } catch(emailErr) {
    Logger.log('[EMAIL ERROR] 신규접수 이메일 발송 실패: ' + emailErr);
  }

  // ★ Supabase TMS 동기화
  try { pushToSupabase_(uid); } catch(e) { Logger.log('[supabase-sync] push 실패: ' + e); }

  } finally {
    lock.releaseLock();
  }
  return { ok:true, message:'신청이 완료되었습니다. 곧 안내톡을 보내드릴게요.' };
}

/****************************** 확정 처리 ******************************/
// ===== PATCH: 매장 방문 확정(토큰 기반) 안전 버전 =====
function confirmByToken_(token){
  const sh  = sh_();
  const col = headerMap_();
  const data = sh.getDataRange().getValues();

  // 1) 토큰으로 행 찾기
  let row = -1;
  for (let i = 1; i < data.length; i++){
    if (String(data[i][col['ConfirmationToken']]) === String(token)) {
      row = i + 1;
      break;
    }
  }
  if (row < 0) {
    return { ok:false, error:'invalid_token' };
  }

  const r         = data[row-1];
  const type      = r[col['상담방식']];
  const name      = r[col['성함']];
  const phone     = r[col['연락처']];
  const visitDate = r[col['방문일']];
  const visitTime = r[col['방문시간']];
  const evIdPrev  = r[col['CalendarEventID']] || '';

  // 2) 시트 상태 업데이트
  sh.getRange(row, col['Status']+1).setValue('CONFIRMED');
  sh.getRange(row, col['ConfirmationToken']+1).setValue('');

  // 3) 캘린더 이벤트 생성/수정 (매장 캘린더 기준)
  let evId = evIdPrev;
  try {
    const cal = getCal_(CAL_STORE_NAME);

    if (evIdPrev) {
      // 기존 이벤트가 있으면 제목/설명만 정리
      const ev = cal.getEventById(evIdPrev);
      if (ev) {
        ev.setTitle(`[방문] ${name || ''}`);
        ev.setDescription(`고객:${name||''} / ${phone||''} / 확정됨`);
      }
    } else if (String(type) === '매장 방문' && visitDate) {
      // 이벤트가 없으면 새로 생성
      const { start, end } = makeDateTimes_(visitDate, visitTime);
      const ev = cal.createEvent(
        `[방문] ${name || ''}`,
        start,
        end,
        { description: `고객:${name||''} / ${phone||''} / 토큰:${token}` }
      );
      evId = ev.getId();
      sh.getRange(row, col['CalendarEventID']+1).setValue(evId);
    }
  } catch (e) {
    Logger.log('Confirm calendar op fail: ' + e);
  }

  // 4) 슬롯 캐시 무효화
  try{
    const ymd = visitDate instanceof Date
      ? Utilities.formatDate(visitDate, TIMEZONE, 'yyyy-MM-dd')
      : String(visitDate || '');
    if (ymd) slotsCacheInvalidate_(ymd);
  } catch(_){}

  // 5) 알림톡 — "확정" 템플릿 호출
  try {
    const uid = r[col['UniqueID']];

    const dtStr = visitDate instanceof Date
      ? Utilities.formatDate(visitDate, TIMEZONE, 'yyyy-MM-dd')
      : String(visitDate || '');
    const tmStr = visitTime instanceof Date
      ? Utilities.formatDate(visitTime, TIMEZONE, 'HH:mm')
      : String(visitTime || '');

    // 방문 주소 텍스트 (도로명+상세 우선, 없으면 "주소" 컬럼)
    let addrText = '';
    if (col['addressRoad'] != null && r[col['addressRoad']]) {
      addrText += String(r[col['addressRoad']]).trim();
    }
    if (col['addressDetail'] != null && r[col['addressDetail']]) {
      addrText += (addrText ? ' ' : '') + String(r[col['addressDetail']]).trim();
    }
    if (!addrText && col['주소'] != null) {
      addrText = String(r[col['주소']] || '').trim();
    }

    const payload = {
      topic:    'alrimtalk',
      template: 'confirmed',           // ✔ 확정 템플릿
      event:    'CONFIRMED_BY_TOKEN',  // 필요시 Make 시나리오에서 분기
      id:       uid,
      name,
      phone,
      type,
      date:     dtStr,
      time:     tmStr,
      address:  addrText || '',
      channel:  'kakao',
      sms_fallback: false,
      change_request_link: GITHUB_PAGES_CONSULT + '/page_change_request.html?uid=' + uid  // 일정 변경/취소 요청 링크
    };
    postMake_('CONFIRMED', payload);
  } catch (e) {
    Logger.log('Make post error(CONFIRMED): ' + e);
  }

  return { ok:true };
}

// ===== PATCH 2 END =====


/****************************** 관리자 API ******************************/
function adminList(){
  try {
    // [Updated] 에러 처리 추가
    const sh = sh_(), col = headerMap_(), rows = sh.getDataRange().getValues();
    
    // 데이터 검증
    if (!rows || rows.length < 2) {
      Logger.log('adminList: No data rows found');
      return [];
    }
    
    const tz = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const toDateStr = v => !v ? '' : (v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v).slice(0,10));
    const toTimeStr = v => !v ? '' : (v instanceof Date ? Utilities.formatDate(v, tz, 'HH:mm') : String(v));
    const toCreated = v => !v ? '' : (v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm:ss') : String(v));
    const toCreatedMs = v => {
      if (!v) return 0;
      if (v instanceof Date) return v.getTime();
      const s = String(v);
      const m = /^(\d{4}-\d{2}-\d{2})(?:[ T](\d{2}:\d{2}(?::\d{2})?))?/.exec(s);
      return m ? new Date(m[1] + 'T' + (m[2] || '00:00')).getTime() : new Date(s).getTime() || 0;
    };

    const out = [];
    for (let i=1;i<rows.length;i++){
      const r = rows[i];
      const dt = toDateStr(r[col['방문일']]);
      const type = r[col['상담방식']];

      // 주소지(가능하면 도로명+상세, 없으면 기존 "주소" 컬럼)
      const addressParts = [];
      if (col['addressRoad'] != null && r[col['addressRoad']]) {
        addressParts.push(r[col['addressRoad']]);
      }
      if (col['addressDetail'] != null && r[col['addressDetail']]) {
        addressParts.push(r[col['addressDetail']]);
      }
      let address = addressParts.join(' ');
      if (!address && col['주소'] != null) {
        address = r[col['주소']] || '';
      }

      if (type === '매장 방문' && dt && dt < todayStr) continue;
      const created = r[col['접수시각']];
      out.push({
        row: i+1,
        uid: r[col['UniqueID']],
        status: r[col['Status']],
        name: r[col['성함']],
        phone: r[col['연락처']],
        address: address,
        type: type,
        date: dt,
        time: toTimeStr(r[col['방문시간']]),
        days: r[col['가능요일']] || '',
        timePrefs: r[col['선호시간대']] || '',
        eventId: r[col['CalendarEventID']] || '',
        createdAt: toCreated(created),
        _createdMs: toCreatedMs(created)
      });
    }

    return out.sort((a,b)=>{
      const aHas = !!a.date, bHas = !!b.date;
      if (aHas && bHas){
        const aMs = new Date(a.date+'T'+(a.time||'00:00')).getTime();
        const bMs = new Date(b.date+'T'+(b.time||'00:00')).getTime();
        return aMs - bMs;
      }
      if (!aHas && !bHas){
        return b._createdMs - a._createdMs;
      }
      return aHas ? -1 : 1;
    });
    
  } catch(e) {
    Logger.log('adminList error: ' + e);
    throw new Error('데이터 로딩 중 오류가 발생했습니다: ' + e.message);
  }
}
function adminConfirm(uid){
  const sh=sh_(), col=headerMap_(), rows=sh.getDataRange().getValues();
  let row=-1; for(let i=1;i<rows.length;i++){ if(String(rows[i][col['UniqueID']])===String(uid)){ row=i+1; break; } }
  if(row<0) throw new Error('not found');
  const curTok=rows[row-1][col['ConfirmationToken']]||Utilities.getUuid();
  if(!rows[row-1][col['ConfirmationToken']]) sh.getRange(row,col['ConfirmationToken']+1).setValue(curTok);
  return confirmByToken_(curTok);
}

// ===== PATCH 1: adminCancel (매장/출장/홀드 모두 삭제) =====
function adminCancel(uid, options){
  options = options || {};
  const sh = sh_(), col = headerMap_(), rows = sh.getDataRange().getValues();
  let row = -1;
  for (let i=1;i<rows.length;i++){
    if (String(rows[i][col['UniqueID']]) === String(uid)) { row = i+1; break; }
  }
  if (row < 0) throw new Error('not found');

  const r = rows[row-1];
  const evId = String(r[col['CalendarEventID']] || '');

  // 1) 본 이벤트 삭제: 매장/출장 양쪽에서 시도
  if (evId){
    try{ const ev = getCal_(CAL_STORE_NAME).getEventById(evId); if (ev) ev.deleteEvent(); }catch(_){}
    try{ const ev = getCal_(CAL_FIELD_NAME).getEventById(evId); if (ev) ev.deleteEvent(); }catch(_){}
  }

  // 2) HOLD 이벤트 삭제 + 시트 초기화 (모든 캘린더에서)
  try{
    if (col['holdEventIds']!=null){
      let ids=[]; try{ ids = JSON.parse(String(r[col['holdEventIds']]||'[]')); }catch(_){}
      
      // 매장 캘린더에서 삭제
      try {
        const storeCal = getCal_(CAL_STORE_NAME);
        ids.forEach(id=>{
          try{ const hev = storeCal.getEventById(id); if (hev) hev.deleteEvent(); }catch(_){}
        });
      } catch(_){}
      
      // 출장 캘린더에서 삭제
      try {
        const fieldCal = getCal_(CAL_FIELD_NAME);
        ids.forEach(id=>{
          try{ const hev = fieldCal.getEventById(id); if (hev) hev.deleteEvent(); }catch(_){}
        });
      } catch(_){}
      
      // 기본 캘린더에서도 삭제
      try {
        const baseCals = CalendarApp.getCalendarsByName(CAL_NAME);
        if (baseCals && baseCals.length > 0) {
          ids.forEach(id=>{
            try{ const hev = baseCals[0].getEventById(id); if (hev) hev.deleteEvent(); }catch(_){}
          });
        }
      } catch(_){}
      
      sh.getRange(row, col['holdEventIds']+1).setValue('[]');
    }
    if (col['holdExpireAt']!=null) sh.getRange(row, col['holdExpireAt']+1).setValue('');
  }catch(_){}

  // 3) 시트 상태 정리
  sh.getRange(row, col['Status']+1).setValue('CANCELLED');
  sh.getRange(row, col['CalendarEventID']+1).setValue('');
  sh.getRange(row, col['ConfirmationToken']+1).setValue('');

  // 4) 슬롯 캐시 무효화
  try{
    const dateStr = r[col['방문일']] instanceof Date
      ? Utilities.formatDate(r[col['방문일']], TIMEZONE, 'yyyy-MM-dd')
      : String(r[col['방문일']]||'');
    if (dateStr) slotsCacheInvalidate_(dateStr);
  }catch(_){}

  // 5) 알림톡 (TMS에서 호출 시 skipNotify=true로 중복 방지)
  if (!options.skipNotify) {
    try {
      const visitDate = r[col['방문일']] instanceof Date ? Utilities.formatDate(r[col['방문일']], TIMEZONE, 'yyyy-MM-dd') : r[col['방문일']];
      const visitTime = r[col['방문시간']] instanceof Date ? Utilities.formatDate(r[col['방문시간']], TIMEZONE, 'HH:mm') : r[col['방문시간']];
      const payload3 = { topic: 'alrimtalk', template: 'cancelled', id: r[col['UniqueID']], name: r[col['성함']], phone: r[col['연락처']], type: r[col['상담방식']], date: visitDate, time: visitTime };
      postMake_('CANCELLED', payload3);
    } catch (_){}
  }

  return { ok:true };
}
function adminChatSetStatus(uid, status){
  // 기존에 쓰고 있는 상태 업데이트 헬퍼 활용
  // 예: updateStatus_(uid, status);
  updateStatus_(uid, status);
  return true;
}

// ===== PATCH 1 END =====



// ===== PATCH 3: adminReschedule (매장 캘린더 기준) =====
function adminReschedule(uid, newDate, newTime){
  if (!/^\d{4}-\d{2}-\d{2}$/.test(newDate)) throw new Error('날짜 형식이 아닙니다(YYYY-MM-DD).');
  if (!/^\d{2}:\d{2}$/.test(newTime))       throw new Error('시간 형식이 아닙니다(HH:mm).');

  const sh  = sh_();
  const col = headerMap_();
  const rows = sh.getDataRange().getValues();

  // 해당 UID 행 찾기
  let row = -1;
  for (let i = 1; i < rows.length; i++){
    if (String(rows[i][col['UniqueID']]) === String(uid)) {
      row = i + 1;
      break;
    }
  }
  if (row < 0) throw new Error('not found');

  const r = rows[row - 1];
  const name  = r[col['성함']]        || '';
  const phone = r[col['연락처']]      || '';
  const type  = r[col['상담방식']]    || '';

  // 기존 날짜/시간(있으면) 문자열로 정리
  const oldDate = r[col['방문일']] instanceof Date
    ? Utilities.formatDate(r[col['방문일']], TIMEZONE, 'yyyy-MM-dd')
    : (r[col['방문일']] || '');
  const oldTime = r[col['방문시간']] instanceof Date
    ? Utilities.formatDate(r[col['방문시간']], TIMEZONE, 'HH:mm')
    : (r[col['방문시간']] || '');

  const evId = String(r[col['CalendarEventID']] || '');

  // ▶ 상담방식에 따라 사용할 캘린더 선택
  const isField = String(type).trim() === '출장 요청';
  const cal = getCal_(isField ? CAL_FIELD_NAME : CAL_STORE_NAME);

  // 새 시작/종료 시각 계산
  const { start, end } = makeDateTimes_(newDate, newTime);
  const nowSvr = new Date();

  // 휴무 / 임시휴무 체크
  const CLOSED_WEEKDAYS_ = getClosedWeekdays_();
  const CLOSED_DATES_ = (() => {
    const a = getClosedDates_();
    const b = getClosedDatesFromSheet_();
    return Array.from(new Set([...(a || []), ...(b || [])]));
  })();

  if (start.getTime() < nowSvr.getTime()) {
    throw new Error('과거 시간으로 변경할 수 없습니다.');
  }
  if (CLOSED_WEEKDAYS_.includes(start.getDay())) {
    throw new Error('휴무 요일에는 변경할 수 없습니다.');
  }
  const ymdNew = Utilities.formatDate(start, TIMEZONE, 'yyyy-MM-dd');
  if (CLOSED_DATES_.includes(ymdNew)) {
    throw new Error('임시휴무일에는 변경할 수 없습니다.');
  }

  // 다른 예약과 겹치는지 (매장+출장 전체 캘린더 기준)
  if (hasTimeConflict_(start, end, evId || null)) {
    throw new Error('해당 시간대에 이미 다른 예약이 있습니다.');
  }

  // ▶ 기존 이벤트가 있으면 시간만 옮기고, 없으면 새로 생성
  if (evId) {
    const ev = cal.getEventById(evId);
    if (!ev) throw new Error('calendar event missing');
    ev.setTime(start, end);
  } else {
    const titlePrefix = isField ? '[출장]' : '[방문]';
    const ev = cal.createEvent(`${titlePrefix} ${name || ''}`, start, end, {});
    sh.getRange(row, col['CalendarEventID'] + 1).setValue(ev.getId());
  }

  // 시트의 방문일/시간 + 리마인드 플래그 초기화
  sh.getRange(row, col['방문일']    + 1).setValue(newDate);
  sh.getRange(row, col['방문시간']  + 1).setValue(newTime);
  if (col['Remind24'] != null) sh.getRange(row, col['Remind24'] + 1).setValue('');
  if (col['Remind2']  != null) sh.getRange(row, col['Remind2']  + 1).setValue('');

  // 슬롯 캐시 무효화
  try{
    if (oldDate) slotsCacheInvalidate_(oldDate);
    if (newDate) slotsCacheInvalidate_(newDate);
  }catch(_){}

  // ▼ 방문 주소 텍스트 (도로명+상세 우선, 없으면 "주소" 컬럼 사용)
  let addrText = '';
  if (col['addressRoad'] != null && r[col['addressRoad']]) {
    addrText += String(r[col['addressRoad']]).trim();
  }
  if (col['addressDetail'] != null && r[col['addressDetail']]) {
    addrText += (addrText ? ' ' : '') + String(r[col['addressDetail']]).trim();
  }
  if (!addrText && col['주소'] != null) {
    addrText = String(r[col['주소']] || '').trim();
  }

  // Make 로 시간 변경 결과 전달 (주소 포함)
  const payload = {
    topic:    'alrimtalk',
    template: 'rescheduled',      // 나중에 솔라피에서 이 템플릿명으로 라우팅
    id:       r[col['UniqueID']],
    name,
    phone,
    type,               // '매장 방문' / '출장 요청' 구분에 사용 가능
    old_date: oldDate,
    old_time: oldTime,
    new_date: newDate,
    new_time: newTime,
    address:  addrText, // ★ 방문 주소
    channel:  'kakao',
    sms_fallback: false,
    change_request_link: GITHUB_PAGES_CONSULT + '/page_change_request.html?uid=' + r[col['UniqueID']]  // 일정 변경/취소 요청 링크
  };
  postMake_('RESCHEDULED', payload);  // _meta.func = 'RESCHEDULED'

  return { ok:true };
}


/* ───────────────────────────────
   딜러 관련 함수
─────────────────────────────── */

// 딜러 목록 가져오기
function getDealers() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(DEALER_SHEET_NAME);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  const headers = data[0];
  const dealers = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dealer = {};
    headers.forEach((h, idx) => { dealer[h] = row[idx]; });
    
    if (dealer.status === 'active') {
      dealers.push(dealer);
    }
  }
  
  return dealers;
}

// 지역에 맞는 딜러 목록
function getDealersByRegion(region) {
  const dealers = getDealers();
  if (!region) return dealers;
  
  return dealers.filter(d => {
    const regions = String(d.regions || '').split(',').map(r => r.trim());
    return regions.some(r => region.includes(r) || r.includes(region));
  });
}

// 관리자용: 딜러 목록 + 현재 배정 건수
function adminGetDealersWithStats() {
  const dealers = getDealers();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const dealerIdIdx = headers.indexOf('dealerId');
  const statusIdx = headers.indexOf('Status');
  
  // 딜러별 배정 건수 계산
  const counts = {};
  for (let i = 1; i < data.length; i++) {
    const did = data[i][dealerIdIdx];
    const st = data[i][statusIdx];
    if (did && st !== 'CANCELLED') {
      counts[did] = (counts[did] || 0) + 1;
    }
  }
  
  return dealers.map(d => ({
    ...d,
    activeCount: counts[d.dealerId] || 0
  }));
}

// 관리자: 딜러 배정
function adminAssignDealer(uid, dealerId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 예약 찾기
  const uidIdx = headers.indexOf('UniqueID');
  let rowIndex = -1;
  let rowData = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][uidIdx] === uid) {
      rowIndex = i + 1;
      rowData = {};
      headers.forEach((h, idx) => { rowData[h] = data[i][idx]; });
      break;
    }
  }
  
  if (rowIndex < 0) throw new Error('예약을 찾을 수 없습니다.');
  
  // 딜러 정보 가져오기
  const dealers = getDealers();
  const dealer = dealers.find(d => d.dealerId === dealerId);
  if (!dealer) throw new Error('딜러를 찾을 수 없습니다.');
  
  // 컬럼 인덱스
  const dealerIdColIdx = headers.indexOf('dealerId');
  const dealerNameColIdx = headers.indexOf('dealerName');
  const assignedAtColIdx = headers.indexOf('assignedAt');
  const statusColIdx = headers.indexOf('Status');
  
  // 시트 업데이트
  const now = new Date();
  const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  if (dealerIdColIdx >= 0) sheet.getRange(rowIndex, dealerIdColIdx + 1).setValue(dealerId);
  if (dealerNameColIdx >= 0) sheet.getRange(rowIndex, dealerNameColIdx + 1).setValue(dealer.name);
  if (assignedAtColIdx >= 0) sheet.getRange(rowIndex, assignedAtColIdx + 1).setValue(nowStr);
  if (statusColIdx >= 0) sheet.getRange(rowIndex, statusColIdx + 1).setValue('ASSIGNED');
  
  // 주소 조합
  let addrText = '';
  const addrRoadIdx = headers.indexOf('addressRoad');
  const addrDetailIdx = headers.indexOf('addressDetail');
  const addrIdx = headers.indexOf('주소');
  if (addrRoadIdx >= 0 && rowData['addressRoad']) addrText += String(rowData['addressRoad']).trim();
  if (addrDetailIdx >= 0 && rowData['addressDetail']) addrText += ' ' + String(rowData['addressDetail']).trim();
  if (!addrText.trim() && addrIdx >= 0) addrText = String(rowData['주소'] || '').trim();
  
  // 딜러에게 카톡 알림 발송 (Make 웹훅)
  try {
    // GitHub Pages URL — 딜러 확정 페이지
    const confirmUrl = 'https://' + GITHUB_PAGES_CONSULT + '/page_dealer_confirm.html?uid=' + uid + '&did=' + dealerId;
    
    const payload = {
      type: 'DEALER_ASSIGNED',
      dealerPhone: dealer.phone,
      dealerName: dealer.name,
      customerName: rowData['성함'],
      customerPhone: rowData['연락처'],
      address: addrText,
      days: rowData['가능요일'] || '미지정',
      timePrefs: rowData['선호시간대'] || '미지정',
      confirmUrl: confirmUrl
    };
    
    UrlFetchApp.fetch(MAKE_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('딜러 알림 발송 실패:', e);
  }
  
  // 고객에게 카톡 알림 (담당자 배정됨)
  try {
    const payload = {
      type: 'CUSTOMER_DEALER_ASSIGNED',
      phone: rowData['연락처'],
      name: rowData['성함'],
      dealerName: dealer.name
    };
    
    UrlFetchApp.fetch(MAKE_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('고객 알림 발송 실패:', e);
  }
  
  return { ok: true, message: dealer.name + ' 딜러에게 배정되었습니다.' };
}

// 딜러용: 일정 확정
function dealerConfirmSchedule(uid, dealerId, date, time, memo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 예약 찾기
  const uidIdx = headers.indexOf('UniqueID');
  const dealerIdIdx = headers.indexOf('dealerId');
  let rowIndex = -1;
  let rowData = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][uidIdx] === uid) {
      // 배정된 딜러인지 확인
      if (data[i][dealerIdIdx] !== dealerId) {
        throw new Error('이 예약에 대한 권한이 없습니다.');
      }
      rowIndex = i + 1;
      rowData = {};
      headers.forEach((h, idx) => { rowData[h] = data[i][idx]; });
      break;
    }
  }
  
  if (rowIndex < 0) throw new Error('예약을 찾을 수 없습니다.');
  
  // 컬럼 인덱스
  const dateColIdx = headers.indexOf('방문일');
  const timeColIdx = headers.indexOf('방문시간');
  const statusColIdx = headers.indexOf('Status');
  const notesColIdx = headers.indexOf('비고');
  
  // 시트 업데이트
  if (dateColIdx >= 0) sheet.getRange(rowIndex, dateColIdx + 1).setValue(date);
  if (timeColIdx >= 0) sheet.getRange(rowIndex, timeColIdx + 1).setValue(time);
  if (statusColIdx >= 0) sheet.getRange(rowIndex, statusColIdx + 1).setValue('CONFIRMED');
  if (notesColIdx >= 0 && memo) {
    const existingNotes = rowData['비고'] || '';
    const newNotes = existingNotes ? existingNotes + '\n[딜러메모] ' + memo : '[딜러메모] ' + memo;
    sheet.getRange(rowIndex, notesColIdx + 1).setValue(newNotes);
  }
  
  // 주소 조합
  let addrText = '';
  if (rowData['addressRoad']) addrText += String(rowData['addressRoad']).trim();
  if (rowData['addressDetail']) addrText += ' ' + String(rowData['addressDetail']).trim();
  if (!addrText.trim() && rowData['주소']) addrText = String(rowData['주소'] || '').trim();
  
  // 딜러 캘린더에 일정 생성 (calendarId가 있으면)
  const dealers = getDealers();
  const dealer = dealers.find(d => d.dealerId === dealerId);
  
  if (dealer && dealer.calendarId) {
    try {
      const cal = CalendarApp.getCalendarById(dealer.calendarId);
      if (cal) {
        const startDt = new Date(date + 'T' + time + ':00');
        const endDt = new Date(startDt.getTime() + 60 * 60 * 1000); // 1시간
        cal.createEvent(
          '출장상담: ' + rowData['성함'],
          startDt,
          endDt,
          {
            description: '고객: ' + rowData['성함'] + '\n연락처: ' + rowData['연락처'] + '\n주소: ' + addrText,
            location: addrText
          }
        );
      }
    } catch (e) {
      console.error('딜러 캘린더 생성 실패:', e);
    }
  }
  
  // 고객에게 확정 카톡 발송
  try {
    const payload = {
      type: 'FIELD_CONFIRMED',
      phone: rowData['연락처'],
      name: rowData['성함'],
      date: date,
      time: time,
      address: addrText,
      dealerName: dealer ? dealer.name : ''
    };
    
    UrlFetchApp.fetch(MAKE_WEBHOOK_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('고객 확정 알림 발송 실패:', e);
  }
  
  // 딜러에게 확정 완료 알림
  if (dealer) {
    try {
      const payload = {
        type: 'DEALER_CONFIRM_DONE',
        dealerPhone: dealer.phone,
        dealerName: dealer.name,
        customerName: rowData['성함'],
        date: date,
        time: time,
        address: addrText
      };
      
      UrlFetchApp.fetch(MAKE_WEBHOOK_URL, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
    } catch (e) {
      console.error('딜러 확정완료 알림 발송 실패:', e);
    }
  }
  
  return { ok: true, message: '일정이 확정되었습니다.' };
}

// 딜러용: 배정된 예약 조회
function getDealerAssignments(dealerId) {
  const sh = sh_();
  const col = headerMap_();
  const rows = sh.getDataRange().getValues();
  const results = [];
  
  // dealerId 컬럼 인덱스 확인
  const dealerIdColIdx = col['dealerId'];
  if (dealerIdColIdx == null) {
    Logger.log('dealerId 컬럼을 찾을 수 없습니다.');
    return [];
  }
  
  for (let i = 1; i < rows.length; i++) {
    const rowDealerId = String(rows[i][dealerIdColIdx] || '').trim();
    
    if (rowDealerId === String(dealerId).trim()) {
      // 주소 조합
      let addrText = '';
      if (col['addressRoad'] != null && rows[i][col['addressRoad']]) {
        addrText += String(rows[i][col['addressRoad']]).trim();
      }
      if (col['addressDetail'] != null && rows[i][col['addressDetail']]) {
        addrText += ' ' + String(rows[i][col['addressDetail']]).trim();
      }
      if (!addrText.trim() && col['주소'] != null) {
        addrText = String(rows[i][col['주소']] || '').trim();
      }
      
      results.push({
        uid: rows[i][col['UniqueID']] || '',
        status: rows[i][col['Status']] || '',
        name: rows[i][col['성함']] || '',
        phone: rows[i][col['연락처']] || '',
        address: addrText,
        days: rows[i][col['가능요일']] || '',
        timePrefs: rows[i][col['선호시간대']] || '',
        date: rows[i][col['방문일']] || '',
        time: rows[i][col['방문시간']] || ''
      });
    }
  }
  
  Logger.log('getDealerAssignments 결과: ' + results.length + '건, dealerId=' + dealerId);
  return results;
}

/************** 출장 시간 제안 **************/
function adminSuggestTimes(uid, suggestions){
  if (!uid || !Array.isArray(suggestions) || !suggestions.length) throw new Error('no suggestions');

  const sh = sh_(), col = headerMap_(), rows = sh.getDataRange().getValues();

  let row = -1;
  for (let i=1;i<rows.length;i++){
    if (String(rows[i][col['UniqueID']]) === String(uid)) { row = i+1; break; }
  }
  if (row < 0) throw new Error('uid not found');

  const durMin = getBusiness_().durMin || DEFAULT_DUR_MIN;
  const now = new Date();
  const expire = addMin_(now, (typeof HOLD_EXPIRE_HOURS==='number'?HOLD_EXPIRE_HOURS:6)*60);

  // 제안 → 병합(HOLD 버퍼 포함)
  const ranges = suggestions.map(s=>{
    const { start: st, end: et } = makeDateTimes_(s.date, s.time);
    return { s:addMin_(st, -BLOCK_BUFFER_MIN), e:addMin_(et, BLOCK_BUFFER_MIN), src:s };
  }).sort((a,b)=>a.s-b.s);

  const merged = [];
  for (const r of ranges){
    const last = merged[merged.length-1];
    if (!last) merged.push({ s:r.s, e:r.e, src:[r.src] });
    else if (r.s <= last.e){ last.e = (r.e>last.e? r.e:last.e); last.src.push(r.src); }
    else merged.push({ s:r.s, e:r.e, src:[r.src] });
  }

    // 매장 캘린더에 HOLD
  const cal = getCal_(CAL_STORE_NAME);
  const holdIds = [];

  // 캘린더 제목에 고객명 / 연락처 / 주소가 보이도록 한 번만 준비
  const baseRow = rows[row - 1];

  // 주소: 도로명 + 상세주소 우선, 없으면 기존 "주소" 컬럼 사용
  const addrParts = [];
  if (col['addressRoad'] != null && baseRow[col['addressRoad']]) {
    addrParts.push(baseRow[col['addressRoad']]);
  }
  if (col['addressDetail'] != null && baseRow[col['addressDetail']]) {
    addrParts.push(baseRow[col['addressDetail']]);
  }
  let addrCombined = addrParts.join(' ');
  if (!addrCombined && col['주소'] != null) {
    addrCombined = baseRow[col['주소']] || '';
  }

  // buildCalTitle_ 에 넘길 형태로 정리
  const calTitleRow = {
    name:    baseRow[col['성함']],
    phone:   baseRow[col['연락처']],
    address: addrCombined
  };

  // 병합된 구간마다 HOLD 생성 (제목은 고객 정보 + "출장 시간 제안")
  merged.forEach((m, i) => {
    const title = buildCalTitle_(calTitleRow, '출장 시간 제안', true); // true => [HOLD] 접두사
    const ev = cal.createEvent(title, m.s, m.e, {
      description: JSON.stringify({
        uid,
        expire: expire.toISOString(),
        type: 'FIELD_SUGGEST',
        hold_ts: Date.now()
      })
    });
    holdIds.push(ev.getId());
  });


  // 시트 저장
  const shortTok = generateShortToken_();
  const human = suggestions.map((s, idx) => `제안${idx+1}: ${s.date} ${s.time}`).join('\n');
  if (col['Status']       != null) sh.getRange(row, col['Status']+1).setValue('SUGGESTED');
  if (col['제안내용']     != null) sh.getRange(row, col['제안내용']+1).setValue(human);
  if (col['단축토큰']     != null) sh.getRange(row, col['단축토큰']+1).setValue(shortTok);
  if (col['holdExpireAt'] != null) sh.getRange(row, col['holdExpireAt']+1).setValue(expire);
  if (col['holdEventIds'] != null) sh.getRange(row, col['holdEventIds']+1).setValue(JSON.stringify(holdIds));

  // 고객 통지 — 알림톡 템플릿 변수 맞춤(suggest1~3, days[])
  const nameRaw  = String(rows[row-1][col['성함']] || '').trim();
  const phoneRaw = String(rows[row-1][col['연락처']] || '');
  const phone    = phoneRaw.replace(/\D/g, '');

  const labels = suggestions.map(s => `${s.date} ${s.time}`); // "YYYY-MM-DD HH:mm"
  const suggest1 = labels[0] || '';
  const suggest2 = labels[1] || '';
  const suggest3 = labels[2] || '';

      // GitHub Pages URL로 변경 — 카카오 인앱 커스텀 스킴 호환
  const confirm_link = GITHUB_PAGES_CONSULT + '/page_suggest.html?t=' + encodeURIComponent(shortTok);
  const resched_link = GITHUB_PAGES_CONSULT + '/page_reschedule.html?t=' + encodeURIComponent(shortTok);


  const payload = {
    topic: 'alrimtalk',          // ★ Make 라우팅 키 추가
    template: 'suggest',         // ★ 솔라피 템플릿명 추가
    id: uid,                     // ★ 다른 페이로드와 키 통일
    uid,
    token: shortTok,
    name: nameRaw,
    phone,                     // Solapi to
    confirm_link,              // #{confirm_link}
    resched_link,              // #{resched_link}
    suggest1,                  // #{제안1}
    suggest2,                  // #{제안2}
    suggest3,                  // #{제안3}
    days: labels,              // days[]
    suggestions,
    holdExpireAt: expire.toISOString(),
    channel: 'kakao',
    sms_fallback: false
  };

  postMake_('SUGGESTED_TIMES', payload);


  // 캐시 무효화
  try {
    const affected = [...new Set(suggestions.map(s => String(s.date).slice(0,10)))];
    affected.forEach(ds => slotsCacheInvalidate_(ds));
  } catch (_) {}
  
  // [NEW] HOLD 제목 생성
function composeHoldTitle_(rec, dStr, tStr){
  const nm   = String(rec.name || '');
  const ph   = String(rec.phone || '');
  const addr = String(rec.address || '');
  const uid  = String(rec.uid || '');
  return `${HOLD_PREFIX} 출장요청 | ${nm} ${ph} | ${addr} | 제안: ${dStr} ${tStr} [UID:${uid}]`;
}

// [NEW] 제안 목록으로 HOLD 생성(기존 UID의 HOLD는 정리 후 재생성)
function createFieldHoldEventsForSuggestions_(uid, suggestions){
  if (!Array.isArray(suggestions) || suggestions.length === 0) return;

  // 요청 레코드 조회(필드명: name, phone, address, uid 가정)
  const rec = getRequestByUid_(uid);
  if (!rec) return;

  const cal = CalendarApp.getCalendarsByName(CAL_NAME)[0];
  if (!cal) throw new Error('Calendar not found: '+CAL_NAME);

  // 기존 HOLD(같은 UID) 정리: 향후 30일 범위
  const now   = new Date();
  const until = new Date(now.getTime() + 30*24*60*60*1000);
  cal.getEvents(now, until).forEach(ev=>{
    const title = ev.getTitle() || '';
    if (title.includes(HOLD_PREFIX) && title.includes(`[UID:${uid}]`)) {
      try{ ev.deleteEvent(); }catch(_){}
    }
  });

  // 제안시각 그대로 시작, 기본 DUR 만큼 종료
  suggestions.forEach(s=>{
    const [y, m, d]    = String(s.date).split('-').map(Number);
    const [hh, mm]     = String(s.time).split(':').map(Number);
    const start        = new Date(y, m-1, d, hh, mm, 0);
    const end          = new Date(start.getTime() + (DEFAULT_DUR_MIN*60*1000));
    const title        = composeHoldTitle_(rec, s.date, s.time);
    const description  =
      `유형: 출장 요청\n`+
      `이름: ${rec.name || ''}\n`+
      `연락처: ${rec.phone || ''}\n`+
      `주소지: ${rec.address || ''}\n`+
      `제안시각: ${s.date} ${s.time}\n`+
      `UID: ${uid}`;

    cal.createEvent(title, start, end, { description });
  });
}
// [NEW] UID로 접수 레코드 조회
function getRequestByUid_(uid){
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const rows = sh.getDataRange().getValues();
  if (!rows || rows.length < 2) return null;

  const head = rows[0].map(v=>String(v||'').trim());
  const idx  = {};
  head.forEach((h,i)=> idx[h] = i);

  for (let r=1; r<rows.length; r++){
    if (String(rows[r][idx['UniqueID']]) === String(uid)){
      // 주소 조합
      let addrText = '';
      if (idx['addressRoad'] != null && rows[r][idx['addressRoad']]) {
        addrText += String(rows[r][idx['addressRoad']]).trim();
      }
      if (idx['addressDetail'] != null && rows[r][idx['addressDetail']]) {
        addrText += ' ' + String(rows[r][idx['addressDetail']]).trim();
      }
      if (!addrText && idx['주소'] != null) {
        addrText = String(rows[r][idx['주소']] || '').trim();
      }
      
      return {
        uid:     rows[r][idx['UniqueID']],
        name:    rows[r][idx['성함']],
        phone:   rows[r][idx['연락처']],
        address: addrText,
      };
    }
  }
  return null;
}


  return { ok:true, message:`${suggestions.length}개 제안. HOLD ${merged.length}개 생성.` };
}

// ===== PATCH 4: hasTimeConflict_ (매장+출장 동시 충돌 검사) =====
function hasTimeConflict_(start, end, exceptId){
  const dayStart = new Date(start.getFullYear(), start.getMonth(), start.getDate(), 0,0,0);
  const dayEnd   = new Date(dayStart.getTime() + 24*3600*1000);

  const cals = [ getCal_(CAL_STORE_NAME), getCal_(CAL_FIELD_NAME) ];
  for (const cal of cals){
    const evs = cal.getEvents(dayStart, dayEnd);
    for (const ev of evs){
      if (exceptId && ev.getId() === exceptId) continue;
      const s = ev.getStartTime().getTime();
      const e = ev.getEndTime().getTime();
      if (!(end.getTime() <= s || start.getTime() >= e)) return true;
    }
  }
  return false;
}
// ===== PATCH 4 END =====


/************** 알림 트리거 **************/
function sendReminders_() {
  const sh  = sh_(), col = headerMap_(), rows = sh.getDataRange().getValues();
  const now = new Date();
  let scanned=0, confirmed=0, cand24=0, sent24=0, cand2=0, sent2=0, errors=0;

  for (let i = 1; i < rows.length; i++) {
    scanned++;
    const r = rows[i];
    if (String(r[col['Status']]) !== 'CONFIRMED') continue;
    confirmed++;

    const toYMD = v =>
      v instanceof Date
        ? Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd')
        : String(v).slice(0, 10);
    const toHM = v =>
      v instanceof Date
        ? Utilities.formatDate(v, TIMEZONE, 'HH:mm')
        : String(v).slice(0, 5);

    const dt = toYMD(r[col['방문일']]);
    if (!dt) continue;
    const tm = toHM(r[col['방문시간']] || '10:00');

    const [yy, mm, dd] = dt.split('-').map(v => +v);
    const [HH, MM]     = tm.split(':').map(v => +v);
    const when    = new Date(yy, (mm || 1) - 1, (dd || 1), HH || 10, MM || 0, 0);
    const diffMin = (when.getTime() - now.getTime()) / 60000;

    // ▼ 주소 텍스트 생성 (도로명+상세 우선, 없으면 기존 "주소" 컬럼 사용)
    let addrText = '';
    if (col['addressRoad'] != null && r[col['addressRoad']]) {
      addrText += String(r[col['addressRoad']]).trim();
    }
    if (col['addressDetail'] != null && r[col['addressDetail']]) {
      addrText += (addrText ? ' ' : '') + String(r[col['addressDetail']]).trim();
    }
    if (!addrText && col['주소'] != null) {
      addrText = String(r[col['주소']] || '').trim();
    }

    // 24시간 전 리마인드 (15분 윈도우)
    if (diffMin <= 24 * 60 && diffMin > 24 * 60 - 15 && !r[col['Remind24']]) {
      try {
        const payload = {
          topic: 'alrimtalk',
          template: 'remind24',
          id: r[col['UniqueID']],
          name: r[col['성함']],
          phone: r[col['연락처']],
          type: r[col['상담방식']],
          date: dt,
          time: tm,
          address: addrText,      // ★ 방문 주소 추가
          trigger: 'time',
          channel: 'kakao',
          sms_fallback: false
        };
        const res = postMake_('REMINDER_24H', payload);

        if (res && res.ok === true) {
          sh.getRange(i + 1, col['Remind24'] + 1)
            .setValue(Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm'));
          sent24++;
        }
      } catch(e) {
        errors++;
        Logger.log('remind24 fail: ' + e);
      }
    }

    // 2시간 전 리마인드 (15분 윈도우)
    if (diffMin <= 120 && diffMin > 105 && !r[col['Remind2']]) {
      try {
        const payload2 = {
          topic: 'alrimtalk',
          template: 'remind2',
          id: r[col['UniqueID']],
          name: r[col['성함']],
          phone: r[col['연락처']],
          type: r[col['상담방식']],
          date: dt,
          time: tm,
          address: addrText,      // ★ 방문 주소 추가
          trigger: 'time',
          channel: 'kakao',
          sms_fallback: false
        };
        const res2 = postMake_('REMINDER_2H', payload2);

        if (res2 && res2.ok === true) {
          sh.getRange(i + 1, col['Remind2'] + 1)
            .setValue(Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm'));
          sent2++;
        }
      } catch(e) {
        errors++;
        Logger.log('remind2 fail: ' + e);
      }
    }
  }

  const report = { ok:true, scanned, confirmed, cand24, sent24, cand2, sent2, errors };
  Logger.log('[REPORT sendReminders_] ' + JSON.stringify(report));
  return report;
}


/****************************** 트리거 ******************************/
function INIT_TRIGGERS(){
  CLEAR_TRIGGERS();
  ScriptApp.newTrigger('cleanupExpiredHolds_').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('sendReminders_').timeBased().everyMinutes(15).create();
}
function CLEAR_TRIGGERS(){
  ScriptApp.getProjectTriggers().forEach(t=>{
    if (['cleanupExpiredHolds_','sendReminders_'].includes(t.getHandlerFunction())){
      ScriptApp.deleteTrigger(t);
    }
  });
}

/************** HOLD 자동만료 **************/
function cleanupExpiredHolds_(){
  const cal = CalendarApp.getCalendarById(ensureCalendar_());
  const now = Date.now();
  const expireMs = (typeof HOLD_EXPIRE_HOURS !== 'undefined' ? HOLD_EXPIRE_HOURS : 6) * 3600 * 1000;
  const keepBufMs= (typeof KEEP_BUFFER_HOURS !== 'undefined' ? KEEP_BUFFER_HOURS : 12) * 3600 * 1000;
  const biz = getBusiness_();
  let daysScanned=0, eventsChecked=0, holdsFound=0, eligible=0, deleted=0, sheetCancelled=0, makeSent=0, errors=0;
  for (let d = 0; d <= 30; d++) {
    daysScanned++;
    const base = new Date(); base.setHours(0,0,0,0); base.setDate(base.getDate()+d);
    const dayStart = new Date(base.getFullYear(), base.getMonth(), base.getDate(), biz.startHour, 0, 0);
    const dayEnd   = new Date(base.getFullYear(), base.getMonth(), base.getDate(), biz.endHour,   0, 0);
    cal.getEvents(dayStart, dayEnd).forEach(ev => {
      eventsChecked++;
      const title = ev.getTitle() || '';
      if (!title.startsWith(HOLD_PREFIX)) return;
      holdsFound++;
      const desc = ev.getDescription() || '';
      const mTs  = /hold_ts:(\d+)/.exec(desc);
      const mUid = /uid:([0-9a-f-]+)/.exec(desc);
      const isPending = /status:PENDING_ADMIN/.test(desc);
      if (!mTs || !isPending) return;
      const createdTs = +mTs[1];
      const startMs   = ev.getStartTime().getTime();
      const ageOk = (now - createdTs) > expireMs;
      const farOk = (startMs - now) > keepBufMs;
      if (!(ageOk && farOk)) return;
      eligible++;
      try { ev.deleteEvent(); deleted++; } catch(e){ errors++; Logger.log('delete hold fail: ' + e); }
      try{
        const ymd = Utilities.formatDate(new Date(startMs), TIMEZONE, 'yyyy-MM-dd');
        slotsCacheInvalidate_(ymd);
      }catch(e){}
      try {
        const sh = sh_(), col = headerMap_(), rows = sh.getDataRange().getValues();
        const uid = mUid ? mUid[1] : '';
        if (!uid) return;
        for (let i=1;i<rows.length;i++){
          if (String(rows[i][col['UniqueID']])===uid && String(rows[i][col['Status']])==='PENDING_ADMIN') {
            sh.getRange(i+1, col['Status']+1).setValue('CANCELLED'); sheetCancelled++;
            sh.getRange(i+1, col['CalendarEventID']+1).setValue('');
            sh.getRange(i+1, col['ConfirmationToken']+1).setValue('');
            const toYMD = v => v instanceof Date ? Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd') : String(v).slice(0, 10);
            const toHM = v => v instanceof Date ? Utilities.formatDate(v, TIMEZONE, 'HH:mm') : String(v).slice(0, 5);
            try {
              const payloadX = { topic: 'alrimtalk', template: 'expired', event: 'HOLD_EXPIRED', id: rows[i][col['UniqueID']], name: rows[i][col['성함']], phone: rows[i][col['연락처']], type: rows[i][col['상담방식']], date: toYMD(rows[i][col['방문일']]), time: toHM(rows[i][col['방문시간']]), trigger: 'time' };
              const r = postMake_('HOLD_EXPIRED', payloadX);
              if (r && r.ok) makeSent++;
            } catch(e){ errors++; Logger.log('Make webhook (expired) error: ' + e); }
            break;
          }
        }
      } catch(e){ errors++; Logger.log('sheet sync fail: ' + e); }
    });
  }
  const report = { ok:true, daysScanned, eventsChecked, holdsFound, eligible, deleted, sheetCancelled, makeSent, errors };
  Logger.log('[REPORT cleanupExpiredHolds_] ' + JSON.stringify(report));
  return report;
}

function INIT_FORMAT_SHEET_VIEW(){
  const sh = sh_(), col = headerMap_();
  sh.setFrozenRows(1);
  sh.setColumnWidths(col['성함']+1,1,120);
  sh.setColumnWidths(col['연락처']+1,1,140);
  sh.setColumnWidths(col['상담방식']+1,1,100);
  sh.setColumnWidths(col['방문일']+1,1,100);
  sh.setColumnWidths(col['방문시간']+1,1,80);
  sh.setColumnWidths(col['Status']+1,1,120);
  sh.getRange(2, col['방문일']+1, sh.getMaxRows()-1, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, col['방문시간']+1, sh.getMaxRows()-1, 1).setNumberFormat('HH:mm');
  const r = sh.getRange(2, col['Status']+1, sh.getMaxRows()-1, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('PENDING_ADMIN').setBackground('#FFF7ED').setFontColor('#92400E').setRanges([r]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('SUGGESTED').setBackground('#eff6ff').setFontColor('#1e40af').setRanges([r]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('CONFIRMED').setBackground('#ECFDF5').setFontColor('#065F46').setRanges([r]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('CANCELLED').setBackground('#FEE2E2').setFontColor('#991B1B').setRanges([r]).build()
  ];
  sh.setConditionalFormatRules(rules);
}

function RUN_cleanup(){
  const a = (typeof cleanupExpiredHolds_ === 'function') ? cleanupExpiredHolds_() : { ok:true, removed:0 };
  const b = cleanupCalendarHolds_();
  return { ok: (a.ok !== false) && (b.ok !== false), removed: (a.removed||0) + (b.removed||0) };
}

/** 캘린더 만료 HOLD 제거 + 시트 초기화 */
function cleanupCalendarHolds_(){
  const sh = sh_(), col = headerMap_(), rows = sh.getDataRange().getValues();
  const cal = getCal_(CAL_STORE_NAME);
  const now = new Date();
  let removed = 0;
  if (col['holdExpireAt']==null || col['holdEventIds']==null) return { ok:true, removed:0 };
  for (let r=2; r<=rows.length; r++){
    const expire = rows[r-1][col['holdExpireAt']];
    if (!expire) continue;
    const exp = expire instanceof Date ? expire : new Date(String(expire));
    if (!isFinite(exp)) continue;
    if (exp > now) continue;
    let ids = [];
    try{ ids = JSON.parse(String(rows[r-1][col['holdEventIds']]||'[]')); }catch(_){}
    ids.forEach(id=>{
      try{ const ev = cal.getEventById(id); if (ev){ ev.deleteEvent(); removed++; } }catch(_){}
    });
    if (col['holdEventIds']!=null) sh.getRange(r, col['holdEventIds']+1).setValue('[]');
    if (col['holdExpireAt']!=null) sh.getRange(r, col['holdExpireAt']+1).setValue('');
    try{
      const dateStr = rows[r-1][col['방문일']] instanceof Date
        ? Utilities.formatDate(rows[r-1][col['방문일']], TIMEZONE, 'yyyy-MM-dd')
        : String(rows[r-1][col['방문일']]||'');
      if (dateStr) slotsCacheInvalidate_(dateStr);
    }catch(_){}
  }
  return { ok:true, removed };
}

function RUN_reminders(){ return sendReminders_(); }


/************** 고객 선택 처리 **************/
function getSuggestionDataByToken_(token) {
  const sh = sh_();
  const col = headerMap_();
  const rows = sh.getDataRange().getValues();
  const baseUrl = ScriptApp.getService().getUrl();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][col['단축토큰']] === token) {
      const status = rows[i][col['Status']];
      if (status !== 'SUGGESTED') return null;

      const uid = rows[i][col['UniqueID']];
      const suggestions = (rows[i][col['제안내용']] || '')
        .split('\n')
        .map(line => line.replace(/제안\d: /, '').trim())
        .filter(Boolean);

      return {
        name: rows[i][col['성함']],
        buttons: suggestions.map(dt => ({
          label: dt,
          url: `${baseUrl}?a=c&u=${uid}&d=${encodeURIComponent(dt)}`
        })),
        rescheduleUrl: `https://${GITHUB_PAGES_CONSULT}/page_reschedule.html?t=${token}`
      };
    }
  }
  return null;
}
function getUidByToken_(token){
  const sh   = sh_();
  const col  = headerMap_();
  const rows = sh.getDataRange().getValues();

  const t = String(token || '');

  // 단축토큰 / UniqueID 컬럼이 없으면 바로 종료
  if (col['단축토큰'] == null || col['UniqueID'] == null) return null;

  for (let i = 1; i < rows.length; i++) {
    const rowTok = String(rows[i][col['단축토큰']] || '');
    if (rowTok === t) {
      return rows[i][col['UniqueID']];
    }
  }
  return null;
}
function clearHoldsForUid_(uid){
  if (!uid) return;

  const sh  = sh_();
  const col = headerMap_();
  const rows = sh.getDataRange().getValues();

  let row = -1;
  for (let i = 1; i < rows.length; i++){
    if (String(rows[i][col['UniqueID']]) === String(uid)){
      row = i + 1;
      break;
    }
  }
  if (row < 0) return;

  const r = rows[row - 1];

  // 1) holdEventIds에 저장된 HOLD 이벤트 삭제 (모든 캘린더에서)
  try {
    if (col['holdEventIds'] != null){
      let ids = [];
      try { ids = JSON.parse(String(r[col['holdEventIds']] || '[]')); } catch(_){}
      
      // 매장 캘린더에서 삭제
      try {
        const storeCal = getCal_(CAL_STORE_NAME);
        ids.forEach(id => {
          try {
            const hev = storeCal.getEventById(id);
            if (hev) hev.deleteEvent();
          } catch(_){}
        });
      } catch(_){}
      
      // ★ 출장 캘린더에서도 삭제 (CAL_FIELD_NAME)
      try {
        const fieldCal = getCal_(CAL_FIELD_NAME);
        ids.forEach(id => {
          try {
            const hev = fieldCal.getEventById(id);
            if (hev) hev.deleteEvent();
          } catch(_){}
        });
      } catch(_){}
      
      // 기본 캘린더(CAL_NAME)에서도 삭제
      try {
        const baseCals = CalendarApp.getCalendarsByName(CAL_NAME);
        if (baseCals && baseCals.length > 0) {
          const baseCal = baseCals[0];
          ids.forEach(id => {
            try {
              const hev = baseCal.getEventById(id);
              if (hev) hev.deleteEvent();
            } catch(_){}
          });
        }
      } catch(_){}
      
      sh.getRange(row, col['holdEventIds'] + 1).setValue('[]');
    }
    if (col['holdExpireAt'] != null){
      sh.getRange(row, col['holdExpireAt'] + 1).setValue('');
    }
  } catch(_){}

  // 2) 제안내용에서 날짜 추출하여 해당 날짜 범위의 HOLD 이벤트도 제목/설명으로 검색 삭제
  try {
    if (col['제안내용'] != null){
      const text = String(r[col['제안내용']] || '');
      const dates = text.split('\n')
        .map(line => line.replace(/제안\d:\s*/, '').trim())
        .filter(Boolean)
        .map(line => String(line).slice(0,10))
        .filter(ds => /^\d{4}-\d{2}-\d{2}$/.test(ds));
      const uniq = Array.from(new Set(dates));
      
      // 삭제할 캘린더 목록
      const calendarsToCheck = [];
      try { calendarsToCheck.push(getCal_(CAL_STORE_NAME)); } catch(_){}
      try { calendarsToCheck.push(getCal_(CAL_FIELD_NAME)); } catch(_){}
      try {
        const baseCals = CalendarApp.getCalendarsByName(CAL_NAME);
        if (baseCals && baseCals.length > 0) calendarsToCheck.push(baseCals[0]);
      } catch(_){}
      
      // 각 날짜에 대해 UID가 포함된 HOLD 이벤트 검색 삭제
      uniq.forEach(ds => {
        try {
          const [y, m, d] = ds.split('-').map(Number);
          const dayStart = new Date(y, m - 1, d, 0, 0, 0);
          const dayEnd = new Date(y, m - 1, d, 23, 59, 59);
          
          calendarsToCheck.forEach(cal => {
            if (!cal) return;
            try {
              const events = cal.getEvents(dayStart, dayEnd);
              events.forEach(ev => {
                const title = ev.getTitle() || '';
                const desc = ev.getDescription() || '';
                // HOLD 접두사가 있고, UID가 제목이나 설명에 포함된 경우 삭제
                if (title.includes(HOLD_PREFIX) && (title.includes(uid) || desc.includes(uid))) {
                  try { ev.deleteEvent(); } catch(_){}
                }
              });
            } catch(_){}
          });
          
          slotsCacheInvalidate_(ds);
        } catch(_){}
      });
    }
  } catch(_){}
  
  // 3) 제안내용 컬럼 초기화
  try {
    if (col['제안내용'] != null){
      sh.getRange(row, col['제안내용'] + 1).setValue('');
    }
  } catch(_){}
}




function confirmFieldRequest_(uid, datetime) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const [date, time] = datetime.split(' ');
    if (!date || !time) throw new Error("날짜시간 형식이 올바르지 않습니다.");

    const { start, end } = makeDateTimes_(date, time);


    const sh = sh_();
    const col = headerMap_();
    const rows = sh.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][col['UniqueID']] === uid) { row = i + 1; break; }
    }
    if (row < 0) return { ok: false, error: '예약 정보를 찾을 수 없습니다.' };

    const r = rows[row-1];
    if (r[col['Status']] !== 'SUGGESTED') return { ok: false, error: '이미 확정되었거나 취소된 예약입니다.'};

    sh.getRange(row, col['Status'] + 1).setValue('CONFIRMED');
    sh.getRange(row, col['방문일'] + 1).setValue(date);
    sh.getRange(row, col['방문시간'] + 1).setValue(time);
    sh.getRange(row, col['단축토큰'] + 1).setValue('');

    const name = r[col['성함']];
    const phone = r[col['연락처']];

       // 출장 확정은 출장 캘린더 생성 + 매장 HOLD 제거
    // 위치(Location)에 주소(도로명 + 상세주소)를 넣고, 우편번호는 제외
    const addrRoad   = (col['addressRoad']   != null ? String(r[col['addressRoad']]   || '').trim() : '');
    const addrDetail = (col['addressDetail'] != null ? String(r[col['addressDetail']] || '').trim() : '');
    let addrText = (addrRoad + ' ' + addrDetail).trim();

    // 분리된 컬럼이 비어 있으면 기존 "주소" 컬럼을 fallback 으로 사용
    if (!addrText && col['주소'] != null) {
      addrText = String(r[col['주소']] || '').trim();
    }

    const ev = getCal_(CAL_FIELD_NAME).createEvent(`[출장] ${name||''}`, start, end, {
      description: `고객:${name||''} / ${phone||''}`,
      location: addrText || ''
    });
    sh.getRange(row, col['CalendarEventID'] + 1).setValue(ev.getId());


    try{
      if (col['holdEventIds']!=null){
        const ids = JSON.parse(String(rows[row-1][col['holdEventIds']]||'[]'));
        const storeCal = getCal_(CAL_STORE_NAME);
        ids.forEach(id=>{ try{ const hev = storeCal.getEventById(id); if (hev) hev.deleteEvent(); }catch(_){ } });
        sh.getRange(row, col['holdEventIds']+1).setValue('[]');
      }
      if (col['holdExpireAt']!=null) sh.getRange(row, col['holdExpireAt']+1).setValue('');
    }catch(_){}

    postMake_('FIELD_CONFIRMED', {
  topic:    'alrimtalk',
  template: 'field_confirmed',
  id:       uid,
  name,
  phone,
  date,
  time,
  // ★ 추가: 방문 주소 (도로명+상세, 없으면 '주소' 컬럼)
  address:  addrText || '',
  // (선택) 다른 템플릿들과 통일
  channel:      'kakao',
  sms_fallback: false,
  change_request_link: GITHUB_PAGES_CONSULT + '/page_change_request.html?uid=' + uid  // 일정 변경/취소 요청 링크
});


    try {
      const dateStr = Utilities.formatDate(start, TIMEZONE, 'yyyy-MM-dd');
      slotsCacheInvalidate_(dateStr);
    } catch (e) { Logger.log('slotsCacheInvalidate_ fail(confirmFieldRequest_): ' + e); }

    // TMS Supabase 동기화 — 고객 확정 시 TMS에도 confirmed 반영
    try { pushToSupabase_(uid); } catch(e) { Logger.log('[supabase-sync] confirmField push 실패: ' + e); }

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ─── 고객 일정 변경/취소 요청 헬퍼 ───

/** uid로 예약 행 검색 → 핵심 정보 반환 */
function getReservationByUid_(uid) {
  const sh = sh_(), col = headerMap_(), rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][col['UniqueID']]) !== String(uid)) continue;
    const r = rows[i];
    const rawDate = r[col['방문일']];
    const rawTime = r[col['방문시간']];
    const dateStr = rawDate instanceof Date ? Utilities.formatDate(rawDate, TIMEZONE, 'yyyy-MM-dd') : String(rawDate || '');
    const timeStr = rawTime instanceof Date ? Utilities.formatDate(rawTime, TIMEZONE, 'HH:mm') : String(rawTime || '');
    let addrText = '';
    if (col['addressRoad'] != null && r[col['addressRoad']]) addrText += String(r[col['addressRoad']]).trim();
    if (col['addressDetail'] != null && r[col['addressDetail']]) addrText += (addrText ? ' ' : '') + String(r[col['addressDetail']]).trim();
    if (!addrText && col['주소'] != null) addrText = String(r[col['주소']] || '').trim();
    return {
      row: i + 1,
      name: String(r[col['성함']] || ''),
      phone: String(r[col['연락처']] || ''),
      type: String(r[col['상담방식']] || ''),
      date: dateStr,
      time: timeStr,
      status: String(r[col['Status']] || ''),
      address: addrText
    };
  }
  return null;
}

/** 고객 변경/취소 요청 처리 */
function submitChangeRequest_(uid, reqType, reason, memo, hopeDate) {
  if (!uid) return { ok: false, error: 'uid 누락' };
  if (!['change', 'cancel'].includes(reqType)) return { ok: false, error: '잘못된 요청 유형' };

  const info = getReservationByUid_(uid);
  if (!info) return { ok: false, error: '예약 정보를 찾을 수 없습니다.' };
  if (!['CONFIRMED', 'ASSIGNED'].includes(info.status.toUpperCase())) {
    return { ok: false, error: '변경/취소 요청이 불가능한 상태입니다. (현재: ' + info.status + ')' };
  }

  const sh = sh_(), col = headerMap_();
  const now = Utilities.formatDate(new Date(), TIMEZONE, 'MM-dd HH:mm');
  const isField = info.type.indexOf('출장') !== -1; // 출장 여부

  // 비고 컬럼에 요청 내용 append
  var noteCol = col['비고'];
  var prevNote = String(sh.getRange(info.row, noteCol + 1).getValue() || '');
  var newNote = '';
  if (reqType === 'change') {
    newNote = '[고객 변경요청 ' + now + '] 요청: ' + (hopeDate || '미입력') + (memo ? ' | 메모: ' + memo : '');
  } else {
    newNote = '[고객 취소요청 ' + now + '] 사유: ' + (reason || '미선택') + (memo ? ' | 메모: ' + memo : '');
  }
  sh.getRange(info.row, noteCol + 1).setValue(prevNote ? prevNote + '\n' + newNote : newNote);

  // ── 취소: 즉시 CANCELLED (캘린더 정리 포함, 알림톡 미발송) ──
  if (reqType === 'cancel') {
    // adminCancel의 캘린더/슬롯 정리 로직 재사용 (skipNotify=true → 알림톡 안 보냄)
    try { adminCancel(uid, { skipNotify: true }); } catch(e) { Logger.log('[selfCancel] adminCancel 실패, 상태만 변경: ' + e); sh.getRange(info.row, col['Status'] + 1).setValue('CANCELLED'); }
    try { pushToSupabase_(uid); } catch(e) { Logger.log('[supabase-sync] cancel push 실패: ' + e); }

    // 관리자 이메일 (취소)
    try {
      var emailSubject = '[MAMORU] 고객 예약 취소 - ' + info.name;
      var emailBody =
        '고객이 예약을 취소했습니다.\n\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
        '■ 취소된 예약 정보\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
        '• 고객명: ' + info.name + '\n' +
        '• 연락처: ' + info.phone + '\n' +
        '• 상담방식: ' + info.type + '\n' +
        '• 예약일: ' + info.date + '\n' +
        '• 예약시간: ' + info.time + '\n' +
        (info.address ? '• 주소: ' + info.address + '\n' : '') +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
        '• 취소 사유: ' + (reason || '미선택') + '\n' +
        (memo ? '• 메모: ' + memo + '\n' : '') +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
        '취소시각: ' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
      GmailApp.sendEmail('bsm@mamoru.kr', emailSubject, emailBody);
    } catch(emailErr) {
      Logger.log('[EMAIL ERROR] 취소 요청 이메일 발송 실패: ' + emailErr);
    }

    return { ok: true };
  }

  // ── 일정 변경 (출장만 여기 도달) ──
  sh.getRange(info.row, col['Status'] + 1).setValue('CHANGE_REQUESTED');
  try { pushToSupabase_(uid); } catch(e) { Logger.log('[supabase-sync] changeRequest push 실패: ' + e); }

  // 관리자 이메일 (출장 일정변경)
  try {
    var emailSubject = '[MAMORU] 출장 일정 변경 요청 - ' + info.name;
    var emailBody =
      '고객이 출장 일정 변경을 요청했습니다.\n\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      '■ 현재 예약 정보\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      '• 고객명: ' + info.name + '\n' +
      '• 연락처: ' + info.phone + '\n' +
      '• 예약일: ' + info.date + '\n' +
      '• 예약시간: ' + info.time + '\n' +
      (info.address ? '• 주소: ' + info.address + '\n' : '') +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      '■ 요청사항\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      (hopeDate || '미입력') + '\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
      '요청시각: ' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    GmailApp.sendEmail('bsm@mamoru.kr', emailSubject, emailBody);
  } catch(emailErr) {
    Logger.log('[EMAIL ERROR] 변경 요청 이메일 발송 실패: ' + emailErr);
  }

  // 접수 확인 알림톡 (출장 일정변경 전용 — Make → 솔라피)
  try {
    postMake_('CHANGE_REQUEST_RECEIVED', {
      topic: 'alrimtalk',
      template: 'change_request_received',
      id: uid,
      name: info.name,
      phone: info.phone,
      date: info.date,
      time: info.time,
      address: info.address || '',
      request_detail: hopeDate || '미입력',
      channel: 'kakao',
      sms_fallback: false
    });
  } catch(makeErr) {
    Logger.log('Make post error(CHANGE_REQUEST_RECEIVED): ' + makeErr);
  }

  return { ok: true };
}

function updateStatus_(uid, status) {
  const sh = sh_();
  const col = headerMap_();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][col['UniqueID']] === uid) {
      sh.getRange(i + 1, col['Status'] + 1).setValue(status);
      break;
    }
  }
}
/* ==========================================
 * [AppSheet 연동 전용 래퍼 함수251216 오후 11시34분 제미나이로 앱시트 연결시 추가한것 ]
 * 앱시트에서 들어오는 문자열 데이터를
 * 기존 함수가 원하는 형식으로 변환해주는 통역사들입니다.
 * ========================================== */

function appSheet_SuggestTimes(uid, suggestionsText) {
  // 앱시트에서 "2023-10-25 10:00, 2023-10-26 14:00" 형태로 들어오면
  // 기존 함수가 원하는 [{date:.., time:..}, ...] 형태로 변환합니다.
  
  if (!suggestionsText) throw new Error("제안 내용이 비어있습니다.");
  
  // 콤마(,)나 줄바꿈(Enter)으로 구분된 텍스트를 분리
  const list = String(suggestionsText).split(/[\n,]+/).map(s => s.trim()).filter(Boolean);
  const suggestions = [];

  list.forEach(item => {
    // "YYYY-MM-DD HH:mm" 형식 파싱
    const parts = item.split(' ');
    if (parts.length >= 2) {
      suggestions.push({ date: parts[0], time: parts[1] });
    }
  });

  if (suggestions.length === 0) throw new Error("날짜와 시간 형식이 올바르지 않습니다. (예: 2023-10-25 10:00)");

  // 기존 함수 호출
  return adminSuggestTimes(uid, suggestions);
}

function appSheet_AssignDealer(uid, dealerId) {
  // 딜러 배정 연결
  return adminAssignDealer(uid, dealerId);
}

function appSheet_Confirm(uid) {
  // 확정 연결
  return adminConfirm(uid);
}

function appSheet_Cancel(uid) {
  // 취소 연결
  return adminCancel(uid);
}

/**
 * 시트에서 차단된 슬롯 조회
 * - CONFIRMED/ASSIGNED: 방문일+방문시간 기준
 * - SUGGESTED: 제안내용 컬럼의 제안 시간들 기준
 */
function getBlockedSlotsFromSheet_(year, month1to12) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sh) return {};

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return {};

  const headers = data[0].map(h => String(h).trim());
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });

  const blocked = {};
  const biz = getBusiness_();
  const stepMin = biz.stepMin || 10;
  const durMin = biz.durMin || 60;

  // 특정 날짜+시간을 차단 목록에 추가하는 헬퍼
  function addBlockedSlots(dateStr, timeStr, isField) {
    if (!dateStr || !/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return;
    
    const [y, m] = dateStr.split('-').map(Number);
    if (y !== year || m !== month1to12) return;
    
    if (!timeStr) return;
    const [hh, mm] = timeStr.split(':').map(Number);
    if (isNaN(hh) || isNaN(mm)) return;
    
    if (!blocked[dateStr]) blocked[dateStr] = [];

    const baseMin = hh * 60 + mm;
    const beforeBuf = isField ? FIELD_BLOCK_BEFORE_MIN : 0;
    const afterBuf = isField ? FIELD_BLOCK_AFTER_MIN : 0;

    const startMin = baseMin - beforeBuf;
    const endMin = baseMin + durMin + afterBuf;

    for (let t = startMin; t < endMin; t += stepMin) {
      if (t < 0 || t >= 24 * 60) continue;
      const h = Math.floor(t / 60);
      const m = t % 60;
      const slot = String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
      if (!blocked[dateStr].includes(slot)) {
        blocked[dateStr].push(slot);
      }
    }
  }

  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][col['Status']] || '').toUpperCase();
    if (!['CONFIRMED', 'SUGGESTED', 'ASSIGNED'].includes(status)) continue;

    const consultType = String(data[i][col['상담방식']] || '');
    const isField = consultType === '출장 요청';

    // ★ SUGGESTED 상태: 제안내용 컬럼에서 제안 시간들 추출
    if (status === 'SUGGESTED') {
      const suggestions = String(data[i][col['제안내용']] || '');
      if (suggestions) {
        // "제안1: 2026-02-10 14:00\n제안2: 2026-02-11 15:00" 형태 파싱
        const lines = suggestions.split('\n');
        lines.forEach(line => {
          const match = line.match(/(\d{4}-\d{2}-\d{2})\s+(\d{2}:\d{2})/);
          if (match) {
            addBlockedSlots(match[1], match[2], true); // 제안은 항상 출장이므로 버퍼 적용
          }
        });
      }
      continue; // SUGGESTED는 방문일/방문시간이 비어있으므로 여기서 끝
    }

    // ★ CONFIRMED/ASSIGNED 상태: 방문일+방문시간 기준
    const rawDate = data[i][col['방문일']];
    const rawTime = data[i][col['방문시간']];

    let dateStr = '';
    if (rawDate instanceof Date) {
      dateStr = Utilities.formatDate(rawDate, TIMEZONE, 'yyyy-MM-dd');
    } else if (rawDate) {
      dateStr = String(rawDate).slice(0, 10);
    }

    let timeStr = '';
    if (rawTime instanceof Date) {
      timeStr = Utilities.formatDate(rawTime, TIMEZONE, 'HH:mm');
    } else if (rawTime) {
      const t = String(rawTime);
      const match = t.match(/(\d{1,2}):(\d{2})/);
      if (match) timeStr = match[1].padStart(2, '0') + ':' + match[2];
    }

    addBlockedSlots(dateStr, timeStr, isField);
  }

  return blocked;
}

/**
 * 시트 기반 월별 슬롯 조회
 */
function getSlotsMonth_SheetBased_(year, month1to12, type) {
  const settings = getSettings();
  const biz = settings.BUSINESS;
  const CLOSED_WEEKDAYS_ = settings.CLOSED_WEEKDAYS || [];
  const CLOSED_DATES_ = settings.CLOSED_DATES || [];

  const blocked = getBlockedSlotsFromSheet_(year, month1to12);
  const tz = Session.getScriptTimeZone();
  const daysInMonth = new Date(year, month1to12, 0).getDate();
  const now = new Date();
  const nowMs = now.getTime();
  const stepMin = biz.stepMin || 10;
  const durMin = biz.durMin || 60;

  const out = {};

  for (let d = 1; d <= daysInMonth; d++) {
    const dayObj = new Date(year, month1to12 - 1, d);
    const ymd = Utilities.formatDate(dayObj, tz, 'yyyy-MM-dd');

    // 과거 날짜
    if (dayObj.getTime() < new Date().setHours(0, 0, 0, 0)) {
      out[ymd] = [];
      continue;
    }

    // 휴무 요일
    if (CLOSED_WEEKDAYS_.includes(dayObj.getDay())) {
      out[ymd] = [];
      continue;
    }

    // 휴무일
    if (CLOSED_DATES_.includes(ymd)) {
      out[ymd] = [];
      continue;
    }

    const dayBlocked = blocked[ymd] || [];
    const slots = [];
    const isToday = dayObj.toDateString() === now.toDateString();

    for (let h = biz.startHour; h < biz.endHour; h++) {
      for (let m = 0; m < 60; m += stepMin) {
        const slotTime = new Date(year, month1to12 - 1, d, h, m);
        
        // 오늘이면 현재 시간 이전 슬롯 제외
        if (isToday && slotTime.getTime() < nowMs) continue;

        // 슬롯 종료 시간이 영업시간 초과하면 제외
        const endTime = new Date(slotTime.getTime() + durMin * 60000);
        if (endTime.getHours() > biz.endHour || 
            (endTime.getHours() === biz.endHour && endTime.getMinutes() > 0)) {
          continue;
        }

        const slot = String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');

        if (!dayBlocked.includes(slot)) {
          slots.push(slot);
        }
      }
    }

    out[ymd] = slots;
  }

  return out;
}

/**
 * 시트 기반 단일 날짜 슬롯 조회
 */
function getAvailableTimeSlots_SheetBased_(dateStr, type) {
  const match = /^(\d{4})-(\d{2})-\d{2}$/.exec(dateStr);
  if (!match) return [];

  const year = parseInt(match[1], 10);
  const month = parseInt(match[2], 10);
  const monthSlots = getSlotsMonth_SheetBased_(year, month, type);

  return monthSlots[dateStr] || [];
}

/* =====================================================
   [기존 함수 교체] - 시트 기반으로 전환
   ===================================================== */

// 기존 getSlotsMonth_ 백업 (필요시 복원용)
const _getSlotsMonth_Calendar_ = getSlotsMonth_;

// 기존 getAvailableTimeSlots_ 백업
const _getAvailableTimeSlots_Calendar_ = getAvailableTimeSlots_;

/**
 * 앱시트 취소 버튼 클릭 시 호출될 함수
 * @param {string} uniqueId 삭제할 예약의 고유 ID
 */
function cancelAppointmentByAppSheet(uniqueId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("상담접수"); // 시트 이름 확인 필수
  const data = sheet.getDataRange().getValues();
  
  // 1. 시트에서 해당 UniqueID 찾기
  for (let i = 1; i < data.length; i++) {
    if (data[i][10] === uniqueId) { // 11번째 열 (K열: UniqueID)
      
      const eventId = data[i][13]; // 14번째 열 (N열: CalendarEventID)
      const status = data[i][11];  // 12번째 열 (L열: Status)
      
      // 이미 취소된 건이면 중단
      if (status === "CANCELLED") return "이미 취소된 예약입니다.";

      // 2. 구글 캘린더 일정 삭제
      if (eventId) {
        try {
          // '마모루 출장방문' 달력 혹은 기본 달력에서 해당 ID 삭제
          const calendar = CalendarApp.getCalendarById("primary"); // 혹은 전용 달력 ID
          const event = calendar.getEventById(eventId);
          if (event) {
            event.deleteEvent();
          }
        } catch (e) {
          console.log("캘린더 삭제 중 오류(이미 없을 수 있음): " + e);
        }
      }

      // 3. 시트 상태 업데이트 (삭제 대신 'CANCELLED'로 변경 추천)
      sheet.getRange(i + 1, 12).setValue("CANCELLED"); // L열 (Status)
      
      // 4. (선택사항) 고객 취소 알림톡 발송 로직 추가 가능
      // sendCancelAlimtalk(data[i]); 

      return "취소 처리가 완료되었습니다.";
    }
  }
  return "해당 예약을 찾을 수 없습니다.";
}

/* 이메일 발송 테스트 함수 */
function testEmailSend_Consulting(){
  try{
    Logger.log('[TEST] 이메일 테스트 시작...');
    const testSubject = '[MAMORU 상담] 이메일 테스트';
    const testBody = '상담접수 이메일 발송 테스트입니다.\n시간: ' + new Date().toLocaleString('ko-KR');
    
    GmailApp.sendEmail('bsm@mamoru.kr', testSubject, testBody);
    Logger.log('[TEST] 이메일 발송 성공!');
    return { success: true, message: '이메일 발송 성공' };
  }catch(e){
    Logger.log('[TEST ERROR] ' + e.toString());
    return { success: false, error: e.toString() };
  }
}
