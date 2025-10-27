/** 설정은 스크립트 속성에서 관리 */
const PROP = PropertiesService.getScriptProperties();
const SHEET_ID = PROP.getProperty('SHEET_ID'); // 필수
const SHEET_NAME = 'AS';
const CACHE_TTL = 20;
const CACHE_PREFIX = 'as_list_';

const MAKE_WEBHOOK_URL = PROP.getProperty('MAKE_WEBHOOK_URL') || '';
const MAKE_STATUS_WEBHOOK_URL = PROP.getProperty('MAKE_STATUS_WEBHOOK_URL') || '';
const ADMIN_TOKEN = PROP.getProperty('ADMIN_TOKEN') || '';

/* =========================
 * 진단/환경
 * ========================= */
function selfTestEnv_(){
  const P = PropertiesService.getScriptProperties();
  const out = {
    LOTTE_ENV: (P.getProperty('LOTTE_ENV') || 'dev').toLowerCase(),
    api_url_dev:  P.getProperty('LOTTE_API_URL_DEV')  || '',
    api_url_prod: P.getProperty('LOTTE_API_URL_PROD') || '',
    cancel_url_dev:  P.getProperty('LOTTE_CANCEL_API_URL_DEV')  || '',
    cancel_url_prod: P.getProperty('LOTTE_CANCEL_API_URL_PROD') || '',
    addr_api_url_dev:  P.getProperty('LOTTE_ADDR_API_URL_DEV')  || '',
    addr_api_url_prod: P.getProperty('LOTTE_ADDR_API_URL_PROD') || '',
    client_key_set_dev:  !!(P.getProperty('LOTTE_CLIENT_KEY_DEV')  || ''),
    client_key_set_prod: !!(P.getProperty('LOTTE_CLIENT_KEY_PROD') || ''),
    jobcustcd_set_dev:  !!(P.getProperty('LOTTE_JOBCUSTCD_DEV')  || ''),
    jobcustcd_set_prod: !!(P.getProperty('LOTTE_JOBCUSTCD_PROD') || ''),
    sender: {
      name: P.getProperty('LOTTE_SENDER_NAME') || '',
      tel:  P.getProperty('LOTTE_SENDER_TEL')  || '',
      zip:  P.getProperty('LOTTE_SENDER_ZIP')  || '',
      addr: P.getProperty('LOTTE_SENDER_ADDR') || ''
    },
    fare: P.getProperty('LOTTE_DEFAULT_FARE') || ''
  };
  Logger.log(JSON.stringify(out,null,2));
  return out;
}
function selfTestEnv(){  return selfTestEnv_(); }
function lotteResetCurrent(){ lotteResetCurrent_(); }
function lottePeekDemo(){ lottePeekDemo_(); }
function diagAddr(){ return diagAddr_(); }
function addrCheckZipAddr(){ 
  const P=PropertiesService.getScriptProperties();
  return addrCheck_(P.getProperty('LOTTE_SENDER_ZIP')||'', P.getProperty('LOTTE_SENDER_ADDR')||''); 
}
function diagAddrShow(){ const out=diagAddr_(); Logger.log(JSON.stringify(out,null,2)); return out; }
function addrCheckShow(){
  const P=PropertiesService.getScriptProperties();
  const out=addrCheck_(P.getProperty('LOTTE_SENDER_ZIP')||'', P.getProperty('LOTTE_SENDER_ADDR')||'');
  Logger.log(JSON.stringify(out,null,2)); return out;
}

/* =========================
 * 컬럼 맵
 * ========================= */
const COLS = { 접수번호:1, 고객명:2, 연락처:3, 진행방식:4, 우편번호:5, 주소:6, 상세주소:7, 수거요청일:8, 전달방법:9, 마모루수량:10, 타사수량:11, 메모:12, 비용:13, 생성일시:14, 입고완료:15, 입금완료:16, 출고완료:17, 송장번호:18, 택배사:19, 출고일:20, 배송완료:21 };

/* =========================
 * LOTTE 고정 운영 설정
 * ========================= */
function lotteConfig_(){
  const P = PropertiesService.getScriptProperties();
  return {
    env: 'prod',
    url: (P.getProperty('LOTTE_API_URL_PROD')||'').trim(),
    cancelUrl: (P.getProperty('LOTTE_CANCEL_API_URL_PROD')||'').trim(),
    clientKey: (P.getProperty('LOTTE_CLIENT_KEY_PROD')||'').trim(),
    jobCustCd: (P.getProperty('LOTTE_JOBCUSTCD_PROD')||'').trim(),
    sender: {
      name: P.getProperty('LOTTE_SENDER_NAME') || '',
      tel:  P.getProperty('LOTTE_SENDER_TEL')  || '',
      zip:  P.getProperty('LOTTE_SENDER_ZIP')  || '',
      addr: P.getProperty('LOTTE_SENDER_ADDR') || ''
    },
    fareSctCd: P.getProperty('LOTTE_DEFAULT_FARE') || '03'
  };
}

/* =========================
 * HTTP 유틸
 * ========================= */
function httpPostJson_(url, headers, payload, tries){
  tries = tries || 3;
  let lastErr = null;
  for (let i=0;i<tries;i++){
    try{
      const res = UrlFetchApp.fetch(url, {
        method:'post',
        contentType:'application/json',
        headers: headers || {},
        payload: JSON.stringify(payload),
        followRedirects:true,
        muteHttpExceptions:true
      });
      const code = res.getResponseCode();
      const text = res.getContentText() || '{}';
      const json = (function(){ try{ return JSON.parse(text); }catch(_){ return {raw:text}; }})();
      if (code >= 200 && code < 300) return {ok:true, code, json};
      if (code === 429 || code >= 500){ Utilities.sleep(600 * (i+1)); continue; }
      return {ok:false, code, json};
    }catch(e){ lastErr = e; Utilities.sleep(600 * (i+1)); }
  }
  if (lastErr) throw lastErr;
  throw new Error('HTTP_POST_FAILED');
}

/* =========================
 * ALPS 페이로드/호출
 * ========================= */
function lotteBuildSnd_(cfg, snd){
  return {
    snd_list: [{
      jobCustCd:   cfg.jobCustCd,
      ustRtgSctCd: snd.ustRtgSctCd || '01',
      ordSct:      snd.ordSct      || '3',
      fareSctCd:   snd.fareSctCd   || cfg.fareSctCd,
      ordNo:       String(snd.ordNo || ''),
      invNo:       String(snd.invNo || ''),

      snperNm:     cfg.sender.name,
      snperTel:    (cfg.sender.tel || '').replace(/\D/g,''),
      snperCpno:   '',
      snperZipcd:  cfg.sender.zip,
      snperAdr:    cfg.sender.addr,

      acperNm:     snd.rcvName,
      acperTel:    (snd.rcvTel || '').replace(/\D/g,''),
      acperCpno:   (snd.rcvCp  || snd.rcvTel || '').replace(/\D/g,''),
      acperZipcd:  snd.rcvZip,
      acperAdr:    snd.rcvAdr,

      boxTypCd:    snd.boxTypCd || 'A',
      gdsNm:       snd.gdsNm    || 'A/S 물품',
      dlvMsgCont:  snd.dlvMsg   || '',
      cusMsgCont:  snd.cusMsg   || '',
      pickReqYmd:  (snd.pickReqYmd || '').replace(/-/g,'')
    }]
  };
}

function lotteSend_(payload, cfg){
  const url = String(cfg.url||'').trim();
  const key = String(cfg.clientKey||'').trim();

  // !!!!!! 점검을 위해 이 1줄을 추가했습니다 !!!!!!
  Logger.log('!!! DEBUG: 롯데 API 호출 URL = ' + url); 

  if (!url || !key) throw new Error('LOTTE_CONFIG_MISSING');
  const r = httpPostJson_(url, { Authorization:'IgtAK '+key, Accept:'application/json' }, payload, 3);
  if (!r.ok) throw new Error('LOTTE_HTTP_'+r.code+': '+String(JSON.stringify(r.json)).slice(0,200));
  const first = (((r.json||{}).rtn_list)||[])[0]||{};
  return { ok:String(first.rtnCd||'').toUpperCase()==='S', code:first.rtnCd||'', msg:first.rtnMsg||'', raw:r.json };
}

/* =========================
 * 운송장 채번기
 * ========================= */
const LOTTE_START11 = parseInt(PROP.getProperty('LOTTE_START11') || '31765377481', 10);
const LOTTE_END11   = parseInt(PROP.getProperty('LOTTE_END11')   || '31765380480', 10);
const LOTTE_CUR11_K = 'LOTTE_CURRENT11';

function lotteCheckDigit_(base11){ return Number(base11) % 7; }
function lotteToInv_(base11){ return String(base11) + String(lotteCheckDigit_(base11)); }
function lotteGetCurrent11_(){ const v=PROP.getProperty(LOTTE_CUR11_K); return (v&&/^\d+$/.test(v))?parseInt(v,10):LOTTE_START11; }
function lotteSetCurrent11_(v){ PROP.setProperty(LOTTE_CUR11_K, String(v)); }
function lotteNextInv_(){
  let cur = lotteGetCurrent11_();
  if (cur > LOTTE_END11) throw new Error('운송장 대역 소진: 추가 대역 요청 필요');
  const inv = lotteToInv_(cur);
  const next = cur + 1;
  lotteSetCurrent11_( next <= LOTTE_END11 ? next : (LOTTE_END11+1) );
  lotteRangeGuard_();
  return inv;
}
function lottePeekInv_(){ const cur=lotteGetCurrent11_(); if(cur>LOTTE_END11) throw new Error('운송장 대역 소진'); return lotteToInv_(cur); }
function lotteResetCurrent_(){ lotteSetCurrent11_(LOTTE_START11); }
function lottePeekDemo_(){ const a=lottePeekInv_(); const b=lotteNextInv_(); const c=lotteNextInv_(); Logger.log('peek=%s next1=%s next2=%s',a,b,c); }

/* =========================
 * ALPS 예약/취소 고수준
 * ========================= */
function lotteBookSingle_(order){
  const cfg = lotteConfig_();
  const invNo = lottePeekInv_();
  const payload = lotteBuildSnd_(cfg, {
    ordNo: order.ordNo || ('TEST-'+Utilities.formatDate(new Date(),'Asia/Seoul','yyyyMMdd-HHmmss')),
    invNo: order.invNo || invNo,
    rcvName: order.rcvName||'', rcvTel:order.rcvTel||'', rcvZip:order.rcvZip||'', rcvAdr:order.rcvAdr||'',
    boxTypCd: order.boxTypCd||'A', gdsNm: order.gdsNm||'A/S 물품', dlvMsg: order.dlvMsg||'',
    pickReqYmd: (order.pickReqYmd||'')
  });

  Logger.log('[BOOK try] ordNo=%s invNo=%s', payload.snd_list[0].ordNo, payload.snd_list[0].invNo);
  const sent = lotteSend_(payload, cfg);
  if (!sent.ok) throw new Error('LOTTE_API_ERROR: ' + (sent.msg||''));
  lotteNextInv_();
  return { ok:true, invNo: payload.snd_list[0].invNo, rtnCd: sent.code||'', rtnMsg: sent.msg||'' };
}

function lotteCancel_(invNo, ordNo){ // ★★★ 1. 다시 invNo, ordNo만 받도록 변경 ★★★
  invNo = String(invNo||'').replace(/\D/g,'');
  ordNo = String(ordNo || '');
  if (!invNo || !ordNo) return {success:false, error:'MISSING_INVNO_OR_ORDNO'};
  const P = PropertiesService.getScriptProperties();
  const url = (P.getProperty('LOTTE_CANCEL_API_URL_PROD')||'').trim();
  const key = (P.getProperty('LOTTE_CLIENT_KEY_PROD')||'').trim();
  const jobCustCd = (P.getProperty('LOTTE_JOBCUSTCD_PROD')||'').trim();
  if (!url || !key || !jobCustCd) return {success:false, error:'CANCEL_CONFIG_MISSING'};

  // ★★★ 2. payload 구조 변경: snd_list 없이 최상위 레벨에 값 포함 ★★★
  const payload = { jobCustCd, invNo, ordNo };

  const r = httpPostJson_(url, { Authorization:'IgtAK '+key, Accept:'application/json' }, payload, 3);
  const rawJson = r.json || {};
  Logger.log('[CANCEL] Raw Response: ' + JSON.stringify(rawJson));

  // --- 롯데 응답 해석 로직 강화 (이전과 동일) ---
  let isSuccess = false;
  let finalCode = '';
  let finalMsg = '';
  if (r.ok) {
    const rtnList = rawJson.rtn_list || []; // 취소 성공 시에도 rtn_list가 올 수 있으므로 유지
    const first = rtnList[0] || {};
    finalCode = String(first.rtnCd || rawJson.code || '').toUpperCase();
    finalMsg = first.rtnMsg || rawJson.message || '';
    if (finalCode === 'S' || (!finalCode && !finalMsg)) { isSuccess = true; }
  } else {
    finalMsg = `HTTP 오류 ${r.code}`;
    if (rawJson.message) finalMsg += `: ${rawJson.message}`;
    finalCode = String(rawJson.code || '');
  }
  // --- 응답 해석 끝 ---

  if (isSuccess) {
    Logger.log('[CANCEL] Parsed Result: OK');
    return {success:true};
  } else {
    const errMsg = `롯데 응답: [코드=${finalCode || '없음'}, 메시지=${finalMsg || '없음'}] (HTTP ${r.code})`;
    Logger.log('[CANCEL] Parsed Result: FAILED (' + errMsg + ')');
    return {success:false, error: errMsg };
  }
} // <--- 여기까지 24줄

/* =========================
 * 접수번호 시퀀스
 * ========================= */
function nextAsId_(){
  const tz='Asia/Seoul', today=Utilities.formatDate(new Date(),tz,'yyyyMMdd');
  const key='AS_SEQ_'+today;
  const cur=parseInt(PROP.getProperty(key)||'0',10);
  PROP.setProperty(key,String(cur+1));
  return 'AS-'+today+'-'+String(cur+1).padStart(3,'0');
}

/* =========================
 * HTTP 엔드포인트
 * ========================= */
function doGet(e){
  const p = (e && e.parameter) ? e.parameter : {};
  if (p.ping === '1') return _json({ ok:true, now:new Date().toISOString() });

  if (p.holidays === '1') return _json({ holidays: getHolidayList_(p.years) });
  if (p.action === 'list') return _json(listAS_(p));
  if (p.action === 'update'){ if (ADMIN_TOKEN && p.token !== ADMIN_TOKEN) return _json({success:false,error:'UNAUTHORIZED'}); return _json(updateAS_(p)); }
  if (p.action === 'delete'){ if (ADMIN_TOKEN && p.token !== ADMIN_TOKEN) return _json({success:false,error:'UNAUTHORIZED'}); return _json(deleteAS_(p.as_id)); }
  if (p.action === 'addr_check'){ if (ADMIN_TOKEN && p.token !== ADMIN_TOKEN) return _json({success:false,error:'UNAUTHORIZED'}); return _json(addrCheckMake_(p.zip||'', p.addr||'')); }
  if (p.action === 'diag_addr'){ if (ADMIN_TOKEN && p.token !== ADMIN_TOKEN) return _json({success:false,error:'UNAUTHORIZED'}); return _json(diagAddr_()); }
  if (p.action === 'dupcheck'){ return _json(findDupAS_(p.phone||'', p.addr1||'')); }
  if (p.action === 'book'){      // 수동 단건 예약
    if (ADMIN_TOKEN && p.token !== ADMIN_TOKEN) return _json({success:false,error:'UNAUTHORIZED'});
    try{ return _json(manualBook(String(p.as_id||'').trim())); }catch(e2){ return _json({success:false,error:String(e2)}); }
  }
  if (p.action === 'cancel_single'){
    if (ADMIN_TOKEN && p.token !== ADMIN_TOKEN) return _json({success:false,error:'UNAUTHORIZED'});
    const inv = String(p.inv||'').replace(/\D/g,''); if(!inv) return _json({success:false,error:'MISSING_INV'});
    const out = lotteCancel_(inv); return _json(out);
  }

  if (p.view === 'admin'){
    return HtmlService.createHtmlOutputFromFile('admin')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e){
  try{
    const q = (e && e.parameter) ? e.parameter : {};
    let body = {};
    if (e && e.postData && e.postData.contents){ try{ body = JSON.parse(e.postData.contents); }catch(_){ body={}; } }

    // 관리자 JSON 이벤트
    if (body.event === 'AS_SHIP_DONE'){ if (ADMIN_TOKEN && body.token !== ADMIN_TOKEN) return _json({success:false,error:'UNAUTHORIZED'}); return _json(shipByAsId_(String(body.as_id||''))); }
    if (body.event === 'AS_SHIP_UNDO'){ if (ADMIN_TOKEN && body.token !== ADMIN_TOKEN) return _json({success:false,error:'UNAUTHORIZED'}); return _json(unshipByAsId_(String(body.as_id||''))); }

    // 관리자 쿼리
    if (q.action === 'update'){ if (ADMIN_TOKEN && q.token !== ADMIN_TOKEN) return _json({success:false,error:'UNAUTHORIZED'}); return _json(updateAS_(q)); }
    if (q.action === 'delete'){ if (ADMIN_TOKEN && q.token !== ADMIN_TOKEN) return _json({success:false,error:'UNAUTHORIZED'}); return _json(deleteAS_(q.as_id)); }
    if (q.action === 'unship'){ if (ADMIN_TOKEN && q.token !== ADMIN_TOKEN) return _json({success:false,error:'UNAUTHORIZED'}); return _json(unshipAS_(q.as_id)); }
    // [ADD_BOOK_ENDPOINT_POST_START]

// [ADD_BOOK_ENDPOINT_POST_START]
if (body.event === 'AS_BOOK'){
  if (ADMIN_TOKEN && body.token !== ADMIN_TOKEN) return _json({success:false, error:'UNAUTHORIZED'});
  return _json(bookAS_(body.as_id));
}
// [ADD_BOOK_ENDPOINT_POST_END]


    // 고객 접수
    if (body.event === 'AS_CREATE'){
      // 중복 차단(최근 24h 동일 주소)
      const dup = findDupAS_(body.phone, body.address1);
      if (dup && dup.found){ return _json({success:false, error:'DUP_FOUND', as_id:dup.as_id}); }

      const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
      if(!sh) return _json({success:false, error:'SHEET_NOT_FOUND'});

      const asId       = nextAsId_(); // 일자+시퀀스
      const createdAt  = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      const fixedPhone = normalizePhone(body.phone);

      const qtyM = parseInt(body.qty_mamoru || '0', 10) || 0;
      const qtyO = parseInt(body.qty_other  || '0', 10) || 0;
      const cost = qtyM * 10000 + qtyO * 20000;

      const row = [
        asId,
        body.name || '',
        fixedPhone,
        body.proceed_type || '',
        body.postcode || '',
        body.address1 || '',
        body.address2 || '',
        body.pickup_date || '',
        body.delivery_method || '',
        body.qty_mamoru || '',
        body.qty_other || '',
        body.memo || '',
        cost,
        createdAt,
        '', '', '',
        '', '', '',
        ''
      ];
      const next = sh.getLastRow() + 1;
      sh.getRange(next, COLS.연락처).setNumberFormat('@');
      sh.getRange(next, 1, 1, row.length).setValues([row]);
      sh.getRange(next, COLS.연락처).setValue("'" + fixedPhone);
      sh.getRange(next, COLS.비용).setNumberFormat('#,##0');

      // 메이크 웹훅
      try{
        UrlFetchApp.fetch(MAKE_WEBHOOK_URL, {
          method:'post', contentType:'application/json', muteHttpExceptions:true,
          payload: JSON.stringify({
            as_id:asId, name:body.name||'', phone:fixedPhone,
            proceed_type:body.proceed_type||'', delivery_method:body.delivery_method||'',
            pickup_date:body.pickup_date||'', qty_mamoru:body.qty_mamoru||'',
            qty_other:body.qty_other||'', created_at:createdAt
          })
        });
      }catch(e1){ Logger.log('[MAKE] '+e1); }

      

      return _json({success:true, as_id:asId});
    }

    return _json({success:false, error:'bad'});
  }catch(err){
    return _json({success:false, error:String(err)});
  }
}

/* =========================
 * 리스트/업데이트/삭제
 * ========================= */
function listAS_(p){
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if(!sh) return {success:false, error:'SHEET_NOT_FOUND'};

  const status = (p.status||'').trim();          // 'progress' | 'done' | ''
  const cacheKey = CACHE_PREFIX + (status || 'all');
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const last = sh.getLastRow();
  if (last < 2){
    const empty = {success:true, rows:[], total:0};
    cache.put(cacheKey, JSON.stringify(empty), CACHE_TTL);
    return empty;
  }

  const start = Math.max(2, last - 500 + 1);
  const vals = sh.getRange(start, 1, last - start + 1, sh.getLastColumn()).getValues();

  const rows = [];
  for (let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i];
    const row = {
      접수번호:r[COLS.접수번호-1], 고객명:r[COLS.고객명-1], 연락처:r[COLS.연락처-1],
      진행방식:r[COLS.진행방식-1], 고객메모:r[COLS.메모-1], 마모루:r[COLS.마모루수량-1],
      타사:r[COLS.타사수량-1], 비용:r[COLS.비용-1], 입고:r[COLS.입고완료-1],
      입금:r[COLS.입금완료-1], 출고:r[COLS.출고완료-1], 송장번호:r[COLS.송장번호-1],
      택배사:r[COLS.택배사-1], 출고일:r[COLS.출고일-1], 배송완료:r[COLS.배송완료-1],
      생성일시:r[COLS.생성일시-1],
      우편번호:r[COLS.우편번호-1], 주소:r[COLS.주소-1], 상세주소:r[COLS.상세주소-1],
      수거요청일:r[COLS.수거요청일-1], 전달방법:r[COLS.전달방법-1]
    };

    const isDone = (row.입고==='Y' && row.입금==='Y' && row.출고==='Y');
    if (status === 'done'     && !isDone) continue;
    if (status === 'progress' &&  isDone) continue;

    rows.push(row);
  }

  const out = {success:true, rows, total:rows.length};
  cache.put(cacheKey, JSON.stringify(out), CACHE_TTL);
  return out;
}

/* 출고 플로우 유틸 */
function _sheet_(){ return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME); }
function _col_(sh, header){ const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; const idx=headers.indexOf(header); if(idx<0) throw new Error('헤더 없음: '+header); return idx+1; }
function _findRowByAsId_(sh, asId){
  const colAsId=_col_(sh,'접수번호'); const last=sh.getLastRow(); if(last<2) return -1;
  const vals=sh.getRange(2,colAsId,last-1,1).getValues().flat();
  const idx=vals.indexOf(asId); return (idx<0)?-1:(2+idx);
}
function shipRow_(row){
  const sh=_sheet_(); const colYn=_col_(sh,'출고완료'); const colWhen=_col_(sh,'출고일');
  sh.getRange(row,colYn).setValue('Y');
  sh.getRange(row,colWhen).setValue(Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd'));
  return true;
}
function unshipRow_(row){
  const sh=_sheet_(); const colYn=_col_(sh,'출고완료'); let colWhen=null;
  try{ colWhen=_col_(sh,'출고일'); }catch(_){}
  sh.getRange(row,colYn).clearContent(); if(colWhen) sh.getRange(row,colWhen).clearContent(); return true;
}
function shipByAsId_(asId){
  const sh=_sheet_(); const row=_findRowByAsId_(sh,asId); if(row<0) return {success:false,error:'AS_ID_NOT_FOUND'};
  shipRow_(row); return {success:true,row};
}
function unshipByAsId_(asId){
  const sh=_sheet_(); const row=_findRowByAsId_(sh,asId); if(row<0) return {success:false,error:'AS_ID_NOT_FOUND'};
  unshipRow_(row); return {success:true,row};
}

function updateAS_(p){
  const asId = (p.as_id||'').trim(); if(!asId) return {success:false, error:'MISSING_AS_ID'};
  const field = (p.field||'').trim(); const value = (p.value||'').trim();

  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if(!sh) return {success:false, error:'SHEET_NOT_FOUND'};

  const last = sh.getLastRow(); if (last < 2) return {success:false, error:'EMPTY'};
  const data = sh.getRange(2, 1, last-1, sh.getLastColumn()).getValues();
  let rowIdx = -1;
  for (let i=0;i<data.length;i++){ if (String(data[i][COLS.접수번호-1]) === asId){ rowIdx = i+2; break; } }
  if (rowIdx < 0) return {success:false, error:'NOT_FOUND'};

  const setVal = (col,val)=> sh.getRange(rowIdx, col).setValue(val);
  let statusToSend = null;

  switch(field){
    case '입고완료':
      setVal(COLS.입고완료, value||'Y'); if((value||'Y')==='Y') statusToSend='inbound'; break;
    case '입금완료':
      setVal(COLS.입금완료, value||'Y'); if((value||'Y')==='Y') statusToSend='paid'; break;
    case '출고완료': {
      const hasInv = !!sh.getRange(rowIdx, COLS.송장번호).getValue();
      if (!hasInv) return {success:false, error:'MISSING_INVOICE'};
      const val=(value||'Y'); setVal(COLS.출고완료,val);
      if (val==='Y'){
        setVal(COLS.출고일, Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd'));
        if(!sh.getRange(rowIdx, COLS.택배사).getValue()) setVal(COLS.택배사,'롯데택배');
        statusToSend='shipped';
      }else{ setVal(COLS.출고일,''); }
      break;
    }
    case '송장번호': setVal(COLS.송장번호, value); if(!sh.getRange(rowIdx,COLS.택배사).getValue()) setVal(COLS.택배사,'롯데택배'); break;
    case '택배사': setVal(COLS.택배사, value); break;
    case '배송완료': setVal(COLS.배송완료, value||'Y'); break;
    default: return {success:false, error:'INVALID_FIELD'};
  }

  if (statusToSend) {
    const rowVals = sh.getRange(rowIdx, 1, 1, sh.getLastColumn()).getValues()[0];
    const payload = {
      event:'AS_STATUS',
      as_id:String(rowVals[COLS.접수번호-1]||''), name:String(rowVals[COLS.고객명-1]||''), phone:String(rowVals[COLS.연락처-1]||''),
      qty_mamoru:String(rowVals[COLS.마모루수량-1]||''), qty_other:String(rowVals[COLS.타사수량-1]||''), total_amount:String(rowVals[COLS.비용-1]||''),
      status:statusToSend, tracking:String(rowVals[COLS.송장번호-1]||''), courier:String(rowVals[COLS.택배사-1]||''),
      when:Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd HH:mm:ss')
    };
    try{
      UrlFetchApp.fetch(MAKE_STATUS_WEBHOOK_URL, { method:'post', contentType:'application/json', payload: JSON.stringify(payload), muteHttpExceptions:true });
    }catch(err){ Logger.log('[STATUS WH][ERR] '+err); }
  }

  const _cache = CacheService.getScriptCache();
  _cache.remove(CACHE_PREFIX + 'all'); _cache.remove(CACHE_PREFIX + 'progress'); _cache.remove(CACHE_PREFIX + 'done');
  return {success:true};
}

function deleteAS_(asId){
  asId = (asId||'').trim();
  if (!asId) return {success:false, error:'MISSING_AS_ID'};

  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if(!sh) return {success:false, error:'SHEET_NOT_FOUND'};

  const last = sh.getLastRow(); if (last < 2) return {success:false, error:'EMPTY'};

  const vals = sh.getRange(2, 1, last-1, sh.getLastColumn()).getValues();
  let rowIdx = -1, row=null;
  for (let i=0;i<vals.length;i++){ if (String(vals[i][COLS.접수번호-1]) === asId) { rowIdx = i + 2; row=vals[i]; break; } }
  if (rowIdx < 0) return {success:false, error:'NOT_FOUND'};

  // ALPS 취소 우선
  var invNow = String(row[COLS.송장번호-1]||'').replace(/\D/g,'');
  if (invNow) {
    try {
      const r = lotteCancel_(invNow);
      if (!r.success) return {success:false, error:'ALPS_CANCEL_FAILED:'+String(r.error||'')};
    } catch (e) { return {success:false, error:'ALPS_CANCEL_EXCEPTION:'+String(e)}; }
  }

  sh.deleteRow(rowIdx);

  const _cache2 = CacheService.getScriptCache();
  _cache2.remove(CACHE_PREFIX + 'all'); _cache2.remove(CACHE_PREFIX + 'progress'); _cache2.remove(CACHE_PREFIX + 'done');

  return {success:true};
}

/* =========================
 * 주소/휴일/기타 유틸
 * ========================= */
function addrCheck_(zip, addr){
  const P = PropertiesService.getScriptProperties();
  const env = (P.getProperty('LOTTE_ENV') || 'dev').toLowerCase();
  const url = (env === 'prod' ? (P.getProperty('LOTTE_ADDR_API_URL_PROD')||'') : (P.getProperty('LOTTE_ADDR_API_URL_DEV')||'')).trim();
  const key = (env === 'prod' ? (P.getProperty('LOTTE_CLIENT_KEY_PROD')||'') : (P.getProperty('LOTTE_CLIENT_KEY_DEV')||'')).trim();
  if (!url || !key) return {success:false, error:'CONFIG_MISSING'};

  const payload = { network:'00', address:String(addr||''), zip_no:String(zip||'') };
  const r = httpPostJson_(url, { 'Authorization':'IgtAK '+key, 'Accept':'application/json' }, payload, 2);
  const body = (r && r.json) || {};
  const ok = String(body.result || body.status || '').toLowerCase() === 'success';
  const deliverable = ok && !body.dlv_msg;
  return { success: ok, status: r.code, deliverable, raw: body, all: body };
}

function addrCheckMake_(zip, addr){
  var url = 'https://hook.eu2.make.com/6a2h9hbp3nzn54ysx5swvtgsa78lmxv8';
  var payload = { address: String(addr||''), zip: String(zip||'') };
  var res = UrlFetchApp.fetch(url, { method:'post', contentType:'application/json', payload: JSON.stringify(payload), followRedirects:true, muteHttpExceptions:true });
  var body = {}; try{ body = JSON.parse(res.getContentText()||'{}'); }catch(_){}
  return { success:true, status: res.getResponseCode(), out: body };
}

function diagAddr_(){
  const P = PropertiesService.getScriptProperties();
  const env = (P.getProperty('LOTTE_ENV')||'dev').toLowerCase();
  const url = (env === 'prod'
    ? (P.getProperty('LOTTE_ADDR_API_URL_PROD') || '')
    : (P.getProperty('LOTTE_ADDR_API_URL_DEV')  || '')
  ).trim();
  const key = (env === 'prod'
    ? (P.getProperty('LOTTE_CLIENT_KEY_PROD') || '')
    : (P.getProperty('LOTTE_CLIENT_KEY_DEV')  || '')
  ).trim();
  return { ok: !!url && !!key, env, url_in_use: url, key_len: key.length, key_head: key.slice(0,12) };
}

function addrCheckTest(){
  const P = PropertiesService.getScriptProperties();
  return addrCheck_(P.getProperty('LOTTE_SENDER_ZIP')||'', P.getProperty('LOTTE_SENDER_ADDR')||'');
}

function normalizePhone(phoneRaw){
  if(!phoneRaw) return '';
  let num=String(phoneRaw).replace(/\D/g,'');
  if(num.length===10&&num.startsWith('10')) num='0'+num;
  else if(num.length===11&&num.startsWith('10')) num='0'+num.slice(0,10);
  else if(!num.startsWith('0')) num='010'+num.slice(-8);
  return num;
}

function setupHeaders(){
  const headers=['접수번호','고객명','연락처','진행방식','우편번호','주소','상세주소','수거요청일','전달방법','마모루수량','타사수량','메모','비용','생성일시','입고완료','입금완료','출고완료','송장번호','택배사','출고일','배송완료'];
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if(!sh) throw new Error('시트 탭명을 확인하세요: '+SHEET_NAME);
  sh.clear();
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('C:C').setNumberFormat('@');
}

function _json(obj){ return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }



function getHolidayList_(yearsParam){
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('HOLIDAYS'); if(!sh) return [];
  const last = sh.getLastRow(); if(last<2) return [];
  const vals = sh.getRange(2,1,last-1,1).getValues().flat().filter(v=>v);
  const dates = vals.map(v=>Utilities.formatDate(new Date(v),'Asia/Seoul','yyyy-MM-dd'));
  if(yearsParam){ const years=String(yearsParam).split(',').map(s=>s.trim()); return dates.filter(d=>years.includes(d.slice(0,4))); }
  return dates;
}

function keepAlive(){
  const url = PROP.getProperty('KEEPALIVE_URL');
  if (!url) return;
  try{ UrlFetchApp.fetch(url,{muteHttpExceptions:true}); }catch(e){ Logger.log(e); }
}

/* 안전 스위치 */
function _prodEnabled_(){ return (PropertiesService.getScriptProperties().getProperty('LOTTE_PROD_ENABLED')||'0') === '1'; }

/* 시트 기록 헬퍼 */
function _writeInvoiceToRow_(sh, rowIdx, invNo){
  sh.getRange(rowIdx, COLS.송장번호).setValue(invNo);
  if (!sh.getRange(rowIdx, COLS.택배사).getValue()) sh.getRange(rowIdx, COLS.택배사).setValue('롯데택배');
  sh.getRange(rowIdx, COLS.출고일).setValue('');
  sh.getRange(rowIdx, COLS.출고완료).setValue('');
}

/* 수동 1건 예약 */
function manualBook(asId){
  if (!_prodEnabled_()) throw new Error('LOTTE_PROD_ENABLED=1 로 설정 후 실행');
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  const sh = ss.getSheetByName('AS'); if(!sh) throw new Error('시트 없음');

  const last = sh.getLastRow(); if(last<2) throw new Error('데이터 없음');
  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  let rowIdx=-1, row=null;
  for (let i=0;i<vals.length;i++){ if (String(vals[i][0])===String(asId)){ rowIdx=i+2; row=vals[i]; break; } }
  if (rowIdx<0) throw new Error('접수번호 미발견: '+asId);

  const COL = { 고객명:2, 연락처:3, 우편번호:5, 주소:6, 상세주소:7, 메모:12 };
  const order = {
    ordNo: asId,
    rcvName: String(row[COL.고객명-1]||''),
    rcvTel:  String(row[COL.연락처-1]||''),
    rcvZip:  String(row[COL.우편번호-1]||''),
    rcvAdr:  String((row[COL.주소-1]||'')+' '+(row[COL.상세주소-1]||'')).trim(),
    gdsNm:   'AS 출고',
    dlvMsg:  String(row[COL.메모-1]||'')
  };

  const out = lotteBookSingle_(order);
  _writeInvoiceToRow_(sh, rowIdx, out.invNo);
  Logger.log('OK inv=%s asId=%s', out.invNo, asId);
  return {success:true, invNo:out.invNo};
}
function manualBookDemo(){ return manualBook('AS-20251020-114503'); }

/* 출고 확정 처리 */
function markShipped(asId){
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  const sh = ss.getSheetByName('AS'); if(!sh) throw new Error('시트 없음');

  const last = sh.getLastRow(); if(last<2) throw new Error('데이터 없음');
  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  let rowIdx=-1; for (let i=0;i<vals.length;i++){ if(String(vals[i][0])===String(asId)){ rowIdx=i+2; break; } }
  if (rowIdx<0) throw new Error('접수번호 미발견: '+asId);

  sh.getRange(rowIdx, COLS.출고완료).setValue('Y');
  sh.getRange(rowIdx, COLS.출고일).setValue(Utilities.formatDate(new Date(),'Asia/Seoul','yyyy-MM-dd'));
  Logger.log('shipped: %s', asId);
}

/* 출고 취소: ALPS 취소 → 성공 시 롤백 */
function unshipAS_(asId){
  asId = String(asId||'').trim();
  if (!asId) return {success:false, error:'MISSING_AS_ID'};

  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if(!sh) return {success:false, error:'SHEET_NOT_FOUND'};
  const last = sh.getLastRow(); if(last<2) return {success:false, error:'EMPTY'};

  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  let rowIdx=-1, row=null;
  for(let i=0;i<vals.length;i++){ if(String(vals[i][0])===asId){ rowIdx=i+2; row=vals[i]; break; } }
  if(rowIdx<0) return {success:false, error:'NOT_FOUND'};

  /* ... (unshipAS_ 함수 내부) ... */
  const inv = String(row[COLS.송장번호-1]||'').replace(/\D/g,'');
  if (inv){
    const c = lotteCancel_(inv, asId); // ★★★ invNo, asId(ordNo)만 전달하도록 원복 ★★★
    if (!c.success) return {success:false, error:'ALPS_CANCEL_FAILED:'+String(c.error||'')};
  }

  

  sh.getRange(rowIdx, COLS.출고완료).setValue('');
  sh.getRange(rowIdx, COLS.출고일).setValue('');
  sh.getRange(rowIdx, COLS.송장번호).setValue('');
  sh.getRange(rowIdx, COLS.택배사).setValue('');

  const cache = CacheService.getScriptCache();
  cache.remove(CACHE_PREFIX + 'all'); cache.remove(CACHE_PREFIX + 'progress'); cache.remove(CACHE_PREFIX + 'done');

  return {success:true};
}
// [ADD_BOOK_HELPER_START]
function bookAS_(asId){
  asId = String(asId||'').trim();
  if (!asId) return {success:false, error:'MISSING_AS_ID'};

  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if(!sh) return {success:false, error:'SHEET_NOT_FOUND'};
  const last = sh.getLastRow(); if(last<2) return {success:false, error:'EMPTY'};

  const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  let rowIdx=-1, row=null;
  for (let i=0;i<vals.length;i++){ if(String(vals[i][0])===asId){ rowIdx=i+2; row=vals[i]; break; } }
  if(rowIdx<0) return {success:false, error:'NOT_FOUND'};

  // 이미 송장 있으면 중복 방지
  const curInv = String(row[COLS.송장번호-1]||'').replace(/\D/g,'');
  if (curInv) return {success:false, error:'ALREADY_HAS_INVOICE'};

  // 수하인 정보 → ALPS 예약
  const order = {
    ordNo: asId,
    rcvName: String(row[COLS.고객명-1]||''),
    rcvTel:  String(row[COLS.연락처-1]||''),
    rcvZip:  String(row[COLS.우편번호-1]||''),
    rcvAdr:  String((row[COLS.주소-1]||'')+' '+(row[COLS.상세주소-1]||'')).trim(),
    boxTypCd:'A',
    gdsNm:   'AS 출고',
    dlvMsg:  String(row[COLS.메모-1]||''),
    pickReqYmd: '' // ★★★ 송장 생성(출고)시에는 수거요청일을 보내지 않음 ★★★
  };

  const out = lotteBookSingle_(order);              // ALPS 채번
  _writeInvoiceToRow_(sh, rowIdx, out.invNo);       // 시트에 송장/택배사만 기록
  return {success:true, invNo: out.invNo};
}
// [ADD_BOOK_HELPER_END]


/* 운송장 대역 임계 */
function lotteRangeGuard_(){
  const start = LOTTE_START11, end = LOTTE_END11, cur = lotteGetCurrent11_();
  const total = (end - start + 1);
  const used  = Math.max(0, Math.min(total, (cur - start)));
  const ratio = used / total;
  if (ratio >= 0.90) {
    try{
      UrlFetchApp.fetch(MAKE_WEBHOOK_URL, {
        method:'post', contentType:'application/json', muteHttpExceptions:true,
        payload: JSON.stringify({
          event:'LOTTE_RANGE_LOW',
          start11:String(start), end11:String(end), cur11:String(cur),
          used, total, ratio:Number(ratio.toFixed(3))
        })
      });
    }catch(e){ Logger.log(e); }
  }
}

/* 주소/전화 정규화 및 중복검사 */
function normalizeAddr_(s){ return String(s||'').replace(/\s+/g,'').replace(/[^\w가-힣0-9]/g,'').toLowerCase(); }
function findDupAS_(phoneRaw, addr1Raw){
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sh) return { found:false };
  const last = sh.getLastRow(); if (last < 2) return { found:false };
  const rng = sh.getRange(2,1, Math.min(500, last-1), sh.getLastColumn()).getValues();
  const now = Date.now(), cutoff = now - 24*60*60*1000;

  const tgtAddr = normalizeAddr_(addr1Raw);
  const tgtPhone = (phoneRaw || '').replace(/\D/g,'');

  for(let i=rng.length-1;i>=0;i--){
    const r = rng[i];
    const created = r[COLS.생성일시-1];
    const ts = created ? new Date(created).getTime() : 0;
    if(!ts || ts < cutoff) continue;

    const addr1 = r[COLS.주소-1]||'';
    const phone = String(r[COLS.연락처-1]||'').replace(/\D/g,'');
    const sameAddr  = normalizeAddr_(addr1) === tgtAddr;
    const samePhone = tgtPhone && phone.endsWith(tgtPhone.slice(-8));
    if (sameAddr /* || (sameAddr && samePhone) */){
      return { found:true, as_id:String(r[COLS.접수번호-1]||'') };
    }
  }
  return { found:false };
}

/* 스프레드시트 UI 메뉴 */
function onOpen(){
  try{
    SpreadsheetApp.getUi()
      .createMenu('관리')
      .addItem('선택행 출고완료', 'uiFinishShip')
      .addItem('선택행 출고취소', 'uiCancelShip')
      .addToUi();
  }catch(e){ Logger.log(e); }
}
function uiFinishShip(){
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const r  = sh.getActiveCell().getRow();
  if (r < 2) { SpreadsheetApp.getUi().alert('데이터 행을 선택하세요.'); return; }
  const asId = String(sh.getRange(r, COLS.접수번호).getValue()||'').trim();
  if (!asId){ SpreadsheetApp.getUi().alert('접수번호가 비어 있습니다.'); return; }
  const res = adminFinishShipById_(asId);
  SpreadsheetApp.getUi().alert(res && res.success ? '출고 처리 완료' : '출고 처리 실패: '+(res&&res.error||'UNKNOWN'));
}
function uiCancelShip(){
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const r  = sh.getActiveCell().getRow();
  if (r < 2) { SpreadsheetApp.getUi().alert('데이터 행을 선택하세요.'); return; }
  const asId = String(sh.getRange(r, COLS.접수번호).getValue()||'').trim();
  if (!asId){ SpreadsheetApp.getUi().alert('접수번호가 비어 있습니다.'); return; }
  const res = adminCancelShipById_(asId);
  SpreadsheetApp.getUi().alert(res && res.success ? '출고 취소 완료' : '출고 취소 실패: '+(res&&res.error||'UNKNOWN'));
}
function adminFinishShipById_(asId){ try{ return updateAS_({ as_id:String(asId||'').trim(), field:'출고완료', value:'Y' }); }catch(e){ Logger.log('[adminFinishShipById_] '+e); return {success:false, error:String(e)}; } }
function adminCancelShipById_(asId){ try{ return unshipAS_(String(asId||'').trim()); }catch(e){ Logger.log('[adminCancelShipById_] '+e); return {success:false, error:String(e)}; } }



