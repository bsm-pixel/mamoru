/**
 * MAMORU TMS — GAS → Supabase 상담 동기화 스크립트
 *
 * 사용법:
 * 1. Google Sheets → 확장 프로그램 → Apps Script
 * 2. 이 코드를 붙여넣기
 * 3. 아래 설정값 수정
 * 4. 트리거 설정: onFormSubmit (폼 제출 시 자동) 또는 syncAllToSupabase (수동 전체 동기화)
 */

// ═══════════════════════════════════════
// 설정 — 여기만 수정하세요
// ═══════════════════════════════════════
const CONFIG = {
  // Vercel 배포 URL (뒤에 /api/consultation/sync)
  API_URL: 'https://여기에-vercel-도메인.vercel.app/api/consultation/sync',
  // .env.local의 CRON_SECRET과 동일한 값
  SYNC_KEY: 'mamoru-tms-cron-2026',
  // 시트 이름
  SHEET_NAME: '상담접수',
};

// ═══════════════════════════════════════
// 컬럼 매핑 (0-based index)
// 시트 구조에 맞게 수정
// ═══════════════════════════════════════
const COL = {
  TIMESTAMP: 0,      // A: 접수시각
  NAME: 1,           // B: 성함
  PHONE: 2,          // C: 연락처
  CONSULT_TYPE: 3,   // D: 상담방식
  VISIT_DATE: 4,     // E: 방문희망일
  VISIT_TIME: 5,     // F: 방문희망시간
  POSTCODE: 6,       // G: 우편번호
  ADDR_ROAD: 7,      // H: 도로명주소
  ADDR_DETAIL: 8,    // I: 상세주소
  ADDR_SIDO: 9,      // J: 시도
  ADDR_SIGUNGU: 10,  // K: 시군구
  ADDR_REGION: 11,   // L: 지역
  MEMO: 12,          // M: 메모
  UNIQUE_ID: 13,     // N: UniqueID
  STATUS: 14,        // O: Status
  DEALER_CODE: 15,   // P: 딜러코드
  DEALER_NAME: 16,   // Q: 딜러명
  SUGGESTED1: 17,    // R: 제안일시1
  SUGGESTED2: 18,    // S: 제안일시2
  SUGGESTED3: 19,    // T: 제안일시3
  CONFIRMED: 20,     // U: 확정일시
  ADMIN_NOTE: 21,    // V: 관리자메모
  SOURCE: 22,        // W: 접수경로
};

/**
 * 시트 행 → API payload 변환
 */
function rowToPayload(row) {
  var val = function(idx) { return row[idx] ? String(row[idx]).trim() : ''; };

  var suggestedDates = [val(COL.SUGGESTED1), val(COL.SUGGESTED2), val(COL.SUGGESTED3)]
    .filter(function(d) { return d !== ''; });

  return {
    uniqueId: val(COL.UNIQUE_ID),
    name: val(COL.NAME),
    phone: val(COL.PHONE),
    consultType: val(COL.CONSULT_TYPE),
    visitDate: val(COL.VISIT_DATE),
    visitTime: val(COL.VISIT_TIME),
    postcode: val(COL.POSTCODE),
    addressRoad: val(COL.ADDR_ROAD),
    addressDetail: val(COL.ADDR_DETAIL),
    addressSido: val(COL.ADDR_SIDO),
    addressSigungu: val(COL.ADDR_SIGUNGU),
    addressRegion: val(COL.ADDR_REGION),
    memo: val(COL.MEMO),
    status: val(COL.STATUS) || 'PENDING_ADMIN',
    dealerCode: val(COL.DEALER_CODE),
    dealerName: val(COL.DEALER_NAME),
    suggestedDates: suggestedDates.length > 0 ? suggestedDates : undefined,
    confirmedDate: val(COL.CONFIRMED),
    adminNote: val(COL.ADMIN_NOTE),
    source: val(COL.SOURCE),
    receivedAt: val(COL.TIMESTAMP) ? new Date(val(COL.TIMESTAMP)).toISOString() : new Date().toISOString(),
  };
}

/**
 * Supabase API로 데이터 전송
 */
function sendToSupabase(payload) {
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-sync-key': CONFIG.SYNC_KEY },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  var response = UrlFetchApp.fetch(CONFIG.API_URL, options);
  var code = response.getResponseCode();
  var body = JSON.parse(response.getContentText());

  if (code !== 200) {
    Logger.log('동기화 실패 (' + code + '): ' + JSON.stringify(body));
    return false;
  }

  Logger.log('동기화 성공: ' + JSON.stringify(body));
  return true;
}

// ═══════════════════════════════════════
// 트리거 함수들
// ═══════════════════════════════════════

/**
 * 폼 제출 시 자동 동기화 (트리거: onFormSubmit)
 * 설정: 트리거 → 트리거 추가 → 이벤트 유형: 양식 제출 시
 */
function onFormSubmit(e) {
  try {
    var row = e.values;
    var payload = rowToPayload(row);

    // uniqueId가 없으면 자동 생성
    if (!payload.uniqueId) {
      payload.uniqueId = 'CONSULT-' + new Date().getTime();
      // 시트에도 UniqueID 기록
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
      var lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, COL.UNIQUE_ID + 1).setValue(payload.uniqueId);
    }

    sendToSupabase(payload);
  } catch (err) {
    Logger.log('onFormSubmit 에러: ' + err);
  }
}

/**
 * 전체 동기화 (수동 실행 또는 주기적 트리거)
 * Apps Script에서 직접 실행하거나, 시간 기반 트리거 설정
 */
function syncAllToSupabase() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    Logger.log('시트를 찾을 수 없음: ' + CONFIG.SHEET_NAME);
    return;
  }

  var data = sheet.getDataRange().getValues();
  var payloads = [];

  // 첫 행은 헤더이므로 건너뛰기
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var payload = rowToPayload(row);

    // 필수값 체크
    if (!payload.name || !payload.phone) continue;

    // uniqueId 없으면 자동 생성 + 시트에 기록
    if (!payload.uniqueId) {
      payload.uniqueId = 'CONSULT-' + new Date().getTime() + '-' + i;
      sheet.getRange(i + 1, COL.UNIQUE_ID + 1).setValue(payload.uniqueId);
    }

    payloads.push(payload);
  }

  if (payloads.length === 0) {
    Logger.log('동기화할 데이터 없음');
    return;
  }

  // 배치 전송 (50건씩)
  var batchSize = 50;
  var totalSynced = 0;

  for (var j = 0; j < payloads.length; j += batchSize) {
    var batch = payloads.slice(j, j + batchSize);

    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-sync-key': CONFIG.SYNC_KEY },
      payload: JSON.stringify({ batch: batch }),
      muteHttpExceptions: true,
    };

    var response = UrlFetchApp.fetch(CONFIG.API_URL, options);
    var body = JSON.parse(response.getContentText());
    totalSynced += (body.synced || 0);

    Logger.log('배치 ' + (Math.floor(j / batchSize) + 1) + ': ' + JSON.stringify(body));
  }

  Logger.log('전체 동기화 완료: ' + totalSynced + '건');
}

/**
 * 시트 편집 시 해당 행만 동기화 (선택적 트리거)
 * 설정: 트리거 → 트리거 추가 → 이벤트 유형: 수정 시
 */
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  var row = sheet.getRange(e.range.getRow(), 1, 1, 26).getValues()[0];
  var payload = rowToPayload(row);

  if (!payload.uniqueId || !payload.name || !payload.phone) return;

  sendToSupabase(payload);
}
