/**
 * MAMORU TMS — Supabase 상담 동기화
 * Code.gs의 submitConsultation 후 자동으로 Supabase에 Push
 *
 * 사용법:
 * 1. Code.gs의 submitConsultation 마지막에 pushToSupabase_(uid, ...) 호출 추가
 * 2. 또는 syncAllToSupabase() 로 전체 일괄 동기화
 */

// ═══════════════════════════════════════
// Supabase 연동 설정
// ═══════════════════════════════════════
var SUPABASE_SYNC_URL = 'https://여기에-vercel-도메인.vercel.app/api/consultation/sync';
var SUPABASE_SYNC_KEY = 'mamoru-tms-cron-2026';

/**
 * 시트 행 데이터 → Supabase payload 변환
 * headerMap_ 기반으로 동적 매핑
 */
function rowToSupabasePayload_(row, col) {
  var val = function(key) {
    return (col[key] != null && row[col[key]] != null) ? String(row[col[key]]).trim() : '';
  };

  var phone = val('연락처').replace(/^'/, ''); // 시트에 '010... 형태로 저장되므로 앞 따옴표 제거

  return {
    uniqueId: val('UniqueID'),
    name: val('성함'),
    phone: phone,
    consultType: val('상담방식'),
    visitDate: val('방문일'),
    visitTime: val('방문시간'),
    postcode: val('addressZip'),
    addressRoad: val('addressRoad'),
    addressDetail: val('addressDetail'),
    memo: val('비고'),
    status: val('Status') || 'PENDING_ADMIN',
    source: '웹폼',
    receivedAt: val('접수시각') ? new Date(val('접수시각')).toISOString() : new Date().toISOString(),
    raw: {
      address: val('주소'),
      days: val('가능요일'),
      timePrefs: val('선호시간대'),
      placeId: val('addressPlaceId'),
      lat: val('addressLat'),
      lng: val('addressLng'),
    },
  };
}

/**
 * 단건 Push — Code.gs에서 접수 직후 호출용
 * @param {string} uid - UniqueID
 */
function pushToSupabase_(uid) {
  try {
    var sh = sh_();
    var col = headerMap_();
    var rows = sh.getDataRange().getValues();

    // uid로 행 찾기
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][col['UniqueID']]) === uid) {
        var payload = rowToSupabasePayload_(rows[i], col);
        return sendToSupabase_(payload);
      }
    }
    Logger.log('[supabase-sync] UID 미발견: ' + uid);
  } catch (e) {
    Logger.log('[supabase-sync] pushToSupabase_ 에러: ' + e);
  }
}

/**
 * 전체 동기화 — 수동 실행 또는 시간 트리거
 * Apps Script 에디터에서 이 함수를 직접 실행
 */
function syncAllToSupabase() {
  var sh = sh_();
  var col = headerMap_();
  var rows = sh.getDataRange().getValues();
  var payloads = [];

  for (var i = 1; i < rows.length; i++) {
    var payload = rowToSupabasePayload_(rows[i], col);
    if (!payload.uniqueId || !payload.name || !payload.phone) continue;
    payloads.push(payload);
  }

  if (payloads.length === 0) {
    Logger.log('[supabase-sync] 동기화할 데이터 없음');
    return;
  }

  // 50건씩 배치 전송
  var batchSize = 50;
  var totalSynced = 0;

  for (var j = 0; j < payloads.length; j += batchSize) {
    var batch = payloads.slice(j, j + batchSize);
    var result = sendToSupabase_({ batch: batch });
    if (result) totalSynced += (result.synced || 0);
    Logger.log('[supabase-sync] 배치 ' + (Math.floor(j / batchSize) + 1) + ': ' + JSON.stringify(result));
  }

  Logger.log('[supabase-sync] 전체 동기화 완료: ' + totalSynced + '건 / ' + payloads.length + '건');
}

/**
 * Supabase API 호출
 */
function sendToSupabase_(payload) {
  try {
    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-sync-key': SUPABASE_SYNC_KEY },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    var response = UrlFetchApp.fetch(SUPABASE_SYNC_URL, options);
    var code = response.getResponseCode();
    var body = JSON.parse(response.getContentText());

    if (code !== 200) {
      Logger.log('[supabase-sync] 실패 (' + code + '): ' + JSON.stringify(body));
      return null;
    }
    return body;
  } catch (e) {
    Logger.log('[supabase-sync] 전송 에러: ' + e);
    return null;
  }
}
