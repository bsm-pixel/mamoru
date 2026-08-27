/**
 * ALPS 클라이언트 — GAS Code.gs와 client.ts의 검증된 로직 사용
 *
 * client.ts의 getConfig/alpsHeaders/httpPost/buildSndPayload를 직접 재활용.
 * bookSingle()은 nextWaybill() RPC를 호출하므로 우회하고,
 * 기존 getNextInvoice() + buildSndPayload + httpPost 조합 사용.
 */

import { createServiceClient } from '@/lib/supabase/server';
import { randomUUID } from 'crypto';

/* ── 환경변수 (GAS lotteConfig_ 포팅 — 모든 값 trim) ── */
const LOTTE_API_URL = (process.env.LOTTE_API_URL || '').trim();
const LOTTE_CANCEL_API_URL = (process.env.LOTTE_CANCEL_API_URL || '').trim();
const LOTTE_TRACK_API_URL = (process.env.LOTTE_TRACK_API_URL || '').trim();
const LOTTE_CLIENT_KEY = (process.env.LOTTE_CLIENT_KEY || '').trim();
const LOTTE_JOBCUSTCD = (process.env.LOTTE_JOB_CUST_CD || process.env.LOTTE_JOBCUSTCD || '').trim();
const LOTTE_FARE = (process.env.LOTTE_DEFAULT_FARE || '03').trim();

/* ── 발송인 (GAS cfg.sender 동일) ── */
const SENDER = {
  name: (process.env.LOTTE_SENDER_NAME || '마모루').trim(),
  tel: (process.env.LOTTE_SENDER_TEL || '').trim(),
  zip: (process.env.LOTTE_SENDER_ZIP || '').trim(),
  addr: (process.env.LOTTE_SENDER_ADDR || '').trim(),
};

/* ── ALPS 헤더 (GAS lotteSend_ 헤더와 동일) ── */
function alpsHeaders() {
  return {
    Authorization: `IgtAK ${LOTTE_CLIENT_KEY}`,
    Accept: 'application/json',
    'Content-Type': 'application/json; charset=utf-8',
    'X-Idempotency-Key': randomUUID(),
    'X-Correlation-Id': randomUUID(),
  };
}

/* ── HTTP POST (GAS httpPostJson_ 동일 — 재시도 3회) ── */
function sleep(ms: number) { return new Promise(r => setTimeout(r, ms)); }

async function httpPost(url: string, headers: Record<string, string>, payload: unknown, tries = 3) {
  let lastErr: Error | null = null;
  for (let i = 0; i < tries; i++) {
    try {
      const res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(payload) });
      const text = await res.text();
      let json: Record<string, unknown>;
      try { json = JSON.parse(text); } catch { json = { raw: text }; }
      if (res.ok) return { ok: true, code: res.status, json };
      if (res.status === 429 || res.status >= 500) { await sleep(600 * (i + 1)); continue; }
      return { ok: false, code: res.status, json };
    } catch (e) {
      lastErr = e as Error;
      await sleep(600 * (i + 1));
    }
  }
  throw lastErr || new Error('HTTP_POST_FAILED');
}

/* ── 체크디짓 ── */
export function checkDigit(base11: number): number { return base11 % 7; }
export function toInvoiceNumber(base11: number): string { return `${base11}${checkDigit(base11)}`; }

/* ── 송장번호 발급 (DB 카운터, 원자적) ── */
export async function getNextInvoice(): Promise<{ invoiceNumber: string; base11: number }> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db = createServiceClient() as any;
  const { data: config, error: fetchErr } = await db
    .from('lotte_waybill_config').select('*').eq('id', 'default').single();
  if (fetchErr || !config) throw new Error('송장번호 설정이 없습니다.');

  const current = Number(config.current_number);
  const end = Number(config.end_number);
  if (current > end) throw new Error(`송장번호 풀 소진 (${current}/${end})`);

  const { error: updateErr } = await db
    .from('lotte_waybill_config')
    .update({ current_number: current + 1, updated_at: new Date().toISOString() })
    .eq('id', 'default').eq('current_number', current);
  if (updateErr) throw new Error('송장번호 발급 충돌. 다시 시도해주세요.');

  return { invoiceNumber: toInvoiceNumber(current), base11: current };
}

/* ── 실패 시 친절 안내 + 사장님 앱 즉시 알림 (재발 조용히 실패 방지) ── */
function friendlyAlpsError(msg: string): string {
  const s = String(msg || '');
  // 롯데 계약/거래처 관련 거절이면 행동 안내를 덧붙임 (거래처계약정보 조회 오류 등)
  if (/계약|거래처/.test(s)) {
    return `${s} — 롯데 거래처 계약 문제일 수 있습니다. 롯데글로벌로지스에 계약 상태를 확인하세요.`;
  }
  return s;
}

// 실패 기록 + 30분 스로틀 앱 푸시. 비차단(송장 응답을 막지 않음).
async function reportAlpsFailure(friendlyErr: string, invNo?: string): Promise<void> {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const now = new Date().toISOString();
    await db.from('system_settings').upsert(
      { key: 'lotte.last_invoice_error', value: JSON.stringify({ at: now, error: friendlyErr, inv: invNo || null }), updated_at: now },
      { onConflict: 'key' },
    );
    // 30분 내 중복 알림 방지
    const { data } = await db.from('system_settings').select('value').eq('key', 'lotte.last_alert_at').maybeSingle();
    const lastAt = data?.value ? new Date(data.value).getTime() : 0;
    if (Date.now() - lastAt < 30 * 60 * 1000) return;
    await db.from('system_settings').upsert(
      { key: 'lotte.last_alert_at', value: now, updated_at: now },
      { onConflict: 'key' },
    );
    const { sendPushToAll } = await import('@/lib/firebase/send-push');
    await sendPushToAll({
      title: '⚠️ 롯데 송장 생성 실패',
      body: (invNo ? `송장 ${invNo} · ` : '') + friendlyErr,
      url: '/sales',
      tag: 'lotte-invoice-fail',
    });
  } catch (e) {
    console.error('[ALPS] 실패 알림 처리 오류:', e);
  }
}

/* ── 송장 발급 (GAS lotteBuildSnd_ + lotteSend_ 100% 동일) ── */
export async function bookShipment(order: {
  invoiceNumber: string;
  receiverName: string;
  receiverTel: string;
  receiverZip: string;
  receiverAddr: string;
  goodsName?: string;
  deliveryMessage?: string;
  ustRtgSctCd?: string;        // '01'=출고(기본) / '02'=반품(회수) — 롯데 IS팀 안내(2026-08-27)
  orgInvoiceNumber?: string;   // 원송장번호(반품 시 선택) → orglInvNo
}): Promise<{ success: boolean; invoiceNumber: string; error?: string }> {
  if (!LOTTE_API_URL || !LOTTE_CLIENT_KEY) {
    return { success: false, invoiceNumber: order.invoiceNumber, error: 'LOTTE API 환경변수 미설정' };
  }
  if (!LOTTE_JOBCUSTCD) {
    return { success: false, invoiceNumber: order.invoiceNumber, error: 'LOTTE_JOBCUSTCD 환경변수 미설정' };
  }

  const now = new Date();
  const pad2 = (n: number) => String(n).padStart(2, '0');
  const ordNo = `TMS-${now.getFullYear()}${pad2(now.getMonth() + 1)}${pad2(now.getDate())}-${pad2(now.getHours())}${pad2(now.getMinutes())}${pad2(now.getSeconds())}`;
  const pickReqYmd = `${now.getFullYear()}${pad2(now.getMonth() + 1)}${pad2(now.getDate())}`;

  // GAS lotteBuildSnd_ 100% 동일 payload
  const payload = {
    snd_list: [{
      jobCustCd:   LOTTE_JOBCUSTCD,
      ustRtgSctCd: order.ustRtgSctCd || '01',   // '02'=반품(회수). 출고/반품 외 양식 동일(롯데 IS팀)
      ordSct:      '1',
      fareSctCd:   LOTTE_FARE,
      ordNo:       ordNo,
      invNo:       order.invoiceNumber,
      orglInvNo:   (order.orgInvoiceNumber || '').replace(/\D/g, ''),  // 원송장번호(반품 시 선택)

      snperNm:     SENDER.name,
      snperTel:    SENDER.tel.replace(/\D/g, ''),
      snperCpno:   '',
      snperZipcd:  SENDER.zip,
      snperAdr:    SENDER.addr,

      acperNm:     order.receiverName,
      acperTel:    (order.receiverTel || '').replace(/\D/g, ''),
      acperCpno:   (order.receiverTel || '').replace(/\D/g, ''),
      acperZipcd:  order.receiverZip,
      acperAdr:    order.receiverAddr,

      boxTypCd:    'A',
      gdsNm:       order.goodsName || '마모루 제품',
      dlvMsgCont:  order.deliveryMessage || '',
      cusMsgCont:  '',
      pickReqYmd:  pickReqYmd,
    }],
  };

  try {
    const r = await httpPost(LOTTE_API_URL, alpsHeaders(), payload, 3);
    if (!r.ok) {
      const fe = friendlyAlpsError(`ALPS HTTP ${r.code}`);
      reportAlpsFailure(fe, order.invoiceNumber).catch(() => {});
      return { success: false, invoiceNumber: order.invoiceNumber, error: fe };
    }

    // ALPS 응답 로그 (디버깅용)
    console.log('[ALPS bookShipment] 응답:', JSON.stringify(r.json).slice(0, 500));

    // GAS lotteSend_ 동일: rtn_list[0].rtnCd === 'S'
    const rtnList = r.json.rtn_list;
    const first = (Array.isArray(rtnList) ? rtnList[0] : {}) as Record<string, unknown>;
    const rtnCd = String(first.rtnCd || '').toUpperCase();

    console.log('[ALPS bookShipment] rtnCd:', rtnCd, 'rtnMsg:', first.rtnMsg);

    if (rtnCd === 'S') {
      return { success: true, invoiceNumber: order.invoiceNumber };
    }

    const fe = friendlyAlpsError(String(first.rtnMsg || rtnCd));
    reportAlpsFailure(fe, order.invoiceNumber).catch(() => {});
    return { success: false, invoiceNumber: order.invoiceNumber, error: fe };
  } catch (err) {
    const fe = friendlyAlpsError(err instanceof Error ? err.message : String(err));
    reportAlpsFailure(fe, order.invoiceNumber).catch(() => {});
    return { success: false, invoiceNumber: order.invoiceNumber, error: fe };
  }
}

/* ── 송장 취소 ── */
export async function cancelShipment(invoiceNumber: string): Promise<{ success: boolean; error?: string }> {
  if (!LOTTE_CANCEL_API_URL || !LOTTE_CLIENT_KEY || !LOTTE_JOBCUSTCD) {
    return { success: false, error: 'LOTTE 취소 환경변수 미설정' };
  }

  const payload = {
    snd_list: [{
      jobCustCd: LOTTE_JOBCUSTCD,
      invNo: invoiceNumber.replace(/\D/g, ''),
      canCd: '01',
      canDtlCd: '19',
      canRmk: '자동 집하취소(TMS)',
    }],
  };

  try {
    const r = await httpPost(LOTTE_CANCEL_API_URL, alpsHeaders(), payload, 1);
    const rtnList = r.json.rtn_list;
    const first = (Array.isArray(rtnList) ? rtnList[0] : {}) as Record<string, unknown>;
    const cd = String(first.rtnCd || '').toUpperCase();
    if (cd === 'S') return { success: true };
    return { success: false, error: `취소 실패: ${first.rtnMsg || cd}` };
  } catch (err) {
    return { success: false, error: String(err) };
  }
}

/* ── 집하(수거) 판정 (109, 2026-07-12) ──
   롯데 기사님이 방문 수거하며 스캔하면 godsStatCd '10'(집하)이 찍힌다.

   ALPS 코드표: 02 출력 · 10 집하 · 12 운송장등록 · 20/21 구간이동 · 40 배달전 · 41 배달완료 · 45 인수자등록 · 09 반품취소

   ⚠️ '10' 만 정확히 매칭하면 안 된다 — 크론이 집하~간선 사이 구간을 놓치고 조회하면 영영 출고 처리가 안 된다.
      그래서 "집하 이후 단계(20 이상)"도 수거된 것으로 본다.
   🚨 그렇다고 ">= 10" 으로 잡으면 **12(운송장등록)가 걸린다.** 12 는 우리가 송장을 발급하는 순간 찍히는 코드라,
      송장만 만들고 아직 기사님이 오지도 않은 건이 전부 출고 처리 + 알림톡 발송되는 사고가 난다. (테스트로 발견) */
function isPickedUpCode(code: string): boolean {
  if (code === '09') return false;              // 반품취소
  const n = Number(code);
  if (!Number.isFinite(n)) return false;
  return n === 10 || n >= 20;                   // 10=집하 / 20↑=간선·배달 (02 출력, 12 운송장등록은 제외)
}

/* 스캔 시각 추출 — ALPS tracking 레코드의 날짜 필드명이 문서상 불확실해 후보 키를 순차 탐색.
   찾지 못하면 undefined 를 돌려주고 호출부가 '감지 시각(now)'으로 대체한다(기능 영향 없음). */
const SCAN_DATE_KEYS = ['scanDtm', 'scnDttm', 'workDtm', 'regDtm', 'procDtm', 'scanDt', 'workDt', 'godsStatDtm'];
function extractScanAt(rec: Record<string, unknown>): string | undefined {
  for (const k of SCAN_DATE_KEYS) {
    const raw = rec[k];
    if (!raw) continue;
    const s = String(raw).replace(/\D/g, '');          // 'yyyyMMddHHmmss' / 'yyyy-MM-dd HH:mm:ss' 모두 흡수
    if (s.length < 8) continue;
    const iso = `${s.slice(0, 4)}-${s.slice(4, 6)}-${s.slice(6, 8)}T${s.slice(8, 10) || '00'}:${s.slice(10, 12) || '00'}:${s.slice(12, 14) || '00'}+09:00`; // ALPS = KST
    const d = new Date(iso);
    if (!isNaN(d.getTime())) return d.toISOString();
  }
  return undefined;
}

/* ── 배송 상태 조회 ──
   ⚠️ state/detail 은 기존 그대로 (호출부 3곳 무영향). pickedUp/pickedUpAt 은 109 신규 추가 필드. */
export async function queryTrackingStatus(invoiceNumber: string): Promise<{
  state: 'ACTIVE' | 'CANCELLED' | 'DELIVERED' | 'NOT_FOUND';
  detail?: string;
  pickedUp?: boolean;        // 109: 기사 수거 스캔 이후로 진행됨
  pickedUpAt?: string;       // 109: 집하 스캔 시각 (파싱 실패 시 undefined)
  trackingKeys?: string[];   // 109: ?debug=1 진단용 — 실제 필드명 확인 후 제거 가능
}> {
  if (!LOTTE_TRACK_API_URL || !LOTTE_CLIENT_KEY || !LOTTE_JOBCUSTCD) {
    return { state: 'NOT_FOUND', detail: 'LOTTE 환경변수 미설정' };
  }
  const url = `${LOTTE_TRACK_API_URL}?invNo=${encodeURIComponent(invoiceNumber)}&jobCustCd=${encodeURIComponent(LOTTE_JOBCUSTCD)}`;
  try {
    const res = await fetch(url, {
      method: 'GET',
      headers: { Authorization: `IgtAK ${LOTTE_CLIENT_KEY}`, Accept: 'application/json' },
    });
    let json: Record<string, unknown>;
    try { json = await res.json(); } catch { json = {}; }
    if (!res.ok) return { state: 'NOT_FOUND', detail: `HTTP ${res.status}` };

    const code = String(json.code || '').toUpperCase();
    if (code !== 'S') return { state: 'NOT_FOUND', detail: `code=${code || 'empty'} msg=${String(json.message || '').slice(0, 100)}` };

    const tracking = Array.isArray(json.tracking) ? json.tracking : [];
    if (tracking.some((x: Record<string, unknown>) => String(x.godsStatCd || '') === '09')) return { state: 'CANCELLED', pickedUp: false };

    // 109: 집하(수거) 판정 — 상태 분기와 독립적으로 계산해 모든 return 에 실어 보낸다
    const pickupRec = tracking.find((x: Record<string, unknown>) => String(x.godsStatCd || '') === '10')
      || tracking.find((x: Record<string, unknown>) => isPickedUpCode(String(x.godsStatCd || '')));
    const pickedUp = !!pickupRec;
    const pickedUpAt = pickupRec ? extractScanAt(pickupRec as Record<string, unknown>) : undefined;
    const trackingKeys = tracking.length > 0 ? Object.keys(tracking[0] as Record<string, unknown>) : [];

    // 배달완료 감지 (2026-05-24 버그 수정): '91' → '41'(배달완료) 또는 '45'(인수자등록)
    // 실제 ALPS API 응답: 41 = 기사 배달완료, 45 = 고객 인수확인. 옛 코드 '91' 은 영원히 매칭 안 됐음.
    if (tracking.some((x: Record<string, unknown>) => {
      const code = String(x.godsStatCd || '');
      return code === '41' || code === '45';
    })) return { state: 'DELIVERED', pickedUp: true, pickedUpAt, trackingKeys };  // 배달완료면 집하는 당연히 지났음

    const result = Array.isArray(json.result) ? json.result : [];
    if (tracking.length === 0 && result.length === 0) return { state: 'NOT_FOUND', detail: 'tracking empty', pickedUp: false };
    return { state: 'ACTIVE', pickedUp, pickedUpAt, trackingKeys };
  } catch (e) {
    return { state: 'NOT_FOUND', detail: `exception: ${String(e).slice(0, 150)}` };
  }
}
