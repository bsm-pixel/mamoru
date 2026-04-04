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

/* ── 송장 발급 (GAS lotteBuildSnd_ + lotteSend_ 100% 동일) ── */
export async function bookShipment(order: {
  invoiceNumber: string;
  receiverName: string;
  receiverTel: string;
  receiverZip: string;
  receiverAddr: string;
  goodsName?: string;
  deliveryMessage?: string;
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
      ustRtgSctCd: '01',
      ordSct:      '3',
      fareSctCd:   LOTTE_FARE,
      ordNo:       ordNo,
      invNo:       order.invoiceNumber,

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
      return { success: false, invoiceNumber: order.invoiceNumber, error: `ALPS HTTP ${r.code}` };
    }

    // GAS lotteSend_ 동일: rtn_list[0].rtnCd === 'S'
    const rtnList = r.json.rtn_list;
    const first = (Array.isArray(rtnList) ? rtnList[0] : {}) as Record<string, unknown>;
    const rtnCd = String(first.rtnCd || '').toUpperCase();

    if (rtnCd === 'S') {
      return { success: true, invoiceNumber: order.invoiceNumber };
    }

    return { success: false, invoiceNumber: order.invoiceNumber, error: String(first.rtnMsg || rtnCd) };
  } catch (err) {
    return { success: false, invoiceNumber: order.invoiceNumber, error: err instanceof Error ? err.message : String(err) };
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

/* ── 배송 상태 조회 ── */
export async function queryTrackingStatus(invoiceNumber: string): Promise<{
  state: 'ACTIVE' | 'CANCELLED' | 'DELIVERED' | 'NOT_FOUND';
  detail?: string;
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
    if (!res.ok) return { state: 'NOT_FOUND' };

    const code = String(json.code || '').toUpperCase();
    if (code !== 'S') return { state: 'NOT_FOUND' };

    const tracking = Array.isArray(json.tracking) ? json.tracking : [];
    if (tracking.some((x: Record<string, unknown>) => String(x.godsStatCd || '') === '09')) return { state: 'CANCELLED' };
    if (tracking.some((x: Record<string, unknown>) => String(x.godsStatCd || '') === '91')) return { state: 'DELIVERED' };

    const result = Array.isArray(json.result) ? json.result : [];
    if (tracking.length === 0 && result.length === 0) return { state: 'NOT_FOUND' };
    return { state: 'ACTIVE' };
  } catch {
    return { state: 'NOT_FOUND' };
  }
}
