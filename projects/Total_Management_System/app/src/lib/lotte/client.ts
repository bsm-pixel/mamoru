/**
 * 롯데택배 ALPS API 클라이언트
 * GAS Code.gs에서 포팅
 */

import type {
  LotteConfig,
  LotteBookRequest,
  LotteBookResult,
  LotteCancelResult,
  LotteTrackResult,
} from './types';
import { createServiceClient } from '@/lib/supabase/server';
import { randomUUID } from 'crypto';

/** 환경변수에서 설정 로드 (lotteConfig_ 포팅) */
function getConfig(): LotteConfig {
  return {
    url: process.env.LOTTE_API_URL || '',
    cancelUrl: process.env.LOTTE_CANCEL_API_URL || '',
    trackingUrl:
      process.env.LOTTE_TRACK_API_URL ||
      'https://apigw.llogis.com:10100/api/pid/cus/714a/custmer-view-tracking',
    clientKey: process.env.LOTTE_CLIENT_KEY || '',
    jobCustCd: process.env.LOTTE_JOBCUSTCD || '',
    sender: {
      name: process.env.LOTTE_SENDER_NAME || '',
      tel: process.env.LOTTE_SENDER_TEL || '',
      zip: process.env.LOTTE_SENDER_ZIP || '',
      addr: process.env.LOTTE_SENDER_ADDR || '',
    },
    fareSctCd: process.env.LOTTE_DEFAULT_FARE || '03',
  };
}

/** ALPS 공통 헤더 */
function alpsHeaders(clientKey: string) {
  return {
    Authorization: `IgtAK ${clientKey}`,
    Accept: 'application/json',
    'Content-Type': 'application/json; charset=utf-8',
    'X-Idempotency-Key': randomUUID(),
    'X-Correlation-Id': randomUUID(),
  };
}

/** 공통 POST (httpPostJson_ 포팅) — 재시도 포함 */
async function httpPost(
  url: string,
  headers: Record<string, string>,
  payload: unknown,
  tries = 3
): Promise<{ ok: boolean; code: number; json: Record<string, unknown> }> {
  let lastErr: Error | null = null;
  for (let i = 0; i < tries; i++) {
    try {
      const res = await fetch(url, {
        method: 'POST',
        headers,
        body: JSON.stringify(payload),
      });
      const text = await res.text();
      let json: Record<string, unknown>;
      try {
        json = JSON.parse(text);
      } catch {
        json = { raw: text };
      }

      if (res.ok) return { ok: true, code: res.status, json };
      if (res.status === 429 || res.status >= 500) {
        await sleep(600 * (i + 1));
        continue;
      }
      return { ok: false, code: res.status, json };
    } catch (e) {
      lastErr = e as Error;
      await sleep(600 * (i + 1));
    }
  }
  throw lastErr || new Error('HTTP_POST_FAILED');
}

function sleep(ms: number) {
  return new Promise((r) => setTimeout(r, ms));
}

/** 체크디짓 계산 (lotteCheckDigit_ 포팅) */
function checkDigit(base11: number): number {
  return base11 % 7;
}

/** base11 → 12자리 송장번호 (lotteToInv_ 포팅) */
function toInvNo(base11: number): string {
  return `${base11}${checkDigit(base11)}`;
}

/** Supabase에서 원자적 송장번호 발급 */
async function nextWaybill(): Promise<string> {
  const supabase = createServiceClient();
  const { data, error } = await supabase.rpc('next_waybill');
  if (error) throw new Error(`송장번호 발급 실패: ${error.message}`);
  return data as string;
}

/** ALPS 페이로드 빌드 (lotteBuildSnd_ 포팅) */
function buildSndPayload(cfg: LotteConfig, req: LotteBookRequest) {
  return {
    snd_list: [
      {
        jobCustCd: cfg.jobCustCd,
        ustRtgSctCd: '01',
        ordSct: '3',
        fareSctCd: cfg.fareSctCd,
        ordNo: req.ordNo || '',
        invNo: req.invNo || '',

        snperNm: cfg.sender.name,
        snperTel: cfg.sender.tel.replace(/\D/g, ''),
        snperCpno: '',
        snperZipcd: cfg.sender.zip,
        snperAdr: cfg.sender.addr,

        acperNm: req.rcvName,
        acperTel: (req.rcvTel || '').replace(/\D/g, ''),
        acperCpno: (req.rcvTel || '').replace(/\D/g, ''),
        acperZipcd: req.rcvZip,
        acperAdr: req.rcvAdr,

        boxTypCd: req.boxTypCd || 'A',
        gdsNm: req.gdsNm || 'A/S 물품',
        dlvMsgCont: req.dlvMsg || '',
        cusMsgCont: '',
        pickReqYmd: (req.pickReqYmd || '').replace(/-/g, ''),
      },
    ],
  };
}

/** 예약 (lotteBookSingle_ 포팅) */
export async function bookSingle(
  order: LotteBookRequest
): Promise<LotteBookResult> {
  const cfg = getConfig();
  if (!cfg.url || !cfg.clientKey) {
    throw new Error('LOTTE_CONFIG_MISSING');
  }

  const invNo = await nextWaybill();
  const payload = buildSndPayload(cfg, { ...order, invNo });

  const r = await httpPost(cfg.url, alpsHeaders(cfg.clientKey), payload, 3);
  if (!r.ok) {
    throw new Error(`LOTTE_HTTP_${r.code}: ${JSON.stringify(r.json).slice(0, 200)}`);
  }

  const rtnList = (r.json as Record<string, unknown>).rtn_list;
  const first = (Array.isArray(rtnList) ? rtnList[0] : {}) as Record<string, unknown>;
  const rtnCd = String(first.rtnCd || '');
  const ok = rtnCd.toUpperCase() === 'S';

  if (!ok) {
    throw new Error(`LOTTE_API_ERROR: ${first.rtnMsg || rtnCd}`);
  }

  return {
    ok: true,
    invNo,
    rtnCd,
    rtnMsg: String(first.rtnMsg || ''),
  };
}

/** 집하취소 (lotteCancelPickup_ 포팅) */
export async function cancelPickup(invNo: string): Promise<LotteCancelResult> {
  invNo = invNo.replace(/\D/g, '');
  if (!invNo) return { success: false, error: 'MISSING_INVNO' };

  const cfg = getConfig();
  if (!cfg.cancelUrl || !cfg.clientKey || !cfg.jobCustCd) {
    return { success: false, error: 'CANCEL_CONFIG_MISSING' };
  }

  const payload = {
    snd_list: [
      {
        jobCustCd: cfg.jobCustCd,
        invNo,
        canCd: '01',
        canDtlCd: '19',
        canRmk: '자동 집하취소(TMS)',
      },
    ],
  };

  const r = await httpPost(cfg.cancelUrl, alpsHeaders(cfg.clientKey), payload, 1);
  const rtnList = (r.json as Record<string, unknown>).rtn_list;
  const first = (Array.isArray(rtnList) ? rtnList[0] : {}) as Record<string, unknown>;
  const cd = String(first.rtnCd || '').toUpperCase();

  if (r.ok && cd === 'S') return { success: true, via: 'pickup_cancel' };

  // 상태 조회로 취소 확인
  const q = await queryStatus(invNo);
  if (q.ok && (q.state === 'CANCELLED' || q.state === 'NOT_FOUND')) {
    return { success: true, via: q.state };
  }

  return { success: false, error: 'PICKUP_CANCEL_FAILED' };
}

/** 상태 조회 (lotteQueryStatus_ 포팅) */
export async function queryStatus(invNo: string): Promise<LotteTrackResult> {
  invNo = invNo.replace(/\D/g, '');
  if (!invNo) return { ok: false, state: 'INVALID_INV', raw: {} };

  const cfg = getConfig();
  if (!cfg.trackingUrl || !cfg.clientKey || !cfg.jobCustCd) {
    return { ok: false, state: 'TRACK_CFG_MISSING', raw: {} };
  }

  const url = `${cfg.trackingUrl}?invNo=${encodeURIComponent(invNo)}&jobCustCd=${encodeURIComponent(cfg.jobCustCd)}`;

  try {
    const res = await fetch(url, {
      method: 'GET',
      headers: {
        Authorization: `IgtAK ${cfg.clientKey}`,
        Accept: 'application/json',
      },
    });

    let json: Record<string, unknown>;
    try {
      json = await res.json();
    } catch {
      json = {};
    }

    if (!res.ok) return { ok: false, state: `HTTP_${res.status}`, raw: json };

    const code = String(json.code || '').toUpperCase();
    if (code !== 'S') return { ok: false, state: 'TRACK_ERR', raw: json };

    const tracking = Array.isArray(json.tracking) ? json.tracking : [];
    const cancelled = tracking.some(
      (x: Record<string, unknown>) => String(x.godsStatCd || '') === '09'
    );
    if (cancelled) return { ok: true, state: 'CANCELLED', raw: json };

    const result = Array.isArray(json.result) ? json.result : [];
    if (tracking.length === 0 && result.length === 0) {
      return { ok: true, state: 'NOT_FOUND', raw: json };
    }

    return { ok: true, state: 'ACTIVE', raw: json };
  } catch (e) {
    return { ok: false, state: 'TRACK_EXC', raw: { error: String(e) } };
  }
}

/** 취소 (lotteCancel_ 포팅) */
export async function cancel(invNo: string): Promise<LotteCancelResult> {
  invNo = invNo.replace(/\D/g, '');
  if (!invNo) return { success: false, error: 'MISSING_INVNO' };

  // 집하취소 먼저 시도
  const pc = await cancelPickup(invNo);
  if (pc.success) return pc;

  // 상태 조회로 확인
  const q = await queryStatus(invNo);
  if (q.ok && (q.state === 'CANCELLED' || q.state === 'NOT_FOUND')) {
    return { success: true, via: q.state };
  }

  return { success: false, error: 'CANCEL_FAILED' };
}
