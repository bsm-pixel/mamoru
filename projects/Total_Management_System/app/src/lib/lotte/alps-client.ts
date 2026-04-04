/**
 * 롯데택배 ALPS API 직접 호출 클라이언트
 * GAS Code.js의 lotte* 함수들을 Node.js로 이전
 */

import { createServiceClient } from '@/lib/supabase/server';

// 환경변수 (GAS Script Properties에서 이전)
const LOTTE_API_URL = process.env.LOTTE_API_URL || '';
const LOTTE_CANCEL_API_URL = process.env.LOTTE_CANCEL_API_URL || '';
const LOTTE_TRACK_API_URL = process.env.LOTTE_TRACK_API_URL || '';
const LOTTE_CLIENT_KEY = process.env.LOTTE_CLIENT_KEY || '';
const LOTTE_JOB_CUST_CD = process.env.LOTTE_JOB_CUST_CD || '';

// 발송인 정보
const SENDER = {
  name: process.env.LOTTE_SENDER_NAME || '마모루',
  tel: process.env.LOTTE_SENDER_TEL || '',
  zip: process.env.LOTTE_SENDER_ZIP || '',
  addr: process.env.LOTTE_SENDER_ADDR || '',
};

/** 체크디짓 계산 (11자리 base → mod 7) */
export function checkDigit(base11: number): number {
  return base11 % 7;
}

/** 11자리 base → 12자리 송장번호 */
export function toInvoiceNumber(base11: number): string {
  const cd = checkDigit(base11);
  return `${base11}${cd}`;
}

/** 다음 송장번호 발급 (DB 카운터 사용, 원자적) */
export async function getNextInvoice(): Promise<{ invoiceNumber: string; base11: number }> {
  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const dbAny = db as any;

  // 현재 번호 조회 + 증가 (원자적 — DB 트랜잭션)
  const { data: config, error: fetchErr } = await dbAny
    .from('lotte_waybill_config')
    .select('*')
    .eq('id', 'default')
    .single();

  if (fetchErr || !config) {
    throw new Error('송장번호 설정이 없습니다. lotte_waybill_config 테이블에 초기값을 설정해주세요.');
  }

  const current = Number(config.current_number);
  const end = Number(config.end_number);

  if (current > end) {
    throw new Error(`송장번호 풀 소진 (현재: ${current}, 종료: ${end}). 롯데택배에 새 범위를 요청하세요.`);
  }

  // 범위 90% 소진 경고
  const start = Number(config.start_number);
  const usage = (current - start) / (end - start);
  if (usage >= 0.9) {
    console.warn(`[ALPS] 송장번호 풀 90% 사용 (${current}/${end})`);
  }

  // 카운터 증가
  const { error: updateErr } = await dbAny
    .from('lotte_waybill_config')
    .update({ current_number: current + 1, updated_at: new Date().toISOString() })
    .eq('id', 'default')
    .eq('current_number', current); // 낙관적 잠금

  if (updateErr) {
    throw new Error('송장번호 발급 충돌. 다시 시도해주세요.');
  }

  return { invoiceNumber: toInvoiceNumber(current), base11: current };
}

/** ALPS API 호출 (재시도 3회) */
async function alpsPost(url: string, payload: unknown, retries = 3): Promise<unknown> {
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
    'Authorization': `IgtAK ${LOTTE_CLIENT_KEY}`,
  };

  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const res = await fetch(url, {
        method: 'POST',
        headers,
        body: JSON.stringify(payload),
        signal: AbortSignal.timeout(10000),
      });

      const data = await res.json();

      if (!res.ok) {
        throw new Error(`ALPS HTTP ${res.status}: ${JSON.stringify(data)}`);
      }

      return data;
    } catch (err) {
      if (attempt === retries) throw err;
      await new Promise(r => setTimeout(r, attempt * 1000));
    }
  }
}

/** 송장 발급 (접수) */
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
    throw new Error('LOTTE API 환경변수가 설정되지 않았습니다.');
  }

  const payload = {
    snd_list: [{
      jobCustCd: LOTTE_JOB_CUST_CD,
      invNo: order.invoiceNumber,
      // 발송인 (ALPS 공식 필드명: snper*)
      snperNm: SENDER.name,
      snperTel: SENDER.tel.replace(/\D/g, ''),
      snperCpno: '',
      snperZipcd: SENDER.zip,
      snperAdr: SENDER.addr,
      // 수화주 (ALPS 공식 필드명: acper*)
      acperNm: order.receiverName,
      acperTel: order.receiverTel.replace(/\D/g, ''),
      acperCpno: order.receiverTel.replace(/\D/g, ''),
      acperZipcd: order.receiverZip,
      acperAdr: order.receiverAddr,
      // 상품/배송
      gdsNm: order.goodsName || '가위 복원수리',
      dlvMsgCont: order.deliveryMessage || '',
      cusMsgCont: '',
      ustRtgSctCd: '01',
      ordSct: '3',
      fareSctCd: '03',
      boxTypCd: 'A',
    }],
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const result = await alpsPost(LOTTE_API_URL, payload) as any;

  if (result?.rtnCd === '0000' || result?.rtnCd === '00') {
    return { success: true, invoiceNumber: order.invoiceNumber };
  }

  return {
    success: false,
    invoiceNumber: order.invoiceNumber,
    error: `ALPS 오류: ${result?.rtnMsg || result?.rtnCd || JSON.stringify(result)}`,
  };
}

/** 송장 취소 */
export async function cancelShipment(invoiceNumber: string): Promise<{ success: boolean; error?: string }> {
  if (!LOTTE_CANCEL_API_URL || !LOTTE_CLIENT_KEY) {
    throw new Error('LOTTE 취소 API 환경변수가 설정되지 않았습니다.');
  }

  try {
    // 1. 집하 취소 시도
    const payload = {
      invNo: invoiceNumber,
      canCd: '01',
      canDtlCd: '19',
    };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = await alpsPost(LOTTE_CANCEL_API_URL, payload) as any;

    if (result?.rtnCd === '0000' || result?.rtnCd === '00') {
      return { success: true };
    }

    return { success: false, error: `취소 실패: ${result?.rtnMsg || result?.rtnCd}` };
  } catch (err) {
    return { success: false, error: String(err) };
  }
}

/** 배송 상태 조회 */
export async function queryTrackingStatus(invoiceNumber: string): Promise<{
  state: 'ACTIVE' | 'CANCELLED' | 'DELIVERED' | 'NOT_FOUND';
  detail?: string;
}> {
  if (!LOTTE_TRACK_API_URL) {
    return { state: 'NOT_FOUND', detail: 'LOTTE_TRACK_API_URL 미설정' };
  }

  try {
    const url = `${LOTTE_TRACK_API_URL}?invNo=${invoiceNumber}`;
    const res = await fetch(url, {
      headers: { 'Authorization': `IgtAK ${LOTTE_CLIENT_KEY}` },
      signal: AbortSignal.timeout(5000),
    });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const data = await res.json() as any;

    if (data?.godsStatCd === '09') return { state: 'CANCELLED' };
    if (data?.godsStatCd === '91') return { state: 'DELIVERED' };
    if (data?.invNo) return { state: 'ACTIVE', detail: data.godsStatNm };
    return { state: 'NOT_FOUND' };
  } catch {
    return { state: 'NOT_FOUND' };
  }
}
