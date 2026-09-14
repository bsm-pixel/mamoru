import { NextRequest, NextResponse, after } from 'next/server';
import { syncSingleOrder } from '@/lib/imweb/sync';
import { createServiceClient } from '@/lib/supabase/server';

/**
 * POST /api/imweb/webhook?key=<시크릿> — 아임웹 주문 웹훅 수신 (실시간)
 *
 * 1) 들어온 웹훅은 이벤트 종류와 무관하게 **원본 그대로 기록**한다 (imweb_webhook_events, 마이그 148)
 *    → "실제로 들어오는지 / 어떤 값이 오는지" 실측 근거 + 알림톡 미발송 문의 시 1차 진단
 * 2) 주문생성·입금완료(SYNC_EVENTS)는 기존대로 해당 주문 1건을 아임웹 API로 재조회해 TMS에 반영
 *    (syncSingleOrder = upsertOrder 재사용 → imweb_order_no 유니크로 크론과 겹쳐도 멱등)
 * 3) 취소·반품·교환·거절 등 그 외 이벤트는 **기록만** 한다 (2026-09-14) — 실제 페이로드 실측 전까지
 *    재고·주문상태에 영향 주지 않음. 알림톡 연결은 실측 후 별도 단계.
 *
 * ⚠️ 보안: 아임웹 웹훅은 서명(signature)을 제공하지 않는다.
 *   1차 — URL 쿼리의 시크릿(key = env IMWEB_WEBHOOK_SECRET)으로 위조 요청 차단 (헤더는 보지 않음)
 *   2차 — 동기화 대상은 받은 주문번호를 아임웹 API로 재조회(syncSingleOrder)하여 실재 확인
 *
 * 페이로드 형식(2026-09-11 테스트 전송 실측): 최상위 `eventType` + `data.orderNo`(number).
 *   취소 요청 예: data.section.sectionItems[].productInfo.prodName / qty, data.section.cancelInfo
 */

/** 기존에 TMS 로 받아 주문 동기화하던 이벤트 — 동작 유지 */
const SYNC_EVENTS = new Set(['ORDER_CREATE', 'ORDER_DEPOSIT_COMPLETE']);

/** 기록에서 제외할 헤더 (민감정보) — Vercel 내부 헤더(x-vercel-*)엔 단기 OIDC 토큰·프록시 서명이 섞여 있어 통째로 제외 (2026-09-14 실측 발견) */
const SKIP_HEADERS = new Set(['cookie', 'authorization']);
const isSkippedHeader = (k: string) => SKIP_HEADERS.has(k) || k.startsWith('x-vercel-');

export async function POST(request: NextRequest) {
  // 1) 시크릿 검증 (미설정이면 fail-closed → 401). 인증 실패 요청은 기록하지 않음(스팸 방지)
  const secret = process.env.IMWEB_WEBHOOK_SECRET;
  const key = request.nextUrl.searchParams.get('key');
  if (!secret || key !== secret) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  // 2) 페이로드 파싱 (JSON 우선, 아니면 텍스트로 보존)
  let payload: unknown = null;
  try {
    payload = await request.json();
  } catch {
    try {
      const text = await request.text();
      payload = text ? { _raw_text: text } : null;
    } catch {
      /* noop */
    }
  }
  // 🔎 Vercel 로그에도 원본 유지 (DB 기록 실패 시 백업 진단 경로)
  console.log('[imweb/webhook] 수신 payload:', JSON.stringify(payload));

  const eventType = extractEventType(payload);
  const orderNo = extractOrderNo(payload);
  // eventType 이 없는 구형/미확인 페이로드는 기존 동작(동기화) 유지
  const shouldSync = !!orderNo && (!eventType || SYNC_EVENTS.has(eventType));
  const action = !orderNo ? 'no_order_no' : shouldSync ? 'sync' : 'logged';

  // 3) 원본 기록 — 실패해도 기존 주문 동기화를 막지 않는다
  const headers: Record<string, string> = {};
  request.headers.forEach((v, k) => { if (!isSkippedHeader(k.toLowerCase())) headers[k] = v; });
  const eventId = await recordEvent({ event_type: eventType, order_no: orderNo, action, payload, headers });

  if (!orderNo) {
    // 못 찾아도 200 응답 (아임웹 재시도 폭주 방지) — 기록으로 추적
    console.warn('[imweb/webhook] 주문번호 추출 실패 — payload 구조 확인 필요');
    return NextResponse.json({ ok: true, note: 'order_no not found' });
  }

  if (!shouldSync) {
    return NextResponse.json({ ok: true, event_type: eventType, order_no: orderNo, logged: true });
  }

  // 4) 응답은 즉시, 처리는 after()로 완주 보장 (fire-and-forget 누락 방지)
  const run = async () => {
    try {
      const r = await syncSingleOrder(orderNo);
      console.log('[imweb/webhook] 동기화 결과:', r);
      await markProcessed(eventId, { process_result: r ?? null });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      console.error('[imweb/webhook] 동기화 실패:', msg);
      await markProcessed(eventId, { process_error: msg });
    }
  };
  try {
    after(run);
  } catch {
    await run().catch(() => {});
  }

  return NextResponse.json({ ok: true, order_no: orderNo });
}

/** imweb_webhook_events 에 원본 기록. 테이블 미생성·DB 오류여도 null 반환(삼킴) */
async function recordEvent(row: {
  event_type: string | null;
  order_no: string | null;
  action: string;
  payload: unknown;
  headers: Record<string, string>;
}): Promise<number | null> {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const { data, error } = await db.from('imweb_webhook_events').insert(row).select('id').single();
    if (error) {
      console.error('[imweb/webhook] 수신 기록 실패:', error.message);
      return null;
    }
    return data?.id ?? null;
  } catch (e) {
    console.error('[imweb/webhook] 수신 기록 예외:', e instanceof Error ? e.message : String(e));
    return null;
  }
}

/** 동기화 결과를 기록 행에 반영 (기록 행이 없으면 no-op) */
async function markProcessed(id: number | null, patch: { process_result?: unknown; process_error?: string }) {
  if (id == null) return;
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    await db.from('imweb_webhook_events').update({ ...patch, processed_at: new Date().toISOString() }).eq('id', id);
  } catch {
    /* 기록 보조 정보 — 실패해도 무시 */
  }
}

/** 최상위 eventType (예: ORDER_CANCEL_REQUEST). 없으면 null */
function extractEventType(payload: unknown): string | null {
  if (!payload || typeof payload !== 'object') return null;
  const t = (payload as Record<string, unknown>).eventType;
  return typeof t === 'string' && t.trim() ? t.trim() : null;
}

/** 아임웹 웹훅 payload에서 주문번호를 방어적으로 추출 (확정 필드 data.orderNo 포함 후보 순회) */
function extractOrderNo(payload: unknown): string | null {
  if (!payload || typeof payload !== 'object') return null;
  const p = payload as Record<string, unknown>;
  const nested = (k: string): Record<string, unknown> | undefined =>
    (p[k] && typeof p[k] === 'object') ? (p[k] as Record<string, unknown>) : undefined;
  const data = nested('data');
  const body = nested('body');
  const inner = nested('payload');
  const candidates: unknown[] = [
    p.order_no, p.orderNo, p.order_code, p.orderCode,
    data?.order_no, data?.orderNo, data?.order_code, data?.orderCode,
    body?.order_no, body?.orderNo,
    inner?.order_no, inner?.orderNo,
  ];
  for (const c of candidates) {
    if (c != null && String(c).trim() !== '') return String(c);
  }
  return null;
}
