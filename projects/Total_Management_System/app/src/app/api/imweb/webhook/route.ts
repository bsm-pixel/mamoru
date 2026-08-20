import { NextRequest, NextResponse, after } from 'next/server';
import { syncSingleOrder } from '@/lib/imweb/sync';

/**
 * POST /api/imweb/webhook?key=<시크릿> — 아임웹 주문 웹훅 수신 (실시간)
 *
 * 아임웹 개발자센터에 등록한 웹훅(ORDER_CREATE / ORDER_DEPOSIT_COMPLETE 등)이
 * 이 URL로 이벤트를 쏘면, 해당 주문 1건만 아임웹 API로 재조회해 TMS에 반영한다.
 * 처리는 기존 upsertOrder(=syncSingleOrder) 재사용 → imweb_order_no 유니크로
 * 크론(하루 1회 안전망)과 겹쳐도 멱등(중복 없음).
 *
 * ⚠️ 보안: 아임웹 웹훅은 서명(signature)을 제공하지 않는다.
 *   1차 — URL 쿼리의 시크릿(key = env IMWEB_WEBHOOK_SECRET)으로 위조 요청 차단
 *   2차 — 받은 주문번호를 아임웹 API로 재조회(syncSingleOrder)하여 실재 확인
 *
 * ⚠️ 페이로드 형식 미확정(2026-08 기준): 아임웹 공식 문서에 필드명 명시 없음.
 *   → 원본을 로깅(Vercel 로그)하고 주문번호를 후보 필드에서 방어적으로 추출한다.
 *   실제 "테스트 보내기" 페이로드 확인 후 extractOrderNo를 확정 필드로 좁힐 것.
 */
export async function POST(request: NextRequest) {
  // 1) 시크릿 검증 (미설정이면 fail-closed → 401)
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
  // 🔎 실제 페이로드 구조 확인용 로그 (매핑 확정 후에도 진단에 유용)
  console.log('[imweb/webhook] 수신 payload:', JSON.stringify(payload));

  // 3) 주문번호 추출 (필드명 미확정 → 후보 다수 시도)
  const orderNo = extractOrderNo(payload);
  if (!orderNo) {
    // 못 찾아도 200 응답 (아임웹 재시도 폭주 방지) — 로그로 추적해 필드 확정
    console.warn('[imweb/webhook] 주문번호 추출 실패 — payload 구조 확인 필요');
    return NextResponse.json({ ok: true, note: 'order_no not found' });
  }

  // 4) 응답은 즉시, 처리는 after()로 완주 보장 (fire-and-forget 누락 방지)
  const run = async () => {
    const r = await syncSingleOrder(orderNo);
    console.log('[imweb/webhook] 동기화 결과:', r);
  };
  try {
    after(run);
  } catch {
    await run().catch(() => {});
  }

  return NextResponse.json({ ok: true, order_no: orderNo });
}

/** 아임웹 웹훅 payload에서 주문번호를 방어적으로 추출 (실제 필드 확정 전 임시) */
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
