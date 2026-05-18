import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * POST /api/serials/swap
 * body: { serial_a_id: string, serial_b_id: string }
 *
 * 두 시리얼의 판매 연결을 양방향으로 동시 교환 (Phase B).
 * 트랜잭션 안전: PL/pgSQL RPC `swap_serials()` 가 SELECT FOR UPDATE + 가드 5겹 검증.
 * 이력 추적: Phase C 트리거가 두 UPDATE 를 audit_log 에 자동 캡처.
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const { serial_a_id, serial_b_id } = body || {};

    if (!serial_a_id || !serial_b_id) {
      return NextResponse.json(
        { error: 'serial_a_id 와 serial_b_id 가 모두 필요합니다' },
        { status: 400 }
      );
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // RPC 호출 — 트랜잭션 + 가드 5겹은 PL/pgSQL 안에서
    const { data, error } = await db.rpc('swap_serials', {
      p_serial_a: serial_a_id,
      p_serial_b: serial_b_id,
    });

    if (error) {
      // RPC RAISE EXCEPTION 메시지 한국어 그대로 노출 (가드 1~5 메시지)
      const message = error.message || error.hint || '시리얼 교환 실패';
      // PostgreSQL 에러 코드 분류
      const statusCode = (() => {
        const m = String(message);
        if (m.includes('SAME_SERIAL') || m.includes('같은 시리얼')) return 400;
        if (m.includes('NOT_FOUND') || m.includes('찾을 수 없')) return 404;
        if (m.includes('NOT_SOLD') || m.includes('판매완료 상태')) return 409;
        if (m.includes('NO_SALE') || m.includes('판매에 연결')) return 409;
        if (m.includes('SAME_SALE') || m.includes('같은 판매 안')) return 400;
        if (m.includes('PRODUCT_MISMATCH') || m.includes('같은 제품')) return 409;
        return 500;
      })();
      return NextResponse.json({ error: message }, { status: statusCode });
    }

    return NextResponse.json(data);
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : JSON.stringify(err);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
