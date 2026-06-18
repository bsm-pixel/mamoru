import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { recalcOutstanding } from '@/lib/outstanding';

/**
 * POST /api/customers/merge — 고객 병합
 * body: { primaryId, victimIds[] }
 * 흡수 대상(victims)의 모든 거래를 주 고객(primary)으로 이관(merge_customers RPC, 단일 트랜잭션)
 * 후 주 고객 미수금 재계산.
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { primaryId, victimIds } = (await req.json()) as { primaryId?: string; victimIds?: string[] };
    if (!primaryId || !Array.isArray(victimIds) || victimIds.length === 0) {
      return NextResponse.json({ error: '주 고객과 흡수 대상을 선택해주세요' }, { status: 400 });
    }
    if (victimIds.includes(primaryId)) {
      return NextResponse.json({ error: '주 고객은 흡수 대상에 포함될 수 없습니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data, error } = await db.rpc('merge_customers', { p_primary: primaryId, p_victims: victimIds });
    if (error) throw error;

    // 주 고객 미수금 재계산 (단일 출처 — RPC 와 중복 정의 회피)
    await recalcOutstanding(db, primaryId);

    return NextResponse.json({ ok: true, result: data });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
