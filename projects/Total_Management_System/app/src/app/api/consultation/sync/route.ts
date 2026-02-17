import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { upsertConsultation, batchUpsertConsultations } from '@/lib/consultation/sync';
import type { GasPushPayload } from '@/lib/consultation/sync';

/**
 * POST /api/consultation/sync
 * 두 가지 인증 경로:
 * 1) GAS 스크립트 → x-sync-key 헤더 (CRON_SECRET)
 * 2) TMS 관리자 UI → Supabase 세션 인증
 */
export async function POST(req: NextRequest) {
  try {
    // 인증 방식 1: API Key (GAS 스크립트용)
    const syncKey = req.headers.get('x-sync-key');
    const isGasAuth = syncKey === process.env.CRON_SECRET;

    // 인증 방식 2: 세션 (TMS 관리자 UI용)
    let isSessionAuth = false;
    if (!isGasAuth) {
      const supabase = await createServerSupabaseClient();
      const { data: { user } } = await supabase.auth.getUser();
      isSessionAuth = !!user;
    }

    if (!isGasAuth && !isSessionAuth) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    // body 파싱 (UI에서 빈 body로 호출할 수 있음)
    let body: Record<string, unknown> = {};
    try {
      body = await req.json();
    } catch {
      // 빈 body → UI 새로고침 요청
    }

    // UI에서 빈 body 호출 시: 캐시 갱신용 (데이터는 React Query invalidation으로 처리)
    if (!body.batch && !body.uniqueId && !body.name) {
      return NextResponse.json({ success: true, synced: 0, message: '동기화 확인 완료' });
    }

    // 배치 모드 (GAS syncAllToSupabase)
    if (body.batch && Array.isArray(body.batch)) {
      const result = await batchUpsertConsultations(body.batch as GasPushPayload[]);
      return NextResponse.json(result);
    }

    // 단건 모드 (GAS pushToSupabase_)
    const result = await upsertConsultation(body as unknown as GasPushPayload);
    if (!result.success) {
      return NextResponse.json(result, { status: 400 });
    }
    return NextResponse.json(result);
  } catch (err) {
    console.error('[consultation-sync] Push 수신 실패:', err);
    return NextResponse.json(
      { error: String(err) },
      { status: 500 }
    );
  }
}
