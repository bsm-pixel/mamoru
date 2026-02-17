import { NextRequest, NextResponse } from 'next/server';
import { upsertConsultation, batchUpsertConsultations } from '@/lib/consultation/sync';
import type { GasPushPayload } from '@/lib/consultation/sync';

/**
 * POST /api/consultation/sync
 * GAS 스크립트에서 호출 — 상담 데이터 Push 수신
 *
 * 단건: { uniqueId, name, phone, ... }
 * 배치: { batch: [{ uniqueId, name, phone, ... }, ...] }
 *
 * 인증: X-Sync-Key 헤더로 CRON_SECRET 확인
 */
export async function POST(req: NextRequest) {
  try {
    // API Key 인증 (GAS에서 호출하므로 세션 인증 대신 키 사용)
    const syncKey = req.headers.get('x-sync-key');
    if (syncKey !== process.env.CRON_SECRET) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const body = await req.json();

    // 배치 모드
    if (body.batch && Array.isArray(body.batch)) {
      const result = await batchUpsertConsultations(body.batch as GasPushPayload[]);
      return NextResponse.json(result);
    }

    // 단건 모드
    const result = await upsertConsultation(body as GasPushPayload);
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
