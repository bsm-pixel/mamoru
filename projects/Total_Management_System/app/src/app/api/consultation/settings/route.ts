/**
 * /api/consultation/settings
 * 사장님 측 consultation_settings 테이블 관리 API.
 *
 * 정책:
 *   - 관리자 전용 (auth 필수)
 *   - GET → consultation_settings 전체 조회
 *   - PATCH { disabled_weekdays?, start_hour?, end_hour?, ... } → 부분 업데이트
 *
 * 옵션 C (078): 휴무 요일 + 특별 휴무일을 달력 관리 화면에서 통합 관리.
 *   설정 → 상담 설정에서 '휴무 요일'과 '특별 휴무일' UI 제거 (이 API + blackouts API로 이전).
 *
 * 사장님 룰 (memory/feedback_consultation_blackout_rule.md):
 *   여기서 변경되는 disabled_weekdays/closed_dates는 고객 셀프 예약 폼에만 영향.
 *   사장님 측 흐름(admin-create, suggest)은 항상 유동.
 */

import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';

const PATCHABLE_FIELDS = [
  'start_hour',
  'end_hour',
  'duration_min',
  'step_min',
  'disabled_weekdays',
  'field_buffer_before',
  'field_buffer_after',
] as const;

export async function GET() {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { data, error } = await (db as any)
    .from('consultation_settings')
    .select('*')
    .eq('id', 'default')
    .single();

  if (error) return NextResponse.json({ error: error.message }, { status: 500 });
  return NextResponse.json(data || {});
}

export async function PATCH(req: NextRequest) {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const body = await req.json().catch(() => ({}));
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const updates: Record<string, any> = {};
  for (const k of PATCHABLE_FIELDS) {
    if (body[k] !== undefined) updates[k] = body[k];
  }
  if (Object.keys(updates).length === 0) {
    return NextResponse.json({ error: '변경할 필드가 없습니다' }, { status: 400 });
  }

  // disabled_weekdays 검증: 0~6 정수 배열만
  if (updates.disabled_weekdays !== undefined) {
    if (!Array.isArray(updates.disabled_weekdays)
        || updates.disabled_weekdays.some((n: unknown) => typeof n !== 'number' || n < 0 || n > 6)) {
      return NextResponse.json({ error: 'disabled_weekdays는 0~6 정수 배열' }, { status: 400 });
    }
  }

  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { error } = await (db as any)
    .from('consultation_settings')
    .upsert({ id: 'default', ...updates, updated_at: new Date().toISOString() });

  if (error) return NextResponse.json({ error: error.message }, { status: 500 });
  return NextResponse.json({ ok: true, updated: Object.keys(updates) });
}
