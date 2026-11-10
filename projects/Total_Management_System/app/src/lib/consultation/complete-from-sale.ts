/**
 * 판매 연결 → 상담 자동 완료 (2026-11-10)
 *
 * 판매가 상담에 연결되면(판매 등록 시 source_consultation_id / 판매 화면의 '상담 연결')
 * 그 상담은 끝난 것 → 사장님이 [상담 완료]를 한 번 더 누르지 않도록 서버가 완료로 옮긴다.
 *
 * - 전이 규칙 SSOT = transitions.ts (확정/진행중 → 완료만. 신규·제안중 등은 건드리지 않음)
 * - 이력(consultation_history) 기록 + 구글 캘린더 갱신 = PATCH /api/consultation/[id] 와 동일
 * - 완료 알림톡은 없음(2026-04-30 제거, 후기 요청은 판매 쪽 단일 진입점) → 고객 발송 없음
 * - 실패해도 판매는 막지 않는다(호출부에서 로그만)
 */
import { after } from 'next/server';
import { isValidTransition } from '@/lib/consultation/transitions';
import { syncConsultationToCalendar } from '@/lib/google/calendar-sync';
import type { ConsultationStatus, ConsultationType } from '@/lib/supabase/types';

export async function completeConsultationFromSale(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  db: any,
  opts: { consultationId: string; saleNumber?: string | null; userId: string },
): Promise<boolean> {
  const { consultationId, saleNumber, userId } = opts;

  const { data: current, error } = await db
    .from('consultations')
    .select('status, consultation_type')
    .eq('id', consultationId)
    .single();
  if (error || !current) return false;

  if (!isValidTransition(
    current.consultation_type as ConsultationType,
    current.status as ConsultationStatus,
    'completed',
  )) return false;

  // 조건부 업데이트 — 그 사이 다른 화면에서 상태가 바뀌었으면 덮어쓰지 않음(이력 중복 방지)
  const { data: updated, error: updErr } = await db
    .from('consultations')
    .update({ status: 'completed', hold_reason: null })
    .eq('id', consultationId)
    .eq('status', current.status)
    .select('id');
  if (updErr) throw updErr;
  if (!updated || updated.length === 0) return false;

  await db.from('consultation_history').insert({
    consultation_id: consultationId,
    from_status: current.status,
    to_status: 'completed',
    changed_by: userId,
    note: `판매 등록으로 자동 완료${saleNumber ? ` (${saleNumber})` : ''}`,
  });

  // 캘린더 갱신은 응답 뒤 백그라운드 (내부에서 예외 모두 캐치)
  after(() => syncConsultationToCalendar(consultationId));
  return true;
}
