/**
 * 복원수리 직접방문(당일수리) → Google Calendar 동기화 orchestrator
 *
 * 2026-05-25 Phase 3-B 신규
 * 컨설팅 calendar-sync.ts 패턴 동일 (재사용) — repairs.proceed_type='직접방문' 전용
 *
 * 호출 규칙: fire-and-forget — 복원수리 로직 절대 블록 X
 * 모든 오류 내부 캐치 + 로그만 남김
 *
 * 사용처:
 *   - api/repair/public/submit/route.ts : 접수 시 (직접방문만)
 *   - api/repair/[id]/route.ts          : 시간 변경 / 취소 / 완료
 */

import { createServiceClient } from '@/lib/supabase/server';
import { createCalendarEvent, updateCalendarEvent, deleteCalendarEvent } from './calendar-client';
import { formatRepairToEvent, type RepairForCalendar, type EventFormatSettings } from './event-formatter';

const BASE_URL = process.env.NEXT_PUBLIC_APP_URL || 'https://app-eta-sandy-75.vercel.app';

/** 이벤트 삭제 대상 상태 (취소 = 캘린더에서 제거) */
const DELETE_STATES = new Set(['cancelled']);

/**
 * 복원수리 직접방문 건 → Google Calendar 동기화
 *
 * @param repairId - repairs.id (UUID)
 *
 * 동기화 대상: proceed_type='직접방문' 만
 * 다른 진행방식(방문수거/직접발송)은 일정 캘린더 표시 의미 없음 → 자동 skip
 */
export async function syncRepairToCalendar(repairId: string): Promise<void> {
  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // 현재 복원수리 조회
    const { data: r, error } = await dbAny
      .from('repairs')
      .select('*')
      .eq('id', repairId)
      .single();

    if (error || !r) {
      console.warn('[repair-calendar-sync] repair 조회 실패', repairId, error?.message);
      return;
    }

    // 직접방문만 동기화 대상 (방문수거/직접발송 = skip)
    if (r.proceed_type !== '직접방문') return;

    const currentStatus = r.status as string;
    const eventId = r.google_event_id as string | null;

    // 1) 삭제 대상 (cancelled) — 이벤트 있으면 삭제
    if (DELETE_STATES.has(currentStatus)) {
      if (eventId) {
        await deleteCalendarEvent({ eventId });
        await dbAny
          .from('repairs')
          .update({ google_event_id: null, google_event_updated_at: new Date().toISOString() })
          .eq('id', repairId);
      }
      return;
    }

    // 2) 유지/생성/업데이트 대상 — visit_date/time 필수
    if (!r.visit_date || !r.visit_time) {
      // 시간 없는 건 skip (방어 — 직접방문은 visit_date/time 필수지만 데이터 정합성 가드)
      return;
    }

    // 설정 로드 (consultation_settings + system_settings — 컨설팅과 공유)
    const settings = await loadFormatSettings();

    const repairForEvent: RepairForCalendar = {
      id: r.id,
      as_id: r.as_id,
      name: r.name,
      phone: r.phone,
      visit_date: r.visit_date,
      visit_time: r.visit_time,
      visit_duration_min: r.visit_duration_min,
      status: r.status,
      qty_mamoru: r.qty_mamoru,
      qty_other: r.qty_other,
      memo: r.memo,
      service_cost: r.service_cost,
      total_amount: r.total_amount,
      created_at: r.created_at,
    };

    const eventPayload = formatRepairToEvent(repairForEvent, settings, BASE_URL);

    if (eventId) {
      // UPDATE 시도
      const res = await updateCalendarEvent({ eventId, event: eventPayload });
      if (res.ok) {
        await dbAny
          .from('repairs')
          .update({ google_event_updated_at: new Date().toISOString() })
          .eq('id', repairId);
      } else if (res.error && /\b(404|410|notFound|deleted)\b/i.test(res.error)) {
        // 이벤트가 사라진 경우 → 재생성
        const createRes = await createCalendarEvent({ event: eventPayload });
        if (createRes.ok && createRes.eventId) {
          await dbAny
            .from('repairs')
            .update({
              google_event_id: createRes.eventId,
              google_event_updated_at: new Date().toISOString(),
            })
            .eq('id', repairId);
        }
      }
      // not_connected / 기타 오류 = silent skip (운영 차단 X)
    } else {
      // CREATE
      const res = await createCalendarEvent({ event: eventPayload });
      if (res.ok && res.eventId) {
        await dbAny
          .from('repairs')
          .update({
            google_event_id: res.eventId,
            google_event_updated_at: new Date().toISOString(),
          })
          .eq('id', repairId);
      }
    }
  } catch (err) {
    console.error('[repair-calendar-sync] 예상치 못한 오류', {
      repairId,
      error: err instanceof Error ? err.message : String(err),
    });
  }
}

// (제거) fireAndForgetRepairSync — void 반환이라 after()/서버리스에서 요청이 잘려 미기록되던 footgun.
//   호출부는 `after(() => syncRepairToCalendar(id))` 로 Promise 를 await 하게 통일(2026-08-04).

async function loadFormatSettings(): Promise<EventFormatSettings> {
  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const dbAny = db as any;

  const [csRes, ssRes] = await Promise.all([
    dbAny.from('consultation_settings').select('*').eq('id', 'default').single(),
    dbAny
      .from('system_settings')
      .select('key, value')
      .in('key', ['business.store_address', 'business.store_name']),
  ]);

  const cs = csRes.data || {};
  const ssMap: Record<string, string> = {};
  (ssRes.data || []).forEach((row: { key: string; value: string | null }) => {
    if (row.value) ssMap[row.key] = String(row.value).replace(/^"|"$/g, '');
  });

  return {
    store_address: ssMap['business.store_address'],
    store_name: ssMap['business.store_name'],
    duration_min: cs.duration_min ?? 60,
    field_buffer_before: cs.field_buffer_before ?? 60,
    field_buffer_after: cs.field_buffer_after ?? 60,
  };
}
