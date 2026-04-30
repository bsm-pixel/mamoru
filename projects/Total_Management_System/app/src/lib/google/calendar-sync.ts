/**
 * 상담 상태 변경 → Google Calendar 동기화 orchestrator
 *
 * 호출 규칙: fire-and-forget — 상담 로직을 절대 블록하면 안 됨
 * 모든 오류는 내부에서 캐치하고 로그만 남김
 */

import { createServiceClient } from '@/lib/supabase/server';
import { createCalendarEvent, updateCalendarEvent, deleteCalendarEvent } from './calendar-client';
import { formatConsultationToEvent, type ConsultationForCalendar, type EventFormatSettings } from './event-formatter';

const BASE_URL = process.env.NEXT_PUBLIC_APP_URL || 'https://app-eta-sandy-75.vercel.app';

/** 캘린더 미표시 상태 (이미 이벤트 있으면 삭제) */
const HIDDEN_STATES = new Set(['suggested', 'assigned', 'pending_admin']);

/** 이벤트 삭제 대상 */
const DELETE_STATES = new Set(['cancelled', 'on_hold']);

/**
 * 상담 건을 Google Calendar와 동기화
 *
 * @param consultationId - consultations.id (UUID)
 */
export async function syncConsultationToCalendar(consultationId: string): Promise<void> {
  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // 현재 상담 조회
    const { data: c, error } = await dbAny
      .from('consultations')
      .select('*')
      .eq('id', consultationId)
      .single();

    if (error || !c) {
      console.warn('[calendar-sync] 상담 조회 실패', consultationId, error?.message);
      return;
    }

    // 톡상담은 캘린더 대상 아님 (일정 없음)
    if (c.consultation_type === 'talk_consult') return;

    const currentStatus = c.status as string;
    const eventId = c.google_event_id as string | null;

    // 1) 숨김 상태 — 이미 이벤트 있으면 삭제
    if (HIDDEN_STATES.has(currentStatus)) {
      if (eventId) {
        await deleteCalendarEvent({ eventId });
        await dbAny
          .from('consultations')
          .update({ google_event_id: null, google_event_updated_at: new Date().toISOString() })
          .eq('id', consultationId);
      }
      return;
    }

    // 2) 삭제 대상 상태 (cancelled / on_hold)
    if (DELETE_STATES.has(currentStatus)) {
      if (eventId) {
        await deleteCalendarEvent({ eventId });
        await dbAny
          .from('consultations')
          .update({ google_event_id: null, google_event_updated_at: new Date().toISOString() })
          .eq('id', consultationId);
      }
      return;
    }

    // 3) 유지/생성/업데이트 대상 (confirmed / reschedule_requested / change_requested / completed)
    if (!c.visit_date || !c.visit_time) {
      // 시간 없는 건 — 스킵
      return;
    }

    // 설정 로드 (consultation_settings + system_settings)
    const settings = await loadFormatSettings();

    const consultationForEvent: ConsultationForCalendar = {
      id: c.id,
      name: c.name,
      phone: c.phone,
      consultation_type: c.consultation_type,
      visit_date: c.visit_date,
      visit_time: c.visit_time,
      status: c.status,
      address_road: c.address_road,
      address_detail: c.address_detail,
      address_sigungu: c.address_sigungu,
      memo: c.memo,
      unique_id: c.unique_id,
      created_at: c.created_at,
      // consultations 테이블에 completed_at 컬럼이 없으므로 completed 상태일 때 updated_at으로 대체
      completed_at: c.status === 'completed' ? c.updated_at : null,
      gas_raw: c.gas_raw,
    };

    const eventPayload = formatConsultationToEvent(consultationForEvent, settings, BASE_URL);

    if (eventId) {
      // UPDATE 시도
      const res = await updateCalendarEvent({ eventId, event: eventPayload });
      if (res.ok) {
        await dbAny
          .from('consultations')
          .update({ google_event_updated_at: new Date().toISOString() })
          .eq('id', consultationId);
      } else if (res.error && /\b(404|410|notFound|deleted)\b/i.test(res.error)) {
        // 이벤트가 사라진 경우 → 재생성
        const createRes = await createCalendarEvent({ event: eventPayload });
        if (createRes.ok && createRes.eventId) {
          await dbAny
            .from('consultations')
            .update({
              google_event_id: createRes.eventId,
              google_event_updated_at: new Date().toISOString(),
            })
            .eq('id', consultationId);
        }
      }
      // not_connected나 기타 오류는 조용히 종료 (상담 로직 블록 안 함)
    } else {
      // CREATE
      const res = await createCalendarEvent({ event: eventPayload });
      if (res.ok && res.eventId) {
        await dbAny
          .from('consultations')
          .update({
            google_event_id: res.eventId,
            google_event_updated_at: new Date().toISOString(),
          })
          .eq('id', consultationId);
      }
    }
  } catch (err) {
    console.error('[calendar-sync] 예상치 못한 오류', {
      consultationId,
      error: err instanceof Error ? err.message : String(err),
    });
  }
}

/** 비동기 호출용 헬퍼 — Promise 무시, 예외 모두 삼킴 */
export function fireAndForgetSync(consultationId: string): void {
  syncConsultationToCalendar(consultationId).catch(() => {
    /* 이미 내부에서 캐치함 */
  });
}

async function loadFormatSettings(): Promise<EventFormatSettings> {
  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const dbAny = db as any;

  const [csRes, ssRes] = await Promise.all([
    dbAny.from('consultation_settings').select('*').eq('id', 'default').single(),
    dbAny
      .from('system_settings')
      .select('key, value')
      .in('key', [
        'business.store_address',
        'business.store_name',
        'consultation.duration_by_type',
      ]),
  ]);

  const cs = csRes.data || {};
  const ssMap: Record<string, string> = {};
  (ssRes.data || []).forEach((r: { key: string; value: string | null }) => {
    if (r.value) ssMap[r.key] = String(r.value).replace(/^"|"$/g, '');
  });

  // duration_by_type은 JSON으로 저장돼 있을 수 있음
  let durationByType: { store_visit?: number; field_request?: number } = {};
  try {
    const raw = ssMap['consultation.duration_by_type'];
    if (raw) durationByType = JSON.parse(raw);
  } catch {
    /* ignore */
  }

  return {
    store_address: ssMap['business.store_address'],
    store_name: ssMap['business.store_name'],
    duration_min: cs.duration_min ?? 60,
    field_buffer_before: cs.field_buffer_before ?? 90,
    field_buffer_after: cs.field_buffer_after ?? 90,
    duration_store_visit: durationByType.store_visit,
    duration_field_request: durationByType.field_request,
  };
}
