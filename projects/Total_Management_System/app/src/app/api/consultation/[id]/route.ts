import { NextRequest, NextResponse, after } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { isValidTransition } from '@/lib/consultation/transitions';
import { sendNotification, type NotifyTemplate } from '@/lib/notification/make-webhook';
import { syncConsultationToCalendar } from '@/lib/google/calendar-sync';
import { deleteCalendarEvent } from '@/lib/google/calendar-client';
import { getServerSetting } from '@/hooks/use-settings';
import type { ConsultationStatus, ConsultationType } from '@/lib/supabase/types';

/** 상태→알림 템플릿 매핑 (consultation_type 기반 분기) */
function getAutoNotifyTemplate(
  newStatus: string,
  consultationType: string
): NotifyTemplate | null {
  if (newStatus === 'confirmed') {
    return consultationType === 'field_request' ? 'field_confirmed' : 'confirmed';
  }
  if (newStatus === 'cancelled') {
    if (consultationType === 'field_request') return 'field_cancelled';
    if (consultationType === 'talk_consult') return null; // 톡상담은 카톡으로 이미 커뮤니케이션 중 → 알림톡 불필요
    return 'cancelled';
  }
  if (newStatus === 'completed') {
    return 'review_request'; // 상담완료 → 리뷰 요청 알림톡 (MAKE_WEBHOOK_URL)
  }
  if (newStatus === 'in_progress' && consultationType === 'talk_consult') {
    return 'talk_ready'; // 톡상담 시작 → 안내 알림톡
  }
  return null;
}

/* GAS cancelViaGAS 제거 완료 — Calendar 삭제 불필요, 슬롯 차단은 Supabase 쿼리 */

/** GET /api/consultation/[id] — 상담 단건 + 이력 */
export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const [consultRes, historyRes] = await Promise.all([
      db.from('consultations').select('*').eq('id', id).single(),
      db
        .from('consultation_history')
        .select('*')
        .eq('consultation_id', id)
        .order('created_at', { ascending: false }),
    ]);

    if (consultRes.error) throw consultRes.error;

    return NextResponse.json({
      consultation: consultRes.data,
      history: historyRes.data || [],
    });
  } catch (err) {
    console.error('[consultation] 상세 조회 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/consultation/[id] — 상담 상태/정보 변경 (Phase 2-2: 전이 검증 + hold_reason) */
export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();
    const { status: newStatus, note, ...rest } = body;

    // 현재 상담 조회 (캘린더 동기화 판단용 visit_date/visit_time 포함)
    const { data: current, error: fetchErr } = await db
      .from('consultations')
      .select('status, consultation_type, visit_date, visit_time')
      .eq('id', id)
      .single();

    if (fetchErr || !current) {
      return NextResponse.json({ error: '상담을 찾을 수 없습니다' }, { status: 404 });
    }

    // 상태 전이 유효성 검증
    if (newStatus && newStatus !== current.status) {
      const valid = isValidTransition(
        current.consultation_type as ConsultationType,
        current.status as ConsultationStatus,
        newStatus as ConsultationStatus
      );
      if (!valid) {
        return NextResponse.json(
          { error: `상태 전이 불가: ${current.status} → ${newStatus}` },
          { status: 400 }
        );
      }
    }

    // 업데이트 데이터 구성
    const updateData = { ...rest };
    if (newStatus) updateData.status = newStatus;

    // on_hold가 아닌 상태로 전이 시 hold_reason 초기화
    if (newStatus && newStatus !== 'on_hold') {
      updateData.hold_reason = null;
    }

    const { data, error } = await db
      .from('consultations')
      .update(updateData)
      .eq('id', id)
      .select()
      .single();

    if (error) throw error;

    // 상태 변경 또는 일정 변경 감지 (캘린더 동기화 트리거 조건)
    const statusChanged = newStatus && newStatus !== current.status;
    const dateChanged = 'visit_date' in updateData && updateData.visit_date !== current.visit_date;
    const timeChanged = 'visit_time' in updateData && updateData.visit_time !== current.visit_time;
    const scheduleChanged = dateChanged || timeChanged;

    // 상태 변경 시 이력 기록
    if (statusChanged) {
      await db.from('consultation_history').insert({
        consultation_id: id,
        from_status: current.status,
        to_status: newStatus,
        changed_by: user.id,
        note: note || null,
      });
    }

    // 후속 작업: after()로 응답 반환 후 백그라운드 실행 — UI 빠른 응답
    if (statusChanged || scheduleChanged) {
      after(async () => {
        const sideEffects: Promise<unknown>[] = [];

        // Google Calendar 동기화 (await로 완료 보장, 내부에서 예외 모두 캐치)
        sideEffects.push(syncConsultationToCalendar(id));

        // 자동 알림톡 발송 — 상태 변경일 때만
        const template = statusChanged ? getAutoNotifyTemplate(newStatus, data.consultation_type) : null;
        if (template && data.phone) {
          // 067: 후기 요청 자동 발송 가드 — 정책 토글 OFF / 약속 ✓ / 이미 발송 시 skip
          let allowSend = true;
          if (template === 'review_request') {
            const autoEnabled = await getServerSetting<boolean>(db, 'review.auto_request_on_completion', false);
            if (!autoEnabled) allowSend = false;
            else if (data.review_promised_at) allowSend = false;     // 약속 고객은 사장님 수동만
            else if (data.review_request_sent_at) allowSend = false; // 중복 방지
          }

          if (allowSend) {
            const address = [data.address_road, data.address_detail].filter(Boolean).join(' ');
            const typeLabel = data.consultation_type === 'store_visit' ? '매장 방문'
              : data.consultation_type === 'field_request' ? '출장 요청' : '온라인상담';
            sideEffects.push(
              sendNotification({
                template,
                phone: data.phone,
                name: data.name,
                data: {
                  id: data.unique_id,
                  type: typeLabel,
                  date: data.visit_date || '',
                  time: data.visit_time || '',
                  address,
                  uid: data.unique_id,                 // 리뷰 폼 uid 파라미터
                  review_type: 'consult',              // 리뷰 폼 type 파라미터
                  type_label: typeLabel,               // 알림톡 치환 변수
                  subtype: data.consultation_type || '', // store_visit, field_request, talk_consult
                  consult_uid: data.unique_id,         // 솔라피 #{consult_uid} 치환용
                },
              }).then(async (r) => {
                if (!r.success) {
                  console.error('[auto-notify] 발송 실패:', r.error);
                } else {
                  console.log('[auto-notify] 발송 성공:', template);
                  // 067: 후기 요청 자동 발송 성공 시 시각 기록 (중복 방지)
                  if (template === 'review_request') {
                    try {
                      await db.from('consultations')
                        .update({ review_request_sent_at: new Date().toISOString() })
                        .eq('id', id);
                    } catch (e) { console.error('[auto-notify] review_request_sent_at 기록 실패:', e); }
                  }
                }
              })
            );
          } else {
            console.log('[auto-notify] review_request skip (policy/promised/duplicate)');
          }
        }

        if (sideEffects.length > 0) {
          const results = await Promise.allSettled(sideEffects);
          results.forEach((r, i) => {
            if (r.status === 'rejected') {
              console.error(`[side-effect ${i}] 실패:`, r.reason);
            }
          });
        }
      });
    }

    return NextResponse.json(data);
  } catch (err) {
    console.error('[consultation] 업데이트 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/consultation/[id] — 상담 건 완전 삭제 (알림톡 없음) */
export async function DELETE(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: consultation, error: fetchErr } = await db
      .from('consultations')
      .select('id, unique_id, name, google_event_id')
      .eq('id', id)
      .single();

    if (fetchErr || !consultation) {
      return NextResponse.json({ error: '상담 건을 찾을 수 없습니다' }, { status: 404 });
    }

    // Google Calendar 이벤트 정리 (있으면 삭제, 실패해도 상담 삭제는 진행)
    if (consultation.google_event_id) {
      try {
        await deleteCalendarEvent({ eventId: consultation.google_event_id });
      } catch (e) {
        console.warn('[consultation delete] 캘린더 이벤트 삭제 실패:', e);
      }
    }

    // 이력 삭제 → 본건 삭제
    await db.from('consultation_history').delete().eq('consultation_id', id);
    const { error: delErr } = await db.from('consultations').delete().eq('id', id);

    if (delErr) throw delErr;

    return NextResponse.json({ ok: true, deleted: consultation.unique_id });
  } catch (err) {
    console.error('[consultation] 삭제 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
