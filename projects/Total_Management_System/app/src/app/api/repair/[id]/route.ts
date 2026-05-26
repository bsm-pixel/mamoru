import { NextRequest, NextResponse, after } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { isValidRepairTransition } from '@/lib/repair/transitions';
import { sendNotification, type NotifyTemplate } from '@/lib/notification/make-webhook';
import { getServerSetting } from '@/hooks/use-settings';
import type { RepairStatus } from '@/lib/supabase/types';
import { fireAndForgetRepairSync } from '@/lib/google/repair-calendar-sync';

/** 상태→자동 알림톡 매핑 (payment_confirmed는 paid_at 플래그로 분리됨) */
function getAutoNotifyTemplate(newStatus: string): NotifyTemplate | null {
  const map: Record<string, NotifyTemplate> = {
    shipped: 'as_shipped',
    cancelled: 'as_cancelled',
    delivered: 'as_review_request', // 배송완료 → 리뷰 요청 (MAKE_REPAIR_WEBHOOK_URL)
  };
  return map[newStatus] || null;
}

/** GET /api/repair/[id] — 단건 + 검수 + 이력 */
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
    const [repairRes, inspectionsRes, historyRes] = await Promise.all([
      db.from('repairs').select('*').eq('id', id).single(),
      db.from('repair_inspections').select('*').eq('repair_id', id).order('scissor_number', { ascending: true }),
      db.from('repair_history').select('*').eq('repair_id', id).order('created_at', { ascending: false }),
    ]);

    if (repairRes.error) throw repairRes.error;

    return NextResponse.json({
      repair: repairRes.data,
      inspections: inspectionsRes.data || [],
      history: historyRes.data || [],
    });
  } catch (err) {
    console.error('[repair] 상세 조회 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/repair/[id] — 상태/정보 변경 */
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
    const { status: newStatus, note, skip_notify, ...rest } = body;

    // 현재 조회
    const { data: current, error: fetchErr } = await db
      .from('repairs')
      .select('*')
      .eq('id', id)
      .single();

    if (fetchErr || !current) {
      return NextResponse.json({ error: '복원수리 건을 찾을 수 없습니다' }, { status: 404 });
    }

    // 상태 전이 검증
    if (newStatus && newStatus !== current.status) {
      const valid = isValidRepairTransition(
        current.status as RepairStatus,
        newStatus as RepairStatus
      );
      if (!valid) {
        return NextResponse.json(
          { error: `상태 전이 불가: ${current.status} → ${newStatus}` },
          { status: 400 }
        );
      }
    }

    // 업데이트 데이터
    const updateData = { ...rest };
    if (newStatus) updateData.status = newStatus;

    const { data, error } = await db
      .from('repairs')
      .update(updateData)
      .eq('id', id)
      .select()
      .single();

    if (error) throw error;

    // paid_at 설정 시 → 입금확인 알림톡 (상태 변경과 독립)
    const justPaid = rest.paid_at && !current.paid_at;

    // 상태 변경 시 이력 기록 + 백그라운드 알림톡
    if (newStatus && newStatus !== current.status) {
      await db.from('repair_history').insert({
        repair_id: id,
        from_status: current.status,
        to_status: newStatus,
        changed_by: user.id,
        note: note || null,
      });

      after(async () => {
        if (skip_notify) return;  // 합포장 출고 등 알림톡 우회 케이스
        const template = getAutoNotifyTemplate(newStatus);
        if (template && data.phone) {
          // 067: 후기 요청 자동 발송 가드
          // 2026-05-26 정책 정정 (사장님 의도): 약속 ✓ 고객만 자동 발송 — 보수적
          //   기존 정책 (약속 X 고객만 자동) 폐기 → 약속 받은 고객만 자동, 약속 X 고객은 사장님 수동만
          let allowSend = true;
          if (template === 'as_review_request') {
            const autoEnabled = await getServerSetting<boolean>(db, 'review.auto_request_on_completion', false);
            if (!autoEnabled) allowSend = false;
            else if (!data.review_promised_at) allowSend = false;   // 약속 X → 자동 발송 안 함
            else if (data.review_request_sent_at) allowSend = false; // 중복 방지
          }

          if (!allowSend) {
            console.log('[repair auto-notify] as_review_request skip (policy/promised/duplicate)');
            return;
          }

          const result = await sendNotification({
            template,
            phone: data.phone,
            name: data.name,
            data: {
              id: data.as_id,
              as_amount: String(data.service_cost || 0),
              shipping_amount: String(data.shipping_fee || 0),
              total_amount: String(data.total_amount || 0),
              tracking: data.invoice_number || '',
              courier: '롯데택배',
              uid: data.as_id,                   // 리뷰 폼 uid 파라미터
              as_uid: data.as_id,                // 솔라피 #{as_uid} 치환용
              review_type: 'repair',             // 리뷰 폼 type 파라미터
              type_label: '복원수리',             // 알림톡 치환 변수
            },
          });
          if (!result.success) {
            console.error('[repair auto-notify] 실패:', result.error);
          } else {
            console.log('[repair auto-notify] 성공:', template);
            // 067: 후기 요청 자동 발송 성공 시 시각 기록 (중복 방지)
            if (template === 'as_review_request') {
              try {
                await db.from('repairs')
                  .update({ review_request_sent_at: new Date().toISOString() })
                  .eq('id', id);
              } catch (e) { console.error('[repair auto-notify] review_request_sent_at 기록 실패:', e); }
            }
          }
        }
      });
    }

    // 입금확인 알림톡 (paid_at 플래그 설정 시, skip_notify가 아닐 때)
    if (justPaid && !skip_notify) {
      after(async () => {
        if (data.phone) {
          const result = await sendNotification({
            template: 'as_payment_confirmed',
            phone: data.phone,
            name: data.name,
            data: {
              id: data.as_id,
              as_amount: String(data.service_cost || 0),
              shipping_amount: String(data.shipping_fee || 0),
              total_amount: String(data.total_amount || 0),
              tracking: data.invoice_number || '',
              courier: '롯데택배',
            },
          });
          if (!result.success) console.error('[repair paid_at notify] 실패:', result.error);
          else console.log('[repair paid_at notify] 입금확인 알림톡 발송');
        }
      });
    }

    // 2026-05-25 Phase 3-B: 직접방문 건 → Google Calendar 자동 동기화 (fire-and-forget)
    //   - status / visit_date / visit_time 등 변경 시 자동 반영
    //   - cancelled 시 이벤트 자동 삭제 (sync 내부 로직)
    //   - 비직접방문 건은 sync 내부에서 자연 skip (proceed_type 가드)
    if (data.proceed_type === '직접방문') {
      after(() => fireAndForgetRepairSync(data.id));
    }

    return NextResponse.json(data);
  } catch (err) {
    console.error('[repair] 업데이트 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/repair/[id] — 복원수리 건 완전 삭제 (알림톡 없음) */
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

    // 현재 건 조회 (송장 + Google 이벤트 확인용)
    const { data: repair, error: fetchErr } = await db
      .from('repairs')
      .select('id, as_id, invoice_number, google_event_id, proceed_type')
      .eq('id', id)
      .single();

    if (fetchErr || !repair) {
      return NextResponse.json({ error: '복원수리 건을 찾을 수 없습니다' }, { status: 404 });
    }

    // 2026-05-25 Phase 3-B: 직접방문 + Google 이벤트 있으면 이벤트 먼저 삭제
    //   삭제는 best-effort — 실패해도 DB 삭제 진행
    if (repair.proceed_type === '직접방문' && repair.google_event_id) {
      try {
        const { deleteCalendarEvent } = await import('@/lib/google/calendar-client');
        await deleteCalendarEvent({ eventId: repair.google_event_id });
      } catch (e) {
        console.warn(`[repair DELETE] Google Calendar 이벤트 삭제 실패 (DB 삭제 진행):`, e);
      }
    }

    // 송장이 있으면 ALPS 취소 시도 (실패해도 삭제 진행)
    let alpsWarning: string | undefined;
    if (repair.invoice_number) {
      try {
        const { cancelShipment } = await import('@/lib/lotte/alps-client');
        const cancelResult = await cancelShipment(repair.invoice_number);
        if (!cancelResult.success) {
          alpsWarning = `ALPS 취소 실패: ${cancelResult.error}`;
          console.warn(`[repair delete] ${alpsWarning}`);
        }
      } catch (e) {
        alpsWarning = `ALPS 취소 오류: ${String(e)}`;
      }
    }

    // 사진 파일 + 메타데이터 삭제
    try {
      const { data: photos } = await db.from('repair_photos').select('file_path').eq('repair_id', id);
      if (photos?.length > 0) {
        const paths = photos.map((p: { file_path: string }) => p.file_path).filter(Boolean);
        if (paths.length > 0) await supabase.storage.from('repair-photos').remove(paths);
        await db.from('repair_photos').delete().eq('repair_id', id);
      }
    } catch { /* 사진 삭제 실패해도 본건 삭제 계속 */ }

    // 이력 삭제 → 검수 삭제 → 본건 삭제
    await db.from('repair_history').delete().eq('repair_id', id);
    await db.from('repair_inspections').delete().eq('repair_id', id);
    const { error: delErr } = await db.from('repairs').delete().eq('id', id);

    if (delErr) throw delErr;

    // Realtime fallback용 push_notifications 행도 함께 정리 (tag 기반)
    try {
      await db.from('push_notifications').delete().eq('tag', `mamoru-as_received-${repair.as_id}`);
    } catch { /* 푸시 알림 정리 실패해도 삭제는 성공 */ }

    return NextResponse.json({ ok: true, deleted: repair.as_id, warning: alpsWarning });
  } catch (err) {
    console.error('[repair] 삭제 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
