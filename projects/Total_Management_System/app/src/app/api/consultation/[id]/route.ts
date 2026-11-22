import { NextRequest, NextResponse, after } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { isValidTransition } from '@/lib/consultation/transitions';
import { sendNotification, type NotifyTemplate } from '@/lib/notification/make-webhook';
import { syncConsultationToCalendar } from '@/lib/google/calendar-sync';
import { deleteCalendarEvent } from '@/lib/google/calendar-client';
import { geocodeForConsultation } from '@/lib/kakao/geocode';
import type { ConsultationStatus, ConsultationType } from '@/lib/supabase/types';

const GITHUB_PAGES = 'page.mamoru.kr/projects/consulting';

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
  // 2026-04-30: 상담완료 → 리뷰 요청 자동 발송 제거. 후기는 sale source 단일 진입점.
  if (newStatus === 'in_progress' && consultationType === 'talk_consult') {
    return 'talk_ready'; // 톡상담 시작 → 안내 알림톡
  }
  return null;
}

/* GAS cancelViaGAS 제거 완료 — Calendar 삭제 불필요, 슬롯 차단은 Supabase 쿼리 */

import { errMsg } from '@/lib/utils/err';

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
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
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
    // 항상 발송 원칙(2026-09-14): 알림 끄기 옵션 없음. 캐시된 구버전 화면이 보내는 skip_notify 는 DB 컬럼이 아니므로 버림
    delete body.skip_notify;
    // also_customer: 주소를 고칠 때 고객 정보 주소도 같이 바꿀지(DB 컬럼 아님, 2026-11-22)
    // address_to_customer: 'apply' | 'keep' — 접수 주소가 고객 정보 주소와 다를 때 사장님이 확인 창에서 고른 것(2026-11-22)
    const { status: newStatus, note, also_customer: alsoCustomer, address_to_customer: addressToCustomer, ...rest } = body;

    // 현재 상담 조회 (캘린더 동기화 판단용 visit_date/visit_time · 주소 고치기 판단용 주소 포함)
    const { data: current, error: fetchErr } = await db
      .from('consultations')
      .select('status, consultation_type, visit_date, visit_time, admin_note, address_road, address_detail, postcode, customer_id, gas_raw')
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

    // 2026-11-22 주소 고치기 — 고객이 출장 요청 주소를 잘못 넣은 경우(작업실 t-261121-field-address)
    //   좌표를 새 주소로 다시 계산한다(지도 핀 · 휴대폰 길찾기). 못 찾으면 비워 둔다 — 옛 주소 핀이 남는 것보다 낫고,
    //   backfill-coords 가 나중에 다시 시도한다. 알림톡은 보내지 않는다(아래 template 은 상태 · 일정 변경에만 정해짐).
    const addressChanged = 'address_road' in updateData && (
      (updateData.address_road || '') !== (current.address_road || '') ||
      (updateData.address_detail || '') !== (current.address_detail || '')
    );
    if ('address_road' in updateData) {
      const road = String(updateData.address_road || '').trim();
      if (!road) return NextResponse.json({ error: '도로명 주소를 넣어 주세요' }, { status: 400 });
      updateData.address_road = road;
      updateData.address_detail = String(updateData.address_detail || '').trim() || null;
      if ('postcode' in updateData) updateData.postcode = String(updateData.postcode || '').trim() || null;
    }
    // 고객 주소에 대한 결정을 이 건에 남긴다(gas_raw.address_decision) — 상세를 다시 열어도 같은 확인 창이 뜨지 않게
    //   주소를 고쳤으면 그때 고른 것(고객 주소도 바꿈 = applied / 이 건만 = kept)으로 기록한다
    const decision: 'applied' | 'kept' | null =
      addressToCustomer === 'apply' ? 'applied'
      : addressToCustomer === 'keep' ? 'kept'
      : addressChanged ? (alsoCustomer ? 'applied' : 'kept')
      : null;
    if (decision) {
      const raw = (current.gas_raw && typeof current.gas_raw === 'object') ? current.gas_raw : {};
      updateData.gas_raw = { ...raw, address_decision: decision, address_decided_at: new Date().toISOString() };
    }
    if (addressChanged) {
      const geo = await geocodeForConsultation(updateData.address_road as string, (updateData.address_detail as string | null) || undefined);
      updateData.latitude = geo?.lat ?? null;
      updateData.longitude = geo?.lng ?? null;
      // 옛 접수(GAS)의 시도 · 시군구 · 지역 칸은 옛 주소 기준이라 비운다(상담 전체 화면에 그대로 보이던 칸)
      updateData.address_sido = null;
      updateData.address_sigungu = null;
      updateData.address_region = null;
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
    // 상담자 메모 변경 시에도 캘린더 설명 갱신 (108)
    const adminNoteChanged = 'admin_note' in updateData && updateData.admin_note !== current.admin_note;

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

    // 주소를 고쳤으면 이력(상태 그대로 = 화면에 「내용 수정」) + 고객 정보 주소(고른 경우만)
    if (addressChanged) {
      const before = [current.address_road, current.address_detail].filter(Boolean).join(' ') || '(없음)';
      await db.from('consultation_history').insert({
        consultation_id: id,
        from_status: current.status,
        to_status: current.status,
        changed_by: user.id,
        note: `주소 변경 — 이전: ${before}${alsoCustomer && current.customer_id ? ' (고객 정보 주소도 바꿈)' : ''}`,
      });
      if (alsoCustomer && current.customer_id) {
        const patch: Record<string, unknown> = { address_road: updateData.address_road, address_detail: updateData.address_detail };
        if ('postcode' in updateData) patch.postcode = updateData.postcode;
        await db.from('customers').update(patch).eq('id', current.customer_id);
      }
    }

    // 접수 주소를 고객 주소로 (주소는 그대로, 확인 창에서 [고객 주소로 바꾸기]) — 2026-11-22
    if (addressToCustomer === 'apply' && !addressChanged && current.customer_id && current.address_road) {
      await db.from('customers').update({
        address_road: current.address_road,
        address_detail: current.address_detail || null,
        postcode: current.postcode || null,
      }).eq('id', current.customer_id);
      await db.from('consultation_history').insert({
        consultation_id: id,
        from_status: current.status,
        to_status: current.status,
        changed_by: user.id,
        note: '고객 정보 주소를 이 접수 주소로 바꿈',
      });
    }

    // 후속 작업: after()로 응답 반환 후 백그라운드 실행 — UI 빠른 응답
    //   adminNoteChanged 만인 경우: 캘린더 설명만 갱신되고 알림톡 template은 null 유지(고객 발송 없음)
    //   addressChanged 만인 경우(2026-11-22): 구글 달력 장소만 새 주소로 — 알림톡 없음
    if (statusChanged || scheduleChanged || adminNoteChanged || addressChanged) {
      after(async () => {
        const sideEffects: Promise<unknown>[] = [];

        // Google Calendar 동기화 (await로 완료 보장, 내부에서 예외 모두 캐치)
        sideEffects.push(syncConsultationToCalendar(id));

        // 자동 알림톡 발송 — 상태 변경일 때 또는 (확정 상태에서 일정만 변경됐을 때) 발송
        //   확정 상태에서 일정만 바뀐 경우(수동 일정 변경) → field_rescheduled/rescheduled 로 분기 — Make `출장_일정변경(관리자변경)` 모듈로 라우팅됨
        let template: NotifyTemplate | null = null;
        if (statusChanged) {
          template = getAutoNotifyTemplate(newStatus, data.consultation_type);
        } else if (scheduleChanged && current.status === 'confirmed') {
          template = data.consultation_type === 'field_request' ? 'field_rescheduled' : 'rescheduled';
        }
        // 일정변경 알림은 여기 한 곳에서만 발송 (모달이 /notify 로 한 번 더 쏘던 2통 발송 제거 — 2026-09-13)
        if (template && data.phone) {
          const address = [data.address_road, data.address_detail].filter(Boolean).join(' ');
          const typeLabel = data.consultation_type === 'store_visit' ? '매장 방문'
            : data.consultation_type === 'field_request' ? '출장 요청' : '온라인상담';
          const effectiveStatus = (newStatus || current.status || '').toString().toUpperCase();
          sideEffects.push(
            sendNotification({
              template,
              phone: data.phone,
              name: data.name,
              data: {
                // admin-create와 동일 필드 (카카오 알림톡 템플릿 변수 매칭 보장 — 누락 시 3109 SMS 대체 발송 사고)
                id: data.unique_id,
                status: effectiveStatus,
                name: data.name,
                phone: data.phone,
                type: typeLabel,
                date: data.visit_date || '',
                time: data.visit_time || '',
                // field_rescheduled 등 #{visit_date}/#{visit_time} 변수를 쓰는 템플릿 호환 (date/time 도 함께 보냄)
                visit_date: data.visit_date || '',
                visit_time: data.visit_time || '',
                // 2026-11-22 Make 「매장방문_일정변경」은 #{new_date}·#{new_time} 을 new_date·new_time 에서 읽는다 — 보내지 않아 바뀐 날짜·시간이 빈칸이었다
                new_date: data.visit_date || '',
                new_time: data.visit_time || '',
                address,
                days: '',
                memo: data.memo || '',
                change_request_link: `${GITHUB_PAGES}/page_change_request.html?uid=${data.unique_id}`,
                // 추가 메타 (다른 곳에서 사용 가능)
                uid: data.unique_id,
                type_label: typeLabel,
                subtype: data.consultation_type || '',
                consult_uid: data.unique_id,
              },
            }).then((r) => {
              if (!r.success) console.error('[auto-notify] 발송 실패:', r.error);
              else console.log('[auto-notify] 발송 성공:', template);
            })
          );
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
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
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
    const calendarResult: { attempted: boolean; ok: boolean; error?: string } = {
      attempted: false,
      ok: true,
    };
    if (consultation.google_event_id) {
      calendarResult.attempted = true;
      try {
        await deleteCalendarEvent({ eventId: consultation.google_event_id });
      } catch (e) {
        calendarResult.ok = false;
        calendarResult.error = errMsg(e);
        console.warn('[consultation delete] 캘린더 이벤트 삭제 실패:', e);
      }
    }

    // 이력 삭제 (실패 시 명확히 throw — 침묵 금지)
    const { error: histErr } = await db
      .from('consultation_history')
      .delete()
      .eq('consultation_id', id);
    if (histErr) {
      console.error('[consultation delete] 이력 삭제 실패:', histErr);
      throw histErr;
    }

    // 본건 삭제
    const { error: delErr } = await db.from('consultations').delete().eq('id', id);
    if (delErr) throw delErr;

    return NextResponse.json({
      ok: true,
      deleted: consultation.unique_id,
      calendar: calendarResult,
    });
  } catch (err) {
    console.error('[consultation] 삭제 실패:', err);
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}
