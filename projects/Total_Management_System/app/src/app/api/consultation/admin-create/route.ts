/**
 * POST /api/consultation/admin-create
 * 관리자 직접 상담 등록 (인스타 DM / 유선 등 외부 채널로 들어온 건을 TMS에 수기 입력)
 *
 * 정책:
 *   - 매장방문 / 출장요청 2가지만 지원 (톡상담 제외)
 *   - status: 'confirmed' 로 즉시 확정 (날짜·시간 필수)
 *   - 중복 체크: phone_normalized 기준, 오늘 이후 같은 visit_date+visit_time → 409 반환
 *   - 알림톡 발송 (옵션, 기본 On): confirmed / field_confirmed 템플릿
 *   - Google Calendar 자동 동기화
 *   - 리마인더 cron 자동 포함 (status=confirmed + visit_date/time)
 *   - gas_raw.source = 'admin_manual' 로 수기 접수 마킹
 */

import { NextRequest, NextResponse, after } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';
import { syncConsultationToCalendar } from '@/lib/google/calendar-sync';
import { geocodeForConsultation } from '@/lib/kakao/geocode';
import crypto from 'crypto';

const GITHUB_PAGES = 'bsm-pixel.github.io/mamoru/projects/consulting';

export async function POST(req: NextRequest) {
  // 관리자 인증 필수
  const supabase = await createServerSupabaseClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  try {
    const body = await req.json();
    const {
      type,          // 'store_visit' | 'field_request'
      name,
      phone,
      visitDate,
      visitTime,
      addressRoad,
      addressDetail,
      postcode,
      memo,
      notify = true, // 기본 On
    } = body;

    // ── 입력 검증 ──────────────────────────────────────────────
    if (!['store_visit', 'field_request'].includes(type)) {
      return NextResponse.json({ ok: false, error: '타입은 매장방문 또는 출장요청만 가능합니다' }, { status: 400 });
    }
    if (!name?.trim() || !phone?.trim()) {
      return NextResponse.json({ ok: false, error: '고객명과 연락처는 필수입니다' }, { status: 400 });
    }
    if (!visitDate || !visitTime) {
      return NextResponse.json({ ok: false, error: '날짜와 시간을 입력해주세요' }, { status: 400 });
    }
    if (type === 'field_request' && !addressRoad?.trim()) {
      return NextResponse.json({ ok: false, error: '출장 등록 시 주소는 필수입니다' }, { status: 400 });
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const phoneNorm = phone.replace(/\D/g, '');
    const today = new Date().toISOString().slice(0, 10);

    // ── 중복 체크: 오늘 이후 · 같은 번호 · 같은 일시 ──────────
    const { data: dupes } = await dbAny
      .from('consultations')
      .select('id, unique_id, name, phone, consultation_type, visit_date, visit_time, status')
      .eq('phone_normalized', phoneNorm)
      .gte('visit_date', today)
      .eq('visit_date', visitDate)
      .eq('visit_time', visitTime)
      .in('status', [
        'pending_admin',
        'assigned',
        'suggested',
        'confirmed',
        'reschedule_requested',
        'change_requested',
      ]);

    if (dupes && dupes.length > 0) {
      return NextResponse.json(
        {
          ok: false,
          error: 'duplicate',
          existing: dupes[0],
        },
        { status: 409 },
      );
    }

    // ── 좌표 (출장만) — lib/kakao/geocode 공통 helper 사용 ──
    let lat: number | null = null;
    let lng: number | null = null;
    if (type === 'field_request' && addressRoad) {
      const geo = await geocodeForConsultation(addressRoad, addressDetail);
      if (geo) {
        lat = geo.lat;
        lng = geo.lng;
      }
    }

    // ── INSERT ───────────────────────────────────────────────
    const uniqueId = crypto.randomUUID();
    const { data: consultation, error: insertErr } = await dbAny
      .from('consultations')
      .insert({
        name: name.trim(),
        phone: phone.trim(),
        consultation_type: type,
        visit_date: visitDate,
        visit_time: visitTime,
        address_road: addressRoad?.trim() || null,
        address_detail: addressDetail?.trim() || null,
        postcode: postcode?.trim() || null,
        latitude: lat,
        longitude: lng,
        status: 'confirmed', // 수기 등록은 즉시 확정
        memo: memo?.trim() || null,
        unique_id: uniqueId,
        gas_raw: {
          source: 'admin_manual', // 수기 접수 마킹 (리포트 집계용)
          registered_by: user.id,
          registered_at: new Date().toISOString(),
        },
        received_at: new Date().toISOString(),
      })
      .select()
      .single();

    if (insertErr) throw new Error(insertErr.message || JSON.stringify(insertErr));

    // 이력
    await dbAny.from('consultation_history').insert({
      consultation_id: consultation.id,
      to_status: 'confirmed',
      changed_by: user.id,
      note: '관리자 직접 등록',
    });

    // ── 알림톡 + Google Calendar 동기화 (after로 응답 후 실행 보장) ──
    after(async () => {
      // 알림톡 (notify=true일 때만)
      if (notify) {
        try {
          const isField = type === 'field_request';
          const fullAddress = [addressRoad, addressDetail].filter(Boolean).join(' ');
          const typeLabel = isField ? '출장 요청' : '매장 방문';

          await sendNotification({
            template: isField ? 'field_confirmed' : 'confirmed',
            phone: phoneNorm,
            name: name.trim(),
            data: {
              id: uniqueId,
              status: 'CONFIRMED',
              name: name.trim(),
              phone: phoneNorm,
              type: typeLabel,
              date: visitDate,
              time: visitTime,
              address: fullAddress,
              days: '',
              memo: memo || '',
              change_request_link: `${GITHUB_PAGES}/page_change_request.html?uid=${uniqueId}`,
            },
          });
        } catch (e) {
          console.error('[admin-create] 알림톡 발송 실패:', e);
        }
      }

      // Google Calendar 동기화
      try {
        await syncConsultationToCalendar(consultation.id);
      } catch (e) {
        console.error('[admin-create] 캘린더 동기화 실패:', e);
      }
    });

    return NextResponse.json({ ok: true, data: consultation });
  } catch (err) {
    console.error('[consultation/admin-create] 등록 실패:', err);
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
