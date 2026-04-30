import { NextRequest, NextResponse, after } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';
import { sendAdminEmail } from '@/lib/notification/email';
import { syncConsultationToCalendar } from '@/lib/google/calendar-sync';
import { geocodeForConsultation } from '@/lib/kakao/geocode';
import { matchOrCreateCustomer } from '@/lib/customer/match-or-create';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

const GITHUB_PAGES = 'bsm-pixel.github.io/mamoru/projects/consulting'; // Make 시나리오가 https:// 추가

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** POST /api/consultation/public/submit — 고객 상담 접수 (비인증, CORS) */
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { name, phone, type, visitDate, visitTime, address, addressZip, addressRoad, addressDetail, addressLat, addressLng, days, timePrefs, memo } = body;

    // 필수값 검증
    if (!name?.trim() || !phone?.trim() || !type?.trim()) {
      return NextResponse.json(
        { ok: false, error: '이름, 연락처, 상담유형은 필수입니다' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    if (type === '매장 방문' && (!visitDate || !visitTime)) {
      return NextResponse.json(
        { ok: false, error: '방문 날짜와 시간을 선택해주세요' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const phoneNorm = phone.replace(/\D/g, '');

    // 중복 접수 체크
    if (type === '매장 방문' && visitDate && visitTime) {
      const { data: dup } = await dbAny
        .from('consultations')
        .select('id')
        .eq('phone_normalized', phoneNorm)
        .eq('consultation_type', 'store_visit')
        .eq('visit_date', visitDate)
        .eq('visit_time', visitTime)
        .in('status', ['pending_admin', 'confirmed'])
        .limit(1);
      if (dup && dup.length > 0) {
        return NextResponse.json(
          { ok: false, error: '이미 동일한 일정으로 접수된 내역이 있습니다' },
          { status: 409, headers: CORS_HEADERS }
        );
      }
    }

    // 상담유형 매핑
    const typeMap: Record<string, string> = {
      '매장 방문': 'store_visit',
      '출장 요청': 'field_request',
      '톡 상담': 'talk_consult',
    };
    const consultationType = typeMap[type] || 'store_visit';
    const initialStatus = (consultationType === 'store_visit' && visitDate && visitTime) ? 'confirmed' : 'pending_admin';
    const uniqueId = crypto.randomUUID();
    const memoText = memo?.trim() || null;

    // 희망 요일/시간대 배열 정리
    const daysArr = Array.isArray(days) ? days : (days ? String(days).split(',').filter(Boolean) : []);
    const timePrefsArr = Array.isArray(timePrefs) ? timePrefs : (timePrefs ? String(timePrefs).split(',').filter(Boolean) : []);
    const fullAddress = address || (addressRoad ? `${addressRoad} ${addressDetail || ''}`.trim() : '');

    // 좌표: 프론트에서 넘어오면 사용, 없으면 서버에서 geocoding (lib/kakao/geocode 공통 helper)
    // ⚠️ 회귀 방지: addressRoad만 1단계에 사용. fullAddress(상세 포함) 합치면 매칭 실패.
    let lat = addressLat ? parseFloat(addressLat) : null;
    let lng = addressLng ? parseFloat(addressLng) : null;
    if ((!lat || !lng) && addressRoad) {
      const geo = await geocodeForConsultation(addressRoad, addressDetail);
      if (geo) { lat = geo.lat; lng = geo.lng; }
    }

    // 고객 자동 매칭/생성 — phone 기준 SSOT
    const { customerId } = await matchOrCreateCustomer(dbAny, {
      phone: phone.trim(),
      name: name.trim(),
      source: 'consultation',
      extra: {
        addressRoad: addressRoad || null,
        addressDetail: addressDetail || null,
        postcode: addressZip || null,
      },
    });

    // INSERT
    const { data: consultation, error: insertErr } = await dbAny
      .from('consultations')
      .insert({
        customer_id: customerId,
        name: name.trim(),
        phone: phone.trim(),
        consultation_type: consultationType,
        visit_date: visitDate || null,
        visit_time: visitTime || null,
        address_road: addressRoad || null,
        address_detail: addressDetail || null,
        status: initialStatus,
        memo: memoText,
        unique_id: uniqueId,
        latitude: lat,
        longitude: lng,
        gas_raw: {
          days: daysArr.join(','),
          timePrefs: timePrefsArr.join(','),
          addressZip: addressZip || '',
          addressRoad: addressRoad || '',
          addressDetail: addressDetail || '',
          addressLat: addressLat || '',
          addressLng: addressLng || '',
        },
        received_at: new Date().toISOString(),
      })
      .select()
      .single();

    if (insertErr) throw new Error(insertErr.message || JSON.stringify(insertErr));

    // 상태 이력
    await dbAny.from('consultation_history').insert({
      consultation_id: consultation.id,
      to_status: initialStatus,
      note: '고객 접수',
    });

    // ── 알림톡 발송 (GAS postMake_ payload와 동일한 구조) ──
    let notifyResult: { success: boolean; error?: string } = { success: false, error: 'skipped' };
    try {
      if (consultationType === 'store_visit' && initialStatus === 'confirmed') {
        // 매장방문 확정 — GAS CONFIRMED payload 동일
        notifyResult = await sendNotification({
          template: 'confirmed',
          phone: phoneNorm,
          name: name.trim(),
          data: {
            id: uniqueId,
            status: initialStatus.toUpperCase(),
            name: name.trim(),
            phone: phoneNorm,
            type,
            date: visitDate || '',
            time: visitTime || '',
            address: fullAddress,
            days: daysArr.join(','),
            memo: memoText || '',
            change_request_link: `${GITHUB_PAGES}/page_change_request.html?uid=${uniqueId}`,
          },
        });
      } else if (consultationType === 'field_request') {
        // 출장요청 접수 — GAS CONSULT_REQUEST payload 동일
        notifyResult = await sendNotification({
          template: 'request',
          phone: phoneNorm,
          name: name.trim(),
          data: {
            id: uniqueId,
            status: initialStatus.toUpperCase(),
            name: name.trim(),
            phone: phoneNorm,
            type,
            date: visitDate || '',
            time: visitTime || '',
            address: fullAddress,
            days: daysArr.join(','),
            timePrefs: timePrefsArr.join(','),
            hope_days_text: daysArr.join(', '),
            hope_times_text: timePrefsArr.join(', '),
            memo: memoText || '',
          },
        });
      } else if (consultationType === 'talk_consult') {
        // 톡상담 접수 — GAS TALK_RECEIVED payload 동일
        notifyResult = await sendNotification({
          template: 'talk_received',
          phone: phoneNorm,
          name: name.trim(),
          data: {
            id: uniqueId,
            name: name.trim(),
            phone: phoneNorm,
            type,
          },
        });
      }
    } catch (notifyErr) {
      console.error('[consultation/submit] 알림톡 발송 실패 (접수는 완료):', notifyErr);
      notifyResult = { success: false, error: String(notifyErr) };
    }

    // ── Gmail 알림 발송 (GAS GmailApp.sendEmail 대체) ──
    let emailResult = false;
    try {
      const typeLabel = type || consultationType;
      const emailLines = [
        `■ ${typeLabel} 접수 알림`,
        ``,
        `이름: ${name.trim()}`,
        `연락처: ${phone}`,
        `유형: ${typeLabel}`,
      ];
      if (visitDate) emailLines.push(`날짜: ${visitDate}`);
      if (visitTime) emailLines.push(`시간: ${visitTime}`);
      if (fullAddress) emailLines.push(`주소: ${fullAddress}`);
      if (daysArr.length > 0) emailLines.push(`희망요일: ${daysArr.join(', ')}`);
      if (timePrefsArr.length > 0) emailLines.push(`희망시간대: ${timePrefsArr.join(', ')}`);
      if (memoText) emailLines.push(`메모: ${memoText}`);
      emailLines.push(``, `접수번호: ${uniqueId}`);

      emailResult = await sendAdminEmail(
        `[MAMORU 상담] 새로운 ${typeLabel} 접수`,
        emailLines.join('\n')
      );
    } catch (emailErr) {
      console.error('[consultation/submit] 이메일 발송 실패:', emailErr);
    }

    // Google Calendar 동기화 — after()로 응답 후 실행 보장 (Vercel 서버리스)
    if (initialStatus === 'confirmed') {
      after(async () => {
        try {
          await syncConsultationToCalendar(consultation.id);
        } catch (e) {
          console.error('[calendar-sync after submit] 실패:', e);
        }
      });
    }

    return NextResponse.json(
      {
        ok: true,
        data: { id: consultation.id, unique_id: uniqueId, status: initialStatus },
        _notifications: { alrimtalk: notifyResult, email: emailResult },
      },
      { headers: CORS_HEADERS }
    );
  } catch (err) {
    console.error('[consultation/public/submit] 접수 실패:', err);
    return NextResponse.json(
      { ok: false, error: err instanceof Error ? err.message : JSON.stringify(err) },
      { status: 500, headers: CORS_HEADERS }
    );
  }
}
