import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

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

    // 매장방문은 날짜/시간 필수
    if (type === '매장 방문' && (!visitDate || !visitTime)) {
      return NextResponse.json(
        { ok: false, error: '방문 날짜와 시간을 선택해주세요' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // 중복 접수 체크 (같은 전화번호 + 유형 + 날짜 + 시간이 pending/confirmed 상태)
    const phoneNorm = phone.replace(/\D/g, '');
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

    // 상태 결정: 매장방문(날짜 있음) → confirmed, 출장/톡 → pending_admin
    const initialStatus = (consultationType === 'store_visit' && visitDate && visitTime) ? 'confirmed' : 'pending_admin';

    // unique_id 생성
    const uniqueId = crypto.randomUUID();

    // 메모 구성 (진단 데이터 포함 가능)
    const memoText = memo?.trim() || null;

    // INSERT
    const insertData = {
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
      latitude: addressLat ? parseFloat(addressLat) : null,
      longitude: addressLng ? parseFloat(addressLng) : null,
      gas_raw: {
        days: Array.isArray(days) ? days.join(',') : (days || ''),
        timePrefs: Array.isArray(timePrefs) ? timePrefs.join(',') : (timePrefs || ''),
        addressZip: addressZip || '',
        addressRoad: addressRoad || '',
        addressDetail: addressDetail || '',
        addressLat: addressLat || '',
        addressLng: addressLng || '',
      },
      received_at: new Date().toISOString(),
    };

    const { data: consultation, error: insertErr } = await dbAny
      .from('consultations')
      .insert(insertData)
      .select()
      .single();

    if (insertErr) throw new Error(insertErr.message || JSON.stringify(insertErr));

    // 상태 이력 기록
    await dbAny.from('consultation_history').insert({
      consultation_id: consultation.id,
      to_status: initialStatus,
      note: '고객 접수',
    });

    // 알림톡 발송
    try {
      if (consultationType === 'store_visit' && initialStatus === 'confirmed') {
        // 매장방문 확정 → 확정 알림톡
        await sendNotification({
          template: 'confirmed',
          phone: phoneNorm,
          name: name.trim(),
          data: {
            id: uniqueId,
            type: '매장방문',
            date: visitDate || '',
            time: visitTime || '',
          },
        });
      } else if (consultationType === 'field_request') {
        // 출장요청 → 접수 안내 (관리자 확인 후 시간 제안)
        // GAS에서는 'request' 템플릿으로 발송했으나, TMS에서는 아직 해당 템플릿 미등록 시 스킵
      } else if (consultationType === 'talk_consult') {
        await sendNotification({
          template: 'talk_received',
          phone: phoneNorm,
          name: name.trim(),
          data: { id: uniqueId },
        });
      }
    } catch (notifyErr) {
      console.error('[consultation/submit] 알림톡 발송 실패 (접수는 완료):', notifyErr);
    }

    return NextResponse.json(
      { ok: true, data: { id: consultation.id, unique_id: uniqueId, status: initialStatus } },
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
