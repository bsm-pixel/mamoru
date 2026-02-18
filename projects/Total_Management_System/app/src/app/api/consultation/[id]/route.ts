import { NextRequest, NextResponse, after } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { isValidTransition } from '@/lib/consultation/transitions';
import { sendNotification, type NotifyTemplate } from '@/lib/notification/make-webhook';
import type { ConsultationStatus, ConsultationType } from '@/lib/supabase/types';

/** 상태→알림 템플릿 매핑 (자동 발송 대상) */
const AUTO_NOTIFY_MAP: Partial<Record<string, NotifyTemplate>> = {
  confirmed: 'confirmed',
  cancelled: 'cancelled',
};

/** GAS 웹앱에 취소 요청 — 캘린더 삭제 + 시트 상태 + 슬롯 캐시 무효화 */
async function cancelViaGAS(uniqueId: string): Promise<{ ok: boolean; detail?: string }> {
  const baseUrl = process.env.GAS_CONSULTING_URL;
  if (!baseUrl) {
    console.error('[GAS cancel] GAS_CONSULTING_URL 환경변수 미설정');
    return { ok: false, detail: 'GAS_CONSULTING_URL 미설정' };
  }
  if (!uniqueId) {
    console.error('[GAS cancel] uniqueId 없음');
    return { ok: false, detail: 'uniqueId 없음' };
  }
  try {
    const key = process.env.CRON_SECRET || 'mamoru-tms-cron-2026';
    // GAS 웹앱은 POST body를 외부에서 받지 못하는 이슈 → GET 쿼리 파라미터 방식 사용
    const params = new URLSearchParams({
      action: 'cancelConsultation',
      uid: uniqueId,
      key,
      skipNotify: 'true',
    });
    const url = `${baseUrl}?${params.toString()}`;
    console.log('[GAS cancel] 요청:', { uid: uniqueId });

    // GAS 웹앱은 302 리다이렉트로 응답 전달
    const res = await fetch(url, {
      method: 'GET',
      redirect: 'follow',
      signal: AbortSignal.timeout(15000),
    });

    const text = await res.text();
    let body;
    try { body = JSON.parse(text); } catch { body = text; }

    if (body?.ok === true) {
      console.log('[GAS cancel] 성공:', body);
      return { ok: true, detail: JSON.stringify(body) };
    }

    console.error('[GAS cancel] 실패 응답:', { status: res.status, body });
    return { ok: false, detail: `HTTP ${res.status}: ${JSON.stringify(body)}` };
  } catch (err) {
    console.error('[GAS cancel] fetch 에러:', err);
    return { ok: false, detail: String(err) };
  }
}

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

    // 현재 상담 조회
    const { data: current, error: fetchErr } = await db
      .from('consultations')
      .select('status, consultation_type')
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

    // 상태 변경 시 이력 기록
    if (newStatus && newStatus !== current.status) {
      await db.from('consultation_history').insert({
        consultation_id: id,
        from_status: current.status,
        to_status: newStatus,
        changed_by: user.id,
        note: note || null,
      });

      // 후속 작업: after()로 응답 반환 후 백그라운드 실행 — UI 빠른 응답
      after(async () => {
        const sideEffects: Promise<unknown>[] = [];

        // 취소 시 GAS 연동 — 구글 캘린더 삭제 + 아임웹 슬롯 열기
        if (newStatus === 'cancelled' && data.unique_id) {
          sideEffects.push(cancelViaGAS(data.unique_id));
        }

        // 자동 알림톡 발송
        const template = AUTO_NOTIFY_MAP[newStatus];
        if (template && data.phone) {
          const address = [data.address_road, data.address_detail].filter(Boolean).join(' ');
          const typeLabel = data.consultation_type === 'store_visit' ? '매장 방문'
            : data.consultation_type === 'field_request' ? '출장 요청' : '톡상담';
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
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
