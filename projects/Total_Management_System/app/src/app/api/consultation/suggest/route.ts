import { NextRequest, NextResponse, after } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { isValidTransition } from '@/lib/consultation/transitions';
import type { ConsultationStatus, ConsultationType } from '@/lib/supabase/types';

/** GAS 웹앱에 시간제안 요청 — 캘린더 HOLD + 시트 상태 + 슬롯 차단 + 알림톡 */
async function suggestViaGAS(
  uniqueId: string,
  suggestions: { date: string; time: string }[]
): Promise<{ ok: boolean; detail?: string }> {
  const baseUrl = process.env.GAS_CONSULTING_URL;
  if (!baseUrl) {
    console.error('[GAS suggest] GAS_CONSULTING_URL 환경변수 미설정');
    return { ok: false, detail: 'GAS_CONSULTING_URL 미설정' };
  }
  if (!uniqueId) {
    console.error('[GAS suggest] uniqueId 없음');
    return { ok: false, detail: 'uniqueId 없음' };
  }
  try {
    const key = process.env.CRON_SECRET || 'mamoru-tms-cron-2026';
    const params = new URLSearchParams({
      action: 'suggestTimes',
      uid: uniqueId,
      key,
      suggestions: JSON.stringify(suggestions),
    });
    const url = `${baseUrl}?${params.toString()}`;
    console.log('[GAS suggest] 요청:', { uid: uniqueId, count: suggestions.length });

    const res = await fetch(url, {
      method: 'GET',
      redirect: 'follow',
      signal: AbortSignal.timeout(20000), // HOLD 생성 포함이므로 넉넉히
    });

    const text = await res.text();
    let body;
    try { body = JSON.parse(text); } catch { body = text; }

    if (body?.ok === true) {
      console.log('[GAS suggest] 성공:', body);
      return { ok: true, detail: JSON.stringify(body) };
    }

    console.error('[GAS suggest] 실패 응답:', { status: res.status, body });
    return { ok: false, detail: `HTTP ${res.status}: ${JSON.stringify(body)}` };
  } catch (err) {
    console.error('[GAS suggest] fetch 에러:', err);
    return { ok: false, detail: String(err) };
  }
}

/** POST /api/consultation/suggest — 시간 제안 (출장요청용) */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const body = await req.json();
    const { consultationId, suggestions } = body as {
      consultationId: string;
      suggestions: { date: string; time: string }[];
    };

    if (!consultationId || !suggestions?.length) {
      return NextResponse.json({ error: 'consultationId, suggestions 필수' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data: current, error: fetchErr } = await db
      .from('consultations')
      .select('status, consultation_type, name, phone')
      .eq('id', consultationId)
      .single();

    if (fetchErr || !current) {
      return NextResponse.json({ error: '상담을 찾을 수 없습니다' }, { status: 404 });
    }

    // 상태 전이 검증: → suggested
    const canTransition = isValidTransition(
      current.consultation_type as ConsultationType,
      current.status as ConsultationStatus,
      'suggested'
    );
    if (!canTransition) {
      return NextResponse.json(
        { error: `현재 상태(${current.status})에서 시간 제안 불가` },
        { status: 400 }
      );
    }

    // suggestions JSONB 저장 + status → suggested
    const { data, error } = await db
      .from('consultations')
      .update({
        suggestions: { dates: suggestions },
        status: 'suggested',
      })
      .eq('id', consultationId)
      .select()
      .single();

    if (error) throw error;

    // 이력 기록
    await db.from('consultation_history').insert({
      consultation_id: consultationId,
      from_status: current.status,
      to_status: 'suggested',
      changed_by: user.id,
      note: `시간 제안: ${suggestions.map((s) => `${s.date} ${s.time}`).join(', ')}`,
    });

    // GAS 연동 — 캘린더 HOLD + 시트 상태 SUGGESTED + 슬롯 차단 + 알림톡 발송
    // 백그라운드 실행 (UI 빠른 응답)
    after(async () => {
      if (data.unique_id) {
        const result = await suggestViaGAS(data.unique_id, suggestions);
        if (!result.ok) {
          console.error('[suggest] GAS 연동 실패:', result.detail);
        }
      } else {
        console.warn('[suggest] unique_id 없음 — GAS 연동 건너뜀 (id:', consultationId, ')');
      }
    });

    return NextResponse.json(data);
  } catch (err) {
    console.error('[suggest] 시간 제안 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
