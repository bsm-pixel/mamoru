import { NextRequest, NextResponse, after } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { isValidTransition } from '@/lib/consultation/transitions';
import { sendNotification } from '@/lib/notification/make-webhook';
import type { ConsultationStatus, ConsultationType } from '@/lib/supabase/types';

const GITHUB_PAGES = 'https://bsm-pixel.github.io/mamoru/projects/consulting';

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
      .select('status, consultation_type, name, phone, unique_id, address_road, address_detail, gas_raw')
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

    // 알림톡 발송 — GAS 대신 직접 (suggest 템플릿)
    const uid = data.unique_id || current.unique_id;
    const phoneNorm = (current.phone || '').replace(/\D/g, '');
    // 단축토큰 생성 (6자리 hex)
    const shortToken = Array.from(crypto.getRandomValues(new Uint8Array(3)))
      .map(b => b.toString(16).padStart(2, '0')).join('');

    // 단축토큰을 gas_raw에 저장 (page_suggest.html에서 사용)
    await db.from('consultations').update({
      gas_raw: {
        ...(current.gas_raw || {}),
        shortToken,
      },
    }).eq('id', consultationId);

    after(async () => {
      try {
        const confirmLink = `${GITHUB_PAGES}/page_suggest.html?t=${encodeURIComponent(shortToken)}`;
        const suggestText = suggestions.map((s, i) => `제안${i + 1}: ${s.date} ${s.time}`).join('\n');

        await sendNotification({
          template: 'suggest',
          phone: phoneNorm,
          name: current.name,
          data: {
            id: uid,
            name: current.name,
            phone: phoneNorm,
            type: '출장 요청',
            address: [current.address_road, current.address_detail].filter(Boolean).join(' '),
            suggest_text: suggestText,
            suggest_count: String(suggestions.length),
            confirm_link: confirmLink,
            // 개별 제안 시간 (Make 템플릿에서 사용)
            ...(suggestions[0] ? { date1: suggestions[0].date, time1: suggestions[0].time } : {}),
            ...(suggestions[1] ? { date2: suggestions[1].date, time2: suggestions[1].time } : {}),
            ...(suggestions[2] ? { date3: suggestions[2].date, time3: suggestions[2].time } : {}),
          },
        });
      } catch (notifyErr) {
        console.error('[suggest] 알림톡 발송 실패:', notifyErr);
      }
    });

    return NextResponse.json(data);
  } catch (err) {
    console.error('[suggest] 시간 제안 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
