import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';

const CRON_SECRET = process.env.CRON_SECRET || 'mamoru-tms-cron-2026';
const GITHUB_PAGES = 'page.mamoru.kr/projects/as'; // Make 시나리오가 https:// 추가

/** 'YYYY-MM-DD' → '7월 30일 (수)' */
function formatKoreanDate(dateStr: string): string {
  const d = new Date(`${dateStr}T00:00:00+09:00`);
  const days = ['일', '월', '화', '수', '목', '금', '토'];
  return `${d.getMonth() + 1}월 ${d.getDate()}일 (${days[d.getDay()]})`;
}

/**
 * GET /api/cron/repair-visit-remind — 직접방문(매장방문) 리마인드 자동 발송
 * Vercel Cron 10분 간격. 상담 send-reminders 패턴 복제(repairs 버전).
 * - 대상: proceed_type='직접방문' AND status='intake' AND 방문일 오늘~내일 AND visit_time 있음
 * - 24h(2~24h 전, 등록-방문 간격 24h+ 만 = 당일예약 제외) / 2h(0.5~2h 전)
 * - visit_remind_24h_sent_at / visit_remind_2h_sent_at 선(先)마킹으로 중복/동시실행 방지
 * - 일정 변경(resched) 시 두 플래그 NULL 리셋되어 새 일정에 재발송됨
 */
export async function GET(req: NextRequest) {
  const authHeader = req.headers.get('authorization');
  if (authHeader !== `Bearer ${CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const now = new Date();
    const kstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
    const today = kstNow.toISOString().slice(0, 10);
    const tomorrow = new Date(kstNow.getTime() + 24 * 60 * 60 * 1000).toISOString().slice(0, 10);

    const sent24h: string[] = [];
    const sent2h: string[] = [];
    const errors: string[] = [];

    const { data: repairs } = await dbAny
      .from('repairs')
      .select('id, as_id, name, phone, manage_token, visit_date, visit_time, qty_mamoru, qty_other, visit_remind_24h_sent_at, visit_remind_2h_sent_at, created_at')
      .eq('proceed_type', '직접방문')
      .eq('status', 'intake')
      .gte('visit_date', today)
      .lte('visit_date', tomorrow)
      .not('visit_time', 'is', null);

    if (!repairs || repairs.length === 0) {
      return NextResponse.json({ sent24h: 0, sent2h: 0, checked: 0 });
    }

    for (const r of repairs) {
      if (!r.visit_date || !r.visit_time) continue;

      const [h, m] = r.visit_time.split(':').map(Number);
      const visitAt = new Date(`${r.visit_date}T${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:00+09:00`);
      const diffHours = (visitAt.getTime() - now.getTime()) / (1000 * 60 * 60);
      if (diffHours < 0) continue;

      const qty = (r.qty_mamoru || 0) + (r.qty_other || 0);
      const changeLink = `${GITHUB_PAGES}/page_change_request.html?uid=${r.manage_token}`;
      const intervalHours = r.created_at
        ? (visitAt.getTime() - new Date(r.created_at).getTime()) / (1000 * 60 * 60)
        : 0;

      // 24h 리마인드 (2~24h 전) — 등록-방문 간격 24h 이상일 때만(당일 예약 제외)
      if (diffHours <= 24 && diffHours > 2 && !r.visit_remind_24h_sent_at && intervalHours >= 24) {
        try {
          const { count } = await dbAny
            .from('repairs')
            .update({ visit_remind_24h_sent_at: now.toISOString() })
            .eq('id', r.id)
            .is('visit_remind_24h_sent_at', null)
            .select('id', { count: 'exact', head: true });
          if (count === 0) continue; // 다른 Cron이 이미 마킹

          await sendNotification({
            template: 'as_visit_remind_24h',
            phone: r.phone,
            name: r.name,
            data: {
              as_id: r.as_id,
              visit_date: formatKoreanDate(r.visit_date),
              visit_time: r.visit_time,
              qty: String(qty),
              change_request_link: changeLink,
            },
          });
          sent24h.push(r.name);
        } catch (err) {
          errors.push(`24h ${r.name}: ${String(err)}`);
        }
      }

      // 2h 리마인드 (0.5~2h 전)
      if (diffHours <= 2 && diffHours > 0.5 && !r.visit_remind_2h_sent_at) {
        try {
          const { count } = await dbAny
            .from('repairs')
            .update({ visit_remind_2h_sent_at: now.toISOString() })
            .eq('id', r.id)
            .is('visit_remind_2h_sent_at', null)
            .select('id', { count: 'exact', head: true });
          if (count === 0) continue;

          await sendNotification({
            template: 'as_visit_remind_2h',
            phone: r.phone,
            name: r.name,
            data: {
              as_id: r.as_id,
              visit_date: formatKoreanDate(r.visit_date),
              visit_time: r.visit_time,
              qty: String(qty),
              change_request_link: changeLink,
            },
          });
          sent2h.push(r.name);
        } catch (err) {
          errors.push(`2h ${r.name}: ${String(err)}`);
        }
      }
    }

    return NextResponse.json({
      sent24h: sent24h.length,
      sent2h: sent2h.length,
      checked: repairs.length,
      errors: errors.length > 0 ? errors : undefined,
    });
  } catch (err) {
    console.error('[cron/repair-visit-remind] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
