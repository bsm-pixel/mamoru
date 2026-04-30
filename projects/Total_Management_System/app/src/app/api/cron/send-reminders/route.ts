import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';

const CRON_SECRET = process.env.CRON_SECRET || 'mamoru-tms-cron-2026';

/**
 * GET /api/cron/send-reminders — 상담 리마인더 자동 발송
 * Vercel Cron: 10분 간격 실행
 *
 * GAS sendReminders_() 로직 이전:
 * - confirmed 상태 상담 중 visit_date가 24h/2h 이내인 건 조회
 * - remind_24h_at, remind_2h_at 컬럼으로 중복 발송 방지
 */
export async function GET(req: NextRequest) {
  // 인증 확인
  const authHeader = req.headers.get('authorization');
  if (authHeader !== `Bearer ${CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const now = new Date();
    const sent24h: string[] = [];
    const sent2h: string[] = [];
    const errors: string[] = [];

    // KST 기준 오늘 + 내일 (UTC 자정~09시 사이 날짜 오류 방지)
    const kstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
    const today = kstNow.toISOString().slice(0, 10);
    const tomorrow = new Date(kstNow.getTime() + 24 * 60 * 60 * 1000).toISOString().slice(0, 10);

    const { data: consultations } = await dbAny
      .from('consultations')
      .select('id, name, phone, consultation_type, visit_date, visit_time, unique_id, address_road, address_detail, remind_24h_at, remind_2h_at')
      .eq('status', 'confirmed')
      .gte('visit_date', today)
      .lte('visit_date', tomorrow)
      .not('visit_time', 'is', null);

    if (!consultations || consultations.length === 0) {
      return NextResponse.json({ sent24h: 0, sent2h: 0, checked: 0 });
    }

    for (const c of consultations) {
      if (!c.visit_date || !c.visit_time) continue;

      // 방문 시각 계산
      const [h, m] = c.visit_time.split(':').map(Number);
      const visitAt = new Date(`${c.visit_date}T${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:00+09:00`);
      const diffMs = visitAt.getTime() - now.getTime();
      const diffHours = diffMs / (1000 * 60 * 60);

      // 이미 지난 상담은 스킵
      if (diffHours < 0) continue;

      const isField = c.consultation_type === 'field_request';
      const address = [c.address_road, c.address_detail].filter(Boolean).join(' ');

      // 24h 리마인더 (2~24시간 전)
      if (diffHours <= 24 && diffHours > 2 && !c.remind_24h_at) {
        try {
          // DB 먼저 마킹 → 중복 발송 방지 (동시 Cron 실행 대비)
          const { count } = await dbAny
            .from('consultations')
            .update({ remind_24h_at: now.toISOString() })
            .eq('id', c.id)
            .is('remind_24h_at', null)
            .select('id', { count: 'exact', head: true });

          if (count === 0) continue; // 다른 Cron이 이미 마킹

          await sendNotification({
            // 매장: remind24 / 출장: field_remind_24h (Make 시나리오 필터와 정확히 일치)
            template: isField ? 'field_remind_24h' : 'remind24',
            phone: c.phone,
            name: c.name,
            data: {
              id: c.unique_id || c.id,
              date: c.visit_date,
              time: c.visit_time,
              type: isField ? '출장' : '매장방문',
              address,                  // 출장 리마인더 방문주소 치환 변수
            },
          });

          sent24h.push(c.name);
        } catch (err) {
          errors.push(`24h ${c.name}: ${String(err)}`);
        }
      }

      // 2h 리마인더 (0.5~2시간 전)
      if (diffHours <= 2 && diffHours > 0.5 && !c.remind_2h_at) {
        try {
          // DB 먼저 마킹 → 중복 발송 방지
          const { count } = await dbAny
            .from('consultations')
            .update({ remind_2h_at: now.toISOString() })
            .eq('id', c.id)
            .is('remind_2h_at', null)
            .select('id', { count: 'exact', head: true });

          if (count === 0) continue; // 다른 Cron이 이미 마킹

          await sendNotification({
            // 매장: remind2 / 출장: field_remind_2h (Make 시나리오 필터와 정확히 일치)
            template: isField ? 'field_remind_2h' : 'remind2',
            phone: c.phone,
            name: c.name,
            data: {
              id: c.unique_id || c.id,
              date: c.visit_date,
              time: c.visit_time,
              type: isField ? '출장' : '매장방문',
              address,                  // 출장 리마인더 방문주소 치환 변수
            },
          });

          sent2h.push(c.name);
        } catch (err) {
          errors.push(`2h ${c.name}: ${String(err)}`);
        }
      }
    }

    return NextResponse.json({
      sent24h: sent24h.length,
      sent2h: sent2h.length,
      checked: consultations.length,
      errors: errors.length > 0 ? errors : undefined,
    });
  } catch (err) {
    console.error('[cron/send-reminders] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
