import { NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** GET /api/consultation/public/settings — 영업시간 + 휴무 설정 (비인증) */
export async function GET() {
  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // 상담 설정 조회
    const { data: settings } = await dbAny
      .from('consultation_settings')
      .select('*')
      .eq('id', 'default')
      .single();

    // 휴무일 조회 (오늘 이후만)
    const today = new Date().toISOString().slice(0, 10);
    const { data: closedDates } = await dbAny
      .from('closed_dates')
      .select('date, reason')
      .gte('date', today)
      .order('date', { ascending: true });

    const result = {
      ok: true,
      data: {
        BUSINESS: {
          startHour: settings?.start_hour ?? 10,
          endHour: settings?.end_hour ?? 20,
          durMin: settings?.duration_min ?? 60,
          stepMin: settings?.step_min ?? 10,
        },
        CLOSED_WEEKDAYS: settings?.disabled_weekdays ?? [0],
        CLOSED_DATES: (closedDates || []).map((d: { date: string }) => d.date),
        FIELD_BUFFER: {
          before: settings?.field_buffer_before ?? 90,
          after: settings?.field_buffer_after ?? 90,
        },
      },
    };

    return NextResponse.json(result, { headers: CORS_HEADERS });
  } catch (err) {
    console.error('[consultation/settings] 조회 실패:', err);
    return NextResponse.json(
      { ok: false, error: String(err) },
      { status: 500, headers: CORS_HEADERS }
    );
  }
}
