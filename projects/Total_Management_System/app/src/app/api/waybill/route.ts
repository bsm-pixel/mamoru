import { NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/waybill — 운송장 번호 잔여 현황 */
export async function GET() {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data: config, error } = await db
      .from('lotte_waybill_config')
      .select('current_number, start_number, end_number, updated_at')
      .eq('id', 'default')
      .single();

    if (error || !config) {
      return NextResponse.json({ error: '송장번호 설정이 없습니다' }, { status: 404 });
    }

    const current = Number(config.current_number);
    const start = Number(config.start_number);
    const end = Number(config.end_number);
    const total = end - start + 1;
    const used = current - start;
    const remaining = end - current + 1;

    return NextResponse.json({
      current_number: current,
      start_number: start,
      end_number: end,
      total,
      used,
      remaining,
      updated_at: config.updated_at,
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
