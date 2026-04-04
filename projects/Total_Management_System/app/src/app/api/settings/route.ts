import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/settings — 전체 설정 조회 → { [key]: value } 맵 */
export async function GET() {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { data, error } = await (supabase as any)
      .from('system_settings')
      .select('key, value');
    if (error) throw error;

    const settings: Record<string, unknown> = {};
    for (const row of data || []) {
      settings[row.key] = row.value;
    }
    return NextResponse.json(settings);
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/settings — 부분 업데이트. body: { items: [{key, value}] } */
export async function PATCH(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { items } = (await req.json()) as { items: { key: string; value: unknown }[] };
    if (!items?.length) return NextResponse.json({ error: 'items required' }, { status: 400 });

    for (const { key, value } of items) {
      const { error } = await (supabase as any)
        .from('system_settings')
        .upsert(
          { key, value: typeof value === 'string' ? value : JSON.stringify(value), updated_at: new Date().toISOString() },
          { onConflict: 'key' },
        );
      if (error) throw error;
    }
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}