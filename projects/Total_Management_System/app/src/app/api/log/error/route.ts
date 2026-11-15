import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { logError } from '@/lib/log/error-log';
import { errMsg } from '@/lib/utils/err';

/** GET /api/log/error — 최근 오류 · 실패 50건 (화면 오류 · 알림톡 발송 실패 · 동기화 실패) */
export async function GET() {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data, error } = await (supabase as any)
      .from('sync_log')
      .select('id, sync_type, status, error_message, started_at')
      .or('status.eq.failed,sync_type.like.error_%')
      .order('started_at', { ascending: false })
      .limit(50);
    if (error) throw error;
    return NextResponse.json({ errors: data || [] });
  } catch (err) {
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}

/** POST /api/log/error — 화면에서 난 오류를 남긴다 { message, where } */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json().catch(() => ({}));
    const message = String(body?.message || '').slice(0, 300);
    const where = String(body?.where || '').slice(0, 120);
    if (!message) return NextResponse.json({ error: '내용이 비어 있습니다' }, { status: 400 });
    await logError('screen', `${where} — ${message}`);
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}
