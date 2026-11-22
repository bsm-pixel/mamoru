import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { errMsg } from '@/lib/utils/err';
import { signUrl } from '@/lib/contracts/sign-link';

/**
 * POST /api/contracts/[id]/sign-link — 고객 서명 링크 만들기 (2026-11-22, 마이그 172)
 *  이미 있으면 같은 열쇠를 돌려준다(톡으로 보낸 주소가 바뀌지 않게). 취소된 계약은 거부.
 */
export async function POST(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    const { id } = await params;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: c, error } = await db.from('contracts').select('id, status, sign_token').eq('id', id).single();
    if (error || !c) {
      const noCol = /sign_token|column/i.test(String(error?.message || ''));
      return NextResponse.json({ error: noCol ? '서명 링크 칸이 아직 없습니다 — 데이터베이스 SQL 172 실행이 필요합니다' : '계약서를 찾을 수 없습니다' }, { status: noCol ? 500 : 404 });
    }
    if (c.status === 'cancelled') return NextResponse.json({ error: '취소된 계약서는 링크를 만들 수 없습니다' }, { status: 400 });
    if (c.sign_token) return NextResponse.json({ token: c.sign_token, url: signUrl(c.sign_token) });

    const token = crypto.randomUUID();
    const { error: upErr } = await db.from('contracts')
      .update({ sign_token: token, sign_link_at: new Date().toISOString() })
      .eq('id', id).is('sign_token', null);
    if (upErr) throw upErr;
    // 동시에 두 번 눌렸으면 먼저 저장된 열쇠를 쓴다
    const { data: again } = await db.from('contracts').select('sign_token').eq('id', id).single();
    const t = again?.sign_token || token;
    return NextResponse.json({ token: t, url: signUrl(t) });
  } catch (err) {
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}
