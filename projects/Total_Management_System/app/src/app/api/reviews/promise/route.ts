import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/reviews/promise — 리뷰 약속 토글
 *  body: { source: 'consultation' | 'repair' | 'sale', id: string, on: boolean }
 *  - on=true  → review_promised_at = now() (이미 set이면 무시)
 *  - on=false → review_promised_at = null
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const source = body.source as 'consultation' | 'repair' | 'sale';
    const id = body.id as string;
    const on = body.on === true;

    if (!['consultation', 'repair', 'sale'].includes(source)) {
      return NextResponse.json({ error: 'source는 consultation/repair/sale 중 하나' }, { status: 400 });
    }
    if (!id) {
      return NextResponse.json({ error: 'id 필수' }, { status: 400 });
    }

    const tableMap: Record<typeof source, string> = {
      consultation: 'consultations',
      repair: 'repairs',
      sale: 'offline_sales',
    };
    const table = tableMap[source];

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    if (on) {
      // 이미 약속되어 있으면 그대로 둠 (멱등성)
      const { data: row, error: fetchErr } = await db
        .from(table)
        .select('id, review_promised_at')
        .eq('id', id)
        .single();
      if (fetchErr || !row) {
        return NextResponse.json({ error: '대상을 찾을 수 없습니다' }, { status: 404 });
      }
      if (!row.review_promised_at) {
        await db
          .from(table)
          .update({ review_promised_at: new Date().toISOString() })
          .eq('id', id);
      }
    } else {
      await db
        .from(table)
        .update({ review_promised_at: null })
        .eq('id', id);
    }

    return NextResponse.json({ ok: true });
  } catch (err) {
    console.error('[reviews/promise] Error:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
