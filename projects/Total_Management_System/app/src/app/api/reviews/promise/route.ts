import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/reviews/promise — 리뷰 약속 토글 + 유형
 *  body: { source: 'consultation' | 'repair' | 'sale', id: string, on: boolean, type?: 'purchase'|'repair'|'consult' }
 *  - on=true  → review_promised_at = now() + review_promised_type = type
 *               (type 필수 — 자동 발송 시 어떤 솔라피 템플릿 쓸지 결정용. 094 마이그레이션)
 *  - on=false → review_promised_at = null, review_promised_type = null
 *
 *  자동 발송 분기 (cron/track-delivery + repair/[id]):
 *    type='repair'   → as_review_request 템플릿
 *    type='consult'  → review_request 템플릿
 *    type='purchase' → purchase_review_request 템플릿
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
    const type = body.type as 'purchase' | 'repair' | 'consult' | undefined;

    if (!['consultation', 'repair', 'sale'].includes(source)) {
      return NextResponse.json({ error: 'source는 consultation/repair/sale 중 하나' }, { status: 400 });
    }
    if (!id) {
      return NextResponse.json({ error: 'id 필수' }, { status: 400 });
    }
    if (on && !type) {
      return NextResponse.json({ error: 'on=true 일 때 type(purchase|repair|consult) 필수' }, { status: 400 });
    }
    if (on && !['purchase', 'repair', 'consult'].includes(type!)) {
      return NextResponse.json({ error: "type은 'purchase' | 'repair' | 'consult' 중 하나" }, { status: 400 });
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
      // ON: 약속 시각 + 유형 동시 set (유형은 이미 설정되어 있어도 새 값으로 갱신)
      const { data: row, error: fetchErr } = await db
        .from(table)
        .select('id, review_promised_at')
        .eq('id', id)
        .single();
      if (fetchErr || !row) {
        return NextResponse.json({ error: '대상을 찾을 수 없습니다' }, { status: 404 });
      }
      const update: Record<string, unknown> = { review_promised_type: type };
      if (!row.review_promised_at) {
        update.review_promised_at = new Date().toISOString();
      }
      await db.from(table).update(update).eq('id', id);
    } else {
      // OFF: 시각 + 유형 모두 NULL
      await db
        .from(table)
        .update({ review_promised_at: null, review_promised_type: null })
        .eq('id', id);
    }

    return NextResponse.json({ ok: true });
  } catch (err) {
    console.error('[reviews/promise] Error:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
