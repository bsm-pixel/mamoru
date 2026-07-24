import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/reviews/promise — 리뷰 약속 토글 + 유형 + 세부유형
 *  body: {
 *    source: 'consultation' | 'repair' | 'sale',
 *    id: string,
 *    on: boolean,
 *    type?: 'purchase' | 'repair' | 'consult',  // on=true 일 때 필수 (094)
 *    subtype?: 'direct_visit' | 'pickup' | 'store_visit' | 'field_request' | 'talk_consult' | null  // (095) repair/consult 일 때 선택
 *  }
 *  - on=true  → review_promised_at = now() + review_promised_type = type + review_promised_subtype = subtype
 *  - on=false → 모두 NULL
 *
 *  자동 발송 분기:
 *    type=repair   + subtype=direct_visit/pickup → as_review_request (subtype 전달)
 *    type=consult  + subtype=store_visit/field_request/talk_consult → review_request (subtype 전달)
 *    type=purchase                                                    → purchase_review_request (subtype X)
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
    // subtype 은 undefined / null / 빈문자 / 실제값 모두 가능
    const subtypeRaw = body.subtype as string | null | undefined;
    const subtype: string | null = (subtypeRaw === undefined || subtypeRaw === null || subtypeRaw === '') ? null : subtypeRaw;

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
    // subtype 유효성: type=repair → direct_visit|pickup|delivery(119, 직접발송=택배 수리), type=consult → store_visit|field_request|talk_consult, type=purchase → null 권장
    if (on && subtype !== null) {
      const repairSub = ['direct_visit', 'pickup', 'delivery'];
      const consultSub = ['store_visit', 'field_request', 'talk_consult'];
      if (type === 'repair' && !repairSub.includes(subtype)) {
        return NextResponse.json({ error: "repair subtype은 'direct_visit' | 'pickup' | 'delivery'" }, { status: 400 });
      }
      if (type === 'consult' && !consultSub.includes(subtype)) {
        return NextResponse.json({ error: "consult subtype은 'store_visit' | 'field_request' | 'talk_consult'" }, { status: 400 });
      }
      // type === 'purchase' → subtype 무시 (NULL 강제)
    }
    const effectiveSubtype: string | null = (on && type === 'purchase') ? null : subtype;

    const tableMap: Record<typeof source, string> = {
      consultation: 'consultations',
      repair: 'repairs',
      sale: 'offline_sales',
    };
    const table = tableMap[source];

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    if (on) {
      const { data: row, error: fetchErr } = await db
        .from(table)
        .select('id, review_promised_at')
        .eq('id', id)
        .single();
      if (fetchErr || !row) {
        return NextResponse.json({ error: '대상을 찾을 수 없습니다' }, { status: 404 });
      }
      const update: Record<string, unknown> = {
        review_promised_type: type,
        review_promised_subtype: effectiveSubtype,
      };
      if (!row.review_promised_at) {
        update.review_promised_at = new Date().toISOString();
      }
      await db.from(table).update(update).eq('id', id);
    } else {
      await db
        .from(table)
        .update({
          review_promised_at: null,
          review_promised_type: null,
          review_promised_subtype: null,
        })
        .eq('id', id);
    }

    return NextResponse.json({ ok: true });
  } catch (err) {
    console.error('[reviews/promise] Error:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
