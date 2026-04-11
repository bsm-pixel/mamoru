import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/serials/move — 창고 간 시리얼 이동 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { product_id, from_zone, to_zone, count } = await req.json() as {
      product_id: string;
      from_zone: 'raw' | 'ready' | 'display';
      to_zone: 'raw' | 'ready' | 'display';
      count: number;
    };

    if (!product_id || !from_zone || !to_zone || !count || count < 1) {
      return NextResponse.json({ error: '필수 값이 누락되었습니다' }, { status: 400 });
    }
    if (from_zone === to_zone) {
      return NextResponse.json({ error: '출발지와 도착지가 같습니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // ── 시리얼 창고 → 보관 (역이동: 시리얼 삭제 + raw_stock 증가) ──
    if (to_zone === 'raw') {
      // from_zone에서 해당 제품의 in_stock 시리얼 count개 조회
      const { data: serials } = await db
        .from('product_serials')
        .select('id')
        .eq('product_id', product_id)
        .eq('status', 'in_stock')
        .eq('warehouse_zone', from_zone)
        .limit(count);

      if (!serials || serials.length < count) {
        return NextResponse.json(
          { error: `${from_zone === 'ready' ? '준비' : '디스플레이'} 창고에 이동 가능한 시리얼이 ${serials?.length || 0}개뿐입니다` },
          { status: 400 }
        );
      }

      // 시리얼 삭제
      const ids = serials.map((s: { id: string }) => s.id);
      await db.from('product_serials').delete().in('id', ids);

      // raw_stock 증가 (stock_quantity는 유지 — 시리얼 삭제로 자동 감소분 = raw_stock 증가분)
      const { data: prod } = await db.from('products').select('raw_stock').eq('id', product_id).single();
      if (prod) {
        await db.from('products').update({ raw_stock: (prod.raw_stock || 0) + count }).eq('id', product_id);
      }

      return NextResponse.json({ success: true, moved: count, action: 'reverse' });
    }

    // ── 시리얼 창고 간 이동 (ready↔display: zone 변경만) ──
    if (from_zone !== 'raw') {
      const { data: serials } = await db
        .from('product_serials')
        .select('id')
        .eq('product_id', product_id)
        .eq('status', 'in_stock')
        .eq('warehouse_zone', from_zone)
        .limit(count);

      if (!serials || serials.length < count) {
        return NextResponse.json(
          { error: `${from_zone === 'ready' ? '준비' : '디스플레이'} 창고에 이동 가능한 시리얼이 ${serials?.length || 0}개뿐입니다` },
          { status: 400 }
        );
      }

      const ids = serials.map((s: { id: string }) => s.id);
      await db.from('product_serials').update({ warehouse_zone: to_zone }).in('id', ids);

      return NextResponse.json({ success: true, moved: count, action: 'zone_transfer' });
    }

    // from_zone === 'raw' → 보관에서 시리얼 창고로 이동은 /api/serials/batch에서 처리
    return NextResponse.json({ error: '보관→시리얼 창고 이동은 시리얼 생성 API를 사용하세요' }, { status: 400 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
