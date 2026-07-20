import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { generateRackLocations } from '@/lib/warehouse/location-code';

/**
 * 창고 로케이션(정위치) API — 112, 2026-07-18
 *
 * GET    /api/warehouse/locations            로케이션 목록 (+ 칸별 배정 제품/재고 집계)
 * POST   /api/warehouse/locations            렉 자동생성 { rack_no, levels, bins? }
 * DELETE /api/warehouse/locations?id=…       로케이션 삭제 (배정된 제품은 ON DELETE SET NULL 로 '미지정'이 됨)
 *
 * ⚠️ 재고 수량(stock_quantity / raw_stock)은 읽기만 한다. 절대 쓰지 않는다.
 */

interface LocationRow {
  id: string; code: string; label: string | null;
  rack_no: number; level_no: number; bin_no: number | null;
  zone_type: string; sort_order: number; is_active: boolean; memo: string | null;
}
interface ProductLite {
  id: string; name: string; sku: string; category: string;
  stock_quantity: number; location_id: string | null;
}

/** GET — 로케이션 + 각 칸에 배정된 제품 요약 */
export async function GET() {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();

  const { data: locations, error: locErr } = await supabase
    .from('warehouse_locations')
    .select('*')
    .eq('is_active', true)
    .order('sort_order');
  if (locErr) return NextResponse.json({ error: locErr.message }, { status: 500 });

  // 위치가 배정된 활성 제품만 (재고는 읽기 전용)
  const { data: products, error: prodErr } = await supabase
    .from('products')
    .select('id, name, sku, category, stock_quantity, location_id')
    .eq('is_active', true)
    .order('name');
  if (prodErr) return NextResponse.json({ error: prodErr.message }, { status: 500 });

  // 칸별 집계 — 제품 수 / 재고 합
  const byLoc = new Map<string, { product_count: number; stock_total: number; products: ProductLite[] }>();
  const unassigned: ProductLite[] = [];
  for (const p of (products || []) as ProductLite[]) {
    if (!p.location_id) { unassigned.push(p); continue; }
    const cur = byLoc.get(p.location_id) || { product_count: 0, stock_total: 0, products: [] };
    cur.product_count += 1;
    cur.stock_total += p.stock_quantity || 0;
    cur.products.push(p);
    byLoc.set(p.location_id, cur);
  }

  const enriched = ((locations || []) as LocationRow[]).map((l) => {
    const agg = byLoc.get(l.id);
    return {
      ...l,
      product_count: agg?.product_count || 0,
      stock_total: agg?.stock_total || 0,
      products: agg?.products || [],
    };
  });

  return NextResponse.json({
    locations: enriched,
    // 위치 미지정 제품 — 배치도에서 칸에 담을 후보
    unassigned,
    unassigned_count: unassigned.length,
    total_locations: enriched.length,
  });
}

/** POST — 렉 자동생성 (단/칸을 펼쳐 일괄 INSERT) */
export async function POST(req: NextRequest) {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();
  try {
    const body = await req.json();
    const rackNo = Number(body.rack_no);
    const levels = Number(body.levels);
    const bins = body.bins == null || body.bins === '' ? 0 : Number(body.bins);
    const zoneType = (body.zone_type as string) || 'storage';

    if (!Number.isInteger(rackNo) || rackNo < 1 || rackNo > 99) {
      return NextResponse.json({ error: '렉 번호는 1~99 사이여야 합니다' }, { status: 400 });
    }
    if (!Number.isInteger(levels) || levels < 1 || levels > 20) {
      return NextResponse.json({ error: '단 수는 1~20 사이여야 합니다' }, { status: 400 });
    }
    if (!Number.isInteger(bins) || bins < 0 || bins > 26) {
      return NextResponse.json({ error: '칸 수는 0~26 사이여야 합니다 (0=칸 안 나눔)' }, { status: 400 });
    }

    const rows = generateRackLocations(rackNo, levels, bins).map((g) => ({
      code: g.code,
      label: g.label,
      rack_no: g.rack_no,
      level_no: g.level_no,
      bin_no: g.bin_no,
      sort_order: g.sort_order,
      zone_type: zoneType,
    }));

    // 같은 렉을 두 번 만들면 code UNIQUE 로 걸린다 → 친절한 메시지로 변환
    const { data, error } = await supabase
      .from('warehouse_locations')
      .insert(rows)
      .select('id, code');

    if (error) {
      if (String(error.code) === '23505') {
        return NextResponse.json({ error: `${rackNo}번 렉은 이미 등록돼 있습니다` }, { status: 409 });
      }
      throw error;
    }
    return NextResponse.json({ success: true, created: data?.length || 0, locations: data });
  } catch (err) {
    return NextResponse.json({ error: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }
}

/** DELETE — 로케이션 1건 삭제. 배정 제품은 products.location_id 가 NULL 로 풀린다(FK ON DELETE SET NULL) */
export async function DELETE(req: NextRequest) {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();
  const id = req.nextUrl.searchParams.get('id');
  const rackNo = req.nextUrl.searchParams.get('rack_no');

  if (!id && !rackNo) {
    return NextResponse.json({ error: 'id 또는 rack_no 가 필요합니다' }, { status: 400 });
  }

  // 삭제 전, 몇 개 제품의 위치가 풀리는지 미리 알려준다(사장님이 놀라지 않게)
  let query = supabase.from('warehouse_locations').delete();
  query = id ? query.eq('id', id) : query.eq('rack_no', Number(rackNo));

  const { error } = await query;
  if (error) return NextResponse.json({ error: error.message }, { status: 500 });
  return NextResponse.json({ success: true });
}
