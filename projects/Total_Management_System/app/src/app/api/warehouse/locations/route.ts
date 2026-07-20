import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { generateRackLocations, makeLocationCode, makeLocationLabel, locationSortOrder } from '@/lib/warehouse/location-code';

/**
 * 창고 로케이션(정위치) API — 112 / 113(단별 칸 수), 2026-07-18
 *
 * GET    /api/warehouse/locations        렉 + 칸 + 칸별 배정 제품/재고
 * POST   /api/warehouse/locations        action: create_rack | add_bin | add_level
 * DELETE /api/warehouse/locations?id=…   칸 1개 삭제 / ?rack_no=… 렉 통째 삭제
 *
 * ⚠️ 재고 수량(stock_quantity / raw_stock)은 읽기만 한다. 절대 쓰지 않는다.
 */

interface ProductLite {
  id: string; name: string; sku: string; category: string;
  stock_quantity: number; location_id: string | null;
}

/** GET */
export async function GET() {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();

  const [racksRes, locsRes, prodRes] = await Promise.all([
    supabase.from('warehouse_racks').select('*').order('sort_order').order('rack_no'),
    supabase.from('warehouse_locations').select('*').eq('is_active', true).order('sort_order'),
    supabase.from('products').select('id, name, sku, category, stock_quantity, location_id').eq('is_active', true).order('name'),
  ]);

  if (racksRes.error) return NextResponse.json({ error: racksRes.error.message }, { status: 500 });
  if (locsRes.error) return NextResponse.json({ error: locsRes.error.message }, { status: 500 });
  if (prodRes.error) return NextResponse.json({ error: prodRes.error.message }, { status: 500 });

  // 칸별 집계
  const byLoc = new Map<string, { product_count: number; stock_total: number; products: ProductLite[] }>();
  const unassigned: ProductLite[] = [];
  for (const p of (prodRes.data || []) as ProductLite[]) {
    if (!p.location_id) { unassigned.push(p); continue; }
    const cur = byLoc.get(p.location_id) || { product_count: 0, stock_total: 0, products: [] };
    cur.product_count += 1;
    cur.stock_total += p.stock_quantity || 0;
    cur.products.push(p);
    byLoc.set(p.location_id, cur);
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const locations = (locsRes.data || []).map((l: any) => {
    const agg = byLoc.get(l.id);
    return { ...l, product_count: agg?.product_count || 0, stock_total: agg?.stock_total || 0, products: agg?.products || [] };
  });

  return NextResponse.json({
    racks: racksRes.data || [],
    locations,
    unassigned,
    unassigned_count: unassigned.length,
    total_locations: locations.length,
  });
}

/** POST — action 분기 */
export async function POST(req: NextRequest) {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();
  try {
    const body = await req.json();
    const action = body.action as string;

    /* ── 렉 생성 — 단별 칸 수를 배열로 받는다 ── */
    if (action === 'create_rack') {
      const rackNo = Number(body.rack_no);
      const label = (body.label as string) || null;
      const levelBins: number[] = Array.isArray(body.level_bins) ? body.level_bins.map((n: unknown) => Number(n) || 0) : [];
      const zoneType = (body.zone_type as string) || 'storage';

      if (!Number.isInteger(rackNo) || rackNo < 1 || rackNo > 99) {
        return NextResponse.json({ error: '렉 번호는 1~99 사이여야 합니다' }, { status: 400 });
      }
      if (levelBins.length < 1 || levelBins.length > 20) {
        return NextResponse.json({ error: '단은 1~20개까지 만들 수 있습니다' }, { status: 400 });
      }
      if (levelBins.some((b) => b < 0 || b > 26)) {
        return NextResponse.json({ error: '한 단의 칸 수는 0~26개까지입니다 (0=칸 없이 선반)' }, { status: 400 });
      }
      // 렉 열 수 = 가장 칸이 많은 단 기준 (그리드 기준선)
      const columns = Math.max(1, ...levelBins);

      const { error: rackErr } = await supabase
        .from('warehouse_racks')
        .insert({ rack_no: rackNo, label, columns, sort_order: rackNo });
      if (rackErr) {
        if (String(rackErr.code) === '23505') {
          return NextResponse.json({ error: `${rackNo}번 렉은 이미 등록돼 있습니다` }, { status: 409 });
        }
        throw rackErr;
      }

      const rows = generateRackLocations(rackNo, levelBins).map((g) => ({
        code: g.code, label: g.label, rack_no: g.rack_no, level_no: g.level_no,
        bin_no: g.bin_no, sort_order: g.sort_order, zone_type: zoneType,
      }));
      const { data, error } = await supabase.from('warehouse_locations').insert(rows).select('id');
      if (error) {
        // 칸 생성이 실패하면 렉만 남는 유령 상태가 되므로 되돌린다
        await supabase.from('warehouse_racks').delete().eq('rack_no', rackNo);
        if (String(error.code) === '23505') {
          return NextResponse.json({ error: `${rackNo}번 렉 자리가 이미 있습니다` }, { status: 409 });
        }
        throw error;
      }
      return NextResponse.json({ success: true, created: data?.length || 0, columns });
    }

    /* ── 특정 단에 칸 1개 추가 ── */
    if (action === 'add_bin') {
      const rackNo = Number(body.rack_no);
      const levelNo = Number(body.level_no);

      const { data: existing, error: exErr } = await supabase
        .from('warehouse_locations')
        .select('id, bin_no')
        .eq('rack_no', rackNo)
        .eq('level_no', levelNo);
      if (exErr) throw exErr;
      if (!existing || existing.length === 0) {
        return NextResponse.json({ error: '해당 단을 찾을 수 없습니다' }, { status: 404 });
      }

      // 이 단이 '선반 통째'(bin_no NULL)면 칸으로 나누기 전에 제품을 비워야 한다
      const shelfRow = existing.find((r: { bin_no: number | null }) => r.bin_no == null);
      if (shelfRow) {
        const { count } = await supabase
          .from('products')
          .select('id', { count: 'exact', head: true })
          .eq('location_id', shelfRow.id);
        if ((count || 0) > 0) {
          return NextResponse.json(
            { error: '이 단은 선반으로 쓰는 중이고 제품이 있습니다. 먼저 제품을 빼주세요.' },
            { status: 409 },
          );
        }
        await supabase.from('warehouse_locations').delete().eq('id', shelfRow.id);
      }

      const maxBin = existing.reduce((m: number, r: { bin_no: number | null }) => Math.max(m, r.bin_no || 0), 0);
      const nextBin = maxBin + 1;
      if (nextBin > 26) return NextResponse.json({ error: '한 단에 26칸까지입니다' }, { status: 400 });

      const totalLevels = await countLevels(supabase, rackNo);
      const { error } = await supabase.from('warehouse_locations').insert({
        code: makeLocationCode(rackNo, levelNo, nextBin),
        label: makeLocationLabel(rackNo, levelNo, nextBin, totalLevels),
        rack_no: rackNo, level_no: levelNo, bin_no: nextBin,
        sort_order: locationSortOrder(rackNo, levelNo, nextBin),
      });
      if (error) throw error;

      // 렉 열 수가 부족하면 넓힌다 (그리드 기준선 유지)
      await widenRackIfNeeded(supabase, rackNo, nextBin);
      return NextResponse.json({ success: true, bin_no: nextBin });
    }

    /* ── 렉에 단 1개 추가 (기본은 칸 없이 선반) ── */
    if (action === 'add_level') {
      const rackNo = Number(body.rack_no);
      const bins = Math.max(0, Math.min(26, Number(body.bins) || 0));

      const { data: rows, error: exErr } = await supabase
        .from('warehouse_locations').select('level_no').eq('rack_no', rackNo);
      if (exErr) throw exErr;
      const maxLevel = (rows || []).reduce((m: number, r: { level_no: number }) => Math.max(m, r.level_no), 0);
      const nextLevel = maxLevel + 1;
      if (nextLevel > 20) return NextResponse.json({ error: '한 렉에 20단까지입니다' }, { status: 400 });

      const gen = generateRackLocations(rackNo, [bins]).map((g) => ({
        code: makeLocationCode(rackNo, nextLevel, g.bin_no),
        label: makeLocationLabel(rackNo, nextLevel, g.bin_no, nextLevel),
        rack_no: rackNo, level_no: nextLevel, bin_no: g.bin_no,
        sort_order: locationSortOrder(rackNo, nextLevel, g.bin_no),
      }));
      const { error } = await supabase.from('warehouse_locations').insert(gen);
      if (error) throw error;

      if (bins > 0) await widenRackIfNeeded(supabase, rackNo, bins);
      return NextResponse.json({ success: true, level_no: nextLevel, created: gen.length });
    }

    return NextResponse.json({ error: '알 수 없는 action 입니다' }, { status: 400 });
  } catch (err) {
    return NextResponse.json({ error: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }
}

/** 렉의 단 수 (라벨 상/중/하 판단용) */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function countLevels(supabase: any, rackNo: number): Promise<number> {
  const { data } = await supabase.from('warehouse_locations').select('level_no').eq('rack_no', rackNo);
  const set = new Set<number>((data || []).map((r: { level_no: number }) => r.level_no));
  return set.size || 1;
}

/** 칸이 늘어 렉 열 수를 넘으면 열 수를 넓힌다 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function widenRackIfNeeded(supabase: any, rackNo: number, neededColumns: number) {
  const { data: rack } = await supabase.from('warehouse_racks').select('id, columns').eq('rack_no', rackNo).maybeSingle();
  if (rack && neededColumns > rack.columns) {
    await supabase.from('warehouse_racks').update({ columns: neededColumns }).eq('id', rack.id);
  }
}

/** DELETE — 칸 1개(id) 또는 렉 통째(rack_no) */
export async function DELETE(req: NextRequest) {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();
  const id = req.nextUrl.searchParams.get('id');
  const rackNo = req.nextUrl.searchParams.get('rack_no');

  if (!id && !rackNo) {
    return NextResponse.json({ error: 'id 또는 rack_no 가 필요합니다' }, { status: 400 });
  }

  if (id) {
    const { error } = await supabase.from('warehouse_locations').delete().eq('id', id);
    if (error) return NextResponse.json({ error: error.message }, { status: 500 });
    return NextResponse.json({ success: true });
  }

  // 렉 통째 삭제 — 칸 먼저(제품은 ON DELETE SET NULL 로 미지정), 그다음 렉 정보
  const { error: locErr } = await supabase.from('warehouse_locations').delete().eq('rack_no', Number(rackNo));
  if (locErr) return NextResponse.json({ error: locErr.message }, { status: 500 });
  const { error: rackErr } = await supabase.from('warehouse_racks').delete().eq('rack_no', Number(rackNo));
  if (rackErr) return NextResponse.json({ error: rackErr.message }, { status: 500 });
  return NextResponse.json({ success: true });
}
