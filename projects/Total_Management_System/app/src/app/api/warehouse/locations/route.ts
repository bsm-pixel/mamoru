import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { generateRackLocations, makeLocationCode, makeLocationLabel, locationSortOrder, type LevelSpec } from '@/lib/warehouse/location-code';

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
/** 한 단의 셀 (행×열) */
interface Cell { id: string; bin_no: number | null; bin_row: number | null }

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

    /* ── 렉 생성 — 단별 {열, 행} 구조를 배열로 받는다 (행>1 이면 수납함) ── */
    if (action === 'create_rack') {
      const rackNo = Number(body.rack_no);
      const label = (body.label as string) || null;
      const levels: LevelSpec[] = Array.isArray(body.levels)
        ? body.levels.map((l: { cols?: unknown; rows?: unknown }) => ({
            cols: Number(l?.cols) || 0,
            rows: Math.max(1, Number(l?.rows) || 1),
          }))
        : [];
      const zoneType = (body.zone_type as string) || 'storage';

      if (!Number.isInteger(rackNo) || rackNo < 1 || rackNo > 99) {
        return NextResponse.json({ error: '렉 번호는 1~99 사이여야 합니다' }, { status: 400 });
      }
      if (levels.length < 1 || levels.length > 20) {
        return NextResponse.json({ error: '단은 1~20개까지 만들 수 있습니다' }, { status: 400 });
      }
      if (levels.some((l) => l.cols < 0 || l.cols > 26)) {
        return NextResponse.json({ error: '한 단의 열은 0~26개까지입니다 (0=칸 없이 선반)' }, { status: 400 });
      }
      if (levels.some((l) => (l.rows ?? 1) < 1 || (l.rows ?? 1) > 20)) {
        return NextResponse.json({ error: '한 단의 행은 1~20개까지입니다' }, { status: 400 });
      }
      const totalCells = levels.reduce((s, l) => s + (l.cols > 0 ? l.cols * (l.rows ?? 1) : 1), 0);
      if (totalCells > 600) {
        return NextResponse.json({ error: `자리가 너무 많습니다 (${totalCells}개). 600개 이하로 만들어주세요` }, { status: 400 });
      }

      // 렉 그리드 기준선 = 수납함이 아닌 단들 중 가장 열이 많은 값
      // (수납함은 단 전체를 차지하는 별도 블록으로 그려서 기준선에 영향을 주지 않는다)
      const simpleCols = levels.filter((l) => (l.rows ?? 1) === 1 && l.cols > 0).map((l) => l.cols);
      const columns = Math.max(1, ...(simpleCols.length ? simpleCols : [1]));

      const { error: rackErr } = await supabase
        .from('warehouse_racks')
        .insert({ rack_no: rackNo, label, columns, sort_order: rackNo });
      if (rackErr) {
        if (String(rackErr.code) === '23505') {
          return NextResponse.json({ error: `${rackNo}번 렉은 이미 등록돼 있습니다` }, { status: 409 });
        }
        throw rackErr;
      }

      const rows = generateRackLocations(rackNo, levels).map((g) => ({
        code: g.code, label: g.label, rack_no: g.rack_no, level_no: g.level_no,
        bin_no: g.bin_no, bin_row: g.bin_row, sort_order: g.sort_order, zone_type: zoneType,
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

    /* ── 특정 단에 열 또는 행 1개 추가 ──
       add_col : 모든 행에 열 하나씩 (단순 칸막이면 칸 1개 추가와 같음)
       add_row : 모든 열에 행 하나씩 (한 줄짜리 단을 수납함으로 만들 때) */
    if (action === 'add_col' || action === 'add_row') {
      const rackNo = Number(body.rack_no);
      const levelNo = Number(body.level_no);

      const { data: existing, error: exErr } = await supabase
        .from('warehouse_locations')
        .select('id, bin_no, bin_row')
        .eq('rack_no', rackNo)
        .eq('level_no', levelNo);
      if (exErr) throw exErr;
      if (!existing || existing.length === 0) {
        return NextResponse.json({ error: '해당 단을 찾을 수 없습니다' }, { status: 404 });
      }

      // '선반 통째'였다면 칸으로 나누기 전에 제품을 비워야 한다
      const shelfRow = existing.find((r: Cell) => r.bin_no == null);
      if (shelfRow) {
        const { count } = await supabase
          .from('products').select('id', { count: 'exact', head: true }).eq('location_id', shelfRow.id);
        if ((count || 0) > 0) {
          return NextResponse.json(
            { error: '이 단은 선반으로 쓰는 중이고 제품이 있습니다. 먼저 제품을 빼주세요.' },
            { status: 409 },
          );
        }
        await supabase.from('warehouse_locations').delete().eq('id', shelfRow.id);
      }

      const cells = (existing as Cell[]).filter((r) => r.bin_no != null);
      const curCols = cells.reduce((m, r) => Math.max(m, r.bin_no || 0), 0);
      const curRows = cells.reduce((m, r) => Math.max(m, r.bin_row || 1), 0);

      const nextCols = action === 'add_col' ? curCols + 1 : curCols || 1;
      const nextRows = action === 'add_row' ? Math.max(curRows, 1) + 1 : curRows || 1;
      if (nextCols > 26) return NextResponse.json({ error: '한 단에 26열까지입니다' }, { status: 400 });
      if (nextRows > 20) return NextResponse.json({ error: '한 단에 20행까지입니다' }, { status: 400 });


      // 새로 생기는 셀만 추가
      const toInsert: Record<string, unknown>[] = [];
      const has = new Set(cells.map((r) => `${r.bin_row || 1}:${r.bin_no}`));
      for (let r = 1; r <= nextRows; r++) {
        for (let c = 1; c <= nextCols; c++) {
          if (has.has(`${r}:${c}`)) continue;
          toInsert.push({
            code: makeLocationCode(rackNo, levelNo, c, r, nextRows),
            label: makeLocationLabel(rackNo, levelNo, c, r, nextRows),
            rack_no: rackNo, level_no: levelNo, bin_no: c, bin_row: r,
            sort_order: locationSortOrder(rackNo, levelNo, c, r),
          });
        }
      }

      // 행이 1 → 2 이상으로 늘면 기존 칸 코드에 행 번호가 붙어야 한다 (R01-2-A → R01-2-A1)
      if (nextRows > 1 && curRows <= 1) {
        for (const cell of cells) {
          await supabase.from('warehouse_locations').update({
            code: makeLocationCode(rackNo, levelNo, cell.bin_no, 1, nextRows),
            label: makeLocationLabel(rackNo, levelNo, cell.bin_no, 1, nextRows),
            bin_row: 1,
          }).eq('id', cell.id);
        }
      }

      if (toInsert.length) {
        const { error } = await supabase.from('warehouse_locations').insert(toInsert);
        if (error) throw error;
      }
      if (nextRows === 1) await widenRackIfNeeded(supabase, rackNo, nextCols);
      return NextResponse.json({ success: true, cols: nextCols, rows: nextRows, created: toInsert.length });
    }

    /* ── 렉에 단 1개 추가 (기본은 칸 없이 선반) ── */
    if (action === 'add_level') {
      const rackNo = Number(body.rack_no);
      const cols = Math.max(0, Math.min(26, Number(body.cols) || 0));
      const rowsCount = Math.max(1, Math.min(20, Number(body.rows) || 1));

      const { data: rows, error: exErr } = await supabase
        .from('warehouse_locations').select('level_no').eq('rack_no', rackNo);
      if (exErr) throw exErr;
      const maxLevel = (rows || []).reduce((m: number, r: { level_no: number }) => Math.max(m, r.level_no), 0);
      const nextLevel = maxLevel + 1;
      if (nextLevel > 20) return NextResponse.json({ error: '한 렉에 20단까지입니다' }, { status: 400 });

      const gen = generateRackLocations(rackNo, [{ cols, rows: rowsCount }]).map((g) => ({
        code: makeLocationCode(rackNo, nextLevel, g.bin_no, g.bin_row, rowsCount),
        label: makeLocationLabel(rackNo, nextLevel, g.bin_no, g.bin_row, rowsCount),
        rack_no: rackNo, level_no: nextLevel, bin_no: g.bin_no, bin_row: g.bin_row,
        sort_order: locationSortOrder(rackNo, nextLevel, g.bin_no, g.bin_row),
      }));
      const { error } = await supabase.from('warehouse_locations').insert(gen);
      if (error) throw error;

      if (cols > 0 && rowsCount === 1) await widenRackIfNeeded(supabase, rackNo, cols);
      return NextResponse.json({ success: true, level_no: nextLevel, created: gen.length });
    }

    return NextResponse.json({ error: '알 수 없는 action 입니다' }, { status: 400 });
  } catch (err) {
    return NextResponse.json({ error: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }
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
