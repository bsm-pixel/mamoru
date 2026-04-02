import { NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/analytics/product-lifecycle — 제품 수명주기 분석
 *  시리얼 판매일 → 동일 고객 복원수리 접수일 간격 분석
 */
export async function GET() {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 1) 판매된 시리얼 (sold_at 있는 것)
    const { data: serials } = await db
      .from('product_serials')
      .select('id, serial_number, product_id, sold_at, sold_to_name, sold_to_phone, products:product_id(name, category)')
      .eq('status', 'sold')
      .not('sold_at', 'is', null)
      .limit(500);

    // 2) 복원수리 접수 건 (고객 전화번호 기준 매칭)
    const { data: repairs } = await db
      .from('repairs')
      .select('id, as_id, name, phone, created_at, qty_mamoru, qty_other')
      .not('status', 'eq', 'cancelled')
      .limit(500);

    // 3) 고객 전화번호 기반 매칭 (시리얼 판매 → 복원수리)
    const phoneRepairMap = new Map<string, Array<{ created_at: string }>>();
    for (const r of (repairs || [])) {
      const phone = (r.phone || '').replace(/\D/g, '');
      if (!phone) continue;
      if (!phoneRepairMap.has(phone)) phoneRepairMap.set(phone, []);
      phoneRepairMap.get(phone)!.push({ created_at: r.created_at });
    }

    // 4) 시리얼별 판매→수리 간격 계산
    type LifecycleItem = { serial: string; product_name: string; category: string; sold_at: string; repair_at: string | null; days: number | null };
    const items: LifecycleItem[] = [];

    for (const s of (serials || [])) {
      const phone = (s.sold_to_phone || '').replace(/\D/g, '');
      const product = s.products as { name: string; category: string } | null;
      const soldDate = new Date(s.sold_at);

      // 해당 고객의 복원수리 중 판매일 이후 가장 빠른 건
      const customerRepairs = phone ? (phoneRepairMap.get(phone) || []) : [];
      const afterSale = customerRepairs
        .filter((r) => new Date(r.created_at) > soldDate)
        .sort((a, b) => a.created_at.localeCompare(b.created_at));

      const firstRepair = afterSale[0] || null;
      const days = firstRepair
        ? Math.floor((new Date(firstRepair.created_at).getTime() - soldDate.getTime()) / (1000 * 60 * 60 * 24))
        : null;

      items.push({
        serial: s.serial_number,
        product_name: product?.name || '알 수 없음',
        category: product?.category || '',
        sold_at: s.sold_at,
        repair_at: firstRepair?.created_at || null,
        days,
      });
    }

    // 5) 제품별 평균 수명 집계
    const productStats: Record<string, { name: string; total_sold: number; repaired: number; avg_days: number; items: number[] }> = {};

    for (const item of items) {
      const key = item.product_name;
      if (!productStats[key]) productStats[key] = { name: key, total_sold: 0, repaired: 0, avg_days: 0, items: [] };
      productStats[key].total_sold++;
      if (item.days !== null) {
        productStats[key].repaired++;
        productStats[key].items.push(item.days);
      }
    }

    const summary = Object.values(productStats)
      .map((p) => ({
        name: p.name,
        total_sold: p.total_sold,
        repaired: p.repaired,
        repair_rate: p.total_sold > 0 ? Math.round((p.repaired / p.total_sold) * 100) : 0,
        avg_days: p.items.length > 0 ? Math.round(p.items.reduce((s, d) => s + d, 0) / p.items.length) : null,
        avg_months: p.items.length > 0 ? Math.round(p.items.reduce((s, d) => s + d, 0) / p.items.length / 30 * 10) / 10 : null,
      }))
      .filter((p) => p.total_sold >= 1)
      .sort((a, b) => b.total_sold - a.total_sold);

    // 전체 통계
    const allDays = items.filter((i) => i.days !== null).map((i) => i.days as number);
    const overallAvg = allDays.length > 0 ? Math.round(allDays.reduce((s, d) => s + d, 0) / allDays.length) : null;

    return NextResponse.json({
      overall: {
        total_sold: items.length,
        total_repaired: allDays.length,
        repair_rate: items.length > 0 ? Math.round((allDays.length / items.length) * 100) : 0,
        avg_days: overallAvg,
        avg_months: overallAvg ? Math.round(overallAvg / 30 * 10) / 10 : null,
      },
      by_product: summary,
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
