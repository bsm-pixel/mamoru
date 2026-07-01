import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * POST /api/products/[id]/renumber-sku
 * 제품 SKU를 현재 카테고리 코드에 맞게 재채번한다.
 *   - 카테고리 오지정 후 변경했는데 SKU 접두어가 옛 카테고리로 남은 경우 교정용.
 *   - 판매·시리얼·계약·아임웹 이력이 있으면 { needConfirm } 로 응답 → UI에서 확인 후 confirm:true 재요청.
 *   - 번호는 새 카테고리의 다음 번호로 서버가 계산(next-sku 와 동일 규칙).
 * body: { confirm?: boolean }
 */
export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json().catch(() => ({}));
    const confirm = body?.confirm === true;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: product, error: fetchErr } = await db
      .from('products')
      .select('id, sku, category, imweb_product_no')
      .eq('id', id)
      .single();
    if (fetchErr || !product) {
      return NextResponse.json({ error: '제품을 찾을 수 없습니다' }, { status: 404 });
    }

    const prefix = String(product.category || '').trim();
    if (!prefix) {
      return NextResponse.json({ error: '카테고리가 없어 SKU를 채번할 수 없습니다' }, { status: 400 });
    }

    // 이미 카테고리와 일치(접두어 + 숫자)하면 변경 불필요
    const sku = String(product.sku || '');
    if (sku.startsWith(prefix) && /^\d+$/.test(sku.slice(prefix.length))) {
      return NextResponse.json({ ok: true, unchanged: true, sku });
    }

    // 이력 확인 — 있으면 confirm 필요(라벨·연동 재확인 경고)
    const [serialsRes, saleItemsRes, contractItemsRes] = await Promise.all([
      db.from('product_serials').select('id', { count: 'exact', head: true }).eq('product_id', id),
      db.from('offline_sale_items').select('id', { count: 'exact', head: true }).eq('product_id', id),
      db.from('contract_items').select('id', { count: 'exact', head: true }).eq('product_id', id),
    ]);
    const linked: string[] = [];
    if ((serialsRes.count || 0) > 0) linked.push(`시리얼 ${serialsRes.count}개`);
    if ((saleItemsRes.count || 0) > 0) linked.push(`판매 ${saleItemsRes.count}건`);
    if ((contractItemsRes.count || 0) > 0) linked.push(`계약서 ${contractItemsRes.count}건`);
    if (product.imweb_product_no) linked.push('아임웹 연동');

    if (linked.length > 0 && !confirm) {
      return NextResponse.json({ ok: false, needConfirm: true, linked });
    }

    // 새 카테고리의 다음 SKU 계산 (next-sku 규칙과 동일)
    const { data: existing } = await db
      .from('products')
      .select('sku')
      .like('sku', `${prefix}%`)
      .order('sku', { ascending: false })
      .limit(1);
    let nextNum = 1;
    if (existing && existing.length > 0) {
      const parsed = parseInt(String(existing[0].sku).slice(prefix.length), 10);
      if (!isNaN(parsed)) nextNum = parsed + 1;
    }
    const nextSku = `${prefix}${String(nextNum).padStart(3, '0')}`;

    const oldSku = product.sku;
    const { error: updErr } = await db.from('products').update({ sku: nextSku }).eq('id', id);
    if (updErr) {
      return NextResponse.json({ error: updErr.message || '재채번 실패' }, { status: 400 });
    }

    return NextResponse.json({ ok: true, sku: nextSku, oldSku, prefix });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
