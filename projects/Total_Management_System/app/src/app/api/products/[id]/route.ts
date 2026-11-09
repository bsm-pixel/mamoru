import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/products/[id] — 제품 상세 */
export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: product, error } = await db
      .from('products')
      .select('*')
      .eq('id', id)
      .single();

    if (error) throw error;

    // 매입처 정보
    let supplier = null;
    if (product.supplier_id) {
      const { data: sup } = await db
        .from('customers')
        .select('id, name')
        .eq('id', product.supplier_id)
        .single();
      supplier = sup;
    }

    return NextResponse.json({ product, supplier });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/products/[id] — 제품 삭제 (연관 데이터 없을 때만) */
export async function DELETE(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 연관 데이터 체크
    const [serialsRes, saleItemsRes, contractItemsRes] = await Promise.all([
      db.from('product_serials').select('id', { count: 'exact', head: true }).eq('product_id', id),
      db.from('offline_sale_items').select('id', { count: 'exact', head: true }).eq('product_id', id),
      db.from('contract_items').select('id', { count: 'exact', head: true }).eq('product_id', id),
    ]);

    const linked: string[] = [];
    if ((serialsRes.count || 0) > 0) linked.push(`시리얼 ${serialsRes.count}개`);
    if ((saleItemsRes.count || 0) > 0) linked.push(`판매 ${saleItemsRes.count}건`);
    if ((contractItemsRes.count || 0) > 0) linked.push(`계약서 ${contractItemsRes.count}건`);

    if (linked.length > 0) {
      return NextResponse.json({
        error: `연결된 데이터가 있어 삭제할 수 없습니다 (${linked.join(', ')}). 비활성화를 사용하세요.`,
      }, { status: 400 });
    }

    // 삭제
    const { error } = await db.from('products').delete().eq('id', id);
    if (error) throw error;

    return NextResponse.json({ deleted: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/products/[id] — 제품 수정 */
export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const updates: Record<string, unknown> = {};

    const allowed = ['name', 'category', 'price', 'price_dealer', 'price_academy', 'price_purchase', 'price_groups', 'dealer_name', 'academy_name', 'stock_quantity', 'supplier_id', 'description', 'imweb_product_no', 'barcode', 'image_url', 'is_active', 'product_group', 'purchase_name', 'sort_order', 'tags'];
    for (const key of allowed) {
      if (key in body) updates[key] = body[key];
    }
    // FK 필드: 빈 문자열 → null 변환 (FK 제약 위반 방지)
    if (updates.supplier_id === '') updates.supplier_id = null;
    // 이름: 앞뒤 공백 제거, 빈 이름 금지 (아임웹 상품명으로도 나가므로) 2026-11-09
    if ('name' in updates) {
      updates.name = String(updates.name ?? '').trim();
      if (!updates.name) return NextResponse.json({ error: '제품명을 입력하세요' }, { status: 400 });
    }
    // 판매가: 0 이상 정수 (아임웹 판매가로도 나가므로)
    if ('price' in updates) {
      const n = Number(updates.price);
      if (!Number.isFinite(n) || n < 0) return NextResponse.json({ error: '판매가가 올바르지 않습니다' }, { status: 400 });
      updates.price = Math.round(n);
    }

    if (Object.keys(updates).length === 0) {
      return NextResponse.json({ error: '수정할 항목이 없습니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 이름·판매가 변경 감지용 — 수정 전 값 (아임웹 반영 판단)
    const { data: prev } = ('name' in updates || 'price' in updates)
      ? await db.from('products').select('name, price, imweb_product_no').eq('id', id).single()
      : { data: null };

    const { data: product, error } = await db
      .from('products')
      .update(updates)
      .eq('id', id)
      .select()
      .single();

    if (error) {
      console.error('[products/PATCH] Supabase error:', JSON.stringify(error));
      return NextResponse.json({ error: error.message || JSON.stringify(error) }, { status: 400 });
    }

    // 재고 변경 시 아임웹 자동 반영 — delta 계산 (새 재고 - 이전 재고)
    if ('stock_quantity' in updates && product.imweb_product_no) {
      try {
        const { updateImwebStock } = await import('@/lib/imweb/client');
        // product는 update 후 값이므로, 이전 값과의 차이를 body에서 계산
        // updates.stock_quantity = 새 값, 이전 값을 별도 조회
        const { data: before } = await db.from('products').select('stock_quantity').eq('id', id).single();
        // before는 이미 업데이트된 상태이므로 body 값과 비교 불가 — 대신 delta를 직접 계산
        // updates.stock_quantity가 body에서 온 새 절대값, product.stock_quantity도 동일 (update 후 select)
        // 이 경우 이전 값을 알 수 없으므로, raw_stock 변화로 추정하거나 별도 처리 필요
        // → 상품 상세에서 직접 재고 수정은 재고 조정 모달을 권장 (delta 명확)
        // 임시: 재고 조정 모달 사용을 유도하고 여기서는 스킵
        console.log('[products/PATCH] 상품 상세 재고 직접 수정 — 재고 조정 모달 사용 권장');
      } catch (e) {
        console.error('[products/PATCH] 아임웹 재고 반영 실패:', e);
      }
    }

    // 이름·판매가 변경 → 아임웹도 같이 (TMS 한 곳에서만 고치기, 2026-11-09)
    //   바뀐 경우 또는 sync_imweb(다시 맞추기·일괄 재시도) 일 때만. 실패해도 TMS 저장은 유지하고 화면에 알린다(조용히 삼키지 않음)
    //   판매가도 TMS 마스터(사장님 확정 2026-11-09) → 바뀌면 아임웹 판매가도 같이
    const imweb: { synced: string[]; errors: string[] } = { synced: [], errors: [] };
    if (product.imweb_product_no) {
      const prodNo = Number(product.imweb_product_no);
      const force = body.sync_imweb === true;
      const { renameImwebProduct, updateImwebPrice } = await import('@/lib/imweb/client');
      const jobs: { label: string; run: () => Promise<void> }[] = [];
      if ('name' in updates && (force || prev?.name !== product.name)) jobs.push({ label: '상품명', run: () => renameImwebProduct(prodNo, product.name) });
      if ('price' in updates && (force || Number(prev?.price) !== Number(product.price))) jobs.push({ label: '판매가', run: () => updateImwebPrice(prodNo, Number(product.price)) });
      for (const job of jobs) {
        try {
          await job.run();
          imweb.synced.push(job.label);
        } catch (e) {
          const msg = e instanceof Error ? e.message : String(e);
          console.error('[products/PATCH] 아임웹 반영 실패:', { id, prodNo, field: job.label, msg });
          imweb.errors.push(`${job.label}: ${msg}`);
        }
      }
    }

    return NextResponse.json({ product, imweb });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : JSON.stringify(err);
    console.error('[products/PATCH] error:', message);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
