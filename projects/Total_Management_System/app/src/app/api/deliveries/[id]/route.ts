import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { updateImwebStock } from '@/lib/imweb/client';
import { recalcOutstanding } from '@/lib/outstanding';
import { computeDeliveryTotals } from '@/lib/deliveries/totals';

/** GET /api/deliveries/[id] — 납품 상세 */
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

    const [dlRes, itemsRes] = await Promise.all([
      db.from('deliveries').select('*').eq('id', id).single(),
      db.from('delivery_items').select('*').eq('delivery_id', id),
    ]);

    if (dlRes.error) throw dlRes.error;

    return NextResponse.json({
      delivery: dlRes.data,
      items: itemsRes.data || [],
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/deliveries/[id] — 상태 전환 + 편집 */
export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();
    const { action } = body as { action: string };

    // 납품 건 조회
    const { data: dl, error: dlErr } = await db.from('deliveries').select('*').eq('id', id).single();
    if (dlErr || !dl) return NextResponse.json({ error: '납품 건을 찾을 수 없습니다' }, { status: 404 });

    // ── 납품 확정 (draft → confirmed): 재고 차감 ──
    if (action === 'confirm') {
      if (dl.status !== 'draft') return NextResponse.json({ error: '작성중 상태에서만 확정 가능합니다' }, { status: 400 });

      // 재고 차감
      const { data: items } = await db.from('delivery_items').select('product_id, quantity').eq('delivery_id', id);
      if (items) {
        const qtyMap: Record<string, number> = {};
        for (const item of items) {
          if (item.product_id) qtyMap[item.product_id] = (qtyMap[item.product_id] || 0) + item.quantity;
        }
        for (const [productId, qty] of Object.entries(qtyMap)) {
          const { data: prod } = await db.from('products').select('stock_quantity, raw_stock, imweb_product_no').eq('id', productId).single();
          if (!prod) continue;
          const newStock = Math.max(0, (prod.stock_quantity || 0) - qty);
          const newRaw = Math.max(0, (prod.raw_stock || 0) - qty);
          await db.from('products').update({ stock_quantity: newStock, raw_stock: newRaw }).eq('id', productId);

          // 아임웹 동기화 — 납품 확정 시 차감(-)
          if (prod.imweb_product_no) {
            try {
              await updateImwebStock(Number(prod.imweb_product_no), -qty);
            } catch (e) { console.error('[imweb] 납품 재고 동기화 실패:', e); }
          }
        }
      }

      await db.from('deliveries').update({ status: 'confirmed', updated_at: new Date().toISOString() }).eq('id', id);
      return NextResponse.json({ success: true, status: 'confirmed' });
    }

    // ── 출고 완료 (confirmed → shipped) ── 수동 처리 (집하 자동감지 실패 시 백업)
    if (action === 'ship') {
      if (dl.status !== 'confirmed') return NextResponse.json({ error: '납품확정 상태에서만 출고 가능합니다' }, { status: 400 });
      // 🐛 110 버그 수정: 기존엔 `body.tracking_number || null` 이라, 송장이 이미 발급된 건에
      //    수동 [출고 완료] 를 누르면 **송장번호가 null 로 지워졌다.**
      //    (전엔 송장 생성 즉시 shipped 라 이 버튼을 누를 일이 없어 안 드러났던 버그)
      const trackingNumber = body.tracking_number || dl.tracking_number || null;
      await db.from('deliveries').update({
        status: 'shipped',
        shipped_date: new Date().toISOString().slice(0, 10),
        tracking_number: trackingNumber,
        shipped_source: 'manual',   // 110: 집하 자동감지와 구분
        updated_at: new Date().toISOString(),
      }).eq('id', id);
      return NextResponse.json({ success: true, status: 'shipped' });
    }

    // ── 정산 완료 (shipped → settled) ──
    if (action === 'settle') {
      if (dl.status !== 'shipped') return NextResponse.json({ error: '출고완료 상태에서만 정산 가능합니다' }, { status: 400 });

      // 미수금: 아래 정산 반영 후 recalcOutstanding 로 재계산
      await db.from('deliveries').update({
        status: 'settled',
        payment_status: 'paid',
        // 납품 total_amount는 이미 net(할인 반영) → 완납액 = total (할인 재차감 시 과소수납)
        paid_amount: (dl.total_amount || 0),
        updated_at: new Date().toISOString(),
      }).eq('id', id);
      await recalcOutstanding(db, dl.customer_id);
      return NextResponse.json({ success: true, status: 'settled' });
    }

    // ── 취소 (draft/confirmed → cancelled) ──
    if (action === 'cancel') {
      if (dl.status === 'shipped' || dl.status === 'settled') {
        return NextResponse.json({ error: '출고/정산 완료 건은 취소할 수 없습니다' }, { status: 400 });
      }
      if (dl.cancelled_at) return NextResponse.json({ error: '이미 취소된 납품입니다' }, { status: 400 });

      const reason = body.reason || '';

      // confirmed였다면 재고 복원
      if (dl.status === 'confirmed') {
        const { data: items } = await db.from('delivery_items').select('product_id, quantity').eq('delivery_id', id);
        if (items) {
          for (const item of items) {
            if (!item.product_id) continue;
            const { data: prod } = await db.from('products').select('stock_quantity, raw_stock, imweb_product_no').eq('id', item.product_id).single();
            if (!prod) continue;
            const newStock = (prod.stock_quantity || 0) + item.quantity;
            const newRaw = (prod.raw_stock || 0) + item.quantity;
            await db.from('products').update({ stock_quantity: newStock, raw_stock: newRaw }).eq('id', item.product_id);

            if (prod.imweb_product_no) {
              try {
                await updateImwebStock(Number(prod.imweb_product_no), +item.quantity); // 취소 → 재고 복구(+)
              } catch (e) { console.error('[imweb] 납품취소 재고 동기화 실패:', e); }
            }
          }
        }
      }

      // 미수금: 아래 취소 반영 후 recalcOutstanding 로 재계산
      await db.from('deliveries').update({
        cancelled_at: new Date().toISOString(),
        cancelled_reason: reason,
        updated_at: new Date().toISOString(),
      }).eq('id', id);
      await recalcOutstanding(db, dl.customer_id);
      return NextResponse.json({ success: true, action: 'cancelled' });
    }

    // ── 결제상태 변경 ──
    if (action === 'update_payment') {
      const { payment_status: newStatus, paid_amount: newPaid } = body as {
        payment_status: string; paid_amount?: number;
      };
      // 납품 total_amount는 이미 net → 완납액 = total
      const newPaidVal = newStatus === 'paid' ? (dl.total_amount || 0) : (newPaid || 0);

      await db.from('deliveries').update({
        payment_status: newStatus,
        paid_amount: newPaidVal,
        updated_at: new Date().toISOString(),
      }).eq('id', id);

      // 미수금: 멱등 재계산 (±diff 누적 X)
      await recalcOutstanding(db, dl.customer_id);

      return NextResponse.json({ success: true });
    }

    // ── 메모/기타 필드 + 품목 편집 ──
    const updates: Record<string, unknown> = {};
    if (body.memo !== undefined) updates.memo = body.memo;
    if (body.expected_date !== undefined) updates.expected_date = body.expected_date;
    if (body.tracking_number !== undefined) updates.tracking_number = body.tracking_number;

    // 품목 편집 — draft 에서만. 재고는 '확정' 시 현재 delivery_items 기준으로 차감하므로
    //   draft 교체는 재고에 영향 없음(안전). 단 미수금(outstanding)은 생성 시 이미 반영됐으므로 총액 변동분 조정.
    if (Array.isArray(body.items)) {
      if (dl.status !== 'draft') {
        return NextResponse.json({ error: '작성중(draft) 상태에서만 품목을 수정할 수 있습니다' }, { status: 400 });
      }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const editItems = body.items as Array<any>;
      if (editItems.length === 0) {
        return NextResponse.json({ error: '품목이 비어 있습니다' }, { status: 400 });
      }
      // 기존 품목 교체
      await db.from('delivery_items').delete().eq('delivery_id', id);
      await db.from('delivery_items').insert(editItems.map((it) => ({
        delivery_id: id,
        product_id: it.product_id || null,
        product_name: it.product_name,
        sku: it.sku || null,
        category: it.category || null,
        quantity: it.quantity,
        unit_price: it.unit_price,
        total_price: it.quantity * it.unit_price,
      })));

      // 총액 재계산 (생성 POST 와 동일 공식 — 복원수리 RS는 VAT 제외)
      const vatTypeVal = (dl.vat_type || 'included') as 'included' | 'separate' | 'none';
      const discountVal = dl.discount_amount || 0;
      const { supplyAmount, vatAmount, totalAmount } = computeDeliveryTotals(editItems, vatTypeVal, discountVal);
      updates.total_amount = totalAmount;
      updates.supply_amount = supplyAmount;
      updates.vat_amount = vatAmount;

      // 미수금: 아래 deliveries 업데이트(총액 변동) 후 recalcOutstanding 로 재계산
    }

    if (Object.keys(updates).length > 0) {
      updates.updated_at = new Date().toISOString();
      await db.from('deliveries').update(updates).eq('id', id);
    }

    // 품목 편집으로 총액이 바뀌었으면 미수금 멱등 재계산
    if (Array.isArray(body.items) && dl.customer_id) {
      await recalcOutstanding(db, dl.customer_id);
    }

    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
