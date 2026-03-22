import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { updateImwebStock } from '@/lib/imweb/client';

/** PATCH /api/sales/[id] — 취소 / 결제상태 변경 / 메모 수정 */
export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
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

    // 현재 판매 건 조회
    const { data: sale, error: saleErr } = await db
      .from('offline_sales')
      .select('*')
      .eq('id', id)
      .single();
    if (saleErr || !sale) {
      return NextResponse.json({ error: '판매 건을 찾을 수 없습니다' }, { status: 404 });
    }

    // --- A) 판매 취소 ---
    if (action === 'cancel') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '이미 취소된 판매입니다' }, { status: 400 });
      }

      const reason = body.reason || '';

      // 1. 시리얼 복원
      const { data: serials } = await db
        .from('product_serials')
        .select('id, product_id')
        .eq('offline_sale_id', id);

      if (serials && serials.length > 0) {
        await db
          .from('product_serials')
          .update({
            status: 'in_stock',
            sold_via: null,
            offline_sale_id: null,
            sold_at: null,
            sold_to_name: null,
            sold_to_phone: null,
          })
          .eq('offline_sale_id', id);
      }

      // 2. 재고 복원 + 3. 아임웹 재고 동기화
      const { data: items } = await db
        .from('offline_sale_items')
        .select('product_id, quantity')
        .eq('sale_id', id);

      if (items && items.length > 0) {
        for (const item of items) {
          if (!item.product_id || !item.quantity) continue;

          // 재고 복원
          const { data: prod } = await db
            .from('products')
            .select('stock_quantity, imweb_product_no')
            .eq('id', item.product_id)
            .single();

          if (prod) {
            const newStock = (prod.stock_quantity || 0) + item.quantity;
            await db
              .from('products')
              .update({ stock_quantity: newStock })
              .eq('id', item.product_id);

            // 아임웹 재고 동기화 (실패해도 취소 자체는 완료)
            if (prod.imweb_product_no) {
              try {
                await updateImwebStock(Number(prod.imweb_product_no), newStock);
              } catch (e) {
                console.error('[imweb] 취소 재고 동기화 실패:', prod.imweb_product_no, e);
              }
            }
          }
        }
      }

      // 4. 미수금 차감 (미결제/부분결제였던 경우)
      if (sale.customer_id && sale.payment_status !== 'paid') {
        const unpaidAmount = sale.total_amount - (sale.paid_amount || 0);
        if (unpaidAmount > 0) {
          const { data: cust } = await db
            .from('customers')
            .select('outstanding_balance')
            .eq('id', sale.customer_id)
            .single();
          if (cust) {
            await db
              .from('customers')
              .update({
                outstanding_balance: Math.max(0, (cust.outstanding_balance || 0) - unpaidAmount),
              })
              .eq('id', sale.customer_id);
          }
        }
      }

      // 5. 판매 상태 업데이트
      const { error: updateErr } = await db
        .from('offline_sales')
        .update({
          cancelled_at: new Date().toISOString(),
          cancelled_reason: reason,
          cancelled_by: user.id,
        })
        .eq('id', id);

      if (updateErr) throw updateErr;
      return NextResponse.json({ success: true, action: 'cancelled' });
    }

    // --- B) 결제상태 변경 ---
    if (action === 'update_payment') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '취소된 판매는 변경할 수 없습니다' }, { status: 400 });
      }

      const { payment_status, paid_amount } = body as {
        payment_status: string;
        paid_amount?: number;
      };

      if (!payment_status) {
        return NextResponse.json({ error: 'payment_status 필수' }, { status: 400 });
      }

      // 미수금 조정
      if (sale.customer_id) {
        const oldUnpaid = sale.payment_status !== 'paid'
          ? sale.total_amount - (sale.paid_amount || 0) : 0;
        const newPaidAmount = paid_amount ?? (payment_status === 'paid' ? sale.total_amount : sale.paid_amount);
        const newUnpaid = payment_status !== 'paid'
          ? sale.total_amount - newPaidAmount : 0;
        const diff = oldUnpaid - newUnpaid; // 양수면 미수금 감소

        if (diff !== 0) {
          const { data: cust } = await db
            .from('customers')
            .select('outstanding_balance')
            .eq('id', sale.customer_id)
            .single();
          if (cust) {
            await db
              .from('customers')
              .update({
                outstanding_balance: Math.max(0, (cust.outstanding_balance || 0) - diff),
              })
              .eq('id', sale.customer_id);
          }
        }
      }

      const updateData: Record<string, unknown> = { payment_status };
      if (paid_amount !== undefined) updateData.paid_amount = paid_amount;

      const { error: updateErr } = await db
        .from('offline_sales')
        .update(updateData)
        .eq('id', id);

      if (updateErr) throw updateErr;
      return NextResponse.json({ success: true, action: 'payment_updated' });
    }

    // --- C) 메모 수정 ---
    if (action === 'update_memo') {
      const { memo } = body as { memo: string };
      const { error: updateErr } = await db
        .from('offline_sales')
        .update({ memo: memo || null })
        .eq('id', id);

      if (updateErr) throw updateErr;
      return NextResponse.json({ success: true, action: 'memo_updated' });
    }

    return NextResponse.json({ error: '알 수 없는 action' }, { status: 400 });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : JSON.stringify(err);
    console.error('[sales PATCH] error:', message);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
