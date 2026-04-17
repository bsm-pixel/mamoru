import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { updateImwebStock } from '@/lib/imweb/client';

/** GET /api/purchasing/[id] — 발주 상세 */
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

    const [orderRes, itemsRes] = await Promise.all([
      db.from('purchase_orders').select('*').eq('id', id).single(),
      db.from('purchase_order_items').select('*').eq('po_id', id),
    ]);

    if (orderRes.error) throw orderRes.error;

    let supplier = null;
    if (orderRes.data.supplier_id) {
      const { data: sup } = await db
        .from('customers')
        .select('id, name')
        .eq('id', orderRes.data.supplier_id)
        .single();
      supplier = sup;
    }

    return NextResponse.json({
      order: orderRes.data,
      items: itemsRes.data || [],
      supplier,
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/purchasing/[id] — 발주 상태 전환 + 수정 */
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
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 현재 발주 조회
    const { data: current, error: fetchErr } = await db
      .from('purchase_orders')
      .select('*')
      .eq('id', id)
      .single();

    if (fetchErr) throw fetchErr;

    const updates: Record<string, unknown> = {};
    const now = new Date().toISOString();

    // 상태 전환 처리
    if (body.status) {
      const newStatus = body.status;

      if (newStatus === 'ordered') {
        if (current.status !== 'draft') {
          return NextResponse.json({ error: 'draft 상태에서만 발주 확정 가능합니다' }, { status: 400 });
        }
        updates.status = 'ordered';
      } else if (newStatus === 'deposit_paid') {
        if (current.status !== 'ordered' && current.status !== 'draft') {
          return NextResponse.json({ error: `${current.status} 상태에서는 선납 처리할 수 없습니다` }, { status: 400 });
        }
        updates.status = 'deposit_paid';
        updates.deposit_amount = body.deposit_amount || current.deposit_amount;
        updates.deposit_paid_at = now;
        updates.balance_amount = current.total_amount - (body.deposit_amount || current.deposit_amount || 0);
      } else if (newStatus === 'received') {
        // 멱등성 가드: 이미 received 이상이면 재고 중복 증가 방지
        if (current.status === 'received' || current.status === 'balance_paid') {
          return NextResponse.json({ error: '이미 입고 처리된 발주입니다' }, { status: 409 });
        }
        // 유효 전이: ordered 또는 deposit_paid에서만 입고 가능
        if (current.status !== 'ordered' && current.status !== 'deposit_paid') {
          return NextResponse.json({ error: `${current.status} 상태에서는 입고 처리할 수 없습니다` }, { status: 400 });
        }

        updates.status = 'received';
        updates.received_date = body.received_date || now.slice(0, 10);

        // 입고 시 재고 증가
        const { data: poItems } = await db
          .from('purchase_order_items')
          .select('product_id, quantity')
          .eq('po_id', id);

        if (poItems) {
          for (const item of poItems) {
            if (item.product_id) {
              const { data: prod } = await db
                .from('products')
                .select('stock_quantity, raw_stock, imweb_product_no')
                .eq('id', item.product_id)
                .single();
              if (prod) {
                const newQty = (prod.stock_quantity || 0) + item.quantity;
                const newRaw = (prod.raw_stock || 0) + item.quantity; // 보관창고에 입고
                await db
                  .from('products')
                  .update({ stock_quantity: newQty, raw_stock: newRaw })
                  .eq('id', item.product_id);

                // 아임웹 재고 동기화 — 입고 시 증가(+)
                if (prod.imweb_product_no && newQty >= 0) {
                  try {
                    await updateImwebStock(Number(prod.imweb_product_no), +item.quantity);
                  } catch (e) {
                    console.error('[imweb] 재고 동기화 실패:', prod.imweb_product_no, e);
                  }
                }
              }
            }
          }
        }
      } else if (newStatus === 'balance_paid') {
        if (current.status !== 'received') {
          return NextResponse.json({ error: '입고 완료 후에만 잔금 처리 가능합니다' }, { status: 400 });
        }
        updates.status = 'balance_paid';
        updates.balance_paid_at = now;
        updates.balance_amount = 0;
      } else if (newStatus === 'cancelled') {
        if (current.status === 'received' || current.status === 'balance_paid') {
          return NextResponse.json({ error: '입고 이후에는 취소할 수 없습니다' }, { status: 400 });
        }
        updates.status = 'cancelled';
      } else {
        return NextResponse.json({ error: `유효하지 않은 상태입니다: ${newStatus}` }, { status: 400 });
      }
    }

    // 일반 필드 수정
    if (body.memo !== undefined) updates.memo = body.memo;
    if (body.expected_date !== undefined) updates.expected_date = body.expected_date;
    if (body.supplier_name !== undefined) updates.supplier_name = body.supplier_name;
    if (body.supplier_id !== undefined) updates.supplier_id = body.supplier_id || null;
    if (body.order_date !== undefined) updates.order_date = body.order_date;
    if (body.currency !== undefined) updates.currency = body.currency;
    if (body.exchange_rate !== undefined) updates.exchange_rate = body.exchange_rate;

    // 품목 수정 (draft 상태에서만)
    if (body.items && Array.isArray(body.items)) {
      if (current.status !== 'draft') {
        return NextResponse.json({ error: '작성중 상태에서만 품목 수정이 가능합니다' }, { status: 400 });
      }

      // 기존 품목 삭제
      await db.from('purchase_order_items').delete().eq('po_id', id);

      // 새 품목 삽입
      const newItems = body.items.map((item: { product_id?: string; product_name: string; sku?: string; quantity: number; unit_price: number }) => ({
        po_id: id,
        product_id: item.product_id || null,
        product_name: item.product_name,
        sku: item.sku || null,
        quantity: item.quantity,
        unit_price: item.unit_price,
        total_price: item.quantity * item.unit_price,
      }));

      if (newItems.length > 0) {
        await db.from('purchase_order_items').insert(newItems);
      }

      // 합계 재계산 (외화 → KRW 환산)
      const rate = updates.exchange_rate || current.exchange_rate || 1;
      const foreignTotal = newItems.reduce((s: number, i: { quantity: number; unit_price: number }) => s + i.quantity * i.unit_price, 0);
      const krwTotal = Math.round(foreignTotal * rate);
      updates.total_amount = krwTotal;
      updates.balance_amount = krwTotal - (current.deposit_amount || 0);
    }

    if (Object.keys(updates).length === 0) {
      return NextResponse.json({ error: '수정할 항목이 없습니다' }, { status: 400 });
    }

    updates.updated_at = now;

    const { data: order, error } = await db
      .from('purchase_orders')
      .update(updates)
      .eq('id', id)
      .select()
      .single();

    if (error) throw error;

    return NextResponse.json({ order });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/purchasing/[id] — 취소된 발주 삭제 */
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

    // 현재 상태 확인
    const { data: current } = await db.from('purchase_orders').select('status').eq('id', id).single();
    if (!current) return NextResponse.json({ error: '발주를 찾을 수 없습니다' }, { status: 404 });
    if (current.status !== 'cancelled') {
      return NextResponse.json({ error: '취소된 발주만 삭제할 수 있습니다' }, { status: 400 });
    }

    // 품목 삭제 → 발주 삭제
    await db.from('purchase_order_items').delete().eq('po_id', id);
    await db.from('purchase_orders').delete().eq('id', id);

    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
