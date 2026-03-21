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
                .select('stock_quantity, imweb_product_no')
                .eq('id', item.product_id)
                .single();
              if (prod) {
                const newQty = (prod.stock_quantity || 0) + item.quantity;
                await db
                  .from('products')
                  .update({ stock_quantity: newQty })
                  .eq('id', item.product_id);

                // 아임웹 재고 동기화 (재고 미사용(-1) 상품은 스킵)
                if (prod.imweb_product_no && newQty >= 0) {
                  try {
                    await updateImwebStock(Number(prod.imweb_product_no), newQty);
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
