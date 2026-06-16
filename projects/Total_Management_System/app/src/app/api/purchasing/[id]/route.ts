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
      } else if (newStatus === 'received' || newStatus === 'partial') {
        // 입고 처리 — newStatus='partial'(분할: 발주 열어둠) / 'received'(완료: 나머지 안 옴, 받은 수량 기준 종료)
        // received_items = 이번에 받은 수량(배치) [{ id, received_quantity }]. received_quantity 컬럼은 누적값으로 저장.
        if (current.status === 'received' || current.status === 'balance_paid') {
          return NextResponse.json({ error: '이미 입고 완료된 발주입니다' }, { status: 409 });
        }
        // 유효 전이: ordered / deposit_paid / partial 에서만 입고 가능
        if (current.status !== 'ordered' && current.status !== 'deposit_paid' && current.status !== 'partial') {
          return NextResponse.json({ error: `${current.status} 상태에서는 입고 처리할 수 없습니다` }, { status: 400 });
        }

        // 이번 배치 수량 맵
        const batchMap: Record<string, number> = {};
        if (Array.isArray(body.received_items)) {
          for (const ri of body.received_items) {
            if (ri && ri.id != null) batchMap[String(ri.id)] = Math.max(0, Math.floor(Number(ri.received_quantity) || 0));
          }
        }

        const { data: poItems } = await db
          .from('purchase_order_items')
          .select('id, product_id, quantity, unit_price, received_quantity')
          .eq('po_id', id);

        let foreignTotalCum = 0; // 누적 입고 기준 합계
        let allFull = true;      // 모든 품목이 주문 수량까지 입고됐는가
        let anyAdjusted = false; // 최종 누적이 주문과 다른가
        if (poItems) {
          for (const item of poItems) {
            const prev = item.received_quantity || 0;
            // 명시값 있으면 그 배치, 없으면 완료(received)는 남은 전량, 분할(partial)은 0
            let batch: number;
            if (Object.prototype.hasOwnProperty.call(batchMap, item.id)) batch = batchMap[item.id];
            else batch = newStatus === 'received' ? Math.max(0, item.quantity - prev) : 0;
            const cum = prev + batch;
            foreignTotalCum += cum * (item.unit_price || 0);
            if (cum < item.quantity) allFull = false;
            if (cum !== item.quantity) anyAdjusted = true;

            // 누적 received_quantity 기록
            await db.from('purchase_order_items').update({ received_quantity: cum }).eq('id', item.id);

            // 재고는 이번 배치만큼만 증가 (중복 방지)
            if (item.product_id && batch > 0) {
              const { data: prod } = await db
                .from('products')
                .select('stock_quantity, raw_stock, imweb_product_no')
                .eq('id', item.product_id)
                .single();
              if (prod) {
                await db.from('products')
                  .update({ stock_quantity: (prod.stock_quantity || 0) + batch, raw_stock: (prod.raw_stock || 0) + batch })
                  .eq('id', item.product_id);
                if (prod.imweb_product_no) {
                  try { await updateImwebStock(Number(prod.imweb_product_no), +batch); }
                  catch (e) { console.error('[imweb] 재고 동기화 실패:', prod.imweb_product_no, e); }
                }
              }
            }
          }
        }

        // 분할인데 결과적으로 전량 입고되면 자동 완료
        const finalize = newStatus === 'received' || allFull;
        if (finalize) {
          const balanceAlreadyPaid = !!current.balance_paid_at || (current.balance_amount ?? 0) <= 0;
          updates.status = balanceAlreadyPaid ? 'balance_paid' : 'received';
          updates.received_date = current.received_date || body.received_date || now.slice(0, 10);
          // 누적이 주문과 다르면 total/balance 재계산 (받은 수량 기준)
          if (anyAdjusted) {
            const rate = current.exchange_rate || 1;
            const vatType = current.vat_type || (current.is_vat_included ? 'included' : 'none');
            const krwTotal = Math.round(foreignTotalCum * rate);
            let newTotal = krwTotal;
            if (vatType === 'separate') {
              const vatAmt = Math.round(krwTotal * 0.1);
              updates.supply_amount = krwTotal; updates.vat_amount = vatAmt; newTotal = krwTotal + vatAmt;
            } else if (vatType === 'none') {
              updates.supply_amount = krwTotal; updates.vat_amount = 0; newTotal = krwTotal;
            } else {
              const supplyAmt = Math.round(krwTotal / 1.1);
              updates.supply_amount = supplyAmt; updates.vat_amount = krwTotal - supplyAmt; newTotal = krwTotal;
            }
            updates.total_amount = newTotal;
            updates.balance_amount = current.balance_paid_at ? 0 : Math.max(0, newTotal - (current.deposit_amount || 0));
          }
        } else {
          // 분할 입고 — 발주 열어둠. 금액/잔금은 확정하지 않음(완료 시 확정).
          updates.status = 'partial';
          if (!current.received_date) updates.received_date = body.received_date || now.slice(0, 10);
        }
      } else if (newStatus === 'balance_paid') {
        // 잔금 지불 기록 — 입고 전(ordered/deposit_paid)이어도 허용. 결제와 입고를 독립 처리.
        if (current.status === 'balance_paid' || current.balance_paid_at) {
          return NextResponse.json({ error: '이미 잔금 지불 처리된 발주입니다' }, { status: 409 });
        }
        if (current.status === 'cancelled' || current.status === 'draft') {
          return NextResponse.json({ error: `${current.status} 상태에서는 잔금 처리할 수 없습니다` }, { status: 400 });
        }
        updates.balance_paid_at = now;
        updates.balance_amount = 0;
        // 입고 완료(received) 후면 → 잔금완료(terminal) 로 전이. 입고 전(ordered/deposit_paid)이면 → status 유지 (입고 확인은 나중에 별도로).
        if (current.status === 'received') {
          updates.status = 'balance_paid';
        }
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

    // 부가세 유형 변경 → total_amount 재계산
    if (body.vat_type !== undefined && body.vat_type !== current.vat_type) {
      updates.vat_type = body.vat_type;
      updates.is_vat_included = body.vat_type === 'included';
      // 품목 합계 재조회
      const { data: poItems } = await db.from('purchase_order_items').select('quantity, unit_price').eq('po_id', id);
      const foreignTotal = (poItems || []).reduce((s: number, i: { quantity: number; unit_price: number }) => s + i.quantity * i.unit_price, 0);
      const rate = updates.exchange_rate || current.exchange_rate || 1;
      const krwTotal = Math.round(foreignTotal * rate);
      let newTotal = krwTotal;
      if (body.vat_type === 'separate') {
        const vatAmt = Math.round(krwTotal * 0.1);
        updates.supply_amount = krwTotal;
        updates.vat_amount = vatAmt;
        newTotal = krwTotal + vatAmt;
      } else if (body.vat_type === 'none') {
        updates.supply_amount = krwTotal;
        updates.vat_amount = 0;
        newTotal = krwTotal;
      } else {
        const supplyAmt = Math.round(krwTotal / 1.1);
        updates.supply_amount = supplyAmt;
        updates.vat_amount = krwTotal - supplyAmt;
        newTotal = krwTotal;
      }
      updates.total_amount = newTotal;
      updates.balance_amount = newTotal - (current.deposit_amount || 0);
    }

    // 품목 수정 — 입고 전(draft/ordered/deposit_paid)까지 허용 (공장 제작 수량이 발주와 다를 때 대응). 입고·잔금지불·취소 이후엔 차단.
    if (body.items && Array.isArray(body.items)) {
      if (current.status === 'received' || current.status === 'balance_paid' || current.status === 'cancelled' || current.balance_paid_at) {
        return NextResponse.json({ error: '입고·잔금 지불·취소 이후에는 품목을 수정할 수 없습니다' }, { status: 400 });
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

      // 합계 재계산 (외화 → KRW 환산). 잔금 = 새 총액 − 선납 (선납이 더 크면 0 — 과지급은 화면에서 안내).
      const rate = updates.exchange_rate || current.exchange_rate || 1;
      const foreignTotal = newItems.reduce((s: number, i: { quantity: number; unit_price: number }) => s + i.quantity * i.unit_price, 0);
      const krwTotal = Math.round(foreignTotal * rate);
      updates.total_amount = krwTotal;
      updates.balance_amount = Math.max(0, krwTotal - (current.deposit_amount || 0));
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
