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

      // 1. 시리얼 복원 — previous_zone으로 원래 위치 복원
      const { data: serials } = await db
        .from('product_serials')
        .select('id, product_id, previous_zone')
        .eq('offline_sale_id', id);

      if (serials && serials.length > 0) {
        for (const serial of serials) {
          await db
            .from('product_serials')
            .update({
              status: 'in_stock',
              warehouse_zone: serial.previous_zone || 'ready', // 원래 zone 복원 (NULL이면 ready fallback)
              previous_zone: null,
              sold_via: null,
              offline_sale_id: null,
              sold_at: null,
              sold_to_name: null,
              sold_to_phone: null,
            })
            .eq('id', serial.id);
        }
      }

      // 2. 재고 복원 + 3. 아임웹 재고 동기화
      const { data: items } = await db
        .from('offline_sale_items')
        .select('product_id, quantity')
        .eq('sale_id', id);

      if (items && items.length > 0) {
        // 상품별 수량 집계 후 병렬 처리
        const productQtyMap: Record<string, number> = {};
        for (const item of items) {
          if (item.product_id && item.quantity) {
            productQtyMap[item.product_id] = (productQtyMap[item.product_id] || 0) + item.quantity;
          }
        }

        await Promise.all(Object.entries(productQtyMap).map(async ([productId, qty]) => {
          const { data: prod } = await db
            .from('products')
            .select('stock_quantity, raw_stock, imweb_product_no')
            .eq('id', productId)
            .single();
          if (!prod) return;

          // 이 상품에 시리얼이 연결되어 있었는지 확인 (B2B는 시리얼 없이 raw_stock 차감)
          const { count: serialCount } = await db
            .from('product_serials')
            .select('id', { count: 'exact', head: true })
            .eq('offline_sale_id', id)
            .eq('product_id', productId);

          const newStock = (prod.stock_quantity || 0) + qty;
          const updateData: Record<string, unknown> = { stock_quantity: newStock };
          // B2B 판매(시리얼 없음)였으면 raw_stock도 복원
          if (!serialCount || serialCount === 0) {
            updateData.raw_stock = (prod.raw_stock || 0) + qty;
          }
          await db.from('products').update(updateData).eq('id', productId);

          if (prod.imweb_product_no) {
            try {
              // 디스플레이 제외 판매 가능 재고만 아임웹에 전달
              const { count: displayCount } = await db
                .from('product_serials')
                .select('id', { count: 'exact', head: true })
                .eq('product_id', productId)
                .eq('status', 'in_stock')
                .eq('warehouse_zone', 'display');
              const sellableStock = Math.max(0, newStock - (displayCount || 0));
              await updateImwebStock(Number(prod.imweb_product_no), sellableStock);
            } catch (e) {
              console.error('[imweb] 취소 재고 동기화 실패:', prod.imweb_product_no, e);
            }
          }
        }));
      }

      // 4. 미수금 차감 (미결제/부분결제였던 경우) — 할인 반영
      if (sale.customer_id && sale.payment_status !== 'paid') {
        const effectiveTotal = sale.total_amount - (sale.discount_amount || 0);
        const unpaidAmount = effectiveTotal - (sale.paid_amount || 0);
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

    // --- A-2) 반품 처리 ---
    if (action === 'return') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '취소된 판매는 반품 처리할 수 없습니다' }, { status: 400 });
      }
      if (sale.returned_at) {
        return NextResponse.json({ error: '이미 반품 처리된 판매입니다' }, { status: 400 });
      }

      const reason = body.reason || '';

      // 1. 시리얼 복원 — previous_zone으로 원래 위치 복원
      const { data: serials } = await db
        .from('product_serials')
        .select('id, product_id, previous_zone')
        .eq('offline_sale_id', id);

      if (serials && serials.length > 0) {
        for (const serial of serials) {
          await db
            .from('product_serials')
            .update({
              status: 'in_stock',
              warehouse_zone: serial.previous_zone || 'ready',
              previous_zone: null,
              sold_via: null,
              offline_sale_id: null,
              sold_at: null,
              sold_to_name: null,
              sold_to_phone: null,
            })
            .eq('id', serial.id);
        }
      }

      // 2. 재고 복원 + 3. 아임웹 재고 동기화
      const { data: items } = await db
        .from('offline_sale_items')
        .select('product_id, quantity')
        .eq('sale_id', id);

      if (items && items.length > 0) {
        const productQtyMap: Record<string, number> = {};
        for (const item of items) {
          if (item.product_id && item.quantity) {
            productQtyMap[item.product_id] = (productQtyMap[item.product_id] || 0) + item.quantity;
          }
        }

        await Promise.all(Object.entries(productQtyMap).map(async ([productId, qty]) => {
          const { data: prod } = await db
            .from('products')
            .select('stock_quantity, raw_stock, imweb_product_no')
            .eq('id', productId)
            .single();
          if (!prod) return;

          const { count: serialCount } = await db
            .from('product_serials')
            .select('id', { count: 'exact', head: true })
            .eq('offline_sale_id', id)
            .eq('product_id', productId);

          const newStock = (prod.stock_quantity || 0) + qty;
          const updateData: Record<string, unknown> = { stock_quantity: newStock };
          if (!serialCount || serialCount === 0) {
            updateData.raw_stock = (prod.raw_stock || 0) + qty;
          }
          await db.from('products').update(updateData).eq('id', productId);

          if (prod.imweb_product_no) {
            try {
              const { count: displayCount } = await db
                .from('product_serials')
                .select('id', { count: 'exact', head: true })
                .eq('product_id', productId)
                .eq('status', 'in_stock')
                .eq('warehouse_zone', 'display');
              const sellableStock = Math.max(0, newStock - (displayCount || 0));
              await updateImwebStock(Number(prod.imweb_product_no), sellableStock);
            } catch (e) {
              console.error('[imweb] 반품 재고 동기화 실패:', prod.imweb_product_no, e);
            }
          }
        }));
      }

      // 4. 미수금 차감
      if (sale.customer_id && sale.payment_status !== 'paid') {
        const effectiveTotal = sale.total_amount - (sale.discount_amount || 0);
        const unpaidAmount = effectiveTotal - (sale.paid_amount || 0);
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

      // 5. 반품 상태 기록
      const { error: updateErr } = await db
        .from('offline_sales')
        .update({
          returned_at: new Date().toISOString(),
          return_reason: reason,
        })
        .eq('id', id);

      if (updateErr) throw updateErr;
      return NextResponse.json({ success: true, action: 'returned' });
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

    // --- D) 판매 정보 수정 (금액/할인/결제방법/날짜) ---
    if (action === 'edit_sale') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '취소된 판매는 수정할 수 없습니다' }, { status: 400 });
      }

      const { total_amount, discount_amount, payment_method, sale_date, payment_detail } = body as {
        total_amount?: number;
        discount_amount?: number;
        payment_method?: string;
        sale_date?: string;
        payment_detail?: Record<string, number>;
      };

      const updateData: Record<string, unknown> = {};
      if (total_amount !== undefined) updateData.total_amount = total_amount;
      if (discount_amount !== undefined) updateData.discount_amount = discount_amount;
      if (payment_method !== undefined) updateData.payment_method = payment_method;
      if (sale_date !== undefined) updateData.sale_date = sale_date;
      if (payment_detail !== undefined) updateData.payment_detail = payment_detail;

      // VAT 재계산 (카드 금액 변경 시)
      if (total_amount !== undefined || payment_method !== undefined) {
        const newTotal = total_amount ?? sale.total_amount;
        const newMethod = payment_method ?? sale.payment_method;
        const cardAmount = newMethod === 'card' ? newTotal
          : newMethod === 'mixed' && payment_detail?.card ? payment_detail.card
          : 0;
        if (cardAmount > 0) {
          updateData.supply_amount = Math.round(cardAmount / 1.1);
          updateData.vat_amount = cardAmount - (updateData.supply_amount as number);
          updateData.is_vat_included = true;
        }
      }

      if (Object.keys(updateData).length === 0) {
        return NextResponse.json({ error: '수정할 항목이 없습니다' }, { status: 400 });
      }

      const { error: updateErr } = await db
        .from('offline_sales')
        .update(updateData)
        .eq('id', id);

      if (updateErr) throw updateErr;
      return NextResponse.json({ success: true, action: 'sale_edited' });
    }

    // --- E) 판매 재구성 (제품 추가/삭제 — 내부적으로 시리얼/재고 복원 후 재적용) ---
    if (action === 'rebuild_sale') {
      if (sale.cancelled_at) {
        return NextResponse.json({ error: '취소된 판매는 수정할 수 없습니다' }, { status: 400 });
      }

      const { items: newItems, sale_info } = body as {
        items: Array<{
          product_id?: string;
          product_name: string;
          sku?: string;
          category?: string;
          quantity: number;
          unit_price: number;
          total_price: number;
          serial_ids?: string[];
          manual_serials?: string[];
        }>;
        sale_info: {
          total_amount: number;
          discount_amount?: number;
          payment_method: string;
          payment_status?: string;
          paid_amount?: number;
          sale_date?: string;
          payment_detail?: Record<string, number>;
          memo?: string;
        };
      };

      if (!newItems || newItems.length === 0) {
        return NextResponse.json({ error: '최소 1개 항목이 필요합니다' }, { status: 400 });
      }

      // ── STEP 1: 기존 시리얼 보존 여부 판단 ──
      // serial_ids나 manual_serials가 명시적으로 있을 때만 시리얼 재할당
      // 없으면 기존 시리얼 유지 (금액/항목만 수정하는 경우)
      const hasNewSerials = newItems.some((it) => (it.serial_ids && it.serial_ids.length > 0) || ((it as Record<string, unknown>).manual_serials as string[] || []).length > 0);

      const { data: oldSerials } = await db
        .from('product_serials')
        .select('id, product_id, warehouse_zone, previous_zone')
        .eq('offline_sale_id', id);

      if (hasNewSerials && oldSerials && oldSerials.length > 0) {
        // 새 시리얼이 명시적으로 있을 때만 기존 시리얼 복원
        for (const serial of oldSerials) {
          await db.from('product_serials').update({
            status: 'in_stock',
            warehouse_zone: serial.previous_zone || 'ready',
            previous_zone: null,
            sold_via: null,
            offline_sale_id: null,
            sold_at: null,
            sold_to_name: null,
            sold_to_phone: null,
            sale_item_id: null,
          }).eq('id', serial.id);
        }
      }

      // ── STEP 2: 기존 재고 복원 ──
      const { data: oldItems } = await db
        .from('offline_sale_items')
        .select('product_id, quantity')
        .eq('sale_id', id);

      if (oldItems) {
        const oldQtyMap: Record<string, number> = {};
        for (const item of oldItems) {
          if (item.product_id) oldQtyMap[item.product_id] = (oldQtyMap[item.product_id] || 0) + item.quantity;
        }
        for (const [productId, qty] of Object.entries(oldQtyMap)) {
          const { data: prod } = await db.from('products').select('stock_quantity, raw_stock').eq('id', productId).single();
          if (prod) {
            const hadSerials = oldSerials?.some((s: { product_id: string }) => s.product_id === productId);
            const updateData: Record<string, unknown> = { stock_quantity: (prod.stock_quantity || 0) + qty };
            if (!hadSerials) updateData.raw_stock = (prod.raw_stock || 0) + qty;
            await db.from('products').update(updateData).eq('id', productId);
          }
        }
      }

      // ── STEP 3: 기존 항목 삭제 → 새 항목 생성 (.select로 ID 확보) ──
      await db.from('offline_sale_items').delete().eq('sale_id', id);

      const saleItems = newItems.map((item) => ({
        sale_id: id,
        product_id: item.product_id || null,
        product_name: item.product_name,
        sku: item.sku || null,
        category: item.category || null,
        quantity: item.quantity,
        unit_price: item.unit_price,
        total_price: item.total_price,
      }));
      const { data: insertedSaleItems } = await db.from('offline_sale_items').insert(saleItems).select('id, product_id');

      // ── STEP 3.5: 기존 시리얼 보존 시 sale_item_id 재매칭 ──
      if (!hasNewSerials && oldSerials && oldSerials.length > 0 && insertedSaleItems) {
        let serialIdx = 0;
        for (const saleItem of insertedSaleItems) {
          const matchSerials = oldSerials.filter((s: { product_id: string | null }) =>
            s.product_id === saleItem.product_id
          );
          if (matchSerials.length > 0) {
            for (const sr of matchSerials) {
              await db.from('product_serials').update({ sale_item_id: saleItem.id }).eq('id', sr.id);
            }
          } else if (serialIdx < oldSerials.length) {
            await db.from('product_serials').update({ sale_item_id: saleItem.id }).eq('id', oldSerials[serialIdx].id);
            serialIdx++;
          }
        }
      }

      // ── STEP 4: 새 시리얼 할당 (아이템별 순회 + sale_item_id 매핑) ──
      if (hasNewSerials && insertedSaleItems) {
        for (let itemIdx = 0; itemIdx < newItems.length; itemIdx++) {
          const item = newItems[itemIdx];
          const saleItemId = insertedSaleItems[itemIdx]?.id || null;

          // 4-A: 기존 등록 시리얼 (serial_ids)
          const serialIds = item.serial_ids || [];
          if (serialIds.length > 0) {
            const { data: serials } = await db
              .from('product_serials')
              .select('id, warehouse_zone')
              .in('id', serialIds)
              .eq('status', 'in_stock');

            if (!serials || serials.length !== serialIds.length) {
              return NextResponse.json({ error: '일부 시리얼이 판매 불가 상태입니다' }, { status: 409 });
            }

            for (const serial of serials) {
              const { count } = await db.from('product_serials').update({
                previous_zone: serial.warehouse_zone,
                status: 'sold',
                sold_via: 'offline',
                offline_sale_id: id,
                sale_item_id: saleItemId,
                sold_at: new Date().toISOString(),
                sold_to_name: sale.customer_name,
                sold_to_phone: sale.customer_phone || null,
              }).eq('id', serial.id).eq('status', 'in_stock')
                .select('id', { count: 'exact', head: true });

              if (count === 0) {
                return NextResponse.json({ error: `시리얼 충돌 — 다시 시도해주세요` }, { status: 409 });
              }
            }
          }

          // 4-B: 직접입력/자동생성 시리얼 (manual_serials)
          const manualSerials = item.manual_serials || [];
          for (const serialNumber of manualSerials) {
            if (!serialNumber.trim()) continue;
            const { data: existing } = await db
              .from('product_serials')
              .select('id')
              .eq('serial_number', serialNumber.trim())
              .limit(1);

            if (existing && existing.length > 0) {
              await db.from('product_serials').update({
                product_id: item.product_id || null,
                sale_item_id: saleItemId,
                status: 'sold',
                sold_via: 'offline',
                offline_sale_id: id,
                sold_at: new Date().toISOString(),
                sold_to_name: sale.customer_name,
                sold_to_phone: sale.customer_phone || null,
                previous_zone: 'ready',
              }).eq('id', existing[0].id);
            } else {
              await db.from('product_serials').insert({
                serial_number: serialNumber.trim(),
                product_id: item.product_id || null,
                sale_item_id: saleItemId,
                status: 'sold',
                warehouse_zone: 'ready',
                previous_zone: 'ready',
                sold_via: 'offline',
                offline_sale_id: id,
                sold_at: new Date().toISOString(),
                sold_to_name: sale.customer_name,
                sold_to_phone: sale.customer_phone || null,
              });
            }
          }
        }
      }

      // ── STEP 5: 새 재고 차감 ──
      const newQtyMap: Record<string, number> = {};
      for (const item of newItems) {
        if (item.product_id && item.quantity > 0) {
          const qty = item.serial_ids?.length || item.quantity;
          newQtyMap[item.product_id] = (newQtyMap[item.product_id] || 0) + qty;
        }
      }
      for (const [productId, qty] of Object.entries(newQtyMap)) {
        const { data: prod } = await db.from('products').select('stock_quantity, raw_stock, imweb_product_no').eq('id', productId).single();
        if (!prod) continue;
        const hasSerials = newItems.some((it) => it.product_id === productId && ((it.serial_ids && it.serial_ids.length > 0) || (it.manual_serials && it.manual_serials.length > 0)));
        const newStock = Math.max(0, (prod.stock_quantity || 0) - qty);
        const updateData: Record<string, unknown> = { stock_quantity: newStock };
        if (!hasSerials) updateData.raw_stock = Math.max(0, (prod.raw_stock || 0) - qty);
        await db.from('products').update(updateData).eq('id', productId);

        // 아임웹 동기화
        if (prod.imweb_product_no) {
          try {
            const { count: displayCount } = await db.from('product_serials')
              .select('id', { count: 'exact', head: true })
              .eq('product_id', productId).eq('status', 'in_stock').eq('warehouse_zone', 'display');
            await updateImwebStock(Number(prod.imweb_product_no), Math.max(0, newStock - (displayCount || 0)));
          } catch { /* 실패해도 계속 */ }
        }
      }

      // ── STEP 6: 판매 레코드 업데이트 ──
      const cardAmount = sale_info.payment_method === 'card' ? sale_info.total_amount
        : sale_info.payment_method === 'mixed' && sale_info.payment_detail?.card ? sale_info.payment_detail.card
        : 0;
      const supplyAmount = cardAmount > 0 ? Math.round(cardAmount / 1.1) : 0;
      const vatAmount = cardAmount > 0 ? cardAmount - supplyAmount : 0;

      await db.from('offline_sales').update({
        total_amount: sale_info.total_amount,
        discount_amount: sale_info.discount_amount || 0,
        payment_method: sale_info.payment_method,
        payment_status: sale_info.payment_status || sale.payment_status,
        paid_amount: sale_info.paid_amount ?? sale_info.total_amount,
        payment_detail: sale_info.payment_detail || null,
        sale_date: sale_info.sale_date || sale.sale_date,
        sale_channel: (sale_info as Record<string, unknown>).sale_channel || sale.sale_channel,
        memo: sale_info.memo ?? sale.memo,
        supply_amount: supplyAmount,
        vat_amount: vatAmount,
        is_vat_included: cardAmount > 0,
      }).eq('id', id);

      return NextResponse.json({ success: true, action: 'sale_rebuilt' });
    }

    return NextResponse.json({ error: '알 수 없는 action' }, { status: 400 });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : JSON.stringify(err);
    console.error('[sales PATCH] error:', message);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
