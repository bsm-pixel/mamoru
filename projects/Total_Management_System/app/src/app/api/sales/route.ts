import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { updateImwebStock } from '@/lib/imweb/client';
import { recalcOutstanding } from '@/lib/outstanding';

/** GET /api/sales — 오프라인 판매 목록 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const url = req.nextUrl;
    const search = url.searchParams.get('search');
    const page = parseInt(url.searchParams.get('page') || '1');
    const limit = parseInt(url.searchParams.get('limit') || '20');
    const from = (page - 1) * limit;
    const to = from + limit - 1;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let query = (supabase as any)
      .from('offline_sales')
      .select('*', { count: 'exact' })
      .order('sale_date', { ascending: false })
      .range(from, to);

    if (search) {
      query = query.or(
        `customer_name.ilike.%${search}%,customer_phone.ilike.%${search}%,sale_number.ilike.%${search}%`
      );
    }

    const { data, count, error } = await query;
    if (error) throw error;

    return NextResponse.json({ sales: data || [], total: count || 0 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/sales — 오프라인 판매 생성 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const { sale, items } = body as {
      sale: {
        customer_id?: string;
        customer_name: string;
        customer_phone?: string;
        sale_date?: string;
        total_amount: number;
        discount_amount?: number;
        paid_amount: number;
        payment_method: string;
        payment_status?: string;
        payment_detail?: Record<string, number>;
        customer_type?: string;
        memo?: string;
        supply_amount?: number;
        vat_amount?: number;
        is_vat_included?: boolean;
        sale_channel?: string;
        source_consultation_id?: string;     // 070: 상담 → 판매 link
      };
      items: Array<{
        product_id?: string;
        product_name: string;
        sku?: string;
        category?: string;
        quantity: number;
        unit_price: number;
        total_price: number;
        supply_amount?: number;
        vat_amount?: number;
        serial_ids?: string[];
        manual_serials?: string[];
      }>;
    };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 무결성 가드 — 같은 판매 안에서 동일 시리얼이 두 품목에 중복 배정되면 거부
    //   (한 물리 시리얼 = 한 품목. 자동생성/수동입력 중복 시 한쪽이 덮어써지는 손상 방지) 2026-06-11
    {
      const allManual: string[] = [];
      const allSerialIds: string[] = [];
      for (const it of items) {
        for (const s of (it.manual_serials || [])) { const v = String(s).trim(); if (v) allManual.push(v); }
        for (const sid of (it.serial_ids || [])) { if (sid) allSerialIds.push(sid); }
      }
      const dupManual = allManual.find((v, i) => allManual.indexOf(v) !== i);
      const dupId = allSerialIds.find((v, i) => allSerialIds.indexOf(v) !== i);
      if (dupManual || dupId) {
        return NextResponse.json({
          error: `같은 판매에 동일 시리얼이 중복 배정되었습니다${dupManual ? `: "${dupManual}"` : ''}. 품목마다 다른 시리얼을 지정해주세요.`,
          code: 'SERIAL_DUPLICATE_IN_PAYLOAD',
        }, { status: 400 });
      }
    }

    // 판매번호 생성: OS-YYYYMMDD-NNN
    const saleDate = sale.sale_date || new Date().toISOString().slice(0, 10);
    const today = saleDate.replace(/-/g, '');
    const { count } = await db
      .from('offline_sales')
      .select('*', { count: 'exact', head: true })
      .gte('sale_date', saleDate);
    const seq = String((count || 0) + 1).padStart(3, '0');
    const saleNumber = `OS-${today}-${seq}`;

    // VAT 자동 계산 — 카드 금액 기준
    let supplyAmount = sale.supply_amount || 0;
    let vatAmount = sale.vat_amount || 0;
    const cardAmount = sale.payment_method === 'mixed'
      ? (sale.payment_detail?.card || 0)
      : sale.payment_method === 'card' ? sale.paid_amount : 0;
    const isVatIncluded = sale.is_vat_included ?? (cardAmount > 0);
    if (isVatIncluded && !supplyAmount && cardAmount > 0) {
      supplyAmount = Math.round(cardAmount / 1.1);
      vatAmount = cardAmount - supplyAmount;
    }

    // 판매 레코드 생성
    const { data: created, error: saleError } = await db
      .from('offline_sales')
      .insert({
        sale_number: saleNumber,
        customer_id: sale.customer_id || null,
        customer_name: sale.customer_name,
        customer_phone: sale.customer_phone || null,
        sale_date: saleDate,
        total_amount: sale.total_amount,
        discount_amount: sale.discount_amount || 0,
        paid_amount: sale.paid_amount,
        payment_method: sale.payment_method,
        payment_status: sale.payment_status || 'paid',
        payment_detail: sale.payment_detail || null,
        memo: sale.memo || null,
        supply_amount: supplyAmount,
        vat_amount: vatAmount,
        is_vat_included: isVatIncluded,
        sale_channel: sale.sale_channel || 'offline',
        customer_type: sale.customer_type || null,
        contract_id: (sale as Record<string, unknown>).contract_id || null,
        source_consultation_id: sale.source_consultation_id || null,
        created_by: user.id,
      })
      .select()
      .single();

    if (saleError) throw saleError;

    // 항목별 id 매핑 (시리얼 → sale_item 연결용)
    const itemIdMap: Record<number, string> = {};

    // 시리얼 수량 서버 검증: 시리얼 지정 시 수량과 일치해야 함
    for (const item of items) {
      if (item.serial_ids && item.serial_ids.length > 0 && item.serial_ids.length !== item.quantity) {
        return NextResponse.json(
          { error: `${item.product_name}: 시리얼 ${item.serial_ids.length}개 ≠ 수량 ${item.quantity}` },
          { status: 400 }
        );
      }
    }

    // 판매 항목 생성
    if (items.length > 0) {
      const saleItems = items.map((item) => {
        let itemSupply = item.supply_amount || 0;
        let itemVat = item.vat_amount || 0;
        if (isVatIncluded && !itemSupply) {
          itemSupply = Math.round(item.total_price / 1.1);
          itemVat = item.total_price - itemSupply;
        }
        return {
          sale_id: created.id,
          product_id: item.product_id || null,
          product_name: item.product_name,
          sku: item.sku || null,
          category: item.category || null,
          quantity: item.quantity,
          unit_price: item.unit_price,
          total_price: item.total_price,
          supply_amount: itemSupply,
          vat_amount: itemVat,
        };
      });

      const { data: insertedItems, error: itemsError } = await db
        .from('offline_sale_items')
        .insert(saleItems)
        .select('id, product_id, product_name');

      if (itemsError) throw itemsError;

      // 항목별 id 매핑 (시리얼 연결용) — 블록 밖에서도 접근 가능하도록
      for (let i = 0; i < (insertedItems || []).length; i++) {
        itemIdMap[i] = insertedItems[i].id;
      }
    }

    // 시리얼 연결: previous_zone 저장 후 status → sold (낙관적 잠금) + sale_item_id 매핑
    for (let itemIdx = 0; itemIdx < items.length; itemIdx++) {
      const item = items[itemIdx];
      const saleItemId = itemIdMap[itemIdx] || null;
      const serialIds = item.serial_ids || [];
      if (serialIds.length === 0) continue;

      const { data: currentSerials } = await db
        .from('product_serials')
        .select('id, warehouse_zone')
        .in('id', serialIds)
        .eq('status', 'in_stock');

      if (!currentSerials || currentSerials.length !== serialIds.length) {
        return NextResponse.json(
          { error: `시리얼 ${serialIds.length}개 중 ${currentSerials?.length || 0}개만 판매 가능 (이미 판매/반품된 시리얼 확인)` },
          { status: 409 }
        );
      }

      for (const serial of currentSerials) {
        const { count } = await db
          .from('product_serials')
          .update({
            previous_zone: serial.warehouse_zone,
            status: 'sold',
            sold_via: 'offline',
            offline_sale_id: created.id,
            sale_item_id: saleItemId, // 준비표 매칭용
            sold_at: new Date().toISOString(),
            sold_to_name: sale.customer_name,
            sold_to_phone: sale.customer_phone || null,
          })
          .eq('id', serial.id)
          .eq('status', 'in_stock')
          .select('id', { count: 'exact', head: true });

        if (count === 0) {
          return NextResponse.json(
            { error: `시리얼 ${serial.id}이(가) 다른 판매에 이미 할당됨` },
            { status: 409 }
          );
        }
      }
    }

    // 직접 입력 시리얼: product_serials에 자동 등록 + 즉시 sold + sale_item_id 연결
    // Phase A 서버 안전망 — 다른 판매에 sold 시리얼은 명시 동의(allow_serial_transfer) 없이 강탈 금지 (2026-05-18)
    const allowTransfer = (body as Record<string, unknown>).allow_serial_transfer === true;
    for (let itemIdx = 0; itemIdx < items.length; itemIdx++) {
      const item = items[itemIdx];
      const saleItemId = itemIdMap[itemIdx] || null;
      const manualSerials = (item as Record<string, unknown>).manual_serials as string[] | undefined;
      if (manualSerials && manualSerials.length > 0) {
        for (const serialNumber of manualSerials) {
          if (!serialNumber.trim()) continue;
          // 중복 체크
          const { data: existing } = await db
            .from('product_serials')
            .select('id, offline_sale_id, status')
            .eq('serial_number', serialNumber.trim())
            .limit(1);

          if (existing && existing.length > 0) {
            const sr = existing[0] as { id: string; offline_sale_id: string | null; status: string };
            // 다른 판매에 이미 매핑되어 있으면 강탈 금지 (명시 동의 없이)
            if (sr.offline_sale_id && sr.offline_sale_id !== created.id && !allowTransfer) {
              return NextResponse.json({
                error: `시리얼 "${serialNumber.trim()}" 은 이미 다른 판매에 등록되어 있습니다. 명시적 이전 동의가 필요합니다.`,
                code: 'SERIAL_DUPLICATE_TRANSFER_BLOCKED',
                conflicting_serial: serialNumber.trim(),
                conflicting_sale_id: sr.offline_sale_id,
              }, { status: 409 });
            }
            // 이미 있으면 sold로 전환 + product_id + sale_item_id 업데이트
            await db.from('product_serials').update({
              product_id: item.product_id || null,
              sale_item_id: saleItemId,
              status: 'sold',
              sold_via: 'offline',
              offline_sale_id: created.id,
              sold_at: new Date().toISOString(),
              sold_to_name: sale.customer_name,
              sold_to_phone: sale.customer_phone || null,
              previous_zone: 'ready',
            }).eq('id', sr.id);
          } else {
            // 없으면 새로 생성 + 즉시 sold + sale_item_id 연결
            await db.from('product_serials').insert({
              serial_number: serialNumber.trim(),
              product_id: item.product_id || null,
              sale_item_id: saleItemId,
              status: 'sold',
              warehouse_zone: 'ready',
              previous_zone: 'ready',
              sold_via: 'offline',
              offline_sale_id: created.id,
              sold_at: new Date().toISOString(),
              sold_to_name: sale.customer_name,
              sold_to_phone: sale.customer_phone || null,
            });
          }
        }
      }
    }

    // 재고 차감 + 아임웹 동기화 — 상품별 병렬 처리
    const productQtyMap: Record<string, number> = {};
    for (const item of items) {
      if (item.product_id && item.quantity > 0) {
        const serialQty = item.serial_ids?.length || 0;
        const qty = serialQty > 0 ? serialQty : item.quantity;
        productQtyMap[item.product_id] = (productQtyMap[item.product_id] || 0) + qty;
      }
    }

    await Promise.all(Object.entries(productQtyMap).map(async ([productId, qty]) => {
      const { data: prod } = await db
        .from('products')
        .select('stock_quantity, raw_stock, imweb_product_no')
        .eq('id', productId)
        .single();
      if (!prod) return;

      // 해당 제품의 시리얼 사용 여부 확인
      const itemsForProduct = items.filter((i) => i.product_id === productId);
      const hasSerials = itemsForProduct.some((i) => i.serial_ids && i.serial_ids.length > 0);

      const newStock = Math.max(0, (prod.stock_quantity || 0) - qty);
      const updateData: Record<string, unknown> = { stock_quantity: newStock };

      // 시리얼 없는 판매(B2B) → 보관창고(raw_stock)에서 차감
      if (!hasSerials) {
        updateData.raw_stock = Math.max(0, (prod.raw_stock || 0) - qty);
      }

      await db.from('products').update(updateData).eq('id', productId);

      // 아임웹 재고 동기화 — delta(증감값) 전달
      if (prod.imweb_product_no) {
        try {
          await updateImwebStock(Number(prod.imweb_product_no), -qty);
        } catch (e) {
          console.error('[imweb] 판매 재고 동기화 실패:', prod.imweb_product_no, e);
        }
      }
    }));

    const paymentStatus = sale.payment_status || 'paid';

    // 미수금 자동 반영: 멱등 재계산 (±diff 누적 X — drift 방지)
    if (sale.customer_id) {
      await recalcOutstanding(db, sale.customer_id);
    }

    // 077: 회계 자동 연동 — 입금된 금액을 cash_transactions(income)에 자동 기록
    // (paid 또는 partial이고 paid_amount > 0인 경우만 — unpaid는 입금 발생 X)
    if (sale.paid_amount > 0 && paymentStatus !== 'unpaid') {
      try {
        await db.from('cash_transactions').insert({
          transaction_date: saleDate,
          type: 'income',
          category: '매출입금',
          amount: sale.paid_amount,
          memo: `${saleNumber} ${sale.customer_name}`,
          source_type: 'offline_sale',
          source_id: created.id,
          created_by: user.id,
        });
      } catch (e) {
        // cashflow INSERT 실패해도 판매는 유지 (회계 페이지에서 수동 보정 가능)
        console.error('[sales POST] cashflow auto-insert 실패:', e);
      }
    }

    return NextResponse.json({ sale: created, saleNumber });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : JSON.stringify(err);
    console.error('[sales POST] error:', message);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
