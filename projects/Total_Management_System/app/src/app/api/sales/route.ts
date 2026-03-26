import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { updateImwebStock } from '@/lib/imweb/client';

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
        memo?: string;
        supply_amount?: number;
        vat_amount?: number;
        is_vat_included?: boolean;
        sale_channel?: string;
      };
      items: Array<{
        product_id?: string;
        product_name: string;
        sku?: string;
        quantity: number;
        unit_price: number;
        total_price: number;
        supply_amount?: number;
        vat_amount?: number;
        serial_ids?: string[];
      }>;
    };

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 판매번호 생성: OS-YYYYMMDD-NNN
    const saleDate = sale.sale_date || new Date().toISOString().slice(0, 10);
    const today = saleDate.replace(/-/g, '');
    const { count } = await db
      .from('offline_sales')
      .select('*', { count: 'exact', head: true })
      .gte('sale_date', saleDate);
    const seq = String((count || 0) + 1).padStart(3, '0');
    const saleNumber = `OS-${today}-${seq}`;

    // VAT 자동 계산 (카드결제 시)
    let supplyAmount = sale.supply_amount || 0;
    let vatAmount = sale.vat_amount || 0;
    const isVatIncluded = sale.is_vat_included ?? (sale.payment_method === 'card');
    if (isVatIncluded && !supplyAmount) {
      supplyAmount = Math.round(sale.paid_amount / 1.1);
      vatAmount = sale.paid_amount - supplyAmount;
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
        memo: sale.memo || null,
        supply_amount: supplyAmount,
        vat_amount: vatAmount,
        is_vat_included: isVatIncluded,
        sale_channel: sale.sale_channel || 'offline',
        created_by: user.id,
      })
      .select()
      .single();

    if (saleError) throw saleError;

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
          quantity: item.quantity,
          unit_price: item.unit_price,
          total_price: item.total_price,
          supply_amount: itemSupply,
          vat_amount: itemVat,
        };
      });

      const { error: itemsError } = await db
        .from('offline_sale_items')
        .insert(saleItems);

      if (itemsError) throw itemsError;
    }

    // 시리얼 연결: previous_zone 저장 후 status → sold
    const allSerialIds = items.flatMap((item) => item.serial_ids || []);
    if (allSerialIds.length > 0) {
      // 현재 zone 조회 → previous_zone에 저장
      const { data: currentSerials } = await db
        .from('product_serials')
        .select('id, warehouse_zone')
        .in('id', allSerialIds);

      if (currentSerials && currentSerials.length > 0) {
        // 각 시리얼의 현재 zone을 previous_zone에 보존
        for (const serial of currentSerials) {
          await db
            .from('product_serials')
            .update({
              previous_zone: serial.warehouse_zone,
              status: 'sold',
              sold_via: 'offline',
              offline_sale_id: created.id,
              sold_at: new Date().toISOString(),
              sold_to_name: sale.customer_name,
              sold_to_phone: sale.customer_phone || null,
            })
            .eq('id', serial.id);
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
        .select('stock_quantity, imweb_product_no')
        .eq('id', productId)
        .single();
      if (!prod) return;

      const newStock = Math.max(0, (prod.stock_quantity || 0) - qty);
      await db.from('products').update({ stock_quantity: newStock }).eq('id', productId);

      // 아임웹 재고 동기화 (실패해도 판매 완료)
      if (prod.imweb_product_no && newStock >= 0) {
        try {
          await updateImwebStock(Number(prod.imweb_product_no), newStock);
        } catch (e) {
          console.error('[imweb] 판매 재고 동기화 실패:', prod.imweb_product_no, e);
        }
      }
    }));

    // 미수금 자동 반영: 미결제/부분결제 시 고객 outstanding_balance 업데이트
    const paymentStatus = sale.payment_status || 'paid';
    if (sale.customer_id && paymentStatus !== 'paid') {
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
            .update({ outstanding_balance: (cust.outstanding_balance || 0) + unpaidAmount })
            .eq('id', sale.customer_id);
        }
      }
    }

    return NextResponse.json({ sale: created, saleNumber });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : JSON.stringify(err);
    console.error('[sales POST] error:', message);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
