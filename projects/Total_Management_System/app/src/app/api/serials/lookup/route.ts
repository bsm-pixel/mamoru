import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { parseSerial, formatSerial } from '@/lib/serial/format';

/** GET /api/serials/lookup?q=시리얼번호 — 시리얼 풀 정보 조회 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const q = req.nextUrl.searchParams.get('q')?.trim();
    if (!q || q.length < 2) {
      return NextResponse.json({ error: '검색어 2자 이상 필요' }, { status: 400 });
    }

    // 1. 시리얼 검색 (번호 or 바코드)
    let { data: serial } = await db
      .from('product_serials')
      .select('*')
      .eq('serial_number', q)
      .single();

    if (!serial) {
      const res = await db
        .from('product_serials')
        .select('*')
        .eq('barcode', q)
        .single();
      serial = res.data;
    }

    // M{YY}-{NNNN} 하이픈 유무 무시 — 정식형(M26-0042)으로 재시도 (고객이 M260042로 입력해도 매칭)
    if (!serial) {
      const p = parseSerial(q);
      if (p) {
        const canonical = formatSerial(p.year2, p.seq);
        if (canonical !== q) {
          const res = await db.from('product_serials').select('*').eq('serial_number', canonical).single();
          serial = res.data;
        }
      }
    }

    // 부분 매칭 (시리얼번호 포함)
    if (!serial) {
      const res = await db
        .from('product_serials')
        .select('*')
        .ilike('serial_number', `%${q}%`)
        .limit(1)
        .single();
      serial = res.data;
    }

    if (!serial) {
      return NextResponse.json({ serial: null, product: null, sale: null, repairs: [] });
    }

    // 2. 제품 정보
    let product = null;
    if (serial.product_id) {
      const { data } = await db
        .from('products')
        .select('id, name, sku, category, image_url, price')
        .eq('id', serial.product_id)
        .single();
      product = data;
    }

    // 3. 판매 정보 (sold일 때)
    let sale = null;
    if (serial.offline_sale_id) {
      const { data } = await db
        .from('offline_sales')
        .select('id, sale_number, sale_date, sale_channel, payment_method, payment_status, customer_name, customer_phone, total_amount')
        .eq('id', serial.offline_sale_id)
        .single();
      sale = data;

      // product_id NULL인 경우 → sale_item_id로 정확 매칭, 없으면 판매 항목에서 추정
      if (!product && sale) {
        let matchItem = null;

        // 1순위: sale_item_id로 정확 매칭
        if (serial.sale_item_id) {
          const { data: exactItem } = await db
            .from('offline_sale_items')
            .select('product_name, sku, unit_price')
            .eq('id', serial.sale_item_id)
            .single();
          matchItem = exactItem;
        }

        // 2순위: 판매 항목에서 추정
        if (!matchItem) {
          const { data: saleItems } = await db
            .from('offline_sale_items')
            .select('product_name, sku, unit_price')
            .eq('sale_id', serial.offline_sale_id)
            .limit(5);
          matchItem = saleItems?.find((si: { product_name: string }) =>
            si.product_name && !si.product_name.includes('복원수리')
          ) || saleItems?.[0];
        }

        if (matchItem) {
          product = {
            id: null,
            name: matchItem.product_name,
            sku: matchItem.sku || null,
            category: null,
            image_url: null,
            price: matchItem.unit_price || 0,
          };
        }
      }
    }

    // 4. 복원수리 이력 (sold_to_phone 매칭)
    let repairs: Array<{ id: string; status: string; created_at: string }> = [];
    if (serial.sold_to_phone) {
      const { data } = await db
        .from('repairs')
        .select('id, status, created_at')
        .eq('customer_phone', serial.sold_to_phone)
        .order('created_at', { ascending: false })
        .limit(5);
      repairs = data || [];
    }

    return NextResponse.json({ serial, product, sale, repairs });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : JSON.stringify(err);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
