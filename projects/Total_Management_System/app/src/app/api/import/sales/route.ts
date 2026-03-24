import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * POST /api/import/sales — 이카운트 판매내역 CSV 업로드
 *
 * 예상 컬럼: 일자, 품목명, 수량, 공급가액, 부가세, 합계, 거래처명, 거래유형, 시리얼
 * → offline_sales + offline_sale_items 생성
 * → 시리얼 있으면 product_serials 생성 + 연결
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const rows: Array<Record<string, string>> = body.rows;

    if (!rows || rows.length === 0) {
      return NextResponse.json({ error: '데이터가 없습니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    let created = 0;
    let serialsCreated = 0;
    const errors: string[] = [];

    // 일자+거래처명 기준으로 그룹핑 (같은 날 같은 고객 = 1건 판매)
    const salesMap = new Map<string, Array<Record<string, string>>>();
    for (const row of rows) {
      const rawDate = (row['일자'] || '').trim();
      const date = rawDate.replace(/^(\d{4}-\d{2}-\d{2}).*$/, '$1'); // 2025-06-15-1 → 2025-06-15
      const customer = (row['거래처명'] || '').trim();
      if (!date || !customer) continue;
      const key = `${date}__${customer}`;
      if (!salesMap.has(key)) salesMap.set(key, []);
      salesMap.get(key)!.push(row);
    }

    for (const [key, items] of salesMap) {
      try {
        const [saleDate, customerName] = key.split('__');

        // 고객 찾기 (이름 기준)
        const { data: customers } = await db
          .from('customers')
          .select('id')
          .eq('name', customerName)
          .limit(1);
        const customerId = customers?.[0]?.id || null;

        // 합계 계산
        const totalAmount = items.reduce((s, r) => s + (parseInt(r['합계'] || '0') || 0), 0);
        const supplyAmount = items.reduce((s, r) => s + (parseInt(r['공급가액'] || '0') || 0), 0);
        const vatAmount = items.reduce((s, r) => s + (parseInt(r['부가세'] || '0') || 0), 0);

        // sale_number 자동 채번
        const dateStr = saleDate.replace(/-/g, '').slice(0, 8);
        const { count } = await db
          .from('offline_sales')
          .select('*', { count: 'exact', head: true })
          .like('sale_number', `OS-${dateStr}%`);
        const saleNumber = `OS-${dateStr}-${String((count || 0) + 1).padStart(3, '0')}`;

        // 판매 레코드 생성
        const { data: sale, error: saleErr } = await db
          .from('offline_sales')
          .insert({
            sale_number: saleNumber,
            customer_id: customerId,
            customer_name: customerName,
            sale_date: saleDate,
            total_amount: totalAmount,
            supply_amount: supplyAmount,
            vat_amount: vatAmount,
            paid_amount: totalAmount,
            payment_method: vatAmount > 0 ? 'card' : 'cash',
            payment_status: 'paid',
            is_vat_included: vatAmount > 0,
            memo: '이카운트 이관',
            created_by: user.id,
          })
          .select()
          .single();

        if (saleErr) throw saleErr;

        // 판매 항목 생성
        const saleItems = items.map((r) => ({
          sale_id: sale.id,
          product_name: r['품목명'] || '',
          quantity: parseInt(r['수량'] || '1') || 1,
          unit_price: parseInt(r['공급가액'] || '0') || 0,
          total_price: parseInt(r['합계'] || '0') || 0,
          supply_amount: parseInt(r['공급가액'] || '0') || 0,
          vat_amount: parseInt(r['부가세'] || '0') || 0,
        }));

        await db.from('offline_sale_items').insert(saleItems);

        // 시리얼 처리
        for (const r of items) {
          const serial = (r['시리얼'] || r['시리얼번호'] || '').trim();
          if (!serial) continue;

          // 시리얼 중복 체크
          const { data: existingSerial } = await db
            .from('product_serials')
            .select('id')
            .eq('serial_number', serial)
            .limit(1);

          if (!existingSerial || existingSerial.length === 0) {
            await db.from('product_serials').insert({
              serial_number: serial,
              status: 'sold',
              sold_via: 'offline',
              offline_sale_id: sale.id,
              sold_at: saleDate,
              sold_to_name: customerName,
              warehouse_zone: 'storage',
            });
            serialsCreated++;
          }
        }

        created++;
      } catch (err) {
        errors.push(`${key}: ${String(err)}`);
      }
    }

    return NextResponse.json({
      sales_created: created,
      serials_created: serialsCreated,
      errors,
      total_rows: rows.length,
      total_sales: salesMap.size,
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
