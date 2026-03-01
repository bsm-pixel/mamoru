import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { saveSale } from '@/lib/ecount/sales';
import { saveCustomer } from '@/lib/ecount/customer';
import { isStatusOk } from '@/lib/ecount/client';

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
      };
      items: Array<{
        product_id?: string;
        product_name: string;
        sku?: string;
        quantity: number;
        unit_price: number;
        total_price: number;
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
        created_by: user.id,
      })
      .select()
      .single();

    if (saleError) throw saleError;

    // 판매 항목 생성
    if (items.length > 0) {
      const saleItems = items.map((item) => ({
        sale_id: created.id,
        product_id: item.product_id || null,
        product_name: item.product_name,
        sku: item.sku || null,
        quantity: item.quantity,
        unit_price: item.unit_price,
        total_price: item.total_price,
      }));

      const { error: itemsError } = await db
        .from('offline_sale_items')
        .insert(saleItems);

      if (itemsError) throw itemsError;
    }

    // === 이카운트 자동 동기화 ===
    let ecountSyncStatus = 'pending';
    try {
      // 1) 고객 이카운트 거래처 코드 확인/자동등록
      let customerCode = '';
      if (sale.customer_id) {
        const { data: cust } = await db
          .from('customers')
          .select('name, phone, email, address_road, address_detail, ecount_customer_code')
          .eq('id', sale.customer_id)
          .single();

        if (cust?.ecount_customer_code) {
          customerCode = cust.ecount_customer_code;
        } else if (cust) {
          // 거래처 코드 없는 기존 고객 → 자동 채번 + 이카운트 등록
          try {
            const { data: maxRow } = await db
              .from('customers')
              .select('ecount_customer_code')
              .not('ecount_customer_code', 'is', null)
              .like('ecount_customer_code', 'MM-%')
              .order('ecount_customer_code', { ascending: false })
              .limit(1);

            let nextNum = 1;
            if (maxRow && maxRow.length > 0 && maxRow[0].ecount_customer_code) {
              const match = maxRow[0].ecount_customer_code.match(/MM-(\d+)/);
              if (match) nextNum = parseInt(match[1]) + 1;
            }
            const padLen = Math.max(3, String(nextNum).length);
            const newCode = `MM-${String(nextNum).padStart(padLen, '0')}`;

            const custResult = await saveCustomer({
              custCode: newCode,
              custName: cust.name,
              phone: cust.phone || undefined,
              email: cust.email || undefined,
              address: [cust.address_road, cust.address_detail].filter(Boolean).join(' ') || undefined,
              remarks: 'TMS 자동등록',
            });

            if (isStatusOk(custResult.Status)) {
              customerCode = newCode;
              await db
                .from('customers')
                .update({ ecount_customer_code: newCode })
                .eq('id', sale.customer_id);
            }
          } catch (custErr) {
            console.error('이카운트 거래처 자동등록 실패:', custErr);
          }
        }
      }

      // 2) 판매전표 생성
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const ecountItems = items.map((item: any) => ({
        PROD_CD: item.sku || '',
        PROD_DES: item.product_name,
        QTY: String(item.quantity),
        PRICE: String(item.unit_price),
        SUPPLY_AMT: String(item.total_price),
      }));

      const result = await saveSale({
        saleDate,
        customerCode,
        items: ecountItems,
        remarks: `TMS ${saleNumber}`,
      });

      // 3) 동기화 상태 업데이트
      if (isStatusOk(result.Status)) {
        ecountSyncStatus = 'synced';
        await db
          .from('offline_sales')
          .update({
            ecount_sync_status: 'synced',
            ecount_synced_at: new Date().toISOString(),
          })
          .eq('id', created.id);
      } else {
        ecountSyncStatus = 'failed';
        await db
          .from('offline_sales')
          .update({ ecount_sync_status: 'failed' })
          .eq('id', created.id);
      }
    } catch (ecountErr) {
      // 이카운트 실패해도 판매 저장은 성공 유지
      ecountSyncStatus = 'failed';
      console.error('이카운트 자동 동기화 실패:', ecountErr);
      await db
        .from('offline_sales')
        .update({ ecount_sync_status: 'failed' })
        .eq('id', created.id)
        .catch(() => {}); // update 실패도 무시
    }

    return NextResponse.json({ sale: created, saleNumber, ecountSyncStatus });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
