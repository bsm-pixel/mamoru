import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * POST /api/import/serials — 이카운트 CSV에서 시리얼만 재임포트
 *
 * CSV 컬럼: 일자, 품목명, 수량, 공급가액, 부가세, 합계, 거래처명, 시리얼
 * → 기존 판매 건(일자+거래처명)을 찾아서 시리얼을 product_serials에 연결
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
    let linked = 0;
    let skipped = 0;
    let noSale = 0;
    let duplicate = 0;
    const errors: string[] = [];

    for (const row of rows) {
      try {
        // 시리얼 값 추출 (컬럼명 여러 변형 대응)
        const serial = (
          row['시리얼'] || row['시리얼번호'] || row['일련번호'] || row['S/N'] || row['serial'] || ''
        ).toString().trim();

        if (!serial) { skipped++; continue; }

        // 일자, 거래처명, 품목명 추출
        const rawDate = (row['일자'] || '').trim();
        const date = rawDate.replace(/^(\d{4}\/\d{2}\/\d{2}).*$/, '$1').replace(/\//g, '-'); // 2024/02/24-2 → 2024-02-24
        const customerName = (row['거래처명'] || '').trim();
        const productName = (row['품목명'] || '').trim();

        if (!date || !customerName || !productName) { skipped++; continue; }

        // 시리얼 중복 체크
        const { data: existingSerial } = await db
          .from('product_serials')
          .select('id')
          .eq('serial_number', String(serial))
          .limit(1);

        if (existingSerial && existingSerial.length > 0) { duplicate++; continue; }

        // 기존 판매 건 찾기 (일자+거래처명, 이카운트 이관 메모)
        const { data: sales } = await db
          .from('offline_sales')
          .select('id')
          .eq('sale_date', date)
          .eq('customer_name', customerName)
          .eq('memo', '이카운트 이관')
          .limit(1);

        const saleId = sales?.[0]?.id || null;
        if (!saleId) {
          noSale++;
          errors.push(`매칭실패: ${date} / ${customerName} / ${productName} / 시리얼:${serial}`);
          continue;
        }

        // 제품 매칭 (품목명 기준)
        const { data: products } = await db
          .from('products')
          .select('id')
          .eq('name', productName)
          .limit(1);

        const productId = products?.[0]?.id || null;

        // 시리얼 생성 + 판매 건 연결
        const { error: insertErr } = await db.from('product_serials').insert({
          product_id: productId,
          serial_number: String(serial),
          barcode: null,
          status: 'sold',
          sold_via: 'offline',
          offline_sale_id: saleId,
          sold_at: date,
          sold_to_name: customerName,
          warehouse_zone: 'raw',
        });

        if (insertErr) throw insertErr;
        linked++;
      } catch (err) {
        errors.push(`${row['시리얼'] || '?'}: ${String(err)}`);
      }
    }

    return NextResponse.json({
      linked,       // 연결 성공
      skipped,      // 시리얼 없는 행
      noSale,       // 판매 건 매칭 실패
      duplicate,    // 이미 존재하는 시리얼
      errors,
      total_rows: rows.length,
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
