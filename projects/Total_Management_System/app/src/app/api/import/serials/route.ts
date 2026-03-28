import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * POST /api/import/serials — 이카운트 CSV에서 시리얼만 재임포트
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
        // 시리얼 값 추출
        const rawSerial = (
          row['시리얼'] || row['시리얼번호'] || row['일련번호'] || row['S/N'] || row['serial'] || ''
        ).toString().trim();

        // 숫자만 추출 (소수점, 공백 등 제거)
        const serial = rawSerial.replace(/[^0-9]/g, '');
        if (!serial) { skipped++; continue; }

        // 일자 파싱 — 다양한 형식 대응
        const rawDate = (row['일자'] || '').trim();
        // 2024/02/27 -1, 2024/02/27-1, 2024-02-27-1 등 → 2024-02-27
        const dateMatch = rawDate.match(/(\d{4})[\/\-](\d{2})[\/\-](\d{2})/);
        if (!dateMatch) { skipped++; continue; }
        const date = `${dateMatch[1]}-${dateMatch[2]}-${dateMatch[3]}`;

        const customerName = (row['거래처명'] || '').trim();
        const productName = (row['품목명'] || '').trim();
        if (!customerName || !productName) { skipped++; continue; }

        // 시리얼 중복 체크
        const { data: existingSerial } = await db
          .from('product_serials')
          .select('id')
          .eq('serial_number', serial)
          .limit(1);

        if (existingSerial && existingSerial.length > 0) { duplicate++; continue; }

        // 기존 판매 건 찾기 (일자+거래처명)
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
          errors.push(`판매매칭실패: ${date} / ${customerName} / ${productName} / S/N:${serial}`);
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
          serial_number: serial,
          barcode: null,
          status: 'sold',
          sold_via: 'offline',
          offline_sale_id: saleId,
          sold_at: `${date}T00:00:00+09:00`,
          sold_to_name: customerName,
          warehouse_zone: 'raw',
        });

        if (insertErr) {
          errors.push(`INSERT실패: S/N:${serial} / ${insertErr.message || JSON.stringify(insertErr)}`);
          continue;
        }
        linked++;
      } catch (err) {
        const msg = err instanceof Error ? err.message : JSON.stringify(err);
        errors.push(`오류: S/N:${row['시리얼'] || '?'} / ${msg}`);
      }
    }

    return NextResponse.json({ linked, skipped, noSale, duplicate, errors, total_rows: rows.length });
  } catch (err) {
    const msg = err instanceof Error ? err.message : JSON.stringify(err);
    return NextResponse.json({ error: msg }, { status: 500 });
  }
}
