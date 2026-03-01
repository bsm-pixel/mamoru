import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { listCustomers, type EcountCustomer } from '@/lib/ecount/customer';
import { isStatusOk } from '@/lib/ecount/client';

/** POST /api/customers/sync-ecount — 이카운트 거래처 → TMS customers 일괄 동기화 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // 1) 이카운트 거래처 전체 조회
    const result = await listCustomers();
    if (!isStatusOk(result.Status)) {
      return NextResponse.json(
        { error: `이카운트 거래처 조회 실패: ${result.Error?.Message || JSON.stringify(result)}` },
        { status: 502 },
      );
    }

    // Data 구조: 배열 또는 { Datas: [...] }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rawData = result.Data as any;
    const ecountList: EcountCustomer[] = Array.isArray(rawData)
      ? rawData
      : Array.isArray(rawData?.Datas)
        ? rawData.Datas
        : [];

    if (ecountList.length === 0) {
      return NextResponse.json({ synced: 0, message: '이카운트 거래처가 없습니다' });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 2) TMS 기존 고객 ecount_customer_code 조회 (중복 방지)
    const { data: existingCustomers } = await db
      .from('customers')
      .select('ecount_customer_code')
      .not('ecount_customer_code', 'is', null);

    const existingCodes = new Set(
      (existingCustomers || []).map((c: { ecount_customer_code: string }) => c.ecount_customer_code)
    );

    // 3) 신규 거래처만 INSERT
    const newCustomers = ecountList
      .filter((ec) => ec.CUST_CD && !existingCodes.has(ec.CUST_CD))
      .map((ec) => ({
        name: ec.CUST_DES || ec.CUST_CD,
        phone: ec.TEL || null,
        email: ec.EMAIL || null,
        address_road: ec.ADDR || null,
        source: 'manual' as const,
        ecount_customer_code: ec.CUST_CD,
      }));

    let inserted = 0;
    if (newCustomers.length > 0) {
      // 50건씩 배치 INSERT (Supabase 제한 대비)
      for (let i = 0; i < newCustomers.length; i += 50) {
        const batch = newCustomers.slice(i, i + 50);
        const { error } = await db.from('customers').insert(batch);
        if (error) {
          console.error(`배치 ${i} INSERT 실패:`, error);
        } else {
          inserted += batch.length;
        }
      }
    }

    return NextResponse.json({
      total: ecountList.length,
      existing: existingCodes.size,
      inserted,
      message: `이카운트 거래처 ${ecountList.length}건 중 ${inserted}건 신규 등록`,
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
