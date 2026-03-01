import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { saveCustomer } from '@/lib/ecount/customer';
import { isStatusOk } from '@/lib/ecount/client';

/** 이카운트 거래처 코드 채번: MM-001, MM-002, ... */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function generateEcountCustCode(db: any): Promise<string> {
  const { data } = await db
    .from('customers')
    .select('ecount_customer_code')
    .not('ecount_customer_code', 'is', null)
    .like('ecount_customer_code', 'MM-%')
    .order('ecount_customer_code', { ascending: false })
    .limit(1);

  let nextNum = 1;
  if (data && data.length > 0 && data[0].ecount_customer_code) {
    const match = data[0].ecount_customer_code.match(/MM-(\d+)/);
    if (match) nextNum = parseInt(match[1]) + 1;
  }

  const padLen = Math.max(3, String(nextNum).length);
  return `MM-${String(nextNum).padStart(padLen, '0')}`;
}

/** POST /api/customers — 고객 신규 등록 (TMS + 이카운트 거래처 동시) */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const { name, phone, email, address, postcode, addressDetail } = body as {
      name: string;
      phone?: string;
      email?: string;
      address?: string;
      postcode?: string;
      addressDetail?: string;
    };

    if (!name?.trim()) {
      return NextResponse.json({ error: '고객명은 필수입니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 1) 이카운트 거래처 코드 채번 + 등록
    let ecountCode: string | null = null;
    let ecountSynced = false;

    try {
      const custCode = await generateEcountCustCode(db);
      const result = await saveCustomer({
        custCode,
        custName: name.trim(),
        phone: phone || undefined,
        email: email || undefined,
        address: [address, addressDetail].filter(Boolean).join(' ') || undefined,
        remarks: 'TMS 자동등록',
      });

      if (isStatusOk(result.Status)) {
        ecountCode = custCode;
        ecountSynced = true;
      }
    } catch (e) {
      // 이카운트 실패해도 TMS 등록은 진행
      console.error('이카운트 거래처 등록 실패:', e);
    }

    // 2) TMS customers 테이블 INSERT
    const { data: customer, error } = await db
      .from('customers')
      .insert({
        name: name.trim(),
        phone: phone || null,
        email: email || null,
        postcode: postcode || null,
        address_road: address || null,
        address_detail: addressDetail || null,
        source: 'manual',
        ecount_customer_code: ecountCode,
      })
      .select('id, name, phone, email, address_road, address_detail, postcode, ecount_customer_code, source')
      .single();

    if (error) throw error;

    return NextResponse.json({
      customer,
      ecountSynced,
      ecountCode,
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
