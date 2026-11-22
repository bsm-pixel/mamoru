import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { errMsg } from '@/lib/utils/err';
import { todayKST } from '@/lib/utils/kst';
import { linkContractCustomer } from '@/lib/contracts/link-customer';

/** GET /api/contracts — 계약서 목록 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const url = req.nextUrl;
    const status = url.searchParams.get('status');
    const search = url.searchParams.get('search');
    const page = parseInt(url.searchParams.get('page') || '1');
    const limit = parseInt(url.searchParams.get('limit') || '20');
    const from = (page - 1) * limit;
    const to = from + limit - 1;

    let query = db
      .from('contracts')
      .select('*', { count: 'exact' })
      .order('created_at', { ascending: false })
      .range(from, to);

    if (status && status !== 'all') {
      query = query.eq('status', status);
    }
    if (search) {
      query = query.or(
        `customer_name.ilike.%${search}%,customer_phone.ilike.%${search}%,contract_number.ilike.%${search}%`
      );
    }

    const { data, count, error } = await query;
    if (error) throw error;

    return NextResponse.json({ contracts: data || [], total: count || 0 });
  } catch (err) {
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}

/** POST /api/contracts — 계약서 생성 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();
    const { contract, items } = body;

    // 계약번호: CT-YYYYMMDD-NNN
    const today = todayKST();
    const todayCompact = today.replace(/-/g, '');
    const { count } = await db
      .from('contracts')
      .select('*', { count: 'exact', head: true })
      .gte('created_at', `${today}T00:00:00`);
    const seq = String((count || 0) + 1).padStart(3, '0');
    const contractNumber = `CT-${todayCompact}-${seq}`;

    // 서명 여부로 상태 결정
    const status = contract.signature_data ? 'signed' : 'draft';
    // 172 이니셜 각인: 문구가 있으면 받음, 없으면 「나중에 받기」
    const engraving = !!contract.engraving;
    const engravingText = engraving ? String(contract.engraving_text || '').trim().slice(0, 30) : '';

    const base = {
        contract_number: contractNumber,
        customer_id: contract.customer_id || null, // 2026-11-22 화면이 보내던 고객 연결을 저장하지 않던 것 고침
        customer_name: contract.customer_name,
        customer_phone: contract.customer_phone || null,
        customer_email: contract.customer_email || null,
        customer_address: contract.customer_address || null,
        total_amount: contract.total_amount,
        discount_amount: contract.discount_amount || 0,
        final_amount: contract.final_amount,
        payment_method: contract.payment_method,
        installment_months: contract.installment_months || 0,
        signature_data: contract.signature_data || null,
        signed_at: contract.signature_data ? new Date().toISOString() : null,
        status,
        memo: contract.memo || null,
        // 전자문서 확장 필드
        delivery_method: contract.delivery_method || 'shipping',
        unavailable_days: contract.unavailable_days || null,
        deposit_amount: contract.deposit_amount || 0,
        balance_amount: contract.balance_amount || 0,
        seller_signature: contract.seller_signature || null,
        customer_title: contract.customer_title || null,
        shop_name: contract.shop_name || null,
        shop_address: contract.shop_address || null,
        // 필기 + 상담 연결
        consultation_id: contract.consultation_id || null,
        handwriting_name: contract.handwriting_name || null,
        handwriting_phone: contract.handwriting_phone || null,
        handwriting_address: contract.handwriting_address || null,
        created_by: user.id,
    };
    // 172 칸 — SQL(마이그 172) 실행 전이면 이 칸 없이 다시 저장한다(각인이 있으면 저장하지 않고 알린다)
    const ext = {
      signed_via: contract.signature_data ? 'store' : null,
      terms_agreed_at: contract.signature_data && contract.terms_agreed ? new Date().toISOString() : null, // 2026-11-23 「위 내용을 확인했습니다」
      ...(engraving ? {
        engraving: true,
        engraving_text: engravingText || null,
        engraving_status: engravingText ? 'received' : 'pending',
        engraving_at: engravingText ? new Date().toISOString() : null,
      } : {}),
    };
    let { data: created, error: createError } = await db.from('contracts').insert({ ...base, ...ext }).select().single();
    if (createError && /signed_via|terms_agreed|engraving|column/i.test(String(createError.message || ''))) {
      if (engraving) throw new Error('이니셜 각인 칸이 아직 준비되지 않았습니다 — 데이터베이스 SQL 172 실행이 필요합니다');
      ({ data: created, error: createError } = await db.from('contracts').insert(base).select().single());
    }

    if (createError) throw createError;

    // 항목 생성
    if (items && items.length > 0) {
      const contractItems = items.map((item: Record<string, unknown>) => ({
        contract_id: created.id,
        product_id: item.product_id || null,
        product_name: item.product_name,
        sku: item.sku || null,
        quantity: item.quantity,
        unit_price: item.unit_price,
        total_price: item.total_price,
        option_text: item.option_text || null,
      }));

      const { error: itemsError } = await db
        .from('contract_items')
        .insert(contractItems);

      if (itemsError) throw itemsError;
    }

    // 2026-11-23 이 자리에서 서명한 계약서는 고객을 붙인다(같은 번호 고객 → 그 고객, 없으면 새로).
    //   링크로 보낸 계약서는 고객이 서명할 때 붙는다(/api/contracts/public).
    if (created.signature_data && !created.customer_id) {
      const cid = await linkContractCustomer(db, created, { road: contract.customer_address_road, detail: contract.customer_address_detail, postcode: contract.customer_postcode });
      if (cid) created.customer_id = cid;
    }

    return NextResponse.json({ contract: created, contractNumber });
  } catch (err) {
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}
