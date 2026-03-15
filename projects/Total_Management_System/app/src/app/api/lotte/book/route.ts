import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { updateInvoice } from '@/lib/imweb/client';

/** POST /api/lotte/book — 송장 생성 (GAS ALPS 경유) */
export async function POST(request: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await request.json();

    // 상품명 자동 조합: order_items에서 가져오기
    if (!body.gdsNm && body.orderId) {
      const { data: items } = await (supabase as any)
        .from('order_items')
        .select('product_name, quantity')
        .eq('order_id', body.orderId);
      if (items && items.length > 0) {
        body.gdsNm = items
          .map((it: { product_name: string; quantity: number }) =>
            it.quantity > 1 ? `${it.product_name} x${it.quantity}` : it.product_name
          )
          .join(', ')
          .slice(0, 750);
      }
    }

    // GAS 경유 ALPS 송장 생성
    const gasUrl = process.env.GAS_AS_URL;
    const adminToken = process.env.GAS_AS_ADMIN_TOKEN;
    if (!gasUrl || !adminToken) {
      return NextResponse.json({ error: 'GAS_AS_URL 또는 GAS_AS_ADMIN_TOKEN 미설정' }, { status: 500 });
    }

    const gasParams = new URLSearchParams({
      action: 'book_order',
      token: adminToken,
      ordNo: body.ordNo || '',
      rcvName: body.rcvName || '',
      rcvTel: body.rcvTel || '',
      rcvZip: body.rcvZip || '',
      rcvAdr: body.rcvAdr || '',
      gdsNm: body.gdsNm || '마모루 제품',
      dlvMsg: body.dlvMsg || '',
      boxTypCd: body.boxTypCd || 'A',
    });

    const gasRes = await fetch(`${gasUrl}?${gasParams}`, {
      method: 'GET',
      redirect: 'follow',
      signal: AbortSignal.timeout(30000),
    });

    const gasText = await gasRes.text();
    let gasBody;
    try { gasBody = JSON.parse(gasText); } catch { gasBody = { ok: false, error: gasText }; }

    if (!gasBody.ok || !gasBody.invNo) {
      return NextResponse.json(
        { error: `GAS 송장 생성 실패: ${gasBody.error || gasBody.msg || gasText}` },
        { status: 502 }
      );
    }

    const result = { ok: true, invNo: gasBody.invNo, rtnCd: gasBody.rtnCd || '', rtnMsg: gasBody.rtnMsg || '' };

    // 주문에 송장 정보 저장
    if (body.orderId) {
      await (supabase as any)
        .from('orders')
        .update({
          invoice_number: result.invNo,
          courier_code: 'LOTTE',
          courier_name: '롯데택배',
          status: 'shipping',
          shipped_at: new Date().toISOString(),
        })
        .eq('id', body.orderId);
    }

    // 아임웹에 송장번호 반영 (배송대기 상태일 때만 성공)
    let imwebSynced = false;
    let imwebNeedsManual = false;
    if (body.ordNo) {
      try {
        const imwebResult = await updateInvoice(body.ordNo, {
          parcel_code: 'LOTTE',
          invoice_no: result.invNo,
        });
        imwebSynced = imwebResult.success;
        imwebNeedsManual = imwebResult.needsManual;
      } catch (imwebErr) {
        console.warn('[lotte/book] 아임웹 반영 실패:', imwebErr);
        imwebNeedsManual = true;
      }
    }

    return NextResponse.json({
      ...result,
      imwebSynced,
      imwebNeedsManual,
    });
  } catch (err) {
    console.error('[lotte/book] 송장 생성 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
