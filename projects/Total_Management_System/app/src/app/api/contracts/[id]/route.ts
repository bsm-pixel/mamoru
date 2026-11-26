import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { errMsg } from '@/lib/utils/err';
import { cleanEngraving } from '@/lib/contracts/sign-link';
import { syncSaleEngravingMemo } from '@/lib/contracts/engraving-memo';

/**
 * PATCH /api/contracts/[id] — 이니셜 각인 문구 적기 (2026-11-22, 마이그 172)
 *  body: { engraving_text } — 톡 등으로 나중에 받은 문구를 직원이 적어 넣는다. 비우면 다시 「나중에 받기」.
 *  각인을 하지 않는 계약에 각인을 켜는 것도 여기서: { engraving: true }
 */
export async function PATCH(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    const { id } = await params;
    const body = await req.json().catch(() => ({}));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const patch: Record<string, unknown> = {};
    if (body.engraving === false) {
      Object.assign(patch, { engraving: false, engraving_text: null, engraving_status: null, engraving_at: null });
    } else if ('engraving_text' in body || body.engraving === true) {
      const text = cleanEngraving(body.engraving_text);
      Object.assign(patch, { engraving: true, engraving_text: text || null, engraving_status: text ? 'received' : 'pending', engraving_at: text ? new Date().toISOString() : null });
    }
    if (!Object.keys(patch).length) return NextResponse.json({ error: '바꿀 내용이 없습니다' }, { status: 400 });

    const { data, error } = await db.from('contracts').update(patch).eq('id', id).select('id, engraving, engraving_text, engraving_status').single();
    if (error) throw error;
    // 2026-11-26 이미 판매로 바꾼 계약이면 판매 메모의 각인 부분도 맞춘다(출고 준비표에 「각인 문구 대기」가 남던 문제)
    await syncSaleEngravingMemo(db, id, { on: !!data.engraving, text: data.engraving_text || null });
    return NextResponse.json({ contract: data });
  } catch (err) {
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}

/**
 * DELETE /api/contracts/[id] — 계약서 영구 삭제
 *
 * 가드:
 *  - offline_sale_id != null → 409 Conflict (판매 전환된 계약은 삭제 불가)
 *
 * 종속 처리 (트랜잭션 안전 순서):
 *  1) product_serials.contract_id = NULL (orphan FK 방지)
 *  2) contracts DELETE → contract_items 자동 CASCADE
 *
 * 주의: signature_data / seller_signature / image_url 은 contracts 컬럼이라 자동 소실
 */
export async function DELETE(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { id } = await params;
    if (!id) return NextResponse.json({ error: '계약서 ID가 필요합니다' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 1) 가드 — 계약서 존재 + 판매 전환 여부
    const { data: contract, error: fetchError } = await db
      .from('contracts')
      .select('id, contract_number, offline_sale_id')
      .eq('id', id)
      .single();

    if (fetchError || !contract) {
      return NextResponse.json({ error: '계약서를 찾을 수 없습니다' }, { status: 404 });
    }

    if (contract.offline_sale_id) {
      return NextResponse.json(
        { error: '판매 전환된 계약은 삭제할 수 없습니다. 먼저 연결된 판매 건을 취소해주세요.' },
        { status: 409 }
      );
    }

    // 2) product_serials.contract_id NULL 처리 (FK orphan 방지)
    const { error: serialError } = await db
      .from('product_serials')
      .update({ contract_id: null })
      .eq('contract_id', id);

    if (serialError) {
      return NextResponse.json(
        { error: `시리얼 연결 해제 실패: ${serialError.message}` },
        { status: 500 }
      );
    }

    // 3) 계약서 DELETE (contract_items 는 ON DELETE CASCADE 로 자동 정리)
    const { error: deleteError } = await db
      .from('contracts')
      .delete()
      .eq('id', id);

    if (deleteError) {
      return NextResponse.json(
        { error: `계약서 삭제 실패: ${deleteError.message}` },
        { status: 500 }
      );
    }

    return NextResponse.json({
      success: true,
      contract_number: contract.contract_number,
    });
  } catch (err) {
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}
