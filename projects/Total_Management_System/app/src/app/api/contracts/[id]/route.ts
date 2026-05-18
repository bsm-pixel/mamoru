import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

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
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
