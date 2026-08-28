import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { getNextInvoice, bookShipment } from '@/lib/lotte/alps-client';

/** POST /api/repair/[id]/recall-pickup — 복원수리 재수거(정밀 재점검) 롯데 반품접수(ustRtgSctCd='02')
 *  출고된 복원수리를 고객집에서 다시 회수. 롯데 IS팀 회신: 02=반품, 양식 출고와 동일.
 *  ⚠️ 취소 API 미지원 → 접수 후 취소는 ALPS 수동. 알림톡은 발송하지 않음(사장님 별도 안내).
 *  🔒 고객 노출 문구는 '정밀 재점검'(중립) — 송장 품명도 브랜드 세이프.
 */
export async function POST(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: repair } = await db.from('repairs').select('*').eq('id', id).single();
    if (!repair) return NextResponse.json({ error: '복원수리 건을 찾을 수 없습니다' }, { status: 404 });
    if (repair.recall_invoice_number) {
      return NextResponse.json({ error: '이미 재수거 접수가 되어 있습니다' }, { status: 400 });
    }
    if (!repair.phone) {
      return NextResponse.json({ error: '고객 연락처가 없습니다.' }, { status: 400 });
    }
    if (!repair.postcode || !repair.address) {
      return NextResponse.json({ error: '고객 주소(우편번호+주소)가 없습니다.' }, { status: 400 });
    }

    const fullAddress = [repair.address, repair.address_detail].filter(Boolean).join(' ');
    const { invoiceNumber } = await getNextInvoice();
    const result = await bookShipment({
      invoiceNumber,
      receiverName: repair.name,
      receiverTel: repair.phone || '',
      receiverZip: repair.postcode || '',
      receiverAddr: fullAddress,
      goodsName: '[MAMORU] 정밀 재점검 회수',   // 고객 노출 대비 중립 문구
      deliveryMessage: '정밀 재점검 수거',
      ustRtgSctCd: '02',                          // 반품(회수)
    });
    if (!result.success) {
      return NextResponse.json({ error: `롯데 재수거 접수 실패: ${result.error}` }, { status: 502 });
    }

    await db.from('repairs').update({
      recall_invoice_number: invoiceNumber,
      recall_courier_name: '롯데택배',
      recall_booked_at: new Date().toISOString(),
    }).eq('id', id);

    await db.from('repair_history').insert({
      repair_id: id,
      from_status: repair.status,
      to_status: repair.status,
      changed_by: user.id,
      note: `재수거 접수(정밀 재점검): ${invoiceNumber}`,
    });

    return NextResponse.json({ success: true, invoiceNumber });
  } catch (err) {
    console.error('[repair] 재수거 접수 실패:', err);
    return NextResponse.json({ error: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }
}
