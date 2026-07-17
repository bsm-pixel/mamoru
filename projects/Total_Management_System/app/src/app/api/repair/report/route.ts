import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { generateWorkSummary } from '@/lib/repair/inspection-text';

/** CORS 헤더 — GitHub Pages에서 호출 허용 */
const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

/** OPTIONS preflight */
export async function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** GET /api/repair/report?as_id=AS-XXXXXX-NNN — 공개 API (수리내역 페이지용) */
export async function GET(req: NextRequest) {
  try {
    const asId = req.nextUrl.searchParams.get('as_id');
    if (!asId) {
      return NextResponse.json({ error: 'as_id 필수' }, { status: 400, headers: CORS_HEADERS });
    }

    // Service Role — 인증 불필요 (공개 API)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const supabase: any = createServiceClient();

    // 복원수리 조회
    const { data: repair, error: repairErr } = await supabase
      .from('repairs')
      .select('*')
      .eq('as_id', asId)
      .single();

    if (repairErr || !repair) {
      // 🔴 CORS 헤더 필수 — 없으면 GitHub Pages(page.mamoru.kr)의 fetch 가 응답을 차단해
      //    고객이 "수리 내역을 찾을 수 없습니다"(정상 안내) 대신 "서버에 연결할 수 없습니다"를 봄
      return NextResponse.json({ error: '복원수리 건을 찾을 수 없습니다' }, { status: 404, headers: CORS_HEADERS });
    }

    // 검수 데이터 조회
    const { data: inspections } = await supabase
      .from('repair_inspections')
      .select('*')
      .eq('repair_id', repair.id)
      .order('scissor_number', { ascending: true });

    const scissors = (inspections || []).map((insp: Record<string, unknown>) => ({
      number: insp.scissor_number,
      type: insp.scissor_type || '',
      blade_tip: insp.blade_tip,
      blade_mid: insp.blade_mid,
      blade_inner: insp.blade_inner,
      comb: insp.comb,
      tension: insp.tension,
      parts: insp.parts,
      stopper: insp.stopper,
      photo: insp.photo_url || '',
      photo_marks: insp.photo_marks || [],
      comment: insp.comment || '',  // 097: 가위별 진단 및 내역
      worker: insp.worker,
    }));

    // 타입별 카운트
    const typeCounts: Record<string, number> = {};
    for (const s of scissors) {
      const t = s.type || '기타';
      typeCounts[t] = (typeCounts[t] || 0) + 1;
    }

    // 자동 문구
    const workSummary = generateWorkSummary(inspections || []);

    return NextResponse.json({
      success: true,
      as_id: repair.as_id,
      customer_name: repair.name,
      qty_mamoru: repair.qty_mamoru || 0,
      qty_other: repair.qty_other || 0,
      total_scissors: (repair.qty_mamoru || 0) + (repair.qty_other || 0),
      service_cost: repair.service_cost || 0,
      shipping_fee: repair.shipping_fee || 0,
      total_amount: repair.total_amount || 0,
      in_date: repair.received_at ? new Date(repair.received_at).toLocaleDateString('ko-KR') : '',
      out_date: repair.shipped_at ? new Date(repair.shipped_at).toLocaleDateString('ko-KR') : '',
      worker: scissors[0]?.worker || '백성민',
      comment: repair.admin_note || '',
      work_summary: workSummary,
      scissors,
      type_counts: typeCounts,
    }, { headers: CORS_HEADERS });
  } catch (err) {
    console.error('[repair/report] 조회 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
