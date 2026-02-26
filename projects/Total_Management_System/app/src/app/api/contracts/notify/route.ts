import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * POST /api/contracts/notify — 계약서 알림톡 발송
 * 계약서 PDF/이미지 URL을 포함하여 알림톡 전송
 * 현재는 상태만 업데이트 (솔라피 템플릿 등록 후 실연동)
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { contractId } = await req.json() as { contractId: string };

    const { data: contract, error } = await db
      .from('contracts')
      .select('*')
      .eq('id', contractId)
      .single();

    if (error || !contract) {
      return NextResponse.json({ error: '계약서를 찾을 수 없습니다' }, { status: 404 });
    }

    if (!contract.customer_phone) {
      return NextResponse.json({ error: '고객 연락처가 없습니다' }, { status: 400 });
    }

    // TODO: 솔라피 계약서 템플릿 등록 후 Make webhook 호출
    // await sendNotification({
    //   template: 'contract_sent',
    //   phone: contract.customer_phone,
    //   name: contract.customer_name,
    //   data: {
    //     contract_number: contract.contract_number,
    //     total_amount: String(contract.final_amount),
    //     pdf_url: contract.pdf_url || '',
    //   },
    // });

    // 상태 업데이트
    await db
      .from('contracts')
      .update({
        status: 'sent',
        notification_sent_at: new Date().toISOString(),
      })
      .eq('id', contractId);

    return NextResponse.json({ success: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
