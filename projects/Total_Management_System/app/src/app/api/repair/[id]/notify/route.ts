import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { sendNotification, type NotifyTemplate } from '@/lib/notification/make-webhook';

/** POST /api/repair/[id]/notify — 수동 알림톡 발송 */
export async function POST(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();
    const { template, extraData } = body;

    if (!template) {
      return NextResponse.json({ error: 'template 필수' }, { status: 400 });
    }

    // 복원수리 조회
    const { data: repair, error } = await db
      .from('repairs')
      .select('*')
      .eq('id', id)
      .single();

    if (error || !repair) {
      return NextResponse.json({ error: '복원수리 건을 찾을 수 없습니다' }, { status: 404 });
    }

    const result = await sendNotification({
      template: template as NotifyTemplate,
      phone: repair.phone,
      name: repair.name,
      data: {
        id: repair.as_id,
        as_amount: String(repair.service_cost || 0),
        shipping_amount: String(repair.shipping_fee || 0),
        total_amount: String(repair.total_amount || 0),
        tracking: repair.invoice_number || '',
        courier: '롯데택배',
        proceed_type: repair.proceed_type || '',
        qty_mamoru: String(repair.qty_mamoru || 0),
        qty_other: String(repair.qty_other || 0),
        ...extraData,
      },
    });

    if (!result.success) {
      return NextResponse.json({ error: result.error }, { status: 502 });
    }

    return NextResponse.json({ success: true });
  } catch (err) {
    console.error('[repair] 알림톡 발송 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
