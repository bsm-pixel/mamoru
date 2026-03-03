import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { sendNotification, type NotifyTemplate } from '@/lib/notification/make-webhook';

/** POST /api/consultation/notify — 알림톡 발송 트리거 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const body = await req.json();
    const { consultationId, template, extraData } = body as {
      consultationId: string;
      template: NotifyTemplate;
      extraData?: Record<string, string>;
    };

    if (!consultationId || !template) {
      return NextResponse.json({ error: 'consultationId, template 필수' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data: c, error: fetchErr } = await db
      .from('consultations')
      .select('unique_id, name, phone, consultation_type, visit_date, visit_time, address_road, address_detail')
      .eq('id', consultationId)
      .single();

    if (fetchErr || !c) {
      return NextResponse.json({ error: '상담을 찾을 수 없습니다' }, { status: 404 });
    }

    // GAS postMake_ payload 키와 동일하게 매칭
    const address = [c.address_road, c.address_detail].filter(Boolean).join(' ');
    const result = await sendNotification({
      template,
      phone: c.phone,
      name: c.name,
      data: {
        id: c.unique_id,
        type: c.consultation_type === 'store_visit' ? '매장 방문'
            : c.consultation_type === 'field_request' ? '출장 요청'
            : '톡상담',
        type_code: c.consultation_type === 'store_visit' ? 'STORE'
            : c.consultation_type === 'field_request' ? 'FIELD'
            : 'TALK',
        date: c.visit_date || '',
        time: c.visit_time || '',
        address,
        ...extraData,
      },
    });

    if (!result.success) {
      return NextResponse.json({ error: result.error }, { status: 502 });
    }

    return NextResponse.json({ success: true });
  } catch (err) {
    console.error('[notify] 알림톡 발송 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
