import { NextRequest, NextResponse, after } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { isValidRepairTransition } from '@/lib/repair/transitions';
import { sendNotification, type NotifyTemplate } from '@/lib/notification/make-webhook';
import type { RepairStatus } from '@/lib/supabase/types';

/** 상태→자동 알림톡 매핑 (payment_confirmed는 paid_at 플래그로 분리됨) */
function getAutoNotifyTemplate(newStatus: string): NotifyTemplate | null {
  const map: Record<string, NotifyTemplate> = {
    shipped: 'as_shipped',
  };
  return map[newStatus] || null;
}

/** GET /api/repair/[id] — 단건 + 검수 + 이력 */
export async function GET(
  _req: NextRequest,
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
    const [repairRes, inspectionsRes, historyRes] = await Promise.all([
      db.from('repairs').select('*').eq('id', id).single(),
      db.from('repair_inspections').select('*').eq('repair_id', id).order('scissor_number', { ascending: true }),
      db.from('repair_history').select('*').eq('repair_id', id).order('created_at', { ascending: false }),
    ]);

    if (repairRes.error) throw repairRes.error;

    return NextResponse.json({
      repair: repairRes.data,
      inspections: inspectionsRes.data || [],
      history: historyRes.data || [],
    });
  } catch (err) {
    console.error('[repair] 상세 조회 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/repair/[id] — 상태/정보 변경 */
export async function PATCH(
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
    const { status: newStatus, note, ...rest } = body;

    // 현재 조회
    const { data: current, error: fetchErr } = await db
      .from('repairs')
      .select('*')
      .eq('id', id)
      .single();

    if (fetchErr || !current) {
      return NextResponse.json({ error: '복원수리 건을 찾을 수 없습니다' }, { status: 404 });
    }

    // 상태 전이 검증
    if (newStatus && newStatus !== current.status) {
      const valid = isValidRepairTransition(
        current.status as RepairStatus,
        newStatus as RepairStatus
      );
      if (!valid) {
        return NextResponse.json(
          { error: `상태 전이 불가: ${current.status} → ${newStatus}` },
          { status: 400 }
        );
      }
    }

    // 업데이트 데이터
    const updateData = { ...rest };
    if (newStatus) updateData.status = newStatus;

    const { data, error } = await db
      .from('repairs')
      .update(updateData)
      .eq('id', id)
      .select()
      .single();

    if (error) throw error;

    // paid_at 설정 시 → 입금확인 알림톡 (상태 변경과 독립)
    const justPaid = rest.paid_at && !current.paid_at;

    // 상태 변경 시 이력 기록 + 백그라운드 알림톡
    if (newStatus && newStatus !== current.status) {
      await db.from('repair_history').insert({
        repair_id: id,
        from_status: current.status,
        to_status: newStatus,
        changed_by: user.id,
        note: note || null,
      });

      after(async () => {
        const template = getAutoNotifyTemplate(newStatus);
        if (template && data.phone) {
          const result = await sendNotification({
            template,
            phone: data.phone,
            name: data.name,
            data: {
              id: data.as_id,
              as_amount: String(data.service_cost || 0),
              shipping_amount: String(data.shipping_fee || 0),
              total_amount: String(data.total_amount || 0),
              tracking: data.invoice_number || '',
            },
          });
          if (!result.success) console.error('[repair auto-notify] 실패:', result.error);
          else console.log('[repair auto-notify] 성공:', template);
        }
      });
    }

    // 입금확인 알림톡 (paid_at 플래그 설정 시)
    if (justPaid) {
      after(async () => {
        if (data.phone) {
          const result = await sendNotification({
            template: 'as_payment_confirmed',
            phone: data.phone,
            name: data.name,
            data: {
              id: data.as_id,
              as_amount: String(data.service_cost || 0),
              shipping_amount: String(data.shipping_fee || 0),
              total_amount: String(data.total_amount || 0),
              tracking: data.invoice_number || '',
            },
          });
          if (!result.success) console.error('[repair paid_at notify] 실패:', result.error);
          else console.log('[repair paid_at notify] 입금확인 알림톡 발송');
        }
      });
    }

    return NextResponse.json(data);
  } catch (err) {
    console.error('[repair] 업데이트 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
