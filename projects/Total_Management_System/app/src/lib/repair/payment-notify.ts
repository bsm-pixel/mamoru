import { sendNotification } from '@/lib/notification/make-webhook';

/** 입금확인 알림톡에 필요한 수리 건 필드 */
export interface RepairPaidRow {
  as_id: string;
  name: string;
  phone: string | null;
  service_cost: number | null;
  shipping_fee: number | null;
  total_amount: number | null;
  invoice_number: string | null;
}

/**
 * 복원수리 입금확인 알림톡(as_payment_confirmed) — 발송 경로 SSOT
 *   · PATCH /api/repair/[id] (사장님 수동 입금확인)
 *   · POST /api/bank/deposit (161: 은행 입금 알림 자동 매칭)
 * 무상·현장결제 생략 판단은 호출하는 쪽에서 한다.
 */
export async function sendRepairPaymentConfirmed(data: RepairPaidRow): Promise<void> {
  if (!data.phone) return;
  const result = await sendNotification({
    template: 'as_payment_confirmed',
    phone: data.phone,
    name: data.name,
    data: {
      id: data.as_id,
      as_id: data.as_id,  // 2026-09-13: #{as_id}/#{as_uid} 템플릿 호환
      as_uid: data.as_id,
      as_amount: String(data.service_cost || 0),
      shipping_amount: String(data.shipping_fee || 0),
      total_amount: String(data.total_amount || 0),
      tracking: data.invoice_number || '',
      courier: '롯데택배',
    },
  });
  if (!result.success) console.error('[repair paid_at notify] 실패:', result.error);
  else console.log('[repair paid_at notify] 입금확인 알림톡 발송');
}
