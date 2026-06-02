/* ──────────────────────────────────────────────────────────────
   고객 미수금(outstanding_balance) 멱등 재계산
   ──────────────────────────────────────────────────────────────
   기존: 판매/납품 변경 시 outstanding_balance 를 ±diff 로 누적 → 금액편집·취소·
        순서꼬임 때 실제값과 drift (예: 2026-06-02 김은혜 650000 잔여).
   변경: 변경이 일어난 고객의 outstanding_balance 를 "실제 미결제 합"으로 항상 재설정.
        매 호출이 자기교정(idempotent)되어 drift 원천 차단.

   미수 정의:
   - offline_sales: 취소(cancelled_at)·반품(returned_at)·완납(payment_status='paid') 제외,
                    (total - discount - paid) 합
   - deliveries(B2B): 취소·완납 제외, (total - paid) 합
   ────────────────────────────────────────────────────────────── */
/* eslint-disable @typescript-eslint/no-explicit-any */

export async function recalcOutstanding(db: any, customerId: string | null | undefined): Promise<void> {
  if (!customerId) return;

  const [salesRes, dlRes] = await Promise.all([
    db.from('offline_sales')
      .select('total_amount, discount_amount, paid_amount')
      .eq('customer_id', customerId)
      .neq('payment_status', 'paid')
      .is('cancelled_at', null)
      .is('returned_at', null),
    db.from('deliveries')
      .select('total_amount, paid_amount')
      .eq('customer_id', customerId)
      .neq('payment_status', 'paid')
      .is('cancelled_at', null),
  ]);

  const salesUnpaid = (salesRes.data || []).reduce(
    (a: number, r: any) => a + Math.max(0, (r.total_amount || 0) - (r.discount_amount || 0) - (r.paid_amount || 0)),
    0
  );
  const dlUnpaid = (dlRes.data || []).reduce(
    (a: number, r: any) => a + Math.max(0, (r.total_amount || 0) - (r.paid_amount || 0)),
    0
  );

  const outstanding = Math.max(0, salesUnpaid + dlUnpaid);
  await db.from('customers').update({ outstanding_balance: outstanding }).eq('id', customerId);
}
