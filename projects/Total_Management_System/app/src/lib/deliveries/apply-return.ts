/* eslint-disable @typescript-eslint/no-explicit-any */
/**
 * 납품(B2B) 반품 반영 — 단일 출처 (164, 2026-10-21)
 *
 * 언제: 반품 물건이 **실제로 입고된 순간** 딱 한 번.
 *   · [반품 처리]에서 「이미 받음」 → 접수와 동시에
 *   · 「택배수거」 → 반품·교환 메뉴에서 [입고완료] 누를 때 (api/returns/[id] PATCH)
 *   수거 접수만 하고 아직 안 온 건은 금액·재고가 그대로다 → 수거가 취소돼도 되돌릴 것이 없다.
 *
 * 무엇을:
 *   ① delivery_items.quantity −= 반품수량, returned_qty += 반품수량, total_price 재계산
 *   ② 합계(total/supply/vat)는 computeDeliveryTotals 로 다시 계산 — 매출 집계 15곳이 수정 없이 자동으로 맞는다.
 *      매출은 **원래 납품한 달**에서 빠진다(사장님 확정 2026-10-21).
 *   ③ products.return_stock += 반품수량 (반품창고 — 판매가능·아임웹 불변. 재고 화면 [재판매]/[불량]으로 정리)
 *   ④ 결제액이 새 합계보다 크면 결제액을 합계로 맞추고 차액을 returns.refund_amount 에 「환불 필요」로 남긴다.
 *      (납품 입금은 원래 cash_transactions 에 기록하지 않으므로 환불도 쓰지 않는다 — 장부 일관성)
 *   ⑤ recalcOutstanding 로 미수금 재계산
 *
 * 멱등: returns.applied_at 을 조건부(CAS)로 먼저 찍고, 찍은 쪽만 반영한다 → 두 번 눌러도 한 번만 빠진다.
 */
import { computeDeliveryTotals } from '@/lib/deliveries/totals';
import { recalcOutstanding } from '@/lib/outstanding';

/** 반품 품목 한 줄 (returns.items) */
export interface DeliveryReturnItem {
  delivery_item_id: string;
  product_id: string | null;
  product_name: string;
  category: string | null;
  qty: number;
  unit_price: number;
}

/** 반품할 수 있는 품목인가 — 복원수리·배송비(category='RS')는 물건이 아니라서 제외 */
export function isReturnableDeliveryItem(item: { category?: string | null; quantity?: number | null }): boolean {
  return item.category !== 'RS' && (item.quantity || 0) > 0;
}

/** returns.product_name 요약 — "NS263F 2" / "NS263F 1 외 2종" */
export function summarizeReturnItems(items: DeliveryReturnItem[]): string {
  if (items.length === 0) return '반품 상품';
  const first = `${items[0].product_name} ${items[0].qty}`;
  return items.length === 1 ? first : `${first} 외 ${items.length - 1}종`;
}

export interface ApplyDeliveryReturnResult {
  applied: boolean;
  refundDue: number;
  returnStatus: 'partial' | 'full' | null;
}

export async function applyDeliveryReturn(db: any, returnId: string): Promise<ApplyDeliveryReturnResult> {
  const none: ApplyDeliveryReturnResult = { applied: false, refundDue: 0, returnStatus: null };

  // CAS — 아직 반영 안 된 건만 잡는다
  const nowIso = new Date().toISOString();
  const { data: ret } = await db.from('returns')
    .update({ applied_at: nowIso })
    .eq('id', returnId).is('applied_at', null).not('delivery_id', 'is', null)
    .select('*').maybeSingle();
  if (!ret) return none;

  let wrote = false;   // 한 줄이라도 썼는지 — 쓴 뒤 실패면 가드를 풀지 않는다(재시도 시 이중 차감 방지)
  try {
    const { data: dl } = await db.from('deliveries').select('*').eq('id', ret.delivery_id).single();
    const { data: dlItems } = await db.from('delivery_items').select('*').eq('delivery_id', ret.delivery_id);
    if (!dl || !dlItems) throw new Error('납품 건을 찾을 수 없습니다');

    const retItems = (Array.isArray(ret.items) ? ret.items : []) as DeliveryReturnItem[];
    const byId = new Map<string, any>(dlItems.map((d: any) => [d.id, d]));

    // ① 품목 수량 차감 + ③ 반품창고
    for (const ri of retItems) {
      const di = byId.get(ri.delivery_item_id);
      if (!di || di.category === 'RS') continue;
      const qty = Math.min(Math.max(0, Math.floor(ri.qty || 0)), di.quantity || 0);   // 남은 수량을 넘지 않게
      if (qty <= 0) continue;
      const newQty = (di.quantity || 0) - qty;
      const { error: iErr } = await db.from('delivery_items').update({
        quantity: newQty,
        returned_qty: (di.returned_qty || 0) + qty,
        total_price: newQty * (di.unit_price || 0),
      }).eq('id', di.id);
      if (iErr) throw iErr;
      wrote = true;
      di.quantity = newQty;   // 아래 합계 재계산에 쓰도록 메모리 값도 갱신

      if (di.product_id) {
        const { data: p } = await db.from('products').select('return_stock').eq('id', di.product_id).single();
        await db.from('products').update({ return_stock: ((p?.return_stock as number) || 0) + qty }).eq('id', di.product_id);
      }
    }

    // ② 합계 재계산 (생성·편집과 같은 공식 — RS 무세·할인 소진 규칙 유지)
    const vatType = (dl.vat_type || 'included') as 'included' | 'separate' | 'none';
    const { supplyAmount, vatAmount, totalAmount } = computeDeliveryTotals(
      dlItems.map((d: any) => ({ category: d.category, quantity: d.quantity || 0, unit_price: d.unit_price || 0 })),
      vatType,
      dl.discount_amount || 0,
    );

    // ④ 결제액 정리
    const paid = dl.paid_amount || 0;
    const newPaid = Math.min(paid, totalAmount);
    const refundDue = paid - newPaid;
    const paymentStatus = newPaid >= totalAmount ? 'paid' : newPaid > 0 ? 'partial' : 'unpaid';

    const returnStatus: 'partial' | 'full' = dlItems.every((d: any) => (d.quantity || 0) <= 0) ? 'full' : 'partial';

    const { error: dErr } = await db.from('deliveries').update({
      total_amount: totalAmount,
      supply_amount: supplyAmount,
      vat_amount: vatAmount,
      paid_amount: newPaid,
      payment_status: paymentStatus,
      return_status: returnStatus,
      returned_at: nowIso,
      updated_at: nowIso,
    }).eq('id', dl.id);
    if (dErr) throw dErr;

    if (refundDue > 0) {
      await db.from('returns').update({ refund_amount: refundDue }).eq('id', returnId);
    }

    // ⑤ 미수금
    if (dl.customer_id) await recalcOutstanding(db, dl.customer_id);

    return { applied: true, refundDue, returnStatus };
  } catch (e) {
    // 아무것도 안 쓴 실패 → 가드를 풀어 다시 시도 가능. 일부라도 쓴 뒤 실패 → 가드 유지(이중 차감 방지) + 로그로 수동 확인
    console.error(`[applyDeliveryReturn] 실패 (${wrote ? '일부 반영됨 — 수동 확인 필요' : '반영 없음 — 재시도 가능'}):`, returnId, e);
    if (!wrote) await db.from('returns').update({ applied_at: null }).eq('id', returnId);
    throw e;
  }
}
