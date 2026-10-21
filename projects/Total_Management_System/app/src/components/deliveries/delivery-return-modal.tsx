'use client';

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { formatKRW } from '@/lib/utils/format';
import { computeDeliveryTotals } from '@/lib/deliveries/totals';
import { isReturnableDeliveryItem } from '@/lib/deliveries/apply-return';
import { PICKUP_COURIER, PICKUP_IN_PERSON } from '@/lib/returns/transitions';
import { Minus, Plus } from 'lucide-react';

interface DLItemLite {
  id: string;
  product_name: string;
  category?: string | null;
  quantity: number;
  unit_price: number;
}

export interface DeliveryReturnSubmit {
  items: Array<{ delivery_item_id: string; qty: number }>;
  pickup_method: typeof PICKUP_COURIER | typeof PICKUP_IN_PERSON;
  reason: string;
}

interface Props {
  open: boolean;
  onClose: () => void;
  items: DLItemLite[];
  /** 아직 입고 안 된 반품이 잡아둔 수량 (품목 id → 수량) */
  pendingQty: Record<string, number>;
  vatType: 'included' | 'separate' | 'none';
  discount: number;
  totalAmount: number;
  paidAmount: number;
  /** 롯데 반품수거 수거지 (그 납품의 별도 배송지 → 없으면 거래처 등록 주소). null = 주소 없음 */
  pickupAddress: string | null;
  busy?: boolean;
  onSubmit: (v: DeliveryReturnSubmit) => void;
}

/**
 * 납품 반품 처리 모달 (164, 2026-10-21) — 전체·부분 반품
 *
 * 회수 방법 기본값 = 롯데 반품수거 접수(누르면 ALPS 까지 즉시, 2026-10-21). 「이미 받음」은 수리 박스에 같이 왔거나
 * 직접 가져온 **예외** 경우라 두 번째 선택지로만 둔다(사장님: 예외를 일반화하지 말 것).
 * 금액·재고는 물건이 입고되는 순간 반영 → 미리보기는 "입고되면 이렇게 바뀝니다".
 */
export function DeliveryReturnModal({
  open, onClose, items, pendingQty, vatType, discount, totalAmount, paidAmount, pickupAddress, busy, onSubmit,
}: Props) {
  const [qty, setQty] = useState<Record<string, number>>({});
  const [method, setMethod] = useState<DeliveryReturnSubmit['pickup_method']>(PICKUP_COURIER);
  const [reason, setReason] = useState('');

  // 열릴 때 초기화 — 렌더 중 보정(React 공식 패턴, ship-address-modal 과 동일)
  const [wasOpen, setWasOpen] = useState(open);
  if (open !== wasOpen) {
    setWasOpen(open);
    if (open) { setQty({}); setMethod(PICKUP_COURIER); setReason(''); }
  }

  const rows = items.filter(isReturnableDeliveryItem).map((it) => ({
    ...it,
    left: Math.max(0, it.quantity - (pendingQty[it.id] || 0)),
  }));
  const pick = (id: string) => qty[id] || 0;
  const totalPick = rows.reduce((s, r) => s + pick(r.id), 0);

  // 입고 후 예상 금액 — 서버 반영(apply-return.ts)과 같은 공식
  const after = computeDeliveryTotals(
    items.map((it) => ({ category: it.category, quantity: it.quantity - pick(it.id), unit_price: it.unit_price })),
    vatType,
    discount,
  ).totalAmount;
  const diff = totalAmount - after;
  const refundDue = Math.max(0, paidAmount - after);

  const setRowQty = (id: string, v: number, max: number) =>
    setQty((p) => ({ ...p, [id]: Math.max(0, Math.min(max, v)) }));

  const inHand = method === PICKUP_IN_PERSON;

  return (
    <Modal open={open} onClose={onClose} title="반품 처리">
      <div className="space-y-4">
        {/* ① 반품 수량 */}
        <div>
          <p className="text-xs font-semibold text-neutral-500 mb-1.5">반품할 수량</p>
          {rows.length === 0 ? (
            <p className="text-sm text-neutral-400 py-2">반품할 수 있는 제품이 없습니다 (복원수리·배송비는 제외)</p>
          ) : (
            <div className="divide-y divide-neutral-100 rounded-lg border border-neutral-200">
              {rows.map((r) => (
                <div key={r.id} className="flex items-center gap-3 px-3 py-2.5">
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-medium truncate">{r.product_name}</p>
                    <p className="text-xs text-neutral-500">
                      납품 {r.quantity}개 · {formatKRW(r.unit_price)}
                      {(pendingQty[r.id] || 0) > 0 && <span className="text-amber-600"> · 회수 대기 {pendingQty[r.id]}개</span>}
                    </p>
                  </div>
                  <div className="flex items-center gap-1 shrink-0">
                    <button type="button" aria-label="수량 빼기" onClick={() => setRowQty(r.id, pick(r.id) - 1, r.left)}
                      disabled={pick(r.id) <= 0}
                      className="w-8 h-8 rounded-lg bg-neutral-100 flex items-center justify-center hover:bg-neutral-200 disabled:opacity-40"><Minus size={14} /></button>
                    <span className="w-8 text-center text-sm font-bold">{pick(r.id)}</span>
                    <button type="button" aria-label="수량 더하기" onClick={() => setRowQty(r.id, pick(r.id) + 1, r.left)}
                      disabled={pick(r.id) >= r.left}
                      className="w-8 h-8 rounded-lg bg-neutral-100 flex items-center justify-center hover:bg-neutral-200 disabled:opacity-40"><Plus size={14} /></button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* ② 회수 방법 — 기본은 롯데 반품수거(누르면 ALPS 즉시 접수) */}
        <div className="space-y-2">
          <p className="text-xs font-semibold text-neutral-500">회수 방법</p>
          <button type="button" onClick={() => setMethod(PICKUP_COURIER)}
            className={`w-full text-left rounded-lg border px-3 py-2.5 transition ${!inHand ? 'border-neutral-900 bg-white' : 'border-neutral-200 bg-white hover:bg-stone-50'}`}>
            <span className="block text-sm font-semibold text-neutral-900">롯데 반품수거 접수</span>
            <span className="block text-xs text-neutral-500 mt-0.5 break-keep">
              누르는 즉시 롯데에 반품수거가 접수됩니다. 물건이 도착해 반품·교환 메뉴에서 [입고완료]를 누르면 금액·재고가 반영됩니다.
            </span>
          </button>
          {/* 롯데 취소 API 가 없어 잘못 접수하면 ALPS 에서 수동 취소 → 어디로 기사님이 가는지 누르기 전에 보여준다 */}
          {!inHand && (
            <p className={`text-xs px-1 break-keep ${pickupAddress ? 'text-neutral-600' : 'text-red-600'}`}>
              {pickupAddress ? `수거지 · ${pickupAddress}` : '수거지 주소가 없습니다 — 거래처 정보에 주소를 등록해 주세요'}
            </p>
          )}
          <button type="button" onClick={() => setMethod(PICKUP_IN_PERSON)}
            className={`w-full text-left rounded-lg border px-3 py-2.5 transition ${inHand ? 'border-neutral-900 bg-white' : 'border-neutral-200 bg-white hover:bg-stone-50'}`}>
            <span className="block text-sm font-semibold text-neutral-900">이미 받았어요</span>
            <span className="block text-xs text-neutral-500 mt-0.5 break-keep">
              직접 가져왔거나 다른 택배(수리 박스 등)에 같이 와서 이미 사무실에 있는 경우. 롯데 접수 없이 지금 바로 반영됩니다.
            </span>
          </button>
        </div>

        {/* ③ 사유 */}
        <input value={reason} onChange={(e) => setReason(e.target.value)} placeholder="반품 사유 (예: 날 불량)"
          className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-neutral-300" />

        {/* ④ 금액 미리보기 */}
        {totalPick > 0 && (
          <div className="rounded-lg bg-stone-50 px-3 py-2.5 text-sm space-y-1">
            <div className="flex justify-between">
              <span className="text-neutral-500">납품 합계</span>
              <span className="font-semibold">{formatKRW(totalAmount)} → {formatKRW(after)}</span>
            </div>
            {refundDue > 0 ? (
              <div className="flex justify-between text-violet-700">
                <span>환불 필요</span><span className="font-semibold">{formatKRW(refundDue)}</span>
              </div>
            ) : (
              <div className="flex justify-between text-neutral-600">
                <span>미수금</span><span className="font-semibold">−{formatKRW(diff)}</span>
              </div>
            )}
            <p className="text-[11px] text-neutral-400 break-keep">
              {inHand ? '지금 반영됩니다. ' : '금액·재고는 물건이 입고될 때 반영됩니다. '}
              매출은 원래 납품한 달에서 빠지고, 돌아온 제품은 반품창고로 들어갑니다(판매가능·아임웹 재고 불변).
            </p>
          </div>
        )}

        <div className="flex gap-2 pt-1">
          <Button variant="secondary" className="flex-1" onClick={onClose} disabled={busy}>취소</Button>
          <Button className="flex-1" disabled={totalPick === 0 || busy || (!inHand && !pickupAddress)} loading={busy}
            onClick={() => onSubmit({
              items: rows.filter((r) => pick(r.id) > 0).map((r) => ({ delivery_item_id: r.id, qty: pick(r.id) })),
              pickup_method: method,
              reason: reason.trim(),
            })}>
            {inHand ? `${totalPick}개 반품 처리` : `${totalPick}개 롯데 반품수거 접수`}
          </Button>
        </div>
      </div>
    </Modal>
  );
}
