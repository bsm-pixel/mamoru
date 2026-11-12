'use client';

import { useState } from 'react';
import { useCreateReturn, useUpdateReturn } from '@/hooks/use-returns';
import { Wrench, Truck, Store } from 'lucide-react';
import toast from 'react-hot-toast';
import { backdropClose } from '@/lib/ui/backdrop';
import { EscClose } from '@/components/ui/esc-close';
import { PICKUP_COURIER, PICKUP_IN_PERSON } from '@/lib/returns/transitions';

/**
 * 점검·조정 후 재발송 접수 (2026-11-12)
 *
 * 판매한 제품을 고객이 불편해할 때 — 같은 제품을 회수해 손본 뒤 다시 보낸다.
 * 교환(새 제품)·반품(환불)과 달리 **판매·재고·시리얼·매출은 그대로**. returns 에 return_type='adjust' 한 건만 만든다.
 * 이후 진행(수거접수 → 입고완료 → 재발송 송장 → 집하)은 반품·교환 메뉴에서. 고객 알림톡은 없음 — 사장님이 직접 연락(11-12 결정).
 *
 * 회수 방법
 *   택배수거 = 롯데 반품수거(반품·교환 메뉴의 [롯데 반품 수거접수] 버튼). 접수 알림톡은 보내지 않는다(사장님이 이미 통화)
 *   이미 받음 = ALPS 에서 직접 수거했거나 직접 받아 이미 사무실에 있음 → 만들자마자 입고완료
 */
interface SaleItem { id: string; product_id?: string | null; product_name: string; category?: string | null; quantity: number; total_price: number; }
interface SerialRow { id: string; serial_number: string; product_id?: string | null; sale_item_id?: string | null; }
interface SaleLike { id: string; customer_id?: string | null; customer_name?: string | null; customer_phone?: string | null; }

export function AdjustReturnModal({ sale, items, serials, onClose }: {
  sale: SaleLike; items: SaleItem[]; serials: SerialRow[]; onClose: () => void;
}) {
  const createReturn = useCreateReturn();
  const updateReturn = useUpdateReturn();
  // 복원수리(RS) 줄은 제품이 아니다 — 제외
  const candidates = items.filter((it) => it.category !== 'RS');
  const [itemId, setItemId] = useState<string | null>(candidates.length === 1 ? candidates[0].id : null);
  const [method, setMethod] = useState<'already' | typeof PICKUP_COURIER>('already');
  const [reason, setReason] = useState('');
  const [busy, setBusy] = useState(false);

  const item = candidates.find((i) => i.id === itemId) || null;
  const itemSerials = (it: SaleItem) => serials.filter(
    (s) => (s.sale_item_id && s.sale_item_id === it.id) || (!s.sale_item_id && s.product_id && s.product_id === it.product_id),
  );
  const serial = item ? itemSerials(item)[0] || null : null;

  async function handleConfirm() {
    if (!item || busy) return;
    setBusy(true);
    try {
      const { return: created } = await createReturn.mutateAsync({
        return_type: 'adjust',
        sale_id: sale.id,
        product_id: item.product_id || null,
        product_name: item.product_name,
        serial_id: serial?.id || null,
        serial_number: serial?.serial_number || null,
        qty: 1,
        customer_id: sale.customer_id || null,
        name: sale.customer_name || null,
        phone: sale.customer_phone || null,
        // '이미 받음' 은 대면수령으로 기록하고 곧바로 입고완료 (알림톡 없음)
        pickup_method: method === 'already' ? PICKUP_IN_PERSON : PICKUP_COURIER,
        reason: reason.trim() || '점검·조정 후 재발송',
      });
      if (method === 'already') {
        await updateReturn.mutateAsync({ id: created.id, status: 'inbound' });
        toast.success(`점검·조정 접수 ${created.return_number} — 입고완료. 반품·교환 메뉴에서 재발송 송장을 발행하세요`);
      } else {
        toast.success(`점검·조정 접수 ${created.return_number} — 반품·교환 메뉴에서 [롯데 반품 수거접수]를 누르세요`);
      }
      onClose();
    } catch (e) {
      toast.error(e instanceof Error ? e.message : '접수 실패');
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4" {...backdropClose(onClose)}>
      <EscClose onClose={onClose} />
      <div className="bg-white rounded-xl shadow-2xl flex flex-col w-full max-w-[480px] max-h-[90vh]" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800 flex items-center gap-1.5"><Wrench size={14} /> 점검·조정 후 재발송</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600 text-lg leading-none">×</button>
        </div>

        <div className="overflow-y-auto flex-1 p-4 space-y-4">
          <p className="text-[11px] text-neutral-500 leading-relaxed">
            같은 제품을 회수해 손본 뒤 다시 보냅니다. <b>판매·재고·시리얼·매출은 바뀌지 않습니다.</b>
            <b>고객 알림톡은 나가지 않습니다</b> — 재발송 후 송장번호로 직접 연락해 주세요.
          </p>

          <div>
            <p className="text-xs font-semibold text-neutral-500 mb-1.5">1. 대상 제품</p>
            {candidates.length === 0 ? (
              <p className="text-xs text-neutral-400 py-3 text-center bg-neutral-50 rounded-lg">대상 제품이 없습니다.</p>
            ) : (
              <div className="space-y-1.5">
                {candidates.map((it) => {
                  const on = itemId === it.id;
                  const srs = itemSerials(it);
                  return (
                    <button key={it.id} onClick={() => setItemId(it.id)}
                      className={`w-full text-left px-3 py-2 rounded-lg border transition ${on ? 'border-stone-900 bg-stone-50' : 'border-neutral-200 hover:border-neutral-300'}`}>
                      <span className="text-sm font-medium text-neutral-800">{it.product_name}</span>
                      {srs.length > 0 && <span className="ml-2 text-[11px] font-mono text-neutral-500">{srs.map((s) => s.serial_number).join(', ')}</span>}
                    </button>
                  );
                })}
              </div>
            )}
          </div>

          <div>
            <p className="text-xs font-semibold text-neutral-500 mb-1.5">2. 회수 방법</p>
            <div className="grid grid-cols-2 gap-1.5">
              {([
                { v: 'already', label: '이미 받았어요', icon: Store },
                { v: PICKUP_COURIER, label: '택배로 회수', icon: Truck },
              ] as const).map((o) => {
                const on = method === o.v;
                return (
                  <button key={o.v} onClick={() => setMethod(o.v)}
                    className={`flex items-center gap-1.5 px-3 py-2 rounded-lg border text-xs font-medium transition ${on ? 'border-stone-900 bg-stone-50 text-stone-900' : 'border-neutral-200 text-neutral-600 hover:border-neutral-300'}`}>
                    <o.icon size={13} /> {o.label}
                  </button>
                );
              })}
            </div>
            <p className="text-[11px] text-neutral-400 mt-1">
              {method === 'already'
                ? 'ALPS에서 직접 수거했거나 직접 받아 이미 사무실에 있는 경우 — 바로 입고완료로 만듭니다.'
                : '반품·교환 메뉴에서 [롯데 반품 수거접수]를 누르면 기사님이 고객집에서 회수합니다. 접수 알림톡은 나가지 않습니다.'}
            </p>
          </div>

          <div>
            <p className="text-xs font-semibold text-neutral-500 mb-1.5">3. 고객이 말한 불편 (선택)</p>
            <input type="text" value={reason} onChange={(e) => setReason(e.target.value)}
              placeholder="예: 틴닝 걸림 — 틴닝 홈 조정"
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          </div>
        </div>

        <div className="flex items-center justify-end gap-2 px-4 py-3 border-t border-neutral-200">
          <button onClick={onClose} className="px-3 py-2 rounded-lg text-sm text-neutral-500 hover:bg-neutral-100">취소</button>
          <button onClick={handleConfirm} disabled={!item || busy}
            className="px-4 py-2 rounded-lg bg-stone-900 text-white text-sm font-semibold hover:bg-stone-800 transition disabled:opacity-50">
            {busy ? '접수 중…' : '점검·조정 접수'}
          </button>
        </div>
      </div>
    </div>
  );
}
