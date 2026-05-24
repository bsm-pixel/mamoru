'use client';

import { useEffect, useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Package, Truck, AlertCircle, CheckCircle } from 'lucide-react';
import { formatKRW, formatDateTime } from '@/lib/utils/format';
import toast from 'react-hot-toast';

interface MergedShipModalProps {
  open: boolean;
  onClose: () => void;
  repairId: string;
  /** 처리 성공 후 콜백 (React Query invalidate 등) */
  onSuccess?: () => void;
}

interface SaleItem {
  id: string;
  sale_number: string;
  customer_name: string;
  customer_phone: string | null;
  sale_date: string;
  total_amount: number;
  invoice_number: string;
  courier_name: string | null;
  shipped_at: string | null;
}

interface OrderItem {
  id: string;
  imweb_order_no: string;
  orderer_name: string;
  orderer_phone: string | null;
  total_price: number;
  invoice_number: string;
  shipped_at: string | null;
  status: string;
  paid_at: string | null;
}

interface RelatedShipmentsResponse {
  phone: string;
  name: string;
  sales: SaleItem[];
  orders: OrderItem[];
}

type SelectedShipment =
  | { type: 'sale'; item: SaleItem }
  | { type: 'order'; item: OrderItem }
  | { type: 'manual'; invoice: string; courier: string }
  | null;

/**
 * 복원수리 합포장 출고 모달
 *
 * 흐름:
 *   1. 모달 열기 → 같은 고객의 송장 보유 판매/주문건 자동 검색
 *   2. 검색 결과 리스트에서 1클릭 선택 (또는 fallback 직접 송장번호 입력)
 *   3. confirm → API 호출 → 송장 복사 + status='shipped' + 알림톡 발송
 */
export function MergedShipModal({ open, onClose, repairId, onSuccess }: MergedShipModalProps) {
  const [loading, setLoading] = useState(false);
  const [data, setData] = useState<RelatedShipmentsResponse | null>(null);
  const [selected, setSelected] = useState<SelectedShipment>(null);
  const [manualInvoice, setManualInvoice] = useState('');
  const [manualCourier, setManualCourier] = useState('롯데택배');
  const [skipNotify, setSkipNotify] = useState(false);
  const [submitting, setSubmitting] = useState(false);

  // 모달 열릴 때 검색
  useEffect(() => {
    if (!open) {
      setData(null);
      setSelected(null);
      setManualInvoice('');
      setManualCourier('롯데택배');
      setSkipNotify(false);
      return;
    }

    setLoading(true);
    fetch(`/api/repair/${repairId}/related-shipments`)
      .then(r => r.json())
      .then((d: RelatedShipmentsResponse) => setData(d))
      .catch(err => {
        console.error('[merged-ship-modal] 검색 실패:', err);
        toast.error('판매건 검색 실패');
      })
      .finally(() => setLoading(false));
  }, [open, repairId]);

  const handleSubmit = async () => {
    if (!selected) {
      toast.error('송장을 선택하거나 직접 입력해주세요');
      return;
    }

    const payload =
      selected.type === 'sale'
        ? {
            invoice_number: selected.item.invoice_number,
            courier_name: selected.item.courier_name || '롯데택배',
            source_type: 'sale',
            source_id: selected.item.id,
            skip_notify: skipNotify,
          }
        : selected.type === 'order'
        ? {
            invoice_number: selected.item.invoice_number,
            courier_name: '롯데택배', // orders 는 courier 컬럼 없음 (기본값)
            source_type: 'order',
            source_id: selected.item.id,
            skip_notify: skipNotify,
          }
        : {
            invoice_number: selected.invoice.trim(),
            courier_name: selected.courier.trim() || '롯데택배',
            source_type: 'manual',
            skip_notify: skipNotify,
          };

    if (selected.type === 'manual' && !payload.invoice_number) {
      toast.error('송장번호를 입력해주세요');
      return;
    }

    setSubmitting(true);
    try {
      const res = await fetch(`/api/repair/${repairId}/merged-ship`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      const result = await res.json();
      if (!res.ok) throw new Error(result.error || '처리 실패');
      toast.success(skipNotify ? '합포장 출고 처리 완료 (알림톡 미발송)' : '합포장 출고 처리 완료 + 알림톡 발송');
      onSuccess?.();
      onClose();
    } catch (err) {
      toast.error('합포장 출고 실패: ' + String(err));
    } finally {
      setSubmitting(false);
    }
  };

  const hasResults = data && (data.sales.length > 0 || data.orders.length > 0);

  return (
    <Modal open={open} onClose={onClose} title="판매건 합포장 출고" className="max-w-2xl">
      <div className="space-y-4">
        {/* 안내 */}
        <div className="flex items-start gap-2 p-3 bg-warm-ivory rounded-lg text-sm text-neutral-700">
          <AlertCircle size={16} className="flex-shrink-0 mt-0.5 text-neutral-500" />
          <div>
            <p>같은 고객의 다른 주문 송장에 복원수리를 합쳐 발송한 경우 선택하세요.</p>
            <p className="text-xs text-neutral-500 mt-1">선택한 송장번호가 복원수리 출고 알림톡 (수리내역 조회 + 배송조회) 에 사용됩니다.</p>
          </div>
        </div>

        {/* 로딩 */}
        {loading && (
          <div className="text-center py-8 text-sm text-neutral-500">
            같은 고객의 송장 보유 판매/주문건 검색 중...
          </div>
        )}

        {/* 검색 결과 — 판매(offline_sales) */}
        {!loading && data && data.sales.length > 0 && (
          <div>
            <h3 className="text-sm font-bold mb-2 flex items-center gap-1.5">
              <Package size={14} />
              오프라인 판매 ({data.sales.length}건)
            </h3>
            <div className="space-y-1.5">
              {data.sales.map(s => (
                <button
                  key={s.id}
                  onClick={() => setSelected({ type: 'sale', item: s })}
                  className={`w-full text-left p-3 rounded-lg border transition ${
                    selected?.type === 'sale' && selected.item.id === s.id
                      ? 'border-indigo-black bg-warm-ivory'
                      : 'border-neutral-200 hover:border-neutral-400'
                  }`}
                >
                  <div className="flex items-center justify-between text-sm">
                    <span className="font-mono font-bold">{s.sale_number}</span>
                    <span className="text-xs text-neutral-500">{s.sale_date}</span>
                  </div>
                  <div className="text-xs text-neutral-600 mt-1 flex items-center gap-3">
                    <span className="font-mono">{s.invoice_number}</span>
                    <span>{s.courier_name || '롯데택배'}</span>
                    <span>{formatKRW(s.total_amount)}</span>
                  </div>
                  {selected?.type === 'sale' && selected.item.id === s.id && (
                    <div className="mt-1 text-xs text-indigo-black font-medium flex items-center gap-1">
                      <CheckCircle size={12} /> 선택됨
                    </div>
                  )}
                </button>
              ))}
            </div>
          </div>
        )}

        {/* 검색 결과 — 아임웹 주문 */}
        {!loading && data && data.orders.length > 0 && (
          <div>
            <h3 className="text-sm font-bold mb-2 flex items-center gap-1.5">
              <Truck size={14} />
              아임웹 주문 ({data.orders.length}건)
            </h3>
            <div className="space-y-1.5">
              {data.orders.map(o => (
                <button
                  key={o.id}
                  onClick={() => setSelected({ type: 'order', item: o })}
                  className={`w-full text-left p-3 rounded-lg border transition ${
                    selected?.type === 'order' && selected.item.id === o.id
                      ? 'border-indigo-black bg-warm-ivory'
                      : 'border-neutral-200 hover:border-neutral-400'
                  }`}
                >
                  <div className="flex items-center justify-between text-sm">
                    <span className="font-mono font-bold">{o.imweb_order_no}</span>
                    <span className="text-xs text-neutral-500">
                      {o.shipped_at ? formatDateTime(o.shipped_at) : (o.paid_at ? formatDateTime(o.paid_at) : '-')}
                    </span>
                  </div>
                  <div className="text-xs text-neutral-600 mt-1 flex items-center gap-3">
                    <span className="font-mono">{o.invoice_number}</span>
                    <span>{formatKRW(o.total_price)}</span>
                    <span className="text-neutral-400">{o.status}</span>
                  </div>
                  {selected?.type === 'order' && selected.item.id === o.id && (
                    <div className="mt-1 text-xs text-indigo-black font-medium flex items-center gap-1">
                      <CheckCircle size={12} /> 선택됨
                    </div>
                  )}
                </button>
              ))}
            </div>
          </div>
        )}

        {/* 검색 결과 0건 */}
        {!loading && data && !hasResults && (
          <div className="text-sm text-neutral-500 text-center py-4 bg-neutral-50 rounded-lg">
            같은 고객({data.name} / {data.phone})의 송장 보유 판매·주문건이 없습니다.
            <br />아래 직접 입력으로 진행해주세요.
          </div>
        )}

        {/* Fallback — 직접 송장번호 입력 */}
        <div className="border-t border-neutral-200 pt-4">
          <button
            onClick={() =>
              setSelected({ type: 'manual', invoice: manualInvoice, courier: manualCourier })
            }
            className={`w-full text-left p-3 rounded-lg border transition mb-2 ${
              selected?.type === 'manual'
                ? 'border-indigo-black bg-warm-ivory'
                : 'border-neutral-200 hover:border-neutral-400'
            }`}
          >
            <span className="text-sm font-bold">직접 송장번호 입력</span>
            <p className="text-xs text-neutral-500 mt-0.5">검색 결과에 없는 송장에 합쳐 발송한 경우</p>
          </button>
          {selected?.type === 'manual' && (
            <div className="space-y-2 px-1">
              <Input
                placeholder="송장번호 (필수)"
                value={manualInvoice}
                onChange={e => {
                  setManualInvoice(e.target.value);
                  setSelected({ type: 'manual', invoice: e.target.value, courier: manualCourier });
                }}
              />
              <Input
                placeholder="택배사 (기본: 롯데택배)"
                value={manualCourier}
                onChange={e => {
                  setManualCourier(e.target.value);
                  setSelected({ type: 'manual', invoice: manualInvoice, courier: e.target.value });
                }}
              />
            </div>
          )}
        </div>

        {/* 알림톡 발송 옵션 */}
        <label className="flex items-center gap-2 cursor-pointer p-3 bg-neutral-50 rounded-lg">
          <input
            type="checkbox"
            checked={!skipNotify}
            onChange={e => setSkipNotify(!e.target.checked)}
            className="w-4 h-4 rounded border-neutral-300"
          />
          <span className="text-sm">
            고객에게 출고 알림톡 발송 <span className="text-neutral-500">(수리내역 조회 + 배송조회 버튼 포함)</span>
          </span>
        </label>

        {/* 액션 */}
        <div className="flex justify-end gap-2 pt-2">
          <Button variant="ghost" onClick={onClose} disabled={submitting}>
            취소
          </Button>
          <Button
            variant="primary"
            onClick={handleSubmit}
            loading={submitting}
            disabled={!selected || submitting}
          >
            합포장 출고 처리
          </Button>
        </div>
      </div>
    </Modal>
  );
}
