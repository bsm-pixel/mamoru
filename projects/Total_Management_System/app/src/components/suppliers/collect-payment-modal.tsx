'use client';

import { useState, useEffect, useMemo } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { Check } from 'lucide-react';
import toast from 'react-hot-toast';
import { createClient } from '@/lib/supabase/client';

interface UnpaidItem {
  id: string;
  type: 'delivery' | 'sale'; // 납품 or 판매
  number: string; // DL-... or OS-...
  date: string;
  totalAmount: number;
  discountAmount: number;
  paidAmount: number;
  unpaidAmount: number;
  paymentStatus: string;
}

interface Props {
  open: boolean;
  customerId: string;
  customerName: string;
  onClose: () => void;
  onComplete: () => void;
}

export function CollectPaymentModal({ open, customerId, customerName, onClose, onComplete }: Props) {
  const queryClient = useQueryClient();
  const [items, setItems] = useState<UnpaidItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [checkedIds, setCheckedIds] = useState<Set<string>>(new Set());
  const [paymentMethod, setPaymentMethod] = useState('transfer');
  const [processing, setProcessing] = useState(false);
  const [progress, setProgress] = useState({ done: 0, total: 0 });

  // 미결제 건 조회
  useEffect(() => {
    if (!open || !customerId) return;
    setLoading(true);

    const supabase = createClient();

    Promise.all([
      // 납품 미결제 건
      (supabase as ReturnType<typeof createClient>)
        .from('deliveries')
        .select('id, dl_number, delivery_date, total_amount, discount_amount, paid_amount, payment_status')
        .eq('customer_id', customerId)
        .in('payment_status', ['unpaid', 'partial'])
        .is('cancelled_at', null)
        .order('delivery_date', { ascending: true }),
      // 판매 미결제 건
      (supabase as ReturnType<typeof createClient>)
        .from('offline_sales')
        .select('id, sale_number, sale_date, total_amount, discount_amount, paid_amount, payment_status')
        .eq('customer_id', customerId)
        .in('payment_status', ['unpaid', 'partial'])
        .is('cancelled_at', null)
        .order('sale_date', { ascending: true }),
    ]).then(([dlRes, saleRes]) => {
      const unpaid: UnpaidItem[] = [];

      for (const d of (dlRes.data || []) as Record<string, unknown>[]) {
        const total = (d.total_amount as number) || 0;
        const discount = (d.discount_amount as number) || 0;
        const paid = (d.paid_amount as number) || 0;
        unpaid.push({
          id: d.id as string,
          type: 'delivery',
          number: d.dl_number as string,
          date: d.delivery_date as string,
          totalAmount: total,
          discountAmount: discount,
          paidAmount: paid,
          unpaidAmount: total - discount - paid,
          paymentStatus: d.payment_status as string,
        });
      }

      for (const s of (saleRes.data || []) as Record<string, unknown>[]) {
        const total = (s.total_amount as number) || 0;
        const discount = (s.discount_amount as number) || 0;
        const paid = (s.paid_amount as number) || 0;
        unpaid.push({
          id: s.id as string,
          type: 'sale',
          number: s.sale_number as string,
          date: s.sale_date as string,
          totalAmount: total,
          discountAmount: discount,
          paidAmount: paid,
          unpaidAmount: total - discount - paid,
          paymentStatus: s.payment_status as string,
        });
      }

      setItems(unpaid);
      setCheckedIds(new Set(unpaid.map(i => i.id)));
      setLoading(false);
    });
  }, [open, customerId]);

  const selectedItems = useMemo(() => items.filter(i => checkedIds.has(i.id)), [items, checkedIds]);
  const selectedTotal = useMemo(() => selectedItems.reduce((s, i) => s + i.unpaidAmount, 0), [selectedItems]);
  const allChecked = items.length > 0 && checkedIds.size === items.length;

  const toggleAll = () => {
    if (allChecked) setCheckedIds(new Set());
    else setCheckedIds(new Set(items.map(i => i.id)));
  };

  const toggleOne = (id: string) => {
    setCheckedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  // 일괄 수금 처리
  const handleCollect = async () => {
    if (selectedItems.length === 0) return;
    setProcessing(true);
    setProgress({ done: 0, total: selectedItems.length });

    let success = 0;
    let fail = 0;

    for (const item of selectedItems) {
      try {
        const endpoint = item.type === 'delivery'
          ? `/api/deliveries/${item.id}`
          : `/api/sales/${item.id}`;

        const res = await fetch(endpoint, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'update_payment',
            payment_status: 'paid',
            payment_method: paymentMethod,
          }),
        });

        if (res.ok) success++;
        else fail++;
      } catch {
        fail++;
      }
      setProgress(prev => ({ ...prev, done: prev.done + 1 }));
    }

    setProcessing(false);

    if (fail === 0) {
      toast.success(`${success}건 수금 처리 완료`);
    } else {
      toast.error(`${success}건 성공, ${fail}건 실패`);
    }

    // 캐시 무효화 — 미수금/매출 대시보드 포함
    queryClient.invalidateQueries({ queryKey: ['customers'] });
    queryClient.invalidateQueries({ queryKey: ['deliveries'] });
    queryClient.invalidateQueries({ queryKey: ['delivery-stats'] });
    queryClient.invalidateQueries({ queryKey: ['sales'] });
    queryClient.invalidateQueries({ queryKey: ['sales-stats'] });
    queryClient.invalidateQueries({ queryKey: ['sales-tab-counts'] });
    queryClient.invalidateQueries({ queryKey: ['outstanding-alert'] });
    queryClient.invalidateQueries({ queryKey: ['dashboard-stats'] });

    onComplete();
    onClose();
  };

  const PAYMENT_METHODS = [
    { key: 'transfer', label: '계좌이체' },
    { key: 'card', label: '카드' },
    { key: 'cash', label: '현금' },
  ];

  const handleClose = () => {
    if (processing) return;
    onClose();
  };

  return (
    <Modal open={open} onClose={handleClose} title={`수금 처리 — ${customerName}`} className="max-w-lg">
      {loading ? (
        <p className="text-sm text-neutral-400 text-center py-8">미결제 내역 조회 중...</p>
      ) : items.length === 0 ? (
        <p className="text-sm text-neutral-400 text-center py-8">미결제 내역이 없습니다</p>
      ) : (
        <>
          {/* 전체 선택 */}
          <label className="flex items-center gap-2 px-3 py-2 rounded-lg bg-neutral-50 mb-3 cursor-pointer select-none">
            <input
              type="checkbox"
              checked={allChecked}
              onChange={toggleAll}
              className="w-4 h-4 rounded accent-neutral-900"
            />
            <span className="text-xs font-semibold text-neutral-600">
              전체 선택 ({items.length}건)
            </span>
          </label>

          {/* 미결제 내역 */}
          <div className="max-h-[50vh] overflow-y-auto space-y-1">
            {/* 납품 건 */}
            {items.filter(i => i.type === 'delivery').length > 0 && (
              <p className="text-[11px] font-semibold text-neutral-400 px-1 pt-1">B2B 납품</p>
            )}
            {items.filter(i => i.type === 'delivery').map(item => (
              <ItemRow key={item.id} item={item} checked={checkedIds.has(item.id)} onToggle={() => toggleOne(item.id)} />
            ))}

            {/* 판매 건 */}
            {items.filter(i => i.type === 'sale').length > 0 && (
              <p className="text-[11px] font-semibold text-neutral-400 px-1 pt-2">판매</p>
            )}
            {items.filter(i => i.type === 'sale').map(item => (
              <ItemRow key={item.id} item={item} checked={checkedIds.has(item.id)} onToggle={() => toggleOne(item.id)} />
            ))}
          </div>

          {/* 하단: 합계 + 결제방식 + 버튼 */}
          <div className="mt-4 pt-3 border-t border-neutral-200 space-y-3">
            <div className="flex items-center justify-between">
              <span className="text-sm text-neutral-500">선택 {selectedItems.length}건 합계</span>
              <span className="text-lg font-bold text-red-600">{formatKRW(selectedTotal)}</span>
            </div>

            <div className="flex gap-1.5">
              {PAYMENT_METHODS.map(m => (
                <button
                  key={m.key}
                  onClick={() => setPaymentMethod(m.key)}
                  className={`flex-1 py-1.5 text-xs rounded-md border transition ${
                    paymentMethod === m.key
                      ? 'bg-neutral-900 text-white border-neutral-900'
                      : 'bg-white text-neutral-500 border-neutral-200'
                  }`}
                >
                  {m.label}
                </button>
              ))}
            </div>

            <div className="flex gap-2">
              <Button variant="ghost" size="sm" onClick={handleClose} disabled={processing} className="flex-1">
                취소
              </Button>
              <Button size="sm" onClick={handleCollect} disabled={processing || selectedItems.length === 0} className="flex-1">
                {processing ? `${progress.done}/${progress.total} 처리 중...` : `수금 처리 (${formatKRW(selectedTotal)})`}
              </Button>
            </div>
          </div>
        </>
      )}
    </Modal>
  );
}

function ItemRow({ item, checked, onToggle }: { item: UnpaidItem; checked: boolean; onToggle: () => void }) {
  return (
    <label className={`flex items-center gap-3 px-3 py-2.5 rounded-lg border transition cursor-pointer select-none ${
      checked ? 'border-neutral-300 bg-neutral-50' : 'border-transparent hover:bg-neutral-50/50'
    }`}>
      <input
        type="checkbox"
        checked={checked}
        onChange={onToggle}
        className="w-4 h-4 rounded accent-neutral-900 shrink-0"
      />
      <div className="flex-1 min-w-0">
        <p className="text-sm font-medium text-neutral-800">{item.number}</p>
        <p className="text-xs text-neutral-400">{formatDate(item.date)}</p>
      </div>
      <Badge className={item.paymentStatus === 'partial' ? 'bg-yellow-100 text-yellow-700' : 'bg-red-100 text-red-700'}>
        {item.paymentStatus === 'partial' ? '부분결제' : '미결제'}
      </Badge>
      <span className="text-sm font-bold text-red-600 shrink-0 w-24 text-right">
        {formatKRW(item.unpaidAmount)}
      </span>
    </label>
  );
}
