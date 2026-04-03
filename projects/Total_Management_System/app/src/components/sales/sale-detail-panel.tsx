'use client';

import { useState, useRef } from 'react';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useSale, useCancelSale, useUpdatePaymentStatus, useUpdateSaleMemo, useEditSale, useRebuildSale, useProducts, useShipSale, useCancelSaleShipment } from '@/hooks/use-sales';
import { CustomerQuickModal } from '@/components/customers/customer-quick-modal';
import { formatKRW, formatDate, formatPhone } from '@/lib/utils/format';
import { Hash, Ban, CheckCircle, AlertTriangle, Pencil, Save, FileText, Printer, Download, Truck, Package } from 'lucide-react';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import type { SaleChannel, OfflineSale, OfflineSaleItem } from '@/lib/supabase/types';

const PAYMENT_METHOD_LABEL: Record<string, string> = {
  card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합',
};

const PAYMENT_STATUS_COLOR: Record<string, string> = {
  paid: 'bg-green-100 text-green-700',
  unpaid: 'bg-red-100 text-red-700',
  partial: 'bg-yellow-100 text-yellow-700',
};

const PAYMENT_STATUS_LABEL: Record<string, string> = {
  paid: '결제완료', unpaid: '미결제', partial: '부분결제',
};

const CHANNEL_CHIP: Record<string, { label: string; className: string }> = {
  offline: { label: '오프라인', className: 'bg-neutral-100 text-neutral-600' },
  online:  { label: '온라인',  className: 'bg-blue-100 text-blue-700' },
  talk:    { label: '톡상담',  className: 'bg-yellow-100 text-yellow-700' },
};

interface Props {
  saleId: string;
}

export function SaleDetailPanel({ saleId }: Props) {
  const { data, isLoading } = useSale(saleId);
  const cancelSale = useCancelSale();
  const updatePayment = useUpdatePaymentStatus();
  const updateMemo = useUpdateSaleMemo();
  const editSale = useEditSale();
  const rebuildSale = useRebuildSale();
  const shipSale = useShipSale();
  const cancelShipment = useCancelSaleShipment();
  const [cancelMode, setCancelMode] = useState(false);
  const [cancelReason, setCancelReason] = useState('');
  const [showPaidConfirm, setShowPaidConfirm] = useState(false);
  const [showCancelConfirm, setShowCancelConfirm] = useState(false);
  const [showEditModal, setShowEditModal] = useState(false);
  const [showReceipt, setShowReceipt] = useState(false);
  const [showCustomer, setShowCustomer] = useState(false);
  const [editingMemo, setEditingMemo] = useState(false);
  const [memoValue, setMemoValue] = useState('');

  if (isLoading) {
    return (
      <div className="p-4 space-y-3">
        <Skeleton className="h-20 w-full" />
        <Skeleton className="h-32 w-full" />
      </div>
    );
  }

  if (!data) {
    return <p className="text-sm text-neutral-400 text-center py-8">판매 정보를 찾을 수 없습니다</p>;
  }

  const { sale, items, serials = [] } = data;
  const s = sale as unknown as OfflineSale;
  const channel = CHANNEL_CHIP[(s.sale_channel || 'offline') as SaleChannel] || CHANNEL_CHIP.offline;

  const handleCancel = async () => {
    await cancelSale.mutateAsync({ id: saleId, reason: cancelReason });
    setCancelMode(false);
    setCancelReason('');
  };

  const handleMarkPaid = () => {
    updatePayment.mutate({
      id: saleId,
      payment_status: 'paid',
      paid_amount: s.total_amount,
    });
  };

  const handleSaveMemo = () => {
    updateMemo.mutate({ id: saleId, memo: memoValue });
    setEditingMemo(false);
  };

  return (
    <div className="p-4 space-y-4">
      {/* 헤더 */}
      <div>
        <div className="flex items-center gap-2 mb-2">
          <span className="text-sm font-bold text-indigo-black">{s.sale_number}</span>
          <Badge className={PAYMENT_STATUS_COLOR[s.payment_status] || ''}>
            {PAYMENT_STATUS_LABEL[s.payment_status] || s.payment_status}
          </Badge>
          <Badge className={channel.className}>{channel.label}</Badge>
        </div>
        <div className="grid grid-cols-2 gap-2 text-sm">
          <div>
            <span className="text-xs text-neutral-500">고객명</span>
            <p className="font-semibold">
              {s.customer_id ? (
                <button onClick={() => setShowCustomer(true)} className="text-blue-600 hover:underline">{s.customer_name}</button>
              ) : s.customer_name}
            </p>
          </div>
          <div>
            <span className="text-xs text-neutral-500">연락처</span>
            <p>{formatPhone(s.customer_phone) || '-'}</p>
          </div>
          <div>
            <span className="text-xs text-neutral-500">판매일</span>
            <p>{formatDate(s.sale_date, 'yyyy.MM.dd')}</p>
          </div>
          <div>
            <span className="text-xs text-neutral-500">결제방법</span>
            <p>{PAYMENT_METHOD_LABEL[s.payment_method] || s.payment_method}</p>
            {/* 복합 결제 상세 */}
            {s.payment_method === 'mixed' && (() => {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const pd = (s as any).payment_detail as Record<string, number> | null;
              if (!pd) return null;
              return (
                <div className="mt-1 text-xs text-neutral-500 space-y-0.5">
                  {pd.card > 0 && <p>카드 {formatKRW(pd.card)}</p>}
                  {pd.cash > 0 && <p>현금 {formatKRW(pd.cash)}</p>}
                  {pd.transfer > 0 && <p>이체 {formatKRW(pd.transfer)}</p>}
                </div>
              );
            })()}
          </div>
        </div>
        {/* 메모 인라인 편집 */}
        <div className="mt-2 pt-2 border-t border-neutral-100">
          {editingMemo ? (
            <div className="flex gap-2 items-center">
              <input
                type="text"
                value={memoValue}
                onChange={(e) => setMemoValue(e.target.value)}
                placeholder="메모 입력"
                className="flex-1 h-8 px-2 rounded border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                autoFocus
              />
              <button onClick={handleSaveMemo} className="p-1 hover:bg-neutral-100 rounded"><Save size={14} className="text-terracotta" /></button>
              <button onClick={() => setEditingMemo(false)} className="p-1 hover:bg-neutral-100 rounded text-xs text-neutral-500">취소</button>
            </div>
          ) : (
            <div className="flex items-center gap-1.5 text-sm text-neutral-600">
              <span>{s.memo || '메모 없음'}</span>
              <button onClick={() => { setMemoValue(s.memo || ''); setEditingMemo(true); }} className="p-0.5 hover:bg-neutral-100 rounded">
                <Pencil size={11} className="text-neutral-400" />
              </button>
            </div>
          )}
        </div>
      </div>

      {/* 판매 항목 */}
      <div className="border-t border-neutral-100 pt-3">
        <h4 className="text-xs font-semibold text-neutral-500 mb-2">판매 항목</h4>
        <div className="space-y-2">
          {(() => {
            // 시리얼-항목 매칭: sale_item_id 우선 → product_id → 순서 배분
            type Sr = { id: string; serial_number: string; product_id: string | null; sale_item_id?: string | null };
            const allSerials = serials as Sr[];
            const usedIds = new Set<string>();

            return (items as OfflineSaleItem[]).map((item) => {
              // 1순위: sale_item_id로 정확 매칭
              const bySaleItem = allSerials.filter((sr) => sr.sale_item_id === item.id && !usedIds.has(sr.id));
              bySaleItem.forEach((sr) => usedIds.add(sr.id));

              // 2순위: product_id 매칭 (sale_item_id 없는 경우)
              const byProduct = bySaleItem.length === 0
                ? allSerials.filter((sr) => sr.product_id && sr.product_id === item.product_id && !usedIds.has(sr.id))
                : [];
              byProduct.forEach((sr) => usedIds.add(sr.id));

              // 3순위: 나머지 unmatched 순서 배분
              const remaining = bySaleItem.length === 0 && byProduct.length === 0
                ? allSerials.filter((sr) => !sr.product_id && !sr.sale_item_id && !usedIds.has(sr.id)).slice(0, item.quantity)
                : [];
              remaining.forEach((sr) => usedIds.add(sr.id));

              const itemSerials = [...bySaleItem, ...byProduct, ...remaining];

              return (
                <div key={item.id} className="py-1.5">
                  <div className="flex items-center justify-between">
                    <div>
                      <p className="text-sm font-medium">{item.product_name}</p>
                      <p className="text-xs text-neutral-500">
                        {item.sku && `${item.sku} · `}{formatKRW(item.unit_price)} x {item.quantity}
                      </p>
                    </div>
                    <span className="text-sm font-bold">{formatKRW(item.total_price)}</span>
                  </div>
                  {itemSerials.length > 0 && (
                    <div className="mt-1 flex flex-wrap gap-1">
                      {itemSerials.map((sr) => (
                        <span key={sr.id} className="inline-flex items-center gap-0.5 text-[10px] font-mono bg-neutral-100 text-neutral-600 px-1.5 py-0.5 rounded">
                          <Hash size={8} />{sr.serial_number}
                        </span>
                      ))}
                    </div>
                  )}
                </div>
              );
            });
          })()}
        </div>
      </div>

      {/* 합계 */}
      <div className="border-t border-neutral-200 pt-3 space-y-1">
        <div className="flex justify-between text-sm">
          <span className="text-neutral-500">소계</span>
          <span>{formatKRW(s.total_amount)}</span>
        </div>
        {s.discount_amount > 0 && (
          <div className="flex justify-between text-sm">
            <span className="text-neutral-500">할인</span>
            <span className="text-red-600">-{formatKRW(s.discount_amount)}</span>
          </div>
        )}
        <div className="flex justify-between text-sm font-bold">
          <span>결제 금액</span>
          <span className="text-terracotta">{formatKRW(s.paid_amount)}</span>
        </div>
        {s.payment_status !== 'paid' && (
          <div className="flex justify-between text-sm font-semibold text-red-500">
            <span>미수금</span>
            <span>{formatKRW(s.total_amount - s.paid_amount)}</span>
          </div>
        )}
        {s.supply_amount > 0 && (
          <div className="mt-2 pt-2 border-t border-dashed border-neutral-200 space-y-0.5">
            <div className="flex justify-between text-xs text-neutral-500">
              <span>공급가액</span>
              <span>{formatKRW(s.supply_amount)}</span>
            </div>
            <div className="flex justify-between text-xs text-neutral-500">
              <span>부가세 (10%)</span>
              <span>{formatKRW(s.vat_amount)}</span>
            </div>
          </div>
        )}
      </div>

      {/* 거래명세서 + 수정 */}
      <div className="flex gap-2">
        <button
          onClick={() => setShowReceipt(true)}
          className="flex-1 flex items-center justify-center gap-2 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600 hover:bg-neutral-50 transition"
        >
          <FileText size={14} />
          거래명세서
        </button>
        {!s.cancelled_at && (
          <button
            onClick={() => setShowEditModal(true)}
            className="flex-1 flex items-center justify-center gap-2 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600 hover:bg-neutral-50 transition"
          >
            <Pencil size={14} />
            수정
          </button>
        )}
      </div>

      {/* 액션 */}
      {s.cancelled_at ? (
        <div className="flex items-center gap-2 px-3 py-2.5 rounded-lg bg-red-50 border border-red-100">
          <Ban size={14} className="text-red-500 shrink-0" />
          <div className="text-xs">
            <p className="font-semibold text-red-700">취소됨 — {formatDate(s.cancelled_at)}</p>
            {s.cancelled_reason && <p className="text-red-600 mt-0.5">{s.cancelled_reason}</p>}
          </div>
        </div>
      ) : (
        <div className="flex items-center gap-2 pt-2 border-t border-neutral-100">
          {s.payment_status !== 'paid' && (
            <Button size="sm" onClick={() => setShowPaidConfirm(true)} disabled={updatePayment.isPending}>
              <CheckCircle size={14} />
              {updatePayment.isPending ? '처리 중...' : '결제완료로 변경'}
            </Button>
          )}
          <button onClick={() => setShowCancelConfirm(true)} className="text-xs text-red-500 hover:text-red-700 transition">
            판매 취소
          </button>
        </div>
      )}

      {/* 결제완료 확인 모달 */}
      <ConfirmModal
        open={showPaidConfirm}
        onClose={() => setShowPaidConfirm(false)}
        onConfirm={() => updatePayment.mutateAsync({ id: saleId, payment_status: 'paid', paid_amount: s.total_amount })}
        title="결제완료 처리"
        message={<>{s.customer_name}님의 결제를 완료 처리합니다.<br />금액: {formatKRW(s.total_amount)}</>}
        confirmLabel="결제완료"
      />

      {/* 판매 취소 확인 모달 */}
      <ConfirmModal
        open={showCancelConfirm}
        onClose={() => { setShowCancelConfirm(false); setCancelReason(''); }}
        onConfirm={() => cancelSale.mutateAsync({ id: saleId, reason: cancelReason })}
        title="판매 취소"
        message={
          <div className="space-y-2">
            <div className="flex items-center gap-1.5 text-xs text-red-600">
              <AlertTriangle size={12} />
              <span>시리얼/재고가 복원됩니다. 이 작업은 되돌릴 수 없습니다.</span>
            </div>
            <input
              type="text"
              value={cancelReason}
              onChange={(e) => setCancelReason(e.target.value)}
              placeholder="취소 사유 (선택)"
              className="w-full h-8 px-3 rounded-lg border border-red-200 bg-red-50/50 text-sm focus:outline-none focus:ring-2 focus:ring-red-300"
            />
          </div>
        }
        confirmLabel="취소 확정"
        variant="danger"
      />

      {/* 택배 발송 */}
      {!s.cancelled_at && (
        <div className="pt-2 border-t border-neutral-100">
          {(s as Record<string, unknown>).invoice_number ? (
            <div className="flex items-center gap-2">
              <Package size={14} className="text-green-600" />
              <span className="text-sm font-mono font-medium">{(s as Record<string, unknown>).invoice_number as string}</span>
              <span className="text-xs text-neutral-400">롯데택배</span>
              <button onClick={() => cancelShipment.mutate(saleId)}
                disabled={cancelShipment.isPending}
                className="ml-auto text-xs text-red-400 hover:text-red-600">
                {cancelShipment.isPending ? '취소 중...' : '송장 취소'}
              </button>
            </div>
          ) : (
            <button
              onClick={() => shipSale.mutate(saleId)}
              disabled={shipSale.isPending}
              className="w-full flex items-center justify-center gap-2 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600 hover:bg-neutral-50 transition"
            >
              <Truck size={14} />
              {shipSale.isPending ? '송장 생성 중...' : '택배 발송 (송장 생성)'}
            </button>
          )}
        </div>
      )}

      {/* 고객 퀵뷰 모달 */}
      {s.customer_id && (
        <CustomerQuickModal customerId={s.customer_id} open={showCustomer} onClose={() => setShowCustomer(false)} />
      )}

      {/* 거래명세서 모달 */}
      {showReceipt && data && (
        <ReceiptModal
          sale={s}
          items={data.items}
          onClose={() => setShowReceipt(false)}
        />
      )}

      {/* 판매 수정 모달 */}
      {showEditModal && data && (
        <FullEditSaleModal
          sale={s}
          items={data.items}
          saleId={saleId}
          onClose={() => setShowEditModal(false)}
          rebuildSale={rebuildSale}
        />
      )}
    </div>
  );
}

/** 판매 전체 수정 모달 (제품 추가/삭제 + 금액/결제 수정) */
function FullEditSaleModal({ sale, items: originalItems, saleId, onClose, rebuildSale }: {
  sale: OfflineSale;
  items: OfflineSaleItem[];
  saleId: string;
  onClose: () => void;
  rebuildSale: ReturnType<typeof useRebuildSale>;
}) {
  const { data: products = [] } = useProducts();
  const [editItems, setEditItems] = useState(
    originalItems.map((it) => ({
      product_id: it.product_id || undefined,
      product_name: it.product_name,
      sku: it.sku || undefined,
      quantity: it.quantity,
      unit_price: it.unit_price,
      total_price: it.total_price,
      serial_ids: [] as string[], // 수정 시 시리얼은 새로 할당 불필요 (B2B)
    }))
  );
  const [paymentMethod, setPaymentMethod] = useState<string>(sale.payment_method);
  const [paymentStatus, setPaymentStatus] = useState<string>(sale.payment_status || 'paid');
  const [paidAmount, setPaidAmount] = useState(sale.paid_amount || 0);
  const [saleChannel, setSaleChannel] = useState<string>((sale as Record<string, unknown>).sale_channel as string || 'offline');
  const [discountAmount, setDiscountAmount] = useState(sale.discount_amount || 0);
  const [saleDate, setSaleDate] = useState(sale.sale_date);
  const [editMemo, setEditMemo] = useState(sale.memo || '');
  const [productSearch, setProductSearch] = useState('');

  const totalAmount = editItems.reduce((s, it) => s + it.unit_price * it.quantity, 0);
  const finalAmount = totalAmount - discountAmount;

  const removeItem = (idx: number) => setEditItems((prev) => prev.filter((_, i) => i !== idx));
  const updateQty = (idx: number, delta: number) => {
    setEditItems((prev) => prev.map((it, i) => i === idx ? { ...it, quantity: Math.max(1, it.quantity + delta), total_price: it.unit_price * Math.max(1, it.quantity + delta) } : it));
  };

  const addProduct = (p: { id: string; name: string; sku?: string; price: number }) => {
    const existing = editItems.findIndex((it) => it.product_id === p.id);
    if (existing >= 0) {
      updateQty(existing, 1);
    } else {
      setEditItems((prev) => [...prev, {
        product_id: p.id,
        product_name: p.name,
        sku: p.sku,
        quantity: 1,
        unit_price: p.price,
        total_price: p.price,
        serial_ids: [],
      }]);
    }
    setProductSearch('');
  };

  const filteredProducts = productSearch.length >= 1
    ? products.filter((p) => p.name.toLowerCase().includes(productSearch.toLowerCase()) || (p.sku || '').toLowerCase().includes(productSearch.toLowerCase())).slice(0, 8)
    : [];

  const handleSave = async () => {
    await rebuildSale.mutateAsync({
      id: saleId,
      items: editItems.map((it) => ({
        ...it,
        total_price: it.unit_price * it.quantity,
      })),
      sale_info: {
        total_amount: totalAmount,
        discount_amount: discountAmount,
        payment_method: paymentMethod,
        payment_status: paymentStatus,
        paid_amount: paymentStatus === 'paid' ? finalAmount : paymentStatus === 'unpaid' ? 0 : paidAmount,
        sale_date: saleDate,
        memo: editMemo.trim() || undefined,
        sale_channel: saleChannel,
      },
    });
    onClose();
  };

  const METHODS = [
    { value: 'card', label: '카드' },
    { value: 'cash', label: '현금' },
    { value: 'transfer', label: '이체' },
    { value: 'mixed', label: '복합' },
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-xl w-[95vw] max-w-[550px] max-h-[90vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
        <div className="px-5 py-3 border-b border-neutral-200 flex items-center justify-between">
          <h3 className="text-sm font-bold text-neutral-800">판매 수정</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600 text-lg">×</button>
        </div>

        <div className="flex-1 overflow-y-auto px-5 py-4 space-y-4">
          {/* 항목 목록 */}
          <div>
            <label className="text-xs font-semibold text-neutral-600 mb-2 block">품목 ({editItems.length})</label>
            <div className="space-y-2">
              {editItems.map((it, idx) => (
                <div key={idx} className="flex items-center gap-2 p-2 rounded-lg border border-neutral-200 bg-neutral-50">
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-medium truncate">{it.product_name}</p>
                    <p className="text-xs text-neutral-500">{formatKRW(it.unit_price)} × {it.quantity} = {formatKRW(it.unit_price * it.quantity)}</p>
                  </div>
                  <div className="flex items-center gap-1 shrink-0">
                    <button onClick={() => updateQty(idx, -1)} className="w-6 h-6 rounded bg-neutral-200 text-xs font-bold">−</button>
                    <input
                      type="number"
                      min="1"
                      value={it.quantity}
                      onChange={(e) => {
                        const val = parseInt(e.target.value) || 1;
                        setEditItems((prev) => prev.map((item, i) => i === idx ? { ...item, quantity: Math.max(1, val), total_price: item.unit_price * Math.max(1, val) } : item));
                      }}
                      className="w-10 h-6 text-center text-xs font-bold border border-neutral-200 rounded bg-white focus:outline-none [appearance:textfield] [&::-webkit-inner-spin-button]:appearance-none"
                    />
                    <button onClick={() => updateQty(idx, 1)} className="w-6 h-6 rounded bg-neutral-200 text-xs font-bold">+</button>
                    <button onClick={() => removeItem(idx)} className="w-6 h-6 rounded bg-red-100 text-red-500 text-xs font-bold ml-1">×</button>
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* 제품 추가 검색 */}
          <div>
            <label className="text-xs font-semibold text-neutral-600 mb-1 block">제품 추가</label>
            <input
              type="text"
              value={productSearch}
              onChange={(e) => setProductSearch(e.target.value)}
              placeholder="제품명 또는 SKU 검색"
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400"
            />
            {filteredProducts.length > 0 && (
              <div className="mt-1 border border-neutral-200 rounded-lg max-h-40 overflow-y-auto divide-y divide-neutral-100">
                {filteredProducts.map((p) => (
                  <button key={p.id} onClick={() => addProduct(p)}
                    className="w-full text-left px-3 py-2 hover:bg-neutral-50 transition">
                    <span className="text-sm font-medium">{p.name}</span>
                    <span className="text-xs text-neutral-500 ml-2">{p.sku}</span>
                    <span className="text-xs font-bold text-neutral-700 float-right">{formatKRW(p.price)}</span>
                  </button>
                ))}
              </div>
            )}
          </div>

          {/* 판매일 + 결제 */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs text-neutral-500 mb-1 block">판매일</label>
              <input type="date" value={saleDate} max={new Date().toISOString().slice(0, 10)}
                onChange={(e) => setSaleDate(e.target.value)}
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
            </div>
            <div>
              <label className="text-xs text-neutral-500 mb-1 block">할인</label>
              <input type="number" value={discountAmount}
                onChange={(e) => setDiscountAmount(Number(e.target.value))}
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
            </div>
          </div>

          <div>
            <label className="text-xs text-neutral-500 mb-1 block">결제방법</label>
            <div className="flex gap-2">
              {METHODS.map((m) => (
                <button key={m.value} onClick={() => setPaymentMethod(m.value)}
                  className={`flex-1 py-1.5 text-xs rounded-md border transition ${paymentMethod === m.value ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'}`}
                >{m.label}</button>
              ))}
            </div>
          </div>

          {/* 결제상태 */}
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">결제상태</label>
            <div className="flex gap-2">
              {[{ value: 'paid', label: '결제완료' }, { value: 'partial', label: '부분결제' }, { value: 'unpaid', label: '미결제' }].map((s) => (
                <button key={s.value} onClick={() => setPaymentStatus(s.value)}
                  className={`flex-1 py-1.5 text-xs rounded-md border transition ${paymentStatus === s.value ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'}`}
                >{s.label}</button>
              ))}
            </div>
            {paymentStatus === 'partial' && (
              <input type="number" value={paidAmount} onChange={(e) => setPaidAmount(Number(e.target.value))}
                placeholder="입금액" className="w-full h-9 px-3 mt-2 rounded-lg border border-neutral-200 text-sm" />
            )}
          </div>

          {/* 판매채널 */}
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">판매채널</label>
            <div className="flex gap-2">
              {[{ value: 'offline', label: '오프라인' }, { value: 'online', label: '온라인' }, { value: 'talk', label: '톡상담' }].map((c) => (
                <button key={c.value} onClick={() => setSaleChannel(c.value)}
                  className={`flex-1 py-1.5 text-xs rounded-md border transition ${saleChannel === c.value ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'}`}
                >{c.label}</button>
              ))}
            </div>
          </div>

          {/* 메모 */}
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">메모</label>
            <input type="text" value={editMemo} onChange={(e) => setEditMemo(e.target.value)}
              placeholder="메모" className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400" />
          </div>

          {/* 합계 */}
          <div className="bg-neutral-50 rounded-lg p-3 text-sm">
            <div className="flex justify-between"><span>합계</span><span className="font-bold">{formatKRW(totalAmount)}</span></div>
            {discountAmount > 0 && <div className="flex justify-between text-neutral-500"><span>할인</span><span>-{formatKRW(discountAmount)}</span></div>}
            <div className="flex justify-between font-bold text-base pt-1 border-t border-neutral-200 mt-1">
              <span>결제금액</span><span>{formatKRW(finalAmount)}</span>
            </div>
          </div>

          <p className="text-[11px] text-red-500">⚠ 저장 시 기존 시리얼/재고가 복원된 후 새 구성으로 재적용됩니다.</p>
        </div>

        <div className="px-5 py-3 border-t border-neutral-200 flex gap-2">
          <button onClick={onClose} className="flex-1 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600">취소</button>
          <button
            onClick={handleSave}
            disabled={rebuildSale.isPending || editItems.length === 0}
            className="flex-1 py-2 rounded-lg bg-neutral-900 text-white text-sm font-medium disabled:opacity-50"
          >
            {rebuildSale.isPending ? '저장 중...' : '저장'}
          </button>
        </div>
      </div>
    </div>
  );
}

/** 거래명세서 모달 (A4 비율, 인쇄 + 이미지 저장) */
function ReceiptModal({ sale, items, onClose }: {
  sale: OfflineSale;
  items: OfflineSaleItem[];
  onClose: () => void;
}) {
  const receiptRef = useRef<HTMLDivElement>(null);
  const [saving, setSaving] = useState(false);
  const PAYMENT_METHOD: Record<string, string> = { card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합' };

  const handlePrint = () => {
    const el = receiptRef.current;
    if (!el) return;
    const printWindow = window.open('', '_blank');
    if (!printWindow) return;
    printWindow.document.write(`
      <html><head><title>거래명세서 — ${sale.sale_number}</title>
      <style>
        body { font-family: 'Noto Sans KR', sans-serif; font-size: 12px; color: #000; padding: 20mm; }
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 6px 8px; text-align: left; }
        thead th { border-bottom: 2px solid #000; }
        tbody td { border-bottom: 1px solid #ddd; }
        .text-right { text-align: right; }
        .text-center { text-align: center; }
        .bold { font-weight: bold; }
        .total-row td { border-top: 2px solid #000; font-weight: bold; }
      </style></head><body>${el.innerHTML}</body></html>
    `);
    printWindow.document.close();
    printWindow.print();
  };

  const handleSaveImage = async () => {
    const el = receiptRef.current;
    if (!el) return;
    setSaving(true);
    try {
      const html2canvas = (await import('html2canvas')).default;
      const canvas = await html2canvas(el, { scale: 2, backgroundColor: '#ffffff' });
      const link = document.createElement('a');
      link.download = `거래명세서_${sale.sale_number}.png`;
      link.href = canvas.toDataURL('image/png');
      link.click();
    } catch (e) {
      console.error('이미지 저장 실패:', e);
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={onClose}>
      <div
        className="bg-white rounded-xl shadow-2xl flex flex-col"
        style={{ width: '595px', maxHeight: '90vh' }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* 상단 버튼 */}
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">거래명세서</h3>
          <div className="flex items-center gap-2">
            <button
              onClick={handleSaveImage}
              disabled={saving}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-neutral-100 text-xs text-neutral-700 hover:bg-neutral-200 transition disabled:opacity-50"
            >
              <Download size={12} />
              {saving ? '저장 중...' : '이미지 저장'}
            </button>
            <button
              onClick={handlePrint}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-neutral-900 text-xs text-white hover:bg-neutral-800 transition"
            >
              <Printer size={12} />
              인쇄
            </button>
            <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600 text-lg leading-none">×</button>
          </div>
        </div>

        {/* A4 비율 콘텐츠 */}
        <div className="overflow-y-auto flex-1 p-6" ref={receiptRef}>
          {/* 헤더 */}
          <h1 style={{ fontSize: '20px', fontWeight: 'bold', textAlign: 'center', marginBottom: '24px' }}>
            거 래 명 세 서
          </h1>

          <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '20px', fontSize: '13px' }}>
            <div>
              <p style={{ fontWeight: 'bold', fontSize: '15px' }}>MAMORU (마모루)</p>
              <p style={{ color: '#666' }}>미용가위 전문 브랜드</p>
              <p style={{ color: '#666' }}>서울특별시 구로구 부광로 88 SKV1, B동 311호</p>
              <p style={{ color: '#666' }}>TEL: 02-6326-0426</p>
            </div>
            <div style={{ textAlign: 'right' }}>
              <p>판매번호: {sale.sale_number}</p>
              <p>판매일: {sale.sale_date}</p>
              <p>발행일: {new Date().toISOString().slice(0, 10)}</p>
            </div>
          </div>

          {/* 고객 */}
          <div style={{ border: '1px solid #ddd', borderRadius: '4px', padding: '10px 14px', marginBottom: '16px', fontSize: '13px' }}>
            <span style={{ color: '#888' }}>고객명: </span>
            <strong>{sale.customer_name}</strong>
            {sale.customer_phone && (
              <span style={{ marginLeft: '24px' }}><span style={{ color: '#888' }}>연락처: </span>{sale.customer_phone}</span>
            )}
          </div>

          {/* 품목 */}
          <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '16px', fontSize: '13px' }}>
            <thead>
              <tr style={{ borderTop: '2px solid #000', borderBottom: '2px solid #000' }}>
                <th style={{ padding: '8px', textAlign: 'left' }}>품명</th>
                <th style={{ padding: '8px', textAlign: 'center', width: '60px' }}>수량</th>
                <th style={{ padding: '8px', textAlign: 'right', width: '100px' }}>단가</th>
                <th style={{ padding: '8px', textAlign: 'right', width: '120px' }}>금액</th>
              </tr>
            </thead>
            <tbody>
              {items.map((item, i) => (
                <tr key={i} style={{ borderBottom: '1px solid #eee' }}>
                  <td style={{ padding: '8px' }}>{item.product_name}{item.sku ? ` (${item.sku})` : ''}</td>
                  <td style={{ padding: '8px', textAlign: 'center' }}>{item.quantity}</td>
                  <td style={{ padding: '8px', textAlign: 'right' }}>{formatKRW(item.unit_price)}</td>
                  <td style={{ padding: '8px', textAlign: 'right', fontWeight: '600' }}>{formatKRW(item.total_price)}</td>
                </tr>
              ))}
            </tbody>
          </table>

          {/* 합계 */}
          <div style={{ borderTop: '2px solid #000', paddingTop: '12px', fontSize: '13px' }}>
            {(sale.discount_amount || 0) > 0 && (
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '4px' }}>
                <span>할인</span><span>-{formatKRW(sale.discount_amount)}</span>
              </div>
            )}
            {(sale.supply_amount || 0) > 0 && (
              <>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '4px', color: '#666' }}>
                  <span>공급가액</span><span>{formatKRW(sale.supply_amount)}</span>
                </div>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '4px', color: '#666' }}>
                  <span>부가세</span><span>{formatKRW(sale.vat_amount)}</span>
                </div>
              </>
            )}
            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '16px', fontWeight: 'bold', paddingTop: '8px', borderTop: '1px solid #ddd' }}>
              <span>합계</span>
              <span>{formatKRW(sale.total_amount - (sale.discount_amount || 0))}</span>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '4px', fontSize: '12px', color: '#888' }}>
              <span>결제방법</span>
              <span>{PAYMENT_METHOD[sale.payment_method] || sale.payment_method}</span>
            </div>
          </div>

          {/* 서명란 */}
          <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '60px', gap: '48px' }}>
            <div style={{ textAlign: 'center' }}>
              <div style={{ width: '100px', borderBottom: '1px solid #000', height: '40px' }}></div>
              <p style={{ fontSize: '11px', marginTop: '4px' }}>공급자</p>
            </div>
            <div style={{ textAlign: 'center' }}>
              <div style={{ width: '100px', borderBottom: '1px solid #000', height: '40px' }}></div>
              <p style={{ fontSize: '11px', marginTop: '4px' }}>공급받는자</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
