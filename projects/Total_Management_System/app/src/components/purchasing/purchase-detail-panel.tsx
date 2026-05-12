'use client';

import { useState } from 'react';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import { usePurchaseOrder, useUpdatePurchaseOrder, useDeletePurchaseOrder } from '@/hooks/use-purchasing';
import { formatKRW, formatDate, calcVAT } from '@/lib/utils/format';
import { useProducts } from '@/hooks/use-sales';
import { useQueryClient } from '@tanstack/react-query';
import { Truck, Pencil, Minus, Plus, Trash2, X, Save, Printer } from 'lucide-react';
import toast from 'react-hot-toast';
import { POPrintModal } from './po-print-modal';

const STATUS_LABEL: Record<string, string> = {
  draft: '작성중', ordered: '발주완료', deposit_paid: '선납완료',
  received: '입고완료', balance_paid: '잔금완료', cancelled: '취소',
};
const STATUS_COLOR: Record<string, string> = {
  draft: 'bg-neutral-100 text-neutral-600', ordered: 'bg-blue-100 text-blue-700',
  deposit_paid: 'bg-yellow-100 text-yellow-700', received: 'bg-green-100 text-green-700',
  balance_paid: 'bg-emerald-100 text-emerald-700', cancelled: 'bg-red-100 text-red-700',
};

interface Props {
  purchaseId: string;
}

export function PurchaseDetailPanel({ purchaseId }: Props) {
  const { data, isLoading } = usePurchaseOrder(purchaseId);
  const updatePO = useUpdatePurchaseOrder();
  const deletePO = useDeletePurchaseOrder();
  const queryClient = useQueryClient();
  const { data: products = [] } = useProducts();
  const [depositInput, setDepositInput] = useState('');
  const [editing, setEditing] = useState(false);
  const [editItems, setEditItems] = useState<Array<{ product_id?: string; product_name: string; sku?: string; quantity: number; unit_price: number }>>([]);
  const [editMemo, setEditMemo] = useState('');
  const [editExpectedDate, setEditExpectedDate] = useState('');
  const [saving, setSaving] = useState(false);
  const [showPrint, setShowPrint] = useState(false);
  // 입고검수 모달 — 품목별 실수령 수량
  const [receiveItems, setReceiveItems] = useState<Array<{ id: string; name: string; ordered: number; received: number; unit_price: number }>>([]);
  const [showReceive, setShowReceive] = useState(false);
  const [receiving, setReceiving] = useState(false);
  const [pendingAction, setPendingAction] = useState<{
    status: string; label: string; msg: string; variant?: 'danger' | 'default'; extra?: Record<string, unknown>;
  } | null>(null);

  if (isLoading) {
    return <div className="space-y-4"><Skeleton className="h-48" /><Skeleton className="h-32" /></div>;
  }
  if (!data?.order) {
    return (
      <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
        <Truck size={28} className="mb-2 opacity-40" />
        <p className="text-sm">발주 정보를 찾을 수 없습니다</p>
      </div>
    );
  }

  const { order: po, items } = data;
  const poVatType = ((po as Record<string, unknown>).vat_type as string) || 'included';
  const poCurrency = ((po as Record<string, unknown>).currency as string) || 'KRW';
  const poRate = ((po as Record<string, unknown>).exchange_rate as number) || 1;
  const isForeign = poCurrency !== 'KRW';
  const { supply, vat } = calcVAT(po.total_amount, poVatType as 'included' | 'separate' | 'none');
  const CURRENCY_SYMBOL: Record<string, string> = { KRW: '₩', USD: '$', CNY: '¥' };

  async function handleAction(status: string, extra?: Record<string, unknown>) {
    await updatePO.mutateAsync({ id: purchaseId, status, ...extra });
  }

  return (
    <div className="flex-1 overflow-y-auto space-y-4 pr-1">
      {/* 헤더 */}
      <div className="flex items-center justify-between">
        <div>
          <h3 className="text-sm font-bold text-indigo-black">{po.po_number}</h3>
          <p className="text-xs text-neutral-500 mt-0.5">
            {po.supplier_name} · {formatDate(po.order_date)}
            {poCurrency !== 'KRW' && <span className="ml-1 text-neutral-400">({poCurrency} 환율 {poRate})</span>}
          </p>
        </div>
        <div className="flex items-center gap-2">
          {po.status !== 'draft' && po.status !== 'cancelled' && (
            <Button variant="ghost" size="sm" onClick={() => setShowPrint(true)}>
              <Printer size={14} />발주서
            </Button>
          )}
          {(po.status === 'draft' || po.status === 'ordered' || po.status === 'deposit_paid') && !po.balance_paid_at && !editing && (
            <Button variant="ghost" size="sm" onClick={() => {
              setEditing(true);
              setEditItems(items.map((i) => ({ product_id: i.product_id || undefined, product_name: i.product_name, sku: i.sku || undefined, quantity: i.quantity, unit_price: i.unit_price })));
              setEditMemo(po.memo || '');
              setEditExpectedDate(po.expected_date || '');
            }}>
              <Pencil size={14} />
              편집
            </Button>
          )}
          <Badge className={STATUS_COLOR[po.status] || ''}>{STATUS_LABEL[po.status] || po.status}</Badge>
        </div>
      </div>

      {/* 발주 정보 */}
      <Card>
        <div className="grid grid-cols-2 gap-3 text-sm">
          <div>
            <span className="text-xs text-neutral-500">매입처</span>
            <p className="font-semibold">{po.supplier_name}</p>
          </div>
          <div>
            <span className="text-xs text-neutral-500">발주일</span>
            <p>{formatDate(po.order_date)}</p>
          </div>
          {po.expected_date && (
            <div>
              <span className="text-xs text-neutral-500">입고 예정일</span>
              <p>{formatDate(po.expected_date)}</p>
            </div>
          )}
          {po.received_date && (
            <div>
              <span className="text-xs text-neutral-500">입고일</span>
              <p>{formatDate(po.received_date)}</p>
            </div>
          )}
        </div>
        {po.memo && <p className="mt-3 pt-3 border-t border-neutral-100 text-sm text-neutral-600">{po.memo}</p>}
      </Card>

      {/* 품목 */}
      {editing ? (
        /* ── 편집 모드 ── */
        <Card>
          <div className="flex items-center justify-between mb-2">
            <h4 className="text-xs font-semibold text-neutral-500">품목 편집</h4>
            <div className="flex gap-1">
              <Button variant="ghost" size="sm" onClick={() => setEditing(false)}><X size={14} />취소</Button>
              <Button size="sm" loading={saving} onClick={async () => {
                if (editItems.length === 0) { toast.error('품목을 추가해주세요'); return; }
                setSaving(true);
                try {
                  await updatePO.mutateAsync({
                    id: purchaseId,
                    items: editItems,
                    memo: editMemo,
                    expected_date: editExpectedDate || undefined,
                  });
                  queryClient.invalidateQueries({ queryKey: ['purchase-order', purchaseId] });
                  setEditing(false);
                  toast.success('발주 수정 완료');
                } catch (err) {
                  toast.error(String(err));
                } finally { setSaving(false); }
              }}><Save size={14} />저장</Button>
            </div>
          </div>
          {po.status !== 'draft' && (
            <p className="text-[10px] text-amber-600 mb-2">선납/발주 후 수정 — 저장 시 총액·잔금이 자동 재계산됩니다 (선납이 새 총액보다 많으면 잔금 0·과지급). 입고 후엔 수정 불가.</p>
          )}

          {/* 메모 + 예정일 */}
          <div className="grid grid-cols-2 gap-2 mb-3">
            <div>
              <label className="text-xs text-neutral-500">입고 예정일</label>
              <input type="date" value={editExpectedDate} onChange={(e) => setEditExpectedDate(e.target.value)}
                className="w-full h-8 px-2 rounded border border-neutral-200 bg-warm-ivory text-xs focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
            <div>
              <label className="text-xs text-neutral-500">메모</label>
              <input type="text" value={editMemo} onChange={(e) => setEditMemo(e.target.value)} placeholder="메모"
                className="w-full h-8 px-2 rounded border border-neutral-200 bg-warm-ivory text-xs focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
            </div>
          </div>

          {/* 편집 품목 목록 */}
          <div className="space-y-2 mb-3">
            {editItems.map((item, idx) => (
              <div key={idx} className="flex items-center gap-2 py-1 border-b border-neutral-50 last:border-0">
                <div className="flex-1 min-w-0">
                  <p className="text-xs font-medium truncate">{item.product_name}</p>
                </div>
                <div className="flex items-center gap-1">
                  <button onClick={() => setEditItems(prev => prev.map((it, i) => i === idx ? { ...it, quantity: Math.max(1, it.quantity - 1) } : it))}
                    className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200"><Minus size={10} /></button>
                  <span className="text-xs font-semibold w-6 text-center">{item.quantity}</span>
                  <button onClick={() => setEditItems(prev => prev.map((it, i) => i === idx ? { ...it, quantity: it.quantity + 1 } : it))}
                    className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200"><Plus size={10} /></button>
                </div>
                <input type="number" value={item.unit_price || ''} onChange={(e) => setEditItems(prev => prev.map((it, i) => i === idx ? { ...it, unit_price: parseInt(e.target.value) || 0 } : it))}
                  className="w-20 h-7 px-2 rounded border border-neutral-200 bg-warm-ivory text-xs text-right" />
                <button onClick={() => setEditItems(prev => prev.filter((_, i) => i !== idx))}
                  className="w-6 h-6 rounded bg-red-50 flex items-center justify-center hover:bg-red-100 text-red-500"><Trash2 size={10} /></button>
              </div>
            ))}
          </div>

          {/* 제품 추가 — 검색 + 스크롤 리스트 */}
          <div>
            <p className="text-xs text-neutral-400 mb-1">제품 추가</p>
            <input
              type="text"
              placeholder="제품명 검색..."
              onChange={(e) => {
                const el = e.target.nextElementSibling;
                if (el) el.setAttribute('data-search', e.target.value.toLowerCase());
                // 강제 리렌더 위해 state 불필요 — DOM 직접 필터
                el?.querySelectorAll('[data-product-row]').forEach((row) => {
                  const name = row.getAttribute('data-product-name') || '';
                  (row as HTMLElement).style.display = name.includes(e.target.value.toLowerCase()) ? '' : 'none';
                });
              }}
              className="w-full h-8 px-3 mb-1.5 rounded-lg border border-neutral-200 text-xs placeholder:text-neutral-400 focus:outline-none focus:ring-1 focus:ring-neutral-300"
            />
            <div className="max-h-[200px] overflow-y-auto border border-neutral-100 rounded-lg divide-y divide-neutral-50">
              {products.filter(p => !editItems.find(ei => ei.product_id === p.id)).map(p => (
                <button key={p.id} data-product-row data-product-name={p.name.toLowerCase()}
                  onClick={() => setEditItems(prev => [...prev, {
                    product_id: p.id, product_name: p.name, sku: p.sku, quantity: 1, unit_price: p.price_purchase || p.price,
                  }])}
                  className="w-full flex items-center justify-between px-3 py-2 text-xs hover:bg-neutral-50 transition text-left">
                  <span className="truncate font-medium">{p.name}</span>
                  <span className="text-neutral-400 shrink-0 ml-2">{formatKRW(p.price_purchase || p.price)}</span>
                </button>
              ))}
            </div>
          </div>

          {/* 합계 */}
          <div className="mt-3 pt-3 border-t border-neutral-200 flex justify-between text-sm font-bold">
            <span>합계</span>
            <span className="text-terracotta">{formatKRW(editItems.reduce((s, i) => s + i.quantity * i.unit_price, 0))}</span>
          </div>
        </Card>
      ) : (
      <Card>
        <h4 className="text-xs font-semibold text-neutral-500 mb-2">발주 품목</h4>
        <div className="space-y-2">
          {items.map((item) => {
            const adjusted = item.received_quantity != null && item.received_quantity !== item.quantity;
            const lineKrw = Math.round((item.received_quantity ?? item.quantity) * item.unit_price * poRate);
            return (
            <div key={item.id} className="flex items-center justify-between py-1.5">
              <div>
                <p className="text-sm font-medium">{item.product_name}</p>
                <p className="text-xs text-neutral-500">
                  {item.sku && !item.sku.startsWith('IW-') && `${item.sku} · `}
                  {isForeign ? `${CURRENCY_SYMBOL[poCurrency]}${item.unit_price.toLocaleString()}` : formatKRW(item.unit_price)} ×{' '}
                  {adjusted ? (
                    <span className="text-orange-600 font-medium">입고 {item.received_quantity} / 주문 {item.quantity}</span>
                  ) : (
                    <>{item.quantity}{item.received_quantity != null ? ' (입고완료)' : ''}</>
                  )}
                </p>
              </div>
              <div className="text-right shrink-0">
                {isForeign && (
                  <span className="text-[10px] text-neutral-400 mr-1">{CURRENCY_SYMBOL[poCurrency]}{((item.received_quantity ?? item.quantity) * item.unit_price).toLocaleString()}</span>
                )}
                <span className={`text-sm font-bold ${adjusted ? 'text-orange-600' : ''}`}>{formatKRW(lineKrw)}</span>
              </div>
            </div>
          ); })}
        </div>
        <div className="mt-3 pt-3 border-t border-neutral-200 space-y-1">
          <div className="flex justify-between text-sm font-bold">
            <span>합계 {items.some((i) => i.received_quantity != null && i.received_quantity !== i.quantity) && <span className="text-[10px] font-normal text-orange-600">(입고 기준)</span>}</span>
            <span className="text-terracotta">{formatKRW(po.total_amount)}</span>
          </div>
          <div className="flex justify-between text-xs text-neutral-500">
            <span>공급가액 {formatKRW(supply)}</span>
            {poVatType !== 'none' && <span>부가세 {formatKRW(vat)}</span>}
            {poVatType === 'none' && <span className="text-neutral-400">부가세 미적용</span>}
          </div>
          {/* 부가세 유형 변경 (입고 전까지) */}
          {po.status !== 'received' && po.status !== 'balance_paid' && po.status !== 'cancelled' && (
            <div className="flex gap-1 mt-2">
              {([['included', '포함'], ['separate', '별도'], ['none', '미적용']] as const).map(([key, label]) => (
                <button key={key} onClick={() => {
                  if (key !== poVatType) updatePO.mutate({ id: purchaseId, vat_type: key });
                }}
                  className={`flex-1 py-1 text-[10px] rounded border transition ${poVatType === key ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-400 border-neutral-200 hover:border-neutral-300'}`}>
                  부가세 {label}
                </button>
              ))}
            </div>
          )}
        </div>
      </Card>
      )}

      {/* 결제 현황 */}
      <Card>
        <h4 className="text-xs font-semibold text-neutral-500 mb-2">결제 현황</h4>
        <div className="grid grid-cols-3 gap-3 text-center">
          <div>
            <p className="text-xs text-neutral-400">총액</p>
            <p className="text-sm font-bold">{formatKRW(po.total_amount)}</p>
          </div>
          <div>
            <p className="text-xs text-neutral-400">선납</p>
            <p className="text-sm font-bold text-blue-600">{formatKRW(po.deposit_amount)}</p>
          </div>
          <div>
            <p className="text-xs text-neutral-400">잔금</p>
            {po.balance_paid_at ? (
              <>
                <p className="text-sm font-bold text-green-600">{formatKRW(Math.max(0, po.total_amount - po.deposit_amount))}</p>
                <p className="text-[10px] text-green-600">지불완료 ✓ {formatDate(po.balance_paid_at)}</p>
              </>
            ) : (
              <p className={`text-sm font-bold ${po.balance_amount > 0 ? 'text-red-500' : 'text-green-600'}`}>
                {formatKRW(po.balance_amount)}
              </p>
            )}
          </div>
        </div>
      </Card>

      {/* 액션 */}
      {po.status !== 'cancelled' && po.status !== 'balance_paid' && (
        <Card>
          <h4 className="text-xs font-semibold text-neutral-500 mb-2">액션</h4>
          <div className="space-y-2">
            {po.status === 'draft' && (
              <Button className="w-full" onClick={() => setPendingAction({ status: 'ordered', label: '발주 확정', msg: '이 발주를 확정합니다.' })} disabled={updatePO.isPending}>
                발주 확정
              </Button>
            )}
            {(po.status === 'ordered' || po.status === 'draft') && (
              <div className="flex gap-2">
                <input type="number" value={depositInput} onChange={(e) => setDepositInput(e.target.value)}
                  placeholder={`선납금`}
                  className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
                <Button variant="secondary" onClick={() => setPendingAction({
                  status: 'deposit_paid', label: '선납 처리',
                  msg: `선납금 ${formatKRW(parseInt(depositInput) || Math.round(po.total_amount / 2))}을 처리합니다.`,
                  extra: { deposit_amount: parseInt(depositInput) || Math.round(po.total_amount / 2) },
                })} disabled={updatePO.isPending}>
                  선납
                </Button>
              </div>
            )}
            {(po.status === 'ordered' || po.status === 'deposit_paid') && (
              <Button className="w-full bg-green-600 hover:bg-green-700" onClick={() => {
                setReceiveItems(items.map((i) => ({ id: i.id, name: i.product_name, ordered: i.quantity, received: i.quantity, unit_price: i.unit_price })));
                setShowReceive(true);
              }} disabled={updatePO.isPending}>
                입고 확인 (재고 증가)
              </Button>
            )}
            {/* 입고 전 잔금 지불 기록 — 업체가 발송 시작했고 잔금만 먼저 보낸 경우 */}
            {(po.status === 'ordered' || po.status === 'deposit_paid') && po.balance_amount > 0 && !po.balance_paid_at && (
              <Button className="w-full" variant="secondary" onClick={() => setPendingAction({
                status: 'balance_paid', label: '잔금 지불 처리',
                msg: `잔금 ${formatKRW(po.balance_amount)} 을 지불완료로 기록합니다.\n(입고 상태는 그대로 — 물건 도착 후 "입고 확인"을 누르면 자동으로 잔금완료가 됩니다)`,
              })} disabled={updatePO.isPending}>
                잔금 지불 처리 (입고 전)
              </Button>
            )}
            {po.status === 'received' && po.balance_amount > 0 && !po.balance_paid_at && (
              <Button className="w-full" onClick={() => setPendingAction({ status: 'balance_paid', label: '잔금 완료', msg: `잔금 ${formatKRW(po.balance_amount)}을 완료 처리합니다.` })} disabled={updatePO.isPending}>
                잔금 완료
              </Button>
            )}
            <Button variant="ghost" className="w-full text-red-500 hover:text-red-600" onClick={() => setPendingAction({ status: 'cancelled', label: '발주 취소', msg: '이 발주를 취소합니다.', variant: 'danger' })} disabled={updatePO.isPending}>
              발주 취소
            </Button>
          </div>
        </Card>
      )}

      {/* 취소 건 삭제 */}
      {po.status === 'cancelled' && (
        <Card>
          <Button
            variant="ghost"
            className="w-full text-red-500 hover:text-red-600"
            onClick={() => {
              if (!confirm('이 발주를 영구 삭제합니다. 되돌릴 수 없습니다.')) return;
              deletePO.mutate(purchaseId);
            }}
            disabled={deletePO.isPending}
          >
            <Trash2 size={14} />
            {deletePO.isPending ? '삭제 중...' : '발주 삭제'}
          </Button>
        </Card>
      )}

      {/* 확인 모달 */}
      {pendingAction && (
        <ConfirmModal
          open={!!pendingAction}
          onClose={() => setPendingAction(null)}
          onConfirm={() => handleAction(pendingAction.status, pendingAction.extra)}
          title={pendingAction.label}
          message={<span className="whitespace-pre-wrap">{pendingAction.msg}</span>}
          confirmLabel={pendingAction.label}
          variant={pendingAction.variant || 'default'}
        />
      )}
      {showPrint && <POPrintModal purchaseId={purchaseId} onClose={() => setShowPrint(false)} />}

      {/* 입고검수 모달 — 품목별 실수령 수량 (제작품이라 주문≠입고 흔함) */}
      {showReceive && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4" onClick={() => { if (!receiving) setShowReceive(false); }}>
          <div className="bg-white rounded-xl shadow-xl w-full max-w-md max-h-[85vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
            <div className="p-4 border-b border-neutral-100">
              <h3 className="text-sm font-bold text-indigo-black">입고 검수</h3>
              <p className="text-xs text-neutral-400 mt-0.5">실제로 받은 수량을 확인하세요. 주문과 다른 품목만 고치면 됩니다.</p>
            </div>
            <div className="p-4 overflow-y-auto flex-1 space-y-2">
              {receiveItems.map((r, idx) => (
                <div key={r.id} className="flex items-center gap-2">
                  <span className="flex-1 text-sm truncate">{r.name}</span>
                  <span className="text-xs text-neutral-400 shrink-0">주문 {r.ordered}</span>
                  <input type="number" min={0} value={r.received}
                    onChange={(e) => { const v = Math.max(0, parseInt(e.target.value) || 0); setReceiveItems((prev) => prev.map((it, i) => i === idx ? { ...it, received: v } : it)); }}
                    className={`w-16 h-8 px-2 rounded border text-sm text-right focus:outline-none focus:ring-1 focus:ring-neutral-300 ${r.received !== r.ordered ? 'border-orange-400 bg-orange-50' : 'border-neutral-200'}`} />
                  <span className="text-[10px] text-neutral-400 w-6 shrink-0">자루</span>
                </div>
              ))}
            </div>
            {(() => {
              const newForeign = receiveItems.reduce((s, r) => s + r.received * r.unit_price, 0);
              const newKrw = Math.round(newForeign * poRate);
              const newTotal = poVatType === 'separate' ? newKrw + Math.round(newKrw * 0.1) : newKrw;
              const changed = receiveItems.some((r) => r.received !== r.ordered);
              const newBalance = po.balance_paid_at ? 0 : Math.max(0, newTotal - po.deposit_amount);
              const overpaid = !po.balance_paid_at && po.deposit_amount > newTotal ? po.deposit_amount - newTotal : 0;
              return (
                <div className="px-4 py-3 border-t border-neutral-100 text-xs space-y-1">
                  <div className="flex justify-between">
                    <span className="text-neutral-500">{changed ? '입고 기준 총액 (재계산)' : '총액'}</span>
                    <span className="font-semibold">{formatKRW(newTotal)}{changed && newTotal !== po.total_amount && <span className="text-orange-600 ml-1">({newTotal < po.total_amount ? '−' : '+'}{formatKRW(Math.abs(newTotal - po.total_amount))})</span>}</span>
                  </div>
                  <div className="flex justify-between"><span className="text-neutral-500">선납</span><span className="text-blue-600">{formatKRW(po.deposit_amount)}</span></div>
                  <div className="flex justify-between"><span className="text-neutral-500">잔금</span><span className={po.balance_paid_at ? 'text-green-600' : newBalance > 0 ? 'text-red-500' : 'text-green-600'}>{po.balance_paid_at ? '지불완료 ✓' : formatKRW(newBalance)}</span></div>
                  {overpaid > 0 && <p className="text-[11px] text-amber-600">⚠ 선납이 입고 기준 총액보다 {formatKRW(overpaid)} 많습니다 — 환불 또는 다음 발주 이월 검토</p>}
                  <p className="text-[10px] text-neutral-400 pt-1">확인 시 재고가 입고 수량만큼 늘어나고(아임웹 동기화 포함){po.balance_paid_at ? ', 잔금 이미 지불 → 잔금완료로' : ''} 입고완료 처리됩니다. 주문 수량 기록은 그대로 보존됩니다.</p>
                </div>
              );
            })()}
            <div className="p-4 border-t border-neutral-100 flex gap-2">
              <button onClick={() => { if (!receiving) setShowReceive(false); }} disabled={receiving}
                className="flex-1 py-2 rounded-lg bg-neutral-100 text-neutral-600 text-sm font-semibold hover:bg-neutral-200 disabled:opacity-50">취소</button>
              <button disabled={receiving} onClick={async () => {
                setReceiving(true);
                try {
                  await updatePO.mutateAsync({ id: purchaseId, status: 'received', received_items: receiveItems.map((r) => ({ id: r.id, received_quantity: r.received })) });
                  queryClient.invalidateQueries({ queryKey: ['purchase-order', purchaseId] });
                  setShowReceive(false);
                } catch (err) { toast.error(String(err)); }
                finally { setReceiving(false); }
              }} className="flex-1 py-2 rounded-lg bg-green-600 text-white text-sm font-semibold hover:bg-green-700 disabled:opacity-50">{receiving ? '처리 중...' : '입고 확인'}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
