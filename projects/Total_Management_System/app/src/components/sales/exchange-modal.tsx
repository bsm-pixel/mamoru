'use client';

import { useState, useEffect, useMemo } from 'react';
import { createClient } from '@/lib/supabase/client';
import { useProducts, useRebuildSale } from '@/hooks/use-sales';
import { formatKRW } from '@/lib/utils/format';
import { Search, RefreshCw, ArrowRight } from 'lucide-react';
import toast from 'react-hot-toast';

/**
 * 교환 모달 (2026-08-25 · Phase 1 매장 직접 교환)
 * 구제품(시리얼 각인) 반납 → 반품창고(returned/return zone) 격리, 새 제품에 새 시리얼 배정.
 * 내부는 검증된 rebuild_sale 엔진 재사용 + exchange_returned_serial_ids로 반품분만 반품창고行.
 * ⚠️ 구 시리얼을 새 제품에 절대 덮어쓰지 않음(각각 자기 시리얼).
 */
interface SaleItem { id: string; product_id?: string | null; product_name: string; sku?: string | null; category?: string | null; quantity: number; unit_price: number; total_price: number; }
interface SerialRow { id: string; serial_number: string; product_id?: string | null; sale_item_id?: string | null; }
interface SaleLike { id: string; total_amount: number; discount_amount?: number | null; paid_amount?: number | null; payment_method?: string | null; sale_date?: string | null; memo?: string | null; sale_channel?: string | null; customer_name?: string | null; }
interface ProductLike { id: string; name: string; sku?: string | null; category?: string | null; price?: number | null; }

export function ExchangeModal({ sale, items, serials, onClose, onDone }: {
  sale: SaleLike; items: SaleItem[]; serials: SerialRow[]; onClose: () => void; onDone?: () => void;
}) {
  const rebuild = useRebuildSale();
  const { data: products = [] } = useProducts();
  const [returnItemId, setReturnItemId] = useState<string | null>(null);
  const [prodSearch, setProdSearch] = useState('');
  const [newProductId, setNewProductId] = useState<string | null>(null);
  const [newSerialId, setNewSerialId] = useState<string | null>(null);
  const [availSerials, setAvailSerials] = useState<{ id: string; serial_number: string }[]>([]);
  const [receivedInput, setReceivedInput] = useState<string>('');

  const itemSerials = (it: SaleItem) => serials.filter(
    (s) => (s.sale_item_id && s.sale_item_id === it.id) || (!s.sale_item_id && s.product_id && s.product_id === it.product_id),
  );
  // 반납 대상 = 전 품목 (시리얼 제품=반품창고 격리 / 비시리얼=판매가능 재고 복구)
  const returnableItems = items;
  const returnItem = items.find((i) => i.id === returnItemId) || null;
  const returnSerials = returnItem ? itemSerials(returnItem) : [];
  const returnSerialIds = returnSerials.map((s) => s.id);
  const newProduct = (products as ProductLike[]).find((p) => p.id === newProductId) || null;

  // 새 제품 in_stock 시리얼 로드
  useEffect(() => {
    if (!newProductId) { setAvailSerials([]); setNewSerialId(null); return; }
    let alive = true;
    (async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = createClient() as any;
      const { data } = await db.from('product_serials').select('id, serial_number').eq('product_id', newProductId).eq('status', 'in_stock').order('serial_number');
      if (alive) { setAvailSerials(data || []); setNewSerialId(null); }
    })();
    return () => { alive = false; };
  }, [newProductId]);

  const newUnitPrice = newProduct?.price || 0;
  const oldLineTotal = returnItem?.total_price || 0;
  const newLineTotal = newUnitPrice; // qty 1
  const diff = newLineTotal - oldLineTotal; // + 추가징수 / - 환불
  const newTotal = (sale.total_amount || 0) - oldLineTotal + newLineTotal;
  const defaultReceived = (sale.paid_amount || 0) + diff;

  const filteredProducts = useMemo(() => (products as ProductLike[]).filter((p) =>
    p.id !== returnItem?.product_id &&
    (!prodSearch || p.name.toLowerCase().includes(prodSearch.toLowerCase()) || (p.sku || '').toLowerCase().includes(prodSearch.toLowerCase())),
  ).slice(0, 40), [products, prodSearch, returnItem]);

  const canConfirm = !!returnItem && !!newProduct && (availSerials.length === 0 || !!newSerialId) && !rebuild.isPending;

  async function handleConfirm() {
    if (!returnItem || !newProduct) return;
    if (availSerials.length > 0 && !newSerialId) { toast.error('새 제품 시리얼을 선택하세요'); return; }

    // 유지 품목(반납품목 제외) — 기존 시리얼 재전송(release→re-assign)
    const keptItems = items.filter((i) => i.id !== returnItem.id).map((i) => ({
      product_id: i.product_id || undefined,
      product_name: i.product_name, sku: i.sku || undefined, category: i.category || undefined,
      quantity: i.quantity, unit_price: i.unit_price, total_price: i.total_price,
      serial_ids: itemSerials(i).map((s) => s.id),
    }));
    const newItem = {
      product_id: newProduct.id, product_name: newProduct.name, sku: newProduct.sku || undefined, category: newProduct.category || undefined,
      quantity: 1, unit_price: newUnitPrice, total_price: newLineTotal,
      serial_ids: newSerialId ? [newSerialId] : [],
    };
    const received = receivedInput !== '' ? (parseInt(receivedInput) || 0) : defaultReceived;
    const net = newTotal - (sale.discount_amount || 0);
    const paid = Math.max(0, Math.min(received, net));
    const payment_status = paid >= net ? 'paid' : paid <= 0 ? 'unpaid' : 'partial';

    try {
      await rebuild.mutateAsync({
        id: sale.id,
        items: [...keptItems, newItem],
        sale_info: {
          total_amount: newTotal,
          discount_amount: sale.discount_amount || 0,
          payment_method: sale.payment_method || 'cash',
          payment_status,
          paid_amount: paid,
          sale_date: sale.sale_date || undefined,
          memo: sale.memo || undefined,
          sale_channel: sale.sale_channel || undefined,
        },
        exchange_returned_serial_ids: returnSerialIds,
      });
      toast.success('교환 완료 — 구제품은 반품창고로, 새 제품 시리얼이 배정되었습니다');
      onDone?.();
      onClose();
    } catch (e) {
      toast.error(e instanceof Error ? e.message : '교환 실패');
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-2xl flex flex-col w-[560px] max-h-[90vh]" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800 flex items-center gap-1.5"><RefreshCw size={14} /> 제품 교환</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600 text-lg leading-none">×</button>
        </div>

        <div className="overflow-y-auto flex-1 p-4 space-y-4">
          {/* 1) 반납 품목 */}
          <div>
            <p className="text-xs font-semibold text-neutral-500 mb-1.5">1. 반납받을 품목 (구제품 → 반품창고)</p>
            {returnableItems.length === 0 ? (
              <p className="text-xs text-neutral-400 py-3 text-center bg-neutral-50 rounded-lg">교환할 품목이 없습니다.</p>
            ) : (
              <div className="space-y-1.5">
                {returnableItems.map((it) => {
                  const srs = itemSerials(it);
                  const on = returnItemId === it.id;
                  return (
                    <button key={it.id} onClick={() => setReturnItemId(it.id)}
                      className={`w-full text-left px-3 py-2 rounded-lg border transition ${on ? 'border-purple-500 bg-purple-50' : 'border-neutral-200 hover:border-neutral-300'}`}>
                      <div className="flex items-center justify-between">
                        <span className="text-sm font-medium text-neutral-800">{it.product_name}</span>
                        <span className="text-sm font-semibold text-neutral-700">{formatKRW(it.total_price)}</span>
                      </div>
                      <div className="text-[11px] text-neutral-400 mt-0.5">
                        {srs.length > 0 ? `시리얼 ${srs.map((s) => s.serial_number).join(', ')} → 반품창고` : '비시리얼 제품 → 판매가능 재고로 복구'}
                      </div>
                    </button>
                  );
                })}
              </div>
            )}
          </div>

          {/* 2) 새 제품 */}
          {returnItem && (
            <div>
              <p className="text-xs font-semibold text-neutral-500 mb-1.5">2. 교환해줄 새 제품</p>
              <div className="relative mb-1.5">
                <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
                <input value={prodSearch} onChange={(e) => setProdSearch(e.target.value)} placeholder="제품명·SKU 검색"
                  className="w-full h-9 pl-9 pr-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
              </div>
              <div className="max-h-40 overflow-y-auto rounded-lg border border-neutral-100 divide-y divide-neutral-50">
                {filteredProducts.map((p) => (
                  <button key={p.id} onClick={() => setNewProductId(p.id)}
                    className={`w-full flex items-center justify-between px-3 py-2 text-left transition ${newProductId === p.id ? 'bg-blue-50' : 'hover:bg-neutral-50'}`}>
                    <span className="text-sm truncate">{p.name} <span className="text-[11px] text-neutral-400">{p.sku}</span></span>
                    <span className="text-xs text-neutral-500 shrink-0">{formatKRW(p.price || 0)}</span>
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* 3) 새 시리얼 */}
          {newProduct && (
            <div>
              <p className="text-xs font-semibold text-neutral-500 mb-1.5">3. 새 제품 시리얼 (판매가능 재고에서 배정)</p>
              {availSerials.length === 0 ? (
                <p className="text-xs text-neutral-500 py-2 px-3 bg-neutral-50 rounded-lg">배정 가능한 재고 시리얼이 없습니다. <b>비시리얼 제품</b>이면 이대로 진행하세요(시리얼 없이 배정). 시리얼 제품인데 재고가 없다면 시리얼을 먼저 등록하세요.</p>
              ) : (
                <select value={newSerialId || ''} onChange={(e) => setNewSerialId(e.target.value || null)}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-stone-400">
                  <option value="">시리얼 선택…</option>
                  {availSerials.map((s) => <option key={s.id} value={s.id}>{s.serial_number}</option>)}
                </select>
              )}
            </div>
          )}

          {/* 4) 차액·수납 */}
          {returnItem && newProduct && (
            <div className="rounded-lg bg-stone-50 p-3 space-y-2">
              <div className="flex items-center justify-center gap-2 text-sm">
                <span className="text-neutral-500">{returnItem.product_name} {formatKRW(oldLineTotal)}</span>
                <ArrowRight size={14} className="text-neutral-400" />
                <span className="font-semibold text-neutral-800">{newProduct.name} {formatKRW(newLineTotal)}</span>
              </div>
              <div className="flex items-center justify-between text-sm pt-1 border-t border-neutral-200">
                <span className="font-semibold">차액</span>
                <span className={`font-bold ${diff > 0 ? 'text-rose-600' : diff < 0 ? 'text-blue-600' : 'text-neutral-500'}`}>
                  {diff > 0 ? `추가 ${formatKRW(diff)} 받기` : diff < 0 ? `${formatKRW(-diff)} 환불` : '차액 없음'}
                </span>
              </div>
              <div className="flex items-center justify-between gap-2">
                <span className="text-xs text-neutral-500 shrink-0">받은 금액(누적)</span>
                <input type="number" value={receivedInput} onChange={(e) => setReceivedInput(e.target.value)}
                  placeholder={String(defaultReceived)}
                  className="w-36 h-8 px-2 rounded-lg border border-neutral-200 text-sm text-right" />
              </div>
              <p className="text-[11px] text-neutral-400">비우면 기본값({formatKRW(defaultReceived)}) 적용. 새 판매 합계 {formatKRW(newTotal)}.</p>
            </div>
          )}
        </div>

        <div className="flex items-center justify-end gap-2 px-4 py-3 border-t border-neutral-200">
          <button onClick={onClose} className="px-3 py-2 rounded-lg text-sm text-neutral-500 hover:bg-neutral-100">취소</button>
          <button onClick={handleConfirm} disabled={!canConfirm}
            className="px-4 py-2 rounded-lg bg-stone-900 text-white text-sm font-semibold hover:bg-stone-800 transition disabled:opacity-50">
            {rebuild.isPending ? '교환 처리 중…' : '교환 확정'}
          </button>
        </div>
      </div>
    </div>
  );
}
