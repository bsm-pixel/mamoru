'use client';

import { useState, useEffect, useMemo } from 'react';
import { createClient } from '@/lib/supabase/client';
import { useProducts } from '@/hooks/use-sales';
import { useExchangeOrder } from '@/hooks/use-orders';
import { formatKRW } from '@/lib/utils/format';
import { Search, RefreshCw, Store, Truck, Plus, X } from 'lucide-react';
import toast from 'react-hot-toast';
import type { Order, OrderItem } from '@/lib/supabase/types';

/**
 * 주문 교환 모달 (아임웹 온라인 주문) — 매출·카드 불변, 상품/재고만 스왑.
 * 반납 → 반품창고 / 새 제품(여러 개) → 시리얼 출고 / 차액 → cash_transaction.
 */
interface ProductLike { id: string; name: string; sku?: string | null; category?: string | null; price?: number | null; }
interface NewRow { key: number; productId: string | null; serialId: string | null; avail: { id: string; serial_number: string }[]; serialInput: string; creating: boolean; }

export function OrderExchangeModal({ order, items, onClose }: { order: Order; items: OrderItem[]; onClose: () => void; }) {
  const exchange = useExchangeOrder();
  const { data: products = [] } = useProducts();

  // 반납 대상(주문 품목 중 선택) — 기본 첫 품목
  const [returnItemIdx, setReturnItemIdx] = useState(0);
  const returnItem = items[returnItemIdx] || null;
  const [returnSerials, setReturnSerials] = useState<{ id: string; serial_number: string }[]>([]);

  // 회수 방식
  const [pickupMode, setPickupMode] = useState<'직접수거' | '방문수거' | '택배수거' | '고객반납'>('직접수거');

  // 새 제품 여러 줄
  const [rows, setRows] = useState<NewRow[]>([{ key: 1, productId: null, serialId: null, avail: [], serialInput: '', creating: false }]);
  const [prodSearch, setProdSearch] = useState('');
  const [activeRow, setActiveRow] = useState(1);

  const [receivedInput, setReceivedInput] = useState('');
  const [diffMethod, setDiffMethod] = useState<'현금' | '카드' | '이체' | '없음'>('현금');

  // 반납 품목의 주문 배정 시리얼 로드(있으면 반품창고行에 포함)
  useEffect(() => {
    (async () => {
      const pid = returnItem?.product_id;
      if (!pid) { setReturnSerials([]); return; }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = createClient() as any;
      const { data } = await db.from('product_serials').select('id, serial_number')
        .eq('order_id', order.id).eq('product_id', pid).eq('status', 'sold');
      setReturnSerials((data || []) as { id: string; serial_number: string }[]);
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [returnItemIdx, order.id]);

  const orderPaid = order.paid_amount ?? order.total_price ?? 0;

  const productById = (id: string | null) => (products as ProductLike[]).find((p) => p.id === id) || null;
  const newTotal = rows.reduce((s, r) => s + (productById(r.productId)?.price || 0), 0);
  const diff = newTotal - orderPaid; // + 추가수령 / - 환불
  const defaultReceived = diff;

  const filteredProducts = useMemo(() => (products as ProductLike[]).filter((p) =>
    !prodSearch || p.name.toLowerCase().includes(prodSearch.toLowerCase()) || (p.sku || '').toLowerCase().includes(prodSearch.toLowerCase()),
  ).slice(0, 40), [products, prodSearch]);

  function setRow(key: number, patch: Partial<NewRow>) {
    setRows((rs) => rs.map((r) => (r.key === key ? { ...r, ...patch } : r)));
  }
  function addRow() {
    const key = Math.max(0, ...rows.map((r) => r.key)) + 1;
    setRows((rs) => [...rs, { key, productId: null, serialId: null, avail: [], serialInput: '', creating: false }]);
    setActiveRow(key);
  }
  function removeRow(key: number) { setRows((rs) => (rs.length > 1 ? rs.filter((r) => r.key !== key) : rs)); }

  async function loadAvail(key: number, productId: string, selectNumber?: string) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createClient() as any;
    const { data } = await db.from('product_serials').select('id, serial_number')
      .eq('product_id', productId).eq('status', 'in_stock').order('serial_number');
    const list = (data || []) as { id: string; serial_number: string }[];
    setRow(key, { avail: list, serialId: selectNumber ? (list.find((s) => s.serial_number === selectNumber)?.id || null) : null });
  }

  async function pickProduct(key: number, p: ProductLike) {
    setRow(key, { productId: p.id, serialId: null, serialInput: '' });
    await loadAvail(key, p.id);
  }

  async function autoGen(key: number, productId: string) {
    setRow(key, { creating: true });
    try {
      const res = await fetch('/api/serials/batch', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ product_id: productId, count: 1, warehouse_zone: 'ready' }) });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || '시리얼 생성 실패');
      await loadAvail(key, productId, j.first_serial);
      toast.success(`시리얼 ${j.first_serial} 생성`);
    } catch (e) { toast.error(e instanceof Error ? e.message : '생성 실패'); }
    finally { setRow(key, { creating: false }); }
  }

  async function manualCreate(key: number, productId: string, num: string) {
    if (!num.trim()) return;
    setRow(key, { creating: true });
    try {
      const res = await fetch('/api/serials', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ product_id: productId, serial_number: num.trim() }) });
      if (!res.ok) { const t = await res.text(); throw new Error(/duplicate|unique|23505/i.test(t) ? `이미 존재: ${num}` : (t || '등록 실패')); }
      await loadAvail(key, productId, num.trim());
      setRow(key, { serialInput: '' });
      toast.success(`시리얼 ${num.trim()} 생성`);
    } catch (e) { toast.error(e instanceof Error ? e.message : '등록 실패'); }
    finally { setRow(key, { creating: false }); }
  }

  const validRows = rows.filter((r) => r.productId);
  // 시리얼 있는 제품인데 미배정이면 확정 불가(비시리얼은 avail 0이라 통과)
  const canConfirm = !!returnItem && validRows.length > 0
    && validRows.every((r) => r.avail.length === 0 || !!r.serialId)
    && !exchange.isPending;

  async function handleConfirm() {
    if (!returnItem) return;
    const returns = [{
      product_id: returnItem.product_id || '',
      product_name: returnItem.product_name || '',
      qty: returnItem.quantity || 1,
      serial_ids: returnSerials.map((s) => s.id),
    }];
    const new_items = validRows.map((r) => {
      const p = productById(r.productId);
      return { product_id: r.productId as string, product_name: p?.name, qty: 1, serial_ids: r.serialId ? [r.serialId] : [] };
    });
    const received = receivedInput !== '' ? parseInt(receivedInput) || 0 : defaultReceived;

    try {
      await exchange.mutateAsync({
        orderId: order.id,
        returns,
        new_items,
        recovery_method: pickupMode,
        diff_amount: received,
        diff_method: received === 0 ? '없음' : diffMethod,
      });
      toast.success('교환 처리 완료 — 반납품 반품창고, 새 제품 출고');
      onClose();
    } catch (e) { toast.error(e instanceof Error ? e.message : '교환 실패'); }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-2xl flex flex-col w-[580px] max-h-[92vh]" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800 flex items-center gap-1.5"><RefreshCw size={14} /> 주문 제품 교환</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600 text-lg leading-none">×</button>
        </div>

        <div className="overflow-y-auto flex-1 p-4 space-y-4">
          <p className="text-[11px] text-neutral-500 bg-amber-50 border border-amber-100 rounded-lg px-3 py-2">
            아임웹 주문·카드결제·매출은 <b>그대로</b> 두고 상품/재고만 바꿉니다(재결제 없음). 반납품은 <b>반품창고</b>로.
          </p>

          {/* 1) 반납 품목 */}
          <div>
            <p className="text-xs font-semibold text-neutral-500 mb-1.5">1. 반납받을 품목 (→ 반품창고)</p>
            <div className="space-y-1.5">
              {items.map((it, idx) => {
                const on = returnItemIdx === idx;
                return (
                  <button key={idx} onClick={() => setReturnItemIdx(idx)}
                    className={`w-full text-left px-3 py-2 rounded-lg border transition ${on ? 'border-purple-500 bg-purple-50' : 'border-neutral-200 hover:border-neutral-300'}`}>
                    <div className="flex items-center justify-between">
                      <span className="text-sm font-medium text-neutral-800">{it.product_name || '품목'}</span>
                      <span className="text-xs text-neutral-500">{it.quantity || 1}개</span>
                    </div>
                    {on && returnSerials.length > 0 && (
                      <div className="text-[11px] text-neutral-400 mt-0.5">시리얼 {returnSerials.map((s) => s.serial_number).join(', ')} → 반품창고</div>
                    )}
                  </button>
                );
              })}
            </div>
          </div>

          {/* 2) 회수 방식 */}
          <div>
            <p className="text-xs font-semibold text-neutral-500 mb-1.5">2. 구 제품 회수 방식</p>
            <div className="grid grid-cols-2 gap-1.5">
              {([
                { v: '직접수거', label: '내가 직접 수거함', icon: Store },
                { v: '고객반납', label: '고객 직접반납', icon: Store },
                { v: '방문수거', label: '방문수거', icon: Truck },
                { v: '택배수거', label: '택배수거', icon: Truck },
              ] as const).map((o) => {
                const on = pickupMode === o.v;
                return (
                  <button key={o.v} onClick={() => setPickupMode(o.v)}
                    className={`flex items-center gap-1.5 px-3 py-2 rounded-lg border text-xs font-medium transition ${on ? 'border-purple-500 bg-purple-50 text-purple-700' : 'border-neutral-200 text-neutral-600 hover:border-neutral-300'}`}>
                    <o.icon size={13} /> {o.label}
                  </button>
                );
              })}
            </div>
          </div>

          {/* 3) 새 제품 여러 개 */}
          <div>
            <div className="flex items-center justify-between mb-1.5">
              <p className="text-xs font-semibold text-neutral-500">3. 교환해줄 새 제품 (여러 개 가능 · 아임웹 없는 제품도 검색됨)</p>
              <button onClick={addRow} className="flex items-center gap-1 text-xs text-purple-600 font-semibold"><Plus size={13} /> 추가</button>
            </div>
            <div className="space-y-2">
              {rows.map((r) => {
                const p = productById(r.productId);
                return (
                  <div key={r.key} className="rounded-lg border border-neutral-200 p-2.5">
                    <div className="flex items-center justify-between mb-1.5">
                      <span className="text-xs font-medium text-neutral-700">{p ? <>{p.name} <span className="text-neutral-400">{formatKRW(p.price || 0)}</span></> : '제품 미선택'}</span>
                      {rows.length > 1 && <button onClick={() => removeRow(r.key)} className="text-neutral-300 hover:text-red-500"><X size={14} /></button>}
                    </div>
                    {activeRow === r.key || !r.productId ? (
                      <>
                        <div className="relative mb-1.5" onFocus={() => setActiveRow(r.key)}>
                          <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-neutral-400" />
                          <input value={activeRow === r.key ? prodSearch : ''} onChange={(e) => { setActiveRow(r.key); setProdSearch(e.target.value); }} placeholder="제품명·SKU 검색"
                            className="w-full h-8 pl-8 pr-2 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                        </div>
                        {activeRow === r.key && (
                          <div className="max-h-32 overflow-y-auto rounded-lg border border-neutral-100 divide-y divide-neutral-50 mb-1.5">
                            {filteredProducts.map((fp) => (
                              <button key={fp.id} onClick={() => { pickProduct(r.key, fp); setProdSearch(''); }}
                                className={`w-full flex items-center justify-between px-2.5 py-1.5 text-left text-sm ${r.productId === fp.id ? 'bg-blue-50' : 'hover:bg-neutral-50'}`}>
                                <span className="truncate">{fp.name} <span className="text-[11px] text-neutral-400">{fp.sku}</span></span>
                                <span className="text-xs text-neutral-500 shrink-0">{formatKRW(fp.price || 0)}</span>
                              </button>
                            ))}
                          </div>
                        )}
                      </>
                    ) : (
                      <button onClick={() => setActiveRow(r.key)} className="text-[11px] text-neutral-400 underline mb-1">제품 변경</button>
                    )}

                    {/* 시리얼 */}
                    {r.productId && (
                      <div className="mt-1">
                        {r.avail.length > 0 && (
                          <select value={r.serialId || ''} onChange={(e) => setRow(r.key, { serialId: e.target.value || null })}
                            className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-white text-sm mb-1.5">
                            <option value="">재고 시리얼 선택…</option>
                            {r.avail.map((s) => <option key={s.id} value={s.id}>{s.serial_number}</option>)}
                          </select>
                        )}
                        <div className="flex items-center gap-1.5">
                          <button type="button" onClick={() => autoGen(r.key, r.productId!)} disabled={r.creating}
                            className="shrink-0 flex items-center gap-1 px-2.5 h-8 rounded-lg bg-stone-900 text-white text-xs font-semibold disabled:opacity-50">
                            <RefreshCw size={11} className={r.creating ? 'animate-spin' : ''} /> 자동생성
                          </button>
                          <input value={r.serialInput} onChange={(e) => setRow(r.key, { serialInput: e.target.value })}
                            onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); manualCreate(r.key, r.productId!, r.serialInput); } }}
                            placeholder="직접입력 후 Enter"
                            className="flex-1 h-8 px-2 rounded-lg border border-neutral-200 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                        </div>
                        {r.serialId ? <p className="text-[11px] text-emerald-700 mt-1">✓ 배정: {r.avail.find((s) => s.id === r.serialId)?.serial_number}</p>
                          : r.avail.length > 0 ? <p className="text-[11px] text-amber-600 mt-1">시리얼을 선택/생성하세요</p>
                          : <p className="text-[11px] text-neutral-400 mt-1">시리얼 없는 제품이면 그대로 진행(보관재고에서 차감)</p>}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>

          {/* 4) 차액 */}
          <div className="rounded-lg bg-stone-50 p-3 space-y-2">
            <div className="flex items-center justify-between text-sm">
              <span className="text-neutral-500">새 제품 합계</span><span className="font-semibold">{formatKRW(newTotal)}</span>
            </div>
            <div className="flex items-center justify-between text-sm">
              <span className="text-neutral-500">주문 결제액</span><span>{formatKRW(orderPaid)}</span>
            </div>
            <div className="flex items-center justify-between text-sm pt-1 border-t border-neutral-200">
              <span className="font-semibold">차액</span>
              <span className={`font-bold ${diff > 0 ? 'text-rose-600' : diff < 0 ? 'text-blue-600' : 'text-neutral-500'}`}>
                {diff > 0 ? `추가 ${formatKRW(diff)} 받기` : diff < 0 ? `${formatKRW(-diff)} 환불` : '차액 없음'}
              </span>
            </div>
            <div className="flex items-center justify-between gap-2">
              <span className="text-xs text-neutral-500 shrink-0">실제 받은 차액</span>
              <input type="number" value={receivedInput} onChange={(e) => setReceivedInput(e.target.value)} placeholder={String(defaultReceived)}
                className="w-32 h-8 px-2 rounded-lg border border-neutral-200 text-sm text-right" />
            </div>
            <div className="flex items-center gap-1.5">
              {(['현금', '카드', '이체', '없음'] as const).map((m) => (
                <button key={m} onClick={() => setDiffMethod(m)}
                  className={`flex-1 py-1.5 rounded-lg border text-xs font-medium transition ${diffMethod === m ? 'border-stone-900 bg-stone-900 text-white' : 'border-neutral-200 text-neutral-600'}`}>{m}</button>
              ))}
            </div>
            <p className="text-[11px] text-neutral-400">비우면 기본값({formatKRW(defaultReceived)}) 적용 · 차액은 입출금에 기록됩니다(카드 재결제 아님).</p>
          </div>
        </div>

        <div className="flex items-center justify-end gap-2 px-4 py-3 border-t border-neutral-200">
          <button onClick={onClose} className="px-3 py-2 rounded-lg text-sm text-neutral-500 hover:bg-neutral-100">취소</button>
          <button onClick={handleConfirm} disabled={!canConfirm}
            className="px-4 py-2 rounded-lg bg-stone-900 text-white text-sm font-semibold hover:bg-stone-800 transition disabled:opacity-50">
            {exchange.isPending ? '교환 처리 중…' : '교환 확정'}
          </button>
        </div>
      </div>
    </div>
  );
}
