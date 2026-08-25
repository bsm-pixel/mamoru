'use client';

import { useState, useEffect, useMemo } from 'react';
import { createClient } from '@/lib/supabase/client';
import { useProducts, useRebuildSale } from '@/hooks/use-sales';
import { useCreateReturn } from '@/hooks/use-returns';
import { formatKRW } from '@/lib/utils/format';
import { Search, RefreshCw, ArrowRight, Store, Truck } from 'lucide-react';
import toast from 'react-hot-toast';

/**
 * 교환 모달 (2026-08-25 · Phase 1 매장 직접 교환)
 * 구제품(시리얼 각인) 반납 → 반품창고(returned/return zone) 격리, 새 제품에 새 시리얼 배정.
 * 내부는 검증된 rebuild_sale 엔진 재사용 + exchange_returned_serial_ids로 반품분만 반품창고行.
 * ⚠️ 구 시리얼을 새 제품에 절대 덮어쓰지 않음(각각 자기 시리얼).
 */
interface SaleItem { id: string; product_id?: string | null; product_name: string; sku?: string | null; category?: string | null; quantity: number; unit_price: number; total_price: number; }
interface SerialRow { id: string; serial_number: string; product_id?: string | null; sale_item_id?: string | null; }
interface SaleLike { id: string; total_amount: number; discount_amount?: number | null; paid_amount?: number | null; payment_method?: string | null; sale_date?: string | null; memo?: string | null; sale_channel?: string | null; customer_name?: string | null; customer_id?: string | null; customer_phone?: string | null; shipped_at?: string | null; delivered_at?: string | null; invoice_number?: string | null; }
interface ProductLike { id: string; name: string; sku?: string | null; category?: string | null; price?: number | null; }

export function ExchangeModal({ sale, items, serials, onClose, onDone }: {
  sale: SaleLike; items: SaleItem[]; serials: SerialRow[]; onClose: () => void; onDone?: () => void;
}) {
  const rebuild = useRebuildSale();
  const createReturn = useCreateReturn();
  const { data: products = [] } = useProducts();
  // 구 제품 회수 방식 — 배송건은 수거 필요, 매장/직접건은 즉시 반납 기본
  const wasShipped = !!(sale.shipped_at || sale.delivered_at || sale.invoice_number);
  const [pickupMode, setPickupMode] = useState<'store' | '방문수거' | '택배수거' | '직접반납'>(wasShipped ? '택배수거' : 'store');
  const [pickupDate, setPickupDate] = useState<string>('');
  const [returnItemId, setReturnItemId] = useState<string | null>(null);
  const [prodSearch, setProdSearch] = useState('');
  const [newProductId, setNewProductId] = useState<string | null>(null);
  const [newSerialId, setNewSerialId] = useState<string | null>(null);
  const [availSerials, setAvailSerials] = useState<{ id: string; serial_number: string }[]>([]);
  const [receivedInput, setReceivedInput] = useState<string>('');
  const [serialInput, setSerialInput] = useState('');
  const [creating, setCreating] = useState(false);

  const itemSerials = (it: SaleItem) => serials.filter(
    (s) => (s.sale_item_id && s.sale_item_id === it.id) || (!s.sale_item_id && s.product_id && s.product_id === it.product_id),
  );
  // 반납 대상 = 전 품목 (시리얼 제품=반품창고 격리 / 비시리얼=판매가능 재고 복구)
  const returnableItems = items;
  const returnItem = items.find((i) => i.id === returnItemId) || null;
  const returnSerials = returnItem ? itemSerials(returnItem) : [];
  const returnSerialIds = returnSerials.map((s) => s.id);
  const newProduct = (products as ProductLike[]).find((p) => p.id === newProductId) || null;

  // 새 제품 in_stock 시리얼 로드 (+ 생성 후 특정 번호 자동 선택)
  const loadSerials = async (selectNumber?: string) => {
    if (!newProductId) { setAvailSerials([]); setNewSerialId(null); return; }
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createClient() as any;
    const { data } = await db.from('product_serials').select('id, serial_number').eq('product_id', newProductId).eq('status', 'in_stock').order('serial_number');
    const list: { id: string; serial_number: string }[] = data || [];
    setAvailSerials(list);
    if (selectNumber) {
      const hit = list.find((s) => s.serial_number === selectNumber);
      if (hit) setNewSerialId(hit.id);
    }
  };
  useEffect(() => {
    setNewSerialId(null); setSerialInput('');
    loadSerials();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [newProductId]);

  /** 자동생성 — 서버 SSOT 채번(MR{YY}{NNNNN}, 전역 MAX+1, DB UNIQUE 백스톱). 보관−1·준비+1 후 in_stock 시리얼을 배정 */
  async function handleAutoGen() {
    if (!newProductId || creating) return;
    setCreating(true);
    try {
      const res = await fetch('/api/serials/batch', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ product_id: newProductId, count: 1, warehouse_zone: 'ready' }) });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || '시리얼 생성 실패');
      await loadSerials(j.first_serial);
      toast.success(`시리얼 ${j.first_serial} 생성·배정`);
    } catch (e) { toast.error(e instanceof Error ? e.message : '시리얼 생성 실패'); }
    finally { setCreating(false); }
  }

  /** 직접입력 — 지정 번호로 생성(보관−1·준비+1). 중복은 서버/DB UNIQUE에서 차단 */
  async function handleManualCreate() {
    const num = serialInput.trim();
    if (!newProductId || !num || creating) return;
    setCreating(true);
    try {
      const res = await fetch('/api/serials', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ product_id: newProductId, serial_number: num }) });
      if (!res.ok) {
        const txt = await res.text();
        throw new Error(/duplicate|unique|23505/i.test(txt) ? `이미 존재하는 시리얼 번호입니다: ${num}` : (txt || '등록 실패'));
      }
      setSerialInput('');
      await loadSerials(num);
      toast.success(`시리얼 ${num} 생성·배정`);
    } catch (e) { toast.error(e instanceof Error ? e.message : '시리얼 등록 실패'); }
    finally { setCreating(false); }
  }

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

  const canConfirm = !!returnItem && !!newProduct && (availSerials.length === 0 || !!newSerialId) && !rebuild.isPending && !creating;

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

      // 배송 수거건이면 반품수거 추적 레코드 생성(매장/직접 즉시반납은 생성 안 함)
      if (pickupMode !== 'store') {
        try {
          await createReturn.mutateAsync({
            return_type: 'exchange',
            sale_id: sale.id,
            product_id: returnItem.product_id || null,
            product_name: returnItem.product_name,
            serial_id: returnSerials[0]?.id || null,
            serial_number: returnSerials[0]?.serial_number || null,
            qty: 1,
            customer_id: sale.customer_id || null,
            name: sale.customer_name || null,
            phone: sale.customer_phone || null,
            pickup_method: pickupMode,
            pickup_date: pickupMode === '방문수거' && pickupDate ? pickupDate : null,
            reason: '제품 교환',
          });
        } catch { toast('반품수거 접수 기록 생성 실패 — 반품관리에서 수동 등록하세요', { icon: '⚠️' }); }
      }

      toast.success(pickupMode === 'store'
        ? '교환 완료 — 구제품 반품창고, 새 제품 배정'
        : '교환 완료 — 새 제품 배정 + 반품수거 접수됨(반품관리에서 진행)');
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

          {/* 3) 새 시리얼 — 재고에서 선택 or 자동생성/직접입력(딜러납품 등 미리 마킹 안 한 경우) */}
          {newProduct && (
            <div>
              <p className="text-xs font-semibold text-neutral-500 mb-1.5">3. 새 제품 시리얼</p>
              {availSerials.length > 0 && (
                <select value={newSerialId || ''} onChange={(e) => setNewSerialId(e.target.value || null)}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-stone-400 mb-2">
                  <option value="">재고 시리얼 선택…</option>
                  {availSerials.map((s) => <option key={s.id} value={s.id}>{s.serial_number}</option>)}
                </select>
              )}
              {/* 새로 만들기 */}
              <div className="rounded-lg border border-dashed border-neutral-300 p-2.5 space-y-2 bg-neutral-50/50">
                <p className="text-[11px] text-neutral-500">새 시리얼 만들기 (미리 마킹 안 한 제품 — 보관재고에서 1개 꺼내 배정)</p>
                <div className="flex items-center gap-2">
                  <button type="button" onClick={handleAutoGen} disabled={creating}
                    className="shrink-0 flex items-center gap-1 px-3 h-9 rounded-lg bg-stone-900 text-white text-xs font-semibold hover:bg-stone-800 transition disabled:opacity-50">
                    <RefreshCw size={12} className={creating ? 'animate-spin' : ''} /> 자동생성
                  </button>
                  <input value={serialInput} onChange={(e) => setSerialInput(e.target.value)}
                    onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); handleManualCreate(); } }}
                    placeholder="직접입력 (예: MR2610900)"
                    className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                  <button type="button" onClick={handleManualCreate} disabled={creating || !serialInput.trim()}
                    className="shrink-0 px-3 h-9 rounded-lg border border-neutral-300 text-xs font-semibold text-neutral-700 hover:bg-neutral-100 transition disabled:opacity-50">등록</button>
                </div>
              </div>
              {newSerialId ? (
                <p className="text-[11px] text-emerald-700 mt-1.5">✓ 배정 시리얼: <b>{availSerials.find((s) => s.id === newSerialId)?.serial_number}</b></p>
              ) : (
                <p className="text-[11px] text-amber-600 mt-1.5">시리얼 미배정 상태로 확정 가능(비시리얼 제품이면 정상). 가위 등 시리얼 제품은 위에서 배정하세요.</p>
              )}
            </div>
          )}

          {/* 3.5) 구 제품 회수 방식 */}
          {returnItem && (
            <div>
              <p className="text-xs font-semibold text-neutral-500 mb-1.5">
                4. 구 제품 회수 {wasShipped && <span className="text-amber-600">(배송건 — 수거 필요)</span>}
              </p>
              <div className="grid grid-cols-2 gap-1.5">
                {([
                  { v: 'store', label: '매장/직접 지금 받음', icon: Store },
                  { v: '방문수거', label: '방문수거', icon: Truck },
                  { v: '택배수거', label: '택배수거', icon: Truck },
                  { v: '직접반납', label: '고객 직접반납(나중)', icon: Store },
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
              {pickupMode === '방문수거' && (
                <div className="flex items-center gap-2 mt-1.5">
                  <span className="text-[11px] text-neutral-500 shrink-0">수거 예약일</span>
                  <input type="date" value={pickupDate} onChange={(e) => setPickupDate(e.target.value)}
                    className="h-8 px-2 rounded-lg border border-neutral-200 text-sm" />
                </div>
              )}
              <p className="text-[11px] text-neutral-400 mt-1">
                {pickupMode === 'store' ? '구 제품을 지금 매장에서 받음 → 즉시 반품창고.' : '구 제품은 반품수거로 회수 → 반품관리에서 입고완료까지 추적.'}
              </p>
            </div>
          )}

          {/* 5) 차액·수납 */}
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
