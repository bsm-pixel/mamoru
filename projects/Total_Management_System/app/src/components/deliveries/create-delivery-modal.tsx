'use client';

/**
 * 거래처 납품서 작성 모달
 * 2026-05-26 Phase C: deliveries/page.tsx 안에서 분리 — /sales/new?mode=b2b 에서도 재사용
 * 2026-06-09: 제품 납품 + 복원수리 혼합 입력 통합 (mode 양자택일 제거)
 *   - 한 납품서에 제품 품목 + 복원수리(category='RS') 항목을 함께 담아 한 번에 저장
 *   - 복원수리는 VAT 제외(computeDeliveryTotals), 거래처 default_repair_price 자동 적용
 *   - 집계는 delivery_items.category='RS' 태그 기반으로 그대로 작동
 */

import { useState, useRef } from 'react';
import { Button } from '@/components/ui/button';
import { useCreateDelivery } from '@/hooks/use-deliveries';
import { useProducts } from '@/hooks/use-sales';
import { useCustomerSearch } from '@/hooks/use-customers';
import { useCustomerCatalog } from '@/hooks/use-customer-catalog';
import { formatKRW, formatPhone, toLocalDateString } from '@/lib/utils/format';
import { computeDeliveryTotals } from '@/lib/deliveries/totals';
import { X } from 'lucide-react';
import toast from 'react-hot-toast';
import type { Product } from '@/lib/supabase/types';

const PAYMENT_LABEL: Record<string, string> = { unpaid: '미결제', partial: '부분결제', paid: '결제완료' };
const RECEIPT_LABEL: Record<string, string> = { expense_proof: '지출증빙', tax_invoice: '세금계산서', none: '미적용' };

interface Props {
  /** @deprecated 2026-06-09 모드 통합 후 미사용 (호출부 호환 위해 유지) */
  initialMode?: 'delivery' | 'repair';
  onClose: () => void;
  onCreated: (id: string) => void;
}

export function CreateDeliveryModal({ onClose, onCreated }: Props) {
  const createDelivery = useCreateDelivery();
  const { data: products = [] } = useProducts();
  // 복원수리 (선택 입력 — 자루 0 이면 미포함)
  const [repairQty, setRepairQty] = useState(0);
  const [repairUnitPrice, setRepairUnitPrice] = useState(8000);
  const [repairExtras, setRepairExtras] = useState<Array<{ product_name: string; quantity: number; unit_price: number; category?: string }>>([]);
  const [repairExtraName, setRepairExtraName] = useState('');
  const [repairExtraPrice, setRepairExtraPrice] = useState('');
  // 고객 검색
  const [customerQuery, setCustomerQuery] = useState('');
  const { data: searchResults = [] } = useCustomerSearch(customerQuery);
  const [selectedCustomer, setSelectedCustomer] = useState<{
    id: string; name: string; phone: string | null; customer_type?: string; company_name?: string; default_repair_price?: number | null;
  } | null>(null);
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  const [showCustomerDropdown, setShowCustomerDropdown] = useState(false);

  // 제품
  const [productSearch, setProductSearch] = useState('');
  const [cart, setCart] = useState<Array<{
    product_id?: string; product_name: string; sku?: string; category?: string;
    quantity: number; unit_price: number;
  }>>([]);

  // 결제/옵션
  const [deliveryDate, setDeliveryDate] = useState(toLocalDateString(new Date()));
  const [expectedDate, setExpectedDate] = useState('');
  const [vatType, setVatType] = useState<'included' | 'separate' | 'none'>('none');
  const [receiptType, setReceiptType] = useState<string>('none');
  const [paymentStatus, setPaymentStatus] = useState<'unpaid' | 'partial' | 'paid'>('unpaid');
  const [paymentMethod, setPaymentMethod] = useState('transfer');
  const [discount, setDiscount] = useState(0);
  const [paidAmount, setPaidAmount] = useState(0);
  const [memo, setMemo] = useState('');
  const [customName, setCustomName] = useState('');
  const [customPrice, setCustomPrice] = useState('');

  const customerType = selectedCustomer?.customer_type;

  // 074: catalog 자동 입력 — 거래처별 납품명 + 납품가 우선 사용
  const { data: customerCatalogData } = useCustomerCatalog(selectedCustomer?.id);
  const catalogEntryMap = new Map(
    (customerCatalogData?.catalog || []).map((c) => [c.product_id, c])
  );

  // 제품 필터
  const filteredProducts = products.filter((p) => {
    if (!productSearch) return true;
    const q = productSearch.toLowerCase();
    return p.name.toLowerCase().includes(q) || (p.sku || '').toLowerCase().includes(q);
  });

  // 가격 결정: catalog.unit_price 우선 → 고객 유형별 → 기본가
  function getPrice(p: Product): number {
    const catalogEntry = catalogEntryMap.get(p.id);
    if (catalogEntry?.unit_price && catalogEntry.unit_price > 0) return catalogEntry.unit_price;
    if (customerType === 'dealer' && (p as Record<string, unknown>).price_dealer) return (p as Record<string, unknown>).price_dealer as number;
    if (customerType === 'academy' && (p as Record<string, unknown>).price_academy) return (p as Record<string, unknown>).price_academy as number;
    return p.price;
  }

  // 납품명 결정: catalog.delivery_name 우선 → product.name
  function getDeliveryName(p: Product): string {
    const catalogEntry = catalogEntryMap.get(p.id);
    return catalogEntry?.delivery_name?.trim() || p.name;
  }

  function addProduct(p: Product) {
    const existing = cart.find((c) => c.product_id === p.id);
    if (existing) {
      setCart((prev) => prev.map((c) => c.product_id === p.id ? { ...c, quantity: c.quantity + 1 } : c));
    } else {
      setCart((prev) => [...prev, {
        product_id: p.id,
        product_name: getDeliveryName(p),  // 074: catalog.delivery_name 우선
        sku: p.sku || undefined,
        category: p.category || undefined,
        quantity: 1,
        unit_price: getPrice(p),
      }]);
    }
  }

  // ── 통합 합계 (제품 + 복원수리 혼합, RS는 VAT 제외) ──
  const repairItems = [
    ...(repairQty > 0 ? [{ product_name: '복원수리', category: 'RS', quantity: repairQty, unit_price: repairUnitPrice }] : []),
    ...repairExtras.map((e) => ({ ...e, category: 'RS' })),
  ];
  const allItems = [...cart, ...repairItems];
  const productItemTotal = cart.reduce((s, c) => s + c.quantity * c.unit_price, 0);
  const repairItemTotal = repairItems.reduce((s, c) => s + c.quantity * c.unit_price, 0);
  // 제품이 없으면(복원수리 전용) VAT 미적용 — 기존 repair-only 동작과 동일
  const effectiveVatType = cart.length === 0 ? 'none' : vatType;
  const { supplyAmount: supply, vatAmount: vat, totalAmount } = computeDeliveryTotals(allItems, effectiveVatType, discount);

  async function handleSubmit() {
    const name = selectedCustomer?.company_name || selectedCustomer?.name || customerName.trim();
    if (!name) { toast.error('거래처를 입력해주세요'); return; }
    if (allItems.length === 0) { toast.error('제품 또는 복원수리를 입력해주세요'); return; }

    try {
      const result = await createDelivery.mutateAsync({
        customer_id: selectedCustomer?.id,
        customer_name: name,
        customer_phone: selectedCustomer?.phone || customerPhone.trim() || undefined,
        customer_type: customerType,
        delivery_date: deliveryDate,
        expected_date: expectedDate || undefined,
        memo: memo.trim() || undefined,
        vat_type: effectiveVatType,
        receipt_type: cart.length === 0 ? 'none' : receiptType,
        payment_status: paymentStatus,
        payment_method: paymentMethod,
        paid_amount: paymentStatus === 'partial' ? paidAmount : paymentStatus === 'paid' ? totalAmount : 0,
        discount_amount: discount,
        items: allItems,
      });
      onCreated((result.delivery as Record<string, unknown>).id as string);
    } catch {
      // error handled by hook
    }
  }

  // 드래그 중 모달 바깥으로 나가도 닫히지 않도록 mousedown 위치 체크
  const overlayRef = useRef<HTMLDivElement>(null);
  const mouseDownTarget = useRef<EventTarget | null>(null);

  return (
    <div ref={overlayRef} className="fixed inset-0 z-50 flex items-center justify-center bg-black/50"
      onMouseDown={(e) => { mouseDownTarget.current = e.target; }}
      onClick={(e) => { if (e.target === overlayRef.current && mouseDownTarget.current === overlayRef.current) onClose(); }}>
      <div
        className="bg-white rounded-xl shadow-2xl flex flex-col"
        style={{ width: '780px', maxHeight: '90vh' }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* 헤더 */}
        <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">거래처 매출 입력</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600"><X size={18} /></button>
        </div>

        {/* 본문 */}
        <div className="overflow-y-auto flex-1 p-5 space-y-4">
          {/* 거래처 선택 */}
          <div>
            <label className="text-xs font-semibold text-neutral-500 mb-1 block">거래처 (딜러/아카데미)</label>
            <div className="relative">
              <input
                type="text"
                value={selectedCustomer ? (selectedCustomer.company_name || selectedCustomer.name) : customerQuery}
                onChange={(e) => {
                  if (selectedCustomer) { setSelectedCustomer(null); setCustomerName(''); }
                  setCustomerQuery(e.target.value);
                  setShowCustomerDropdown(true);
                }}
                onFocus={() => customerQuery.length >= 2 && setShowCustomerDropdown(true)}
                placeholder="거래처 검색 (2자 이상)"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-neutral-300"
              />
              {selectedCustomer && (
                <button onClick={() => { setSelectedCustomer(null); setCustomerQuery(''); setCustomerName(''); }}
                  className="absolute right-2 top-1/2 -translate-y-1/2 text-neutral-400 hover:text-neutral-600">
                  <X size={14} />
                </button>
              )}
              {showCustomerDropdown && searchResults.length > 0 && !selectedCustomer && (
                <div className="absolute z-10 w-full mt-1 bg-white border border-neutral-200 rounded-lg shadow-lg max-h-48 overflow-y-auto">
                  {searchResults
                    .filter((c) => {
                      // 딜러/아카데미만 필터 (source로 판별)
                      const src = c.source;
                      const type = c.customer_type || '';
                      return type === 'dealer' || type === 'academy' || src === 'dealer' || src === 'academy';
                    })
                    .map((c) => (
                      <button
                        key={c.id}
                        onClick={() => {
                          const companyName = c.company_name || undefined;
                          setSelectedCustomer({
                            id: c.id,
                            name: c.name,
                            phone: c.phone,
                            customer_type: c.customer_type || '',
                            company_name: companyName || undefined,
                            default_repair_price: c.default_repair_price ?? null,
                          });
                          setCustomerName(companyName || c.name);
                          setCustomerPhone(c.phone || '');
                          setShowCustomerDropdown(false);
                          // 거래처 복원수리 기본 단가 자동 적용 (2-E)
                          if (c.default_repair_price && c.default_repair_price > 0) setRepairUnitPrice(c.default_repair_price);
                          // 장바구니 가격 재계산
                          const type = c.customer_type || '';
                          setCart((prev) => prev.map((item) => {
                            if (!item.product_id) return item;
                            const prod = products.find((p) => p.id === item.product_id);
                            if (!prod) return item;
                            let price = prod.price;
                            if (type === 'dealer' && (prod as Record<string, unknown>).price_dealer) price = (prod as Record<string, unknown>).price_dealer as number;
                            if (type === 'academy' && (prod as Record<string, unknown>).price_academy) price = (prod as Record<string, unknown>).price_academy as number;
                            return { ...item, unit_price: price };
                          }));
                        }}
                        className="w-full flex items-center justify-between px-3 py-2 text-sm hover:bg-neutral-50 transition text-left"
                      >
                        <div>
                          <span className="font-medium">{c.company_name || c.name}</span>
                          {c.company_name && <span className="text-xs text-neutral-400 ml-1">({c.name})</span>}
                          {!c.company_name && c.phone && <span className="text-xs text-neutral-400 ml-2">{formatPhone(c.phone)}</span>}
                        </div>
                        <span className={`text-[10px] px-1.5 py-0.5 rounded font-medium ${
                          c.customer_type === 'dealer' ? 'bg-blue-100 text-blue-700' : 'bg-purple-100 text-purple-700'
                        }`}>
                          {c.customer_type === 'dealer' ? '딜러' : '아카데미'}
                        </span>
                      </button>
                    ))}
                  {searchResults.filter((c) => {
                    const type = c.customer_type || '';
                    return type === 'dealer' || type === 'academy';
                  }).length === 0 && (
                    <div className="px-3 py-2 text-xs text-neutral-400">딜러/아카데미 고객이 없습니다</div>
                  )}
                </div>
              )}
            </div>
            {/* 직접 입력 (검색 결과 없을 때) */}
            {!selectedCustomer && (
              <div className="grid grid-cols-2 gap-2 mt-2">
                <input type="text" value={customerName} onChange={(e) => setCustomerName(e.target.value)}
                  placeholder="거래처명 직접 입력"
                  className="h-8 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
                <input type="text" value={customerPhone} onChange={(e) => setCustomerPhone(e.target.value)}
                  placeholder="연락처"
                  className="h-8 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
              </div>
            )}
            {customerType === 'dealer' && <p className="text-xs text-blue-600 mt-1">딜러가 적용</p>}
            {customerType === 'academy' && <p className="text-xs text-purple-600 mt-1">아카데미가 적용</p>}
          </div>

          {/* 날짜 */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs font-semibold text-neutral-500 mb-1 block">납품일</label>
              <input type="date" value={deliveryDate} onChange={(e) => setDeliveryDate(e.target.value)}
                className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-stone-50 text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
            </div>
            <div>
              <label className="text-xs font-semibold text-neutral-500 mb-1 block">납품 예정일</label>
              <input type="date" value={expectedDate} onChange={(e) => setExpectedDate(e.target.value)}
                className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-stone-50 text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
            </div>
          </div>

          {/* 제품 선택 */}
          <div>
            <label className="text-xs font-semibold text-neutral-500 mb-1 block">품목 (제품)</label>
            <div className="relative">
              <input
                type="text"
                value={productSearch}
                onChange={(e) => setProductSearch(e.target.value)}
                placeholder="제품명 또는 SKU 검색"
                className="w-full h-8 px-3 rounded-lg border border-neutral-200 text-xs placeholder:text-neutral-400 focus:outline-none focus:ring-1 focus:ring-neutral-300"
              />
              {productSearch && (
                <div className="absolute z-20 w-full mt-1 max-h-[180px] overflow-y-auto bg-white border border-neutral-200 rounded-lg shadow-lg divide-y divide-neutral-50">
                  {filteredProducts.map((p) => {
                    const inCart = cart.find((c) => c.product_id === p.id);
                    const price = getPrice(p);
                    return (
                      <button
                        key={p.id}
                        onClick={() => { addProduct(p); setProductSearch(''); }}
                        className={`w-full flex items-center justify-between px-3 py-2 text-xs hover:bg-neutral-50 transition text-left ${inCart ? 'bg-neutral-900/5' : ''}`}
                      >
                        <span className="truncate font-medium">
                          {p.name}
                          {inCart && <span className="ml-1.5 text-neutral-900 font-bold">x{inCart.quantity}</span>}
                        </span>
                        <span className="text-neutral-400 shrink-0 ml-2">{formatKRW(price)}</span>
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
            {/* 임시 제품 직접 입력 */}
            <div className="flex gap-2 mt-2">
              <input type="text" value={customName} onChange={(e) => setCustomName(e.target.value)}
                placeholder="품목명 직접 입력" className="flex-1 h-7 px-2 rounded border border-neutral-200 text-xs placeholder:text-neutral-400" />
              <input type="number" value={customPrice} onChange={(e) => setCustomPrice(e.target.value)}
                placeholder="금액" className="w-24 h-7 px-2 rounded border border-neutral-200 text-xs text-right placeholder:text-neutral-400" />
              <button onClick={() => {
                if (!customName.trim() || !parseInt(customPrice)) return;
                setCart((prev) => [...prev, { product_name: customName.trim(), quantity: 1, unit_price: parseInt(customPrice) || 0 }]);
                setCustomName(''); setCustomPrice('');
              }} className="h-7 px-3 rounded bg-neutral-900 text-white text-[10px] font-semibold shrink-0">추가</button>
            </div>
            <button onClick={() => {
              if (cart.some((c) => c.product_name === '배송비')) return;
              setCart((prev) => [...prev, { product_name: '배송비', quantity: 1, unit_price: 3000 }]);
            }} className="mt-1 h-7 px-3 rounded border border-neutral-200 text-[10px] font-medium text-neutral-500 hover:bg-neutral-50 transition">
              + 배송비 3,000원
            </button>

            {/* 장바구니 */}
            {cart.length > 0 && (
              <div className="mt-2 space-y-1.5">
                {cart.map((item, idx) => (
                  <div key={idx} className="flex items-center gap-2 py-1 border-b border-neutral-50 last:border-0">
                    <div className="flex-1 min-w-0">
                      <p className="text-xs font-medium truncate">{item.product_name}</p>
                    </div>
                    <div className="flex items-center gap-1">
                      <button onClick={() => setCart((prev) => prev.map((c, i) => i === idx ? { ...c, quantity: Math.max(1, c.quantity - 1) } : c))}
                        className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200 text-xs">-</button>
                      <input
                        type="number"
                        min={1}
                        value={item.quantity}
                        onChange={(e) => setCart((prev) => prev.map((c, i) => i === idx ? { ...c, quantity: Math.max(1, parseInt(e.target.value) || 1) } : c))}
                        className="w-10 h-6 text-center text-xs font-bold border border-neutral-200 rounded bg-white focus:outline-none [appearance:textfield] [&::-webkit-inner-spin-button]:appearance-none"
                      />
                      <button onClick={() => setCart((prev) => prev.map((c, i) => i === idx ? { ...c, quantity: c.quantity + 1 } : c))}
                        className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200 text-xs">+</button>
                    </div>
                    <input
                      type="number"
                      value={item.unit_price || ''}
                      onChange={(e) => setCart((prev) => prev.map((c, i) => i === idx ? { ...c, unit_price: parseInt(e.target.value) || 0 } : c))}
                      className="w-20 h-7 px-2 rounded border border-neutral-200 bg-stone-50 text-xs text-right"
                    />
                    <span className="text-xs text-neutral-500 w-20 text-right">{formatKRW(item.quantity * item.unit_price)}</span>
                    <button onClick={() => setCart((prev) => prev.filter((_, i) => i !== idx))}
                      className="w-6 h-6 rounded bg-red-50 flex items-center justify-center hover:bg-red-100 text-red-500 text-xs">x</button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* ═══ 복원수리 (선택) — 제품과 한 납품서에 함께 ═══ */}
          <div className="border border-neutral-200 rounded-lg p-3 space-y-2.5 bg-stone-50/40">
            <p className="text-xs font-bold text-neutral-700">복원수리 (선택)</p>
            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="text-[10px] text-neutral-500 block mb-0.5">자루 수</label>
                <input type="number" min={0} value={repairQty || ''} placeholder="0"
                  onChange={(e) => setRepairQty(Math.max(0, parseInt(e.target.value) || 0))}
                  className="w-full h-8 px-2 rounded border border-neutral-200 bg-white text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
              </div>
              <div>
                <label className="text-[10px] text-neutral-500 block mb-0.5">단가 (원)</label>
                <input type="number" value={repairUnitPrice}
                  onChange={(e) => setRepairUnitPrice(parseInt(e.target.value) || 0)}
                  className="w-full h-8 px-2 rounded border border-neutral-200 bg-white text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
              </div>
            </div>
            <p className="text-[10px] text-neutral-400">
              {selectedCustomer?.default_repair_price && selectedCustomer.default_repair_price > 0
                ? '거래처 기본 단가 자동 적용 · 수정 가능'
                : '기본 8,000원 · 수정 가능 (거래처 정보에 기본 단가 등록 시 자동 적용)'}
            </p>
            {/* 추가 항목 */}
            <div className="flex gap-2">
              <input type="text" value={repairExtraName} onChange={(e) => setRepairExtraName(e.target.value)}
                placeholder="추가 항목 (예: 급행료)" className="flex-1 h-7 px-2 rounded border border-neutral-200 text-xs placeholder:text-neutral-400" />
              <input type="number" value={repairExtraPrice} onChange={(e) => setRepairExtraPrice(e.target.value)}
                placeholder="금액" className="w-24 h-7 px-2 rounded border border-neutral-200 text-xs text-right placeholder:text-neutral-400" />
              <button onClick={() => {
                if (!repairExtraName.trim() || !parseInt(repairExtraPrice)) return;
                setRepairExtras((prev) => [...prev, { product_name: repairExtraName.trim(), quantity: 1, unit_price: parseInt(repairExtraPrice) || 0, category: 'RS' }]);
                setRepairExtraName(''); setRepairExtraPrice('');
              }} className="h-7 px-3 rounded bg-neutral-900 text-white text-[10px] font-semibold shrink-0">추가</button>
            </div>
            <button onClick={() => {
              if (repairExtras.some((e) => e.product_name === '배송비')) return;
              setRepairExtras((prev) => [...prev, { product_name: '배송비', quantity: 1, unit_price: 3000, category: 'RS' }]);
            }} className="h-7 px-3 rounded border border-neutral-200 text-[10px] font-medium text-neutral-500 hover:bg-neutral-50 transition">
              + 배송비 3,000원
            </button>
            {repairExtras.length > 0 && (
              <div className="space-y-1">
                {repairExtras.map((item, idx) => (
                  <div key={idx} className="flex items-center justify-between py-1 border-b border-neutral-100 last:border-0">
                    <span className="text-xs">{item.product_name}</span>
                    <div className="flex items-center gap-2">
                      <span className="text-xs text-neutral-500">{formatKRW(item.unit_price)}</span>
                      <button onClick={() => setRepairExtras((prev) => prev.filter((_, i) => i !== idx))}
                        className="w-5 h-5 rounded bg-red-50 flex items-center justify-center text-red-500 text-[10px]">x</button>
                    </div>
                  </div>
                ))}
              </div>
            )}
            {repairItemTotal > 0 && (
              <div className="flex justify-between text-xs font-semibold pt-1 border-t border-neutral-100">
                <span>복원수리 소계 ({repairQty}자루{repairExtras.length > 0 ? ' + 추가' : ''})</span>
                <span>{formatKRW(repairItemTotal)}</span>
              </div>
            )}
          </div>

          {/* VAT + 증빙 + 결제 — 그루핑 */}
          <div className="grid grid-cols-3 gap-3">
            <div className="border border-neutral-200 rounded-lg p-2.5">
              <label className="text-[10px] font-semibold text-neutral-400 mb-1.5 block">부가세</label>
              <div className="flex gap-1">
                {(['included', 'separate', 'none'] as const).map((v) => (
                  <button key={v} onClick={() => setVatType(v)}
                    className={`flex-1 py-1.5 rounded text-[10px] font-semibold transition ${
                      vatType === v ? 'bg-neutral-900 text-white' : 'bg-neutral-50 text-neutral-500 hover:bg-neutral-100'
                    }`}>
                    {{ included: '포함', separate: '별도', none: '미적용' }[v]}
                  </button>
                ))}
              </div>
            </div>
            <div className="border border-neutral-200 rounded-lg p-2.5">
              <label className="text-[10px] font-semibold text-neutral-400 mb-1.5 block">증빙유형</label>
              <div className="flex gap-1">
                {(['expense_proof', 'tax_invoice', 'none'] as const).map((r) => (
                  <button key={r} onClick={() => setReceiptType(r)}
                    className={`flex-1 py-1.5 rounded text-[10px] font-semibold transition ${
                      receiptType === r ? 'bg-neutral-900 text-white' : 'bg-neutral-50 text-neutral-500 hover:bg-neutral-100'
                    }`}>
                    {RECEIPT_LABEL[r]}
                  </button>
                ))}
              </div>
            </div>
            <div className="border border-neutral-200 rounded-lg p-2.5">
              <label className="text-[10px] font-semibold text-neutral-400 mb-1.5 block">결제상태</label>
              <div className="flex gap-1">
                {(['unpaid', 'partial', 'paid'] as const).map((ps) => (
                  <button key={ps} onClick={() => { setPaymentStatus(ps); if (ps !== 'partial') setPaidAmount(0); }}
                    className={`flex-1 py-1.5 rounded text-[10px] font-semibold transition ${
                      paymentStatus === ps
                        ? ps === 'paid' ? 'bg-green-600 text-white' : ps === 'partial' ? 'bg-yellow-500 text-white' : 'bg-red-500 text-white'
                        : 'bg-neutral-50 text-neutral-500 hover:bg-neutral-100'
                    }`}>
                    {PAYMENT_LABEL[ps]}
                  </button>
                ))}
              </div>
              {paymentStatus === 'partial' && (
                <input type="number" value={paidAmount || ''} onChange={(e) => setPaidAmount(parseInt(e.target.value) || 0)}
                  placeholder="선납금 입력"
                  className="w-full h-7 px-2 mt-2 rounded border border-neutral-200 bg-stone-50 text-xs focus:outline-none focus:ring-1 focus:ring-neutral-300" />
              )}
            </div>
          </div>

          {/* 결제수단 */}
          <div className="border border-neutral-200 rounded-lg p-2.5">
            <label className="text-[10px] font-semibold text-neutral-400 mb-1.5 block">결제수단</label>
            <div className="flex gap-1">
              {(['card', 'cash', 'transfer', 'mixed'] as const).map((m) => (
                <button key={m} onClick={() => setPaymentMethod(m)}
                  className={`flex-1 py-1.5 rounded text-xs font-semibold transition ${
                    paymentMethod === m ? 'bg-neutral-900 text-white' : 'bg-neutral-50 text-neutral-500 hover:bg-neutral-100'
                  }`}>
                  {{ card: '카드', cash: '현금', transfer: '이체', mixed: '복합' }[m]}
                </button>
              ))}
            </div>
          </div>

          {/* 할인 + 메모 */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs font-semibold text-neutral-500 mb-1 block">할인금액 (제품)</label>
              <input type="number" value={discount || ''} onChange={(e) => setDiscount(parseInt(e.target.value) || 0)}
                placeholder="0"
                className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-stone-50 text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
            </div>
            <div>
              <label className="text-xs font-semibold text-neutral-500 mb-1 block">메모</label>
              <input type="text" value={memo} onChange={(e) => setMemo(e.target.value)}
                placeholder="메모 (선택)"
                className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-stone-50 text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
            </div>
          </div>

          {/* 합계 */}
          {allItems.length > 0 && (
            <div className="bg-neutral-50 rounded-lg p-3 space-y-1">
              {cart.length > 0 && (
                <div className="flex justify-between text-xs text-neutral-500">
                  <span>제품 합계</span><span>{formatKRW(productItemTotal)}</span>
                </div>
              )}
              {repairItemTotal > 0 && (
                <div className="flex justify-between text-xs text-neutral-500">
                  <span>복원수리 (부가세 제외)</span><span>{formatKRW(repairItemTotal)}</span>
                </div>
              )}
              {discount > 0 && (
                <div className="flex justify-between text-xs text-red-500">
                  <span>할인</span><span>-{formatKRW(discount)}</span>
                </div>
              )}
              <div className="flex justify-between text-xs text-neutral-500">
                <span>공급가액</span><span>{formatKRW(supply)}</span>
              </div>
              {effectiveVatType !== 'none' && (
                <div className="flex justify-between text-xs text-neutral-500">
                  <span>부가세</span><span>{formatKRW(vat)}</span>
                </div>
              )}
              <div className="flex justify-between text-sm font-bold pt-1 border-t border-neutral-200">
                <span>합계</span><span>{formatKRW(totalAmount)}</span>
              </div>
            </div>
          )}
        </div>

        {/* 푸터 */}
        <div className="flex justify-end gap-2 px-5 py-3 border-t border-neutral-200">
          <Button variant="ghost" onClick={onClose}>취소</Button>
          <Button
            onClick={handleSubmit}
            disabled={allItems.length === 0 || (!selectedCustomer && !customerName.trim()) || createDelivery.isPending}
          >
            {createDelivery.isPending ? '생성 중...' : '납품서 생성'}
          </Button>
        </div>
      </div>
    </div>
  );
}
