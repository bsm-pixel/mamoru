'use client';

import { useState, Suspense, useEffect } from 'react';
import { useRouter, useSearchParams } from 'next/navigation';
import Link from 'next/link';
import toast from 'react-hot-toast';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useProducts, useCreateSale } from '@/hooks/use-sales';
import { formatKRW, calcVAT, toLocalDateString } from '@/lib/utils/format';
import { Minus, Plus, ShoppingBag, Trash2 } from 'lucide-react';
import { CustomerAutocomplete, type SelectedCustomer } from '@/components/shared/customer-autocomplete';
import { CustomerCreateModal } from '@/components/customers/customer-create-modal';
import { SerialPicker } from '@/components/sales/serial-picker';
import { getUnitPrice, getProductDisplayName, hasGroupPrice } from '@/lib/utils/pricing';
import { usePriceGroups } from '@/hooks/use-price-groups';
import { useCustomerCatalog } from '@/hooks/use-customer-catalog';
import { TagSelector } from '@/components/shared/tag-selector';
import { useSetting } from '@/hooks/use-settings';
import type { Product } from '@/lib/supabase/types';

interface CartItem {
  product: Product | null; // null이면 임시 제품
  customName?: string;     // 임시 제품명
  customPrice?: number;    // 임시 제품 가격
  quantity: number;
  unitPrice: number;
  selectedSerialIds: string[];
  manualSerials?: string[]; // 직접 입력 시리얼 번호
}

/** cart item의 고유 키 */
function cartItemKey(item: CartItem): string {
  return item.product ? item.product.id : `custom-${item.customName}-${item.customPrice}`;
}

export default function NewSalePage() {
  return (
    <Suspense fallback={null}>
      <NewSaleContent />
    </Suspense>
  );
}

function NewSaleContent() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const contractId = searchParams?.get('contract_id') || null;
  const fromConsultationId = searchParams?.get('from_consultation') || null;
  const { data: products = [], isLoading: productsLoading } = useProducts();
  const createSale = useCreateSale();
  const priceGroups = usePriceGroups();
  const availableTags = useSetting<string[]>('customer.tags', []);

  const [saleDate, setSaleDate] = useState(toLocalDateString(new Date()));
  const [cart, setCart] = useState<CartItem[]>([]);
  const [selectedCustomer, setSelectedCustomer] = useState<SelectedCustomer | null>(null);

  // 073/074: B2B 납품처 catalog (customer별 납품명 + 단가 자동 입력용)
  const { data: customerCatalogData } = useCustomerCatalog(selectedCustomer?.id);
  const catalogEntryMap = new Map(
    (customerCatalogData?.catalog || []).map((c) => [c.product_id, c])
  );
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  const [paymentMethod, setPaymentMethod] = useState('card');
  const [paymentStatus, setPaymentStatus] = useState<'paid' | 'unpaid' | 'partial'>('paid');
  const [depositAmount, setDepositAmount] = useState(0);
  const [saleChannel, setSaleChannel] = useState('offline');
  const [discount, setDiscount] = useState(0);
  const [memo, setMemo] = useState('');

  // 복합 결제 분리 금액
  const [mixedCard, setMixedCard] = useState(0);
  const [mixedCash, setMixedCash] = useState(0);
  const [mixedTransfer, setMixedTransfer] = useState(0);

  const [showCreateCustomer, setShowCreateCustomer] = useState(false);
  const [sourceConsultation, setSourceConsultation] = useState<{ id: string; unique_id: string } | null>(null);
  const customerType = selectedCustomer?.customer_type;

  // 070: 출장/매장상담 → 판매 link prefill
  // 077: 자동매칭된 기존 고객(이름이 다를 때) 안내 토스트
  useEffect(() => {
    if (!fromConsultationId) return;
    let cancelled = false;
    (async () => {
      try {
        const res = await fetch(`/api/consultation/${fromConsultationId}`);
        if (!res.ok) return;
        const { consultation } = await res.json();
        if (cancelled || !consultation) return;
        setSourceConsultation({ id: consultation.id, unique_id: consultation.unique_id });
        setCustomerName(consultation.name || '');
        setCustomerPhone(consultation.phone || '');
        if (consultation.customer_id) {
          // 077: 실제 customer 레코드 fetch — phone_normalized로 자동 매칭된 기존 고객의 정확한 이름 사용
          let customerName: string = consultation.name;
          try {
            const custRes = await fetch(`/api/customers/${consultation.customer_id}`);
            if (custRes.ok) {
              const { customer } = await custRes.json();
              if (customer?.name) {
                customerName = customer.name;
                // 이름이 다르면 자동매칭 안내 (전화번호 기준 SSOT)
                if (customer.name !== consultation.name) {
                  toast(
                    `동일 전화번호로 자동매칭됨\n접수명: ${consultation.name} → 기존 고객: ${customer.name}`,
                    { duration: 5000, icon: 'ℹ️' }
                  );
                }
              }
            }
          } catch {
            /* customer fetch 실패해도 consultation 데이터로 진행 */
          }
          setSelectedCustomer({
            id: consultation.customer_id,
            name: customerName,
            phone: consultation.phone,
            email: null,
            address_road: consultation.address_road || null,
            address_detail: consultation.address_detail || null,
            postcode: consultation.postcode || null,
            ecount_customer_code: null,
          } as SelectedCustomer);
        }
      } catch {
        /* 조용히 실패, 수동 입력 가능 */
      }
    })();
    return () => { cancelled = true; };
  }, [fromConsultationId]);
  const totalAmount = cart.reduce((s, item) => s + item.unitPrice * item.quantity, 0);
  const finalAmount = totalAmount - discount;
  const paidAmount = paymentStatus === 'paid' ? finalAmount
    : paymentStatus === 'unpaid' ? 0
    : Math.min(depositAmount, finalAmount);

  // 임시 제품 입력 상태
  const [customProductName, setCustomProductName] = useState('');
  const [customProductPrice, setCustomProductPrice] = useState('');

  function addToCart(product: Product) {
    // 074: catalog.unit_price 우선 사용 (해당 customer 등록된 맞춤가) → 없으면 group 단가 fallback
    const catalogEntry = catalogEntryMap.get(product.id);
    const customPrice = catalogEntry?.unit_price;
    const price = (customPrice && customPrice > 0)
      ? customPrice
      : getUnitPrice(product, customerType, priceGroups);
    setCart((prev) => {
      const existing = prev.find((item) => item.product?.id === product.id);
      if (existing) {
        return prev.map((item) =>
          item.product?.id === product.id ? { ...item, quantity: item.quantity + 1 } : item
        );
      }
      return [...prev, { product, quantity: 1, unitPrice: price, selectedSerialIds: [] }];
    });
  }

  function addCustomProduct() {
    const name = customProductName.trim();
    const price = parseInt(customProductPrice) || 0;
    if (!name || price <= 0) return;
    setCart((prev) => [
      ...prev,
      { product: null, customName: name, customPrice: price, quantity: 1, unitPrice: price, selectedSerialIds: [] },
    ]);
    setCustomProductName('');
    setCustomProductPrice('');
  }

  // 고객 변경 시 장바구니 가격 재계산 (등록 제품만)
  function recalcCartPrices(type?: string) {
    setCart((prev) => prev.map((item) => ({
      ...item,
      unitPrice: item.product ? getUnitPrice(item.product, type, priceGroups) : item.unitPrice,
    })));
  }

  function updateSerialIds(key: string, serialIds: string[]) {
    setCart((prev) =>
      prev.map((item) =>
        cartItemKey(item) === key ? { ...item, selectedSerialIds: serialIds } : item
      )
    );
  }

  function updateManualSerials(key: string, serials: string[]) {
    setCart((prev) =>
      prev.map((item) =>
        cartItemKey(item) === key ? { ...item, manualSerials: serials } : item
      )
    );
  }

  function updateQuantity(key: string, delta: number) {
    setCart((prev) =>
      prev
        .map((item) =>
          cartItemKey(item) === key
            ? { ...item, quantity: Math.max(0, item.quantity + delta) }
            : item
        )
        .filter((item) => item.quantity > 0)
    );
  }

  function removeFromCart(key: string) {
    setCart((prev) => prev.filter((item) => cartItemKey(item) !== key));
  }

  async function handleSubmit() {
    const name = selectedCustomer?.name || customerName.trim();
    if (!name) return;
    if (cart.length === 0) return;

    await createSale.mutateAsync({
      sale: {
        customer_id: selectedCustomer?.id || undefined,
        customer_name: name,
        customer_phone: selectedCustomer?.phone || customerPhone.trim() || undefined,
        total_amount: totalAmount,
        discount_amount: discount,
        paid_amount: paidAmount,
        payment_method: paymentMethod,
        payment_status: paymentStatus,
        payment_detail: paymentMethod === 'mixed'
          ? { card: mixedCard, cash: mixedCash, transfer: mixedTransfer }
          : { [paymentMethod]: paidAmount },
        sale_date: saleDate,
        sale_channel: saleChannel,
        customer_type: customerType || undefined,
        contract_id: contractId,
        source_consultation_id: sourceConsultation?.id || undefined,
        memo: memo.trim() || undefined,
      },
      items: cart.map((item) => ({
        product_id: item.product?.id || undefined,
        // 073/074: catalog delivery_name 우선 사용 → 없으면 기존 fallback (price_groups display_name 또는 product.name)
        product_name: item.product
          ? ((catalogEntryMap.get(item.product.id)?.delivery_name?.trim()) || getProductDisplayName(item.product, customerType, priceGroups))
          : (item.customName || '임시 제품'),
        sku: item.product?.sku || undefined,
        category: item.product?.category || undefined,
        quantity: item.quantity,
        unit_price: item.unitPrice,
        total_price: item.unitPrice * item.quantity,
        serial_ids: item.selectedSerialIds,
        manual_serials: item.manualSerials || [],
      })),
    });

    router.push('/sales');
  }

  // 제품 검색/필터
  const [productSearch, setProductSearch] = useState('');
  const [productCategory, setProductCategory] = useState<string>('all');

  const CATEGORY_LABEL: Record<string, string> = {
    BL: '블런트',
    TH: '틴닝',
    LO: '장가위',
    SL: '슬라이싱',
    RS: '복원수리',
  };

  const filteredProducts = products.filter((p) => {
    const matchSearch = !productSearch ||
      p.name.toLowerCase().includes(productSearch.toLowerCase()) ||
      (p.sku || '').toLowerCase().includes(productSearch.toLowerCase());
    const matchCat = productCategory === 'all' || p.category === productCategory;
    return matchSearch && matchCat;
  });

  return (
    <>
      <Topbar title="판매 입력" />

      <div className="px-4 md:px-6 py-4 space-y-3">
        {/* 070: 상담 link 안내 배너 */}
        {sourceConsultation && (
          <div className="flex items-center gap-2 px-3 py-2 rounded-lg bg-blue-50 border border-blue-200 text-xs">
            <span className="text-blue-700 font-semibold">출장/매장상담에서 가져옴 ·</span>
            <Link href={`/consultations/${sourceConsultation.id}`} className="text-blue-700 font-mono font-bold underline">
              {sourceConsultation.unique_id}
            </Link>
            <span className="text-neutral-500 ml-auto">저장 시 이 판매가 상담에 연결되어 후기 관리가 자동 통합됩니다</span>
          </div>
        )}

        {/* 제목 + 검색 (그리드 밖 — 좌열 너비만큼만) */}
        <div className="xl:w-[33.33%]">
          <h3 className="text-sm font-semibold text-neutral-700 mb-2">제품 선택</h3>
          {/* 검색 + 카테고리 필터 */}
            <div className="flex items-center gap-2">
              <input
                type="text"
                value={productSearch}
                onChange={(e) => setProductSearch(e.target.value)}
                placeholder="제품명 또는 SKU 검색"
                className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-neutral-300"
              />
              <div className="flex gap-1 shrink-0">
                <button
                  onClick={() => setProductCategory('all')}
                  className={`px-2.5 py-1.5 text-xs rounded-md border transition ${
                    productCategory === 'all' ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'
                  }`}
                >전체</button>
                {Object.entries(CATEGORY_LABEL).map(([key, label]) => (
                  <button
                    key={key}
                    onClick={() => setProductCategory(key)}
                    className={`px-2.5 py-1.5 text-xs rounded-md border transition ${
                      productCategory === key ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'
                    }`}
                  >{label}</button>
                ))}
              </div>
            </div>
        </div>

        {/* 3열 그리드 — 검색바 아래 라인부터 동일 높이 시작 */}
        <div className="grid grid-cols-1 xl:grid-cols-3 gap-4 xl:items-start">
          {/* 좌: 제품 목록 */}
          <div className="space-y-3">
            {/* 제품 목록 테이블 */}
            {productsLoading ? (
              <div className="text-sm text-neutral-400 py-4">로딩중...</div>
            ) : (
              <Card padding={false}>
                <div className="flex-1 overflow-y-auto" style={{ maxHeight: 'calc(100vh - 320px)', minHeight: '300px' }}>
                  <table className="w-full text-sm">
                    <thead className="bg-neutral-50 sticky top-0">
                      <tr className="text-xs text-neutral-500">
                        <th className="text-left px-3 py-2 font-medium">제품명</th>
                        <th className="text-right px-3 py-2 font-medium w-24">단가</th>
                        <th className="text-right px-3 py-2 font-medium w-14">재고</th>
                        <th className="text-center px-3 py-2 font-medium w-12">추가</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-neutral-100">
                      {filteredProducts.map((p) => {
                        const inCart = cart.find((c) => c.product?.id === p.id);
                        const price = getUnitPrice(p, customerType, priceGroups);
                        const isDiscounted = hasGroupPrice(p, customerType, priceGroups);
                        return (
                          <tr
                            key={p.id}
                            onClick={() => addToCart(p)}
                            className={`cursor-pointer transition ${
                              inCart ? 'bg-neutral-900/5' : 'hover:bg-warm-ivory/60'
                            }`}
                          >
                            <td className="px-3 py-2.5">
                              <span className="font-medium text-indigo-black">{getProductDisplayName(p, customerType, priceGroups)}</span>
                              {inCart && <span className="ml-2 text-xs font-bold text-neutral-900">×{inCart.quantity}</span>}
                            </td>
                            <td className="px-3 py-2.5 text-right">
                              <span className="font-bold">{formatKRW(price)}</span>
                              {isDiscounted && (
                                <span className="block text-[10px] text-neutral-400 line-through">{formatKRW(p.price)}</span>
                              )}
                            </td>
                            <td className="px-3 py-2.5 text-right text-xs text-neutral-500">{p.stock_quantity ?? '-'}</td>
                            <td className="px-3 py-2.5 text-center">
                              <button
                                onClick={(e) => { e.stopPropagation(); addToCart(p); }}
                                className="w-7 h-7 rounded-md bg-neutral-100 hover:bg-neutral-200 text-neutral-600 text-sm font-bold transition"
                              >+</button>
                            </td>
                          </tr>
                        );
                      })}
                      {filteredProducts.length === 0 && (
                        <tr>
                          <td colSpan={4} className="px-3 py-6 text-center text-neutral-400 text-xs">
                            {productSearch ? '검색 결과가 없습니다' : '등록된 제품이 없습니다'}
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </Card>
            )}

            {/* 임시 제품 직접 입력 */}
            <Card>
              <h3 className="text-sm font-semibold text-neutral-700 mb-2">+ 직접 입력</h3>
              <p className="text-xs text-neutral-400 mb-2">등록되지 않은 제품을 직접 입력합니다 (빗, 소모품 등)</p>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={customProductName}
                  onChange={(e) => setCustomProductName(e.target.value)}
                  placeholder="제품명"
                  className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-neutral-300"
                />
                <input
                  type="number"
                  value={customProductPrice}
                  onChange={(e) => setCustomProductPrice(e.target.value)}
                  placeholder="금액"
                  className="w-28 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-neutral-300"
                />
                <Button
                  size="sm"
                  variant="secondary"
                  onClick={addCustomProduct}
                  disabled={!customProductName.trim() || !(parseInt(customProductPrice) > 0)}
                >
                  추가
                </Button>
              </div>
            </Card>
          </div>

          {/* 중: 고객 + 결제 */}
          <div className="space-y-3">
            {/* 고객 정보 */}
            <Card>
              <h3 className="text-sm font-semibold text-indigo-black mb-3">고객 정보</h3>
              <CustomerAutocomplete
                selectedCustomer={selectedCustomer}
                onSelect={(c) => { setSelectedCustomer(c); setCustomerName(c.name); setCustomerPhone(c.phone || ''); recalcCartPrices(c.customer_type); }}
                onClear={() => { setSelectedCustomer(null); setCustomerName(''); setCustomerPhone(''); recalcCartPrices(undefined); }}
              />
              {!selectedCustomer && (
                <button type="button" onClick={() => setShowCreateCustomer(true)}
                  className="w-full mt-2 py-2 rounded-lg border border-dashed border-neutral-300 text-xs text-neutral-500 hover:bg-neutral-50 transition">
                  + 신규 고객 등록
                </button>
              )}
              {selectedCustomer && (
                <div className="mt-3 pt-3 border-t border-neutral-100 space-y-1.5 text-xs text-neutral-500">
                  {selectedCustomer.address_road && (
                    <div>
                      <span className="text-neutral-400">주소: </span>
                      <span>{selectedCustomer.postcode || ''} {selectedCustomer.address_road || ''} {selectedCustomer.address_detail || ''}</span>
                    </div>
                  )}
                  {/* 태그 — 편집 가능 */}
                  {availableTags.length > 0 && (
                    <div>
                      <span className="text-neutral-400 block mb-1">태그</span>
                      <TagSelector
                        availableTags={availableTags}
                        selectedTags={selectedCustomer.tags || []}
                        onChange={async (tags) => {
                          setSelectedCustomer({ ...selectedCustomer, tags });
                          // 즉시 DB 반영
                          try {
                            await fetch(`/api/customers/${selectedCustomer.id}`, {
                              method: 'PATCH',
                              headers: { 'Content-Type': 'application/json' },
                              body: JSON.stringify({ tags }),
                            });
                          } catch { /* 조용히 실패 */ }
                        }}
                      />
                    </div>
                  )}
                  {/* 메모 — 고정 높이 영역 (메모 없어도 공간 유지) */}
                  <div className="min-h-[48px] p-2 rounded-lg bg-neutral-50 border border-neutral-100">
                    {selectedCustomer.memo ? (
                      <p className="text-xs text-neutral-600">{selectedCustomer.memo}</p>
                    ) : (
                      <p className="text-xs text-neutral-300 italic">고객 메모 없음</p>
                    )}
                  </div>
                </div>
              )}
              {customerType === 'dealer' && <p className="text-xs text-purple-600 mt-1">딜러가 적용 중</p>}
              {customerType === 'academy' && <p className="text-xs text-emerald-600 mt-1">아카데미가 적용 중</p>}
            </Card>

            <CustomerCreateModal
              open={showCreateCustomer}
              onClose={() => setShowCreateCustomer(false)}
              onCreated={(c) => { setSelectedCustomer({ id: c.id, name: c.name, phone: c.phone, customer_type: c.customer_type } as SelectedCustomer); setCustomerName(c.name); setCustomerPhone(c.phone); recalcCartPrices(c.customer_type); }}
            />

            {/* 결제 정보 */}
            <Card>
              <h3 className="text-sm font-semibold text-indigo-black mb-3">결제 정보</h3>
              <div className="space-y-3">
                <div>
                  <label className="text-xs text-neutral-500 mb-1 block">판매일</label>
                  <input type="date" value={saleDate} max={toLocalDateString(new Date())} onChange={(e) => setSaleDate(e.target.value)}
                    className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-neutral-300" />
                </div>
                <div>
                  <label className="text-xs text-neutral-500 mb-1 block">판매 채널</label>
                  <div className="flex gap-1">
                    {([{ value: 'offline', label: '오프라인', color: 'bg-neutral-800 text-white' }, { value: 'talk', label: '온라인상담', color: 'bg-yellow-500 text-white' }] as const).map((ch) => (
                      <button key={ch.value} onClick={() => setSaleChannel(ch.value)}
                        className={`flex-1 py-1.5 rounded-lg text-xs font-semibold transition ${saleChannel === ch.value ? ch.color : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>{ch.label}</button>
                    ))}
                  </div>
                </div>
                <div>
                  <label className="text-xs text-neutral-500 mb-1 block">결제 수단</label>
                  <div className="flex gap-1">
                    {(['card', 'cash', 'transfer', 'mixed'] as const).map((method) => (
                      <button key={method} onClick={() => { setPaymentMethod(method); if (method === 'mixed') { setMixedCard(0); setMixedCash(0); setMixedTransfer(0); } }}
                        className={`flex-1 py-1.5 rounded-lg text-xs font-semibold transition ${paymentMethod === method ? 'bg-terracotta text-cream' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>
                        {{ card: '카드', cash: '현금', transfer: '이체', mixed: '복합' }[method]}
                      </button>
                    ))}
                  </div>
                  {paymentMethod === 'mixed' && (
                    <div className="mt-2 p-2 rounded-lg bg-neutral-50 border border-neutral-200 space-y-1.5">
                      {(['카드', '현금', '이체'] as const).map((label, i) => {
                        const val = [mixedCard, mixedCash, mixedTransfer][i];
                        const setter = [setMixedCard, setMixedCash, setMixedTransfer][i];
                        return (
                          <div key={label} className="flex items-center gap-1.5">
                            <span className="text-[10px] text-neutral-500 w-8">{label}</span>
                            <input type="number" min={0} value={val || ''} onChange={(e) => setter(parseInt(e.target.value) || 0)}
                              className="flex-1 h-7 px-2 rounded border border-neutral-200 text-xs text-right bg-white focus:outline-none focus:ring-1 focus:ring-terracotta/40" placeholder="0" />
                          </div>
                        );
                      })}
                      {(() => { const t = mixedCard + mixedCash + mixedTransfer; const ok = t === finalAmount; return (
                        <div className={`flex justify-between text-[10px] font-semibold pt-1 border-t border-neutral-200 ${ok ? 'text-green-600' : 'text-red-500'}`}>
                          <span>합계</span><span>{formatKRW(t)}{ok ? ' ✓' : ` (${formatKRW(Math.abs(finalAmount - t))} ${t > finalAmount ? '초과' : '부족'})`}</span>
                        </div>
                      ); })()}
                    </div>
                  )}
                </div>
                <div>
                  <label className="text-xs text-neutral-500 mb-1 block">결제 상태</label>
                  <div className="flex gap-1">
                    {([{ value: 'paid' as const, label: '결제완료', color: 'bg-green-600 text-white' }, { value: 'partial' as const, label: '부분결제', color: 'bg-yellow-500 text-white' }, { value: 'unpaid' as const, label: '미결제', color: 'bg-red-500 text-white' }]).map((ps) => (
                      <button key={ps.value} onClick={() => setPaymentStatus(ps.value)}
                        className={`flex-1 py-1.5 rounded-lg text-xs font-semibold transition ${paymentStatus === ps.value ? ps.color : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>{ps.label}</button>
                    ))}
                  </div>
                  {paymentStatus === 'partial' && (
                    <div className="mt-1.5">
                      <input type="number" value={depositAmount || ''} onChange={(e) => setDepositAmount(parseInt(e.target.value) || 0)} placeholder="입금액"
                        className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
                    </div>
                  )}
                  {paymentStatus !== 'paid' && finalAmount > 0 && <p className="text-[10px] text-red-500 mt-1 font-medium">미수금: {formatKRW(finalAmount - paidAmount)}</p>}
                </div>
                <div>
                  <label className="text-xs text-neutral-500">할인 금액</label>
                  <input type="number" value={discount || ''} onChange={(e) => setDiscount(parseInt(e.target.value) || 0)} placeholder="0"
                    className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
                </div>
                <div>
                  <label className="text-xs text-neutral-500">메모</label>
                  <input type="text" value={memo} onChange={(e) => setMemo(e.target.value)} placeholder="메모 (선택)"
                    className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
                </div>
              </div>
            </Card>
          </div>

          {/* 우: 장바구니 + 판매등록 */}
          <div className="space-y-3">
            <Card>
              <h3 className="text-sm font-semibold text-indigo-black mb-3 flex items-center gap-2">
                <ShoppingBag size={16} />
                장바구니 ({cart.length})
              </h3>
              {cart.length === 0 ? (
                <p className="text-xs text-neutral-400 text-center py-4">제품을 선택해주세요</p>
              ) : (
                <div className="space-y-2">
                  {cart.map((item) => {
                    const key = cartItemKey(item);
                    const name = item.product ? getProductDisplayName(item.product, customerType) : (item.customName || '임시 제품');
                    const isCustom = !item.product;
                    return (
                      <div key={key}>
                        <div className="flex items-center gap-2">
                          <div className="flex-1 min-w-0">
                            <div className="flex items-center gap-1.5">
                              <p className="text-sm font-medium truncate">{name}</p>
                              {isCustom && <span className="text-[10px] px-1.5 py-0.5 rounded bg-orange-100 text-orange-600 font-medium shrink-0">임시</span>}
                            </div>
                            <p className="text-xs text-neutral-500">{formatKRW(item.unitPrice)} x {item.quantity} = {formatKRW(item.unitPrice * item.quantity)}</p>
                          </div>
                          <div className="flex items-center gap-1">
                            <button onClick={() => updateQuantity(key, -1)} className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200"><Minus size={12} /></button>
                            <input type="number" min="1" value={item.quantity}
                              onChange={(e) => { const val = parseInt(e.target.value) || 1; setCart((prev) => prev.map((it) => cartItemKey(it) === key ? { ...it, quantity: Math.max(1, val) } : it)); }}
                              className="w-10 h-6 text-center text-sm font-semibold border border-neutral-200 rounded bg-white focus:outline-none focus:ring-1 focus:ring-neutral-300 [appearance:textfield] [&::-webkit-inner-spin-button]:appearance-none" />
                            <button onClick={() => updateQuantity(key, 1)} className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200"><Plus size={12} /></button>
                            <button onClick={() => removeFromCart(key)} className="w-6 h-6 rounded bg-red-50 flex items-center justify-center hover:bg-red-100 text-red-500"><Trash2 size={12} /></button>
                          </div>
                        </div>
                        {item.product && customerType !== 'dealer' && customerType !== 'academy' ? (
                          <SerialPicker productId={item.product.id} quantity={item.quantity} selectedSerialIds={item.selectedSerialIds}
                            onSelect={(ids) => updateSerialIds(key, ids)} manualSerials={item.manualSerials || []} onManualSerialsChange={(s) => updateManualSerials(key, s)} />
                        ) : item.product && (customerType === 'dealer' || customerType === 'academy') ? (
                          <p className="text-[10px] text-neutral-400 mt-1">보관창고에서 출고 (시리얼 미부여)</p>
                        ) : !item.product ? (
                          <ManualSerialInput serials={item.manualSerials || []} onChange={(s) => updateManualSerials(key, s)} />
                        ) : null}
                      </div>
                    );
                  })}
                </div>
              )}
              {cart.length > 0 && (
                <div className="mt-3 pt-3 border-t border-neutral-100 space-y-1">
                  <div className="flex justify-between text-sm"><span className="text-neutral-500">소계</span><span className="font-semibold">{formatKRW(totalAmount)}</span></div>
                  {discount > 0 && <div className="flex justify-between text-sm"><span className="text-neutral-500">할인</span><span className="text-red-600">-{formatKRW(discount)}</span></div>}
                  <div className="flex justify-between text-sm font-bold"><span>결제 금액</span><span className="text-terracotta">{formatKRW(paidAmount)}</span></div>
                  {(() => {
                    const cardAmt = paymentMethod === 'card' ? paidAmount : paymentMethod === 'mixed' ? mixedCard : 0;
                    if (cardAmt <= 0) return null;
                    const { supply, vat } = calcVAT(cardAmt);
                    return (
                      <div className="mt-2 pt-2 border-t border-dashed border-neutral-200 space-y-0.5">
                        <div className="flex justify-between text-xs text-neutral-500"><span>카드 공급가액</span><span>{formatKRW(supply)}</span></div>
                        <div className="flex justify-between text-xs text-neutral-500"><span>부가세 (10%)</span><span>{formatKRW(vat)}</span></div>
                        {paymentMethod === 'mixed' && <div className="flex justify-between text-xs text-neutral-400"><span>현금/이체</span><span>{formatKRW(mixedCash + mixedTransfer)}</span></div>}
                      </div>
                    );
                  })()}
                </div>
              )}
            </Card>

            <Button className="w-full"
              disabled={(!selectedCustomer && !customerName.trim()) || cart.length === 0 || createSale.isPending}
              onClick={handleSubmit}>
              {createSale.isPending ? '등록 중...' : `판매 등록 (${formatKRW(paidAmount)})`}
            </Button>
          </div>
        </div>
      </div>
    </>
  );
}

/** 직접입력 품목용 시리얼 직접 입력 */
function ManualSerialInput({ serials, onChange }: { serials: string[]; onChange: (s: string[]) => void }) {
  const [input, setInput] = useState('');
  const [open, setOpen] = useState(false);

  return (
    <div className="mt-1">
      <button type="button" onClick={() => setOpen(!open)}
        className={`text-xs px-2 py-1 rounded transition ${serials.length > 0 ? 'bg-blue-50 text-blue-700' : 'bg-neutral-50 text-neutral-500 hover:bg-neutral-100'}`}>
        # 시리얼 {serials.length > 0 ? `${serials.length}개` : '직접 입력'}
      </button>
      {open && (
        <div className="mt-1 p-2 rounded border border-neutral-200 bg-white space-y-1.5">
          {serials.map((s, i) => (
            <div key={i} className="flex items-center gap-1.5">
              <span className="font-mono text-xs text-neutral-700 flex-1">{s}</span>
              <button type="button" onClick={() => onChange(serials.filter((_, j) => j !== i))} className="text-red-400 hover:text-red-600 text-xs">×</button>
            </div>
          ))}
          <div className="flex gap-1.5">
            <input type="text" value={input} onChange={(e) => setInput(e.target.value)}
              onKeyDown={(e) => { if (e.key === 'Enter' && input.trim()) { e.preventDefault(); onChange([...serials, input.trim()]); setInput(''); } }}
              placeholder="시리얼 번호 입력 후 엔터"
              className="flex-1 h-7 px-2 rounded border border-neutral-200 text-xs font-mono placeholder:text-neutral-400" />
            <button type="button" onClick={() => { if (input.trim()) { onChange([...serials, input.trim()]); setInput(''); } }}
              className="px-2 py-1 text-xs bg-neutral-900 text-white rounded">추가</button>
          </div>
          <button type="button" onClick={async () => {
            try {
              const res = await fetch('/api/serials/batch');
              const data = await res.json();
              let next = data.next_start || 13790001;
              // 이미 추가된 번호와 겹치면 +1
              const existing = serials.map((s) => parseInt(s, 10)).filter((n) => !isNaN(n));
              while (existing.includes(next)) next++;
              onChange([...serials, String(next)]);
            } catch { /* ignore */ }
          }} className="text-[10px] text-green-600 hover:text-green-800 font-medium">
            자동 번호 생성
          </button>
        </div>
      )}
      {!open && serials.length > 0 && (
        <div className="mt-0.5 flex flex-wrap gap-1">
          {serials.map((s, i) => (
            <span key={i} className="px-1.5 py-0.5 rounded bg-blue-50 text-blue-700 text-[10px] font-mono">{s}</span>
          ))}
        </div>
      )}
    </div>
  );
}
