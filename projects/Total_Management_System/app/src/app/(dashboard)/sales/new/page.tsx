'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useProducts, useCreateSale } from '@/hooks/use-sales';
import { formatKRW, calcVAT } from '@/lib/utils/format';
import { Minus, Plus, ShoppingBag, Trash2 } from 'lucide-react';
import { CustomerAutocomplete, type SelectedCustomer } from '@/components/shared/customer-autocomplete';
import { SerialPicker } from '@/components/sales/serial-picker';
import type { Product } from '@/lib/supabase/types';

interface CartItem {
  product: Product | null; // null이면 임시 제품
  customName?: string;     // 임시 제품명
  customPrice?: number;    // 임시 제품 가격
  quantity: number;
  unitPrice: number;
  selectedSerialIds: string[];
}

/** cart item의 고유 키 */
function cartItemKey(item: CartItem): string {
  return item.product ? item.product.id : `custom-${item.customName}-${item.customPrice}`;
}

/** 고객 유형에 따른 단가 결정 */
function getUnitPrice(product: Product, customerType?: string): number {
  if (customerType === 'dealer' && product.price_dealer > 0) return product.price_dealer;
  if (customerType === 'academy' && product.price_academy > 0) return product.price_academy;
  return product.price;
}

export default function NewSalePage() {
  const router = useRouter();
  const { data: products = [], isLoading: productsLoading } = useProducts();
  const createSale = useCreateSale();

  const [cart, setCart] = useState<CartItem[]>([]);
  const [selectedCustomer, setSelectedCustomer] = useState<SelectedCustomer | null>(null);
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  const [paymentMethod, setPaymentMethod] = useState('card');
  const [saleChannel, setSaleChannel] = useState('offline');
  const [discount, setDiscount] = useState(0);
  const [memo, setMemo] = useState('');

  const customerType = selectedCustomer?.customer_type;
  const totalAmount = cart.reduce((s, item) => s + item.unitPrice * item.quantity, 0);
  const paidAmount = totalAmount - discount;

  // 임시 제품 입력 상태
  const [customProductName, setCustomProductName] = useState('');
  const [customProductPrice, setCustomProductPrice] = useState('');

  function addToCart(product: Product) {
    const price = getUnitPrice(product, customerType);
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
      unitPrice: item.product ? getUnitPrice(item.product, type) : item.unitPrice,
    })));
  }

  function updateSerialIds(key: string, serialIds: string[]) {
    setCart((prev) =>
      prev.map((item) =>
        cartItemKey(item) === key ? { ...item, selectedSerialIds: serialIds } : item
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
        sale_channel: saleChannel,
        memo: memo.trim() || undefined,
      },
      items: cart.map((item) => ({
        product_id: item.product?.id || undefined,
        product_name: item.product?.name || item.customName || '임시 제품',
        sku: item.product?.sku || undefined,
        quantity: item.quantity,
        unit_price: item.unitPrice,
        total_price: item.unitPrice * item.quantity,
        serial_ids: item.selectedSerialIds,
      })),
    });

    router.push('/sales');
  }

  // 카테고리별 그룹핑
  const grouped = products.reduce<Record<string, Product[]>>((acc, p) => {
    (acc[p.category] ||= []).push(p);
    return acc;
  }, {});

  const CATEGORY_LABEL: Record<string, string> = {
    BL: '블런트',
    TH: '틴닝',
    LO: '장가위',
    SL: '슬라이싱',
  };

  return (
    <>
      <Topbar title="판매 입력" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
          {/* 좌측: 제품 선택 */}
          <div className="lg:col-span-2 space-y-4">
            <h3 className="text-sm font-semibold text-neutral-700">제품 선택</h3>

            {productsLoading ? (
              <div className="text-sm text-neutral-400">로딩중...</div>
            ) : (
              Object.entries(grouped).map(([cat, prods]) => (
                <div key={cat}>
                  <p className="text-xs font-semibold text-neutral-500 mb-2">
                    {CATEGORY_LABEL[cat] || cat}
                  </p>
                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                    {prods.map((p) => {
                      const inCart = cart.find((c) => c.product?.id === p.id);
                      return (
                        <button
                          key={p.id}
                          onClick={() => addToCart(p)}
                          className={`p-3 rounded-lg border text-left transition ${
                            inCart
                              ? 'border-terracotta bg-terracotta/5'
                              : 'border-neutral-200 bg-card-white hover:border-terracotta/40'
                          }`}
                        >
                          <p className="text-sm font-semibold text-indigo-black truncate">
                            {p.name}
                          </p>
                          <p className="text-xs text-neutral-500 mt-0.5">{p.sku}</p>
                          <p className="text-sm font-bold text-terracotta mt-1">
                            {formatKRW(getUnitPrice(p, customerType))}
                          </p>
                          {customerType === 'dealer' && p.price_dealer > 0 && (
                            <p className="text-[10px] text-neutral-400 line-through">{formatKRW(p.price)}</p>
                          )}
                          {customerType === 'academy' && p.price_academy > 0 && (
                            <p className="text-[10px] text-neutral-400 line-through">{formatKRW(p.price)}</p>
                          )}
                          {inCart && (
                            <p className="text-xs text-terracotta font-semibold mt-1">
                              x{inCart.quantity}
                            </p>
                          )}
                        </button>
                      );
                    })}
                  </div>
                </div>
              ))
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

          {/* 우측: 장바구니 + 고객 + 결제 */}
          <div className="space-y-4">
            {/* 장바구니 */}
            <Card>
              <h3 className="text-sm font-semibold text-indigo-black mb-3 flex items-center gap-2">
                <ShoppingBag size={16} />
                장바구니 ({cart.length})
              </h3>

              {cart.length === 0 ? (
                <p className="text-xs text-neutral-400 text-center py-4">
                  제품을 선택해주세요
                </p>
              ) : (
                <div className="space-y-2">
                  {cart.map((item) => {
                    const key = cartItemKey(item);
                    const name = item.product?.name || item.customName || '임시 제품';
                    const isCustom = !item.product;
                    return (
                      <div key={key}>
                        <div className="flex items-center gap-2">
                          <div className="flex-1 min-w-0">
                            <div className="flex items-center gap-1.5">
                              <p className="text-sm font-medium truncate">{name}</p>
                              {isCustom && (
                                <span className="text-[10px] px-1.5 py-0.5 rounded bg-orange-100 text-orange-600 font-medium shrink-0">임시</span>
                              )}
                            </div>
                            <p className="text-xs text-neutral-500">
                              {formatKRW(item.unitPrice)} x {item.quantity} = {formatKRW(item.unitPrice * item.quantity)}
                            </p>
                          </div>
                          <div className="flex items-center gap-1">
                            <button
                              onClick={() => updateQuantity(key, -1)}
                              className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200"
                            >
                              <Minus size={12} />
                            </button>
                            <span className="text-sm font-semibold w-6 text-center">{item.quantity}</span>
                            <button
                              onClick={() => updateQuantity(key, 1)}
                              className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200"
                            >
                              <Plus size={12} />
                            </button>
                            <button
                              onClick={() => removeFromCart(key)}
                              className="w-6 h-6 rounded bg-red-50 flex items-center justify-center hover:bg-red-100 text-red-500"
                            >
                              <Trash2 size={12} />
                            </button>
                          </div>
                        </div>
                        {/* 시리얼 선택 (등록 제품만) */}
                        {item.product && (
                          <SerialPicker
                            productId={item.product.id}
                            quantity={item.quantity}
                            selectedSerialIds={item.selectedSerialIds}
                            onSelect={(ids) => updateSerialIds(key, ids)}
                          />
                        )}
                      </div>
                    );
                  })}
                </div>
              )}

              {/* 합계 */}
              {cart.length > 0 && (
                <div className="mt-3 pt-3 border-t border-neutral-100 space-y-1">
                  <div className="flex justify-between text-sm">
                    <span className="text-neutral-500">소계</span>
                    <span className="font-semibold">{formatKRW(totalAmount)}</span>
                  </div>
                  {discount > 0 && (
                    <div className="flex justify-between text-sm">
                      <span className="text-neutral-500">할인</span>
                      <span className="text-red-600">-{formatKRW(discount)}</span>
                    </div>
                  )}
                  <div className="flex justify-between text-sm font-bold">
                    <span>결제 금액</span>
                    <span className="text-terracotta">{formatKRW(paidAmount)}</span>
                  </div>
                  {/* 카드결제 시 VAT 분리 표시 */}
                  {paymentMethod === 'card' && paidAmount > 0 && (() => {
                    const { supply, vat } = calcVAT(paidAmount);
                    return (
                      <div className="mt-2 pt-2 border-t border-dashed border-neutral-200 space-y-0.5">
                        <div className="flex justify-between text-xs text-neutral-500">
                          <span>공급가액</span>
                          <span>{formatKRW(supply)}</span>
                        </div>
                        <div className="flex justify-between text-xs text-neutral-500">
                          <span>부가세 (10%)</span>
                          <span>{formatKRW(vat)}</span>
                        </div>
                      </div>
                    );
                  })()}
                </div>
              )}
            </Card>

            {/* 고객 정보 — 자동완성 */}
            <Card>
              <h3 className="text-sm font-semibold text-indigo-black mb-3">고객 정보</h3>
              <CustomerAutocomplete
                selectedCustomer={selectedCustomer}
                onSelect={(c) => {
                  setSelectedCustomer(c);
                  setCustomerName(c.name);
                  setCustomerPhone(c.phone || '');
                  recalcCartPrices(c.customer_type);
                }}
                onClear={() => {
                  setSelectedCustomer(null);
                  setCustomerName('');
                  setCustomerPhone('');
                  recalcCartPrices(undefined);
                }}
              />
              {customerType === 'dealer' && (
                <p className="text-xs text-purple-600 mt-1">딜러가 적용 중</p>
              )}
              {customerType === 'academy' && (
                <p className="text-xs text-emerald-600 mt-1">아카데미가 적용 중</p>
              )}
            </Card>

            {/* 결제 정보 */}
            <Card>
              <h3 className="text-sm font-semibold text-indigo-black mb-3">결제 정보</h3>
              <div className="space-y-3">
                {/* 판매 채널 */}
                <div>
                  <label className="text-xs text-neutral-500 mb-1.5 block">판매 채널</label>
                  <div className="flex gap-2">
                    {([
                      { value: 'offline', label: '오프라인', color: 'bg-neutral-800 text-white' },
                      { value: 'online', label: '온라인', color: 'bg-blue-600 text-white' },
                      { value: 'talk', label: '톡상담', color: 'bg-yellow-500 text-white' },
                    ] as const).map((ch) => (
                      <button
                        key={ch.value}
                        onClick={() => setSaleChannel(ch.value)}
                        className={`flex-1 py-2 rounded-lg text-xs font-semibold transition ${
                          saleChannel === ch.value
                            ? ch.color
                            : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
                        }`}
                      >
                        {ch.label}
                      </button>
                    ))}
                  </div>
                </div>

                <div className="flex gap-2">
                  {(['card', 'cash', 'transfer', 'mixed'] as const).map((method) => (
                    <button
                      key={method}
                      onClick={() => setPaymentMethod(method)}
                      className={`flex-1 py-2 rounded-lg text-xs font-semibold transition ${
                        paymentMethod === method
                          ? 'bg-terracotta text-cream'
                          : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
                      }`}
                    >
                      {{ card: '카드', cash: '현금', transfer: '이체', mixed: '복합' }[method]}
                    </button>
                  ))}
                </div>

                <div>
                  <label className="text-xs text-neutral-500">할인 금액</label>
                  <input
                    type="number"
                    value={discount || ''}
                    onChange={(e) => setDiscount(parseInt(e.target.value) || 0)}
                    placeholder="0"
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                  />
                </div>

                <div>
                  <label className="text-xs text-neutral-500">메모</label>
                  <input
                    type="text"
                    value={memo}
                    onChange={(e) => setMemo(e.target.value)}
                    placeholder="메모 (선택)"
                    className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                  />
                </div>
              </div>
            </Card>

            {/* 제출 */}
            <Button
              className="w-full"
              disabled={(!selectedCustomer && !customerName.trim()) || cart.length === 0 || createSale.isPending}
              onClick={handleSubmit}
            >
              {createSale.isPending ? '등록 중...' : `판매 등록 (${formatKRW(paidAmount)})`}
            </Button>
          </div>
        </div>
      </div>
    </>
  );
}
