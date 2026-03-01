'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useProducts, useCreateSale } from '@/hooks/use-sales';
import { formatKRW } from '@/lib/utils/format';
import { Minus, Plus, ShoppingBag, Trash2 } from 'lucide-react';
import { CustomerAutocomplete, type SelectedCustomer } from '@/components/shared/customer-autocomplete';
import type { Product } from '@/lib/supabase/types';

interface CartItem {
  product: Product;
  quantity: number;
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
  const [discount, setDiscount] = useState(0);
  const [memo, setMemo] = useState('');

  const totalAmount = cart.reduce((s, item) => s + item.product.price * item.quantity, 0);
  const paidAmount = totalAmount - discount;

  function addToCart(product: Product) {
    setCart((prev) => {
      const existing = prev.find((item) => item.product.id === product.id);
      if (existing) {
        return prev.map((item) =>
          item.product.id === product.id ? { ...item, quantity: item.quantity + 1 } : item
        );
      }
      return [...prev, { product, quantity: 1 }];
    });
  }

  function updateQuantity(productId: string, delta: number) {
    setCart((prev) =>
      prev
        .map((item) =>
          item.product.id === productId
            ? { ...item, quantity: Math.max(0, item.quantity + delta) }
            : item
        )
        .filter((item) => item.quantity > 0)
    );
  }

  function removeFromCart(productId: string) {
    setCart((prev) => prev.filter((item) => item.product.id !== productId));
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
        memo: memo.trim() || undefined,
      },
      items: cart.map((item) => ({
        product_id: item.product.id,
        product_name: item.product.name,
        sku: item.product.sku,
        quantity: item.quantity,
        unit_price: item.product.price,
        total_price: item.product.price * item.quantity,
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
                      const inCart = cart.find((c) => c.product.id === p.id);
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
                            {formatKRW(p.price)}
                          </p>
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
                  {cart.map((item) => (
                    <div key={item.product.id} className="flex items-center gap-2">
                      <div className="flex-1 min-w-0">
                        <p className="text-sm font-medium truncate">{item.product.name}</p>
                        <p className="text-xs text-neutral-500">
                          {formatKRW(item.product.price)} x {item.quantity} = {formatKRW(item.product.price * item.quantity)}
                        </p>
                      </div>
                      <div className="flex items-center gap-1">
                        <button
                          onClick={() => updateQuantity(item.product.id, -1)}
                          className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200"
                        >
                          <Minus size={12} />
                        </button>
                        <span className="text-sm font-semibold w-6 text-center">{item.quantity}</span>
                        <button
                          onClick={() => updateQuantity(item.product.id, 1)}
                          className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200"
                        >
                          <Plus size={12} />
                        </button>
                        <button
                          onClick={() => removeFromCart(item.product.id)}
                          className="w-6 h-6 rounded bg-red-50 flex items-center justify-center hover:bg-red-100 text-red-500"
                        >
                          <Trash2 size={12} />
                        </button>
                      </div>
                    </div>
                  ))}
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
                }}
                onClear={() => {
                  setSelectedCustomer(null);
                  setCustomerName('');
                  setCustomerPhone('');
                }}
              />
            </Card>

            {/* 결제 정보 */}
            <Card>
              <h3 className="text-sm font-semibold text-indigo-black mb-3">결제 정보</h3>
              <div className="space-y-3">
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
