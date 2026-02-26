'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { SignatureCanvas } from '@/components/contracts/signature-canvas';
import { useProducts } from '@/hooks/use-sales';
import { useCreateContract } from '@/hooks/use-contracts';
import { formatKRW } from '@/lib/utils/format';
import { Minus, Plus, Trash2, FileSignature } from 'lucide-react';
import type { Product } from '@/lib/supabase/types';

interface CartItem {
  product: Product;
  quantity: number;
  option: string;
}

export default function NewContractPage() {
  const router = useRouter();
  const { data: products = [] } = useProducts();
  const createContract = useCreateContract();

  const [cart, setCart] = useState<CartItem[]>([]);
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  const [customerEmail, setCustomerEmail] = useState('');
  const [customerAddress, setCustomerAddress] = useState('');
  const [paymentMethod, setPaymentMethod] = useState('card');
  const [installment, setInstallment] = useState(0);
  const [discount, setDiscount] = useState(0);
  const [memo, setMemo] = useState('');
  const [signatureData, setSignatureData] = useState('');

  const totalAmount = cart.reduce((s, item) => s + item.product.price * item.quantity, 0);
  const finalAmount = totalAmount - discount;

  function addToCart(product: Product) {
    setCart((prev) => {
      const existing = prev.find((item) => item.product.id === product.id);
      if (existing) {
        return prev.map((item) =>
          item.product.id === product.id ? { ...item, quantity: item.quantity + 1 } : item
        );
      }
      return [...prev, { product, quantity: 1, option: '' }];
    });
  }

  function updateQuantity(productId: string, delta: number) {
    setCart((prev) =>
      prev
        .map((item) =>
          item.product.id === productId ? { ...item, quantity: Math.max(0, item.quantity + delta) } : item
        )
        .filter((item) => item.quantity > 0)
    );
  }

  async function handleSubmit() {
    if (!customerName.trim() || cart.length === 0) return;

    await createContract.mutateAsync({
      contract: {
        customer_name: customerName.trim(),
        customer_phone: customerPhone.trim() || undefined,
        customer_email: customerEmail.trim() || undefined,
        customer_address: customerAddress.trim() || undefined,
        total_amount: totalAmount,
        discount_amount: discount,
        final_amount: finalAmount,
        payment_method: paymentMethod,
        installment_months: installment,
        signature_data: signatureData || undefined,
        memo: memo.trim() || undefined,
      },
      items: cart.map((item) => ({
        product_id: item.product.id,
        product_name: item.product.name,
        sku: item.product.sku,
        quantity: item.quantity,
        unit_price: item.product.price,
        total_price: item.product.price * item.quantity,
        option_text: item.option || undefined,
      })),
    });

    router.push('/contracts');
  }

  const grouped = products.reduce<Record<string, Product[]>>((acc, p) => {
    (acc[p.category] ||= []).push(p);
    return acc;
  }, {});

  const CAT_LABEL: Record<string, string> = { BL: '블런트', TH: '틴닝', LO: '장가위', SL: '슬라이싱' };

  return (
    <>
      <Topbar title="계약서 작성" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
          {/* 좌측: 제품 선택 */}
          <div className="lg:col-span-2 space-y-4">
            <h3 className="text-sm font-semibold text-neutral-700">제품 선택</h3>
            {Object.entries(grouped).map(([cat, prods]) => (
              <div key={cat}>
                <p className="text-xs font-semibold text-neutral-500 mb-2">{CAT_LABEL[cat] || cat}</p>
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
                        <p className="text-sm font-semibold text-indigo-black truncate">{p.name}</p>
                        <p className="text-xs text-neutral-500 mt-0.5">{p.sku}</p>
                        <p className="text-sm font-bold text-terracotta mt-1">{formatKRW(p.price)}</p>
                        {inCart && <p className="text-xs text-terracotta font-semibold mt-1">x{inCart.quantity}</p>}
                      </button>
                    );
                  })}
                </div>
              </div>
            ))}

            {/* 서명 */}
            <Card>
              <h3 className="text-sm font-semibold text-indigo-black mb-3 flex items-center gap-2">
                <FileSignature size={16} />
                고객 서명
              </h3>
              <SignatureCanvas onSign={setSignatureData} width={360} height={180} />
            </Card>
          </div>

          {/* 우측: 장바구니 + 고객 + 결제 */}
          <div className="space-y-4">
            {/* 장바구니 */}
            <Card>
              <h3 className="text-sm font-semibold text-indigo-black mb-3">선택 제품 ({cart.length})</h3>
              {cart.length === 0 ? (
                <p className="text-xs text-neutral-400 text-center py-4">제품을 선택해주세요</p>
              ) : (
                <div className="space-y-2">
                  {cart.map((item) => (
                    <div key={item.product.id} className="flex items-center gap-2">
                      <div className="flex-1 min-w-0">
                        <p className="text-sm font-medium truncate">{item.product.name}</p>
                        <p className="text-xs text-neutral-500">
                          {formatKRW(item.product.price)} x {item.quantity}
                        </p>
                      </div>
                      <div className="flex items-center gap-1">
                        <button onClick={() => updateQuantity(item.product.id, -1)} className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200">
                          <Minus size={12} />
                        </button>
                        <span className="text-sm font-semibold w-6 text-center">{item.quantity}</span>
                        <button onClick={() => updateQuantity(item.product.id, 1)} className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200">
                          <Plus size={12} />
                        </button>
                        <button onClick={() => setCart((prev) => prev.filter((c) => c.product.id !== item.product.id))} className="w-6 h-6 rounded bg-red-50 flex items-center justify-center hover:bg-red-100 text-red-500">
                          <Trash2 size={12} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}

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
                    <span>최종 금액</span>
                    <span className="text-terracotta">{formatKRW(finalAmount)}</span>
                  </div>
                </div>
              )}
            </Card>

            {/* 고객 정보 */}
            <Card>
              <h3 className="text-sm font-semibold text-indigo-black mb-3">고객 정보</h3>
              <div className="space-y-2">
                <input type="text" value={customerName} onChange={(e) => setCustomerName(e.target.value)} placeholder="고객명 *" className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
                <input type="text" value={customerPhone} onChange={(e) => setCustomerPhone(e.target.value)} placeholder="연락처" className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
                <input type="email" value={customerEmail} onChange={(e) => setCustomerEmail(e.target.value)} placeholder="이메일" className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
                <input type="text" value={customerAddress} onChange={(e) => setCustomerAddress(e.target.value)} placeholder="주소" className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
              </div>
            </Card>

            {/* 결제 */}
            <Card>
              <h3 className="text-sm font-semibold text-indigo-black mb-3">결제 정보</h3>
              <div className="space-y-3">
                <div className="flex gap-2">
                  {(['card', 'cash', 'transfer', 'mixed'] as const).map((m) => (
                    <button key={m} onClick={() => setPaymentMethod(m)} className={`flex-1 py-2 rounded-lg text-xs font-semibold transition ${paymentMethod === m ? 'bg-terracotta text-cream' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>
                      {{ card: '카드', cash: '현금', transfer: '이체', mixed: '복합' }[m]}
                    </button>
                  ))}
                </div>

                {paymentMethod === 'card' && (
                  <div>
                    <label className="text-xs text-neutral-500">할부 개월</label>
                    <select value={installment} onChange={(e) => setInstallment(parseInt(e.target.value))} className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm">
                      <option value={0}>일시불</option>
                      {[2, 3, 4, 5, 6, 9, 12].map((m) => (
                        <option key={m} value={m}>{m}개월</option>
                      ))}
                    </select>
                  </div>
                )}

                <div>
                  <label className="text-xs text-neutral-500">할인 금액</label>
                  <input type="number" value={discount || ''} onChange={(e) => setDiscount(parseInt(e.target.value) || 0)} placeholder="0" className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
                </div>

                <div>
                  <label className="text-xs text-neutral-500">메모</label>
                  <input type="text" value={memo} onChange={(e) => setMemo(e.target.value)} placeholder="메모 (선택)" className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
                </div>
              </div>
            </Card>

            <Button
              className="w-full"
              disabled={!customerName.trim() || cart.length === 0 || createContract.isPending}
              onClick={handleSubmit}
            >
              {createContract.isPending ? '생성 중...' : `계약서 생성 (${formatKRW(finalAmount)})`}
            </Button>
          </div>
        </div>
      </div>
    </>
  );
}
