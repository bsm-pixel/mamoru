'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useProducts } from '@/hooks/use-sales';
import { useCreatePurchaseOrder, useSupplierCatalog } from '@/hooks/use-purchasing';
import { useSetting } from '@/hooks/use-settings';
import { formatKRW, calcVAT } from '@/lib/utils/format';
import { SupplierSelect } from '@/components/ui/supplier-select';
import { LowStockPickerModal } from '@/components/purchasing/low-stock-picker-modal';
import { ArrowLeft, Minus, Plus, Trash2, AlertTriangle, Filter, Search } from 'lucide-react';
import type { Product } from '@/lib/supabase/types';

interface POItem {
  product: Product | null;
  product_name: string;
  sku: string;
  quantity: number;
  unit_price: number;
}

export default function NewPurchaseOrderPage() {
  const router = useRouter();
  const { data: products = [] } = useProducts();
  const createPO = useCreatePurchaseOrder();

  const [supplierName, setSupplierName] = useState('');
  const [supplierId, setSupplierId] = useState('');
  const [orderDate, setOrderDate] = useState(new Date().toISOString().slice(0, 10));
  const [expectedDate, setExpectedDate] = useState('');
  const [memo, setMemo] = useState('');
  const [vatType, setVatType] = useState<'included' | 'separate' | 'none'>('included');
  const [currency, setCurrency] = useState<'KRW' | 'USD' | 'CNY'>('KRW');
  const [exchangeRate, setExchangeRate] = useState<number>(1);
  const [items, setItems] = useState<POItem[]>([]);
  const [lowStockOpen, setLowStockOpen] = useState(false);
  const [catalogOnly, setCatalogOnly] = useState(false);
  const [productSearch, setProductSearch] = useState('');

  // 매입처 카탈로그 조회
  const { data: catalogData } = useSupplierCatalog(supplierId);
  const catalogProductIds = new Set((catalogData?.catalog || []).map(c => c.product_id));

  const lowStockThreshold = useSetting<number>('inventory.low_stock_threshold', 3);
  const lowStockCount = products.filter(
    (p) => p.stock_quantity >= 0 && p.stock_quantity <= lowStockThreshold
  ).length;
  const alreadyAddedIds = new Set(items.filter((i) => i.product?.id).map((i) => i.product!.id));

  // 매입품목 + 검색 필터링
  const displayProducts = (() => {
    let list = catalogOnly && supplierId
      ? products.filter(p => catalogProductIds.has(p.id))
      : products;
    if (productSearch) {
      const q = productSearch.toLowerCase();
      list = list.filter(p => p.name.toLowerCase().includes(q) || p.sku.toLowerCase().includes(q) || (p.product_group || '').toLowerCase().includes(q));
    }
    return list;
  })();

  const foreignTotal = items.reduce((s, i) => s + i.quantity * i.unit_price, 0);
  const rate = currency !== 'KRW' && exchangeRate > 0 ? exchangeRate : 1;
  const totalAmount = Math.round(foreignTotal * rate); // KRW 환산
  const CURRENCY_SYMBOL: Record<string, string> = { KRW: '₩', USD: '$', CNY: '¥' };

  function addProduct(product: Product) {
    setItems((prev) => {
      const existing = prev.find((i) => i.product?.id === product.id);
      if (existing) {
        return prev.map((i) =>
          i.product?.id === product.id ? { ...i, quantity: i.quantity + 1 } : i
        );
      }
      return [...prev, {
        product,
        product_name: (product as Record<string, unknown>).purchase_name as string || product.name,
        sku: product.sku,
        quantity: 1,
        unit_price: product.price_purchase || product.price,
      }];
    });
  }

  function addMultipleProducts(selected: Product[]) {
    selected.forEach((p) => addProduct(p));
  }

  function updateItemQty(idx: number, delta: number) {
    setItems((prev) => prev.map((item, i) =>
      i === idx ? { ...item, quantity: Math.max(1, item.quantity + delta) } : item
    ));
  }

  function updateItemPrice(idx: number, price: number) {
    setItems((prev) => prev.map((item, i) =>
      i === idx ? { ...item, unit_price: price } : item
    ));
  }

  function removeItem(idx: number) {
    setItems((prev) => prev.filter((_, i) => i !== idx));
  }

  async function handleSubmit() {
    if (!supplierName.trim() || items.length === 0) return;

    await createPO.mutateAsync({
      supplier_id: supplierId || undefined,
      supplier_name: supplierName.trim(),
      order_date: orderDate,
      expected_date: expectedDate || undefined,
      memo: memo.trim() || undefined,
      vat_type: vatType,
      currency,
      exchange_rate: rate,
      items: items.map((i) => ({
        product_id: i.product?.id,
        product_name: i.product_name,
        sku: i.sku,
        quantity: i.quantity,
        unit_price: i.unit_price,
      })),
    });

    router.push('/purchasing');
  }

  return (
    <>
      <Topbar title="발주 작성" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <Button variant="ghost" size="sm" onClick={() => router.push('/purchasing')}>
          <ArrowLeft size={14} />
          목록으로
        </Button>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
          {/* 좌측: 제품 선택 */}
          <div className="lg:col-span-2">
            <Card>
              <div className="flex items-center justify-between mb-3">
                <h3 className="text-sm font-bold text-indigo-black">제품 선택</h3>
                <div className="flex items-center gap-2">
                  {supplierId && catalogProductIds.size > 0 && (
                    <button
                      onClick={() => setCatalogOnly(!catalogOnly)}
                      className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-semibold transition ${
                        catalogOnly
                          ? 'bg-neutral-900 text-white'
                          : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
                      }`}
                    >
                      <Filter size={12} />
                      매입품목만 ({catalogProductIds.size})
                    </button>
                  )}
                  {lowStockCount > 0 && (
                    <button
                      onClick={() => setLowStockOpen(true)}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-amber-50 text-amber-600 text-xs font-semibold hover:bg-amber-100 transition"
                    >
                      <AlertTriangle size={13} />
                      저재고 {lowStockCount}건
                    </button>
                  )}
                </div>
              </div>
              <div className="relative mb-3">
                <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
                <input
                  type="text"
                  value={productSearch}
                  onChange={(e) => setProductSearch(e.target.value)}
                  placeholder="제품명, SKU, 제품군 검색..."
                  className="w-full h-8 pl-8 pr-3 rounded-lg border border-neutral-200 bg-warm-ivory text-xs placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 transition"
                />
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                {displayProducts.map((p) => {
                  const inList = items.find((i) => i.product?.id === p.id);
                  return (
                    <button
                      key={p.id}
                      onClick={() => addProduct(p)}
                      className={`p-3 rounded-lg border text-left transition ${
                        inList
                          ? 'border-terracotta bg-terracotta/5'
                          : 'border-neutral-200 bg-card-white hover:border-terracotta/40'
                      }`}
                    >
                      <div className="flex items-center justify-between gap-1">
                        <p className="text-sm font-semibold text-indigo-black truncate">{p.name}</p>
                        {p.stock_quantity >= 0 && (
                          <span className={`text-xs font-bold shrink-0 ${p.stock_quantity === 0 ? 'text-red-500' : p.stock_quantity <= 3 ? 'text-amber-500' : 'text-neutral-400'}`}>
                            {p.stock_quantity}
                          </span>
                        )}
                      </div>
                      <p className="text-xs text-neutral-500">{p.sku.startsWith('IW-') ? '' : p.sku}</p>
                      <p className="text-xs text-neutral-600 mt-1">
                        매입가 {p.price_purchase > 0 ? formatKRW(p.price_purchase) : '-'}
                      </p>
                    </button>
                  );
                })}
              </div>
            </Card>
          </div>

          {/* 우측: 발주 정보 */}
          <div className="space-y-4">
            <Card>
              <h3 className="text-sm font-bold text-indigo-black mb-3">매입처 정보</h3>
              <div className="space-y-3">
                <div>
                  <label className="text-xs text-neutral-500">매입처 *</label>
                  <SupplierSelect
                    value={supplierId}
                    displayName={supplierName}
                    onChange={(id, name) => { setSupplierId(id); setSupplierName(name); }}
                    placeholder="매입처 선택"
                  />
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="text-xs text-neutral-500">발주일</label>
                    <input
                      type="date"
                      value={orderDate}
                      onChange={(e) => setOrderDate(e.target.value)}
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                    />
                  </div>
                  <div>
                    <label className="text-xs text-neutral-500">입고 예정일</label>
                    <input
                      type="date"
                      value={expectedDate}
                      onChange={(e) => setExpectedDate(e.target.value)}
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                    />
                  </div>
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
                {/* 통화 + 환율 */}
                <div>
                  <label className="text-xs text-neutral-500">통화</label>
                  <div className="flex gap-1.5 mt-1">
                    {(['KRW', 'USD', 'CNY'] as const).map((c) => (
                      <button key={c} type="button" onClick={() => { setCurrency(c); if (c === 'KRW') setExchangeRate(1); }}
                        className={`flex-1 py-1.5 text-xs rounded-md border transition ${currency === c ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'}`}>
                        {c === 'KRW' ? '₩ 원화' : c === 'USD' ? '$ 달러' : '¥ 위안'}
                      </button>
                    ))}
                  </div>
                </div>
                {currency !== 'KRW' && (
                  <div>
                    <label className="text-xs text-neutral-500">환율 (1 {currency} = ₩)</label>
                    <input
                      type="number"
                      value={exchangeRate || ''}
                      onChange={(e) => setExchangeRate(Number(e.target.value) || 0)}
                      placeholder="예: 195"
                      className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                    />
                  </div>
                )}
              </div>
            </Card>

            {/* 품목 목록 */}
            <Card>
              <h3 className="text-sm font-bold text-indigo-black mb-3">발주 품목 ({items.length})</h3>
              {items.length === 0 ? (
                <p className="text-xs text-neutral-400 text-center py-4">제품을 선택해주세요</p>
              ) : (
                <div className="space-y-3">
                  {items.map((item, idx) => (
                    <div key={idx} className="border-b border-neutral-50 pb-2 last:border-0">
                      <div className="flex items-center justify-between">
                        <p className="text-sm font-medium truncate flex-1">{item.product_name}</p>
                        <button onClick={() => removeItem(idx)} className="w-6 h-6 rounded bg-red-50 flex items-center justify-center hover:bg-red-100 text-red-500">
                          <Trash2 size={12} />
                        </button>
                      </div>
                      <div className="flex items-center gap-2 mt-1">
                        <div className="flex items-center gap-1">
                          <button onClick={() => updateItemQty(idx, -1)} className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200">
                            <Minus size={12} />
                          </button>
                          <span className="text-sm font-semibold w-8 text-center">{item.quantity}</span>
                          <button onClick={() => updateItemQty(idx, 1)} className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200">
                            <Plus size={12} />
                          </button>
                        </div>
                        <span className="text-xs text-neutral-500">x</span>
                        <input
                          type="number"
                          value={item.unit_price || ''}
                          onChange={(e) => updateItemPrice(idx, parseInt(e.target.value) || 0)}
                          className="w-24 h-7 px-2 rounded border border-neutral-200 bg-warm-ivory text-xs text-right"
                        />
                        <span className="text-xs font-semibold ml-auto">
                          {currency !== 'KRW' && <span className="text-neutral-400 mr-1">{CURRENCY_SYMBOL[currency]}{(item.quantity * item.unit_price).toLocaleString()}</span>}
                          {formatKRW(Math.round(item.quantity * item.unit_price * rate))}
                        </span>
                      </div>
                    </div>
                  ))}
                </div>
              )}

              {/* 부가세 유형 */}
              <div className="mt-3 pt-3 border-t border-neutral-200">
                <label className="text-xs font-semibold text-neutral-600 mb-2 block">부가세 유형</label>
                <div className="flex gap-1.5 mb-3">
                  {([['included', '포함'], ['separate', '별도'], ['none', '미적용']] as const).map(([key, label]) => (
                    <button key={key} type="button" onClick={() => setVatType(key)}
                      className={`flex-1 py-1.5 text-xs rounded-md border transition ${vatType === key ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'}`}>
                      {label}
                    </button>
                  ))}
                </div>
              </div>

              {items.length > 0 && (
                <div className="space-y-1">
                  {(() => {
                    const { supply, vat, payment } = calcVAT(totalAmount, vatType);
                    return (
                      <div className="space-y-0.5">
                        <div className="flex justify-between text-xs text-neutral-500">
                          <span>공급가액</span>
                          <span>{formatKRW(supply)}</span>
                        </div>
                        {vatType !== 'none' && (
                          <div className="flex justify-between text-xs text-neutral-500">
                            <span>부가세 (10%)</span>
                            <span>{formatKRW(vat)}</span>
                          </div>
                        )}
                        {currency !== 'KRW' && (
                          <div className="flex justify-between text-xs text-neutral-400 pt-1">
                            <span>외화 합계</span>
                            <span>{CURRENCY_SYMBOL[currency]}{foreignTotal.toLocaleString()} × {rate} = </span>
                          </div>
                        )}
                        <div className="flex justify-between text-sm font-bold pt-1">
                          <span>합계 (KRW)</span>
                          <span className="text-terracotta">{formatKRW(payment)}</span>
                        </div>
                      </div>
                    );
                  })()}
                </div>
              )}
            </Card>

            <Button
              className="w-full"
              disabled={!supplierName.trim() || items.length === 0 || createPO.isPending}
              onClick={handleSubmit}
            >
              {createPO.isPending ? '생성 중...' : `발주 생성 (${formatKRW(calcVAT(totalAmount, vatType).payment)})`}
            </Button>
          </div>
        </div>
      </div>

      <LowStockPickerModal
        open={lowStockOpen}
        onClose={() => setLowStockOpen(false)}
        products={catalogOnly && supplierId ? displayProducts : products}
        threshold={lowStockThreshold}
        alreadyAddedIds={alreadyAddedIds}
        onAdd={addMultipleProducts}
      />
    </>
  );
}
