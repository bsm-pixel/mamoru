'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { ProductDetailPanel } from '@/components/products/product-detail-panel';
import { useProducts } from '@/hooks/use-sales';
import { formatKRW } from '@/lib/utils/format';
import { Plus } from 'lucide-react';
import type { Product } from '@/lib/supabase/types';

const CATEGORY_LABEL: Record<string, string> = {
  BL: '블런트',
  TH: '틴닝',
  LO: '장가위',
  SL: '슬라이싱',
  SUP: '부자재',
};

const CATEGORY_COLOR: Record<string, string> = {
  BL: 'bg-blue-100 text-blue-700',
  TH: 'bg-purple-100 text-purple-700',
  LO: 'bg-green-100 text-green-700',
  SL: 'bg-orange-100 text-orange-700',
  SUP: 'bg-neutral-200 text-neutral-700',
};

export default function ProductsPage() {
  const router = useRouter();
  const { data: products = [], isLoading } = useProducts();
  const [selectedId, setSelectedId] = useState<string | null>(null);

  const grouped = products.reduce<Record<string, Product[]>>((acc, p) => {
    (acc[p.category] ||= []).push(p);
    return acc;
  }, {});

  return (
    <>
      <Topbar title="제품 관리" />

      <div className="flex h-[calc(100vh-3.5rem)]">
        {/* 좌측: 제품 목록 */}
        <div className={`${selectedId ? 'hidden md:block md:w-1/2 lg:w-3/5' : 'w-full'} overflow-y-auto px-4 md:px-6 py-4 space-y-6 border-r border-neutral-100`}>
          <Button size="sm" onClick={() => router.push('/products/new')}>
            <Plus size={14} />
            제품 등록
          </Button>

          {isLoading ? (
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
              {Array.from({ length: 6 }).map((_, i) => (
                <Skeleton key={i} className="h-32" />
              ))}
            </div>
          ) : (
            Object.entries(grouped).map(([cat, prods]) => (
              <div key={cat}>
                <h3 className="text-sm font-semibold text-neutral-700 mb-3">
                  {CATEGORY_LABEL[cat] || cat} ({prods.length})
                </h3>
                <div className={`grid grid-cols-1 ${selectedId ? 'lg:grid-cols-2' : 'sm:grid-cols-2 lg:grid-cols-3'} gap-3`}>
                  {prods.map((product) => (
                    <Card
                      key={product.id}
                      className={`cursor-pointer transition-shadow ${
                        selectedId === product.id
                          ? 'ring-2 ring-neutral-900 shadow-md'
                          : 'hover:shadow-md'
                      }`}
                      onClick={() => setSelectedId(product.id)}
                    >
                      <div className="flex items-start justify-between">
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2 mb-1">
                            <h4 className="text-sm font-bold text-indigo-black truncate">{product.name}</h4>
                            <Badge className={CATEGORY_COLOR[product.category] || 'bg-neutral-100'}>
                              {CATEGORY_LABEL[product.category] || product.category}
                            </Badge>
                          </div>
                          <p className="text-xs text-neutral-500">{product.sku}</p>
                          <div className="flex items-center gap-3 mt-2">
                            <span className="text-sm font-bold">{formatKRW(product.price)}</span>
                            {product.price_dealer > 0 && (
                              <span className="text-xs text-purple-600">도매 {formatKRW(product.price_dealer)}</span>
                            )}
                          </div>
                        </div>
                        <div className="text-right">
                          <p className="text-xs text-neutral-500">재고</p>
                          {product.stock_quantity === -1 ? (
                            <Badge className="bg-neutral-100 text-neutral-500 text-[10px]">미사용</Badge>
                          ) : (
                            <p className={`text-lg font-bold ${product.stock_quantity > 0 ? 'text-indigo-black' : 'text-red-500'}`}>
                              {product.stock_quantity}
                            </p>
                          )}
                        </div>
                      </div>
                    </Card>
                  ))}
                </div>
              </div>
            ))
          )}
        </div>

        {/* 우측: 상세 패널 (PC) */}
        {selectedId && (
          <div className="hidden md:block md:w-1/2 lg:w-2/5 overflow-y-auto px-5 py-4 bg-warm-ivory/30">
            <ProductDetailPanel
              productId={selectedId}
              onClose={() => setSelectedId(null)}
            />
          </div>
        )}
      </div>
    </>
  );
}
