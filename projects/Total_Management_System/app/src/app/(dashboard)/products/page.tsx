'use client';

import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useProducts } from '@/hooks/use-sales';
import { formatKRW } from '@/lib/utils/format';
import { Plus } from 'lucide-react';
import type { Product } from '@/lib/supabase/types';

const CATEGORY_LABEL: Record<string, string> = {
  BL: '블런트',
  TH: '틴닝',
  LO: '장가위',
  SL: '슬라이싱',
};

const CATEGORY_COLOR: Record<string, string> = {
  BL: 'bg-blue-100 text-blue-700',
  TH: 'bg-purple-100 text-purple-700',
  LO: 'bg-green-100 text-green-700',
  SL: 'bg-orange-100 text-orange-700',
};

export default function ProductsPage() {
  const router = useRouter();
  const { data: products = [], isLoading } = useProducts();

  const grouped = products.reduce<Record<string, Product[]>>((acc, p) => {
    (acc[p.category] ||= []).push(p);
    return acc;
  }, {});

  return (
    <>
      <Topbar title="제품 관리" />

      <div className="px-4 md:px-6 py-4 space-y-6">
        <Button size="sm" onClick={() => router.push('/products/new')}>
          <Plus size={14} />
          제품 등록
        </Button>

        {isLoading ? (
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
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
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                {prods.map((product) => (
                  <Card
                    key={product.id}
                    className="cursor-pointer hover:shadow-md transition-shadow"
                    onClick={() => router.push(`/products/${product.id}`)}
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
                        {/* 3단 가격 */}
                        <div className="flex items-center gap-3 mt-2">
                          <span className="text-sm font-bold text-terracotta">{formatKRW(product.price)}</span>
                          {product.price_dealer > 0 && (
                            <span className="text-xs text-purple-600">도매 {formatKRW(product.price_dealer)}</span>
                          )}
                        </div>
                      </div>
                      <div className="text-right">
                        <p className="text-xs text-neutral-500">재고</p>
                        <p className={`text-lg font-bold ${product.stock_quantity > 0 ? 'text-indigo-black' : 'text-red-500'}`}>
                          {product.stock_quantity}
                        </p>
                      </div>
                    </div>
                    {product.imweb_product_no && (
                      <p className="mt-2 text-[10px] text-neutral-400">아임웹 #{product.imweb_product_no}</p>
                    )}
                  </Card>
                ))}
              </div>
            </div>
          ))
        )}
      </div>
    </>
  );
}
