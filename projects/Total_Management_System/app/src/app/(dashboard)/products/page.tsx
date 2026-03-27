'use client';

import { useState, useEffect, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { ProductDetailPanel } from '@/components/products/product-detail-panel';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useProducts } from '@/hooks/use-sales';
import { formatKRW } from '@/lib/utils/format';
import { Plus, Search, Package, AlertTriangle } from 'lucide-react';
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
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [category, setCategory] = useState('all');
  const [search, setSearch] = useState('');

  // PC 여부 감지 (lg:1024px+)
  const [isLg, setIsLg] = useState(false);
  useEffect(() => {
    const mql = window.matchMedia('(min-width: 1024px)');
    setIsLg(mql.matches);
    const handler = (e: MediaQueryListEvent) => setIsLg(e.matches);
    mql.addEventListener('change', handler);
    return () => mql.removeEventListener('change', handler);
  }, []);

  // 카테고리별 카운트
  const categoryCounts = useMemo(() => {
    const counts: Record<string, number> = {};
    products.forEach((p) => {
      counts[p.category] = (counts[p.category] || 0) + 1;
    });
    return counts;
  }, [products]);

  // 재고부족 카운트 (0~3, -1 제외)
  const lowStockCount = useMemo(
    () => products.filter((p) => p.stock_quantity >= 0 && p.stock_quantity <= 3).length,
    [products]
  );

  // 필터링
  const filtered = useMemo(() => {
    let list = products;
    if (category !== 'all') list = list.filter((p) => p.category === category);
    if (search) {
      const q = search.toLowerCase();
      list = list.filter((p) =>
        p.name.toLowerCase().includes(q) || p.sku.toLowerCase().includes(q)
      );
    }
    return list;
  }, [products, category, search]);

  // 카테고리 칩 목록 (동적)
  const categoryChips = useMemo(() => {
    const cats = Object.keys(categoryCounts).sort();
    return [
      { key: 'all', label: '전체', count: products.length },
      ...cats.map((c) => ({
        key: c,
        label: CATEGORY_LABEL[c] || c,
        count: categoryCounts[c],
      })),
    ];
  }, [categoryCounts, products.length]);

  return (
    <>
      <Topbar title="제품 관리" />

      <div className="px-4 md:px-6 py-4 space-y-3">
        {/* 상단: 요약 카드 + 제품 등록 버튼 */}
        <div className="flex items-center gap-3">
          <div className="flex gap-2 flex-1 min-w-0">
            <div className="flex items-center gap-2 px-3 py-2 rounded-lg bg-blue-50 min-w-0">
              <Package size={14} className="text-blue-600 shrink-0" />
              <span className="text-xs text-neutral-500">전체</span>
              <span className="text-sm font-bold text-blue-700">{products.length}</span>
            </div>
            {lowStockCount > 0 && (
              <div className="flex items-center gap-2 px-3 py-2 rounded-lg bg-red-50 min-w-0">
                <AlertTriangle size={14} className="text-red-600 shrink-0" />
                <span className="text-xs text-neutral-500">재고부족</span>
                <span className="text-sm font-bold text-red-700">{lowStockCount}</span>
              </div>
            )}
          </div>
          <Button size="sm" onClick={() => router.push('/products/new')} className="shrink-0">
            <Plus size={14} />
            제품 등록
          </Button>
        </div>

        {/* 검색 */}
        <div className="relative">
          <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="제품명, SKU 검색..."
            className="w-full h-9 pl-9 pr-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm text-indigo-black placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 transition"
          />
        </div>

        {/* 카테고리 칩 탭 */}
        <div className="flex gap-2 overflow-x-auto pb-1 scrollbar-hide">
          {categoryChips.map((chip) => (
            <button
              key={chip.key}
              onClick={() => setCategory(chip.key)}
              className={`shrink-0 flex items-center gap-1.5 px-3 py-1.5 rounded-full text-sm font-medium transition ${
                category === chip.key
                  ? 'bg-neutral-900 text-white'
                  : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'
              }`}
            >
              {chip.label}
              <span className={`text-xs ${category === chip.key ? 'text-neutral-300' : 'text-neutral-400'}`}>
                {chip.count}
              </span>
            </button>
          ))}
        </div>

        {/* PC: 2열 레이아웃 */}
        {isLg && (
          <div className="flex gap-4 h-[calc(100vh-260px)]">
            {/* 좌측: 제품 그리드 */}
            <div className="w-[40%] shrink-0 overflow-y-auto pr-1">
              <p className="text-xs text-neutral-500 mb-2">{filtered.length}개 제품</p>
              {isLoading ? (
                <div className="grid grid-cols-2 gap-2">
                  {Array.from({ length: 6 }).map((_, i) => (
                    <Skeleton key={i} className="h-24 rounded-xl" />
                  ))}
                </div>
              ) : filtered.length === 0 ? (
                <div className="flex flex-col items-center justify-center py-16 text-neutral-400">
                  <Package size={28} className="mb-2 opacity-40" />
                  <p className="text-sm">{search ? '검색 결과가 없습니다' : '제품이 없습니다'}</p>
                </div>
              ) : (
                <div className="grid grid-cols-2 gap-2">
                  {filtered.map((p) => (
                    <CompactProductCard
                      key={p.id}
                      product={p}
                      isSelected={selectedId === p.id}
                      showCategory={category === 'all'}
                      onSelect={() => setSelectedId(p.id)}
                    />
                  ))}
                </div>
              )}
            </div>

            {/* 우측: 상세 모니터 */}
            <div className="flex-1 min-w-0 overflow-y-auto">
              {selectedId ? (
                <ProductDetailPanel productId={selectedId} onClose={() => setSelectedId(null)} />
              ) : (
                <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
                  <Package size={28} className="mb-2 opacity-40" />
                  <p className="text-xs text-center">목록에서 제품을 선택하세요</p>
                </div>
              )}
            </div>
          </div>
        )}

        {/* 모바일: 그리드만 */}
        {!isLg && (
          <>
            <p className="text-xs text-neutral-500">{filtered.length}개 제품</p>
            {isLoading ? (
              <div className="grid grid-cols-2 gap-2">
                {Array.from({ length: 6 }).map((_, i) => (
                  <Skeleton key={i} className="h-24 rounded-xl" />
                ))}
              </div>
            ) : filtered.length === 0 ? (
              <div className="flex flex-col items-center justify-center py-16 text-neutral-400">
                <Package size={28} className="mb-2 opacity-40" />
                <p className="text-sm">{search ? '검색 결과가 없습니다' : '제품이 없습니다'}</p>
              </div>
            ) : (
              <div className="grid grid-cols-2 gap-2">
                {filtered.map((p) => (
                  <CompactProductCard
                    key={p.id}
                    product={p}
                    isSelected={selectedId === p.id}
                    showCategory={category === 'all'}
                    onSelect={() => setSelectedId(p.id)}
                  />
                ))}
              </div>
            )}
          </>
        )}

        {/* 모바일 전용 슬라이드 패널 */}
        {!isLg && (
          <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="제품 상세">
            {selectedId && <ProductDetailPanel productId={selectedId} onClose={() => setSelectedId(null)} />}
          </SlidePanel>
        )}
      </div>
    </>
  );
}

// ── 컴팩트 제품 카드 ──

interface CompactProductCardProps {
  product: Product;
  isSelected: boolean;
  showCategory: boolean;
  onSelect: () => void;
}

function CompactProductCard({ product: p, isSelected, showCategory, onSelect }: CompactProductCardProps) {
  return (
    <Card
      className={`cursor-pointer transition p-3 ${
        isSelected ? 'ring-2 ring-neutral-900 shadow-md' : 'hover:shadow-md'
      }`}
      onClick={onSelect}
    >
      {/* 1행: 이름 + 재고 */}
      <div className="flex items-start justify-between gap-1 mb-1">
        <h4 className="text-sm font-bold text-indigo-black truncate flex-1 min-w-0">{p.name}</h4>
        {p.stock_quantity === -1 ? (
          <Badge className="bg-neutral-100 text-neutral-400 text-[9px] shrink-0">미사용</Badge>
        ) : (
          <span className={`text-sm font-bold shrink-0 ${p.stock_quantity > 0 ? 'text-indigo-black' : p.stock_quantity === 0 ? 'text-red-500' : 'text-indigo-black'}`}>
            {p.stock_quantity}
          </span>
        )}
      </div>

      {/* 2행: SKU + 카테고리 */}
      <div className="flex items-center gap-1.5 mb-1">
        <span className="text-[11px] text-neutral-400 font-mono truncate">{p.sku}</span>
        {showCategory && (
          <Badge className={`${CATEGORY_COLOR[p.category] || 'bg-neutral-100'} text-[9px] px-1 py-0`}>
            {CATEGORY_LABEL[p.category] || p.category}
          </Badge>
        )}
      </div>

      {/* 3행: 가격 */}
      <span className="text-sm font-bold">{formatKRW(p.price)}</span>
    </Card>
  );
}
