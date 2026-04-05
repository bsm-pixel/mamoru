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
import { Plus, Search, Package, AlertTriangle, EyeOff } from 'lucide-react';
import type { Product } from '@/lib/supabase/types';

import { useSetting } from '@/hooks/use-settings';
import { DEFAULT_CAT_LABELS } from '@/lib/utils/setting-defaults';

const CATEGORY_COLOR: Record<string, string> = {
  BL: 'bg-blue-100 text-blue-700',
  TH: 'bg-purple-100 text-purple-700',
  LO: 'bg-green-100 text-green-700',
  SL: 'bg-orange-100 text-orange-700',
};

export default function ProductsPage() {
  const router = useRouter();
  const [showInactive, setShowInactive] = useState(false);
  const { data: products = [], isLoading } = useProducts({ includeInactive: showInactive });
  const CATEGORY_LABEL = useSetting<Record<string, string>>('inventory.category_labels', DEFAULT_CAT_LABELS);
  const catTabVisible = useSetting<Record<string, boolean>>('inventory.category_tab_visible', {});
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [panelMode, setPanelMode] = useState<'view' | 'create' | 'duplicate'>('view');
  const [duplicateData, setDuplicateData] = useState<Record<string, unknown> | null>(null);
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

  // 필터링 (부자재 SUP 제외 — /supplies에서 별도 관리)
  const filtered = useMemo(() => {
    let list = products.filter((p) => p.category !== 'SUP');
    if (category !== 'all') list = list.filter((p) => p.category === category);
    if (search) {
      const q = search.toLowerCase();
      list = list.filter((p) =>
        p.name.toLowerCase().includes(q) || p.sku.toLowerCase().includes(q)
      );
    }
    return list;
  }, [products, category, search]);

  // 카테고리 칩 목록 (동적 — 설정에서 탭 표시 여부 필터)
  const categoryChips = useMemo(() => {
    const cats = Object.keys(categoryCounts).filter((c) => c !== 'SUP').sort();
    const visibleCats = cats.filter((c) => catTabVisible[c] !== false);
    return [
      { key: 'all', label: '전체', count: products.filter((p) => p.category !== 'SUP').length },
      ...visibleCats.map((c) => ({
        key: c,
        label: CATEGORY_LABEL[c] || c,
        count: categoryCounts[c],
      })),
    ];
  }, [categoryCounts, products.length, catTabVisible, CATEGORY_LABEL]);

  return (
    <>
      <Topbar title="제품 관리" />

      <div className="flex flex-col h-full min-h-0 px-4 md:px-6 pt-4 overflow-hidden">
        {/* ── 상단 고정 영역 ── */}
        <div className="shrink-0 space-y-3 pb-3">
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
          <Button size="sm" onClick={() => { setSelectedId(null); setPanelMode('create'); }} className="shrink-0">
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

        {/* 카테고리 칩 탭 + 비활성 포함 토글 */}
        <div className="flex gap-2 overflow-x-auto pb-1 scrollbar-hide items-center">
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
          <button
            onClick={() => setShowInactive(!showInactive)}
            className={`shrink-0 ml-auto flex items-center gap-1 px-3 py-1.5 rounded-full text-xs font-medium transition ${
              showInactive ? 'bg-neutral-700 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
            }`}
          >
            <EyeOff size={11} />
            비활성 포함
          </button>
        </div>
        </div>{/* 상단 고정 영역 끝 */}

        {/* ── 하단 콘텐츠 영역 (flex-1, 내부 스크롤) ── */}

        {/* PC: 2열 레이아웃 */}
        {isLg && (
          <div className="flex gap-4 flex-1 min-h-0">
            {/* 좌측: 제품 그리드 (3/5) */}
            <div className="w-[60%] shrink-0 overflow-y-auto pr-1">
              <p className="text-xs text-neutral-500 mb-2">{filtered.length}개 제품</p>
              {isLoading ? (
                <div className="grid grid-cols-3 gap-2">
                  {Array.from({ length: 6 }).map((_, i) => (
                    <Skeleton key={i} className="h-14 rounded-xl" />
                  ))}
                </div>
              ) : filtered.length === 0 ? (
                <div className="flex flex-col items-center justify-center py-16 text-neutral-400">
                  <Package size={28} className="mb-2 opacity-40" />
                  <p className="text-sm">{search ? '검색 결과가 없습니다' : '제품이 없습니다'}</p>
                </div>
              ) : (
                <div className="grid grid-cols-3 gap-2">
                  {filtered.map((p) => (
                    <CompactProductCard
                      key={p.id}
                      product={p}
                      isSelected={selectedId === p.id}
                      showCategory={category === 'all'}
                      onSelect={() => { setSelectedId(p.id); setPanelMode('view'); }}
                    />
                  ))}
                </div>
              )}
            </div>

            {/* 우측: 상세 모니터 */}
            <div className="flex-1 min-w-0 overflow-y-auto">
              {panelMode === 'create' || panelMode === 'duplicate' ? (
                <ProductDetailPanel
                  mode={panelMode}
                  duplicateData={duplicateData as never}
                  onClose={() => { setPanelMode('view'); setDuplicateData(null); }}
                  onCreated={(id) => {
                    setPanelMode('view');
                    setDuplicateData(null);
                    if (id !== '__duplicate__') setSelectedId(id);
                  }}
                />
              ) : selectedId ? (
                <ProductDetailPanel
                  productId={selectedId}
                  mode="view"
                  onClose={() => setSelectedId(null)}
                  onCreated={(id) => {
                    if (id === '__duplicate__') {
                      // 복제 요청 — 현재 제품 데이터로 duplicate 모드 전환
                      const prod = products.find((p) => p.id === selectedId);
                      if (prod) {
                        setDuplicateData({
                          name: prod.name,
                          category: prod.category,
                          price: prod.price,
                          price_dealer: prod.price_dealer || 0,
                          price_academy: prod.price_academy || 0,
                          price_purchase: prod.price_purchase || 0,
                          description: prod.description || '',
                          imweb_product_no: '',
                          supplier_id: prod.supplier_id || '',
                        });
                        setPanelMode('duplicate');
                        setSelectedId(null);
                      }
                    } else {
                      setSelectedId(id);
                    }
                  }}
                />
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
          <div className="flex-1 min-h-0 overflow-y-auto pt-3">
            <p className="text-xs text-neutral-500 mb-2">{filtered.length}개 제품</p>
            {isLoading ? (
              <div className="grid grid-cols-2 gap-2">
                {Array.from({ length: 6 }).map((_, i) => (
                  <Skeleton key={i} className="h-14 rounded-xl" />
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
        )}

        {/* 모바일 전용 슬라이드 패널 */}
        {!isLg && (
          <SlidePanel open={!!selectedId || panelMode !== 'view'} onClose={() => { setSelectedId(null); setPanelMode('view'); setDuplicateData(null); }} title={panelMode === 'create' ? '제품 등록' : panelMode === 'duplicate' ? '제품 복제' : '제품 상세'}>
            {panelMode === 'create' || panelMode === 'duplicate' ? (
              <ProductDetailPanel mode={panelMode} duplicateData={duplicateData as never}
                onClose={() => { setPanelMode('view'); setDuplicateData(null); }}
                onCreated={(id) => { setPanelMode('view'); setDuplicateData(null); if (id !== '__duplicate__') setSelectedId(id); }} />
            ) : selectedId ? (
              <ProductDetailPanel productId={selectedId} mode="view" onClose={() => setSelectedId(null)}
                onCreated={(id) => {
                  if (id === '__duplicate__') {
                    const prod = products.find((p) => p.id === selectedId);
                    if (prod) {
                      setDuplicateData({ name: prod.name, category: prod.category, price: prod.price, price_dealer: prod.price_dealer || 0, price_academy: prod.price_academy || 0, price_purchase: prod.price_purchase || 0, description: prod.description || '', imweb_product_no: '', supplier_id: prod.supplier_id || '' });
                      setPanelMode('duplicate'); setSelectedId(null);
                    }
                  } else { setSelectedId(id); }
                }} />
            ) : null}
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
  const CATEGORY_LABEL = useSetting<Record<string, string>>('inventory.category_labels', DEFAULT_CAT_LABELS);
  const isUnused = p.stock_quantity === -1;
  return (
    <Card
      className={`cursor-pointer transition px-3 py-2 ${
        isSelected ? 'ring-2 ring-neutral-900 shadow-md' : 'hover:shadow-md'
      } ${!p.is_active ? 'opacity-50' : ''}`}
      onClick={onSelect}
    >
      {/* 1행: 이름 + 카테고리 + 재고 */}
      <div className="flex items-center gap-1.5">
        <h4 className="text-sm font-semibold text-indigo-black truncate flex-1 min-w-0">
          {p.name}
          {!p.is_active && <Badge className="bg-red-100 text-red-500 text-[8px] ml-1">비활성</Badge>}
        </h4>
        <Badge className={`${CATEGORY_COLOR[p.category] || 'bg-neutral-100 text-neutral-500'} text-[9px] px-1.5 py-0 shrink-0`}>
          {CATEGORY_LABEL[p.category] || p.category}
        </Badge>
        <span className={`text-sm font-bold w-8 text-right shrink-0 ${
          isUnused ? '' : p.stock_quantity === 0 ? 'text-red-500' : ''
        }`}>
          {isUnused ? '' : p.stock_quantity}
        </span>
      </div>
      {/* 2행: 가격 */}
      <div className="text-xs text-neutral-500 mt-0.5">{formatKRW(p.price)}</div>
    </Card>
  );
}
