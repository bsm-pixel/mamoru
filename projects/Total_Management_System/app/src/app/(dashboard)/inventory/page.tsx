'use client';

import { useState, useEffect, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { useQueryClient } from '@tanstack/react-query';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Skeleton } from '@/components/ui/skeleton';
import { Button } from '@/components/ui/button';
import { ProductDetailPanel } from '@/components/products/product-detail-panel';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useInventory, type InventoryItem } from '@/hooks/use-inventory';
import { AdjustModal } from '@/components/inventory/adjust-modal';
import { formatKRW } from '@/lib/utils/format';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import {
  AlertTriangle, Package, Boxes, TrendingDown,
  ArrowUpDown, Wrench, Eye, EyeOff,
} from 'lucide-react';

// 대분류 카테고리 (재고 페이지용)
const CATEGORY_GROUPS: { key: string; label: string; codes: string[] }[] = [
  { key: 'all', label: '전체', codes: [] },
  { key: 'scissors', label: '가위', codes: ['BL', 'TH', 'LO', 'SL'] },
  { key: 'comb', label: '빗', codes: ['CB'] },
  { key: 'case', label: '가위집', codes: ['CS'] },
  { key: 'accessory', label: '악세서리', codes: ['AC'] },
  { key: 'etc', label: '기타', codes: [] }, // codes 비어있으면 나머지 전부
];

// 제품 카테고리 코드 → 한글
const CAT_LABEL: Record<string, string> = {
  BL: '블런트', TH: '틴닝', LO: '장가위', SL: '슬라이싱',
  CB: '빗', CS: '가위집', AC: '악세서리',
};
const CAT_COLOR: Record<string, string> = {
  BL: 'bg-blue-100 text-blue-700',
  TH: 'bg-purple-100 text-purple-700',
  LO: 'bg-green-100 text-green-700',
  SL: 'bg-orange-100 text-orange-700',
};

// 알려진 카테고리 코드 모음
const KNOWN_CODES = ['BL', 'TH', 'LO', 'SL', 'CB', 'CS', 'AC'];

type SortKey = 'name' | 'stock_quantity' | 'pending_quantity' | 'value';

export default function InventoryPage() {
  const router = useRouter();
  const queryClient = useQueryClient();
  const [categoryGroup, setCategoryGroup] = useState('all');
  const [search, setSearch] = useState('');
  const [lowStockOnly, setLowStockOnly] = useState(false);
  const [hideUnused, setHideUnused] = useState(true); // 미사용 기본 숨김
  const [sortKey, setSortKey] = useState<SortKey>('name');
  const [sortAsc, setSortAsc] = useState(true);
  const [showAdjust, setShowAdjust] = useState(false);
  const [selectedId, setSelectedId] = useState<string | null>(null);

  // PC 여부 감지
  const [isLg, setIsLg] = useState(false);
  useEffect(() => {
    const mql = window.matchMedia('(min-width: 1024px)');
    setIsLg(mql.matches);
    const handler = (e: MediaQueryListEvent) => setIsLg(e.matches);
    mql.addEventListener('change', handler);
    return () => mql.removeEventListener('change', handler);
  }, []);

  const { data, isLoading } = useInventory({
    search: search || undefined,
    lowStock: lowStockOnly,
    threshold: 3,
  });

  const allItems = data?.items || [];

  // 카테고리 그룹 필터 + 미사용 필터
  const filtered = useMemo(() => {
    let list = allItems;

    // 카테고리 그룹 필터
    if (categoryGroup !== 'all') {
      const group = CATEGORY_GROUPS.find((g) => g.key === categoryGroup);
      if (group) {
        if (group.codes.length > 0) {
          list = list.filter((item) => group.codes.includes(item.category));
        } else if (group.key === 'etc') {
          list = list.filter((item) => !KNOWN_CODES.includes(item.category));
        }
      }
    }

    // 미사용 숨기기
    if (hideUnused) {
      list = list.filter((item) => item.stock_quantity !== -1);
    }

    return list;
  }, [allItems, categoryGroup, hideUnused]);

  // 카테고리별 요약 (선택된 카테고리 기준)
  const summary = useMemo(() => {
    const items = filtered.filter((i) => i.stock_quantity !== -1); // 미사용 제외
    const totalStock = items.reduce((s, i) => s + i.stock_quantity, 0);
    const totalPending = items.reduce((s, i) => s + i.pending_quantity, 0);
    const lowCount = items.filter((i) => i.stock_quantity >= 0 && i.stock_quantity <= 3).length;
    const totalValue = items.reduce((s, i) => s + i.stock_quantity * (i.price_purchase || 0), 0);
    return { totalStock, totalPending, lowCount, totalValue, productCount: items.length };
  }, [filtered]);

  // 정렬
  const sorted = useMemo(() => {
    return [...filtered].sort((a, b) => {
      let cmp = 0;
      switch (sortKey) {
        case 'name': cmp = a.name.localeCompare(b.name); break;
        case 'stock_quantity': cmp = a.stock_quantity - b.stock_quantity; break;
        case 'pending_quantity': cmp = a.pending_quantity - b.pending_quantity; break;
        case 'value': cmp = (a.stock_quantity * a.price_purchase) - (b.stock_quantity * b.price_purchase); break;
      }
      return sortAsc ? cmp : -cmp;
    });
  }, [filtered, sortKey, sortAsc]);

  function toggleSort(key: SortKey) {
    if (sortKey === key) setSortAsc(!sortAsc);
    else { setSortKey(key); setSortAsc(key === 'name'); }
  }

  return (
    <>
      <Topbar title="창고·재고 관리" action={
        <Button size="sm" variant="ghost" onClick={() => setShowAdjust(true)}>
          <Wrench size={14} />
          재고 조정
        </Button>
      } />

      {showAdjust && (
        <AdjustModal
          items={allItems}
          onClose={() => setShowAdjust(false)}
          onSuccess={() => queryClient.invalidateQueries({ queryKey: ['inventory'] })}
        />
      )}

      <div className="px-4 md:px-6 py-4 space-y-3">
        {/* 요약 카드 — 선택 카테고리 기준 */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
          <SummaryCard
            icon={<Package size={16} className="text-blue-500" />}
            label="총 재고"
            value={`${summary.totalStock}개`}
            sub={`${summary.productCount}종`}
          />
          <SummaryCard
            icon={<Boxes size={16} className="text-indigo-500" />}
            label="미입고"
            value={`${summary.totalPending}개`}
            sub="발주 진행 중"
          />
          <SummaryCard
            icon={<AlertTriangle size={16} className="text-amber-500" />}
            label="저재고"
            value={`${summary.lowCount}종`}
            sub="3개 이하"
            alert={summary.lowCount > 0}
          />
          <SummaryCard
            icon={<TrendingDown size={16} className="text-terracotta" />}
            label="재고 원가"
            value={formatKRW(summary.totalValue)}
            sub="매입가 기준"
          />
        </div>

        {/* 검색 + 필터 */}
        <div className="flex items-center gap-2">
          <SearchInput value={search} onChange={setSearch} placeholder="제품명, SKU, 바코드 검색" />
          <button
            onClick={() => setLowStockOnly(!lowStockOnly)}
            className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition flex items-center gap-1 ${
              lowStockOnly ? 'bg-amber-500 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
            }`}
          >
            <AlertTriangle size={11} />
            저재고
          </button>
          <button
            onClick={() => setHideUnused(!hideUnused)}
            className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition flex items-center gap-1 ${
              hideUnused ? 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200' : 'bg-neutral-700 text-white'
            }`}
          >
            {hideUnused ? <Eye size={11} /> : <EyeOff size={11} />}
            미사용
          </button>
        </div>

        {/* 카테고리 칩 탭 */}
        <div className="flex gap-1.5 overflow-x-auto pb-1 scrollbar-hide">
          {CATEGORY_GROUPS.map((g) => (
            <button
              key={g.key}
              onClick={() => setCategoryGroup(g.key)}
              className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition ${
                categoryGroup === g.key
                  ? 'bg-neutral-900 text-white'
                  : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
              }`}
            >
              {g.label}
            </button>
          ))}
        </div>

        {/* PC: 2열 레이아웃 (테이블 + 상세 모니터) */}
        {isLg && (
          <div className="flex gap-4 h-[calc(100vh-320px)]">
            {/* 좌측: 재고 테이블 */}
            <div className={`${selectedId ? 'w-[55%]' : 'w-full'} shrink-0 overflow-y-auto`}>
              <InventoryTable
                items={sorted}
                isLoading={isLoading}
                lowStockOnly={lowStockOnly}
                selectedId={selectedId}
                sortKey={sortKey}
                sortAsc={sortAsc}
                onToggleSort={toggleSort}
                onSelect={setSelectedId}
              />
            </div>

            {/* 우측: 상세 모니터 */}
            {selectedId && (
              <div className="flex-1 min-w-0 overflow-y-auto">
                <ProductDetailPanel productId={selectedId} onClose={() => setSelectedId(null)} />
              </div>
            )}
          </div>
        )}

        {/* 모바일: 테이블만 */}
        {!isLg && (
          <InventoryTable
            items={sorted}
            isLoading={isLoading}
            lowStockOnly={lowStockOnly}
            selectedId={selectedId}
            sortKey={sortKey}
            sortAsc={sortAsc}
            onToggleSort={toggleSort}
            onSelect={setSelectedId}
          />
        )}

        {/* 모바일 전용 슬라이드 패널 */}
        {!isLg && (
          <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="제품 상세">
            {selectedId && <ProductDetailPanel productId={selectedId} onClose={() => setSelectedId(null)} />}
          </SlidePanel>
        )}

        <p className="text-xs text-neutral-400">총 {sorted.length}개 제품</p>
      </div>
    </>
  );
}

/* ── 하위 컴포넌트 ── */

function SummaryCard({ icon, label, value, sub, alert }: {
  icon: React.ReactNode; label: string; value: string; sub?: string; alert?: boolean;
}) {
  return (
    <Card>
      <div className="flex items-start gap-2">
        <div className={`w-7 h-7 rounded-lg flex items-center justify-center shrink-0 ${alert ? 'bg-amber-50' : 'bg-neutral-50'}`}>
          {icon}
        </div>
        <div>
          <p className="text-[11px] text-neutral-500">{label}</p>
          <p className={`text-base font-bold ${alert ? 'text-amber-600' : 'text-indigo-black'}`}>{value}</p>
          {sub && <p className="text-[10px] text-neutral-400">{sub}</p>}
        </div>
      </div>
    </Card>
  );
}

function InventoryTable({ items, isLoading, lowStockOnly, selectedId, sortKey, sortAsc, onToggleSort, onSelect }: {
  items: InventoryItem[];
  isLoading: boolean;
  lowStockOnly: boolean;
  selectedId: string | null;
  sortKey: SortKey;
  sortAsc: boolean;
  onToggleSort: (key: SortKey) => void;
  onSelect: (id: string) => void;
}) {
  return (
    <Card padding={false}>
      {isLoading ? (
        <div className="p-4 space-y-3">
          {Array.from({ length: 6 }).map((_, i) => (
            <Skeleton key={i} className="h-14 w-full" />
          ))}
        </div>
      ) : items.length === 0 ? (
        <EmptyState icon={Boxes} message={lowStockOnly ? '저재고 제품이 없습니다' : '제품이 없습니다'} />
      ) : (
        <>
          {/* 헤더 (PC) */}
          <div className="hidden md:grid grid-cols-12 gap-2 px-4 py-2 bg-neutral-50 text-[11px] font-semibold text-neutral-500 border-b border-neutral-100">
            <SortHeader label="제품" span={4} sortKey="name" currentKey={sortKey} asc={sortAsc} onClick={onToggleSort} />
            <SortHeader label="현재고" span={2} sortKey="stock_quantity" currentKey={sortKey} asc={sortAsc} onClick={onToggleSort} />
            <span className="col-span-3">창고별</span>
            <SortHeader label="미입고" span={1} sortKey="pending_quantity" currentKey={sortKey} asc={sortAsc} onClick={onToggleSort} />
            <SortHeader label="원가" span={2} sortKey="value" currentKey={sortKey} asc={sortAsc} onClick={onToggleSort} />
          </div>

          {/* 행 */}
          <div className="divide-y divide-neutral-100">
            {items.map((item) => (
              <InventoryRow
                key={item.id}
                item={item}
                isSelected={selectedId === item.id}
                onClick={() => onSelect(item.id)}
              />
            ))}
          </div>
        </>
      )}
    </Card>
  );
}

function SortHeader({ label, span, sortKey, currentKey, asc, onClick }: {
  label: string; span: number; sortKey: SortKey; currentKey: SortKey; asc: boolean; onClick: (key: SortKey) => void;
}) {
  const active = currentKey === sortKey;
  return (
    <button
      className={`col-span-${span} flex items-center gap-1 hover:text-indigo-black transition text-left`}
      onClick={() => onClick(sortKey)}
    >
      {label}
      <ArrowUpDown size={9} className={active ? 'text-terracotta' : 'text-neutral-300'} />
      {active && <span className="text-[8px] text-terracotta">{asc ? '↑' : '↓'}</span>}
    </button>
  );
}

function InventoryRow({ item, isSelected, onClick }: { item: InventoryItem; isSelected: boolean; onClick: () => void }) {
  const isNoStock = item.stock_quantity === -1;
  const isLow = !isNoStock && item.stock_quantity >= 0 && item.stock_quantity <= 3;
  const totalValue = isNoStock ? 0 : item.stock_quantity * (item.price_purchase || 0);
  const total = item.zone_raw + item.zone_ready + item.zone_display;

  return (
    <div
      onClick={onClick}
      className={`flex items-center gap-3 px-4 py-2.5 cursor-pointer hover:bg-warm-ivory/60 transition ${
        isSelected ? 'bg-terracotta/5 ring-1 ring-terracotta/20' : ''
      }`}
    >
      {/* 모바일 */}
      <div className="md:hidden flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-indigo-black truncate">{item.name}</span>
          {isNoStock && <Badge className="bg-neutral-100 text-neutral-500 text-[9px]">미사용</Badge>}
          {isLow && <Badge className="bg-amber-100 text-amber-700 text-[9px]">저재고</Badge>}
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          <span>{item.sku}</span>
          {!isNoStock && (
            <span>재고 <strong className={isLow ? 'text-amber-600' : 'text-indigo-black'}>{item.stock_quantity}</strong></span>
          )}
        </div>
        {!isNoStock && total > 0 && (
          <div className="flex items-center gap-2 mt-1">
            <ZoneDots raw={item.zone_raw} ready={item.zone_ready} display={item.zone_display} />
          </div>
        )}
      </div>

      {/* PC */}
      <div className="hidden md:grid grid-cols-12 gap-2 flex-1 items-center">
        <div className="col-span-4">
          <div className="flex items-center gap-2">
            <span className="text-sm font-semibold text-indigo-black truncate">{item.name}</span>
            {CAT_COLOR[item.category] && (
              <Badge className={`${CAT_COLOR[item.category]} text-[9px] px-1 py-0`}>
                {CAT_LABEL[item.category] || item.category}
              </Badge>
            )}
            {isNoStock && <Badge className="bg-neutral-100 text-neutral-500 text-[9px]">미사용</Badge>}
            {isLow && <Badge className="bg-amber-100 text-amber-700 text-[9px]">저재고</Badge>}
          </div>
          <p className="text-[11px] text-neutral-400 font-mono">{item.sku}</p>
        </div>
        <div className="col-span-2">
          {isNoStock ? (
            <span className="text-sm text-neutral-300">-</span>
          ) : (
            <span className={`text-sm font-bold ${isLow ? 'text-amber-600' : 'text-indigo-black'}`}>
              {item.stock_quantity}개
            </span>
          )}
        </div>
        <div className="col-span-3">
          {!isNoStock && total > 0 ? (
            <ZoneDots raw={item.zone_raw} ready={item.zone_ready} display={item.zone_display} />
          ) : (
            <span className="text-xs text-neutral-300">-</span>
          )}
        </div>
        <div className="col-span-1">
          {item.pending_quantity > 0 ? (
            <span className="text-xs font-semibold text-blue-600">+{item.pending_quantity}</span>
          ) : (
            <span className="text-xs text-neutral-300">-</span>
          )}
        </div>
        <div className="col-span-2 text-right">
          <span className="text-xs font-semibold">{totalValue > 0 ? formatKRW(totalValue) : '-'}</span>
        </div>
      </div>
    </div>
  );
}

/** 창고별 색상 도트 + 숫자 */
function ZoneDots({ raw, ready, display }: { raw: number; ready: number; display: number }) {
  return (
    <div className="flex items-center gap-3 text-[11px]">
      {raw > 0 && (
        <span className="flex items-center gap-1">
          <span className="w-2 h-2 rounded-full bg-neutral-400" />
          <span className="text-neutral-500">보관 {raw}</span>
        </span>
      )}
      {ready > 0 && (
        <span className="flex items-center gap-1">
          <span className="w-2 h-2 rounded-full bg-green-500" />
          <span className="text-green-700">준비 {ready}</span>
        </span>
      )}
      {display > 0 && (
        <span className="flex items-center gap-1">
          <span className="w-2 h-2 rounded-full bg-blue-500" />
          <span className="text-blue-700">디스플레이 {display}</span>
        </span>
      )}
    </div>
  );
}
