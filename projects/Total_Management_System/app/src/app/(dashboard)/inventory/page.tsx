'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { useQueryClient } from '@tanstack/react-query';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Skeleton } from '@/components/ui/skeleton';
import { Button } from '@/components/ui/button';
import { useInventory, type InventoryItem } from '@/hooks/use-inventory';
import { AdjustModal } from '@/components/inventory/adjust-modal';
import { formatKRW } from '@/lib/utils/format';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { AlertTriangle, Package, Boxes, TrendingDown, ArrowUpDown, Wrench } from 'lucide-react';

const CATEGORY_TABS = [
  { value: '', label: '전체' },
  { value: '가위', label: '가위' },
  { value: '빗', label: '빗' },
  { value: '케이스', label: '케이스' },
  { value: '악세서리', label: '악세서리' },
  { value: '기타', label: '기타' },
];

type SortKey = 'name' | 'stock_quantity' | 'pending_quantity' | 'value';

export default function InventoryPage() {
  const router = useRouter();
  const queryClient = useQueryClient();
  const [category, setCategory] = useState('');
  const [search, setSearch] = useState('');
  const [lowStockOnly, setLowStockOnly] = useState(false);
  const [sortKey, setSortKey] = useState<SortKey>('name');
  const [sortAsc, setSortAsc] = useState(true);
  const [showAdjust, setShowAdjust] = useState(false);

  const { data, isLoading } = useInventory({
    category: category || undefined,
    search: search || undefined,
    lowStock: lowStockOnly,
    threshold: 3,
  });

  const items = data?.items || [];
  const summary = data?.summary;

  // 정렬
  const sorted = [...items].sort((a, b) => {
    let cmp = 0;
    switch (sortKey) {
      case 'name': cmp = a.name.localeCompare(b.name); break;
      case 'stock_quantity': cmp = a.stock_quantity - b.stock_quantity; break;
      case 'pending_quantity': cmp = a.pending_quantity - b.pending_quantity; break;
      case 'value': cmp = (a.stock_quantity * a.price_purchase) - (b.stock_quantity * b.price_purchase); break;
    }
    return sortAsc ? cmp : -cmp;
  });

  function toggleSort(key: SortKey) {
    if (sortKey === key) {
      setSortAsc(!sortAsc);
    } else {
      setSortKey(key);
      setSortAsc(key === 'name');
    }
  }

  return (
    <>
      <Topbar title="재고 현황" action={
        <Button size="sm" variant="ghost" onClick={() => setShowAdjust(true)}>
          <Wrench size={14} />
          재고 조정
        </Button>
      } />

      {showAdjust && (
        <AdjustModal
          items={items}
          onClose={() => setShowAdjust(false)}
          onSuccess={() => queryClient.invalidateQueries({ queryKey: ['inventory'] })}
        />
      )}

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 요약 카드 */}
        {summary && (
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            <SummaryCard
              icon={<Package size={18} className="text-blue-500" />}
              label="총 재고"
              value={`${summary.total_stock}개`}
            />
            <SummaryCard
              icon={<Boxes size={18} className="text-indigo-500" />}
              label="미입고"
              value={`${summary.total_pending}개`}
              sub="발주 진행 중"
            />
            <SummaryCard
              icon={<AlertTriangle size={18} className="text-amber-500" />}
              label="저재고"
              value={`${summary.low_stock_count}개`}
              sub="3개 이하"
              alert={summary.low_stock_count > 0}
            />
            <SummaryCard
              icon={<TrendingDown size={18} className="text-terracotta" />}
              label="재고 원가"
              value={formatKRW(summary.total_value)}
              sub="매입가 기준"
            />
          </div>
        )}

        {/* 검색 + 필터 */}
        <div className="flex items-center gap-3">
          <SearchInput
            value={search}
            onChange={setSearch}
            placeholder="제품명, SKU, 바코드 검색"
          />
          <button
            onClick={() => setLowStockOnly(!lowStockOnly)}
            className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition flex items-center gap-1 ${
              lowStockOnly
                ? 'bg-amber-500 text-white'
                : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
            }`}
          >
            <AlertTriangle size={12} />
            저재고
          </button>
        </div>

        {/* 카테고리 탭 */}
        <div className="flex gap-1.5 overflow-x-auto pb-1">
          {CATEGORY_TABS.map((tab) => (
            <button
              key={tab.value}
              onClick={() => setCategory(tab.value)}
              className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition ${
                category === tab.value
                  ? 'bg-terracotta text-cream'
                  : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
              }`}
            >
              {tab.label}
            </button>
          ))}
        </div>

        {/* 재고 테이블 */}
        <Card padding={false}>
          {isLoading ? (
            <div className="p-4 space-y-3">
              {Array.from({ length: 6 }).map((_, i) => (
                <Skeleton key={i} className="h-14 w-full" />
              ))}
            </div>
          ) : sorted.length === 0 ? (
            <EmptyState icon={Boxes} message={lowStockOnly ? '저재고 제품이 없습니다' : '제품이 없습니다'} />
          ) : (
            <>
              {/* 헤더 (PC) */}
              <div className="hidden md:grid grid-cols-12 gap-2 px-4 py-2.5 bg-neutral-50 text-xs font-semibold text-neutral-500 border-b border-neutral-100">
                <SortHeader label="제품" span={4} sortKey="name" currentKey={sortKey} asc={sortAsc} onClick={toggleSort} />
                <SortHeader label="현재고" span={2} sortKey="stock_quantity" currentKey={sortKey} asc={sortAsc} onClick={toggleSort} />
                <span className="col-span-2">창고별</span>
                <SortHeader label="미입고" span={2} sortKey="pending_quantity" currentKey={sortKey} asc={sortAsc} onClick={toggleSort} />
                <SortHeader label="원가 합계" span={2} sortKey="value" currentKey={sortKey} asc={sortAsc} onClick={toggleSort} />
              </div>

              {/* 행 */}
              <div className="divide-y divide-neutral-100">
                {sorted.map((item) => (
                  <InventoryRow key={item.id} item={item} onClick={() => router.push(`/products/${item.id}`)} />
                ))}
              </div>
            </>
          )}
        </Card>

        <div className="text-xs text-neutral-400">
          총 {sorted.length}개 제품
        </div>
      </div>
    </>
  );
}

/* ---------- 하위 컴포넌트 ---------- */

function SummaryCard({ icon, label, value, sub, alert }: {
  icon: React.ReactNode;
  label: string;
  value: string;
  sub?: string;
  alert?: boolean;
}) {
  return (
    <Card>
      <div className="flex items-start gap-2.5">
        <div className={`w-8 h-8 rounded-lg flex items-center justify-center shrink-0 ${alert ? 'bg-amber-50' : 'bg-neutral-50'}`}>
          {icon}
        </div>
        <div>
          <p className="text-xs text-neutral-500">{label}</p>
          <p className={`text-lg font-bold ${alert ? 'text-amber-600' : 'text-indigo-black'}`}>{value}</p>
          {sub && <p className="text-[10px] text-neutral-400">{sub}</p>}
        </div>
      </div>
    </Card>
  );
}

function SortHeader({ label, span, sortKey, currentKey, asc, onClick }: {
  label: string;
  span: number;
  sortKey: SortKey;
  currentKey: SortKey;
  asc: boolean;
  onClick: (key: SortKey) => void;
}) {
  const active = currentKey === sortKey;
  return (
    <button
      className={`col-span-${span} flex items-center gap-1 hover:text-indigo-black transition text-left`}
      onClick={() => onClick(sortKey)}
    >
      {label}
      <ArrowUpDown size={10} className={active ? 'text-terracotta' : 'text-neutral-300'} />
      {active && <span className="text-[9px] text-terracotta">{asc ? '↑' : '↓'}</span>}
    </button>
  );
}

function InventoryRow({ item, onClick }: { item: InventoryItem; onClick: () => void }) {
  const isNoStock = item.stock_quantity === -1;
  const isLow = !isNoStock && item.stock_quantity <= 3;
  const totalValue = item.stock_quantity * (item.price_purchase || 0);

  return (
    <div onClick={onClick} className="flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-warm-ivory/60 transition">
      {/* 모바일 레이아웃 */}
      <div className="md:hidden flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-indigo-black truncate">{item.name}</span>
          {isNoStock && <Badge className="bg-neutral-100 text-neutral-500 text-[10px]">미사용</Badge>}
          {isLow && <Badge className="bg-amber-100 text-amber-700 text-[10px]">저재고</Badge>}
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          <span>{item.sku}</span>
          {isNoStock ? (
            <span className="text-neutral-400">재고 미사용</span>
          ) : (
            <span>재고 <strong className={isLow ? 'text-amber-600' : 'text-indigo-black'}>{item.stock_quantity}</strong></span>
          )}
          {item.pending_quantity > 0 && (
            <span className="text-blue-600">+{item.pending_quantity} 미입고</span>
          )}
        </div>
        <div className="flex items-center gap-2 mt-0.5 text-[10px] text-neutral-400">
          {item.zone_raw > 0 && <span className="text-neutral-400">보관 {item.zone_raw}</span>}
          {item.zone_ready > 0 && <span className="text-green-600">준비 {item.zone_ready}</span>}
          {item.zone_display > 0 && <span className="text-blue-600">디스플레이 {item.zone_display}</span>}
        </div>
      </div>

      {/* PC 레이아웃 */}
      <div className="hidden md:grid grid-cols-12 gap-2 flex-1 items-center">
        <div className="col-span-4">
          <div className="flex items-center gap-2">
            <span className="text-sm font-semibold text-indigo-black truncate">{item.name}</span>
            {isNoStock && <Badge className="bg-neutral-100 text-neutral-500 text-[10px]">미사용</Badge>}
            {isLow && <Badge className="bg-amber-100 text-amber-700 text-[10px]">저재고</Badge>}
          </div>
          <p className="text-xs text-neutral-500">{item.sku}</p>
        </div>
        <div className="col-span-2">
          {isNoStock ? (
            <span className="text-sm text-neutral-400">-</span>
          ) : (
            <span className={`text-sm font-bold ${isLow ? 'text-amber-600' : 'text-indigo-black'}`}>
              {item.stock_quantity}개
            </span>
          )}
        </div>
        <div className="col-span-2 text-xs text-neutral-500 space-y-0.5">
          <div className="text-neutral-400">보관 <strong>{item.zone_raw}</strong></div>
          <div className="text-green-600">준비 <strong>{item.zone_ready}</strong></div>
          <div className="text-blue-600">디스플레이 <strong>{item.zone_display}</strong></div>
        </div>
        <div className="col-span-2">
          {item.pending_quantity > 0 ? (
            <span className="text-sm font-semibold text-blue-600">+{item.pending_quantity}</span>
          ) : (
            <span className="text-sm text-neutral-300">-</span>
          )}
        </div>
        <div className="col-span-2 text-right">
          <span className="text-sm font-semibold">{totalValue > 0 ? formatKRW(totalValue) : '-'}</span>
        </div>
      </div>
    </div>
  );
}
