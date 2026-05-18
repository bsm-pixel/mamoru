'use client';

import { useState, memo, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { SlidePanel } from '@/components/ui/slide-panel';
import { SaleDetailPanel } from '@/components/sales/sale-detail-panel';
import { useSales, useSalesTabCounts, useSalesStats } from '@/hooks/use-sales';
import type { SalesTab, SalesChannel, SalesDateRange } from '@/hooks/use-sales';
import { useContracts } from '@/hooks/use-contracts';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Pagination } from '@/components/ui/pagination';
import { Plus, FileSignature, Receipt, TrendingUp, Calendar, AlertCircle, ClipboardList } from 'lucide-react';
import { PrepSheetModal } from '@/components/sales/prep-sheet-modal';
import type { OfflineSale, SaleChannel as SaleChannelType } from '@/lib/supabase/types';

const PAYMENT_METHOD_LABEL: Record<string, string> = {
  card: '카드',
  cash: '현금',
  transfer: '계좌이체',
  mixed: '복합',
};

const PAYMENT_STATUS_COLOR: Record<string, string> = {
  paid: 'bg-green-100 text-green-700',
  unpaid: 'bg-red-100 text-red-700',
  partial: 'bg-yellow-100 text-yellow-700',
};

const PAYMENT_STATUS_LABEL: Record<string, string> = {
  paid: '결제완료',
  unpaid: '미결제',
  partial: '부분결제',
};

const CHANNEL_CHIP: Record<string, { label: string; className: string }> = {
  offline: { label: '오프라인', className: 'bg-neutral-100 text-neutral-600' },
  online:  { label: '온라인',  className: 'bg-blue-100 text-blue-700' },  // 레거시 데이터 호환
  talk:    { label: '온라인상담',  className: 'bg-yellow-100 text-yellow-700' },
};

const TABS: { key: SalesTab; label: string }[] = [
  { key: 'all', label: '전체' },
  { key: 'unpaid', label: '미수금' },
  { key: 'cancelled', label: '취소' },
];

const CHANNELS: { key: SalesChannel | 'b2b'; label: string }[] = [
  { key: 'all', label: '전체' },
  { key: 'offline', label: '오프라인' },
  { key: 'talk', label: '온라인상담' },
  { key: 'b2b', label: 'B2B 납품' },
];

const DATE_RANGES: { key: SalesDateRange; label: string }[] = [
  { key: 'all', label: '전체 기간' },
  { key: 'today', label: '오늘' },
  { key: 'week', label: '이번주' },
  { key: 'month', label: '이번달' },
];

export default function SalesPage() {
  const router = useRouter();
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [tab, setTab] = useState<SalesTab>('all'); // 기본 탭: 전체
  const [channel, setChannel] = useState<SalesChannel | 'b2b'>('all');
  const [dateRange, setDateRange] = useState<SalesDateRange>('month');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  const [selectedSaleId, setSelectedSaleId] = useState<string | null>(null);
  const [isLg, setIsLg] = useState(false);
  const [prepMode, setPrepMode] = useState(false);
  const [checkedIds, setCheckedIds] = useState<Set<string>>(new Set());
  const [showPrepSheet, setShowPrepSheet] = useState(false);

  useEffect(() => {
    const mq = window.matchMedia('(min-width: 1024px)');
    setIsLg(mq.matches);
    const handler = (e: MediaQueryListEvent) => setIsLg(e.matches);
    mq.addEventListener('change', handler);
    return () => mq.removeEventListener('change', handler);
  }, []);

  const { data, isLoading } = useSales({ search, page, limit: 20, tab, channel, dateRange, dateFrom: dateFrom || undefined, dateTo: dateTo || undefined });
  const { data: tabCounts } = useSalesTabCounts();
  const { data: stats } = useSalesStats();
  const { data: contractData } = useContracts({ status: 'signed', limit: 100 });
  const newContractCount = contractData?.contracts?.filter((c) => !c.offline_sale_id).length || 0;
  const sales = data?.sales || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);

  const handleTabChange = (t: SalesTab) => { setTab(t); setPage(1); };
  const toggleCheck = (id: string) => {
    setCheckedIds((prev) => { const next = new Set(prev); if (next.has(id)) next.delete(id); else next.add(id); return next; });
  };
  const handleChannelChange = (c: SalesChannel) => { setChannel(c); setPage(1); };
  const handleDateRangeChange = (d: SalesDateRange) => { setDateRange(d); setPage(1); };

  /* --- 목록 영역 (좌측/모바일) --- */
  const listContent = (
    <>
      {/* 통계 요약 카드 */}
      {stats && (
        <div className="grid grid-cols-3 gap-2">
          <div className="bg-white rounded-lg border border-neutral-200 p-3">
            <div className="flex items-center gap-1.5 text-xs text-neutral-500 mb-1">
              <Calendar size={12} />
              이번주
            </div>
            <p className="text-base font-bold text-neutral-900">{formatKRW(stats.week.amount)}</p>
            <p className="text-[11px] text-neutral-400">{stats.week.count}건</p>
          </div>
          <div className="bg-white rounded-lg border border-neutral-200 p-3">
            <div className="flex items-center gap-1.5 text-xs text-neutral-500 mb-1">
              <TrendingUp size={12} />
              이번달
            </div>
            <p className="text-base font-bold text-neutral-900">{formatKRW(stats.month.amount)}</p>
            <p className="text-[11px] text-neutral-400">{stats.month.count}건</p>
          </div>
          <div className="bg-white rounded-lg border border-neutral-200 p-3">
            <div className="flex items-center gap-1.5 text-xs text-red-500 mb-1">
              <AlertCircle size={12} />
              미수금
            </div>
            <p className="text-base font-bold text-red-600">{formatKRW(stats.outstanding)}</p>
          </div>
        </div>
      )}

      {/* B2B 거래처별 매출 (B2B 필터일 때) */}
      {channel === 'b2b' && stats?.b2b && stats.b2b.length > 0 && (
        <div className="bg-white rounded-lg border border-neutral-200 overflow-hidden">
          <div className="px-3 py-2 bg-neutral-50 text-xs font-semibold text-neutral-600">이번달 거래처별 매출</div>
          <table className="w-full text-sm">
            <tbody className="divide-y divide-neutral-100">
              {stats.b2b.map((b: { name: string; type: string; amount: number; count: number }) => (
                <tr key={b.name} className="hover:bg-warm-ivory/40">
                  <td className="px-3 py-2">
                    <span className="font-medium">{b.name}</span>
                    <span className={`ml-2 px-1.5 py-0.5 rounded text-[10px] font-medium ${b.type === 'dealer' ? 'bg-blue-100 text-blue-700' : 'bg-purple-100 text-purple-700'}`}>
                      {b.type === 'dealer' ? '딜러' : '아카데미'}
                    </span>
                  </td>
                  <td className="px-3 py-2 text-right text-xs text-neutral-500">{b.count}건</td>
                  <td className="px-3 py-2 text-right font-bold">{formatKRW(b.amount)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* 신규 계약서 알림 */}
      {newContractCount > 0 && (
        <button
          onClick={() => router.push('/contracts')}
          className="w-full flex items-center gap-3 px-4 py-3 rounded-lg bg-neutral-100 border border-neutral-200 hover:bg-neutral-150 transition text-left"
        >
          <FileSignature size={18} className="text-neutral-700 shrink-0" />
          <div className="flex-1">
            <p className="text-sm font-semibold text-neutral-800">신규 계약서 {newContractCount}건</p>
            <p className="text-xs text-neutral-500">판매 전환 대기 중인 계약서가 있습니다</p>
          </div>
        </button>
      )}

      {/* 준비표 뽑기 + 검색 */}
      <div className="flex items-center gap-2">
        <button
          onClick={() => { setPrepMode(!prepMode); setCheckedIds(new Set()); }}
          className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium transition shrink-0 ${
            prepMode ? 'bg-neutral-900 text-white' : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'
          }`}
        >
          <ClipboardList size={14} />
          {prepMode ? '선택 취소' : '준비표 뽑기'}
        </button>
        {prepMode && checkedIds.size > 0 && (
          <button
            onClick={() => setShowPrepSheet(true)}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-blue-600 text-white text-xs font-medium hover:bg-blue-700 transition shrink-0"
          >
            <Receipt size={14} />
            준비표 인쇄 ({checkedIds.size}건)
          </button>
        )}
        <SearchInput
          value={search}
          onChange={(v) => { setSearch(v); setPage(1); }}
          placeholder="판매번호, 고객명, 전화번호"
        />
      </div>

      {/* 탭 바 */}
      <div className="flex gap-1 border-b border-neutral-200">
        {TABS.map((t) => {
          const count = tabCounts?.[t.key] ?? 0;
          const active = tab === t.key;
          return (
            <button
              key={t.key}
              onClick={() => handleTabChange(t.key)}
              className={`px-3 py-2 text-sm font-medium border-b-2 transition whitespace-nowrap ${
                active
                  ? 'border-neutral-900 text-neutral-900'
                  : 'border-transparent text-neutral-400 hover:text-neutral-600'
              }`}
            >
              {t.label}
              {count > 0 && t.key === 'unpaid' && (
                <span className={`ml-1.5 text-xs px-1.5 py-0.5 rounded-full ${
                  active ? 'bg-neutral-900 text-white' : 'bg-neutral-200 text-neutral-500'
                }`}>
                  {count}
                </span>
              )}
            </button>
          );
        })}
      </div>

      {/* 채널 + 기간 필터 (한 줄 통합) */}
      <div className="flex flex-wrap items-center gap-3">
        {/* 채널 */}
        <div className="flex gap-1">
          {CHANNELS.map((c) => (
            <button
              key={c.key}
              onClick={() => handleChannelChange(c.key)}
              className={`px-2.5 py-1 text-xs rounded-full border transition ${
                channel === c.key
                  ? 'bg-neutral-900 text-white border-neutral-900'
                  : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-400'
              }`}
            >
              {c.label}
            </button>
          ))}
        </div>

        {/* 구분선 */}
        <div className="w-px h-4 bg-neutral-200" />

        {/* 기간 */}
        <div className="flex gap-1 items-center">
          {DATE_RANGES.map((d) => (
            <button
              key={d.key}
              onClick={() => { handleDateRangeChange(d.key); setDateFrom(''); setDateTo(''); }}
              className={`px-2.5 py-1 text-xs rounded-full border transition ${
                dateRange === d.key && !dateFrom
                  ? 'bg-neutral-900 text-white border-neutral-900'
                  : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-400'
              }`}
            >
              {d.label}
            </button>
          ))}
          {/* 직접 선택 */}
          <div className="flex items-center gap-1 ml-1">
            <input
              type="date"
              value={dateFrom}
              onChange={(e) => { setDateFrom(e.target.value); setDateRange('all'); setPage(1); }}
              className="h-7 px-2 rounded-md border border-neutral-200 text-xs bg-white"
            />
            <span className="text-xs text-neutral-400">~</span>
            <input
              type="date"
              value={dateTo}
              onChange={(e) => { setDateTo(e.target.value); setDateRange('all'); setPage(1); }}
              className="h-7 px-2 rounded-md border border-neutral-200 text-xs bg-white"
            />
            {(dateFrom || dateTo) && (
              <button onClick={() => { setDateFrom(''); setDateTo(''); }} className="text-[10px] text-neutral-400 hover:text-neutral-600">초기화</button>
            )}
          </div>
        </div>
      </div>

      {/* 판매 목록 */}
      <Card padding={false}>
        {isLoading ? (
          <div className="p-4 space-y-3">
            {Array.from({ length: 5 }).map((_, i) => (
              <Skeleton key={i} className="h-14 w-full" />
            ))}
          </div>
        ) : sales.length === 0 ? (
          <EmptyState icon={Receipt} message="판매 기록이 없습니다" />
        ) : (
          <div className="divide-y divide-neutral-100">
            {sales.map((sale) => (
              <SaleRow
                key={sale.id}
                sale={sale}
                selected={selectedSaleId === sale.id}
                onClick={() => setSelectedSaleId(sale.id)}
                prepMode={prepMode}
                checked={checkedIds.has(sale.id)}
                onCheck={() => toggleCheck(sale.id)}
              />
            ))}
          </div>
        )}
      </Card>

      <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} />
    </>
  );

  return (
    <>
      <Topbar
        title="판매 관리"
        action={
          <Button onClick={() => router.push('/sales/new')} size="sm">
            <Plus size={14} />
            판매 입력
          </Button>
        }
      />

      {isLg ? (
        /* PC: 마스터-디테일 2컬럼 */
        <div className="flex gap-4 px-4 md:px-6 py-4 h-full min-h-0">
          {/* 좌측: 목록 */}
          <div className="w-2/5 shrink-0 overflow-y-auto space-y-3 pr-1">
            {listContent}
          </div>
          {/* 우측: 상세 패널 */}
          <div className="flex-1 min-w-0 overflow-y-auto bg-white rounded-xl border border-neutral-200">
            {selectedSaleId ? (
              <SaleDetailPanel saleId={selectedSaleId} />
            ) : (
              <div className="flex items-center justify-center h-full text-sm text-neutral-400">
                좌측 목록에서 판매 건을 선택하세요
              </div>
            )}
          </div>
        </div>
      ) : (
        /* 모바일: 목록 + SlidePanel */
        <div className="px-4 md:px-6 py-4 space-y-4">
          {listContent}
        </div>
      )}

      {/* 모바일 상세 패널 */}
      {!isLg && (
        <SlidePanel
          open={!!selectedSaleId}
          onClose={() => setSelectedSaleId(null)}
          title="판매 상세"
          className="sm:w-[480px]"
        >
          {selectedSaleId && <SaleDetailPanel saleId={selectedSaleId} />}
        </SlidePanel>
      )}

      {/* 준비표 모달 — 복수 건 */}
      {showPrepSheet && (
        <PrepSheetModal
          saleIds={Array.from(checkedIds)}
          onClose={() => setShowPrepSheet(false)}
        />
      )}
    </>
  );
}

/* 목록 행 — PC/모바일 공용 (카드형) */
const SaleRow = memo(function SaleRow({ sale, selected, onClick, prepMode, checked, onCheck }: {
  sale: OfflineSale; selected?: boolean; onClick: () => void;
  prepMode?: boolean; checked?: boolean; onCheck?: () => void;
}) {
  const isCancelled = !!sale.cancelled_at;
  const channelInfo = CHANNEL_CHIP[(sale.sale_channel || 'offline') as SaleChannelType] || CHANNEL_CHIP.offline;

  return (
    <div
      onClick={prepMode ? onCheck : onClick}
      className={`flex items-center gap-4 px-4 py-3 cursor-pointer transition ${
        selected ? 'bg-neutral-50 border-l-2 border-l-neutral-900' : 'hover:bg-warm-ivory/60'
      } ${isCancelled ? 'opacity-50' : ''} ${checked ? 'bg-blue-50 border-l-2 border-l-blue-500' : ''}`}
    >
      {prepMode && (
        <input type="checkbox" checked={checked} onChange={onCheck}
          onClick={(e) => e.stopPropagation()}
          className="w-4 h-4 rounded border-neutral-300 shrink-0" />
      )}
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2 flex-wrap">
          <span className={`text-sm font-semibold text-indigo-black truncate ${isCancelled ? 'line-through' : ''}`}>
            {sale.customer_name}
          </span>
          {isCancelled ? (
            <Badge className="bg-neutral-200 text-neutral-500">취소</Badge>
          ) : (
            <Badge className={PAYMENT_STATUS_COLOR[sale.payment_status] || ''}>
              {PAYMENT_STATUS_LABEL[sale.payment_status] || sale.payment_status}
            </Badge>
          )}
          <Badge className={channelInfo.className}>{channelInfo.label}</Badge>
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          <span>{formatDate(sale.sale_date)}</span>
          <span>{PAYMENT_METHOD_LABEL[sale.payment_method] || sale.payment_method}</span>
        </div>
      </div>
      <div className="text-right shrink-0">
        <span className={`text-sm font-bold ${isCancelled ? 'line-through text-neutral-400' : ''}`}>{formatKRW(sale.total_amount)}</span>
      </div>
    </div>
  );
});
