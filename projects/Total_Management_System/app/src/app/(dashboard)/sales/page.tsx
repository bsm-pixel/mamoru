'use client';

import { useState, memo, useEffect, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { SlidePanel } from '@/components/ui/slide-panel';
import { SaleDetailPanel } from '@/components/sales/sale-detail-panel';
import { DeliveryDetailPanel } from '@/components/deliveries/delivery-detail-panel';
import { useSales, useSalesTabCounts, useSalesStats } from '@/hooks/use-sales';
import type { SalesTab, SalesChannel, SalesDateRange } from '@/hooks/use-sales';
import { useDeliveryStats, useDeliveries } from '@/hooks/use-deliveries';
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

/** 2026-05-26 IA 통합: 영역 칩 — 고객(B2C) / 거래처(B2B) / 전체. default 고객 */
type SalesSection = 'customer' | 'partner' | 'all';
const SECTIONS: { key: SalesSection; label: string }[] = [
  { key: 'customer', label: '고객' },
  { key: 'partner', label: '거래처' },
  { key: 'all', label: '전체' },
];

/** 영역별 채널 옵션 — 영역 칩 선택에 따라 동적 표시 */
const CHANNELS_BY_SECTION: Record<SalesSection, { key: SalesChannel | 'b2b'; label: string }[]> = {
  customer: [
    { key: 'all', label: '전체' },
    { key: 'offline', label: '오프라인' },
    { key: 'talk', label: '온라인상담' },
  ],
  partner: [
    { key: 'b2b', label: '전체' },  // Phase B 에서 dealer/academy 세분화 예정
  ],
  all: [
    { key: 'all', label: '전체' },
    { key: 'offline', label: '오프라인' },
    { key: 'talk', label: '온라인상담' },
    { key: 'b2b', label: 'B2B 납품' },
  ],
};

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
  const [section, setSection] = useState<SalesSection>('customer'); // 2026-05-26: default 고객
  const [channel, setChannel] = useState<SalesChannel | 'b2b'>('all');
  const [dateRange, setDateRange] = useState<SalesDateRange>('month');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  /** 2026-05-26 Phase B: 우측 상세 패널 — sourceType 분기. 'sale' → SaleDetailPanel / 'delivery' → DeliveryDetailPanel */
  const [selected, setSelected] = useState<{ id: string; sourceType: 'sale' | 'delivery' } | null>(null);
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
  const { data: deliveryStats } = useDeliveryStats(); // 2026-05-26: 거래처 카드 — deliveries 합산용
  // 2026-05-26 Phase B: 영역 'partner'/'all' 일 때 deliveries 목록 합집합. limit 30 (사장님 운영 규모 충분)
  const { data: deliveryData, isLoading: isDeliveriesLoading } = useDeliveries({ search, dateRange, page: 1, limit: 30 });
  const { data: contractData } = useContracts({ status: 'signed', limit: 100 });
  const newContractCount = contractData?.contracts?.filter((c) => !c.offline_sale_id).length || 0;
  const sales = data?.sales || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);

  /** 2026-05-26 Phase B: 통합 목록 — 영역 칩 따라 sale/delivery 합집합. 날짜 desc 정렬 */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  type Delivery = any;
  const unifiedItems = useMemo(() => {
    const items: Array<
      | { sourceType: 'sale'; id: string; date: string; data: OfflineSale }
      | { sourceType: 'delivery'; id: string; date: string; data: Delivery }
    > = [];
    if (section === 'customer' || section === 'all') {
      for (const s of sales) {
        items.push({ sourceType: 'sale', id: s.id, date: s.sale_date || s.created_at || '', data: s });
      }
    }
    if (section === 'partner' || section === 'all') {
      const dls = (deliveryData?.deliveries || []) as Delivery[];
      for (const d of dls) {
        if (d.cancelled_at && tab !== 'cancelled') continue; // 영역 partner/all 에서도 취소 탭 외엔 cancelled 숨김
        if (tab === 'unpaid' && !['unpaid', 'partial'].includes(d.payment_status)) continue;
        items.push({ sourceType: 'delivery', id: d.id, date: d.delivery_date || d.created_at || '', data: d });
      }
    }
    return items.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  }, [sales, deliveryData, section, tab]);
  const listLoading = (section === 'customer' && isLoading)
    || (section === 'partner' && isDeliveriesLoading)
    || (section === 'all' && (isLoading || isDeliveriesLoading));

  const handleTabChange = (t: SalesTab) => { setTab(t); setPage(1); };
  const toggleCheck = (id: string) => {
    setCheckedIds((prev) => { const next = new Set(prev); if (next.has(id)) next.delete(id); else next.add(id); return next; });
  };
  const handleChannelChange = (c: SalesChannel | 'b2b') => { setChannel(c); setPage(1); };
  const handleDateRangeChange = (d: SalesDateRange) => { setDateRange(d); setPage(1); };

  /** 영역 칩 변경 — 채널 칩의 첫 옵션으로 자동 매핑 (옵션 셋이 달라지므로 stale 방지) */
  const handleSectionChange = (s: SalesSection) => {
    setSection(s);
    setChannel(CHANNELS_BY_SECTION[s][0].key);
    setPage(1);
  };

  /** 거래처 매출 입력 — Phase C 에서 /sales/new?mode=b2b 라우팅 예정. 현재는 안내. */
  const handlePartnerSaleClick = () => {
    if (confirm('거래처 매출 입력은 현재 [B2B거래] 메뉴(/deliveries)에서 진행해주세요. 이동하시겠습니까?')) {
      router.push('/deliveries');
    }
  };

  /** 거래처(B2B) 카드 합산값 — offline_sales (dealer/academy) + deliveries 매출 */
  const partnerWeek = (stats?.partnerWeek?.amount || 0) + (deliveryStats?.weekAmount || 0);
  const partnerWeekCount = (stats?.partnerWeek?.count || 0) + (deliveryStats?.weekCount || 0);
  const partnerMonth = (stats?.partnerMonth?.amount || 0) + (deliveryStats?.monthAmount || 0);
  const partnerMonthCount = (stats?.partnerMonth?.count || 0) + (deliveryStats?.monthCount || 0);
  const partnerOutstanding = (stats?.partnerOutstanding || 0) + (deliveryStats?.outstanding || 0);

  /* --- 목록 영역 (좌측/모바일) --- */
  const listContent = (
    <>
      {/* 통계 요약 카드 — 2026-05-26: 고객(B2C) / 거래처(B2B) 2섹션 동등 비율 (사장님 A-1) */}
      {stats && (
        <div className="grid grid-cols-2 gap-2">
          {/* 고객 (B2C) 섹션 */}
          <div className="bg-white rounded-lg border border-neutral-200 p-3 space-y-2">
            <div className="text-[11px] font-semibold text-neutral-500 uppercase tracking-wider">고객 (B2C)</div>
            <div className="grid grid-cols-3 gap-1.5">
              <div>
                <div className="flex items-center gap-1 text-[10px] text-neutral-400 mb-0.5"><Calendar size={10} />이번주</div>
                <p className="text-sm font-bold text-neutral-900">{formatKRW(stats.customerWeek?.amount || 0)}</p>
                <p className="text-[10px] text-neutral-400">{stats.customerWeek?.count || 0}건</p>
              </div>
              <div>
                <div className="flex items-center gap-1 text-[10px] text-neutral-400 mb-0.5"><TrendingUp size={10} />이번달</div>
                <p className="text-sm font-bold text-neutral-900">{formatKRW(stats.customerMonth?.amount || 0)}</p>
                <p className="text-[10px] text-neutral-400">{stats.customerMonth?.count || 0}건</p>
              </div>
              <div>
                <div className="flex items-center gap-1 text-[10px] text-red-400 mb-0.5"><AlertCircle size={10} />미수금</div>
                <p className="text-sm font-bold text-red-600">{formatKRW(stats.customerOutstanding || 0)}</p>
              </div>
            </div>
          </div>

          {/* 거래처 (B2B) 섹션 */}
          <div className="bg-white rounded-lg border border-neutral-200 p-3 space-y-2">
            <div className="text-[11px] font-semibold text-neutral-500 uppercase tracking-wider">거래처 (B2B)</div>
            <div className="grid grid-cols-3 gap-1.5">
              <div>
                <div className="flex items-center gap-1 text-[10px] text-neutral-400 mb-0.5"><Calendar size={10} />이번주</div>
                <p className="text-sm font-bold text-neutral-900">{formatKRW(partnerWeek)}</p>
                <p className="text-[10px] text-neutral-400">{partnerWeekCount}건</p>
              </div>
              <div>
                <div className="flex items-center gap-1 text-[10px] text-neutral-400 mb-0.5"><TrendingUp size={10} />이번달</div>
                <p className="text-sm font-bold text-neutral-900">{formatKRW(partnerMonth)}</p>
                <p className="text-[10px] text-neutral-400">{partnerMonthCount}건</p>
              </div>
              <div>
                <div className="flex items-center gap-1 text-[10px] text-red-400 mb-0.5"><AlertCircle size={10} />미수금</div>
                <p className="text-sm font-bold text-red-600">{formatKRW(partnerOutstanding)}</p>
              </div>
            </div>
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

      {/* 영역 칩 (★ 2026-05-26 신규) — 고객(B2C) / 거래처(B2B) / 전체. default 고객 */}
      <div className="flex flex-wrap items-center gap-3">
        <div className="flex gap-1">
          {SECTIONS.map((s) => (
            <button
              key={s.key}
              onClick={() => handleSectionChange(s.key)}
              className={`px-3 py-1 text-xs rounded-full border transition font-medium ${
                section === s.key
                  ? 'bg-neutral-900 text-white border-neutral-900'
                  : 'bg-white text-neutral-600 border-neutral-300 hover:border-neutral-500'
              }`}
            >
              {s.label}
            </button>
          ))}
        </div>
      </div>

      {/* 채널 + 기간 필터 (한 줄 통합) */}
      <div className="flex flex-wrap items-center gap-3">
        {/* 채널 — 영역 칩 따라 동적 표시 */}
        <div className="flex gap-1">
          {CHANNELS_BY_SECTION[section].map((c) => (
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

      {/* 판매 목록 (2026-05-26 Phase B: sale + delivery 통합) */}
      <Card padding={false}>
        {listLoading ? (
          <div className="p-4 space-y-3">
            {Array.from({ length: 5 }).map((_, i) => (
              <Skeleton key={i} className="h-14 w-full" />
            ))}
          </div>
        ) : unifiedItems.length === 0 ? (
          <EmptyState icon={Receipt} message="판매/거래 기록이 없습니다" />
        ) : (
          <div className="divide-y divide-neutral-100">
            {unifiedItems.map((item) => {
              if (item.sourceType === 'sale') {
                return (
                  <SaleRow
                    key={`sale-${item.id}`}
                    sale={item.data}
                    selected={selected?.sourceType === 'sale' && selected.id === item.id}
                    onClick={() => setSelected({ id: item.id, sourceType: 'sale' })}
                    prepMode={prepMode}
                    checked={checkedIds.has(item.id)}
                    onCheck={() => toggleCheck(item.id)}
                  />
                );
              }
              return (
                <DeliveryRow
                  key={`delivery-${item.id}`}
                  delivery={item.data}
                  selected={selected?.sourceType === 'delivery' && selected.id === item.id}
                  onClick={() => setSelected({ id: item.id, sourceType: 'delivery' })}
                />
              );
            })}
          </div>
        )}
      </Card>

      {/* 페이지네이션 — 영역 'customer' 일 때만 (sales 만 서버 페이지네이션). partner/all 은 limit 30 안에서 표시 */}
      {section === 'customer' && (
        <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} />
      )}
      {section !== 'customer' && unifiedItems.length >= 30 && (
        <p className="text-xs text-neutral-400 text-center py-2">최대 30건까지 표시됩니다. 더 자세히는 [B2B거래] 메뉴에서 확인하세요.</p>
      )}
    </>
  );

  return (
    <>
      <Topbar
        title="판매 관리"
        action={
          <div className="flex gap-2">
            <Button onClick={() => router.push('/sales/new')} size="sm">
              <Plus size={14} />
              판매 입력
            </Button>
            {/* 2026-05-26 IA 통합: 거래처 매출 (B2B). Phase C 에서 /sales/new?mode=b2b 라우팅으로 교체 예정 */}
            <Button onClick={handlePartnerSaleClick} size="sm" variant="secondary">
              <Plus size={14} />
              거래처 매출
            </Button>
          </div>
        }
      />

      {isLg ? (
        /* PC: 마스터-디테일 2컬럼 */
        <div className="flex gap-4 px-4 md:px-6 py-4 h-full min-h-0">
          {/* 좌측: 목록 */}
          <div className="w-2/5 shrink-0 overflow-y-auto space-y-3 pr-1">
            {listContent}
          </div>
          {/* 우측: 상세 패널 — 2026-05-26 Phase B: sourceType 분기 */}
          <div className="flex-1 min-w-0 overflow-y-auto bg-white rounded-xl border border-neutral-200">
            {selected ? (
              selected.sourceType === 'sale' ? (
                <SaleDetailPanel saleId={selected.id} />
              ) : (
                <DeliveryDetailPanel deliveryId={selected.id} />
              )
            ) : (
              <div className="flex items-center justify-center h-full text-sm text-neutral-400">
                좌측 목록에서 판매/거래 건을 선택하세요
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

      {/* 모바일 상세 패널 — 2026-05-26 Phase B: sourceType 분기 */}
      {!isLg && (
        <SlidePanel
          open={!!selected}
          onClose={() => setSelected(null)}
          title={selected?.sourceType === 'delivery' ? '거래처 납품 상세' : '판매 상세'}
          className="sm:w-[480px]"
        >
          {selected && selected.sourceType === 'sale' && <SaleDetailPanel saleId={selected.id} />}
          {selected && selected.sourceType === 'delivery' && <DeliveryDetailPanel deliveryId={selected.id} />}
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
          {/* 운영 상태 칩 (2026-05-25): 판매완료 / 배송중 (취소는 위에서 처리) */}
          {!isCancelled && sale.delivered_at && (
            <Badge className="bg-neutral-100 text-neutral-600">판매완료</Badge>
          )}
          {!isCancelled && !sale.delivered_at && sale.shipped_at && (
            <Badge className="bg-green-100 text-green-700">배송중</Badge>
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

/* 거래처 납품 행 — 2026-05-26 Phase B 신규 */
const DELIVERY_STATUS_LABEL: Record<string, string> = {
  draft: '작성중', confirmed: '납품확정', shipped: '출고완료', settled: '정산완료',
};
const DELIVERY_STATUS_COLOR: Record<string, string> = {
  draft: 'bg-neutral-100 text-neutral-600',
  confirmed: 'bg-blue-100 text-blue-700',
  shipped: 'bg-green-100 text-green-700',
  settled: 'bg-emerald-100 text-emerald-700',
};
const DELIVERY_PAYMENT_LABEL: Record<string, string> = { unpaid: '미결제', partial: '부분결제', paid: '결제완료' };
const DELIVERY_PAYMENT_COLOR: Record<string, string> = {
  unpaid: 'bg-red-100 text-red-700',
  partial: 'bg-yellow-100 text-yellow-700',
  paid: 'bg-green-100 text-green-700',
};
const CUSTOMER_TYPE_LABEL: Record<string, string> = { dealer: '딜러', academy: '아카데미' };

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type DeliveryRowData = any;

const DeliveryRow = memo(function DeliveryRow({ delivery, selected, onClick }: {
  delivery: DeliveryRowData; selected?: boolean; onClick: () => void;
}) {
  const isCancelled = !!delivery.cancelled_at;
  const ctypeLabel = CUSTOMER_TYPE_LABEL[delivery.customer_type as string] || delivery.customer_type || '';

  return (
    <div
      onClick={onClick}
      className={`flex items-center gap-4 px-4 py-3 cursor-pointer transition ${
        selected ? 'bg-neutral-50 border-l-2 border-l-neutral-900' : 'hover:bg-warm-ivory/60'
      } ${isCancelled ? 'opacity-50' : ''}`}
    >
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2 flex-wrap">
          <span className={`text-sm font-semibold text-indigo-black truncate ${isCancelled ? 'line-through' : ''}`}>
            {delivery.customer_name || delivery.company_name || '미지정'}
          </span>
          {isCancelled ? (
            <Badge className="bg-neutral-200 text-neutral-500">취소</Badge>
          ) : (
            <Badge className={DELIVERY_PAYMENT_COLOR[delivery.payment_status as string] || ''}>
              {DELIVERY_PAYMENT_LABEL[delivery.payment_status as string] || delivery.payment_status}
            </Badge>
          )}
          {/* 납품 상태 (draft/confirmed/shipped/settled) */}
          {!isCancelled && delivery.status && (
            <Badge className={DELIVERY_STATUS_COLOR[delivery.status as string] || ''}>
              {DELIVERY_STATUS_LABEL[delivery.status as string] || delivery.status}
            </Badge>
          )}
          {/* sourceType 뱃지 — 거래처(B2B) 구분 */}
          <Badge className="bg-indigo-50 text-indigo-700 border border-indigo-200">거래처</Badge>
          {ctypeLabel && (
            <span className="text-[10px] text-neutral-500">{ctypeLabel}</span>
          )}
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          <span>{formatDate(delivery.delivery_date || delivery.expected_date || '')}</span>
          <span>{delivery.dl_number || ''}</span>
        </div>
      </div>
      <div className="text-right shrink-0">
        <span className={`text-sm font-bold ${isCancelled ? 'line-through text-neutral-400' : ''}`}>
          {formatKRW(((delivery.total_amount as number) || 0) - ((delivery.discount_amount as number) || 0))}
        </span>
      </div>
    </div>
  );
});
