'use client';

import { useState, memo, useEffect, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { SlidePanel } from '@/components/ui/slide-panel';
import { SaleDetailPanel } from '@/components/sales/sale-detail-panel';
import { DeliveryDetailPanel } from '@/components/deliveries/delivery-detail-panel';
import { useSales, useSalesTabCounts, useSalesStats, useMarkSalePacked } from '@/hooks/use-sales';
import type { SalesTab, SalesChannel, SalesDateRange } from '@/hooks/use-sales';
import { useDeliveryStats, useDeliveries } from '@/hooks/use-deliveries';
import { useContracts } from '@/hooks/use-contracts';
import { formatKRW, formatDate, SALE_CHANNEL_LABEL, CUSTOMER_TYPE_LABEL } from '@/lib/utils/format';
import { getSaleShipStatus, getDeliveryShipStatus, type ShipStatus } from '@/lib/sales/ship-status';
import { useActivityTypes, type ActivityTypes } from '@/hooks/use-activity-types';
import { ActivityChips } from '@/components/shared/activity-chips';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Pagination } from '@/components/ui/pagination';
import { Plus, FileSignature, Receipt, ClipboardList, LayoutGrid, Package } from 'lucide-react';
import { PrepSheetModal } from '@/components/sales/prep-sheet-modal';
import { RevenueDarkCard } from '@/components/ui/revenue-dark-card';
import type { OfflineSale } from '@/lib/supabase/types';
import { NotePreview } from '@/components/shared/customer-notes';
import { useLatestCustomerNotes, type LatestNote } from '@/hooks/use-customer-notes';

const PAYMENT_METHOD_LABEL: Record<string, string> = {
  card: '카드',
  cash: '현금',
  transfer: '계좌이체',
  mixed: '복합',
};

// 채널 라벨은 공용 SALE_CHANNEL_LABEL(format.ts) 사용 — 전 화면 통일

// 고객유형 라벨은 format.ts SSOT(CUSTOMER_TYPE_LABEL) 사용

/** 2026-05-26 Phase G-4: 안 A 상태 분류 — 좌측 색 줄 + 우측 도트 결정 */
type RowState = 'paid_done' | 'paid_shipping' | 'paid_packed' | 'paid_wait_ship' | 'paid_unhandled' | 'unpaid' | 'partial' | 'shipped_b2b_unpaid' | 'wait_pickup_unpaid' | 'cancelled' | 'returned';

function getRowStateSale(s: OfflineSale): RowState {
  if (s.cancelled_at) return 'cancelled';
  if (s.returned_at) return 'returned';   // 반품 = 종결 상태(결제·배송 상태보다 우선) — '처리 대기'로 안 떨어지게
  if (s.payment_status === 'partial') return 'partial';
  if (s.payment_status === 'unpaid') return 'unpaid';
  // payment_status === 'paid' — 초록 = 사장님이 더 할 일 없는 상태
  if (s.delivered_at) return 'paid_done';            // 배송완료/고객수령 완료 = 끝
  if (s.shipped_at) return 'paid_shipping';          // 출고완료(배송중) — 택배사가 배송, 내 할 일 끝
  // 2026-07-18: 포장까지 끝낸 건 = '준비완료'. 송장 유무와 무관(포장은 송장 전에도 함) — 사장님 확정
  if (s.packed_at) return 'paid_packed';             // 포장 완료 — 내 할 일 끝, 기사님만 오면 됨
  if (s.invoice_number) return 'paid_wait_ship';     // 송장 발급됨, 포장 전 — 준비 필요
  return 'paid_unhandled';                            // 송장X + 수령X — 출고/수령 처리 필요
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function getRowStateDelivery(d: any): RowState {
  if (d.cancelled_at) return 'cancelled';
  if (d.payment_status === 'partial') return 'partial';
  // 110: '출고' = 기사님이 실제로 수거해 감 (집하). 송장만 발급된 건은 아직 출고 전이다.
  //      (전엔 송장 발급 즉시 status='shipped' 라 여기서 '출고완료' 로 잘못 떴다)
  const isShipped = d.status === 'shipped' || d.status === 'settled';
  if (d.payment_status === 'unpaid' && isShipped) return 'shipped_b2b_unpaid'; // 출고됐는데 결제 대기
  // 110: 송장은 발급됐지만 아직 기사님이 안 가져간 건 — '미결제'로만 뭉개면 출고대기 정보가 사라진다
  if (d.payment_status === 'unpaid' && d.tracking_number) return 'wait_pickup_unpaid';
  if (d.payment_status === 'unpaid') return 'unpaid';
  // payment_status === 'paid'
  if (d.delivered_at) return 'paid_done';             // 110: 배송완료 = 끝 (B2C 와 동일 기준)
  if (isShipped) return 'paid_done';                  // 출고완료 + 결제완료 = 판매완료
  return 'paid_wait_ship';                            // 결제완료 + 출고 전 (송장 발급됐으면 '출고 대기')
}

function stripColor(state: RowState): string {
  switch (state) {
    case 'paid_done': return 'bg-green-500';
    case 'paid_shipping': return 'bg-green-500';  // 출고완료(배송중)도 초록 — 사장님 할 일 끝
    case 'paid_packed': return 'bg-emerald-400';  // 준비완료 — 포장 끝, 기사님만 오면 됨
    case 'unpaid': return 'bg-red-500';
    case 'partial': return 'bg-yellow-400';
    case 'cancelled': return 'bg-neutral-300';
    case 'returned': return 'bg-violet-300';      // 반품 = 종결(취소와 구분: 보라)
    default: return 'bg-transparent';             // paid_wait_ship, paid_unhandled — 할 일 있음
  }
}

function rightDot(state: RowState): { color: string; title: string } | null {
  switch (state) {
    case 'paid_shipping': return { color: 'bg-green-500', title: '배송중' };
    case 'paid_packed': return { color: 'bg-emerald-500', title: '준비완료 — 포장 끝, 기사님 대기' };
    case 'paid_wait_ship': return { color: 'bg-amber-400', title: '출고 대기' };
    case 'paid_unhandled': return { color: 'bg-amber-400', title: '처리 대기' };
    case 'shipped_b2b_unpaid': return { color: 'bg-amber-400', title: '출고완료 · 결제 대기' };
    case 'wait_pickup_unpaid': return { color: 'bg-amber-400', title: '출고대기 · 결제 대기' };
    default: return null;
  }
}

function statusLabel(state: RowState): string {
  switch (state) {
    case 'paid_done': return '판매완료';
    case 'paid_shipping': return '배송중';
    case 'paid_packed': return '준비완료';
    case 'paid_wait_ship': return '출고 대기';
    case 'paid_unhandled': return '처리 대기';
    case 'unpaid': return '미결제';
    case 'partial': return '부분결제';
    case 'shipped_b2b_unpaid': return '출고완료 · 결제대기';
    case 'wait_pickup_unpaid': return '출고대기 · 결제대기';
    case 'cancelled': return '취소';
    case 'returned': return '반품';
  }
}

/** 그리드 '상태' 컬럼 전용 — '출고대기'는 별도 '배송상태' 컬럼이 담당하므로 상태엔 '준비중'으로 표기 */
function gridStatusLabel(state: RowState): string {
  if (state === 'paid_wait_ship') return '준비중';
  if (state === 'wait_pickup_unpaid') return '준비중 · 결제대기';
  return statusLabel(state);
}

function statusTextClass(state: RowState): string {
  if (state === 'unpaid') return 'text-red-600 font-medium';
  if (state === 'partial') return 'text-yellow-700 font-medium';
  if (state === 'paid_done' || state === 'paid_shipping') return 'text-green-700 font-medium';
  if (state === 'paid_packed') return 'text-emerald-600 font-medium';   // 준비완료 = 내 할 일 끝
  if (state === 'returned') return 'text-violet-600 font-medium';       // 반품
  return 'text-neutral-500';
}

const TABS: { key: SalesTab; label: string }[] = [
  { key: 'all', label: '전체' },
  { key: 'processing', label: '처리 필요' },
  { key: 'unpaid', label: '미수금' },
  { key: 'cancelled', label: '취소·반품' },
];

/** 2026-05-26 IA 통합: 영역 칩 — 고객(B2C) / 거래처(B2B) / 전체. default 고객 */
type SalesSection = 'customer' | 'partner' | 'all';
const SECTIONS: { key: SalesSection; label: string }[] = [
  { key: 'customer', label: '고객' },
  { key: 'partner', label: '거래처' },
  { key: 'all', label: '전체' },
];

/** 2026-05-26 Phase G-2: 채널 칩 폐기 — 영역 칩만으로 데이터 분기 충분
 *  - section='customer' → useSales(channel='all') + customer_type IN (소매/온라인/상담) 자동
 *  - section='partner'  → useDeliveries + offline_sales 의 dealer/academy 합집합 (useMemo 안에서)
 *  - section='all'      → 둘 다 합집합
 */

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
  // 2026-05-26 Phase G-2: 채널 칩 폐기 — channel state 는 hook 호환용으로 'all' 고정
  const channel: SalesChannel = 'all';
  const [dateRange, setDateRange] = useState<SalesDateRange>('month');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  /** 2026-05-26 Phase B: 우측 상세 패널 — sourceType 분기. 'sale' → SaleDetailPanel / 'delivery' → DeliveryDetailPanel */
  const [selected, setSelected] = useState<{ id: string; sourceType: 'sale' | 'delivery' } | null>(null);
  const [isLg, setIsLg] = useState(false);
  const [prepMode, setPrepMode] = useState(false);
  const markPacked = useMarkSalePacked();   // 일괄 포장완료(준비완료) — 2026-07-18
  // 2026-07-14 PC 그리드(베타) — 기본 OFF. isLg 에서만 토글 노출. localStorage 로 선호 기억
  const [gridMode, setGridMode] = useState(false);
  const toggleGrid = () => setGridMode((v) => { const n = !v; try { localStorage.setItem('sales-pc-grid', n ? '1' : '0'); } catch { /* noop */ } return n; });
  const [checkedIds, setCheckedIds] = useState<Set<string>>(new Set());
  const [checkedDeliveryIds, setCheckedDeliveryIds] = useState<Set<string>>(new Set()); // 2026-05-26 Phase D: 거래처 납품 준비표 통합
  const [showPrepSheet, setShowPrepSheet] = useState(false);

  useEffect(() => {
    const mq = window.matchMedia('(min-width: 1024px)');
    setIsLg(mq.matches);
    // 2026-07-14: PC 그리드 선호 복원 (SSR 안전하게 mount 후 읽음 — 기존 isLg effect 에 합침)
    try { setGridMode(localStorage.getItem('sales-pc-grid') === '1'); } catch { /* noop */ }
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
        // 처리 필요 = 결제완료인데 출고 전(거래처)
        if (tab === 'processing' && !(d.payment_status === 'paid' && !['shipped', 'settled'].includes(d.status))) continue;
        items.push({ sourceType: 'delivery', id: d.id, date: d.delivery_date || d.created_at || '', data: d });
      }
    }
    return items.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  }, [sales, deliveryData, section, tab]);
  const { data: noteMap } = useLatestCustomerNotes(
    unifiedItems.map((i) => (i.sourceType === 'sale' ? i.data.customer_id : null))
  );
  const listLoading = (section === 'customer' && isLoading)
    || (section === 'partner' && isDeliveriesLoading)
    || (section === 'all' && (isLoading || isDeliveriesLoading));

  const handleTabChange = (t: SalesTab) => { setTab(t); setPage(1); };
  const toggleCheck = (id: string) => {
    setCheckedIds((prev) => { const next = new Set(prev); if (next.has(id)) next.delete(id); else next.add(id); return next; });
  };
  const toggleCheckDelivery = (id: string) => {
    setCheckedDeliveryIds((prev) => { const next = new Set(prev); if (next.has(id)) next.delete(id); else next.add(id); return next; });
  };
  const totalChecked = checkedIds.size + checkedDeliveryIds.size;
  const handleDateRangeChange = (d: SalesDateRange) => { setDateRange(d); setPage(1); };

  /** 영역 칩 변경 — 데이터 분기는 unifiedItems useMemo 가 처리 */
  const handleSectionChange = (s: SalesSection) => {
    setSection(s);
    setPage(1);
  };

  /** 2026-05-26 Phase C: 거래처 매출 입력 → /sales/new?mode=b2b 라우팅 (CreateDeliveryModal 풀스크린 표시) */
  const handlePartnerSaleClick = () => {
    router.push('/sales/new?mode=b2b');
  };

  /** 거래처(B2B) 카드 합산값 — offline_sales (dealer/academy) + deliveries 매출 */
  const partnerWeek = (stats?.partnerWeek?.amount || 0) + (deliveryStats?.weekAmount || 0);
  const partnerWeekCount = (stats?.partnerWeek?.count || 0) + (deliveryStats?.weekCount || 0);
  const partnerMonth = (stats?.partnerMonth?.amount || 0) + (deliveryStats?.monthAmount || 0);
  const partnerMonthCount = (stats?.partnerMonth?.count || 0) + (deliveryStats?.monthCount || 0);
  const partnerOutstanding = (stats?.partnerOutstanding || 0) + (deliveryStats?.outstanding || 0);

  /** 2026-06-18 제품/복원수리 분리표기 — 헤드라인=제품(RS 제외, 대시보드 정의와 일치), 보조표기=복원수리. 제품+복원수리 = 리스트 합 */
  const customerMonthRepair = stats?.customerMonthRepair || 0;
  const customerMonthProduct = (stats?.customerMonth?.amount || 0) - customerMonthRepair;
  const partnerMonthRepair = (stats?.partnerMonthRepair || 0) + (deliveryStats?.monthRepair || 0);
  const partnerMonthProduct = partnerMonth - partnerMonthRepair;

  // 활동유형 칩 — 보이는 항목들의 고객 전화로 배치 조회
  const salesActTypes = useActivityTypes(
    unifiedItems.map((i) => (i.data as { customer_phone?: string | null }).customer_phone)
  );

  /* --- 목록 영역 (좌측/모바일) --- */
  const listContent = (
    <>
      {/* 통계 요약 카드 — 2026-05-26 Phase G-4 안 3 + 2026-05-27 공통 RevenueDarkCard 추출 */}
      {stats && (
        <div className="grid grid-cols-2 gap-3">
          <RevenueDarkCard
            label="고객 (B2C) 제품"
            amount={formatKRW(customerMonthProduct)}
            amountSub={`이번달 · ${stats.customerMonth?.count || 0}건`}
            subNote={customerMonthRepair > 0 ? `+ 복원수리 ${formatKRW(customerMonthRepair)}` : undefined}
            bottomGrid={[
              { label: '이번주', value: formatKRW(stats.customerWeek?.amount || 0) },
              { label: '미수금', value: formatKRW(stats.customerOutstanding || 0), highlight: 'amber' },
            ]}
          />
          <RevenueDarkCard
            label="거래처 (B2B) 제품"
            amount={formatKRW(partnerMonthProduct)}
            amountSub={`이번달 · ${partnerMonthCount}건`}
            subNote={partnerMonthRepair > 0 ? `+ 복원수리 ${formatKRW(partnerMonthRepair)}` : undefined}
            bottomGrid={[
              { label: '이번주', value: formatKRW(partnerWeek) },
              { label: '미수금', value: formatKRW(partnerOutstanding), highlight: 'amber' },
            ]}
          />
        </div>
      )}

      {/* B2B 거래처별 매출 (영역 '거래처' 일 때만) — 2026-05-26 Phase G-2: channel='b2b' → section='partner' 매핑 */}
      {section === 'partner' && stats?.b2b && stats.b2b.length > 0 && (
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
          onClick={() => { setPrepMode(!prepMode); setCheckedIds(new Set()); setCheckedDeliveryIds(new Set()); }}
          className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium transition shrink-0 ${
            prepMode ? 'bg-neutral-900 text-white' : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'
          }`}
        >
          <ClipboardList size={14} />
          {prepMode ? '선택 취소' : '준비표 뽑기'}
        </button>
        {prepMode && totalChecked > 0 && (
          <button
            onClick={() => setShowPrepSheet(true)}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-blue-600 text-white text-xs font-medium hover:bg-blue-700 transition shrink-0"
          >
            <Receipt size={14} />
            준비표 인쇄 ({totalChecked}건)
          </button>
        )}
        {/* 일괄 포장완료 — 2026-07-18. 판매 건만 대상(B2B 납품은 별도 흐름) */}
        {prepMode && checkedIds.size > 0 && (
          <button
            onClick={() => { markPacked.mutate({ ids: [...checkedIds], packed: true }); setCheckedIds(new Set()); }}
            disabled={markPacked.isPending}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-600 text-white text-xs font-medium hover:bg-emerald-700 transition shrink-0 disabled:opacity-50"
          >
            <Package size={14} />
            {markPacked.isPending ? '처리 중…' : `포장완료 (${checkedIds.size}건)`}
          </button>
        )}
        <SearchInput
          value={search}
          onChange={(v) => { setSearch(v); setPage(1); }}
          placeholder="판매번호, 고객명, 전화번호"
        />
      </div>

      {/* 탭 바 — 상태 필터 */}
      <div className="flex items-end gap-1 border-b border-neutral-200">
        <span className="text-[11px] font-semibold text-neutral-400 pb-2 pr-1 shrink-0">상태</span>
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
              {count > 0 && (t.key === 'unpaid' || t.key === 'processing') && (
                <span className={`ml-1.5 text-xs px-1.5 py-0.5 rounded-full ${
                  t.key === 'processing'
                    ? (active ? 'bg-amber-500 text-white' : 'bg-amber-100 text-amber-700')
                    : (active ? 'bg-neutral-900 text-white' : 'bg-neutral-200 text-neutral-500')
                }`}>
                  {count}
                </span>
              )}
            </button>
          );
        })}
      </div>

      {/* 영역 칩 + 기간 필터 (한 줄 통합) — 2026-05-26 Phase G-2: 채널 칩 폐기 */}
      <div className="flex flex-wrap items-center gap-3">
        {/* 대상 — 고객(B2C) / 거래처(B2B) / 전체. default 고객 */}
        <div className="flex items-center gap-1">
          <span className="text-[11px] font-semibold text-neutral-400 pr-0.5">대상</span>
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

        {/* 구분선 */}
        <div className="w-px h-4 bg-neutral-200" />

        {/* 기간 */}
        <div className="flex flex-wrap gap-1 items-center">
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
          <div className="flex flex-wrap items-center gap-1 ml-1">
            <input
              type="date"
              value={dateFrom}
              onChange={(e) => { setDateFrom(e.target.value); setDateRange('all'); setPage(1); }}
              className="h-7 px-2 rounded-md border border-neutral-200 text-xs bg-white min-w-0 max-w-[42vw] sm:max-w-none"
            />
            <span className="text-xs text-neutral-400">~</span>
            <input
              type="date"
              value={dateTo}
              onChange={(e) => { setDateTo(e.target.value); setDateRange('all'); setPage(1); }}
              className="h-7 px-2 rounded-md border border-neutral-200 text-xs bg-white min-w-0 max-w-[42vw] sm:max-w-none"
            />
            {(dateFrom || dateTo) && (
              <button onClick={() => { setDateFrom(''); setDateTo(''); }} className="text-[10px] text-neutral-400 hover:text-neutral-600">초기화</button>
            )}
          </div>
        </div>
      </div>

      {/* 색 범례 — 좌측 줄/도트 의미 */}
      <div className="flex flex-wrap items-center gap-x-3 gap-y-1 text-[11px] text-neutral-500 px-1">
        <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm bg-green-500" />완료(출고·수령)</span>
        <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-amber-400" />처리 대기</span>
        <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm bg-red-500" />미결제</span>
        <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm bg-yellow-400" />부분결제</span>
        <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm bg-neutral-300" />취소</span>
        <span className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm bg-violet-300" />반품</span>
      </div>

      {/* 판매 목록 (2026-05-26 Phase B: sale + delivery 통합) — gridMode 면 밀집 표, 아니면 카드 */}
      <Card padding={false}>
        {listLoading ? (
          <div className="p-4 space-y-3">
            {Array.from({ length: 5 }).map((_, i) => (
              <Skeleton key={i} className="h-14 w-full" />
            ))}
          </div>
        ) : unifiedItems.length === 0 ? (
          <EmptyState icon={Receipt} message="판매/거래 기록이 없습니다" />
        ) : gridMode && isLg ? (
          <div className="overflow-x-auto">
            <SalesGridTable
              items={unifiedItems}
              selectedId={selected?.id}
              onSelect={(id, type) => setSelected({ id, sourceType: type })}
              prepMode={prepMode}
              checkedSale={checkedIds}
              checkedDelivery={checkedDeliveryIds}
              onCheckSale={toggleCheck}
              onCheckDelivery={toggleCheckDelivery}
              actLookup={salesActTypes}
            />
          </div>
        ) : (
          <div className="divide-y divide-neutral-100">
            {unifiedItems.map((item) => {
              if (item.sourceType === 'sale') {
                return (
                  <SaleRow
                    key={`sale-${item.id}`}
                    sale={item.data}
                    note={item.data.customer_id ? noteMap?.[item.data.customer_id] : undefined}
                    selected={selected?.sourceType === 'sale' && selected.id === item.id}
                    onClick={() => setSelected({ id: item.id, sourceType: 'sale' })}
                    prepMode={prepMode}
                    checked={checkedIds.has(item.id)}
                    onCheck={() => toggleCheck(item.id)}
                    actTypes={salesActTypes(item.data.customer_phone)}
                  />
                );
              }
              return (
                <DeliveryRow
                  key={`delivery-${item.id}`}
                  delivery={item.data}
                  selected={selected?.sourceType === 'delivery' && selected.id === item.id}
                  onClick={() => setSelected({ id: item.id, sourceType: 'delivery' })}
                  prepMode={prepMode}
                  checked={checkedDeliveryIds.has(item.id)}
                  onCheck={() => toggleCheckDelivery(item.id)}
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
        <p className="text-xs text-neutral-400 text-center py-2">최대 30건까지 표시됩니다. 더 자세히는 <button onClick={() => router.push('/deliveries')} className="underline font-semibold text-neutral-500 hover:text-neutral-700">B2B 납품</button> 메뉴에서 확인하세요.</p>
      )}
    </>
  );

  return (
    <>
      <Topbar
        title="판매 관리"
        action={
          <div className="flex gap-2">
            {/* 2026-07-14 PC 그리드(베타) 토글 — PC 에서만 */}
            {isLg && (
              <Button onClick={toggleGrid} size="sm" variant="secondary" title="목록을 밀집 표로 보기">
                <LayoutGrid size={14} />
                {gridMode ? '카드 보기' : 'PC 그리드'}
              </Button>
            )}
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
        /* PC: 마스터-디테일 2컬럼. gridMode 면 목록(표)을 넓게, 상세를 고정폭으로 뒤집는다 */
        <div className="flex gap-4 px-4 md:px-6 py-4 h-full min-h-0">
          {/* 좌측: 목록 */}
          <div className={`${gridMode ? 'flex-1 min-w-0' : 'w-2/5 shrink-0'} overflow-y-auto space-y-3 pr-1`}>
            {listContent}
          </div>
          {/* 우측: 상세 패널 — 2026-05-26 Phase B: sourceType 분기 */}
          <div className={`${gridMode ? 'w-[420px] shrink-0' : 'flex-1 min-w-0'} overflow-y-auto bg-white rounded-xl border border-neutral-200`}>
            {selected ? (
              selected.sourceType === 'sale' ? (
                <SaleDetailPanel saleId={selected.id} />
              ) : (
                <DeliveryDetailPanel deliveryId={selected.id} />
              )
            ) : (
              <div className="flex items-center justify-center h-full text-sm text-neutral-400">
                {gridMode ? '표에서 건을 선택하세요' : '좌측 목록에서 판매/거래 건을 선택하세요'}
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

      {/* 준비표 모달 — 복수 건 (2026-05-26 Phase D: B2C + B2B 통합) */}
      {showPrepSheet && (
        <PrepSheetModal
          saleIds={Array.from(checkedIds)}
          deliveryIds={Array.from(checkedDeliveryIds)}
          onClose={() => setShowPrepSheet(false)}
        />
      )}
    </>
  );
}

/* 목록 행 — 2026-05-26 Phase G-4: 안 A (좌측 색 줄 + 우측 도트 + 판매완료 라벨 통일) */
/* ─────────────────────────────────────────────────────────────
   2026-07-14 PC-Native 밀집 그리드 (베타) — 카드 대신 표.
   ⚠️ 데이터·상태·핸들러는 카드 뷰와 100% 동일한 것을 재사용한다(기능 드리프트 0).
      행 상태 헬퍼(getRowStateSale/Delivery·statusLabel·rightDot)·라벨 상수 그대로.
   ───────────────────────────────────────────────────────────── */
type UnifiedItem =
  | { sourceType: 'sale'; id: string; date: string; data: OfflineSale }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  | { sourceType: 'delivery'; id: string; date: string; data: any };

/** 배송상태 tone → 컴팩트 배지 (주문관리 배지 스타일과 통일) */
const SHIP_BADGE: Record<ShipStatus['tone'], string> = {
  amber: 'bg-amber-50 text-amber-700',
  green: 'bg-emerald-50 text-emerald-700',
  mute: 'text-neutral-300',
};

const SalesGridTable = memo(function SalesGridTable({
  items, selectedId, onSelect, prepMode, checkedSale, checkedDelivery, onCheckSale, onCheckDelivery, actLookup,
}: {
  items: UnifiedItem[];
  selectedId?: string;
  onSelect: (id: string, type: 'sale' | 'delivery') => void;
  prepMode: boolean;
  checkedSale: Set<string>;
  checkedDelivery: Set<string>;
  onCheckSale: (id: string) => void;
  onCheckDelivery: (id: string) => void;
  actLookup?: (phone?: string | null) => ActivityTypes | undefined;
}) {
  return (
    <table className="w-full text-sm border-collapse">
      <thead>
        <tr className="sticky top-0 bg-stone-50 z-[1] text-left text-[11px] font-semibold text-neutral-500">
          {prepMode && <th className="px-2 py-2.5 w-9"></th>}
          <th className="px-3 py-2.5 whitespace-nowrap">날짜</th>
          <th className="px-3 py-2.5">고객</th>
          <th className="px-3 py-2.5 whitespace-nowrap">상담유형</th>
          <th className="px-3 py-2.5 whitespace-nowrap">상태</th>
          <th className="px-3 py-2.5 whitespace-nowrap">배송상태</th>
          <th className="px-3 py-2.5 text-right whitespace-nowrap">금액</th>
        </tr>
      </thead>
      <tbody>
        {items.map((item) => {
          const isSale = item.sourceType === 'sale';
          const d = item.data;
          const state = isSale ? getRowStateSale(d) : getRowStateDelivery(d);
          const dot = rightDot(state);
          const cancelled = state === 'cancelled';
          const sel = selectedId === item.id;
          const checked = isSale ? checkedSale.has(item.id) : checkedDelivery.has(item.id);
          const onCheck = () => (isSale ? onCheckSale(item.id) : onCheckDelivery(item.id));
          const amt = (d.total_amount || 0) - (d.discount_amount || 0);
          // 채널: 판매는 sale_channel 직접(매장/출장/톡/온라인·레거시오프라인), B2B납품은 '거래처'
          const channelLabel = isSale ? (SALE_CHANNEL_LABEL[d.sale_channel as string] || d.sale_channel || '—') : '거래처';
          // 배송상태: 목록 데이터만으로 파생(추가 쿼리 0)
          const ship: ShipStatus = isSale ? getSaleShipStatus(d) : getDeliveryShipStatus(d);
          return (
            <tr
              key={`${item.sourceType}-${item.id}`}
              onClick={prepMode ? onCheck : () => onSelect(item.id, item.sourceType)}
              className={`border-b border-neutral-100 cursor-pointer transition ${
                sel ? 'bg-[#F4F0EA] shadow-[inset_3px_0_0_#1A1A1A]' : 'hover:bg-warm-ivory/50'
              } ${checked ? 'bg-blue-50' : ''} ${cancelled ? 'opacity-50' : ''}`}
            >
              {prepMode && (
                <td className="px-2 py-2.5">
                  <input type="checkbox" checked={checked} onChange={onCheck} onClick={(e) => e.stopPropagation()}
                    className="w-4 h-4 rounded border-neutral-300" />
                </td>
              )}
              <td className="px-3 py-2.5 text-neutral-600 whitespace-nowrap tabular-nums">{formatDate(item.date, 'yy.MM.dd')}</td>
              <td className="px-3 py-2.5">
                <div className="flex items-center gap-1 min-w-0">
                  <span className={`font-semibold text-indigo-black truncate ${cancelled ? 'line-through' : ''}`}>{d.customer_name || '—'}</span>
                  {isSale && <ActivityChips types={actLookup?.(d.customer_phone)} className="shrink-0" />}
                </div>
                {isSale && d.company_name && <div className="text-[11px] text-neutral-400">{d.company_name}</div>}
              </td>
              <td className="px-3 py-2.5 whitespace-nowrap text-xs text-neutral-500">{channelLabel}</td>
              <td className="px-3 py-2.5 whitespace-nowrap">
                <span className={`inline-flex items-center gap-1.5 text-xs font-medium ${statusTextClass(state)}`}>
                  {dot && <span className={`w-1.5 h-1.5 rounded-full ${dot.color}`} />}
                  {gridStatusLabel(state)}
                </span>
              </td>
              <td className="px-3 py-2.5 whitespace-nowrap">
                {ship.tone === 'mute'
                  ? <span className="text-xs text-neutral-300">{ship.label}</span>
                  : <span className={`inline-block px-2 py-0.5 rounded text-[10.5px] font-semibold ${SHIP_BADGE[ship.tone]}`}>{ship.label}</span>}
                {ship.autoPicked && <span className="ml-1 text-[10px] text-neutral-400">· 수거</span>}
              </td>
              <td className={`px-3 py-2.5 text-right font-bold tabular-nums ${cancelled ? 'line-through text-neutral-400' : 'text-indigo-black'}`}>{formatKRW(amt)}</td>
            </tr>
          );
        })}
      </tbody>
    </table>
  );
});

const SaleRow = memo(function SaleRow({ sale, note, selected, onClick, prepMode, checked, onCheck, actTypes }: {
  sale: OfflineSale; note?: LatestNote; selected?: boolean; onClick: () => void;
  prepMode?: boolean; checked?: boolean; onCheck?: () => void; actTypes?: ActivityTypes;
}) {
  const state = getRowStateSale(sale);
  const isCancelled = state === 'cancelled';
  const dot = rightDot(state);
  const channelLabel = SALE_CHANNEL_LABEL[sale.sale_channel || 'offline'] || '오프라인';

  return (
    <div
      onClick={prepMode ? onCheck : onClick}
      className={`flex items-stretch gap-3 px-4 py-3 cursor-pointer transition ${
        selected ? 'bg-neutral-50' : 'hover:bg-warm-ivory/40'
      } ${isCancelled ? 'opacity-50' : ''} ${checked ? 'bg-blue-50' : ''}`}
    >
      {/* 좌측 색 줄 */}
      <div className={`w-1 rounded ${stripColor(state)}`} style={{ alignSelf: 'stretch' }} />

      {/* 체크박스 (준비표 모드) */}
      {prepMode && (
        <input type="checkbox" checked={checked} onChange={onCheck}
          onClick={(e) => e.stopPropagation()}
          className="w-4 h-4 rounded border-neutral-300 shrink-0 self-center" />
      )}

      {/* 본문 */}
      <div className="flex-1 min-w-0 flex items-center gap-3">
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-1 min-w-0">
            <span className={`text-sm font-semibold text-indigo-black truncate ${isCancelled ? 'line-through' : ''}`}>
              {sale.customer_name}
            </span>
            <ActivityChips types={actTypes} className="shrink-0" />
          </div>
          <div className="text-xs text-neutral-500 mt-0.5 flex items-center gap-1.5 flex-wrap">
            <span>{formatDate(sale.sale_date)}</span>
            <span className="text-neutral-300">·</span>
            <span>{PAYMENT_METHOD_LABEL[sale.payment_method] || sale.payment_method}</span>
            <span className="text-neutral-300">·</span>
            <span>{channelLabel}</span>
            <span className="text-neutral-300">·</span>
            <span className={statusTextClass(state)}>{statusLabel(state)}</span>
          </div>
          <NotePreview note={note} className="mt-0.5" />
        </div>

        {/* 우측: 도트 + 금액 */}
        <div className="flex items-center gap-2 shrink-0">
          {dot && <span className={`w-2 h-2 rounded-full ${dot.color}`} title={dot.title} />}
          <span className={`text-sm font-bold ${isCancelled ? 'line-through text-neutral-400' : 'text-indigo-black'}`}>
            {formatKRW((sale.total_amount || 0) - (sale.discount_amount || 0))}
          </span>
        </div>
      </div>
    </div>
  );
});

/* 거래처 납품 행 — 2026-05-26 Phase G-4: 안 A 통합 디자인 (SaleRow 와 동일 패턴, B2B 분류 로직만 분기) */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type DeliveryRowData = any;

const DeliveryRow = memo(function DeliveryRow({ delivery, selected, onClick, prepMode, checked, onCheck }: {
  delivery: DeliveryRowData; selected?: boolean; onClick: () => void;
  prepMode?: boolean; checked?: boolean; onCheck?: () => void;
}) {
  const state = getRowStateDelivery(delivery);
  const isCancelled = state === 'cancelled';
  const dot = rightDot(state);
  const ctypeLabel = CUSTOMER_TYPE_LABEL[delivery.customer_type as string] || '거래처';

  return (
    <div
      onClick={prepMode ? onCheck : onClick}
      className={`flex items-stretch gap-3 px-4 py-3 cursor-pointer transition ${
        selected ? 'bg-neutral-50' : 'hover:bg-warm-ivory/40'
      } ${isCancelled ? 'opacity-50' : ''} ${checked ? 'bg-blue-50' : ''}`}
    >
      {/* 좌측 색 줄 */}
      <div className={`w-1 rounded ${stripColor(state)}`} style={{ alignSelf: 'stretch' }} />

      {/* 체크박스 (준비표 모드) */}
      {prepMode && (
        <input type="checkbox" checked={checked} onChange={onCheck}
          onClick={(e) => e.stopPropagation()}
          className="w-4 h-4 rounded border-neutral-300 shrink-0 self-center" />
      )}

      {/* 본문 */}
      <div className="flex-1 min-w-0 flex items-center gap-3">
        <div className="flex-1 min-w-0">
          <div className={`text-sm font-semibold text-indigo-black truncate ${isCancelled ? 'line-through' : ''}`}>
            {delivery.company_name || delivery.customer_name || '미지정'}
          </div>
          <div className="text-xs text-neutral-500 mt-0.5 flex items-center gap-1.5 flex-wrap">
            <span>{formatDate(delivery.delivery_date || delivery.expected_date || '')}</span>
            <span className="text-neutral-300">·</span>
            <span>{PAYMENT_METHOD_LABEL[delivery.payment_method as string] || delivery.payment_method || '-'}</span>
            <span className="text-neutral-300">·</span>
            <span>{ctypeLabel}</span>
            <span className="text-neutral-300">·</span>
            <span className={statusTextClass(state)}>{statusLabel(state)}</span>
          </div>
        </div>

        {/* 우측: 도트 + 금액 */}
        <div className="flex items-center gap-2 shrink-0">
          {dot && <span className={`w-2 h-2 rounded-full ${dot.color}`} title={dot.title} />}
          <span className={`text-sm font-bold ${isCancelled ? 'line-through text-neutral-400' : 'text-indigo-black'}`}>
            {/* 납품 total_amount는 이미 net(할인 반영) → 할인 재차감 금지 */}
            {formatKRW((delivery.total_amount as number) || 0)}
          </span>
        </div>
      </div>
    </div>
  );
});
