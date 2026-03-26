'use client';

import { useState, useEffect, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations } from '@/hooks/use-consultations';
import {
  formatPhone,
  formatDday,
  CONSULTATION_STATUS_LABEL,
  CONSULTATION_STATUS_COLOR,
} from '@/lib/utils/format';
import { MapPin, Search, ChevronLeft, ChevronRight, Navigation, CalendarCheck, LayoutGrid } from 'lucide-react';
import { MobileFieldDayView } from './mobile-field-day-view';
import type { Consultation } from '@/lib/supabase/types';

// 5탭: 오늘출장 | 신규접수 | 대응필요 | 확정 | 지난내역
type TabKey = 'today' | 'new_intake' | 'action_needed' | 'confirmed' | 'past';

const TABS: { key: TabKey; label: string }[] = [
  { key: 'today', label: '오늘출장' },
  { key: 'new_intake', label: '신규접수' },
  { key: 'action_needed', label: '일정재요청' },
  { key: 'confirmed', label: '확정' },
  { key: 'past', label: '지난내역' },
];

function getTabFilters(tab: TabKey) {
  switch (tab) {
    case 'today':
      return null;
    case 'new_intake':
      return { status: 'pending_admin', orderBy: 'received_at_desc' as const };
    case 'action_needed':
      return { statuses: ['suggested', 'reschedule_requested', 'change_requested'], orderBy: 'received_at_desc' as const };
    case 'confirmed':
      return { status: 'confirmed', dateFilter: 'upcoming' as const, orderBy: 'visit_date_asc' as const };
    case 'past':
      return { statuses: ['completed', 'cancelled'], orderBy: 'updated_at_desc' as const };
  }
}

interface Props {
  selectedFieldId?: string | null;
  onFieldSelect?: (id: string | null) => void;
  onSelect?: (id: string) => void;  // 상세 패널 열기
  onSubTabChange?: (tab: string) => void; // 서브탭 변경 → 지도 연동
}

export function FieldRequestList({ selectedFieldId, onFieldSelect, onSelect, onSubTabChange }: Props) {
  const router = useRouter();
  const [tab, setTab] = useState<TabKey>('new_intake');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [showRegionFilter, setShowRegionFilter] = useState(false);
  const [selectedRegion, setSelectedRegion] = useState<string | null>(null);

  // 기본 탭 결정용 건수 확인
  const today = new Date().toISOString().slice(0, 10);
  const { data: todayCheck } = useConsultations({
    status: 'confirmed',
    type: 'field_request',
    limit: 100,
  });
  const todayCount = todayCheck?.consultations?.filter(c => c.visit_date === today).length || 0;

  const { data: actionCheck } = useConsultations({
    statuses: ['suggested', 'reschedule_requested', 'change_requested'],
    type: 'field_request',
    limit: 1,
  });
  const actionCount = actionCheck?.total || 0;

  // 최초 로드 시 우선순위: 대응필요 > 오늘출장 > 신규접수
  const [initialTabSet, setInitialTabSet] = useState(false);
  useEffect(() => {
    if (!initialTabSet && todayCheck && actionCheck) {
      if (actionCount > 0) { setTab('action_needed'); onSubTabChange?.('action_needed'); }
      else if (todayCount > 0) { setTab('today'); onSubTabChange?.('today'); }
      setInitialTabSet(true);
    }
  }, [todayCheck, actionCheck, todayCount, actionCount, initialTabSet]);

  const tabFilters = getTabFilters(tab);
  const { data, isLoading } = useConsultations({
    ...(tabFilters || {}),
    type: 'field_request',
    search: tab !== 'today' ? search : '',
    page: tab !== 'today' ? page : 1,
    limit: 20,
  });
  const consultations = data?.consultations || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);

  // 지역 추출 헬퍼
  const getRegion = (c: Consultation): string => {
    let region = (c as any).address_sigungu;
    if (!region && c.address_road) {
      const match = c.address_road.match(/^\S+\s+(\S+[시군구])/);
      region = match?.[1] || null;
    }
    return region || '기타';
  };

  // 지역 칩 목록 (신규접수/대응필요 탭)
  const regionChips = useMemo(() => {
    if (!showRegionFilter || (tab !== 'action_needed' && tab !== 'new_intake')) return null;
    const countMap = new Map<string, number>();
    for (const c of consultations) {
      const r = getRegion(c);
      countMap.set(r, (countMap.get(r) || 0) + 1);
    }
    return Array.from(countMap, ([region, count]) => ({ region, count }))
      .sort((a, b) => b.count - a.count);
  }, [showRegionFilter, tab, consultations]);

  // 지역 필터 적용된 목록
  const filteredConsultations = useMemo(() => {
    if (!selectedRegion) return consultations;
    return consultations.filter(c => getRegion(c) === selectedRegion);
  }, [consultations, selectedRegion]);

  const renderRow = (c: Consultation) => {
    const dday = formatDday(c.visit_date);
    const isHighlighted = selectedFieldId === c.id;
    const isCancelled = c.status === 'cancelled';
    const isCompleted = c.status === 'completed';

    return (
      <div
        key={c.id}
        id={`field-${c.id}`}
        className={`px-4 py-3 hover:bg-warm-ivory/60 transition ${
          isCancelled ? 'opacity-60 bg-neutral-50 border-l-2 border-l-neutral-300'
          : isCompleted ? 'border-l-2 border-l-green-400'
          : isHighlighted ? 'bg-terracotta/5 ring-1 ring-terracotta'
          : dday.isToday ? 'border-l-2 border-l-terracotta'
          : ''
        }`}
      >
        <div
          className="cursor-pointer"
          onClick={() => {
            onFieldSelect?.(c.id);
            onSelect ? onSelect(c.id) : router.push(`/consultations/${c.id}`);
          }}
        >
          <div className="flex items-center gap-2">
            <span className={`text-sm font-semibold truncate ${isCancelled ? 'line-through text-neutral-400' : 'text-indigo-black'}`}>{c.name}</span>
            <Badge className={CONSULTATION_STATUS_COLOR[c.status]}>
              {CONSULTATION_STATUS_LABEL[c.status]}
            </Badge>
            {dday.label && (tab === 'confirmed' || tab === 'past') && (
              <Badge className={
                dday.isToday ? 'bg-terracotta/10 text-terracotta'
                : dday.isPast ? 'bg-error-soft text-error'
                : 'bg-info-soft text-info'
              }>
                {dday.label}
              </Badge>
            )}
          </div>
          <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
            <span>{formatPhone(c.phone)}</span>
            {c.visit_date && <span>{c.visit_date} {c.visit_time || ''}</span>}
          </div>
          {/* 주소 + 네비 버튼 */}
          {c.address_road && (
            <div className="flex items-center gap-2 mt-1.5">
              <MapPin size={12} className="text-neutral-400 shrink-0" />
              <span className="text-xs text-neutral-600 truncate">{c.address_road} {c.address_detail || ''}</span>
              <a
                href={`https://map.kakao.com/link/search/${encodeURIComponent(c.address_road || '')}`}
                target="_blank"
                rel="noopener noreferrer"
                onClick={(e) => e.stopPropagation()}
                className="shrink-0 inline-flex items-center gap-1 px-2 py-1 rounded-md bg-blue-50 text-blue-600 text-[11px] font-semibold hover:bg-blue-100 transition lg:hidden"
              >
                <Navigation size={10} />
                네비
              </a>
            </div>
          )}
          {/* 신규접수/대응필요: 가능요일/선호시간대 칩 */}
          {(tab === 'new_intake' || tab === 'action_needed') && (() => {
            const raw = (c as any).gas_raw;
            const days = (raw?.days as string)?.split(',').filter(Boolean) || [];
            const times = (raw?.timePrefs as string)?.split(',').filter(Boolean) || [];
            if (days.length === 0 && times.length === 0) return null;
            return (
              <div className="flex items-center gap-1 mt-1.5 flex-wrap">
                {days.map(d => (
                  <span key={d} className="px-1.5 py-0.5 rounded bg-blue-50 text-blue-600 text-[10px] font-medium">{d}</span>
                ))}
                {times.length > 0 && days.length > 0 && <span className="text-neutral-300 text-[10px]">·</span>}
                {times.map(t => (
                  <span key={t} className="px-1.5 py-0.5 rounded bg-amber-50 text-amber-600 text-[10px] font-medium">{t}</span>
                ))}
              </div>
            );
          })()}
        </div>
      </div>
    );
  };

  return (
    <div className="space-y-4">
      {/* 검색 — 오늘출장 탭에서는 숨김 (MobileFieldDayView가 자체 UI) */}
      {tab !== 'today' && (
        <div className="relative">
          <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => { setSearch(e.target.value); setPage(1); }}
            placeholder="이름, 전화번호 검색"
            className="w-full h-9 pl-9 pr-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm text-indigo-black placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 transition"
          />
        </div>
      )}

      {/* 4탭 */}
      <div className="flex items-center gap-1 overflow-x-auto pb-1">
        {TABS.map((t) => (
          <button
            key={t.key}
            onClick={() => { setTab(t.key); setPage(1); setSelectedRegion(null); onSubTabChange?.(t.key); }}
            className={`flex items-center gap-1 px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition ${
              tab === t.key
                ? 'bg-terracotta text-cream'
                : 'bg-card-white text-neutral-500 hover:bg-warm-ivory'
            }`}
          >
            {t.key === 'today' && <CalendarCheck size={12} />}
            {t.label}
            {/* 대응필요 건수 뱃지 */}
            {t.key === 'action_needed' && actionCount > 0 && (
              <span className={`ml-0.5 inline-flex items-center justify-center min-w-[16px] h-[16px] px-1 rounded-full text-[10px] font-bold leading-none ${
                tab === 'action_needed' ? 'bg-cream/30 text-cream' : 'bg-red-500 text-white'
              }`}>
                {actionCount}
              </span>
            )}
            {/* 오늘출장 건수 뱃지 */}
            {t.key === 'today' && todayCount > 0 && (
              <span className={`ml-0.5 inline-flex items-center justify-center min-w-[16px] h-[16px] px-1 rounded-full text-[10px] font-bold leading-none ${
                tab === 'today' ? 'bg-cream/30 text-cream' : 'bg-terracotta text-white'
              }`}>
                {todayCount}
              </span>
            )}
          </button>
        ))}
        {tab !== 'today' && <span className="ml-auto text-xs text-neutral-500 shrink-0">{total}건</span>}
      </div>

      {/* 지역 필터 토글 + 칩 */}
      {(tab === 'new_intake' || tab === 'action_needed') && consultations.length > 0 && (
        <div className="space-y-2">
          <div className="flex justify-end">
            <button
              onClick={() => { setShowRegionFilter(!showRegionFilter); setSelectedRegion(null); }}
              className={`flex items-center gap-1 px-2 py-1 rounded-md text-[11px] font-medium transition ${
                showRegionFilter ? 'bg-terracotta/10 text-terracotta' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
              }`}
            >
              <LayoutGrid size={12} />
              지역별 보기
            </button>
          </div>
          {regionChips && regionChips.length > 0 && (
            <div className="flex items-center gap-1.5 overflow-x-auto pb-1">
              <button
                onClick={() => setSelectedRegion(null)}
                className={`shrink-0 px-2.5 py-1 rounded-full text-[11px] font-semibold transition ${
                  !selectedRegion ? 'bg-terracotta text-cream' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
                }`}
              >
                전체 {consultations.length}
              </button>
              {regionChips.map(({ region, count }) => (
                <button
                  key={region}
                  onClick={() => setSelectedRegion(selectedRegion === region ? null : region)}
                  className={`shrink-0 px-2.5 py-1 rounded-full text-[11px] font-semibold transition ${
                    selectedRegion === region ? 'bg-terracotta text-cream' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
                  }`}
                >
                  {region} {count}
                </button>
              ))}
            </div>
          )}
        </div>
      )}

      {/* 오늘출장 탭: MobileFieldDayView 렌더링 */}
      {tab === 'today' ? (
        <MobileFieldDayView onSelect={onSelect} />
      ) : (
        <>
          {/* 목록 */}
          <Card padding={false}>
            {isLoading ? (
              <div className="p-4 space-y-3">
                {Array.from({ length: 5 }).map((_, i) => (
                  <Skeleton key={i} className="h-20 w-full" />
                ))}
              </div>
            ) : filteredConsultations.length === 0 ? (
              <div className="flex items-center justify-center h-32 text-sm text-neutral-400">
                {selectedRegion ? `${selectedRegion} 지역에 해당 건이 없습니다` : tab === 'past' ? '지난 출장 내역이 없습니다' : '해당 출장요청이 없습니다'}
              </div>
            ) : (
              <div className="divide-y divide-neutral-100">
                {filteredConsultations.map(renderRow)}
              </div>
            )}
          </Card>

          {/* 페이지네이션 */}
          {totalPages > 1 && (
            <div className="flex items-center justify-center gap-2">
              <Button variant="ghost" size="sm" disabled={page <= 1} onClick={() => setPage(page - 1)}>
                <ChevronLeft size={16} />
              </Button>
              <span className="text-sm text-neutral-500">{page} / {totalPages}</span>
              <Button variant="ghost" size="sm" disabled={page >= totalPages} onClick={() => setPage(page + 1)}>
                <ChevronRight size={16} />
              </Button>
            </div>
          )}
        </>
      )}
    </div>
  );
}
