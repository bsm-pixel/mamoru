'use client';

import { useState, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations, useUpdateConsultationStatus } from '@/hooks/use-consultations';
import { SuggestTimeModal } from './suggest-time-modal';
import { HoldReasonModal } from './hold-reason-modal';
import {
  formatPhone,
  formatDday,
  formatDateGroup,
  CONSULTATION_STATUS_LABEL,
  CONSULTATION_STATUS_COLOR,
} from '@/lib/utils/format';
import { MapPin, Search, ChevronLeft, ChevronRight, Map as MapIcon, List, CheckCircle, AlertTriangle } from 'lucide-react';
import type { Consultation } from '@/lib/supabase/types';

type TabKey = 'waiting' | 'suggested' | 'upcoming' | 'overdue' | 'completed' | 'on_hold' | 'cancelled';

const TABS: { key: TabKey; label: string; icon?: React.ReactNode }[] = [
  { key: 'waiting', label: '처리대기' },
  { key: 'suggested', label: '제안완료' },
  { key: 'upcoming', label: '예정' },
  { key: 'overdue', label: '완료 필요', icon: <AlertTriangle size={12} /> },
  { key: 'completed', label: '완료', icon: <CheckCircle size={12} /> },
  { key: 'on_hold', label: '보류' },
  { key: 'cancelled', label: '취소' },
];

function getTabFilters(tab: TabKey) {
  switch (tab) {
    case 'waiting':
      return { statuses: ['pending_admin', 'reschedule_requested'], orderBy: 'received_at_desc' };
    case 'suggested':
      return { status: 'suggested' };
    case 'upcoming':
      return { status: 'confirmed', dateFilter: 'upcoming' as const, orderBy: 'visit_date_asc' };
    case 'overdue':
      return { status: 'confirmed', dateFilter: 'past' as const, orderBy: 'visit_date_asc' };
    case 'completed':
      return { status: 'completed', orderBy: 'updated_at_desc' };
    case 'on_hold':
      return { status: 'on_hold', orderBy: 'updated_at_desc' };
    case 'cancelled':
      return { status: 'cancelled', orderBy: 'updated_at_desc' };
  }
}

interface Props {
  onToggleMap?: () => void;
  showMap?: boolean;
}

export function FieldRequestList({ onToggleMap, showMap }: Props) {
  const router = useRouter();
  const [tab, setTab] = useState<TabKey>('waiting');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [suggestTarget, setSuggestTarget] = useState<string | null>(null);
  const [holdTarget, setHoldTarget] = useState<string | null>(null);

  const tabFilters = getTabFilters(tab);
  const { data, isLoading } = useConsultations({
    ...tabFilters,
    type: 'field_request',
    search,
    page,
    limit: 20,
  });
  const consultations = data?.consultations || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);
  const updateStatus = useUpdateConsultationStatus();

  // 날짜 그룹 (예정/완료필요 탭)
  const dateGroups = useMemo(() => {
    if (tab !== 'upcoming' && tab !== 'overdue') return null;
    const map = new Map<string, Consultation[]>();
    for (const c of consultations) {
      const key = c.visit_date || 'no-date';
      if (!map.has(key)) map.set(key, []);
      map.get(key)!.push(c);
    }
    return Array.from(map, ([dateStr, items]) => ({
      dateStr,
      label: dateStr === 'no-date' ? '날짜 미정' : formatDateGroup(dateStr),
      items,
    }));
  }, [tab, consultations]);

  const renderActions = (c: Consultation) => {
    switch (tab) {
      case 'waiting':
        return (
          <>
            <Button variant="primary" size="sm" onClick={() => setSuggestTarget(c.id)}>
              {c.status === 'reschedule_requested' ? '새 시간 제안' : '시간 제안'}
            </Button>
            <Button variant="secondary" size="sm" onClick={() => updateStatus.mutate({ id: c.id, status: 'confirmed' })}>
              즉시 확정
            </Button>
            <Button variant="ghost" size="sm" onClick={() => setHoldTarget(c.id)}>보류</Button>
            <Button variant="danger" size="sm" onClick={() => updateStatus.mutate({ id: c.id, status: 'cancelled' })}>
              취소
            </Button>
          </>
        );
      case 'suggested':
        return (
          <>
            <Button variant="primary" size="sm" onClick={() => updateStatus.mutate({ id: c.id, status: 'confirmed' })}>
              수동 확정
            </Button>
            <Button variant="ghost" size="sm" onClick={() => setHoldTarget(c.id)}>보류</Button>
            <Button variant="danger" size="sm" onClick={() => updateStatus.mutate({ id: c.id, status: 'cancelled' })}>
              취소
            </Button>
          </>
        );
      case 'upcoming':
        return (
          <>
            <Button variant="ghost" size="sm" onClick={() => setHoldTarget(c.id)}>보류</Button>
            <Button variant="danger" size="sm" onClick={() => updateStatus.mutate({ id: c.id, status: 'cancelled' })}>
              취소
            </Button>
          </>
        );
      case 'overdue':
        return (
          <Button variant="primary" size="sm"
            onClick={() => updateStatus.mutate({ id: c.id, status: 'completed', note: '출장 완료 처리' })}
          >
            <CheckCircle size={14} />
            출장 완료
          </Button>
        );
      case 'on_hold':
        return (
          <>
            <Button variant="secondary" size="sm" onClick={() => updateStatus.mutate({ id: c.id, status: 'pending_admin' })}>
              대기 복원
            </Button>
            <Button variant="danger" size="sm" onClick={() => updateStatus.mutate({ id: c.id, status: 'cancelled' })}>
              취소
            </Button>
          </>
        );
      default:
        return null;
    }
  };

  const renderRow = (c: Consultation) => {
    const dday = formatDday(c.visit_date);

    return (
      <div key={c.id} className={`px-4 py-3 hover:bg-warm-ivory/60 transition ${
        dday.isToday ? 'bg-terracotta/5 border-l-2 border-l-terracotta' : ''
      }`}>
        <div className="cursor-pointer" onClick={() => router.push(`/consultations/${c.id}`)}>
          <div className="flex items-center gap-2">
            <span className="text-sm font-semibold text-indigo-black truncate">{c.name}</span>
            <Badge className={CONSULTATION_STATUS_COLOR[c.status]}>
              {CONSULTATION_STATUS_LABEL[c.status]}
            </Badge>
            {dday.label && (tab === 'upcoming' || tab === 'overdue') && (
              <Badge className={
                dday.isToday ? 'bg-terracotta/10 text-terracotta'
                : dday.isPast ? 'bg-error-soft text-error'
                : 'bg-info-soft text-info'
              }>
                {dday.label}
              </Badge>
            )}
            {c.status === 'on_hold' && c.hold_reason && (
              <span className="text-xs text-neutral-400 truncate max-w-[120px]" title={c.hold_reason}>
                ({c.hold_reason})
              </span>
            )}
          </div>
          <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
            <span>{formatPhone(c.phone)}</span>
            {c.address_road && (
              <span className="flex items-center gap-1">
                <MapPin size={11} />
                {c.address_sido} {c.address_sigungu}
              </span>
            )}
            {c.visit_date && (
              <span>{c.visit_date} {c.visit_time || ''}</span>
            )}
          </div>
        </div>
        <div className="flex gap-1 mt-2 flex-wrap">
          {renderActions(c)}
        </div>
      </div>
    );
  };

  return (
    <div className="space-y-4">
      {/* 검색 + 지도 토글 */}
      <div className="flex items-center gap-2">
        <div className="flex-1 relative">
          <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => { setSearch(e.target.value); setPage(1); }}
            placeholder="이름, 전화번호 검색"
            className="w-full h-9 pl-9 pr-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm text-indigo-black placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 transition"
          />
        </div>
        {onToggleMap && (
          <Button variant={showMap ? 'primary' : 'secondary'} size="sm" onClick={onToggleMap}>
            {showMap ? <List size={14} /> : <MapIcon size={14} />}
            {showMap ? '리스트' : '지도'}
          </Button>
        )}
      </div>

      {/* 탭 */}
      <div className="flex items-center gap-1 overflow-x-auto pb-1">
        {TABS.map((t) => (
          <button
            key={t.key}
            onClick={() => { setTab(t.key); setPage(1); }}
            className={`flex items-center gap-1 px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition ${
              tab === t.key
                ? t.key === 'overdue' ? 'bg-error text-cream' : 'bg-terracotta text-cream'
                : 'bg-card-white text-neutral-500 hover:bg-warm-ivory'
            }`}
          >
            {t.icon}
            {t.label}
          </button>
        ))}
        <span className="ml-auto text-xs text-neutral-500 shrink-0">{total}건</span>
      </div>

      {/* 목록 */}
      <Card padding={false}>
        {isLoading ? (
          <div className="p-4 space-y-3">
            {Array.from({ length: 5 }).map((_, i) => (
              <Skeleton key={i} className="h-20 w-full" />
            ))}
          </div>
        ) : consultations.length === 0 ? (
          <div className="flex items-center justify-center h-32 text-sm text-neutral-400">
            해당 상담이 없습니다
          </div>
        ) : dateGroups ? (
          <div>
            {dateGroups.map((group) => (
              <div key={group.dateStr}>
                <div className="px-4 py-2 bg-warm-ivory/80 text-xs font-bold text-neutral-600 sticky top-0">
                  {group.label}
                  <span className="ml-2 font-normal text-neutral-400">{group.items.length}건</span>
                </div>
                <div className="divide-y divide-neutral-100">
                  {group.items.map(renderRow)}
                </div>
              </div>
            ))}
          </div>
        ) : (
          <div className="divide-y divide-neutral-100">
            {consultations.map(renderRow)}
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

      {/* 모달 */}
      {suggestTarget && (
        <SuggestTimeModal open={!!suggestTarget} onClose={() => setSuggestTarget(null)} consultationId={suggestTarget} />
      )}
      {holdTarget && (
        <HoldReasonModal open={!!holdTarget} onClose={() => setHoldTarget(null)} consultationId={holdTarget} />
      )}
    </div>
  );
}
