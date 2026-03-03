'use client';

import { useState, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations, useUpdateConsultationStatus } from '@/hooks/use-consultations';
import { RescheduleModal } from './reschedule-modal';
import {
  formatPhone,
  formatDday,
  formatDateGroup,
} from '@/lib/utils/format';
import { Calendar, Search, ChevronLeft, ChevronRight, CheckCircle, X } from 'lucide-react';
import type { Consultation } from '@/lib/supabase/types';

// R2: 3탭 (확정 / 취소 / 완료)
type TabKey = 'confirmed' | 'cancelled' | 'completed';

const TABS: { key: TabKey; label: string }[] = [
  { key: 'confirmed', label: '확정' },
  { key: 'cancelled', label: '취소' },
  { key: 'completed', label: '완료' },
];

function getTabFilters(tab: TabKey) {
  switch (tab) {
    case 'confirmed':
      return { status: 'confirmed', orderBy: 'visit_date_asc' };
    case 'cancelled':
      return { status: 'cancelled', orderBy: 'updated_at_desc' };
    case 'completed':
      return { status: 'completed', orderBy: 'updated_at_desc' };
  }
}

export function StoreVisitList() {
  const router = useRouter();
  const [tab, setTab] = useState<TabKey>('confirmed');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [rescheduleTarget, setRescheduleTarget] = useState<Consultation | null>(null);

  const tabFilters = getTabFilters(tab);
  const { data, isLoading } = useConsultations({
    ...tabFilters,
    type: 'store_visit',
    search,
    page,
    limit: 20,
  });
  const consultations = data?.consultations || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);

  const updateStatus = useUpdateConsultationStatus();

  // 확정 탭: 날짜 그룹으로 묶기
  const dateGroups = useMemo(() => {
    if (tab !== 'confirmed') return null;
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

  const renderRow = (c: Consultation) => {
    const dday = formatDday(c.visit_date);
    const busy = updateStatus.isPending && updateStatus.variables?.id === c.id;

    return (
      <div
        key={c.id}
        className={`flex items-center gap-3 px-4 py-3 hover:bg-warm-ivory/60 transition ${
          dday.isToday ? 'bg-terracotta/5 border-l-2 border-l-terracotta' : ''
        }`}
      >
        <div
          className="flex-1 min-w-0 cursor-pointer"
          onClick={() => router.push(`/consultations/${c.id}`)}
        >
          <div className="flex items-center gap-2">
            <span className="text-sm font-semibold text-indigo-black truncate">{c.name}</span>
            {dday.label && tab === 'confirmed' && (
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
            {c.visit_date && (
              <span className="flex items-center gap-1">
                <Calendar size={12} />
                {c.visit_date} {c.visit_time || ''}
              </span>
            )}
          </div>
        </div>
        {/* R2: 확정 탭에서만 인라인 [일정변경][취소] */}
        {tab === 'confirmed' && (
          <div className="flex gap-1 shrink-0">
            <Button variant="secondary" size="sm" disabled={busy} onClick={() => setRescheduleTarget(c)}>
              일정변경
            </Button>
            <Button
              variant="danger"
              size="sm"
              disabled={busy}
              loading={updateStatus.variables?.status === 'cancelled' && busy}
              onClick={() => updateStatus.mutate({ id: c.id, status: 'cancelled', note: '매장방문 취소' })}
            >
              <X size={12} />
              취소
            </Button>
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="space-y-4">
      {/* 검색 */}
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

      {/* R2: 3탭 */}
      <div className="flex items-center gap-1 overflow-x-auto pb-1">
        {TABS.map((t) => (
          <button
            key={t.key}
            onClick={() => { setTab(t.key); setPage(1); }}
            className={`flex items-center gap-1 px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition ${
              tab === t.key
                ? 'bg-terracotta text-cream'
                : 'bg-card-white text-neutral-500 hover:bg-warm-ivory'
            }`}
          >
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
              <Skeleton key={i} className="h-16 w-full" />
            ))}
          </div>
        ) : consultations.length === 0 ? (
          <div className="flex items-center justify-center h-32 text-sm text-neutral-400">
            {tab === 'confirmed' ? '확정된 방문이 없습니다' : '해당 상담이 없습니다'}
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

      {/* 일정변경 모달 */}
      {rescheduleTarget && (
        <RescheduleModal
          open={!!rescheduleTarget}
          onClose={() => setRescheduleTarget(null)}
          consultationId={rescheduleTarget.id}
          currentDate={rescheduleTarget.visit_date || ''}
          currentTime={rescheduleTarget.visit_time || ''}
          consultationType={rescheduleTarget.consultation_type}
          uniqueId={rescheduleTarget.unique_id}
        />
      )}
    </div>
  );
}
