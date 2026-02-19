'use client';

import { useState, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations, useSendNotification, useUpdateConsultationStatus } from '@/hooks/use-consultations';
import { RescheduleModal } from './reschedule-modal';
import { HoldReasonModal } from './hold-reason-modal';
import {
  formatPhone,
  formatDday,
  formatDateGroup,
  CONSULTATION_STATUS_LABEL,
  CONSULTATION_STATUS_COLOR,
} from '@/lib/utils/format';
import { Calendar, Search, ChevronLeft, ChevronRight, CheckCircle, AlertTriangle } from 'lucide-react';
import type { Consultation } from '@/lib/supabase/types';

type TabKey = 'upcoming' | 'change_requested' | 'overdue' | 'completed' | 'on_hold' | 'cancelled';

const TABS: { key: TabKey; label: string; icon?: React.ReactNode }[] = [
  { key: 'upcoming', label: '예정' },
  { key: 'change_requested', label: '변경/취소' },
  { key: 'overdue', label: '완료 필요', icon: <AlertTriangle size={12} /> },
  { key: 'completed', label: '완료', icon: <CheckCircle size={12} /> },
  { key: 'on_hold', label: '보류' },
  { key: 'cancelled', label: '취소' },
];

/** 탭별 쿼리 필터 설정 */
function getTabFilters(tab: TabKey) {
  switch (tab) {
    case 'upcoming':
      return { status: 'confirmed', dateFilter: 'upcoming' as const, orderBy: 'visit_date_asc' };
    case 'change_requested':
      return { status: 'change_requested', orderBy: 'updated_at_desc' };
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

export function StoreVisitList() {
  const router = useRouter();
  const [tab, setTab] = useState<TabKey>('upcoming');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [rescheduleTarget, setRescheduleTarget] = useState<Consultation | null>(null);
  const [holdTarget, setHoldTarget] = useState<string | null>(null);

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

  const sendNotify = useSendNotification();
  const updateStatus = useUpdateConsultationStatus();

  // 예정 탭: 날짜 그룹으로 묶기
  const dateGroups = useMemo(() => {
    if (tab !== 'upcoming' && tab !== 'overdue') return null;
    const groups: { dateStr: string; label: string; items: Consultation[] }[] = [];
    const map = new Map<string, Consultation[]>();
    for (const c of consultations) {
      const key = c.visit_date || 'no-date';
      if (!map.has(key)) map.set(key, []);
      map.get(key)!.push(c);
    }
    for (const [dateStr, items] of map) {
      groups.push({
        dateStr,
        label: dateStr === 'no-date' ? '날짜 미정' : formatDateGroup(dateStr),
        items,
      });
    }
    return groups;
  }, [tab, consultations]);

  const renderRow = (c: Consultation) => {
    const dday = formatDday(c.visit_date);
    const busy = updateStatus.isPending && updateStatus.variables?.id === c.id;
    const busyStatus = busy ? updateStatus.variables?.status : null;

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
            {dday.label && (
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
            {c.visit_date && (
              <span className="flex items-center gap-1">
                <Calendar size={12} />
                {c.visit_date} {c.visit_time || ''}
              </span>
            )}
          </div>
        </div>
        {/* 액션 버튼 */}
        <div className="flex gap-1 shrink-0 flex-wrap">
          {tab === 'upcoming' && (
            <>
              <Button variant="secondary" size="sm" disabled={busy} onClick={() => setRescheduleTarget(c)}>
                일정변경
              </Button>
              <Button
                variant="ghost" size="sm"
                disabled={busy || sendNotify.isPending}
                loading={sendNotify.isPending && sendNotify.variables?.consultationId === c.id}
                onClick={() => sendNotify.mutate({ consultationId: c.id, template: 'confirmed' })}
              >
                알림 재발송
              </Button>
              <Button variant="ghost" size="sm" disabled={busy} onClick={() => setHoldTarget(c.id)}>보류</Button>
            </>
          )}
          {tab === 'overdue' && (
            <Button
              variant="primary" size="sm"
              disabled={busy}
              loading={busyStatus === 'completed'}
              onClick={() => updateStatus.mutate({ id: c.id, status: 'completed', note: '방문 완료 처리' })}
            >
              <CheckCircle size={14} />
              방문 완료
            </Button>
          )}
          {tab === 'on_hold' && (
            <>
              <Button variant="secondary" size="sm" disabled={busy} loading={busyStatus === 'confirmed'} onClick={() => updateStatus.mutate({ id: c.id, status: 'confirmed' })}>
                확정 복원
              </Button>
              <Button variant="danger" size="sm" disabled={busy} loading={busyStatus === 'cancelled'} onClick={() => updateStatus.mutate({ id: c.id, status: 'cancelled' })}>
                취소
              </Button>
            </>
          )}
        </div>
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
              <Skeleton key={i} className="h-16 w-full" />
            ))}
          </div>
        ) : consultations.length === 0 ? (
          <div className="flex items-center justify-center h-32 text-sm text-neutral-400">
            {tab === 'upcoming' ? '예정된 방문이 없습니다' : '해당 상담이 없습니다'}
          </div>
        ) : dateGroups ? (
          /* 날짜 그룹 뷰 */
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
          /* 일반 리스트 */
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
      {rescheduleTarget && (
        <RescheduleModal
          open={!!rescheduleTarget}
          onClose={() => setRescheduleTarget(null)}
          consultationId={rescheduleTarget.id}
          currentDate={rescheduleTarget.visit_date || ''}
          currentTime={rescheduleTarget.visit_time || ''}
        />
      )}
      {holdTarget && (
        <HoldReasonModal open={!!holdTarget} onClose={() => setHoldTarget(null)} consultationId={holdTarget} />
      )}
    </div>
  );
}
