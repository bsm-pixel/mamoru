'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations, useUpdateConsultationStatus } from '@/hooks/use-consultations';
import { RescheduleModal } from './reschedule-modal';
import { SuggestTimeModal } from './suggest-time-modal';
import {
  formatPhone,
  formatDday,
  CONSULTATION_STATUS_LABEL,
  CONSULTATION_STATUS_COLOR,
} from '@/lib/utils/format';
import { MapPin, Search, ChevronLeft, ChevronRight, CalendarClock, X } from 'lucide-react';
import type { Consultation } from '@/lib/supabase/types';

// R2: 4탭 (신규요청 / 출장제안 / 재요청 / 확정)
type TabKey = 'new_request' | 'suggested' | 're_request' | 'confirmed';

const TABS: { key: TabKey; label: string }[] = [
  { key: 'new_request', label: '신규요청' },
  { key: 'suggested', label: '출장제안' },
  { key: 're_request', label: '재요청' },
  { key: 'confirmed', label: '확정' },
];

function getTabFilters(tab: TabKey) {
  switch (tab) {
    case 'new_request':
      return { status: 'pending_admin', orderBy: 'received_at_desc' };
    case 'suggested':
      return { status: 'suggested', orderBy: 'received_at_desc' };
    case 're_request':
      return { statuses: ['reschedule_requested', 'change_requested'], orderBy: 'received_at_desc' };
    case 'confirmed':
      return { status: 'confirmed', dateFilter: 'upcoming' as const, orderBy: 'visit_date_asc' };
  }
}

interface Props {
  selectedFieldId?: string | null;
  onFieldSelect?: (id: string | null) => void;
  onSelect?: (id: string) => void;  // 상세 패널 열기
}

export function FieldRequestList({ selectedFieldId, onFieldSelect, onSelect }: Props) {
  const router = useRouter();
  const [tab, setTab] = useState<TabKey>('new_request');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [rescheduleTarget, setRescheduleTarget] = useState<Consultation | null>(null);
  const [suggestTarget, setSuggestTarget] = useState<string | null>(null);

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

  const handleCancel = (c: Consultation) => {
    updateStatus.mutate({ id: c.id, status: 'cancelled', note: '출장 취소' });
  };

  const renderRow = (c: Consultation) => {
    const dday = formatDday(c.visit_date);
    const busy = updateStatus.isPending && updateStatus.variables?.id === c.id;
    const isHighlighted = selectedFieldId === c.id;

    return (
      <div
        key={c.id}
        id={`field-${c.id}`}
        className={`px-4 py-3 hover:bg-warm-ivory/60 transition ${
          isHighlighted ? 'bg-terracotta/5 ring-1 ring-terracotta' : ''
        } ${dday.isToday ? 'border-l-2 border-l-terracotta' : ''}`}
      >
        <div
          className="cursor-pointer"
          onClick={() => {
            onFieldSelect?.(c.id);
            onSelect ? onSelect(c.id) : router.push(`/consultations/${c.id}`);
          }}
        >
          <div className="flex items-center gap-2">
            <span className="text-sm font-semibold text-indigo-black truncate">{c.name}</span>
            <Badge className={CONSULTATION_STATUS_COLOR[c.status]}>
              {CONSULTATION_STATUS_LABEL[c.status]}
            </Badge>
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
            {c.address_road && (
              <span className="flex items-center gap-1">
                <MapPin size={11} />
                {c.address_sido} {c.address_sigungu}
              </span>
            )}
            {c.visit_date && <span>{c.visit_date} {c.visit_time || ''}</span>}
          </div>
        </div>

        {/* R2: 모든 탭에 [일정변경][취소] + 탭별 추가 액션 */}
        <div className="flex gap-1 mt-2 flex-wrap">
          {/* 신규요청 탭: 시간제안 추가 */}
          {tab === 'new_request' && (
            <Button variant="primary" size="sm" disabled={busy} onClick={() => setSuggestTarget(c.id)}>
              시간 제안
            </Button>
          )}
          {/* 재요청 탭: 새 시간 제안 */}
          {tab === 're_request' && (
            <Button variant="primary" size="sm" disabled={busy} onClick={() => setSuggestTarget(c.id)}>
              새 시간 제안
            </Button>
          )}
          {/* 공통: 일정변경 + 취소 */}
          <Button variant="secondary" size="sm" disabled={busy} onClick={() => setRescheduleTarget(c)}>
            <CalendarClock size={12} />
            일정변경
          </Button>
          <Button
            variant="danger"
            size="sm"
            disabled={busy}
            loading={updateStatus.variables?.status === 'cancelled' && busy}
            onClick={() => handleCancel(c)}
          >
            <X size={12} />
            취소
          </Button>
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

      {/* R2: 4탭 */}
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
              <Skeleton key={i} className="h-20 w-full" />
            ))}
          </div>
        ) : consultations.length === 0 ? (
          <div className="flex items-center justify-center h-32 text-sm text-neutral-400">
            해당 출장요청이 없습니다
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
      {suggestTarget && (
        <SuggestTimeModal open={!!suggestTarget} onClose={() => setSuggestTarget(null)} consultationId={suggestTarget} />
      )}
    </div>
  );
}
