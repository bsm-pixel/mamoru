'use client';

import { useState } from 'react';
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
import { MapPin, Search, ChevronLeft, ChevronRight, Navigation } from 'lucide-react';
import type { Consultation } from '@/lib/supabase/types';

// #4: 2탭 (대응필요 / 확정)
type TabKey = 'action_needed' | 'confirmed';

const TABS: { key: TabKey; label: string }[] = [
  { key: 'action_needed', label: '대응필요' },
  { key: 'confirmed', label: '확정' },
];

function getTabFilters(tab: TabKey) {
  switch (tab) {
    case 'action_needed':
      return { statuses: ['pending_admin', 'suggested', 'reschedule_requested', 'change_requested'], orderBy: 'received_at_desc' as const };
    case 'confirmed':
      return { status: 'confirmed', dateFilter: 'upcoming' as const, orderBy: 'visit_date_asc' as const };
  }
}

interface Props {
  selectedFieldId?: string | null;
  onFieldSelect?: (id: string | null) => void;
  onSelect?: (id: string) => void;  // 상세 패널 열기
}

export function FieldRequestList({ selectedFieldId, onFieldSelect, onSelect }: Props) {
  const router = useRouter();
  const [tab, setTab] = useState<TabKey>('action_needed');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);

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
  const renderRow = (c: Consultation) => {
    const dday = formatDday(c.visit_date);
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
            {c.visit_date && <span>{c.visit_date} {c.visit_time || ''}</span>}
          </div>
          {/* #5: 주소 + 네비 버튼 항상 표시 */}
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
        </div>

        {/* #3: 인라인 액션 제거 — 클릭 시 슬라이드 패널에서 처리 */}
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

    </div>
  );
}
