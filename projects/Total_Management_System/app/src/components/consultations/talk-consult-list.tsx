'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations } from '@/hooks/use-consultations';
import { activityDisplay } from '@/lib/customer/display';
import { useActivityTypes } from '@/hooks/use-activity-types';
import { ActivityChips } from '@/components/shared/activity-chips';
import {
  formatRelative,
  formatPhone,
  CONSULTATION_STATUS_COLOR,
  CONSULTATION_STATUS_LABEL,
} from '@/lib/utils/format';
import { Search, ChevronLeft, ChevronRight, MessageCircle } from 'lucide-react';

// 3탭: 신규접수 / 진행중 / 완료
type TabKey = 'new_intake' | 'in_progress' | 'completed';

const TABS: { key: TabKey; label: string }[] = [
  { key: 'new_intake', label: '신규접수' },
  { key: 'in_progress', label: '진행중' },
  { key: 'completed', label: '완료' },
];

function getTabFilters(tab: TabKey) {
  switch (tab) {
    case 'new_intake':
      return { status: 'pending_admin', orderBy: 'received_at_desc' as const };
    case 'in_progress':
      return { status: 'in_progress', orderBy: 'received_at_desc' as const };
    case 'completed':
      return { statuses: ['completed', 'cancelled'], orderBy: 'updated_at_desc' as const };
  }
}

export function TalkConsultList({ onSelect }: { onSelect?: (id: string) => void } = {}) {
  const router = useRouter();
  const [tab, setTab] = useState<TabKey>('new_intake');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const tabFilters = getTabFilters(tab);
  const { data, isLoading } = useConsultations({
    ...(tabFilters || {}),
    type: 'talk_consult',
    search,
    page,
    limit: 20,
  });
  const consultations = data?.consultations || [];
  const consultActTypes = useActivityTypes(consultations.map((c) => c.phone));
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);

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
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition ${
              tab === t.key
                ? 'bg-stone-900 text-white'
                : 'bg-stone-100 text-stone-600 hover:bg-stone-200'
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
            {tab === 'new_intake' ? '신규 톡상담이 없습니다' : tab === 'in_progress' ? '진행중인 톡상담이 없습니다' : '완료된 톡상담이 없습니다'}
          </div>
        ) : (
          <div className="divide-y divide-neutral-100">
            {consultations.map((c) => {
              const isCancelled = c.status === 'cancelled';
              const isCompleted = c.status === 'completed';
              return (
              <div key={c.id} className={`px-4 py-3 hover:bg-warm-ivory/60 transition ${
                isCancelled ? 'opacity-60 bg-neutral-50 border-l-2 border-l-neutral-300'
                : isCompleted ? 'border-l-2 border-l-green-400'
                : ''
              }`}>
                <div className="cursor-pointer" onClick={() => onSelect ? onSelect(c.id) : router.push(`/consultations/${c.id}`)}>
                  <div className="flex items-center gap-2">
                    <MessageCircle size={14} className={isCancelled ? 'text-neutral-400 shrink-0' : 'text-info shrink-0'} />
                    <span className={`text-sm font-semibold truncate ${isCancelled ? 'line-through text-neutral-400' : 'text-indigo-black'}`}>{activityDisplay(c.activity_name, c.name)}</span>
                    <ActivityChips types={consultActTypes(c.phone)} className="shrink-0" />
                    <Badge className={CONSULTATION_STATUS_COLOR[c.status]}>
                      {tab === 'new_intake' ? formatRelative(c.received_at) : (CONSULTATION_STATUS_LABEL[c.status] || c.status)}
                    </Badge>
                  </div>
                  <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
                    <span>{formatPhone(c.phone)}</span>
                    {tab === 'completed' && <span>{formatRelative(c.received_at)}</span>}
                  </div>
                  {c.memo && (
                    <p className="mt-1 text-xs text-neutral-500 truncate">{c.memo}</p>
                  )}
                </div>
              </div>
              );
            })}
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
