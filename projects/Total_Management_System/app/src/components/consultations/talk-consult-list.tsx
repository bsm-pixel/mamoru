'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations, useUpdateConsultationStatus, useStartTalkConsult } from '@/hooks/use-consultations';
import {
  formatRelative,
  formatPhone,
  CONSULTATION_STATUS_COLOR,
} from '@/lib/utils/format';
import { Search, ChevronLeft, ChevronRight, MessageCircle } from 'lucide-react';
import type { Consultation } from '@/lib/supabase/types';

// R2: 4탭 (신규 / 진행중 / 완료 / 취소)
type TabKey = 'pending' | 'in_progress' | 'completed' | 'cancelled';

const TABS: { key: TabKey; label: string }[] = [
  { key: 'pending', label: '신규' },
  { key: 'in_progress', label: '진행중' },
  { key: 'completed', label: '완료' },
  { key: 'cancelled', label: '취소' },
];

function getTabStatus(tab: TabKey): string {
  return tab === 'pending' ? 'pending_admin' : tab;
}

export function TalkConsultList({ onSelect }: { onSelect?: (id: string) => void } = {}) {
  const router = useRouter();
  const [tab, setTab] = useState<TabKey>('pending');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const { data, isLoading } = useConsultations({
    status: getTabStatus(tab),
    type: 'talk_consult',
    search,
    page,
    limit: 20,
    orderBy: (tab === 'completed' || tab === 'cancelled') ? 'updated_at_desc' : undefined,
  });
  const consultations = data?.consultations || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);
  const updateStatus = useUpdateConsultationStatus();
  const startTalk = useStartTalkConsult();

  const renderActions = (c: Consultation) => {
    const busy = (updateStatus.isPending && updateStatus.variables?.id === c.id) ||
                 (startTalk.isPending && startTalk.variables?.id === c.id);
    const busyStatus = (updateStatus.isPending && updateStatus.variables?.id === c.id)
      ? updateStatus.variables?.status : null;
    const isTalkStarting = startTalk.isPending && startTalk.variables?.id === c.id;

    switch (tab) {
      case 'pending':
        return (
          <>
            <Button variant="primary" size="sm" disabled={busy} loading={isTalkStarting} onClick={() => startTalk.mutate({ id: c.id })}>
              상담진행
            </Button>
            <Button variant="danger" size="sm" disabled={busy} loading={busyStatus === 'cancelled'} onClick={() => updateStatus.mutate({ id: c.id, status: 'cancelled' })}>
              취소
            </Button>
          </>
        );
      case 'in_progress':
        return (
          <>
            <Button variant="primary" size="sm" disabled={busy} loading={busyStatus === 'completed'} onClick={() => updateStatus.mutate({ id: c.id, status: 'completed' })}>
              상담완료
            </Button>
            <Button variant="danger" size="sm" disabled={busy} loading={busyStatus === 'cancelled'} onClick={() => updateStatus.mutate({ id: c.id, status: 'cancelled' })}>
              취소
            </Button>
          </>
        );
      default:
        return null;
    }
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
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition ${
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
            {tab === 'pending' ? '대기 중인 톡상담이 없습니다' : '해당 상담이 없습니다'}
          </div>
        ) : (
          <div className="divide-y divide-neutral-100">
            {consultations.map((c) => (
              <div key={c.id} className="px-4 py-3 hover:bg-warm-ivory/60 transition">
                <div className="cursor-pointer" onClick={() => onSelect ? onSelect(c.id) : router.push(`/consultations/${c.id}`)}>
                  <div className="flex items-center gap-2">
                    <MessageCircle size={14} className="text-info shrink-0" />
                    <span className="text-sm font-semibold text-indigo-black truncate">{c.name}</span>
                    {tab === 'pending' && (
                      <Badge className={CONSULTATION_STATUS_COLOR[c.status]}>
                        {formatRelative(c.received_at)}
                      </Badge>
                    )}
                  </div>
                  <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
                    <span>{formatPhone(c.phone)}</span>
                    {tab !== 'pending' && <span>{formatRelative(c.received_at)}</span>}
                  </div>
                  {c.memo && (
                    <p className="mt-1 text-xs text-neutral-500 truncate">{c.memo}</p>
                  )}
                </div>
                <div className="flex gap-1 mt-2 flex-wrap">
                  {renderActions(c)}
                </div>
              </div>
            ))}
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
