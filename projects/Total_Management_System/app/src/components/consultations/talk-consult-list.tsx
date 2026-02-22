'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations, useUpdateConsultationStatus, useStartTalkConsult } from '@/hooks/use-consultations';
import { HoldReasonModal } from './hold-reason-modal';
import {
  formatRelative,
  formatPhone,
  CONSULTATION_STATUS_COLOR,
} from '@/lib/utils/format';
import { Search, ChevronLeft, ChevronRight, MessageCircle } from 'lucide-react';
import type { Consultation } from '@/lib/supabase/types';

type TabKey = 'pending' | 'in_progress' | 'completed' | 'on_hold' | 'cancelled';

const TABS: { key: TabKey; label: string }[] = [
  { key: 'pending', label: '대기' },
  { key: 'in_progress', label: '진행중' },
  { key: 'completed', label: '완료' },
  { key: 'on_hold', label: '보류' },
  { key: 'cancelled', label: '취소' },
];

function getTabStatus(tab: TabKey): string {
  return tab === 'pending' ? 'pending_admin' : tab;
}

export function TalkConsultList() {
  const router = useRouter();
  const [tab, setTab] = useState<TabKey>('pending');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [holdTarget, setHoldTarget] = useState<string | null>(null);

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
              상담 시작
            </Button>
            <Button variant="ghost" size="sm" disabled={busy} onClick={() => setHoldTarget(c.id)}>보류</Button>
            <Button variant="danger" size="sm" disabled={busy} loading={busyStatus === 'cancelled'} onClick={() => updateStatus.mutate({ id: c.id, status: 'cancelled' })}>
              취소
            </Button>
          </>
        );
      case 'in_progress':
        return (
          <>
            <Button variant="primary" size="sm" disabled={busy} loading={busyStatus === 'completed'} onClick={() => updateStatus.mutate({ id: c.id, status: 'completed' })}>
              처리완료
            </Button>
            <Button variant="ghost" size="sm" disabled={busy} onClick={() => setHoldTarget(c.id)}>보류</Button>
            <Button variant="danger" size="sm" disabled={busy} loading={busyStatus === 'cancelled'} onClick={() => updateStatus.mutate({ id: c.id, status: 'cancelled' })}>
              취소
            </Button>
          </>
        );
      case 'on_hold':
        return (
          <>
            <Button variant="secondary" size="sm" disabled={busy} loading={busyStatus === 'pending_admin'} onClick={() => updateStatus.mutate({ id: c.id, status: 'pending_admin' })}>
              대기 복원
            </Button>
            <Button variant="secondary" size="sm" disabled={busy} loading={busyStatus === 'in_progress'} onClick={() => updateStatus.mutate({ id: c.id, status: 'in_progress' })}>
              상담 재개
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
                <div className="cursor-pointer" onClick={() => router.push(`/consultations/${c.id}`)}>
                  <div className="flex items-center gap-2">
                    <MessageCircle size={14} className="text-info shrink-0" />
                    <span className="text-sm font-semibold text-indigo-black truncate">{c.name}</span>
                    <Badge className={CONSULTATION_STATUS_COLOR[c.status]}>
                      {tab === 'pending' ? formatRelative(c.received_at) : ''}
                    </Badge>
                    {c.status === 'on_hold' && c.hold_reason && (
                      <span className="text-xs text-neutral-400 truncate max-w-[120px]" title={c.hold_reason}>
                        ({c.hold_reason})
                      </span>
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

      {/* 모달 */}
      {holdTarget && (
        <HoldReasonModal open={!!holdTarget} onClose={() => setHoldTarget(null)} consultationId={holdTarget} />
      )}
    </div>
  );
}
