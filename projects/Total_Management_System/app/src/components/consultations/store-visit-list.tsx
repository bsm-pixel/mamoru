'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations, useSendNotification, useUpdateConsultationStatus } from '@/hooks/use-consultations';
import { RescheduleModal } from './reschedule-modal';
import { HoldReasonModal } from './hold-reason-modal';
import {
  formatDate,
  formatPhone,
  CONSULTATION_STATUS_LABEL,
  CONSULTATION_STATUS_COLOR,
} from '@/lib/utils/format';
import { Calendar, Search, ChevronLeft, ChevronRight } from 'lucide-react';
import type { Consultation } from '@/lib/supabase/types';

const STATUS_TABS = [
  { value: 'confirmed', label: '확정' },
  { value: 'on_hold', label: '보류' },
  { value: 'cancelled', label: '취소' },
];

export function StoreVisitList() {
  const router = useRouter();
  const [status, setStatus] = useState('confirmed');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [rescheduleTarget, setRescheduleTarget] = useState<Consultation | null>(null);
  const [holdTarget, setHoldTarget] = useState<string | null>(null);

  const { data, isLoading } = useConsultations({
    status,
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

      {/* 상태 탭 */}
      <div className="flex items-center gap-1">
        {STATUS_TABS.map((tab) => (
          <button
            key={tab.value}
            onClick={() => { setStatus(tab.value); setPage(1); }}
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition ${
              status === tab.value
                ? 'bg-terracotta text-cream'
                : 'bg-card-white text-neutral-500 hover:bg-warm-ivory'
            }`}
          >
            {tab.label}
          </button>
        ))}
        <span className="ml-auto text-xs text-neutral-500">{total}건</span>
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
            해당 상담이 없습니다
          </div>
        ) : (
          <div className="divide-y divide-neutral-100">
            {consultations.map((c) => (
              <div
                key={c.id}
                className="flex items-center gap-3 px-4 py-3 hover:bg-warm-ivory/60 transition"
              >
                <div
                  className="flex-1 min-w-0 cursor-pointer"
                  onClick={() => router.push(`/consultations/${c.id}`)}
                >
                  <div className="flex items-center gap-2">
                    <span className="text-sm font-semibold text-indigo-black truncate">{c.name}</span>
                    <Badge className={CONSULTATION_STATUS_COLOR[c.status]}>
                      {CONSULTATION_STATUS_LABEL[c.status]}
                    </Badge>
                    {c.status === 'on_hold' && c.hold_reason && (
                      <span className="text-xs text-neutral-400 truncate max-w-[120px]" title={c.hold_reason}>
                        ({c.hold_reason})
                      </span>
                    )}
                  </div>
                  <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
                    <span>{formatPhone(c.phone)}</span>
                    {c.visit_date && (
                      <span className="flex items-center gap-1 text-terracotta">
                        <Calendar size={12} />
                        {formatDate(c.visit_date, 'M/d')} {c.visit_time || ''}
                      </span>
                    )}
                  </div>
                </div>
                {/* 액션 버튼 */}
                <div className="flex gap-1 shrink-0">
                  {c.status === 'confirmed' && (
                    <>
                      <Button
                        variant="secondary"
                        size="sm"
                        onClick={() => setRescheduleTarget(c)}
                      >
                        일정변경
                      </Button>
                      <Button
                        variant="ghost"
                        size="sm"
                        disabled={sendNotify.isPending}
                        onClick={() => sendNotify.mutate({ consultationId: c.id, template: 'confirmed' })}
                      >
                        알림 재발송
                      </Button>
                      <Button
                        variant="ghost"
                        size="sm"
                        onClick={() => setHoldTarget(c.id)}
                      >
                        보류
                      </Button>
                    </>
                  )}
                  {c.status === 'on_hold' && (
                    <>
                      <Button
                        variant="secondary"
                        size="sm"
                        onClick={() => updateStatus.mutate({ id: c.id, status: 'confirmed' })}
                      >
                        확정 복원
                      </Button>
                      <Button
                        variant="danger"
                        size="sm"
                        onClick={() => updateStatus.mutate({ id: c.id, status: 'cancelled' })}
                      >
                        취소
                      </Button>
                    </>
                  )}
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
        <HoldReasonModal
          open={!!holdTarget}
          onClose={() => setHoldTarget(null)}
          consultationId={holdTarget}
        />
      )}
    </div>
  );
}
