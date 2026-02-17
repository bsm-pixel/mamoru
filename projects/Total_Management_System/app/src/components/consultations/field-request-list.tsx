'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations, useUpdateConsultationStatus } from '@/hooks/use-consultations';
import { SuggestTimeModal } from './suggest-time-modal';
import { HoldReasonModal } from './hold-reason-modal';
import {
  formatDate,
  formatPhone,
  CONSULTATION_STATUS_LABEL,
  CONSULTATION_STATUS_COLOR,
} from '@/lib/utils/format';
import { MapPin, Search, ChevronLeft, ChevronRight, Map, List } from 'lucide-react';
import type { Consultation } from '@/lib/supabase/types';

const STATUS_TABS = [
  { value: 'pending_admin', label: '신규접수' },
  { value: 'suggested', label: '시간제안완료' },
  { value: 'reschedule_requested', label: '재제안요청' },
  { value: 'confirmed', label: '확정' },
  { value: 'on_hold', label: '보류' },
  { value: 'cancelled', label: '취소' },
];

interface Props {
  onToggleMap?: () => void;
  showMap?: boolean;
}

export function FieldRequestList({ onToggleMap, showMap }: Props) {
  const router = useRouter();
  const [status, setStatus] = useState('pending_admin');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [suggestTarget, setSuggestTarget] = useState<string | null>(null);
  const [holdTarget, setHoldTarget] = useState<string | null>(null);

  const { data, isLoading } = useConsultations({
    status,
    type: 'field_request',
    search,
    page,
    limit: 20,
  });
  const consultations = data?.consultations || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);
  const updateStatus = useUpdateConsultationStatus();

  const renderActions = (c: Consultation) => {
    switch (c.status) {
      case 'pending_admin':
        return (
          <>
            <Button variant="primary" size="sm" onClick={() => setSuggestTarget(c.id)}>
              시간 제안
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
      case 'reschedule_requested':
        return (
          <>
            <Button variant="primary" size="sm" onClick={() => setSuggestTarget(c.id)}>
              새 시간 제안
            </Button>
            <Button variant="secondary" size="sm" onClick={() => updateStatus.mutate({ id: c.id, status: 'confirmed' })}>
              즉시 확정
            </Button>
            <Button variant="ghost" size="sm" onClick={() => setHoldTarget(c.id)}>보류</Button>
          </>
        );
      case 'confirmed':
        return (
          <>
            <Button variant="ghost" size="sm" onClick={() => setHoldTarget(c.id)}>보류</Button>
            <Button variant="danger" size="sm" onClick={() => updateStatus.mutate({ id: c.id, status: 'cancelled' })}>
              취소
            </Button>
          </>
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
          <Button
            variant={showMap ? 'primary' : 'secondary'}
            size="sm"
            onClick={onToggleMap}
          >
            {showMap ? <List size={14} /> : <Map size={14} />}
            {showMap ? '리스트' : '지도'}
          </Button>
        )}
      </div>

      {/* 상태 탭 */}
      <div className="flex items-center gap-1 overflow-x-auto pb-1">
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
        ) : (
          <div className="divide-y divide-neutral-100">
            {consultations.map((c) => (
              <div key={c.id} className="px-4 py-3 hover:bg-warm-ivory/60 transition">
                <div
                  className="cursor-pointer"
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
                    {c.address_road && (
                      <span className="flex items-center gap-1">
                        <MapPin size={11} />
                        {c.address_sido} {c.address_sigungu}
                      </span>
                    )}
                    {c.visit_date && (
                      <span className="text-terracotta">{formatDate(c.visit_date, 'M/d')}</span>
                    )}
                  </div>
                </div>
                {/* 액션 버튼 */}
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
      {suggestTarget && (
        <SuggestTimeModal
          open={!!suggestTarget}
          onClose={() => setSuggestTarget(null)}
          consultationId={suggestTarget}
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
