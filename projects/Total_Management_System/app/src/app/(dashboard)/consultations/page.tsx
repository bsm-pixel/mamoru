'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultations, useConsultationSync } from '@/hooks/use-consultations';
import {
  formatDate,
  formatPhone,
  CONSULTATION_STATUS_LABEL,
  CONSULTATION_STATUS_COLOR,
  CONSULTATION_TYPE_LABEL,
} from '@/lib/utils/format';
import { RefreshCw, Search, ChevronLeft, ChevronRight } from 'lucide-react';
import type { Consultation } from '@/lib/supabase/types';

const STATUS_TABS = [
  { value: 'all', label: '전체' },
  { value: 'pending_admin', label: '대기중' },
  { value: 'suggested', label: '일정 제안' },
  { value: 'assigned', label: '배정됨' },
  { value: 'confirmed', label: '확정' },
  { value: 'cancelled', label: '취소' },
];

const TYPE_FILTERS = [
  { value: 'all', label: '전체 타입' },
  { value: 'store_visit', label: '매장 방문' },
  { value: 'field_request', label: '출장 요청' },
];

export default function ConsultationsPage() {
  const router = useRouter();
  const [status, setStatus] = useState('all');
  const [type, setType] = useState('all');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const sync = useConsultationSync();

  const { data, isLoading } = useConsultations({ status, type, search, page, limit: 20 });
  const consultations = data?.consultations || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);

  return (
    <>
      <Topbar title="상담관리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 상단: 동기화 + 검색 */}
        <div className="flex items-center gap-3">
          <Button
            variant="secondary"
            size="sm"
            onClick={() => sync.mutate()}
            disabled={sync.isPending}
          >
            <RefreshCw size={14} className={sync.isPending ? 'animate-spin' : ''} />
            {sync.isPending ? '동기화 중...' : 'GAS 동기화'}
          </Button>

          <div className="flex-1 relative">
            <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
            <input
              type="text"
              value={search}
              onChange={(e) => { setSearch(e.target.value); setPage(1); }}
              placeholder="이름, 전화번호, ID 검색"
              className="w-full h-9 pl-9 pr-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm text-indigo-black placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 transition"
            />
          </div>
        </div>

        {/* 타입 필터 */}
        <div className="flex items-center gap-2">
          <select
            value={type}
            onChange={(e) => { setType(e.target.value); setPage(1); }}
            className="h-8 px-3 rounded-lg border border-neutral-200 bg-card-white text-xs text-indigo-black focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          >
            {TYPE_FILTERS.map((f) => (
              <option key={f.value} value={f.value}>{f.label}</option>
            ))}
          </select>
          <span className="text-xs text-neutral-500">
            총 {total}건
          </span>
        </div>

        {/* 상태 탭 */}
        <div className="flex gap-1 overflow-x-auto pb-1">
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
        </div>

        {/* 상담 목록 */}
        <Card padding={false}>
          {isLoading ? (
            <div className="p-4 space-y-3">
              {Array.from({ length: 5 }).map((_, i) => (
                <Skeleton key={i} className="h-16 w-full" />
              ))}
            </div>
          ) : consultations.length === 0 ? (
            <div className="flex items-center justify-center h-40 text-sm text-neutral-400">
              상담 내역이 없습니다
            </div>
          ) : (
            <div className="divide-y divide-neutral-100">
              {consultations.map((c) => (
                <ConsultationRow
                  key={c.id}
                  consultation={c}
                  onClick={() => router.push(`/consultations/${c.id}`)}
                />
              ))}
            </div>
          )}
        </Card>

        {/* 페이지네이션 */}
        {totalPages > 1 && (
          <div className="flex items-center justify-center gap-2">
            <Button
              variant="ghost"
              size="sm"
              disabled={page <= 1}
              onClick={() => setPage(page - 1)}
            >
              <ChevronLeft size={16} />
            </Button>
            <span className="text-sm text-neutral-500">
              {page} / {totalPages}
            </span>
            <Button
              variant="ghost"
              size="sm"
              disabled={page >= totalPages}
              onClick={() => setPage(page + 1)}
            >
              <ChevronRight size={16} />
            </Button>
          </div>
        )}
      </div>
    </>
  );
}

function ConsultationRow({ consultation, onClick }: { consultation: Consultation; onClick: () => void }) {
  const statusColor = CONSULTATION_STATUS_COLOR[consultation.status] || 'bg-neutral-100 text-neutral-500';

  return (
    <div
      onClick={onClick}
      className="flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-warm-ivory/60 transition"
    >
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-indigo-black truncate">
            {consultation.name}
          </span>
          <Badge className={statusColor}>
            {CONSULTATION_STATUS_LABEL[consultation.status] || consultation.status}
          </Badge>
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          <span>{formatPhone(consultation.phone)}</span>
          <span>{CONSULTATION_TYPE_LABEL[consultation.consultation_type]}</span>
          {consultation.visit_date && (
            <span className="text-terracotta">{formatDate(consultation.visit_date)}</span>
          )}
        </div>
      </div>
      <div className="text-right shrink-0 text-xs text-neutral-400">
        {consultation.received_at && formatDate(consultation.received_at)}
      </div>
    </div>
  );
}
