'use client';

import { useState } from 'react';
import { useParams, useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useConsultation, useUpdateConsultationStatus, useAssignDealer } from '@/hooks/use-consultations';
import { useDealers } from '@/hooks/use-dealers';
import {
  formatDate,
  formatDateTime,
  formatPhone,
  formatRelative,
  CONSULTATION_STATUS_LABEL,
  CONSULTATION_STATUS_COLOR,
  CONSULTATION_TYPE_LABEL,
} from '@/lib/utils/format';
import { ArrowLeft, MapPin, Calendar, User, Clock, UserCheck } from 'lucide-react';
import type { ConsultationStatus } from '@/lib/supabase/types';

/** 상태 전이 가능 맵 */
const STATUS_TRANSITIONS: Record<string, { value: ConsultationStatus; label: string }[]> = {
  pending_admin: [
    { value: 'suggested', label: '일정 제안' },
    { value: 'assigned', label: '딜러 배정' },
    { value: 'cancelled', label: '취소' },
  ],
  suggested: [
    { value: 'confirmed', label: '확정' },
    { value: 'reschedule_requested', label: '일정 변경' },
    { value: 'cancelled', label: '취소' },
  ],
  assigned: [
    { value: 'confirmed', label: '확정' },
    { value: 'reschedule_requested', label: '일정 변경' },
    { value: 'cancelled', label: '취소' },
  ],
  reschedule_requested: [
    { value: 'suggested', label: '재제안' },
    { value: 'assigned', label: '딜러 재배정' },
    { value: 'cancelled', label: '취소' },
  ],
};

export default function ConsultationDetailPage() {
  const params = useParams();
  const router = useRouter();
  const id = params.id as string;
  const { data, isLoading } = useConsultation(id);
  const updateStatus = useUpdateConsultationStatus();
  const assignDealer = useAssignDealer();
  const { data: dealers } = useDealers();
  const [selectedDealer, setSelectedDealer] = useState('');

  if (isLoading) {
    return (
      <>
        <Topbar title="상담 상세" />
        <div className="px-4 md:px-6 py-4 space-y-4">
          <Skeleton className="h-8 w-32" />
          <Skeleton className="h-40 w-full" />
          <Skeleton className="h-40 w-full" />
        </div>
      </>
    );
  }

  if (!data) {
    return (
      <>
        <Topbar title="상담 상세" />
        <div className="flex items-center justify-center h-60 text-neutral-400">
          상담을 찾을 수 없습니다
        </div>
      </>
    );
  }

  const { consultation: c, history } = data;
  const statusColor = CONSULTATION_STATUS_COLOR[c.status] || 'bg-neutral-100 text-neutral-500';
  const transitions = STATUS_TRANSITIONS[c.status] || [];

  return (
    <>
      <Topbar title="상담 상세" />

      <div className="px-4 md:px-6 py-4 space-y-4 max-w-3xl">
        {/* 뒤로가기 */}
        <button
          onClick={() => router.back()}
          className="flex items-center gap-1 text-sm text-neutral-500 hover:text-indigo-black transition"
        >
          <ArrowLeft size={16} />
          목록으로
        </button>

        {/* 상담 헤더 */}
        <div className="flex items-center justify-between">
          <div>
            <h2 className="text-lg font-bold">{c.name}</h2>
            <p className="text-xs text-neutral-500 mt-0.5">
              {c.unique_id} &middot; {formatDateTime(c.received_at)}
            </p>
          </div>
          <Badge className={statusColor}>
            {CONSULTATION_STATUS_LABEL[c.status] || c.status}
          </Badge>
        </div>

        {/* 고객 정보 */}
        <Card>
          <CardHeader>
            <CardTitle>
              <User size={16} className="inline mr-1.5" />
              고객 정보
            </CardTitle>
          </CardHeader>
          <dl className="grid grid-cols-2 gap-y-2 text-sm">
            <dt className="text-neutral-500">이름</dt>
            <dd className="font-medium">{c.name}</dd>
            <dt className="text-neutral-500">전화</dt>
            <dd>{formatPhone(c.phone)}</dd>
            <dt className="text-neutral-500">상담 타입</dt>
            <dd>{CONSULTATION_TYPE_LABEL[c.consultation_type]}</dd>
          </dl>
        </Card>

        {/* 일정 정보 */}
        <Card>
          <CardHeader>
            <CardTitle>
              <Calendar size={16} className="inline mr-1.5" />
              일정
            </CardTitle>
          </CardHeader>
          <dl className="grid grid-cols-2 gap-y-2 text-sm">
            <dt className="text-neutral-500">방문 희망일</dt>
            <dd className="font-medium">
              {c.visit_date ? formatDate(c.visit_date, 'yyyy년 M월 d일') : '-'}
            </dd>
            <dt className="text-neutral-500">방문 시간</dt>
            <dd>{c.visit_time || '-'}</dd>
            {c.suggestions && (c.suggestions as { dates?: string[] }).dates && (
              <>
                <dt className="text-neutral-500">제안 일정</dt>
                <dd className="col-span-1">
                  {((c.suggestions as { dates: string[] }).dates).map((d, i) => (
                    <span key={i} className="block text-xs text-info">{d}</span>
                  ))}
                </dd>
              </>
            )}
          </dl>
        </Card>

        {/* 주소 정보 */}
        <Card>
          <CardHeader>
            <CardTitle>
              <MapPin size={16} className="inline mr-1.5" />
              주소
            </CardTitle>
          </CardHeader>
          <dl className="grid grid-cols-2 gap-y-2 text-sm">
            <dt className="text-neutral-500">우편번호</dt>
            <dd>{c.postcode || '-'}</dd>
            <dt className="text-neutral-500">도로명</dt>
            <dd className="col-span-2">{c.address_road || '-'} {c.address_detail || ''}</dd>
            <dt className="text-neutral-500">지역</dt>
            <dd>{[c.address_sido, c.address_sigungu, c.address_region].filter(Boolean).join(' ') || '-'}</dd>
          </dl>
        </Card>

        {/* 딜러 배정 */}
        <Card>
          <CardHeader>
            <CardTitle>
              <UserCheck size={16} className="inline mr-1.5" />
              딜러 배정
            </CardTitle>
          </CardHeader>
          {c.dealer_id ? (
            <p className="text-sm font-medium text-success">딜러 배정 완료</p>
          ) : (
            <div className="flex items-center gap-2">
              <select
                value={selectedDealer}
                onChange={(e) => setSelectedDealer(e.target.value)}
                className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
              >
                <option value="">딜러 선택</option>
                {dealers?.map((d) => (
                  <option key={d.id} value={d.id}>{d.name} ({d.dealer_code})</option>
                ))}
              </select>
              <Button
                size="sm"
                disabled={!selectedDealer || assignDealer.isPending}
                onClick={() => {
                  assignDealer.mutate({ consultationId: c.id, dealerId: selectedDealer });
                  setSelectedDealer('');
                }}
              >
                {assignDealer.isPending ? '배정 중...' : '배정'}
              </Button>
            </div>
          )}
        </Card>

        {/* 메모 */}
        {c.memo && (
          <Card>
            <CardHeader>
              <CardTitle>메모</CardTitle>
            </CardHeader>
            <p className="text-sm text-neutral-700 whitespace-pre-wrap">{c.memo}</p>
          </Card>
        )}

        {/* 빠른 액션: 상태 변경 */}
        {transitions.length > 0 && (
          <Card>
            <CardHeader>
              <CardTitle>상태 변경</CardTitle>
            </CardHeader>
            <div className="flex flex-wrap gap-2">
              {transitions.map((t) => (
                <Button
                  key={t.value}
                  variant={t.value === 'cancelled' ? 'danger' : 'secondary'}
                  size="sm"
                  disabled={updateStatus.isPending}
                  onClick={() => updateStatus.mutate({ id: c.id, status: t.value })}
                >
                  {t.label}
                </Button>
              ))}
            </div>
          </Card>
        )}

        {/* 상태 이력 */}
        <Card>
          <CardHeader>
            <CardTitle>
              <Clock size={16} className="inline mr-1.5" />
              상태 이력
            </CardTitle>
          </CardHeader>
          {history.length === 0 ? (
            <p className="text-sm text-neutral-400">이력 없음</p>
          ) : (
            <div className="space-y-3">
              {history.map((h) => (
                <div key={h.id} className="flex items-start gap-3 text-sm">
                  <div className="w-2 h-2 rounded-full bg-terracotta mt-1.5 shrink-0" />
                  <div>
                    <div className="flex items-center gap-2">
                      {h.from_status && (
                        <>
                          <span className="text-neutral-500">
                            {CONSULTATION_STATUS_LABEL[h.from_status]}
                          </span>
                          <span className="text-neutral-400">&rarr;</span>
                        </>
                      )}
                      <span className="font-medium">
                        {CONSULTATION_STATUS_LABEL[h.to_status]}
                      </span>
                    </div>
                    {h.note && <p className="text-xs text-neutral-500 mt-0.5">{h.note}</p>}
                    <p className="text-xs text-neutral-400 mt-0.5">
                      {formatRelative(h.created_at)}
                    </p>
                  </div>
                </div>
              ))}
            </div>
          )}
        </Card>
      </div>
    </>
  );
}
