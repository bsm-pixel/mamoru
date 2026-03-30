'use client';

import { useState } from 'react';
import { useConsultation, useUpdateConsultationStatus } from '@/hooks/use-consultations';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { SuggestTimeModal } from './suggest-time-modal';
import { formatPhone, formatDate, CONSULTATION_STATUS_LABEL } from '@/lib/utils/format';
import { Calendar, MapPin, Phone, User, Clock, FileSignature } from 'lucide-react';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import Link from 'next/link';
const TYPE_LABEL: Record<string, string> = {
  store_visit: '매장방문',
  field_request: '출장요청',
  talk_consult: '톡상담',
};

const STATUS_COLOR: Record<string, string> = {
  pending_admin: 'bg-yellow-100 text-yellow-700',
  confirmed: 'bg-blue-100 text-blue-700',
  suggested: 'bg-orange-100 text-orange-700',
  in_progress: 'bg-blue-100 text-blue-700',
  completed: 'bg-green-100 text-green-700',
  cancelled: 'bg-red-100 text-red-700',
  on_hold: 'bg-neutral-100 text-neutral-600',
  reschedule_requested: 'bg-orange-100 text-orange-700',
  change_requested: 'bg-orange-100 text-orange-700',
};

interface Props {
  consultationId: string;
}

export function ConsultationDetailPanel({ consultationId }: Props) {
  const { data, isLoading } = useConsultation(consultationId);
  const updateStatus = useUpdateConsultationStatus();
  const [suggestModalOpen, setSuggestModalOpen] = useState(false);
  const [pendingStatus, setPendingStatus] = useState<string | null>(null);

  if (isLoading) {
    return <div className="space-y-3"><Skeleton className="h-20" /><Skeleton className="h-32" /><Skeleton className="h-20" /></div>;
  }

  if (!data?.consultation) {
    return <p className="text-sm text-neutral-400 text-center py-8">상담 정보를 찾을 수 없습니다</p>;
  }

  const c = data.consultation;
  const history = data.history || [];
  const statusLabel = CONSULTATION_STATUS_LABEL[c.status] || c.status;
  const statusColor = STATUS_COLOR[c.status] || 'bg-neutral-100';

  const handleStatus = (newStatus: string) => {
    setPendingStatus(newStatus);
  };

  const confirmStatusLabels: Record<string, { title: string; msg: string; variant?: 'danger' | 'default' }> = {
    completed: { title: '상담 완료', msg: `${c.name}님의 상담을 완료 처리합니다.` },
    in_progress: { title: '상담 시작', msg: `${c.name}님의 상담을 시작합니다.` },
    cancelled: { title: '상담 취소', msg: `${c.name}님의 상담을 취소합니다. 이 작업은 되돌릴 수 없습니다.`, variant: 'danger' },
  };

  return (
    <div className="space-y-4">
      {/* 헤더 */}
      <div>
        <div className="flex items-center gap-2 mb-1">
          <h3 className="text-base font-bold">{c.name}</h3>
          <Badge className={statusColor}>{statusLabel}</Badge>
        </div>
        <div className="flex items-center gap-2 text-xs text-neutral-500">
          <Badge className="bg-neutral-100 text-neutral-600">{TYPE_LABEL[c.consultation_type] || c.consultation_type}</Badge>
          <span>{formatDate(c.received_at)}</span>
        </div>
      </div>

      {/* 고객 정보 */}
      <div className="bg-neutral-50 rounded-lg p-3 space-y-2">
        <div className="flex items-center gap-2 text-sm">
          <User size={14} className="text-neutral-400" />
          <span className="font-medium">{c.name}</span>
        </div>
        {c.phone && (
          <div className="flex items-center gap-2 text-sm">
            <Phone size={14} className="text-neutral-400" />
            <a href={`tel:${c.phone}`} className="text-blue-600">{formatPhone(c.phone)}</a>
          </div>
        )}
        {c.visit_date && (
          <div className="flex items-center gap-2 text-sm">
            <Calendar size={14} className="text-neutral-400" />
            <span>{c.visit_date} {c.visit_time || ''}</span>
          </div>
        )}
        {c.address_road && (
          <div className="flex items-center gap-2 text-sm">
            <MapPin size={14} className="text-neutral-400" />
            <span>{c.address_road} {c.address_detail || ''}</span>
          </div>
        )}
      </div>

      {/* 출장: 가능요일/선호시간대 */}
      {c.consultation_type === 'field_request' && (() => {
        const raw = c.gas_raw as any;
        const prefDays = (raw?.days as string)?.split(',').filter(Boolean) || [];
        const prefTimes = (raw?.timePrefs as string)?.split(',').filter(Boolean) || [];
        if (prefDays.length === 0 && prefTimes.length === 0) return null;
        return (
          <div className="flex flex-wrap items-center gap-1.5 text-xs">
            {prefDays.length > 0 && (
              <>
                <span className="text-neutral-500 font-medium">가능요일:</span>
                {prefDays.map(d => (
                  <span key={d} className="px-1.5 py-0.5 rounded bg-blue-50 text-blue-600 font-medium">{d}</span>
                ))}
              </>
            )}
            {prefTimes.length > 0 && (
              <>
                <span className="text-neutral-500 font-medium ml-1">시간대:</span>
                {prefTimes.map(t => (
                  <span key={t} className="px-1.5 py-0.5 rounded bg-amber-50 text-amber-600 font-medium">{t}</span>
                ))}
              </>
            )}
          </div>
        );
      })()}

      {/* 메모 — 접수메모 / 재요청메모 분리 */}
      {c.memo && (() => {
        const lines = c.memo.split('\n');
        const reRequestLines = lines.filter(l => l.includes('[고객 재요청'));
        const originalLines = lines.filter(l => !l.includes('[고객 재요청'));
        const originalMemo = originalLines.join('\n').trim();
        const reRequestMemo = reRequestLines.map(l => l.replace(/^\d+\s*/, '').trim()).filter(Boolean);

        return (
          <div className="space-y-2">
            {originalMemo && (
              <div className="bg-white border border-neutral-200 rounded-lg p-3">
                <p className="text-xs text-neutral-500 mb-1">접수 메모</p>
                <p className="text-sm whitespace-pre-wrap">{originalMemo}</p>
              </div>
            )}
            {reRequestMemo.length > 0 && (
              <div className="bg-amber-50 border border-amber-200 rounded-lg p-3">
                <p className="text-xs text-amber-600 font-medium mb-1">재요청 메모</p>
                {reRequestMemo.map((line, i) => (
                  <p key={i} className="text-sm text-amber-800">{line}</p>
                ))}
              </div>
            )}
          </div>
        );
      })()}

      {/* 액션 버튼 — #3: 모든 액션을 패널 내부에 집중 */}
      {!['completed', 'cancelled'].includes(c.status) && (
        <div className="space-y-2 pt-2 border-t border-neutral-100">
          {/* 출장: 신규/재요청 → 시간 제안 모달 */}
          {c.consultation_type === 'field_request' && ['pending_admin', 'reschedule_requested', 'change_requested'].includes(c.status) && (
            <Button size="sm" className="w-full" onClick={() => setSuggestModalOpen(true)} disabled={updateStatus.isPending}>
              <Clock size={14} />
              시간 제안
            </Button>
          )}
          {/* 확정 → 완료 */}
          {c.status === 'confirmed' && (
            <Button size="sm" className="w-full" onClick={() => handleStatus('completed')} disabled={updateStatus.isPending}>
              상담 완료
            </Button>
          )}
          {/* 톡: 신규 → 진행 */}
          {c.status === 'pending_admin' && c.consultation_type === 'talk_consult' && (
            <Button size="sm" className="w-full" onClick={() => handleStatus('in_progress')} disabled={updateStatus.isPending}>
              상담 시작
            </Button>
          )}
          {/* 진행중 → 완료 */}
          {c.status === 'in_progress' && (
            <Button size="sm" className="w-full" onClick={() => handleStatus('completed')} disabled={updateStatus.isPending}>
              상담 완료
            </Button>
          )}
          <Button variant="ghost" size="sm" className="w-full text-red-600" onClick={() => handleStatus('cancelled')} disabled={updateStatus.isPending}>
            취소
          </Button>
        </div>
      )}

      {/* 완료 시 계약서 CTA */}
      {c.status === 'completed' && (
        <Link
          href={`/contracts/new?customer_name=${encodeURIComponent(c.name)}&customer_phone=${encodeURIComponent(c.phone || '')}`}
          className="flex items-center gap-2 p-3 rounded-lg bg-green-50 border border-green-200 hover:bg-green-100 transition"
        >
          <FileSignature size={16} className="text-green-600" />
          <span className="text-sm font-medium text-green-700">계약서 작성하기</span>
        </Link>
      )}

      {/* 상세 페이지 링크 */}
      <Link
        href={`/consultations/${c.id}`}
        className="block text-center text-xs text-neutral-400 hover:text-neutral-600 py-2"
      >
        상세 페이지에서 보기 →
      </Link>

      {/* 상태 변경 확인 모달 */}
      {pendingStatus && confirmStatusLabels[pendingStatus] && (
        <ConfirmModal
          open={!!pendingStatus}
          onClose={() => setPendingStatus(null)}
          onConfirm={() => updateStatus.mutateAsync({ id: c.id, status: pendingStatus })}
          title={confirmStatusLabels[pendingStatus].title}
          message={confirmStatusLabels[pendingStatus].msg}
          confirmLabel={confirmStatusLabels[pendingStatus].title}
          variant={confirmStatusLabels[pendingStatus].variant || 'default'}
        />
      )}

      {/* 이력 */}
      {history.length > 0 && (
        <div className="pt-2 border-t border-neutral-100">
          <p className="text-xs font-semibold text-neutral-500 mb-2">상태 이력</p>
          <div className="space-y-1.5">
            {history.slice(0, 5).map((h) => (
              <div key={h.id} className="flex items-center justify-between text-xs">
                <span className="text-neutral-500">
                  {(h.from_status && CONSULTATION_STATUS_LABEL[h.from_status]) || h.from_status || '-'} → {CONSULTATION_STATUS_LABEL[h.to_status] || h.to_status}
                </span>
                <span className="text-neutral-400">{formatDate(h.created_at)}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 시간 제안 모달 */}
      {c.consultation_type === 'field_request' && (
        <SuggestTimeModal
          open={suggestModalOpen}
          onClose={() => setSuggestModalOpen(false)}
          consultationId={c.id}
          prefDays={((c.gas_raw as any)?.days as string)?.split(',').filter(Boolean)}
          prefTimes={((c.gas_raw as any)?.timePrefs as string)?.split(',').filter(Boolean)}
        />
      )}
    </div>
  );
}
