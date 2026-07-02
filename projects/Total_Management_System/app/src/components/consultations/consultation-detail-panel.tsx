'use client';

import { useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import toast from 'react-hot-toast';
import { useConsultation, useUpdateConsultationStatus, useUpdateConsultationAdminNote, useDeleteConsultation } from '@/hooks/use-consultations';
import { activityDisplay, honorific } from '@/lib/customer/display';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { SuggestTimeModal } from './suggest-time-modal';
import { ManualConfirmModal } from './manual-confirm-modal';
import { formatPhone, formatDate, CONSULTATION_STATUS_LABEL } from '@/lib/utils/format';
import { Calendar, MapPin, Phone, User, Clock, FileSignature, ShoppingCart, CheckCircle2, CheckCircle } from 'lucide-react';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import { Modal } from '@/components/ui/modal';
import Link from 'next/link';
const TYPE_LABEL: Record<string, string> = {
  store_visit: '매장방문',
  field_request: '출장요청',
  talk_consult: '온라인상담',
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
  /** 077: 상담완료 처리 후 호출 (예: 매장방문 탭에서 "지난내역"으로 자동 전환) */
  onAfterComplete?: () => void;
}

export function ConsultationDetailPanel({ consultationId, onAfterComplete }: Props) {
  const router = useRouter();
  const { data, isLoading, refetch } = useConsultation(consultationId);
  const updateStatus = useUpdateConsultationStatus();
  const updateAdminNote = useUpdateConsultationAdminNote();
  const deleteConsultation = useDeleteConsultation();
  const [suggestModalOpen, setSuggestModalOpen] = useState(false);
  const [manualConfirmOpen, setManualConfirmOpen] = useState(false);
  const [pendingStatus, setPendingStatus] = useState<string | null>(null);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  // 077: 상담완료 시 판매전환 분기 모달 (linkedSales 없을 때만)
  const [saleConvertOpen, setSaleConvertOpen] = useState(false);
  // 108: 상담자 메모 로컬 편집 상태 — 선택 상담이 바뀔 때만 서버값으로 초기화(재조회로 타이핑 중 값 안 날아가게 id 기준)
  const [adminNote, setAdminNote] = useState('');
  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => { setAdminNote(data?.consultation?.admin_note ?? ''); }, [data?.consultation?.id]);

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
  const hasLinkedSales = !!(data?.linkedSales && data.linkedSales.length > 0);
  const adminNoteDirty = adminNote !== (c.admin_note ?? '');

  const handleStatus = (newStatus: string) => {
    // 077: 상담완료 + 판매미연결 + 톡상담 아닌 경우 → 판매전환 분기 모달
    if (newStatus === 'completed' && !hasLinkedSales && c.consultation_type !== 'talk_consult') {
      setSaleConvertOpen(true);
      return;
    }
    setPendingStatus(newStatus);
  };

  // 077: 상담완료 mutation 헬퍼 — 성공 시 토스트 + onAfterComplete 호출
  const completeOnly = async () => {
    await updateStatus.mutateAsync({ id: c.id, status: 'completed' });
    toast.success('상담완료 처리됨 — 지난내역 탭에서 확인 가능');
    onAfterComplete?.();
  };

  const confirmStatusLabels: Record<string, { title: string; msg: string; variant?: 'danger' | 'default' }> = {
    completed: { title: '상담 완료', msg: `${honorific(c.activity_name, c.name)}님의 상담을 완료 처리합니다.` },
    in_progress: { title: '상담 시작', msg: `${honorific(c.activity_name, c.name)}님의 상담을 시작합니다.` },
    cancelled: { title: '상담 취소', msg: `${honorific(c.activity_name, c.name)}님의 상담을 취소합니다. 이 작업은 되돌릴 수 없습니다.`, variant: 'danger' },
  };

  return (
    <div className="space-y-4">
      {/* 헤더 */}
      <div>
        <div className="flex items-center gap-2 mb-1">
          <h3 className="text-base font-bold">{activityDisplay(c.activity_name, c.name)}</h3>
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

      {/* 상담자 메모 (관리자 전용) — 고객 비노출 · 구글캘린더 설명란 반영 (108) */}
      <div className="bg-white border border-neutral-200 rounded-lg p-3">
        <div className="flex items-center justify-between mb-1.5">
          <p className="text-xs font-medium text-neutral-700">📝 상담자 메모 <span className="font-normal text-neutral-400">(관리자 전용)</span></p>
          <button
            onClick={() => updateAdminNote.mutate(
              { id: c.id, adminNote },
              { onSuccess: () => { toast.success('메모 저장됨'); refetch(); } },
            )}
            disabled={!adminNoteDirty || updateAdminNote.isPending}
            className="text-xs font-semibold text-stone-900 disabled:text-neutral-300"
          >
            {updateAdminNote.isPending ? '저장 중...' : '저장'}
          </button>
        </div>
        <textarea
          value={adminNote}
          onChange={(e) => setAdminNote(e.target.value)}
          rows={3}
          placeholder="유선·DM·톡 요청, 참고사항 등 — 고객에게 안 보이며 구글캘린더 일정에 표기됩니다"
          className="w-full px-2.5 py-2 rounded-lg border border-neutral-200 bg-stone-50 text-sm resize-none focus:outline-none focus:ring-2 focus:ring-stone-400 placeholder:text-neutral-400"
        />
        {c.consultation_type === 'talk_consult' && (
          <p className="mt-1 text-[10px] text-neutral-400">※ 톡상담은 캘린더 일정이 없어 이 메모는 TMS에만 표시됩니다.</p>
        )}
      </div>

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
          {/* 079: 수동 일정 확정 — DM/유선 협의 후 즉시 확정 (매장/출장 한정, 톡상담 제외). 이미 확정된 건이면 라벨은 "변경" */}
          {c.consultation_type !== 'talk_consult' && c.status !== 'in_progress' && (
            <Button size="sm" variant="secondary" className="w-full" onClick={() => setManualConfirmOpen(true)} disabled={updateStatus.isPending}>
              <CheckCircle size={14} />
              {c.status === 'confirmed' ? '수동 일정 변경' : '수동 일정 확정'}
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
          <button
            onClick={() => setShowDeleteConfirm(true)}
            className="w-full text-center text-xs text-neutral-400 hover:text-red-500 py-1.5 transition"
          >
            삭제 (알림 없이 제거)
          </button>
        </div>
      )}

      {/* 취소/완료 상태에서도 삭제 가능 */}
      {['completed', 'cancelled'].includes(c.status) && (
        <button
          onClick={() => setShowDeleteConfirm(true)}
          className="w-full text-center text-xs text-neutral-400 hover:text-red-500 py-1.5 transition"
        >
          삭제 (알림 없이 제거)
        </button>
      )}

      {/* 070: 이 상담으로 처리된 판매 노출 (역방향 link) */}
      {data?.linkedSales && data.linkedSales.length > 0 && (
        <div className="flex flex-wrap items-center gap-1.5 text-xs">
          <span className="text-neutral-500 font-medium">이 상담으로 판매 처리:</span>
          {data.linkedSales.map((s) => (
            <Link
              key={s.id}
              href={`/sales/${s.id}`}
              className="px-1.5 py-0.5 rounded bg-green-50 text-green-700 font-semibold hover:bg-green-100 transition"
            >
              {s.sale_number}
            </Link>
          ))}
        </div>
      )}

      {/* 070: 판매로 처리 CTA — 톡상담 제외, 확정~완료 단계에서 표시 */}
      {/* 077: linkedSales 있으면 "판매연결완료" 배지로 대체 (중복 처리 방지) */}
      {['confirmed', 'in_progress', 'completed'].includes(c.status) && c.consultation_type !== 'talk_consult' && (
        data?.linkedSales && data.linkedSales.length > 0 ? (
          <div className="flex items-center gap-2 p-3 rounded-lg bg-green-50 border border-green-200">
            <CheckCircle2 size={16} className="text-green-600" />
            <span className="text-sm font-medium text-green-700">
              판매연결완료 ({data.linkedSales.length}건)
            </span>
          </div>
        ) : (
          <Link
            href={`/sales/new?from_consultation=${c.id}`}
            className="flex items-center gap-2 p-3 rounded-lg bg-blue-50 border border-blue-200 hover:bg-blue-100 transition"
          >
            <ShoppingCart size={16} className="text-blue-600" />
            <span className="text-sm font-medium text-blue-700">판매로 처리</span>
          </Link>
        )
      )}

      {/* 계약서 CTA — 톡상담 제외, 확정~완료 단계에서 표시 */}
      {['confirmed', 'in_progress', 'completed'].includes(c.status) && c.consultation_type !== 'talk_consult' && (
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
          onConfirm={async () => {
            // 077: 'completed'는 완료 후 토스트 + 탭 자동전환
            if (pendingStatus === 'completed') {
              await completeOnly();
            } else {
              await updateStatus.mutateAsync({ id: c.id, status: pendingStatus });
            }
          }}
          title={confirmStatusLabels[pendingStatus].title}
          message={confirmStatusLabels[pendingStatus].msg}
          confirmLabel={confirmStatusLabels[pendingStatus].title}
          variant={confirmStatusLabels[pendingStatus].variant || 'default'}
        />
      )}

      {/* 077: 상담완료 시 판매전환 분기 모달 (linkedSales 없는 경우만 표시) */}
      <Modal
        open={saleConvertOpen}
        onClose={() => setSaleConvertOpen(false)}
        title="상담 완료 처리"
        className="max-w-sm"
      >
        <div className="space-y-4">
          <div className="text-sm text-neutral-700">
            <p className="mb-2">{honorific(c.activity_name, c.name)}님의 상담을 완료 처리합니다.</p>
            <p className="text-amber-700 bg-amber-50 border border-amber-200 rounded-lg p-2.5">
              아직 연결된 판매가 없습니다.<br />판매를 먼저 등록하시겠습니까?
            </p>
          </div>
          <div className="flex flex-col gap-2">
            <Button
              variant="primary"
              size="sm"
              onClick={() => {
                setSaleConvertOpen(false);
                router.push(`/sales/new?from_consultation=${c.id}`);
              }}
            >
              <ShoppingCart size={14} />
              판매 등록하기
            </Button>
            <Button
              variant="secondary"
              size="sm"
              onClick={async () => {
                setSaleConvertOpen(false);
                await completeOnly();
              }}
              loading={updateStatus.isPending}
            >
              완료만 처리 (판매 없이)
            </Button>
            <Button
              variant="ghost"
              size="sm"
              onClick={() => setSaleConvertOpen(false)}
            >
              취소
            </Button>
          </div>
        </div>
      </Modal>

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

      {/* 삭제 확인 모달 */}
      <ConfirmModal
        open={showDeleteConfirm}
        onClose={() => setShowDeleteConfirm(false)}
        onConfirm={() => deleteConsultation.mutateAsync(c.id)}
        title="상담 건 삭제"
        message={<>이 건을 <strong>완전히 삭제</strong>합니다.<br />알림톡은 발송되지 않습니다. 복구할 수 없습니다.</>}
        confirmLabel="삭제"
        variant="danger"
      />

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

      {/* 079: 수동 일정 확정 모달 (매장/출장 한정) */}
      {c.consultation_type !== 'talk_consult' && (
        <ManualConfirmModal
          open={manualConfirmOpen}
          onClose={() => setManualConfirmOpen(false)}
          consultationId={c.id}
          customerName={c.name}
          consultationType={c.consultation_type as 'store_visit' | 'field_request'}
          currentStatus={c.status}
          currentVisitDate={c.visit_date}
          currentVisitTime={c.visit_time}
        />
      )}
    </div>
  );
}
