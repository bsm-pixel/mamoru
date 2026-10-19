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
import { FieldDelayAction } from './field-delay-action';
import { formatPhone, formatDate, CONSULTATION_STATUS_LABEL } from '@/lib/utils/format';
import { Calendar, MapPin, Phone, User, Clock, FileSignature, ShoppingCart, CheckCircle2, CheckCircle } from 'lucide-react';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import { Modal } from '@/components/ui/modal';
import Link from 'next/link';
import { ActionNote, MoreActions, DangerZone, DangerLink, SubtleButton } from '@/components/ui/action-section';
import { HoldReasonModal } from './hold-reason-modal';
import { getAllowedTransitions } from '@/lib/consultation/transitions';
import type { ConsultationStatus, ConsultationType } from '@/lib/supabase/types';
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
  // 2026-10-19 보류 — 상세 페이지에만 있던 기능을 패널에도 (사유 입력 모달 공용)
  const [holdOpen, setHoldOpen] = useState(false);
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

  // ── 액션 5칸 규칙 (2026-10-19) — 상태 전이 SSOT = lib/consultation/transitions.ts ──
  const type = c.consultation_type as ConsultationType;
  const allowed = getAllowedTransitions(type, c.status as ConsultationStatus);
  const isTerminal = ['completed', 'cancelled'].includes(c.status);
  const canHold = allowed.includes('on_hold');
  const canCancel = allowed.includes('cancelled');
  const canManualConfirm = !isTerminal && type !== 'talk_consult' && !['in_progress', 'on_hold'].includes(c.status);
  const canSellContract = ['confirmed', 'in_progress', 'completed'].includes(c.status) && type !== 'talk_consult';
  const busy = updateStatus.isPending;
  const visitMeta = c.visit_date ? `${formatDate(c.visit_date)} ${c.visit_time || ''}`.trim() : undefined;

  // ① 지금 상태 한 줄 — 무엇을 기다리고, 다음에 뭘 누르면 되는지
  const renderNote = () => {
    switch (c.status) {
      case 'pending_admin':
        return type === 'talk_consult'
          ? <ActionNote tone="wait" title="신규접수">톡으로 응대를 시작하면 [상담 시작]을 눌러주세요.</ActionNote>
          : <ActionNote tone="wait" title="신규접수">가능 요일·시간을 보고 방문 시간을 제안해주세요.</ActionNote>;
      case 'suggested':
        return <ActionNote tone="wait" title="제안중">고객이 일정을 고르면 자동으로 확정됩니다. 전화로 정했다면 [수동 일정 확정].</ActionNote>;
      case 'reschedule_requested':
      case 'change_requested':
        return <ActionNote tone="wait" title="일정 재요청">고객이 다른 일정을 원합니다. 재요청 메모를 보고 다시 제안해주세요.</ActionNote>;
      case 'confirmed':
        return <ActionNote tone="wait" title="확정" meta={visitMeta}>상담이 끝나면 [상담 완료]를 눌러주세요.</ActionNote>;
      case 'in_progress':
        return <ActionNote tone="wait" title="진행중">응대가 끝나면 [상담 완료]를 눌러주세요.</ActionNote>;
      case 'on_hold':
        return <ActionNote tone="muted" title="보류" meta={c.hold_reason || undefined}>다시 진행할 때 아래 버튼을 눌러주세요.</ActionNote>;
      case 'completed':
        return <ActionNote tone="done" title="완료">{hasLinkedSales || type === 'talk_consult' ? '끝난 상담입니다.' : '판매로 이어졌다면 아래 [판매로 처리]로 연결해주세요.'}</ActionNote>;
      case 'cancelled':
        return <ActionNote tone="muted" title="취소됨" />;
      default:
        return <ActionNote tone="muted" title={statusLabel} />;
    }
  };

  // ② 주 액션 — 상태별 딱 하나
  const resumeTo: ConsultationStatus | null = c.status !== 'on_hold' ? null
    : type === 'talk_consult' ? 'in_progress'
    : type === 'store_visit' ? 'confirmed'
    : 'pending_admin';
  const resumeLabel = type === 'talk_consult' ? '상담 재개' : type === 'store_visit' ? '다시 확정 (보류 해제)' : '신규로 되돌리기 (보류 해제)';
  const primary = (() => {
    if (type === 'field_request' && ['pending_admin', 'reschedule_requested', 'change_requested'].includes(c.status)) {
      return <Button className="w-full" onClick={() => setSuggestModalOpen(true)} disabled={busy}><Clock size={14} />시간 제안</Button>;
    }
    if (type === 'talk_consult' && c.status === 'pending_admin') {
      return <Button className="w-full" onClick={() => handleStatus('in_progress')} disabled={busy}>상담 시작</Button>;
    }
    if (c.status === 'confirmed' || c.status === 'in_progress') {
      return <Button className="w-full" onClick={() => handleStatus('completed')} disabled={busy}><CheckCircle2 size={14} />상담 완료</Button>;
    }
    if (resumeTo && allowed.includes(resumeTo)) {
      return (
        <Button className="w-full" loading={busy} onClick={() => updateStatus.mutate({ id: c.id, status: resumeTo, note: '보류 해제' })}>
          {resumeLabel}
        </Button>
      );
    }
    return null;
  })();

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
        const reRequestRaw = lines.filter(l => l.includes('[고객 재요청'));
        const originalLines = lines.filter(l => !l.includes('[고객 재요청'));
        const originalMemo = originalLines.join('\n').trim();

        // [고객 재요청 <시각>] <사유> → {time, msg}로 분리(시각은 초 제거·간결화: "2026. 7. 11. 오전 8:11:30" → "7/11 오전 8:11")
        const shortTs = (ts: string) => {
          const m = ts.match(/(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\.\s*(오전|오후)?\s*(\d{1,2}):(\d{2})/);
          return m ? `${m[2]}/${m[3]}${m[4] ? ' ' + m[4] : ''} ${m[5]}:${m[6]}` : ts.replace(/(\d{1,2}:\d{2}):\d{2}/, '$1').trim();
        };
        const reRequests = reRequestRaw
          .map(l => {
            const clean = l.replace(/^\d+\s*/, '');
            const m = clean.match(/^\[고객 재요청\s*([^\]]+)\]\s*([\s\S]*)$/);
            return m ? { time: shortTs(m[1].trim()), msg: m[2].trim() } : { time: '', msg: clean.trim() };
          })
          .filter(r => r.msg || r.time);

        // 고객이 반응한 "내가 제안/확정했던 일정" — visit_date(확정) 우선, 없으면 suggestions(제안중). 재요청 시점에도 지워지지 않고 남아있음
        let proposedLabel = '';
        let proposedText = '';
        if (c.visit_date) {
          proposedLabel = '확정했던 일정';
          proposedText = `${formatDate(c.visit_date)} ${c.visit_time || ''}`.trim();
        } else {
          const rawSug = c.suggestions as { dates?: Array<{ date?: string; time?: string }> } | Array<{ date?: string; time?: string }> | null;
          const dates = Array.isArray(rawSug) ? rawSug : (rawSug?.dates ?? []);
          const valid = dates.filter(d => d && d.date);
          if (valid.length > 0) {
            proposedLabel = '내가 제안했던 일정';
            proposedText = valid.map(d => `${formatDate(d.date as string)}${d.time ? ' ' + d.time : ''}`).join(' / ');
          }
        }

        return (
          <div className="space-y-2">
            {originalMemo && (
              <div className="bg-white border border-neutral-200 rounded-lg p-3">
                <p className="text-xs text-neutral-500 mb-1">접수 메모</p>
                <p className="text-sm whitespace-pre-wrap">{originalMemo}</p>
              </div>
            )}
            {reRequests.length > 0 && (
              <div className="bg-amber-50 border border-amber-200 rounded-lg p-3 space-y-2">
                <p className="text-xs text-amber-600 font-medium">재요청 메모</p>
                {/* 고객이 반응한 나의 제안/확정 일정 — "내가 몇일 몇시로 제안했더라?" 즉시 파악 */}
                {proposedText && (
                  <div className="flex items-baseline gap-1.5 text-xs bg-amber-100/60 rounded-md px-2 py-1.5">
                    <span className="text-amber-700 shrink-0">{proposedLabel}</span>
                    <span className="text-amber-900 font-semibold">{proposedText}</span>
                  </div>
                )}
                <div className="space-y-1.5">
                  {reRequests.map((r, i) => (
                    <div key={i}>
                      <p className="text-sm text-amber-800">{r.msg}</p>
                      {r.time && <p className="text-[11px] text-amber-500 mt-0.5">재요청 접수 · {r.time}</p>}
                    </div>
                  ))}
                </div>
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

      {/* 이 상담으로 처리된 판매 (070 역방향 link) — 정보 */}
      {hasLinkedSales && (
        <div className="flex flex-wrap items-center gap-1.5 text-xs">
          <CheckCircle2 size={14} className="text-green-600" />
          <span className="text-neutral-500 font-medium">판매연결완료:</span>
          {data!.linkedSales!.map((s) => (
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

      {/* 액션 — 표준 5칸 (2026-10-19, 복원수리·판매와 같은 순서: ① 상태 한 줄 ② 검은 버튼 1개 ③ 보조 ④ 접힘 ⑤ 위험)
          전엔 상태에 따라 버튼 4~6개가 평면 나열되고 보류가 없었으며, 판매·계약 CTA 가 색 박스로 따로 떠 있었다 */}
      <div className="pt-3 border-t border-neutral-100 space-y-2">
        <h4 className="text-xs font-semibold text-neutral-500">액션</h4>
        {/* ① 지금 상태 */}
        {renderNote()}
        {/* ② 주 액션 — 딱 하나 */}
        {primary}
        {/* ③ 보조 */}
        {canManualConfirm && (
          <Button size="sm" variant="secondary" className="w-full" onClick={() => setManualConfirmOpen(true)} disabled={updateStatus.isPending}>
            <CheckCircle size={14} />
            {c.status === 'confirmed' ? '수동 일정 변경' : '수동 일정 확정'}
          </Button>
        )}
        {/* 출장 도착 지연 안내 (2026-09-22 복원) — 확정 출장만, 알림톡 field_delayed */}
        {c.status === 'confirmed' && type === 'field_request' && (
          <FieldDelayAction consultationId={c.id} visitTime={c.visit_time} className="w-full" />
        )}
        {/* 판매로 처리 · 계약서 — 톡상담 제외, 확정~완료 (판매연결 있으면 판매 버튼 숨김 = 중복 처리 방지 077) */}
        {canSellContract && (
          <div className="grid grid-cols-2 gap-2">
            {!hasLinkedSales ? (
              <Link href={`/sales/new?from_consultation=${c.id}`}
                className="h-9 rounded-lg border border-neutral-200 bg-white text-xs font-semibold text-neutral-800 hover:bg-neutral-50 flex items-center justify-center gap-1.5">
                <ShoppingCart size={14} />판매로 처리
              </Link>
            ) : <span />}
            <Link href={`/contracts/new?customer_name=${encodeURIComponent(c.name)}&customer_phone=${encodeURIComponent(c.phone || '')}`}
              className={`h-9 rounded-lg border border-neutral-200 bg-white text-xs font-semibold text-neutral-800 hover:bg-neutral-50 flex items-center justify-center gap-1.5 ${hasLinkedSales ? 'col-span-2' : ''}`}>
              <FileSignature size={14} />계약서 작성
            </Link>
          </div>
        )}
        {/* ④ 드물게 쓰는 것 — 보류 */}
        {canHold && (
          <MoreActions label="다른 처리 방법">
            <SubtleButton onClick={() => setHoldOpen(true)} disabled={updateStatus.isPending}>
              보류 (사유 입력)
            </SubtleButton>
          </MoreActions>
        )}
        {/* ⑤ 파괴적 액션 */}
        <DangerZone>
          {canCancel && (
            <DangerLink onClick={() => handleStatus('cancelled')} disabled={updateStatus.isPending}>
              상담 취소
            </DangerLink>
          )}
          <DangerLink onClick={() => setShowDeleteConfirm(true)}>
            삭제 (알림 없이 제거)
          </DangerLink>
        </DangerZone>
      </div>

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

      {/* 보류 사유 모달 (상세 페이지와 공용) */}
      {holdOpen && <HoldReasonModal open={holdOpen} onClose={() => setHoldOpen(false)} consultationId={c.id} />}

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
