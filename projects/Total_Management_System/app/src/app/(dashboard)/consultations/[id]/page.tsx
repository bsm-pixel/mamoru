'use client';

import { useState } from 'react';
import { useParams, useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import {
  useConsultation,
  useUpdateConsultationStatus,
  useAssignDealer,
  useSendNotification,
} from '@/hooks/use-consultations';
import { useDealers } from '@/hooks/use-dealers';
import { RescheduleModal } from '@/components/consultations/reschedule-modal';
import { SuggestTimeModal } from '@/components/consultations/suggest-time-modal';
import { HoldReasonModal } from '@/components/consultations/hold-reason-modal';
import { getAllowedTransitions } from '@/lib/consultation/transitions';
import {
  formatDate,
  formatDateTime,
  formatPhone,
  formatRelative,
  formatDday,
  CONSULTATION_STATUS_LABEL,
  CONSULTATION_STATUS_COLOR,
  CONSULTATION_TYPE_LABEL,
} from '@/lib/utils/format';
import { ArrowLeft, MapPin, Calendar, User, Clock, UserCheck, AlertCircle, CheckCircle, FileSignature } from 'lucide-react';
import type { ConsultationStatus, ConsultationType } from '@/lib/supabase/types';

/** 상태 전이 → 버튼 라벨 매핑 */
const STATUS_BUTTON_LABEL: Record<string, string> = {
  pending_admin: '대기 복원',
  suggested: '시간 제안',
  confirmed: '확정',
  cancelled: '취소',
  reschedule_requested: '일정 변경 요청',
  on_hold: '보류',
  in_progress: '상담 시작',
  completed: '처리완료',
};

export default function ConsultationDetailPage() {
  const params = useParams();
  const router = useRouter();
  const id = params.id as string;
  const { data, isLoading } = useConsultation(id);
  const updateStatus = useUpdateConsultationStatus();
  const assignDealer = useAssignDealer();
  const sendNotify = useSendNotification();
  const { data: dealers } = useDealers();
  const [selectedDealer, setSelectedDealer] = useState('');
  const [showReschedule, setShowReschedule] = useState(false);
  const [showSuggest, setShowSuggest] = useState(false);
  const [showHold, setShowHold] = useState(false);
  const [showCancelConfirm, setShowCancelConfirm] = useState(false);

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
  const cType = c.consultation_type as ConsultationType;
  const allowed = getAllowedTransitions(cType, c.status as ConsultationStatus);

  // 방문일 지남 여부
  const visitDday = formatDday(c.visit_date);
  const isOverdue = c.status === 'confirmed' && visitDday.isPast;

  /** 유형별 액션 버튼 렌더링 */
  const renderTypeActions = () => {
    // 종료 상태: 버튼 없음
    if (c.status === 'completed' || c.status === 'cancelled') return null;

    const busy = updateStatus.isPending;
    const busyStatus = busy ? updateStatus.variables?.status : null;

    switch (cType) {
      case 'store_visit':
        return (
          <>
            {isOverdue && (
              <Button variant="primary" size="sm" disabled={busy} loading={busyStatus === 'completed'}
                onClick={() => updateStatus.mutate({ id: c.id, status: 'completed', note: '방문 완료 처리' })}
              >
                <CheckCircle size={14} />
                방문 완료
              </Button>
            )}
            {c.status === 'confirmed' && (
              <>
                <Button variant="secondary" size="sm" disabled={busy} onClick={() => setShowReschedule(true)}>
                  일정변경
                </Button>
                <Button
                  variant="ghost" size="sm"
                  disabled={busy || sendNotify.isPending}
                  loading={sendNotify.isPending}
                  onClick={() => sendNotify.mutate({ consultationId: c.id, template: 'confirmed' })}
                >
                  알림톡 재발송
                </Button>
              </>
            )}
            {allowed.includes('on_hold') && (
              <Button variant="ghost" size="sm" disabled={busy} onClick={() => setShowHold(true)}>보류</Button>
            )}
            {allowed.includes('confirmed') && c.status !== 'confirmed' && (
              <Button variant="secondary" size="sm" disabled={busy} loading={busyStatus === 'confirmed'} onClick={() => updateStatus.mutate({ id: c.id, status: 'confirmed' })}>
                확정 복원
              </Button>
            )}
            {allowed.includes('cancelled') && (
              <Button variant="danger" size="sm" disabled={busy} loading={busyStatus === 'cancelled'} onClick={() => setShowCancelConfirm(true)}>
                취소
              </Button>
            )}
          </>
        );
      case 'field_request':
        return (
          <>
            {isOverdue && (
              <Button variant="primary" size="sm" disabled={busy} loading={busyStatus === 'completed'}
                onClick={() => updateStatus.mutate({ id: c.id, status: 'completed', note: '출장 완료 처리' })}
              >
                <CheckCircle size={14} />
                출장 완료
              </Button>
            )}
            {(c.status === 'pending_admin' || c.status === 'reschedule_requested' || c.status === 'change_requested') && (
              <Button variant="primary" size="sm" disabled={busy} onClick={() => setShowSuggest(true)}>
                {(c.status === 'reschedule_requested' || c.status === 'change_requested') ? '새 시간 제안' : '시간 제안'}
              </Button>
            )}
            {/* 출장 확정건 수동 일정 변경 — 유선 연락 등으로 관리자가 직접 변경 */}
            {c.status === 'confirmed' && (
              <Button variant="secondary" size="sm" disabled={busy} onClick={() => setShowReschedule(true)}>
                일정변경
              </Button>
            )}
            {allowed.includes('confirmed') && (
              <Button variant="secondary" size="sm" disabled={busy} loading={busyStatus === 'confirmed'} onClick={() => updateStatus.mutate({ id: c.id, status: 'confirmed' })}>
                {c.status === 'suggested' ? '수동 확정' : '확정'}
              </Button>
            )}
            {allowed.includes('on_hold') && (
              <Button variant="ghost" size="sm" disabled={busy} onClick={() => setShowHold(true)}>보류</Button>
            )}
            {allowed.includes('cancelled') && (
              <Button variant="danger" size="sm" disabled={busy} loading={busyStatus === 'cancelled'} onClick={() => setShowCancelConfirm(true)}>
                취소
              </Button>
            )}
          </>
        );
      case 'talk_consult':
        return (
          <>
            {allowed.includes('in_progress') && (
              <Button variant="primary" size="sm" disabled={busy} loading={busyStatus === 'in_progress'} onClick={() => updateStatus.mutate({ id: c.id, status: 'in_progress' })}>
                상담 시작
              </Button>
            )}
            {allowed.includes('completed') && (
              <Button variant="primary" size="sm" disabled={busy} loading={busyStatus === 'completed'} onClick={() => updateStatus.mutate({ id: c.id, status: 'completed' })}>
                처리완료
              </Button>
            )}
            {allowed.includes('on_hold') && (
              <Button variant="ghost" size="sm" disabled={busy} onClick={() => setShowHold(true)}>보류</Button>
            )}
            {allowed.includes('pending_admin') && (
              <Button variant="secondary" size="sm" disabled={busy} loading={busyStatus === 'pending_admin'} onClick={() => updateStatus.mutate({ id: c.id, status: 'pending_admin' })}>
                대기 복원
              </Button>
            )}
            {allowed.includes('cancelled') && (
              <Button variant="danger" size="sm" disabled={busy} loading={busyStatus === 'cancelled'} onClick={() => setShowCancelConfirm(true)}>
                취소
              </Button>
            )}
          </>
        );
      default:
        return null;
    }
  };

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
              {c.unique_id} &middot; {formatDateTime(c.received_at)} &middot; {CONSULTATION_TYPE_LABEL[c.consultation_type]}
            </p>
          </div>
          <Badge className={statusColor}>
            {CONSULTATION_STATUS_LABEL[c.status] || c.status}
          </Badge>
        </div>

        {/* 보류 사유 카드 */}
        {c.status === 'on_hold' && c.hold_reason && (
          <Card className="border-warning/30 bg-warning-soft/20">
            <div className="flex items-start gap-2">
              <AlertCircle size={16} className="text-warning mt-0.5 shrink-0" />
              <div>
                <p className="text-sm font-semibold text-warning">보류 중</p>
                <p className="text-sm text-neutral-700 mt-0.5">{c.hold_reason}</p>
              </div>
            </div>
          </Card>
        )}

        {/* 고객 일정변경 요청 카드 */}
        {c.status === 'change_requested' && c.memo && (() => {
          // 비고에서 가장 최근 [고객 변경요청 ...] 파싱
          const lines = c.memo.split('\n');
          const changeLine = [...lines].reverse().find(l => l.includes('[고객 변경요청'));
          if (!changeLine) return null;
          return (
            <Card className="border-orange-300/50 bg-orange-50">
              <div className="flex items-start gap-2">
                <AlertCircle size={16} className="text-orange-600 mt-0.5 shrink-0" />
                <div>
                  <p className="text-sm font-semibold text-orange-700">고객 일정 변경 요청</p>
                  <p className="text-sm text-neutral-700 mt-0.5 whitespace-pre-wrap">{changeLine.replace(/^\[고객 변경요청 [\d-: ]+\]\s*/, '')}</p>
                </div>
              </div>
            </Card>
          );
        })()}

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
            {c.suggestions && (c.suggestions as { dates?: unknown[] }).dates && (
              <>
                <dt className="text-neutral-500">제안 일정</dt>
                <dd className="col-span-1">
                  {((c.suggestions as { dates: { date: string; time: string }[] }).dates).map((d, i) => (
                    <span key={i} className="block text-xs text-info">
                      {typeof d === 'string' ? d : `${d.date} ${d.time}`}
                    </span>
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
            {c.latitude && c.longitude && (
              <>
                <dt className="text-neutral-500">좌표</dt>
                <dd className="text-xs text-neutral-400">{c.latitude}, {c.longitude}</dd>
              </>
            )}
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

        {/* 유형별 액션 버튼 */}
        {allowed.length > 0 && (
          <Card>
            <CardHeader>
              <CardTitle>액션</CardTitle>
            </CardHeader>
            <div className="flex flex-wrap gap-2">
              {renderTypeActions()}
            </div>
          </Card>
        )}

        {/* 다음 단계: 계약서 작성 CTA */}
        {c.status === 'completed' && (
          <Card className="border-blue-200 bg-blue-50/50">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <FileSignature size={18} className="text-blue-600" />
                <div>
                  <p className="text-sm font-semibold text-blue-700">계약서 작성</p>
                  <p className="text-xs text-neutral-500">상담 완료 → 계약서를 작성합니다</p>
                </div>
              </div>
              <Button
                size="sm"
                onClick={() => router.push(`/contracts/new?customer_name=${encodeURIComponent(c.name)}&customer_phone=${encodeURIComponent(c.phone || '')}`)}
              >
                작성하기
              </Button>
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

      {/* 모달 */}
      {showReschedule && (
        <RescheduleModal
          open={showReschedule}
          onClose={() => setShowReschedule(false)}
          consultationId={c.id}
          currentDate={c.visit_date || ''}
          currentTime={c.visit_time || ''}
          consultationType={c.consultation_type}
          uniqueId={c.unique_id}
        />
      )}
      {showSuggest && (
        <SuggestTimeModal
          open={showSuggest}
          onClose={() => setShowSuggest(false)}
          consultationId={c.id}
        />
      )}
      {showHold && (
        <HoldReasonModal
          open={showHold}
          onClose={() => setShowHold(false)}
          consultationId={c.id}
        />
      )}

      {/* 취소 확인 모달 */}
      {showCancelConfirm && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
          <div className="bg-white rounded-xl shadow-xl p-6 w-full max-w-sm mx-4">
            <h3 className="text-lg font-bold text-neutral-900">상담 취소</h3>
            <p className="text-sm text-neutral-600 mt-2">
              정말 이 상담을 취소하시겠습니까?<br />
              <span className="text-danger text-xs">취소 시 캘린더 일정 삭제 및 알림톡이 발송됩니다.</span>
            </p>
            <div className="flex gap-2 mt-5 justify-end">
              <Button variant="ghost" size="sm" onClick={() => setShowCancelConfirm(false)}>
                돌아가기
              </Button>
              <Button
                variant="danger"
                size="sm"
                loading={updateStatus.isPending}
                onClick={() => {
                  updateStatus.mutate(
                    { id: c.id, status: 'cancelled' },
                    { onSettled: () => setShowCancelConfirm(false) }
                  );
                }}
              >
                취소 확정
              </Button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
