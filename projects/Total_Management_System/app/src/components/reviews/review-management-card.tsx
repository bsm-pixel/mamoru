'use client';

/**
 * 067: 리뷰 관리 공용 카드 — 상담/수리/판매 상세 패널에서 동일 사용
 *
 * 역할:
 *  1) 약속 토글: 사장님이 "이 고객 후기 약속 받았다" 체크 (review_promised_at)
 *  2) 후기 요청 발송: ReviewRequestModal 열어 알림톡 수동 발송 (review_request_sent_at)
 *  3) 작성 완료 표시: review_submitted_at 있으면 readonly 정적 라벨로 전환
 *
 * 자동 발송 정책 (system_settings.review.auto_request_on_completion):
 *  - OFF (기본, 핀셋 정책): 약속 ✓ 고객만 사장님 수동 발송
 *  - ON (안내문 정책): 약속 X 고객은 자동 발송 / 약속 ✓ 고객은 항상 사장님 수동만
 */

import { useState, useEffect } from 'react';
import Link from 'next/link';
import { Card } from '@/components/ui/card';
import { Star, MessageCircle, CheckCircle2, Send, Info, Clock } from 'lucide-react';
import toast from 'react-hot-toast';
import { ReviewRequestModal } from '@/components/sales/review-request-modal';
import { useSetting } from '@/hooks/use-settings';
import type { ReviewSource } from '@/lib/notification/review-request';

interface RelatedActivity {
  source: ReviewSource;
  id: string;
  displayId: string;
  typeLabel: string;
  promisedAt: string | null;
  requestSentAt: string | null;
  submittedAt: string | null;
}

type PromiseType = 'purchase' | 'repair' | 'consult';
// 119: 직접발송(택배 수리) subtype 추가 — proceed_type 3종과 1:1 (직접발송이 direct_visit 로 뭉개지던 버그 fix)
type RepairSubtype = 'direct_visit' | 'pickup' | 'delivery';
type ConsultSubtype = 'store_visit' | 'field_request' | 'talk_consult';
type PromiseSubtype = RepairSubtype | ConsultSubtype;

interface Props {
  source: ReviewSource;
  id: string;
  customerName: string;
  customerPhone: string | null;
  promisedAt: string | null;
  /** 094: 약속한 후기 유형 (자동 발송 시 솔라피 템플릿 결정) — NULL이면 source 디폴트 매핑 */
  promisedType?: PromiseType | null;
  /** 095: 약속 세부 유형 (repair: direct_visit|pickup / consult: store_visit|field_request|talk_consult) */
  promisedSubtype?: string | null;
  requestSentAt: string | null;
  submittedAt: string | null;
  /** 변경 후 부모 쿼리 invalidate / refetch 트리거 */
  onChanged?: () => void;
  /** sale source일 때 수리 상품 포함 여부 (ReviewRequestModal 기본값 결정) */
  hasRepairItem?: boolean;
  /** source의 실제 sub-type (consultation_type / proceed_type 등) — ReviewRequestModal subtype 자동 추론 */
  sourceType?: string | null;
  /** 2026-05-26: 컴팩트 모드 — 상세 패널 헤더 우측에 미니 UI 로 표시 (시각 부담 ↓) */
  compact?: boolean;
  /** 2026-05-26: 자동 발송 예정 판정용 (compact 모드에서 "⏳ 자동 발송 예정" 표시) */
  shippedAt?: string | null;
  deliveredAt?: string | null;
  invoiceNumber?: string | null;
}

/** source → 디폴트 유형 매핑 (094) */
function defaultPromiseType(source: ReviewSource, hasRepairItem: boolean): PromiseType {
  if (source === 'repair') return 'repair';
  if (source === 'consultation') return 'consult';
  // source === 'sale'
  return hasRepairItem ? 'repair' : 'purchase';
}

/** type + sourceType → 디폴트 subtype (095) */
function defaultPromiseSubtype(type: PromiseType, sourceType: string | null): PromiseSubtype | null {
  if (type === 'purchase') return null;
  if (type === 'repair') {
    // proceed_type 3종을 1:1 매핑 (119): 직접발송 → delivery(택배 수리)
    if (sourceType === '직접방문' || sourceType === 'direct_visit') return 'direct_visit';
    if (sourceType === '방문수거' || sourceType === 'pickup') return 'pickup';
    if (sourceType === '직접발송' || sourceType === 'delivery') return 'delivery';
    return 'delivery'; // 미지정은 가장 흔한 택배 수리로 (기존엔 direct_visit 로 잘못 뭉갬)
  }
  // type === 'consult'
  if (sourceType === 'store_visit' || sourceType === 'field_request' || sourceType === 'talk_consult') {
    return sourceType;
  }
  return 'store_visit';
}

const TYPE_LABEL: Record<PromiseType, string> = {
  repair: '복원수리',
  consult: '상담',
  purchase: '제품구매',
};

const SUBTYPE_LABEL: Record<PromiseSubtype, string> = {
  direct_visit: '직접방문',
  pickup: '방문수거',
  delivery: '택배 수리',
  store_visit: '직접방문',
  field_request: '출장',
  talk_consult: '톡상담',
};

const REPAIR_SUBTYPES: RepairSubtype[] = ['direct_visit', 'pickup', 'delivery'];
const CONSULT_SUBTYPES: ConsultSubtype[] = ['store_visit', 'field_request', 'talk_consult'];

function formatDate(iso: string | null): string {
  if (!iso) return '';
  try {
    const d = new Date(iso);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  } catch {
    return '';
  }
}

/** 같은 고객의 다른 source 활동 표시 영역 (자동 매칭 X — 정보 표시용) */
function RelatedActivitySection({ items }: { items: RelatedActivity[] }) {
  if (items.length === 0) return null;
  return (
    <div className="mt-3 pt-3 border-t border-neutral-100">
      <div className="flex items-center gap-1.5 mb-2">
        <Info size={11} className="text-neutral-500" />
        <span className="text-[10px] font-semibold text-neutral-500 uppercase tracking-wider">같은 고객 다른 활동</span>
      </div>
      <div className="space-y-1.5">
        {items.map((it) => {
          const detailHref = it.source === 'consultation' ? `/consultations/${it.id}` : it.source === 'repair' ? `/repairs/${it.id}` : `/sales/${it.id}`;
          const isCompleted = !!it.submittedAt;
          const isPending = !isCompleted && !!it.requestSentAt;
          const isPromised = !isCompleted && !isPending && !!it.promisedAt;
          let chipLabel = '';
          let chipClass = '';
          if (isCompleted) { chipLabel = `✅ 작성완료 ${formatDate(it.submittedAt)}`; chipClass = 'bg-green-50 text-green-700'; }
          else if (isPending) { chipLabel = `📤 발송 ${formatDate(it.requestSentAt)} · 대기`; chipClass = 'bg-blue-50 text-blue-700'; }
          else if (isPromised) { chipLabel = `☑ 약속 ${formatDate(it.promisedAt)}`; chipClass = 'bg-amber-50 text-amber-700'; }
          return (
            <Link
              key={`${it.source}-${it.id}`}
              href={detailHref}
              className="flex items-center gap-2 px-2.5 py-1.5 rounded-md hover:bg-neutral-50 transition group"
            >
              <span className="text-[10px] px-1.5 py-0.5 rounded bg-neutral-100 text-neutral-600 font-semibold shrink-0">{it.typeLabel}</span>
              <span className="text-[11px] text-neutral-500 font-mono truncate">{it.displayId}</span>
              <span className={`text-[10px] px-1.5 py-0.5 rounded font-semibold ${chipClass} shrink-0 ml-auto`}>{chipLabel}</span>
              <span className="text-[10px] text-neutral-400 group-hover:text-terracotta">→</span>
            </Link>
          );
        })}
      </div>
    </div>
  );
}

export function ReviewManagementCard({
  source,
  id,
  customerName,
  customerPhone,
  promisedAt,
  promisedType = null,
  promisedSubtype = null,
  requestSentAt,
  submittedAt,
  onChanged,
  hasRepairItem = false,
  sourceType = null,
  compact = false,
  shippedAt = null,
  deliveredAt = null,
  invoiceNumber = null,
}: Props) {
  const [togglingPromise, setTogglingPromise] = useState(false);
  const [showRequestModal, setShowRequestModal] = useState(false);
  const [related, setRelated] = useState<RelatedActivity[]>([]);

  // 2026-05-26: 자동 발송 예정 판정 (사장님 우려 → 시각 신호)
  //   조건 5중: 토글 ON + 약속 ✓ + 미발송 + 송장 있음 + 배송중(shipped + !delivered)
  //   ALPS cron 1시간마다 자동 추적 → '41'/'45' 코드 감지 시 자동 발송 예정
  const autoEnabled = useSetting<boolean>('review.auto_request_on_completion', false);
  const autoSendPending =
    autoEnabled &&
    !!promisedAt &&
    !requestSentAt &&
    !!invoiceNumber &&
    !!shippedAt &&
    !deliveredAt;

  // 같은 phone의 다른 source 활동 조회 (정보 표시용 — 자동 매칭 X)
  useEffect(() => {
    if (!customerPhone) return;
    const params = new URLSearchParams({
      phone: customerPhone,
      excludeSource: source,
      excludeId: id,
    });
    fetch(`/api/reviews/related-activity?${params.toString()}`)
      .then((r) => r.ok ? r.json() : { items: [] })
      .then((d) => setRelated(d.items || []))
      .catch(() => setRelated([]));
  }, [customerPhone, source, id]);

  // 늦게 도착한 리뷰 자동 매칭 — submittedAt 없을 때만 한 번 검사
  // 067 배포 이전에 작성된 리뷰 또는 source_id 매칭 실패로 누락된 review_submitted_at 백필
  useEffect(() => {
    if (submittedAt) return;
    fetch('/api/reviews/auto-match', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ source, id }),
    })
      .then((r) => r.ok ? r.json() : null)
      .then((d) => {
        if (d?.matched && !d.alreadySet) {
          // 백필 성공 — 카드 refresh
          onChanged?.();
        }
      })
      .catch(() => { /* 조용히 실패, 카드 핵심 기능에 영향 X */ });
    // submittedAt이 있으면 호출 자체 안 함 (멱등성 보강은 endpoint에서도 가드)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [source, id]);

  // 작성 완료 — 카드 readonly + 정적 라벨
  if (submittedAt) {
    if (compact) {
      // 컴팩트: 상세 패널 헤더 우측에 미니 표시
      return (
        <div className="flex items-center gap-1.5 text-[11px] text-green-700" title={`${formatDate(submittedAt)} 작성 완료`}>
          <CheckCircle2 size={13} />
          <span className="font-semibold">리뷰 작성 완료</span>
        </div>
      );
    }
    return (
      <Card>
        <div className="flex items-center gap-2 mb-2">
          <Star size={14} className="text-terracotta" />
          <h3 className="text-xs font-bold text-indigo-black">리뷰 관리</h3>
        </div>
        <div className="flex items-center gap-2 px-3 py-2.5 rounded-lg bg-green-50 border border-green-200">
          <CheckCircle2 size={16} className="text-green-700 shrink-0" />
          <div className="flex-1">
            <p className="text-sm font-semibold text-green-800">작성 완료</p>
            <p className="text-[11px] text-green-700">{formatDate(submittedAt)}에 리뷰가 등록되었습니다</p>
          </div>
        </div>
        <RelatedActivitySection items={related} />
      </Card>
    );
  }

  // 094: 활성 유형 — DB 값 우선, 없으면 source 디폴트
  const activeType: PromiseType = promisedType ?? defaultPromiseType(source, hasRepairItem);
  // 095: 활성 subtype — DB 값 우선, 없으면 type+sourceType 디폴트 (purchase 면 null)
  const activeSubtype: PromiseSubtype | null = (() => {
    if (activeType === 'purchase') return null;
    const fromDb = promisedSubtype as PromiseSubtype | null;
    if (fromDb && (REPAIR_SUBTYPES as string[]).concat(CONSULT_SUBTYPES as string[]).includes(fromDb)) return fromDb;
    return defaultPromiseSubtype(activeType, sourceType);
  })();

  /** /api/reviews/promise 호출 공통 헬퍼 — on=true 일 때 type + subtype 모두 전달 */
  const callPromiseApi = async (
    on: boolean,
    nextType?: PromiseType,
    nextSubtype?: PromiseSubtype | null,
    successMsg?: string,
  ) => {
    if (togglingPromise) return;
    setTogglingPromise(true);
    try {
      const body: Record<string, unknown> = { source, id, on };
      if (on) {
        body.type = nextType;
        body.subtype = nextType === 'purchase' ? null : (nextSubtype ?? null);
      }
      const res = await fetch('/api/reviews/promise', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        toast.error(err.error || '저장 실패');
        return;
      }
      if (successMsg) toast.success(successMsg);
      onChanged?.();
    } catch (e) {
      toast.error(`오류: ${String(e)}`);
    } finally {
      setTogglingPromise(false);
    }
  };

  const handleTogglePromise = async () => {
    const next = !promisedAt;
    if (next) {
      await callPromiseApi(true, activeType, activeSubtype, '리뷰 약속 체크');
    } else {
      await callPromiseApi(false, undefined, undefined, '약속 해제');
    }
  };

  // 094: 유형 변경 — subtype 은 새 유형의 디폴트로 리셋
  const handleChangePromiseType = async (next: PromiseType) => {
    if (next === activeType) return;
    const nextSubtype = next === 'purchase' ? null : defaultPromiseSubtype(next, sourceType);
    await callPromiseApi(true, next, nextSubtype, `약속 유형: ${TYPE_LABEL[next]}`);
  };

  // 095: subtype 변경 (현재 유형 유지)
  const handleChangePromiseSubtype = async (next: PromiseSubtype) => {
    if (next === activeSubtype) return;
    await callPromiseApi(true, activeType, next, `세부 유형: ${SUBTYPE_LABEL[next]}`);
  };

  /** 현재 유형에 맞는 subtype 옵션 (purchase 면 빈 배열) */
  const availableSubtypes: PromiseSubtype[] =
    activeType === 'repair' ? [...REPAIR_SUBTYPES]
    : activeType === 'consult' ? [...CONSULT_SUBTYPES]
    : [];

  /** 후기 요청 발송 — 약속 ON 시 모달 우회 (이미 유형/subtype 선택됨), OFF 시 모달 열기 */
  const handleRequestClick = async () => {
    if (!customerPhone) { toast.error('고객 연락처가 없어 발송할 수 없습니다'); return; }
    // 자동 발송 예정 confirm 가드 (compact 모드에서만 의미 — 기존 동작 유지)
    if (autoSendPending) {
      const ok = confirm('이 건은 배송완료 시 자동 발송될 예정입니다.\n\n지금 수동으로 발송하시겠습니까?');
      if (!ok) return;
    }
    // 약속 OFF → 모달 열기 (유형 선택 → 발송)
    if (!promisedAt) {
      setShowRequestModal(true);
      return;
    }
    // 약속 ON → 토글+칩 선택값 그대로 즉시 발송 (IA: SSOT 일관)
    setTogglingPromise(true);
    try {
      const res = await fetch('/api/reviews/request', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          source,
          id,
          review_type: activeType,
          subtype: activeType === 'purchase' ? undefined : (activeSubtype || undefined),
        }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        toast.error(err.error || '발송 실패');
        return;
      }
      toast.success('후기 요청 알림톡을 발송했습니다');
      onChanged?.();
    } catch (e) {
      toast.error(`오류: ${String(e)}`);
    } finally {
      setTogglingPromise(false);
    }
  };

  // 2026-05-26 Phase G-6 후속: 컴팩트 모드 시안 3 (토글 스위치 + 날짜)
  // 2026-05-27 (094): 토글 ON 시 [복원수리/상담/제품구매] 라디오 칩 인라인 표시
  if (compact) {
    return (
      <>
        <div className="flex items-center gap-2 flex-wrap min-w-0">
          {/* 약속 토글 스위치 (시안 3) */}
          <button
            type="button"
            onClick={handleTogglePromise}
            disabled={togglingPromise}
            title={promisedAt ? `리뷰 약속 ON (${formatDate(promisedAt)})` : '리뷰 약속 받음 토글'}
            className="flex items-center gap-2 text-[11px] px-2 py-1 rounded-md hover:bg-neutral-50 transition disabled:opacity-50"
          >
            <span
              className={`relative inline-block w-7 h-4 rounded-full transition ${
                promisedAt ? 'bg-neutral-900' : 'bg-neutral-300'
              }`}
            >
              <span
                className={`absolute top-0.5 w-3 h-3 bg-white rounded-full shadow-sm transition-all ${
                  promisedAt ? 'left-3.5' : 'left-0.5'
                }`}
              />
            </span>
            <span className={promisedAt ? 'font-semibold text-neutral-900' : 'text-neutral-500'}>리뷰 약속</span>
            {promisedAt && <span className="text-[10px] text-neutral-400">{formatDate(promisedAt)}</span>}
          </button>

          {/* 094: 약속 ON 시 유형 라디오 칩 (자동 발송 시 솔라피 템플릿 분기) */}
          {promisedAt && (
            <div className="flex items-center gap-1">
              {(['repair', 'consult', 'purchase'] as PromiseType[]).map((t) => (
                <button
                  key={t}
                  type="button"
                  onClick={() => handleChangePromiseType(t)}
                  disabled={togglingPromise}
                  title={`자동 발송 유형: ${TYPE_LABEL[t]}`}
                  className={`text-[10px] px-1.5 py-0.5 rounded transition disabled:opacity-50 ${
                    activeType === t
                      ? 'bg-stone-900 text-white font-semibold'
                      : 'bg-stone-100 text-stone-600 hover:bg-stone-200'
                  }`}
                >
                  {TYPE_LABEL[t]}
                </button>
              ))}
            </div>
          )}

          {/* 095: 약속 ON + repair/consult 일 때 subtype 라디오 칩 */}
          {promisedAt && availableSubtypes.length > 0 && (
            <div className="flex items-center gap-1 pl-1 border-l border-stone-200">
              {availableSubtypes.map((st) => (
                <button
                  key={st}
                  type="button"
                  onClick={() => handleChangePromiseSubtype(st)}
                  disabled={togglingPromise}
                  title={`${TYPE_LABEL[activeType]} · ${SUBTYPE_LABEL[st]}`}
                  className={`text-[10px] px-1.5 py-0.5 rounded transition disabled:opacity-50 ${
                    activeSubtype === st
                      ? 'bg-stone-700 text-white font-semibold'
                      : 'bg-stone-50 text-stone-500 hover:bg-stone-100'
                  }`}
                >
                  {SUBTYPE_LABEL[st]}
                </button>
              ))}
            </div>
          )}

          {/* 후기 요청 작은 버튼 — 약속 ON 시 모달 없이 즉시 발송, OFF 시 모달 */}
          <button
            type="button"
            onClick={handleRequestClick}
            disabled={togglingPromise}
            title={
              autoSendPending
                ? '배송완료 자동 감지 시 자동 발송 예정 (ALPS cron 1시간마다)'
                : requestSentAt ? `최근 발송: ${formatDate(requestSentAt)}` : '후기 요청 알림톡 발송'
            }
            className={`flex items-center gap-1 text-[11px] px-2 py-1 rounded-md transition ${
              autoSendPending
                ? 'bg-amber-50 text-amber-700 border border-amber-200 hover:bg-amber-100'
                : 'bg-indigo-black text-cream hover:bg-indigo-black/85'
            }`}
          >
            {autoSendPending ? <Clock size={10} /> : <Send size={10} />}
            <span>
              {autoSendPending
                ? '자동 발송 예정'
                : requestSentAt ? '재발송' : '후기 요청'}
            </span>
          </button>
        </div>

        {showRequestModal && (
          <ReviewRequestModal
            saleId={id}
            source={source}
            customerName={customerName}
            customerPhone={customerPhone || ''}
            hasRepairItem={hasRepairItem}
            sourceType={sourceType}
            alreadySent={!!requestSentAt}
            onClose={() => setShowRequestModal(false)}
            onSent={() => { setShowRequestModal(false); onChanged?.(); }}
          />
        )}
      </>
    );
  }

  return (
    <>
      <Card>
        <div className="flex items-center gap-2 mb-3">
          <Star size={14} className="text-terracotta" />
          <h3 className="text-xs font-bold text-indigo-black">리뷰 관리</h3>
        </div>

        {/* 약속 토글 */}
        <button
          type="button"
          onClick={handleTogglePromise}
          disabled={togglingPromise}
          className={`w-full flex items-center justify-between px-3 py-2.5 rounded-lg border transition mb-2 ${
            promisedAt
              ? 'bg-stone-900/5 border-stone-900/40'
              : 'bg-white border-neutral-200 hover:bg-neutral-50'
          }`}
        >
          <div className="flex items-center gap-2">
            <span
              className={`w-4 h-4 rounded border-2 flex items-center justify-center text-[10px] font-bold transition ${
                promisedAt
                  ? 'bg-stone-900 border-stone-900 text-white'
                  : 'bg-white border-neutral-300 text-transparent'
              }`}
            >
              {promisedAt ? '✓' : ''}
            </span>
            <span className={`text-sm ${promisedAt ? 'font-semibold text-indigo-black' : 'text-neutral-600'}`}>
              리뷰 참여 약속
            </span>
          </div>
          {promisedAt && <span className="text-[11px] text-neutral-500">{formatDate(promisedAt)}</span>}
        </button>

        {/* 094: 약속 ON 시 유형 라디오 칩 (자동 발송 시 솔라피 템플릿 분기) */}
        {promisedAt && (
          <div className="mb-2 flex items-center gap-1.5">
            <span className="text-[10px] text-stone-500 uppercase tracking-wider font-semibold">자동 발송 유형</span>
            <div className="flex items-center gap-1 ml-auto">
              {(['repair', 'consult', 'purchase'] as PromiseType[]).map((t) => (
                <button
                  key={t}
                  type="button"
                  onClick={() => handleChangePromiseType(t)}
                  disabled={togglingPromise}
                  className={`text-[11px] px-2 py-0.5 rounded transition disabled:opacity-50 ${
                    activeType === t
                      ? 'bg-stone-900 text-white font-semibold'
                      : 'bg-stone-100 text-stone-600 hover:bg-stone-200'
                  }`}
                >
                  {TYPE_LABEL[t]}
                </button>
              ))}
            </div>
          </div>
        )}

        {/* 095: 약속 ON + repair/consult 일 때 subtype 라디오 칩 */}
        {promisedAt && availableSubtypes.length > 0 && (
          <div className="mb-2 flex items-center gap-1.5">
            <span className="text-[10px] text-stone-500 uppercase tracking-wider font-semibold">세부 유형</span>
            <div className="flex items-center gap-1 ml-auto">
              {availableSubtypes.map((st) => (
                <button
                  key={st}
                  type="button"
                  onClick={() => handleChangePromiseSubtype(st)}
                  disabled={togglingPromise}
                  className={`text-[11px] px-2 py-0.5 rounded transition disabled:opacity-50 ${
                    activeSubtype === st
                      ? 'bg-stone-700 text-white font-semibold'
                      : 'bg-stone-50 text-stone-500 border border-stone-200 hover:bg-stone-100'
                  }`}
                >
                  {SUBTYPE_LABEL[st]}
                </button>
              ))}
            </div>
          </div>
        )}

        {/* 후기 요청 발송 버튼 — 약속 ON 시 모달 없이 즉시 발송, OFF 시 모달 */}
        <button
          type="button"
          onClick={handleRequestClick}
          disabled={togglingPromise}
          className="w-full flex items-center justify-center gap-1.5 px-3 py-2.5 rounded-lg bg-stone-900 text-white text-sm font-semibold hover:bg-stone-800 transition disabled:opacity-50"
        >
          <Send size={13} />
          {requestSentAt ? '재발송' : '후기 요청 보내기'}
        </button>

        {/* 발송 시각 표시 */}
        {requestSentAt && (
          <div className="mt-2 flex items-center gap-1.5 text-[11px] text-neutral-500">
            <MessageCircle size={11} />
            <span>최근 발송 · {formatDate(requestSentAt)}</span>
          </div>
        )}

        <RelatedActivitySection items={related} />
      </Card>

      {showRequestModal && (
        <ReviewRequestModal
          saleId={id}
          source={source}
          customerName={customerName}
          customerPhone={customerPhone || ''}
          hasRepairItem={hasRepairItem}
          sourceType={sourceType}
          alreadySent={!!requestSentAt}
          onClose={() => setShowRequestModal(false)}
          onSent={() => {
            setShowRequestModal(false);
            onChanged?.();
          }}
        />
      )}
    </>
  );
}
