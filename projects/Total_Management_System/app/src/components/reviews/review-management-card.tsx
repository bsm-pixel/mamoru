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
import { Star, MessageCircle, CheckCircle2, Send, Info } from 'lucide-react';
import toast from 'react-hot-toast';
import { ReviewRequestModal } from '@/components/sales/review-request-modal';
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

interface Props {
  source: ReviewSource;
  id: string;
  customerName: string;
  customerPhone: string | null;
  promisedAt: string | null;
  requestSentAt: string | null;
  submittedAt: string | null;
  /** 변경 후 부모 쿼리 invalidate / refetch 트리거 */
  onChanged?: () => void;
  /** sale source일 때 수리 상품 포함 여부 (ReviewRequestModal 기본값 결정) */
  hasRepairItem?: boolean;
}

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
  requestSentAt,
  submittedAt,
  onChanged,
  hasRepairItem = false,
}: Props) {
  const [togglingPromise, setTogglingPromise] = useState(false);
  const [showRequestModal, setShowRequestModal] = useState(false);
  const [related, setRelated] = useState<RelatedActivity[]>([]);

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

  const handleTogglePromise = async () => {
    if (togglingPromise) return;
    setTogglingPromise(true);
    try {
      const next = !promisedAt;
      const res = await fetch('/api/reviews/promise', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ source, id, on: next }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        toast.error(err.error || '저장 실패');
        return;
      }
      toast.success(next ? '리뷰 약속 체크' : '약속 해제');
      onChanged?.();
    } catch (e) {
      toast.error(`오류: ${String(e)}`);
    } finally {
      setTogglingPromise(false);
    }
  };

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
              ? 'bg-terracotta/10 border-terracotta/40'
              : 'bg-white border-neutral-200 hover:bg-neutral-50'
          }`}
        >
          <div className="flex items-center gap-2">
            <span
              className={`w-4 h-4 rounded border-2 flex items-center justify-center text-[10px] font-bold transition ${
                promisedAt
                  ? 'bg-terracotta border-terracotta text-white'
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

        {/* 후기 요청 발송 버튼 */}
        <button
          type="button"
          onClick={() => {
            if (!customerPhone) {
              toast.error('고객 연락처가 없어 발송할 수 없습니다');
              return;
            }
            setShowRequestModal(true);
          }}
          className="w-full flex items-center justify-center gap-1.5 px-3 py-2.5 rounded-lg bg-indigo-black text-cream text-sm font-semibold hover:bg-indigo-black/85 transition disabled:opacity-50"
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
