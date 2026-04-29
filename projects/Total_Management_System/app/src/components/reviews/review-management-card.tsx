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

import { useState } from 'react';
import { Card } from '@/components/ui/card';
import { Star, MessageCircle, CheckCircle2, Send } from 'lucide-react';
import toast from 'react-hot-toast';
import { ReviewRequestModal } from '@/components/sales/review-request-modal';
import type { ReviewSource } from '@/lib/notification/review-request';

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
