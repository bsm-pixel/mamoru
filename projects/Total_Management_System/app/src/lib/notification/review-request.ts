/**
 * 후기 요청 알림톡 발송 헬퍼
 *
 * 자동 발송(consultation/repair completed/delivered 시점) 및 수동 발송(상세 패널 카드 버튼) 양쪽에서
 * 동일한 URL/템플릿/페이로드로 발송하기 위한 추출 헬퍼.
 *
 * 사용처:
 *   - /api/reviews/request (수동 발송)
 *   - /api/consultation/[id] (자동 발송 — 가드 통과 시)
 *   - /api/repair/[id] (자동 발송 — 가드 통과 시)
 */

import { sendNotification, type NotifyTemplate } from '@/lib/notification/make-webhook';

export type ReviewSource = 'consultation' | 'repair' | 'sale';
export type ReviewType = 'consult' | 'repair' | 'purchase';

export interface ReviewSubject {
  source: ReviewSource;
  /** 사장님 페이지 URL과 솔라피 치환에 사용되는 source_id
   *  - consultation: consultations.unique_id
   *  - repair: repairs.as_id
   *  - sale: offline_sales.sale_number */
  sourceId: string;
  customerName: string;
  customerPhone: string;
  reviewType: ReviewType;
  /** 상담/수리의 세부 유형 (store_visit / field_request / talk_consult / repair) */
  subtype?: string;
}

const REVIEW_FORM_BASE = 'https://page.mamoru.kr/projects/reviews/page_review.html';

/**
 * 복원수리 진행유형(proceed_type) → 후기 subtype 자동 유도
 *
 * 153(2026-09-16) 이후 후기 자동발송이 "기본 ON" 이라 사장님이 약속 유형을 직접 고르지 않는 건이 대부분이다.
 * 그때 subtype 이 비면 후기 폼이 어떤 수리였는지 모른 채 열린다 → proceed_type 에서 그대로 유도한다.
 * (사장님이 카드에서 직접 고른 값이 있으면 그 값이 우선 — 이 함수는 폴백일 뿐)
 */
export function repairSubtypeFromProceedType(proceedType?: string | null): string | undefined {
  switch ((proceedType || '').trim()) {
    case '직접방문': return 'direct_visit';
    case '방문수거': return 'pickup';
    case '직접발송': return 'delivery';
    default: return undefined;
  }
}

/** 후기 요청 알림톡 발송 — 템플릿/URL/페이로드 통일된 단일 진입점 */
export async function sendReviewRequestNotification(subject: ReviewSubject) {
  const reviewUrl =
    `${REVIEW_FORM_BASE}?type=${subject.reviewType}` +
    `&uid=${encodeURIComponent(subject.sourceId)}` +
    `&name=${encodeURIComponent(subject.customerName)}` +
    (subject.subtype ? `&subtype=${encodeURIComponent(subject.subtype)}` : '');

  let template: NotifyTemplate;
  if (subject.reviewType === 'repair') template = 'as_review_request';
  else if (subject.reviewType === 'purchase') template = 'purchase_review_request';
  else template = 'review_request';

  const typeLabel =
    subject.reviewType === 'repair' ? '복원수리'
    : subject.reviewType === 'consult' ? '상담'
    : '제품구매';

  return sendNotification({
    template,
    phone: subject.customerPhone,
    name: subject.customerName,
    data: {
      id: subject.sourceId,
      uid: subject.sourceId,
      consult_uid: subject.sourceId, // 솔라피 #{consult_uid} 치환용
      as_uid: subject.sourceId,      // 솔라피 #{as_uid} 치환용
      order_uid: subject.sourceId,   // 솔라피 #{order_uid} 치환용 (아임웹 주문 버튼과 동일 변수 — 오프라인도 버튼 작동)
      review_type: subject.reviewType,
      type_label: typeLabel,
      subtype: subject.subtype || '',
      review_url: reviewUrl,
    },
  });
}
