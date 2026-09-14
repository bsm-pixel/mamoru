# 리뷰(후기) 프로세스 흐름도
> 최종 업데이트: 2026-09-13 · 대상: 후기 요청→작성→저장→노출 전 경로
> ⚠️ `REVIEW_SYSTEM_BRIEF.md`(2026-02, GAS+시트 계획)는 **폐기된 구계획**. 현행은 아래 TMS(Vercel)+Supabase.

---

## 1. 4단계 파이프라인 (한눈에)

```
[배송완료/수리완료/상담완료]
  → (자동/수동) 후기요청 알림톡 발송 (솔라피+Make)
    → 고객이 링크 클릭: page.mamoru.kr/projects/reviews/page_review.html (커스텀 폼)
      ※ 아임웹 네이티브 구매평 아님 (아임웹 구매평/QA 탭은 CSS로 숨김 처리)
  → 고객 별점·내용·사진 작성 → POST /api/reviews/submit
    → Supabase reviews 테이블 저장 (status = pending, 자동승인 토글 ON이면 approved)
  → (사장님) 리뷰 관리(/reviews)에서 승인 → status = approved
  → 노출: GET /api/reviews/public (approved 만) → 4개 화면에 렌더
     · 메인 후기 iframe · 메인 하단 · 리뷰 전체보기(/62) · ★상품 상세 하단 위젯
```

---

## 2. 각 단계 상세

### (A) 후기요청 알림톡 — 어디로 보내나
- **목적지 URL**: `https://page.mamoru.kr/projects/reviews/page_review.html?type={type}&uid={sourceId}&name=...&subtype=...`
- 조립 단일 진입점: `app/src/lib/notification/review-request.ts` (`REVIEW_FORM_BASE`, `sendReviewRequestNotification`)
- 템플릿 분기: `purchase`→`purchase_review_request` / `repair`→`as_review_request` / `consult`→`review_request`
- `sourceId` = offline_sales.sale_number / repairs.as_id / consultations.unique_id / **orders는 order.id**

**발송 트리거(진입점)**
| 소스 | 자동 발송 위치 | 조건 |
|---|---|---|
| **아임웹 주문(orders)** | ① `lib/imweb/sync.ts` (아임웹 상태가 배송완료로 동기화될 때 직접 발송 — **토글 무관**) ② `track-delivery` 크론 [1] 롯데 배달완료 감지 ③ `api/orders/[id]/pickup-complete` 매장 픽업 | 미발송(`orders.review_requested_at`) + 전화O. **orders는 '약속' 개념 없음**. uid = **orders.id(UUID)** 로 통일 |
| 판매(offline_sales) | `track-delivery` 크론 [3] 배송완료 감지 | **review_promised_at 있는 건만**(약속) |
| 복원수리(repairs) | `track-delivery` 크론 [2] + `api/repair/[id]` | **약속 있는 건만** |

> 📌 **2026-09-14 알림톡 항상 발송 원칙** — 설정의 알림톡 on/off 토글(전체·유형별·자동 후기요청)과 화면별 "알림톡 발송" 체크박스를 모두 삭제. 고객 행동·상태 변화마다 항상 발송하고, 같은 내용 2통만 서버 규칙으로 막는다(무상 0원 입금확인 · 직접방문 현장결제 `paid_on_site`). 비상 정지 = 설정의 해당 Make 웹훅 URL 비우기.

| 상담 | 정책상 **영구 수동만** | — |
| 수동 발송 | 판매/주문 상세 "후기 요청" 버튼 → `api/reviews/request` | 언제든 |
- 실제 도착 3요소: ①솔라피 템플릿 검수완료 ②Make 리뷰 분기 ③토글. 미발송 진단은 이 순서.

### (B) 고객 작성 → 저장
- 폼: `projects/reviews/page_review.html` (API 베이스 `app-eta-sandy-75.vercel.app/api/reviews`)
  - 로드 `/api/reviews/info?uid&type` (고객·구매품목 조회) → purchase는 주문 내 여러 제품 중 **리뷰할 제품 선택**(`imweb_product_no` 전송)
  - 사진 `/api/reviews/upload` (Supabase `review-photos` 버킷)
  - 제출 `/api/reviews/submit`
- 저장: **Supabase `reviews` 테이블** (아임웹·구글시트 아님)
  - 컬럼: `type/subtype/name/phone/stars/content/photo_urls/source_id/product/product_group/status/is_best/meta`
  - **상품 매칭 키**: `meta.imweb_product_no` + `product_group` (submit route에서 기록) ← 상품별 필터의 핵심
  - 중복 방지: purchase `source_id = uid:productNo`(제품별 1회), 그 외 uid
  - **purchase uid 정규화** `lib/reviews/resolve-purchase-uid.ts` (info·submit 공용): 숫자 아임웹 주문번호가 오면 `orders.id` 로 변환. 2026-09-13 이전 track-delivery [1] 이 uid=imweb_order_no 로 보내 링크가 404 나던 버그의 **옛 링크 호환**. 새 발송은 orders.id

#### 상담 리뷰 subtype 결정 규칙 (2026-09-13)
- 저장 우선순위(`api/reviews/submit`): **URL subtype → 판매건 약속 subtype(review_promised_subtype) → 판매채널 정규화**
- 판매채널 정규화 SSOT = `lib/reviews/consult-subtype.ts` `saleChannelToConsultSubtype`: `store→store_visit` · `field→field_request` · `talk→talk_consult` · 레거시 `offline→store_visit`(사장님 결정 2026-09-13, 마이그 147) (online 등은 원값 유지)
- ⚠️ 솔라피 `review_request` 버튼 URL엔 `subtype`이 없음 → 약속 칩 없이 보낸 판매건은 판매채널로 결정됨. 과거 `'store'` 원시값이 저장돼 관리자 "상담·매장" / 고객 "상담"으로 갈라졌던 버그 → 코드 정규화 + 마이그 146(과거분 교정)
- 표기 통일: 작성 폼(`api/reviews/info` `consultTypeLabel` + `page_review.html`) = `상담 · 직접방문/출장/톡상담`, 노출 칩 = `상담·직접방문`. `info`는 CT- 상담이 없으면 판매건(OS-) fallback으로 이어짐(과거엔 404)

### (C) ★상품 상세 하단 자동 노출 — YES (승인된 것만)
- 위젯: `projects/reviews/ImwebWidgetCode_product_reviews.html`
  - **아임웹 상품 상세 "공통 영역" 코드위젯**으로 1회 paste → **모든 상품에 자동 적용**(상품마다 수동 X)
  - 상품 자동감지: URL `?idx=`(아임웹 상품번호) → referrer → href 순, 감지 실패 시 위젯 숨김
  - `GET /api/reviews/public?imweb_no={번호}` → **그 상품(군) 리뷰만** 평균별점/분포/포토/필터/정렬/모달 렌더
  - 아임웹 네이티브 구매평/QA/탭은 CSS `display:none` 로 숨김
- **승인 게이트**: `api/reviews/public`는 `status='approved'`만 반환.
  - 제출 시 기본 `pending` → 사장님이 `/reviews`에서 **수동 승인**해야 노출
  - `system_settings['review.auto_approve']='true'`(자동노출 토글)면 작성 즉시 approved (기본 OFF)

### (D) 상품별 매칭 — YES
- `api/reviews/public`: `imweb_no` → `products.product_group` resolve → 같은 **제품군 전체 리뷰 묶어** 노출(색상만 다른 동일 가위 등)
- product_group 미설정 상품은 `meta.imweb_product_no` 단독 매칭. 위젯 `?group=` 수동 지정도 지원(1순위)

---

## 3. 노출 4개 화면 (동일 데이터 `api/reviews/public` 공유)
| # | 화면 | 파일 | 배포 |
|---|---|---|---|
| 1 | 메인 후기 iframe | `projects/main/iframe_main_reviews.html` | 자동(GitHub Pages) |
| 2 | 메인 하단 | `projects/main/page_main_btm.html` | 자동 |
| 3 | 리뷰 전체보기 `/62` | `projects/reviews/page_reviews.html` | 자동 |
| 4 | **상품 상세 하단** | `projects/reviews/ImwebWidgetCode_product_reviews.html` | ⚠️ **아임웹 수동 재-paste**(자동반영 X) |
- 리뷰 표시/모달 로직 수정 시 **4곳 전부 동기화** ([feedback_review_display_sync])

---

## 4. 핵심 파일 맵

### 알림톡 발송
- `app/src/lib/notification/review-request.ts` — URL/템플릿 조립 SSOT
- `app/src/lib/imweb/sync.ts` · `api/cron/track-delivery` [1] · `api/orders/[id]/pickup-complete` — 아임웹 주문 후기 자동 발송 (uid=orders.id)
- `app/src/lib/notification/make-webhook.ts` — 솔라피-Make 연동
- 진입점: `api/reviews/request`, `api/imweb/orders/[id]/review-request`, `api/consultation/[id]`, `api/repair/[id]`

### 폼 / API
- `projects/reviews/page_review.html` — 작성 폼(고객)
- `app/src/app/api/reviews/submit/route.ts` — 저장
- `app/src/app/api/reviews/info/route.ts` — 고객·구매품목 조회
- `app/src/lib/reviews/resolve-purchase-uid.ts` — purchase uid 정규화(아임웹 주문번호→orders.id)
- `app/src/app/api/reviews/public/route.ts` — 공개 조회 + 상품(군)별 필터 (approved만)
- `app/src/app/api/reviews/upload|upload-bulk/route.ts` — 사진
- `app/src/app/api/reviews/[id]/route.ts` — 승인/숨김/베스트/삭제(수동 큐레이션)
- `app/src/app/api/reviews/auto-match|promise|promised/route.ts` — 약속·자동매칭 부가

### 관리 / DB
- `app/src/app/(dashboard)/reviews/page.tsx` — 리뷰 관리(승인, 자동노출 토글 `review.auto_approve`)
- Supabase `reviews` 테이블

### 관련 문서
- `TMS_FLOW_REVIEW_EVENT.md` — 리뷰 이벤트(당첨자) 흐름
- 메모리: `reference_review_modal_unified`(표시 4파일 통일), `project_review_event`, `reference_solapi_templates`(후기 3종)

---

## 5. 요약 판정
- (A) 알림톡 → **커스텀 폼**(page.mamoru.kr/projects/reviews/page_review.html), 아임웹 네이티브 아님
- (B) 저장 = **TMS Supabase `reviews`** (시트/아임웹 아님)
- (C) 상품 상세 하단 자동 노출 = **YES(승인된 것만)** — 공통 코드위젯이 상품 자동감지, 기본은 수동 승인(자동노출 토글 OFF 기본)
- (D) 상품별 매칭 = **YES** (imweb_product_no → product_group)
