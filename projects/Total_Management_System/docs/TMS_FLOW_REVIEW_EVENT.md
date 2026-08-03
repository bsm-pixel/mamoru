# TMS FLOW — 리뷰 이벤트 (자사몰 후기 기반)

> 고객이 진짜 후기를 남기도록 유도 → 월별 베스트 선정 → 추첨 상품 증정. 고객 페이지는 iframe 자동 연동(fetch).

## 전체 흐름
```
제품 배송 → 후기 알림톡(기존 크론) → 고객이 자사몰(아임웹)에 후기 작성
        → reviews 테이블 적재(기존)
                    │
          ┌─────────┴──────────┐  (월말)
          ▼                    ▼
  TMS 「리뷰 이벤트 관리」    이달 이벤트 설정
  (/reviews/event)           (상품 1·2·3등 / 마감일 / 히어로)
   · 그 달 후기 카드         status: draft → live(진행중) → announced(발표)
   · 1·2·3등 등수 토글
   · 표시명·경로 오버라이드
          │ [저장/발표]
          ▼
  DB: review_event_config(월설정) + reviews.event_month/event_rank/…(당첨 마킹)
          │
          ▼
  공개 API  GET /api/reviews/event-public  (마스킹·CORS·발표된 것만)
   { current:{prizes,deadline,hero}, past:[{month,label,winners[]}] }
          │  (fetch)
          ▼
  고객 페이지  page.mamoru.kr/projects/reviews/page_review_event.html (아임웹 iframe)
   · Hero(1등 상품·카운트다운) ← current
   · 이달의 상품 ← current.prizes
   · 지난 당첨자(월 탭 아카이브) ← past
   · fetch 실패/빈값 = 정적 폴백(밴드 숨김·'준비중')
```

## 상태(status) 의미
- `draft` : 임시저장(비공개). 공개 API 미노출.
- `live` : 진행중. current 로 노출 → Hero·이달의 상품.
- `announced` : 발표됨. past 로 노출 → 지난 당첨자 아카이브(해당 월 탭).

## 데이터
- `review_event_config` (마이그 123): month(PK,'YYMM'), deadline, announce_at, hero_image_url, prizes jsonb `[{rank,name,desc,image_url,count}]`, status.
- `reviews` 마킹: `event_month`, `event_rank`(1/2/3), `event_display_name`(선택), `event_route`(선택). 별도 테이블 아닌 SSOT.
- 마스킹: 서버(`src/lib/reviews/mask.ts`) — 홍**님 / 010-****-32**. 공개 API는 원본 전화/실명 미노출.

## 파일
- 관리 화면: `src/app/(dashboard)/reviews/event/page.tsx` (리뷰관리 헤더 [리뷰 이벤트 관리] 버튼 진입)
- 관리 API(인증): `src/app/api/reviews/event/route.ts` (GET 월 설정+후기 / POST 설정 upsert+당첨 마킹 재설정)
- 공개 API: `src/app/api/reviews/event-public/route.ts` (CORS·5분 캐시)
- 이미지 업로드: 기존 `/api/reviews/upload-bulk` (`review-photos` 버킷) 재사용
- 고객 페이지: `projects/reviews/page_review_event.html` (fetch 렌더 + 정적 폴백), 임베드 스니펫 `projects/reviews/iframe_review_event.html`

## RLS 주의 (2026-08-03 버그·수정)
- `review_event_config` 는 **RLS 켜짐 + 정책 없음** → anon/**authenticated 롤 쓰기 42501 차단**(SELECT는 0행 반환).
- ∴ 관리 API(`api/reviews/event`)는 **인증 확인(createServerSupabaseClient.getUser) 후 `createServiceClient()`(service role)로 DB 작업**. 공개 API도 service role. RLS는 켠 채 유지(=service role만 접근, 더 안전).
- 증상이었던 것: 게시 눌러도 반영 안 됨 → 원인은 authenticated 롤 upsert가 RLS에 막혀 저장 실패(사장님이 실패 메시지 놓침). service role 전환으로 해결.

## 응모 시작일(선택, 마이그 124)
- `review_event_config.entry_start`(timestamptz, nullable). 응모자 집계 **하한**. 비우면 그 달 1일부터.
- 용도: 첫 회차 등 **과거 후기까지 포함**해 선정. 예 8월 이벤트에 시작일 4/1 → 4~8월 후기 풀에서 선정(당첨자는 event_month='2508'). 끝은 항상 그 달 말일.
- 관리 GET: 하한 우선순위 = `?start=YYYY-MM-DD`(미리보기) → `config.entry_start` → 그 달 1일. UI에서 날짜 바꾸면 `reloadPool`로 풀만 즉시 갱신(설정 유지, 만지던 마킹 보존).
- 다음 달부턴 비우면 평소대로. 중복 수상 위험 없음(다음 달 풀은 그 달만).

## 월 경계 주의
- 응모자 집계는 `created_at`(timestamptz)을 **KST 월**로 묶음(`kstMonthRange`, UTC밀림 회피). `toISOString().slice(0,7)` 금지. 날짜 하한 변환 `kstDateStartISO`.

## 미구현/후속
- 당첨 알림톡(발표 시 당첨자에게) — 미연결. 필요 시 솔라피 템플릿 추가.
- 인스타 응모(Phase 2) 보류.
