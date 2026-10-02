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

## 당첨자 배송 (2026-10-02, 마이그 156)

```
선정 화면에서 당첨 저장 (reviews.event_month/event_rank = SSOT, 그대로)
   ↓  TMS 「당첨자 배송」 패널 (reviews/event 하단) — 열 때 당첨자 기준 자동 동기화
[당첨 안내 보내기] → review_event_won 알림톡 (#{token} = 배송지 입력 링크)
   │                (템플릿 승인 전: [링크] 복사 → 카톡 채널 채팅으로 직접 전송, 흐름 동일)
   ↓
page.mamoru.kr/projects/reviews/page_event_address.html?t=<token>
   · 이름·연락처 고정 / 주소 = 저장값 → 없으면 예전 주소(고객정보 → 복원수리 → 아임웹 주문) 미리 채움
   · 저장 → review_event_shipments 주소 + customers 주소 갱신(없으면 고객 생성, source=manual) + 사장님 푸시
   ↓
[송장 생성] → 롯데 ALPS bookShipment (품목 "리뷰이벤트 {상품}") → ALPS에서 출력
   ↓
track-delivery 크론 [5] (30분) → 집하 감지 → shipped_at CAS → review_event_shipped 알림톡 1회
                                 → 배달완료 감지 → delivered_at
```

- 표: `review_event_shipments` (당첨자 1명=1행, review_id 유니크, 토큰 32자). 송장 전이면 당첨 취소 시 `cancelled_at`(soft).
- API: 관리 `api/reviews/event/shipments` (GET 목록 / POST notify·invoice·cancel_invoice·save_address) · 고객 `api/reviews/event-address` (GET/POST, 토큰=권한, CORS *)
- 공용 로직: `lib/reviews/event-shipments.ts` (sync·예전주소·고객주소갱신·알림톡 2종)
- 알림톡: `review_event_won`(변수 name·rank·prize·token·address_link) / `review_event_shipped`(name·prize·courier·tracking) → `EVENT_TEMPLATES` → `webhook_event`(Make 06 EVENT)
- 송장 생성 후엔 고객 페이지에서 주소 수정 불가(발송 준비 중 화면). 바꾸려면 TMS에서 송장 취소(집하 전만) → 주소 수정 → 재생성.
- 당첨 안내는 **게시와 동시 자동발송이 아니라 버튼**: 템플릿 승인 전 게시하면 Make 분기가 없어 조용히 사라지는데 TMS엔 '보냄'으로 남기 때문(Make 200=성공 판정 구멍). 승인 후 버튼 한 번에 미발송 전원.

## 홍보물 제작용 비공개 읽기 주소 (2026-10-02)
- `GET /api/reviews/event-preview?key=<키>[&month=YYMM]` — Claude 스킬 `mamoru-review-event-post`(피그마 인스타 게시물·스토리)가 읽는다.
- 공개 API(event-public)와 달리 **비공개(draft) 상품 설정 + 게시 전 당첨자**까지 준다 → 고객 공개 전에 게시물을 먼저 만들 수 있다.
- 🔒 키=`system_settings review_event.preview_key`(불일치·없음=404). 당첨자는 event-public 과 같은 마스킹(`백*민 님 (3562)`)만 — 실명·전체 번호 없음. 읽기 전용·no-store.
- 키·스킬 원본은 **저장소 밖** `C:마모루드라이브ClaudeSecrets`(repo 는 page.mamoru.kr 로 공개되므로 키를 커밋하지 않는다). 키 교체=system_settings 값 변경 + 스킬 파일의 주소 갱신.
- 주소를 못 읽는 환경의 안전망 = 추첨 화면 [게시물용 명단 복사].

## 추첨 초기화 (2026-10-02)
- 위치: **설정 > 시스템 「리뷰 추첨 초기화」** (년·월 선택). 추첨 화면엔 두지 않는다 — 녹화·화면공유에 "다시 뽑기" 버튼이 보이면 조작으로 오해받음(사장님 결정). 같은 이유로 룰렛은 목표 인원을 다 뽑으면 [추첨하기] 버튼이 사라지고 완료 화면이 된다.
- 동작: 그 달 `reviews.event_*` 4필드 해제 + `review_event_shipments` 취소표시(soft). 상품·마감일·응모 시작일은 유지. 다시 추첨해 같은 사람이 뽑히면 배송 행은 되살아난다.
- 🔒 차단: 발표됨(announced) / 당첨 안내 발송 / 고객 주소 저장 / 송장 생성 중 하나라도 있으면 거부(409). 판정은 서버(`lib/reviews/event-reset.ts`)가 실행 직전에 다시 한다.
- 이력: `system_settings` key `review_event.reset_log` (누가·언제·몇 명, 최근 20건). API `api/reviews/event/reset` (GET 미리보기 / POST 실행).

## 미구현/후속
- 솔라피 2종 검수 → Make `06 EVENT`에 분기 2개 연결(템플릿 필터 + 솔라피 모듈 + 변수 매핑, dlq 확인)
- 인스타 응모(Phase 2) 보류.
