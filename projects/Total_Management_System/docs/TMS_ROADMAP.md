# TMS (Total Management System) 전체 작업 로드맵

> 최종 목적: 마모루 운영의 주문·배송·수리·재고·알림을 하나의 시스템에서 관리
> 최종 수정: 2026-07-20 — 창고 로케이션(정위치) 관리 신규 완료. 마이그 112~115 실행 완료, 배포 승인 대기

---

## ✅ 완료 (07-18~20): 창고 로케이션(정위치) 관리 — 마이그 `112`~`115`

"이 제품 어디 있지?"를 없애기 위한 **제품(SKU) 단위 정위치** 관리. 업계 개념으로는 로케이션 코드 + 정위치(Fixed Location).
**수량은 위치별로 쪼개지 않는다** — `raw_stock`이 파생값이 되면 재고 갱신 지점 10곳 이상이 영향권이라 의도적으로 배제. 전 커밋에서 `stock_quantity`/`raw_stock` 대입부 변경 0건 확인.
- [x] `112` `warehouse_locations` 테이블 + `products.location_id`(`ON DELETE SET NULL` — 자리를 지워도 제품이 깨지지 않음)
- [x] `113` `warehouse_racks` — 렉을 1급 개체로 (자리 0개인 렉도 존재 가능)
- [x] `114` `bin_row` — 한 단 안에 **행×열 수납함**(예: 3단에 6행 10열 가위 보관함) 지원. 행이 1개면 코드에 행번호를 안 붙여 **기존 라벨 `R01-2-A` 그대로 유효**
- [x] `115` 1단 = **맨 아래**로 정정 (건물 층수와 동일한 직관). 상/중/하단 라벨 폐기 → `N단`으로 통일, 렌더는 큰 번호가 위
- [x] 배치도 `/inventory/map` — 렉·단·칸 시각화, 칸 클릭 시 우측에 제품+재고, **제품명·SKU 검색 → 해당 칸 하이라이트**
- [x] 렉별 **수정 모달**(단 추가·＋열/＋행·칸 개별 삭제·렉 통째 삭제) — 진입 화면은 보는 것에만 집중시키려 편집 컨트롤을 전부 모달로 이동
- [x] 렉별 **배치도 인쇄** — A4 한 장에 렉 하나. `@page margin 0` + 시트 A4 실치수 고정 + 칸 높이 역산(`fitCellH`)으로 한 장을 보장. 세로/가로, 모든 렉 인쇄(렉마다 한 장). *위치라벨 낱장 인쇄는 폐기 — 사장님이 원한 건 "그림 보고 제품 찾기"였다*
- [x] **칸 위치 명시 배치**(`cellGridPos`) — 칸을 배열 순서대로 흘리면 중간 칸 삭제 시 뒤가 전부 밀려 지도가 실물과 어긋난다. 배치도·수정화면·인쇄 **세 곳 모두** 열·행을 명시 지정
- [x] **한 칸에 여러 품목** — 종수 배지 + 이름 여러 줄(+외 N종), 담기는 체크박스 다중 선택 후 한 번에. (재고 이벤트 제품을 한 자리에 몰아 넣는 실제 운영 패턴)
- [x] **가로 병합**(마이그 116 `col_span`) — 편집화면 '병합 모드'에서 같은 행 인접 칸을 골라 하나로. 흡수 칸 제품은 왼쪽으로 이동(수량 로직 무관). 넓은 칸은 ✂로 분리. 배치도·인쇄가 폭(`col_span`)을 반영
- [x] **삭제↔생성 토글** — 편집화면에서 칸 클릭 시 점선 빈자리로 남고, 그 자리를 다시 클릭하면 재생성(`add_cell`). 전체 격자를 그려 빈 위치를 보여줌
- [x] **메인 배치도 정리** — 칸에서 코드(A1) 숨기고 **모델명 크게**, 빈칸은 점선 테두리만(글자 없음). 코드는 클릭 시 우측 상세에만
- [x] 두 모달 모두 `preventAutoClose` — 작업 중 Escape·바깥 드래그로 닫히지 않음
- Phase 2(미착수): `product_serials.location_id`(준비/전시 존 개별 위치, 086 감사 트리거에 필드만 추가하면 이동 이력이 따라옴) · ABC 슬로팅 추천

---

## ✅ 완료 (05-15): 자체 도메인 전환 — 마무리 운영 작업

사장님이 같은 날 진행한 사용자측 작업 — DNS 1~4시간 전파 + SSL 자동 발급 + 코드위젯 paste + 솔라피 재검수.
- [x] 카페24 호스팅 어드민에서 NS 변경 (`bns1~4.hostcocoa.com`) → Google·Cloudflare·LGU 1~4시간 내 전파, KT 캐시는 약간 지연
- [x] 아임웹 어드민 DNS 레코드 사전 입력 (NS 전환 전): MX `SMTP.GOOGLE.COM` + SPF + google-site-verification + CNAME `page.mamoru.kr` → `bsm-pixel.github.io`. **메일 끊김 0** 보장 흐름
- [x] 아임웹 대표 도메인 `mamoru.kr` 전환
- [x] 아임웹 SSL 결제·발급 — COMODO Basic SSL 38,500원/년(VAT 포함). Sectigo 발급, ~2026-11-29 만료 후 자동 갱신 (사장님 결제 1년 내 두 번 갱신)
- [x] GitHub Pages page.mamoru.kr SSL 발급 — Let's Encrypt 무료, 90일 자동 갱신
- [x] 아임웹 코드위젯 13개 + 헤더 코드 paste (main 5 + brand + consulting 4 + as 2 + reviews)
- [x] 솔라피 후기 템플릿 3종 URL 갱신 + 재검수 신청 — `review_request` (상담, 현재 미사용이지만 카탈로그 청소 + 미래 대비) / `purchase_review_request` (판매·구매) / `as_review_request` (복원수리, `type=as` → `type=repair` 표준 통일). 카카오 검수 1~3 영업일 대기, 옛 템플릿은 검수 중에도 발송 가능 → 운영 무중단
- [x] 자가 검증: `https://mamoru.kr` 자물쇠 정상, 메인 코드위젯 디자인 정상, 옛 `bsm-pixel.github.io` URL 자동 redirect 확인
- 잔여 2건: 카페24 쇼핑몰 연결 해제 (2분) / 아임웹 스크롤 로고 403 Forbidden (아임웹 상담원 연결 — S3 권한 누락 추정)

---

## ✅ 완료 (05-14): 자체 도메인 전환 — 커밋 `749c0af`

GitHub Pages (`bsm-pixel.github.io/mamoru/...`) → **`page.mamoru.kr`**, 아임웹 쇼핑몰 (`mamoruscissors63682.imweb.me`) → **`mamoru.kr` apex**. TMS Vercel 도메인은 영구 유지(사장님 결정).
- [x] WHOIS 분석으로 등록대행자=가비아 / 등록기관 실제 진입은 카페24 호스팅 어드민 확인. 도메인 갱신 22,000원 결제 (만료 2027-05-24)
- [x] 메일 안전 흐름: NS 변경 **전에** 아임웹 DNS 어드민에 MX/SPF/google-site-verification/CNAME page 사전 입력 → 카페24 NS → 아임웹 NS(`bns1~4.hostcocoa.com`) 변경. Google Workspace 메일 끊김 0
- [x] 알림톡 호환: 옛 `*.github.io` URL 은 GitHub Pages custom domain 자동 redirect 로 계속 작동 → 이미 발송된 알림톡 영향 없음
- [x] 코드 일괄 갱신: TMS 7개 + GAS + 아임웹 iframe 13개 + v10_trendy 2개 + 페이지 3개 아임웹 도메인 + 공통 preconnect 2개. TypeScript 0 에러, Vercel/GitHub Pages 배포 success
- [x] `type` 표준 4인 회의 결정: `consult / repair / purchase` (사장님 솔라피 콘솔의 `type=as` 는 재검수 사이클에 `type=repair` 로 통일)
- [x] 메모리 카탈로그 신규 작성: `memory/reference_solapi_templates.md` — 26종 템플릿 trigger/변수/링크/코드 위치 전수
- [x] 문서 6개 도메인 갱신: PAGES_INDEX, TMS_SYSTEM_ARCHITECTURE, SOLAPI_TEMPLATES_ORDERS, FLOW_change_request, figma_master_spec, figma_tokens_usage
- 사장님 마무리 5단계 (출근 후 15~20분): `DOMAIN_MIGRATION_NEXT_STEPS.md` 참조 — 아임웹 SSL 확인 / 대표 도메인 전환 / 카페24 쇼핑몰 연결 해제 / 솔라피 후기 URL 2개 갱신 + 재검수

---

## ✅ 완료 (05-13~14): 출장 상담 알림톡 흐름 fix 묶음 (커밋 209ef65~4a9cd78)
GAS→TMS 이식 시 출장 흐름 알림톡 변수가 곳곳에서 빠져 있어 일괄 복구. DB·Make·솔라피 변경 없음 (TMS 코드만).
- [x] 출장 확정(`field_confirmed`) 알림톡이 SMS로 떨어지던 버그 — `address`·`change_request_link` 누락 복구 (`209ef65`)
- [x] 수동 일정 변경 — 확정 후 시간 변경 시 `field_rescheduled` 알림톡 발송 분기 추가 + data 페이로드에 `visit_date`/`visit_time` 추가 + 버튼/모달 라벨 동적화 (확정 후엔 "수동 일정 변경") (`a215a09`)
- [x] notify route — defensive `visit_date`/`visit_time` 추가 (`useRescheduleConsultation` 훅 경로도 호환) (`a215a09`)
- [x] 일정변경 요청 페이지 404 — `resched` route 가 `uid` 도 받게 + page_change_request.html 가 `?uid=` 로 호출 (`f083f43`)
- [x] 출장 취소(`field_cancelled`) 알림톡 본문에 `#{visit_date}` 등 원본 변수 그대로 발송되던 버그 — `visit_date`/`visit_time`/`address` 추가 (`4a9cd78`)
- 흐름도(TMS_FLOW_CONSULTATION) + 매뉴얼(MANUAL_CONSULTATION) 갱신.

## ✅ 완료 (05-13): 매입(발주→입고) 흐름 보강 — 배포·SQL(081) 완료 (커밋 bc03567~6056a63)
가위는 제작품이라 주문≠입고가 흔하고, 업체가 발송 시작 후 잔금만 먼저 보내는 경우도 많음 → 결제/입고 독립 + 실수령 보정 흐름 추가.
- [x] 입고 전 잔금 지불 기록 — `발주완료`/`선납완료`에서 "잔금 지불 처리 (입고 전)" → `balance_paid_at` 기록·`balance_amount=0`, 상태 유지. 입고 후엔 기존 "잔금 완료"로 → `잔금완료`. (`bc03567`, 실제 UI 패널 누락 fix `5e5c901`)
- [x] 입고검수 — "입고 확인" → 품목별 `주문/실수령` 편집 모달 + 재계산 프리뷰. 재고는 실수령 기준 증가(아임웹도), `purchase_order_items.received_quantity` 기록(`quantity`는 보존), 실수령≠주문이면 `total_amount`·`balance_amount` 재계산. 마이그레이션 081. (`09af5e7`)
- [x] 입고 전 발주 품목 수정 — "편집"을 `작성중`/`발주완료`/`선납완료` + 잔금 미지불일 때 허용 (입고·잔금지불·취소 이후 잠김). 수정 시 총액·잔금 자동 재계산. (`6056a63`)
- [x] 입고/결제 어느 순서로 끝나든 마지막에 자동 `잔금완료` 전이.
- 마이그레이션 081 사장님 SQL Editor 실행 완료. 흐름도/매뉴얼(TMS_FLOW_INVENTORY, MANUAL_INVENTORY) 갱신.

## ✅ 완료 (05-12~13): 매출 3분할 — 배포·SQL 실행 완료 (커밋 5c96d1b~4de93f2)
1단계(05-12, 배포·검증OK)에 이어 2단계. 총매출 = B2C 제품 + B2B 제품(딜러·아카데미·납품) + 복원수리 전체(A 접수 + B 판매RS + C 납품RS). RS는 제품 매출에서 제외 → 복원수리로만. 가정: RS 항목엔 할인 미적용.
- [x] 2-A useHubStats — HubStatsResult.sales 에 salesB2C/salesB2B 신설 (RPC 후처리 + fallback 양쪽 계산)
- [x] 2-B RPC `078_hub_stats_b2c_b2b_split.sql` — get_hub_stats 'sales' 객체에 salesB2C/salesB2B + deliveries 통합
- [x] 2-C 대시보드 — 총매출 KPI = salesB2C+salesB2B+monthRepairAmount, 3분할 내역, '제품 판매' 카드
- [x] 2-D 회계 리포트 — 복원수리=A(접수,paid_at)+B(판매RS)+C(납품RS), 제품=B2C+B2B(납품 포함), by_product/margin RS 제외, total_revenue 중복 제거. offline_sales 취소/반품 필터 추가. reports/page 탭별 표시
- [x] 2-E `079_customers_default_repair_price.sql` — 거래처별 복원수리 기본 단가 + 고객 상세 화면 입력 + 납품 "+B2B수리" 자동 채움
- [x] 3-A 판매 입력 "복원수리" 모드 — 마모루(1만)/타사(2만) 자루+단가, offline_sale_items category='RS' 저장
- [x] 후속 fix (`4f67218`): 회계 "이번 달" 프리셋 UTC 버그(4/30→5/1) fix / 회계 제품 탭 B2B 두 줄 통합 / 판매입력에서 RS 더미 제품·"RS" 카테고리 버튼 숨김 / 복원수리 모드 "배송비 3,000원" 토글 / 납품 "+B2B수리" 추가항목·배송비에 category='RS'
- [x] 배송비 자루 fix (`4de93f2` + RPC `080_hub_stats_exclude_shipping_qty.sql`): 배송비 = 복원수리 매출에 금액 포함, "자루 수"에서만 제외 (대시보드·회계·RPC 전부)
- [x] 마이그레이션 078·079·080 사장님 SQL Editor 실행 완료
- (선택, 미진행) 납품 페이지 "이번달 매출"을 회계 B2B 제품과 동일 기준으로 정렬 / 기존 RS 더미 제품 is_active=false
- 상세: `memory/project_tms_repair_revenue_split.md`

## 🔴 다음 할 일 (2026-03-03 기준)

### 1순위: Make Router 연결 + 솔라피 재검수 대기
- [x] 솔라피 23종 전체 검수 승인 (상담 17종 + 복원수리 5종 + 계약서 1종) ✅ 2026-03-03
- [x] BC 버튼 chatExtra 한글 불가 발견 → 메타데이터 전체 제거 결정 ✅ 2026-03-03
- [ ] 솔라피 BC 메타데이터 제거 후 재검수 대기 (1~3 영업일)
- [x] Make Router 상담 17종 분기 연결 ✅ 2026-03-12
- [x] Make Router 복원수리 6종 분기 연결 ✅ 2026-03-12
- [x] Make Router 계약서 1종 분기 연결 ✅ 2026-03-12
- [x] 판매→후기 요청 알림톡 흐름 완성 (uid 매핑 + info API OS-* 대응) ✅ 2026-04-09
- [ ] Make Router 리뷰 작성 요청 템플릿 연결 (후기 템플릿 미연결)

### 2순위: UI 동작 검증 (수동)
- [ ] /sales/new 판매입력 동작 확인 (딜러→딜러가, 아카데미→아카데미가 자동 적용)
- [ ] /contracts/new 전자문서 UI 동작 확인 (제품 모달 + 서명 2개 + 수령방법 + 선납/잔금)
- [ ] /products/[id]/serials 시리얼 등록 동작 확인
- [ ] /customers 고객 목록/상세 동작 확인
- [ ] /products/new 제품 등록 (4단 가격: 소매/딜러/아카데미/매입) 동작 확인
- [ ] /suppliers B2B 거래처 (딜러/아카데미/매입처 서브탭) 확인
- [ ] /purchasing/new 발주 작성 → 입고 → 재고 증가 흐름 확인
- [ ] /inventory 재고 현황 표시 + 저재고 필터 확인
- [ ] /reports 회계 리포트 + 엑셀 다운로드 확인
- [ ] /reports/transaction 거래내역서 인쇄 확인

### ✅ 완료 (03-22): 거래처 구조 재설계 + 유형별 가격
- [x] customer_type 재정의: retail/online/dealer/academy/supplier ✅
- [x] 제품 4단 가격: 소매가/딜러가/아카데미가/매입가 ✅
- [x] 판매 등록: 고객 유형별 가격 자동 적용 ✅
- [x] B2B 거래처 페이지: 딜러/아카데미/매입처 서브탭 ✅
- [x] 고객 목록 매입처 제외 + 딜러/아카데미 필터 ✅
- [x] 사이드바 B2B 거래처 + Handshake 아이콘 ✅
- [x] 고객 등록/수정: 주소·매장명·메모·유형 필드 추가 ✅

### ✅ 완료 (03-21): 아임웹 연동 + 재고 관리 완성 + UI 개선
- [x] 아임웹 v2 API 상품 동기화 (설정 > 상품 동기화 버튼) ✅
- [x] 아임웹 재고 실시간 연동 (입고/판매/조정 → 자동 반영) ✅
- [x] 온라인 주문 TMS 재고 자동 차감 + 취소/환불 복구 ✅
- [x] 발주 입고 멱등성 버그 수정 (재고 중복 증가 방지) ✅
- [x] 매입처 드롭다운 (제품/발주 폼 자동완성) ✅
- [x] 재고 수동 조정 (파손/실사 보정 모달 + 이력) ✅
- [x] 회계 COGS/마진 분석 (제품별 랭킹 + 마진 엑셀) ✅
- [x] 미지급금/미수금 현황 (매입처 잔금 + 고객 미수금) ✅
- [x] 사이드바 네비 재배치 (업무 플로우 5그룹) ✅
- [x] 사이드바 활성 메뉴 가시성 개선 ✅
- [x] 제품 마스터-디테일 레이아웃 (우측 패널) ✅

### ✅ 완료된 이전 할 일
- [x] **상담 수동 일정 확정 + 재고조사 인쇄 + 알림톡 변수 fix** (2026-05-03 저녁) — ① **상담 수동 일정 확정 (079)**: 우측 상세 패널에 "수동 일정 확정" 버튼 신설. 고객이 접수한 건에 대해 DM/유선 협의된 일정을 즉시 확정 처리. 모달(날짜+시간+메모) → /api/consultation/[id] PATCH → 자동: 캘린더 sync + 알림톡 + 이력 기록. 사장님 룰: closed_dates 검증 X. 신규 파일 `manual-confirm-modal.tsx`. 노출 조건: 매장/출장 + 활성 상태(completed/cancelled/in_progress 제외). ② **재고조사 인쇄 기능**: 창고재고 화면 우측 상단 "재고조사 인쇄" 버튼. 화면 적용 필터/정렬 그대로 인쇄 + 카테고리 자동 그룹화 + 그룹별 소계 + 실측 빈 칸(노란 배경) + 비고 영역. window.open() 새 탭 패턴(po-print-modal과 동일). 신규 파일 `inventory-print-modal.tsx`. ③ **알림톡 변수 매칭 fix**: PATCH 알림톡이 admin-create 대비 status/name/phone/days/memo/change_request_link 누락 → 카카오 3109(잘못된 파라미터) → SMS 대체 발송 사고. admin-create와 동일 필드로 통일하여 모든 PATCH 상태 변경에 일관 적용. 매뉴얼/흐름도 갱신: MANUAL_CONSULTATION + MANUAL_INVENTORY + TMS_FLOW_CONSULTATION. commits: `652ab3a` `4f3f485` `22ccc77`
- [x] **가이드 페이지 위계 정합 + 사장님 룰 박제** (2026-05-03) — 사장님 강조 룰("전체 페이지 흐름 + UI/UX 위계 + 다른 영역 중복 점검 필수") 박제 (`memory/feedback_page_holistic_review.md`). ① **상담/복원수리 가이드 CTA 3중 보장**: Hero + 모바일 floating + 페이지 끝 (Brand Guide ADDENDUM § 5 PC 수치 680px 정합) ② **iframe wrapper 부모 측 floating CTA**: iframe 안 fixed가 iframe document 끝점 기준이라 작동 X → 부모 wrapper에 직접 fixed로 추가 (`consulting/iframe_guide.html` + `as/iframe_guide.html`) ③ **모바일 CTA 중복 제거**: page 내부 floating 영구 숨김 (wrapper로 단일화) + Hero secondary "Q&A 보기" 제거 (탭과 중복) ④ **과정안내 Step 1 위계 강화**: "두 가지 진행방식" 18px 굵게 + method-group 카드 강화 + "또는" divider (다크/라이트 두 페이지 동일 패턴) ⑤ **"마모루 컨설팅" 탭 시장 문제점 3블록 제거**: brand 페이지와 70~80% 중복 → 페이지 간 메시지 분담 명확화 (brand=Why / guide=How) ⑥ **상담철칙 ↔ 다크 메시지 swap**: 실무 원칙 먼저 → 결심 강화. 메인 슬로건 "거짓 없는 본질" + 사장님 직접 상담철칙 03 신규. **TMS 내부 흐름 무변경**. commits: `108d292` `1b115c2` `439e377` `55051c3` `ec6ebfc` `298ebd2` `aeb3128` `e2b9ad3` `f412525` `67da51e` `85e9cad` `1885869`
- [x] **상담 달력 관리 + 휴무 SSOT 통합 (078)** (2026-05-02) — 사장님 요청: 상담 카테고리에 "달력관리" 화면 신설하여 4개월(현재월 ~ +3) 달력에서 클릭 한 번으로 휴무 토글. 정기 휴무 요일(매주 반복) + 임시 휴무일(특정 날짜)을 한 화면에서 통합 관리(SSOT). 핵심 룰: **막힘은 고객 셀프 예약에만 적용, 사장님 측 흐름(일정수동등록/시간제안)은 항상 유동**(memory/feedback_consultation_blackout_rule.md). ① 신규 화면 `/consultations/calendar` (4개월 grid + 7개 요일 토글 + 모달 휴무 사유 입력) ② 신규 API `/api/consultation/blackouts` (closed_dates GET/POST/DELETE) + `/api/consultation/settings` (consultation_settings GET/PATCH) ③ 신규 hook use-blackouts.ts ④ 설정 → 상담 설정의 "휴무 요일" + "특별 휴무일" UI 삭제 → "달력 관리로 이동" 링크로 대체 (SSOT 정리) ⑤ 시각 표시: 임시 휴무(빨간) > 정기 휴무(회색) > 일반 ⑥ 휴무 사유는 고객에게 미노출 (API 응답에서 자동 제외 — `/api/consultation/public/settings/route.ts:46`). 회귀 안전: 알림톡/리마인더/Google Calendar 동기화 모두 closed_dates 검증 0이라 정상 유지. commits: `ca95025` (1차) + `b270439` (옵션 C 통합)
- [x] **리뷰 모달 viewport 중앙 fix** (2026-05-02) — 사장님 보고: 후기 페이지 리뷰 카드 클릭 시 모달이 화면 밖에 떠 backdrop만 보이던 문제. 1차 분석 잘못(단순 `position: absolute → fixed` CSS 치환 제안) → 사장님 지적("아임웹 iframe 안인 거 알고 체크한 거 맞아?")으로 재분석. 진짜 원인: `iframe_reviews.html` wrapper에 `MAMORU_REQUEST_VIEWPORT` 응답 코드 누락 (다른 wrapper `as/iframe_form.html` L50-58은 보유). fix: 같은 패턴으로 응답 코드 추가. 메모리 영구 보관: `memory/reference_iframe_pages.md` (페이지별 iframe 환경 식별 표) + `memory/feedback_consultation_blackout_rule.md` (사장님 룰). commit `f066717`
- [x] **TMS 즉각 반영 풀 연동 + 사장님 보고 3건 fix (075)** (2026-04-30 심야 +3) — ① **경비 카테고리 동적화**: `expenses/page.tsx`가 `useSetting('accounting.expense_categories')`로 읽도록 변경. 설정에서 추가한 카테고리가 즉시 경비 입력 화면에 반영(이전 hard-coded `DEFAULT_EXPENSE_CATEGORIES` 상수). ② **"일정 재요청" 카운트 버그 fix**: TS fallback (`use-dashboard-stats.ts:154`)와 RPC v3 (`migrations/075_hub_stats_rpc_v3.sql`) 양쪽에서 needAction status 배열의 `pending_admin` 잘못 포함 제거 — 신규 출장 상담이 재요청으로 중복 카운트되던 버그 해결. ③ **복원수리 매출 정의 통일 (옵션 A 발생 기준, 사장님 합의)**: A채널 paid_at 조건 제거, B채널 category='RS' 정확화 + 0원 무상 제외, RPC v3에 monthRepairMamoru/Other/B2B 분리 신설. ④ **invalidate 풀 연동**: `lib/query/invalidate-keys.ts` helper 신설 → 모든 sale/repair mutation onSuccess에서 `invalidateFinancialQueries(qc)` — 새 판매·수리 등록·수정·취소 직후 대시보드 매출 즉각 갱신 (이전 60s 지연). ⑤ **staleTime 합리화**: sales-stats 60s→30s, products 5분→1분. 흐름도/매뉴얼 갱신: TMS_FLOW_ACCOUNTING / MANUAL_DASHBOARD / MANUAL_REPAIR / TMS_SYSTEM_ARCHITECTURE. 사장님 SQL Editor에서 075 RPC 실행 완료. commit `20d4752`
- [x] **브랜드 소개 페이지 정비 + 네이버 인앱 page_main_top 잘림 fix** (2026-04-28) — ① `page_intro.html` 카피 톤 재작성(영업 현장 멘트 반영) + 메인 퀵네비 chip 아임웹 페이지 ID 갱신(/53→/60·/58→/61·/intro→/31, '후기→고객후기') ② problem/tech/consult 카드 아이콘을 인라인 SVG에서 `./icons/*.svg` 외부 파일 참조로 통일(page_diag와 동일 패턴), `brand/icons/` 폴더 신설(visit·door·face·gijun·up·blunt 임시 svg) ③ `<img>` 사이즈를 박스 가득 채우는 `100% 100%` + `object-fit: contain` 적용, `consult-step-icon` 박스 사이즈 명시(모바일 64 / PC 80) ④ 모바일 퀵네비(`.intro-quicknav-wrap`) sticky CSS + 부모 sticky 퀵네비 통신용 postMessage(MAMORU_QUICKNAV_POS / MAMORU_SCROLL_TO / MAMORU_SCROLL_TO_OFFSET) 추가 — 아임웹 코드위젯 v3 가이드 작성(iframe ID `mamoruIntroFrame` 매칭 + origin 검증). 단, 회사소개 페이지의 모바일 sticky 동작은 미완 상태 — TODO.md에 보류 등록 ⑤ **네이버 인앱 page_main_top 히어로 잘림 fix**: 기존 7중 안전망(ResizeObserver·이미지·폰트·다단 setTimeout·resize·REQUEST_HEIGHT·4가지 측정값)에 4개 트리거(pageshow/visibilitychange/focus/orientationchange) + 다단 setTimeout 50ms·150ms 초입 송신 추가. BFCache 복원·앱 포그라운드 복귀·viewport 동적 변경 케이스 보완. commits: `f8103db a727e05 23fcaf0 6b0c404 7b2260a 88e999c`
- [x] **빠른 송장 발급 + 상담 모달 개선 + 일정변경 재진입 가드 + cancel 캘린더 + 푸시 dedup** (2026-04-27) — ① **빠른 송장(`/manual-invoices`)** 사이드바 '판매' 그룹에 신설: `manual_invoices` 신규 테이블(매출 KPI 미반영) + `<CustomerAutocomplete>` + 품목명 직접 입력(50자) + 발급 즉시 송장번호 큰글씨 + 클립보드 + 오늘 발급 미니 리스트 인라인 취소. 고객 상세 거래 타임라인에 '빠른송장' 항목 통합. 알림톡 발송 v1 제외(향후 옵션 가능). 상세: `docs/TMS_FLOW_SALES.md` ② **상담 일정 수동 등록 모달 개선** (CreateConsultationModal): 고객명/연락처 텍스트 입력 → `<CustomerAutocomplete>` + 외부 "+ 신규 고객 등록"(중첩 `<CustomerCreateModal>`). 출장요청 주소 입력 → `<DaumPostcodeButton>` (postcode + road + detail 분리 저장). 출장요청 시 consultations.postcode INSERT(컬럼 기존). 매장↔출장 토글 시 비어있을 때만 자동 채움. 자식 모달 떠 있을 때 부모 dialog `open=false`로 일시 닫기 → native dialog × fixed-div 중첩 가려짐 회피. ESC 가드. ③ **다음 우편번호 검색 재사용 컴포넌트 추출**: `<DaumPostcodeButton>` 신설 + 기존 3곳(고객 신규/고객 수정/배송 설정) 동일 코드 75줄 정리. ④ **CustomerAutocomplete `disableInlineNewForm` prop 추가**: 상담 모달은 인라인 폼 비활성, 다른 사용처 기본값 false → 회귀 0. ⑤ **일정변경 페이지(`page_change_request`) 재진입 가드**: GAS getReservationInfo의 CONFIRMED/ASSIGNED 가드가 TMS Vercel API 이행 시 누락됐던 것 복구. reservation API 응답에 `canRequestChange: boolean` 추가, 페이지에서 cancelled→cancel-done / completed→already-completed / reschedule_requested→already-rescheduled 전용 안내 화면 분기. cancel-done 디자인 재사용. ⑥ **page_change_request safeClose 카카오 인앱 닫기 fix**: `kakaoweb://closeBrowser`(iOS 전용)만 호출하던 단순 패턴 → page_suggest/page_result와 동일한 UA 분기 + history.back→window.close fallback. mamoru.kr 강제 이동 폴백 제거(메모리 규칙 위반). ⑦ **cancel API 캘린더 동기화 누락 보강**: admin-create/resched는 호출하던 syncConsultationToCalendar를 cancel만 빠뜨려 DB는 cancelled인데 캘린더 일정이 잔류하던 문제. `after()` 블록 추가. 잔여 정리용 **신규 admin endpoint** `GET /api/consultation/admin-cleanup-calendar` — cancelled+google_event_id 모든 행에 일괄 sync, idempotent. ⑧ **푸시 알림 중복 도착 dedup**: `firebase/client.ts`의 onMessage가 SW의 push 이벤트와 중복으로 알림 생성하던 문제 → onMessage 핸들러 제거(SW가 모든 표시 담당). subscribe API single-token-per-user 정책(같은 user의 옛 토큰 자동 삭제) + 신규 cleanup-others endpoint + 설정 → 알림 → "이 기기만 알림 받기" 버튼. notificationclick의 매칭 범위를 same-origin 전체로 확장(기존: /dashboard·/consultations·/repairs 3경로만 → 다른 페이지에서 클릭 시 새 창 열림 회피, focused 우선). ⑨ **Vercel ignoreCommand ghost commit 회복**: `git diff --quiet $VERCEL_GIT_PREVIOUS_SHA $VERCEL_GIT_COMMIT_SHA -- .`이 shallow clone에서 옛 SHA를 못 찾아 23시간 전부터 모든 TMS 배포가 exit 128로 차단되던 문제 → `bash -c '... 2>/dev/null \|\| exit 1'` fallback으로 자가 치유 사이클. 상세: `docs/TMS_FLOW_CONSULTATION.md` § 푸시·재진입·취소 / `docs/TMS_FLOW_SALES.md` § 빠른 송장
- [x] **메인 페이지 YouTube 섹션 신설 + iframe 잘림 만성 버그 해결 + 진단 Lottie 도입** (2026-04-26 오후) — ① 메인 상단(`page_main_top`)에 [3.5] mm-videos 섹션 신설: Quick Nav ↔ 라인업 배너 사이, lite-youtube 패턴 자체구현(라이브러리 0), 모바일 가로 스크롤(75vw 카드+marquee) / PC 3열 grid, 모노크롬 ▶ 버튼, 16:9 썸네일. 첫 영상 등록 `1LZhDgEyrMA`. ② 모바일 영상 가로 스크롤이 페이지 세로 스크롤을 가로채던 버그 → `touch-action: pan-x` + `overflow-y: hidden` + `overscroll-behavior-x: contain` 적용. ③ 다음 카드 peek 노출 강화(80vw→75vw, 마스크 90%→94%). ④ **메인 영역 3개 파일(top/btm/main) `initIframeComm()` 강화** — 사장님이 모바일 PWA에서 자주 겪던 "Trust Numbers 아래 콘텐츠 통째 잘림" 만성 버그 근본 해결. 🔒 수정금지 마커 → ⚙️ 강화 적용으로 변경(외부 인터페이스 100% 호환). 보수적 over-estimation 7중 안전망(4가지 측정값 최댓값/img load/document.fonts.ready/다단 setTimeout/resize 디바운스/REQUEST_HEIGHT 양방향/?debug=1 콘솔). ⑤ 진단 페이지(`page_diag`) Q_FEEL/Q_STYLE/Q_HABIT 세 질문에 Lottie(.json) 도입 — 인프라 이미 완비됨 확인(dotlottie-player + renderGif 분기 + lottieUrl 슬롯), Q_FEEL 상단 통합 사용 가이드 코멘트만 추가. 진단 SVG 아이콘 18종 정비. 첫 Lottie 작업 등록(`style_go.json`/`style_back.json`). 모두 TMS 외 페이지 영역, TMS 내부 흐름 무변경. commits: `6d67c63 baebe0d 7135ccf d4c6f75 6a149fd e36243a 84cbf89 7f7c79e`
- [x] **고객 대면 페이지 21개 Brand Guide v1.0 정합성 정비** (2026-04-26) — 아임웹 연동/GitHub Pages 페이지 전수 점검. 5개 영역(brand 1 / main 5 / consulting 7 / as 4 / reviews+verify 3) / 5 commit. 변경: ① 폐기 컬러 80→10곳(88% 감소, 잔존은 page_intro 히어로 다크+골드 예외 8곳 + page_diag 가드라인 주석 2곳 — 모두 의도적) ② Pretendard 5파일→0(가이드 서체 통일) ③ `--mm-gold/--gold-dark/--trust-gold` 변수명 60+곳 → `--mm-ink` 일괄(시각 무변, 미래 골드 회귀 차단) ④ page_main_btm body color cream→void (잠재 가독성 폭탄) ⑤ 외부 브랜드 컬러(네이버 #03C75A 등) 모두 모노크롬화 ⑥ page_as_report 카카오 인앱 닫기 fallback iOS/안드로이드 분기 + mamoru.kr 강제 이동 제거 ⑦ 본문 12px→13px 부분 상향(라벨 11px 하한은 유지) ⑧ 더블 br/단독 끝 br 위반분만 정리. **TMS 내부 흐름 무변경** — 흐름도 4종 다이어그램 영향 없음. 펜딩: 메인 후기 외부 API `app-eta-sandy-75.vercel.app` 자체 도메인 이관 + Polish 단계(진단 progress bar / Trust Number 카피 / Masonry 동적 컬럼). 플랜 파일: `C:\Users\user\.claude\plans\validated-spinning-tower.md`
- [x] **상담관리 일정 수동 등록 기능** (2026-04-24) — 인스타DM/유선 등 외부 채널 접수 건을 TMS 상담 탭 우측 "일정수동등록" 버튼으로 수기 입력. 매장방문/출장요청 2유형, 중복 체크(phone_normalized+일시) 시 경고 모달, 출장은 카카오 지오코딩 자동, 알림톡(confirmed/field_confirmed) 발송 선택, Google Calendar/리마인더/푸시 자동 편입. 상세: `docs/TMS_FLOW_CONSULTATION.md § 6`
- [x] **리마인더 알림톡 방문주소 치환 누락 복구** (2026-04-24) — cron SELECT 에 address_road/detail 추가, FIELD_REMIND_24H/2H 의 #{address} 정상 치환.
- [x] **아임웹 배너 모달 X 버튼 제거** (2026-04-24) — 하단 "오늘 하루 보지 않기" / "닫기" 버튼만으로 통일 (미니멀 톤).
- [x] **메인 페이지(/main) 카피·디자인 리뉴얼** (2026-04-23) — Trust 섹션 ZERO 상술 중공형 효과 / Quick Nav 2줄 분리 + 아임웹 카테고리 URL 실제 매핑 / 슬로건 '거짓 없는 본질' / 복원수리·컨설팅·도구섹션 카피 리라이팅. 20-40대 여성 타겟 감성 반영.
- [x] **푸시 알림 중복 표시 버그 수정** (2026-04-23) — FCM SW 경로와 Realtime Notification 경로의 tag 불일치로 같은 알림이 2번 떠 보이던 문제. 3개 파일(send-push.ts / firebase-messaging-sw.js / use-push-notifications.ts)에서 tag 통일 → 브라우저 자동 dedup. 9가지 푸시 전부 영향.
- [x] **상담 리마인더 오발송 수정** (2026-04-23) — Make 시나리오 FIELD_REMIND_24H 모듈 내부 Solapi 템플릿이 매장방문 내용으로 잘못 연결. 사장님이 Make 설정 수정 완료. 코드 변경 없음 (Make만).
- [x] **푸시 알림 테스트 발송 패널** (2026-04-23) — 설정 > 알림·연동에 9가지 타입별 테스트 버튼. 기본/리뷰/상담3종/출장3종/복원수리/주문. 기기 연결·설정 토글·라우트 동작을 단계별 진단 가능.
- [x] **고객 행동 푸시 알림 사일런트 실패 수정** (2026-04-23) — Vercel 서버리스의 fire-and-forget Promise가 응답 반환 후 잘려나가던 문제. reviews/submit · consultation/public/confirm · resched 3개 라우트를 after() 래퍼로 감싸 실행 보장. 리뷰/출장확정/재요청 푸시 정상 동작.
- [x] **아임웹 상품 동기화 사일런트 실패 버그 수정** (2026-04-23) — products.sku UNIQUE 제약 + error 체크 누락으로 31건 조용히 실패하던 버그. 매칭 우선순위 재설계(imweb_product_no → sku → 신규), `.maybeSingle()`, error 반환값 체크, UI에 total_fetched/synced/failed 구분 표시 + 실패 내역 펼쳐보기.
- [x] **아임웹 배너/팝업 원격 관리 (Phase 1+2)** (2026-04-22) — TMS 설정에서 이미지 업로드/토글/5초 슬라이드로 아임웹 메인 모달 배너 원격 관리. 최대 5장, 이미지별 개별 링크, 스와이프 지원. 상세: `docs/TMS_FLOW_IMWEB_BANNER.md`
- [x] **Google Calendar 연동 Phase 1 MVP** (2026-04-21) — OAuth 2.0 + 상담 확정/변경/취소 자동 동기화 / 출장 확정건 "일정변경" 버튼 / 설정 UI / 재동기화. 상세: `docs/TMS_FLOW_CONSULTATION.md § 5`
- [x] **복원수리 접수 MAKE 웹훅 pickup_date 누락 수정** (2026-04-21) — 방문수거 4종 수거예정일 `YYYY년 MM월 DD일 (X요일)` 포맷 추가
- [x] **복원수리 삭제 시 OS 푸시 알림 자동 회수** (2026-04-21) — tag에 as_id 포함 + Service Worker DISMISS 메시지 + push_notifications 행 정리
- [x] **GAS 데드코드 정리** (2026-03-04) — AppSheet 래퍼 4개 + 테스트/데모 함수 제거, consulting -115줄 / as -50줄
- [x] **솔라피 23종 검수 승인 + BC 메타데이터 이슈 해결** (2026-03-03) — 한글 chatExtra 불가 → 메타데이터 제거 + 해피톡 사전 입력 폼으로 대체
- [x] **TMS 일정변경 알림톡 버그 수정** (2026-03-03) — change_request_link 누락 + template 자동 분기 (rescheduled/field_rescheduled)
- [x] **간편진단 태블릿 PWA** (2026-03-01) — `/diagnosis` 13개 질문 조건부 분기, manifest 분리, 결과 테이블
- [x] **Phase A~F 자체 ERP 전환 완료** (2026-03-01) — 아래 Phase ERP 섹션 참조
- [x] 고객 자동완성 검색 (GET /api/customers/search + CustomerAutocomplete 공유 컴포넌트)
- [x] 고객 신규등록 (POST /api/customers)
- [x] 판매 저장 시 이카운트 자동 동기화 → 이카운트 코드 제거 후 TMS 단독 관리
- [x] 계약서 입력에도 CustomerAutocomplete 적용 (email/address 확장)
- [x] 이카운트 ERP 6/6 API 검증 + 정식 인증키 + 프로덕션 검증 (2026-02-28)
- [x] GAS Script Properties 설정 (TMS_REPAIR_SYNC_URL, CRON_SECRET) (2026-02-28)
- [x] 복원수리 접수 → GAS → TMS 동기화 테스트 통과 (2026-02-28)

---

## Phase 1: 주문·배송 코어 ✅ 완료

**목적:** 아임웹 주문을 TMS로 가져오고, 롯데택배(ALPS) 송장을 생성/취소/추적하는 기본 파이프라인 구축

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 1-1 | 아임웹 → TMS 주문 동기화 | ✅ | 아임웹 v2 API 폴링 → Supabase upsert |
| 1-2 | 주문 목록/상세 UI | ✅ | 상태 탭 필터, 검색, 페이지네이션 |
| 1-3 | 송장 생성 (ALPS 접수) | ✅ | 롯데택배 apiSndOut → 12자리 운송장 생성 |
| 1-4 | 송장 취소 (소프트 취소) | ✅ | TMS cancel_pending → ALPS 수동 집하취소 → 자동 감지 |
| 1-5 | 아임웹 송장 연동 | ✅ | 배송대기(STANDBY) 상태에서 PATCH invoice 입력 |
| 1-6 | 동기화 보호 | ✅ | cancel_pending/shipping 상태 덮어쓰기 방지 |

### 운영 플로우 (확정)
```
주문 접수 → 아임웹 "배송대기 처리" (수동 1회)
         → TMS 송장 생성 (ALPS 접수 + 아임웹 자동 연동)
         → 배송 추적

취소 시 → TMS 송장취소 (cancel_pending)
       → ALPS 집하취소 (수동)
       → TMS 자동 감지 → cancelled
       → 아임웹 고객취소 or 관리자 취소 (수동)
```

### 제약사항
- 아임웹 v2 API: 읽기 + 송장입력만 가능 (상태변경 code -99)
- ALPS: 취소 API 없음 → 소프트 취소 패턴
- 고객 취소: 상품준비중까지만 가능 (배송대기 이후 관리자 직접)

---

## Phase 1.5: 상담 취소 연동 ✅ 완료

**목적:** TMS 상담 취소 시 구글 캘린더 삭제 + 아임웹 슬롯 해제 + 알림톡 자동 발송

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 1.5-1 | TMS → GAS 취소 연동 | ✅ | cancelViaGAS GET 호출 → 캘린더 삭제 + 슬롯 해제 |
| 1.5-2 | GAS doGet cancelConsultation 핸들러 | ✅ | POST body 제한 → GET 쿼리 파라미터 방식 전환 |
| 1.5-3 | 취소 알림톡 자동 발송 | ✅ | Make webhook → Solapi 취소 알림톡 |
| 1.5-4 | 취소 확인 모달 (안전장치) | ✅ | 3개 상담유형 모두 확인 모달 거쳐 취소 |
| 1.5-5 | 백그라운드 처리 (빠른 응답) | ✅ | after() API로 응답 즉시 반환, 후속작업 비동기 |

### 운영 플로우 (확정)
```
취소 시 → TMS 취소 버튼 → 확인 모달 ("정말 취소하시겠습니까?")
       → 취소 확정 → Supabase 상태 변경 → UI 즉시 반영
       → (백그라운드) GAS 캘린더 삭제 + 아임웹 슬롯 해제
       → (백그라운드) 알림톡 자동 발송
```

### 해결한 이슈
- GAS 웹앱 외부 POST body 수신 불가 → GET 쿼리 파라미터 전환
- Vercel 환경변수 `\n` 문제 → 재설정
- fire-and-forget → after() 백그라운드 처리 (Vercel 실행 컨텍스트 보장)

---

## Phase 1.6: 고객 일정 변경/취소 요청 ✅ 완료

**목적:** 확정된 예약(매장방문/출장)에 대해 고객이 직접 일정 변경 또는 취소를 요청 — 알림톡 → 셀프서비스 폼 → TMS 연동

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 1.6-1 | page_change_request.html | ✅ | GitHub Pages 고객 대면 페이지 (예약조회+변경/취소 폼) |
| 1.6-2 | GAS API: getReservationInfo | ✅ | uid로 예약 정보 조회 (CONFIRMED/ASSIGNED만) |
| 1.6-3 | GAS API: submitChangeRequest | ✅ | 매장/출장 분기 처리 (아래 참조) |
| 1.6-4 | 확정 알림톡 change_request_link 배치 | ✅ | CONFIRMED/CONFIRMED_BY_TOKEN/RESCHEDULED/FIELD_CONFIRMED에 추가, REMINDER에서 제거 |
| 1.6-5 | TMS 상태/UI 연동 | ✅ | change_requested 타입·라벨·색상·전이·탭·상세카드 |

### 매장방문 / 출장 분기 (2026-02-21 확정)

**매장방문 일정변경:**
```
확정 알림톡 → "일정확인/변경" 버튼 → 셀프서비스 페이지
  → [일정 변경] 선택 → "기존 예약 취소 후 재예약" 안내
  → 버튼 클릭 → adminCancel(skipNotify) 자동취소 → 접수페이지(mamoru.kr/52) 이동
  → 고객이 새로 예약 → 즉시 확정 → 새 confirmed 알림톡
```
- 관리자 개입 없음, 알림톡 없음 (페이지가 피드백)

**출장 일정변경:**
```
확정 알림톡 → "일정확인/변경" 버튼 → 셀프서비스 페이지
  → [일정 변경] 선택 → 요청사항 입력 → 제출
  → GAS: CHANGE_REQUESTED + Supabase + 이메일 + change_request_received 알림톡
  → TMS 출장 "처리대기" 탭 → 관리자 "새 시간 제안"
```

**취소 (공통):**
```
셀프서비스 페이지 → [예약 취소] → 사유 선택 → 제출
  → adminCancel(skipNotify) 즉시 CANCELLED (캘린더 삭제 + 슬롯 해제)
  → 페이지: "예약이 취소되었습니다" + 관리자 이메일
  → 알림톡 미발송 (관리자 수동 취소 시에만 알림톡 발송)
```

### TMS UI 변경 (2026-02-21)
- 매장방문: `변경/취소` 탭 제거 (자동취소+재예약이므로 불필요)
- 출장: `처리대기` 탭에 change_requested 포함, 버튼 "새 시간 제안"
- 상세 페이지: 고객 요청 카드(주황색) — 비고에서 `[고객 변경요청]` 파싱 표시
- 상태 전이: change_requested → suggested/confirmed/on_hold/cancelled

### 수동 작업 필요 (솔라피/Make) — Phase 1.6-6
- [x] 솔라피: change_request_received 템플릿 등록 + 검수 승인 ✅ 2026-03-03
- [ ] 솔라피: confirmed 템플릿에 "일정확인/변경" WL 버튼 추가 → 재검수 (BC 메타데이터 재검수 시 함께 처리)
- [ ] 솔라피: rescheduled 템플릿에 "일정확인/변경" WL 버튼 추가 → 재검수
- [ ] 솔라피: field_confirmed 템플릿에 "일정확인/변경" WL 버튼 추가 → 재검수
- [ ] Make: 확정 3개 시나리오에 change_request_link 변수 매핑 (Make Router 연결 시)
- [ ] Make: CHANGE_REQUEST_RECEIVED 이벤트 분기 + 솔라피 모듈 연결

---

## Phase 1.7: 출장 일정 제안 캘린더 UI ✅ 완료

**목적:** 출장 상담 일정 제안 페이지를 텍스트 버튼 나열 → 캘린더 + 라디오 카드 + 재요청 폼 통합 UI로 업그레이드

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 1.7-1 | page_suggest.html 캘린더 UI | ✅ | 달력(제안 날짜 강조) + 라디오 카드 + 확정 모달 |
| 1.7-2 | page_suggest.html 재요청 통합 | ✅ | 하단 토글 → textarea(reason) + 재요청 모달 |
| 1.7-3 | page_reschedule.html 리다이렉트 | ✅ | 기존 알림톡 링크 호환, page_suggest.html로 자동 이동 |
| 1.7-4 | Code.gs reason 파라미터 | ✅ | markResched에 reason 읽기 → 이메일+비고 컬럼 반영 |
| 1.7-5 | GAS 배포 + GitHub Pages 배포 | ✅ | clasp push @285 + Pages 서빙 확인 |
| 1.7-6 | 실 환경 테스트 | 📋 | 실 토큰 테스트, 카카오 인앱 확인 |

---

## Phase 1.8: 알림톡 13종 출장/매장/톡상담 분기 구현 ✅ 완료

**목적:** 기존 매장방문 전용이던 알림톡을 출장/매장/톡상담 유형별로 분기 — GAS 6개 패치 + TMS 7개 파일 변경

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 1.8-1 | GAS adminFieldDelay 신규 함수 | ✅ | 출장 지연 안내: visit_time + delayMin → visit_time_revised 계산 |
| 1.8-2 | GAS doGet fieldDelay/talkReady 액션 | ✅ | TMS_SYNC_KEY 인증, TMS→GAS 호출 엔드포인트 |
| 1.8-3 | GAS submitConsultation 톡상담 알림 | ✅ | type='톡상담' 시 TALK_RECEIVED 자동 발송 |
| 1.8-4 | GAS adminCancel 출장/매장 분기 | ✅ | 출장→field_cancelled, 매장→cancelled |
| 1.8-5 | GAS adminReschedule 출장/매장 분기 | ✅ | 출장→field_rescheduled, 매장→rescheduled |
| 1.8-6 | GAS sendReminders_ 출장/매장 분기 | ✅ | 출장→field_remind_24h/2h, 매장→remind24/2 |
| 1.8-7 | TMS make-webhook.ts 타입 확장 | ✅ | NotifyTemplate 7종 추가 + TEMPLATE_EVENT_MAP |
| 1.8-8 | TMS delay API route 생성 | ✅ | POST /api/consultation/delay → GAS fieldDelay 호출 |
| 1.8-9 | TMS hooks 추가 | ✅ | useFieldDelay + useStartTalkConsult |
| 1.8-10 | TMS [id]/route.ts 자동 알림 분기 | ✅ | getAutoNotifyTemplate() — 상담유형별 템플릿 선택 |
| 1.8-11 | TMS talk-consult-list.tsx 상담시작 | ✅ | "상담 시작" → useStartTalkConsult → talk_ready 발송 |
| 1.8-12 | TMS field-request-list.tsx 지연안내 | ✅ | "지연 안내" 버튼 + DelaySelectModal (5~30분) |
| 1.8-13 | 고객 페이지 경고 문구 추가 | ✅ | 출장 변경/취소 시 스케줄 조율 안내 (change_request + suggest) |
| 1.8-14 | 솔라피 버튼·메타데이터 총정리 | ✅ | 17개 템플릿 BC/WL/퀵버튼 + 메타데이터 #{type}_#{template}_#{name} |
| 1.8-15 | GAS 배포 @286 + GitHub Pages 배포 | ✅ | clasp push + git push |

### 알림톡 템플릿 17종 체계
```
매장방문 (5종): confirmed, cancelled, rescheduled, remind24, remind2
출장요청 (9종): request, suggest, field_confirmed, field_cancelled,
                field_rescheduled, field_remind_24h, field_remind_2h,
                field_delayed, change_request_received
톡상담   (2종): talk_received, talk_ready
복원수리 (1종): as_received
```

### 버튼·메타데이터 설계 (2026-03-03 최종 확정)
- WL 버튼: 6개 (일정확인/변경, 일정 선택하기, 복원수리 안내 확인)
  - URL에 `https://` 프로토콜 필수 → 솔라피 템플릿에서 `https://#{변수}` 형태
- BC 버튼: 12개 (1:1 문의하기)
  - ~~메타데이터 #{type}_#{template}_#{name}~~ → **메타데이터 제거** (한글 chatExtra 3080 에러)
  - 고객 식별: 해피톡 진입 시 사전 입력 폼(성함/연락처)으로 대체
- 퀵버튼: 4개 (리마인드용 1:1 문의하기)

### 잔여 작업
- [x] 솔라피 검수 제출 (17종 전체) → **승인 완료** ✅ 2026-03-03
- [x] BC 메타데이터 한글 불가 발견 → 전체 제거 결정 ✅ 2026-03-03
- [ ] BC 메타데이터 제거 후 솔라피 재검수 대기 (1~3 영업일)
- [x] Make Router 상담 17종 + 복원수리 6종 + 계약서 1종 분기 연결 ✅ 2026-03-12
- [ ] Make Router 리뷰 작성 요청 템플릿 연결
- [ ] 재검수 승인 후 E2E 테스트
- [ ] 검수 안정화 후 Make→솔라피 직접 호출 전환 (FLOW_change_request.md 참조)

---

## Phase 2: 대시보드 (허브 + 카테고리) ✅ 완료

**목적:** 운영자가 한눈에 현황 파악 — 허브(3초 전체 파악) + 카테고리별 전용 대시보드 3개

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 2-1 | 허브 대시보드 | ✅ | /dashboard → 주문/상담/복원수리 3개 HubCategoryCard (핵심 수치 + 클릭 이동) |
| 2-2 | 주문 전용 대시보드 | ✅ | /orders/dashboard → 파이프라인바 + 통계 4개 + 결제완료 UrgentList |
| 2-3 | 상담 전용 대시보드 | ✅ | /consultations/dashboard → 통계 4개 + 오늘 일정 타임라인 + 미확인 UrgentList |
| 2-4 | 복원수리 전용 대시보드 | ✅ | /repairs/dashboard → 6단계 파이프라인바 + 경과일 경고 + 수거접수/직접발송 UrgentList 2열 |
| 2-5 | 공유 컴포넌트 | ✅ | hub-category-card, pipeline-bar, urgent-list 3개 재사용 컴포넌트 |
| 2-6 | 내비게이션 업데이트 | ✅ | NAV_ITEMS matchPrefix 기반 active 판정, href → 카테고리 대시보드 |
| 2-7 | 캐시 무효화 통합 | ✅ | hub-stats / order-dashboard-stats / consultation-dashboard-stats / repair-dashboard-stats |

### 아키텍처
```
/dashboard (허브)
  ├─ 주문 카드 → /orders/dashboard (주문 전용)
  ├─ 상담 카드 → /consultations/dashboard (상담 전용)
  └─ 복원수리 카드 → /repairs/dashboard (복원수리 전용)

각 카테고리 대시보드 → "전체 목록" → /orders, /consultations, /repairs
```

### 주요 파일
- `hooks/use-dashboard-stats.ts` — 4개 통계 훅 (useHubStats, useOrderDashboardStats, useConsultationDashboardStats, useRepairDashboardStats)
- `components/dashboard/hub-category-card.tsx` — 허브 대형 클릭 카드
- `components/dashboard/pipeline-bar.tsx` — 수평 파이프라인 시각화
- `components/dashboard/urgent-list.tsx` — 긴급 건 리스트

---

## Phase 3: 아임웹 자동 동기화 📋 미착수

**목적:** 수동 동기화 버튼 없이, 주기적으로 신규 주문 자동 유입

| # | 작업 | 상태 | 설명 |
|---|------|------|------|
| 3-1 | Vercel Cron 설정 | 📋 | 5~10분 간격 자동 폴링 |
| 3-2 | 증분 동기화 최적화 | 📋 | 마지막 sync 시점 이후 변경분만 |
| 3-3 | 동기화 상태 대시보드 표시 | 📋 | 마지막 동기화 시간, 에러 로그 |

---

## Phase 4: 알림톡 연동 (Make + Solapi) 🔧 검수 승인 완료 / Make 연결 진행 중

**목적:** 주문 상태 변경 시 고객에게 자동 알림톡 발송

| # | 작업 | 상태 | 설명 |
|---|------|------|------|
| 4-1 | 일반 주문 알림 | 📋 | 아임웹 자동 알림톡 활용 (TMS 개입 불필요) |
| 4-2 | 상담 알림톡 17종 분기 | ✅ | Phase 1.8 구현 + 솔라피 검수 승인 (BC 메타데이터 제거 재검수 중) |
| 4-3 | 복원수리 알림톡 6종 | ✅ | Phase 7 구현 + as_cancelled 추가 (2026-03-05) + 솔라피 검수 승인 |
| 4-4 | Make → 솔라피 직접 호출 전환 | 📋 | 검수 안정화 후 전환 예정 (비용 절감) |

### 알림톡 역할 분리
- **일반 주문 (가위/주변제품)**: 아임웹 알림톡 (결제완료→발송→배송완료)
- **상담 (매장/출장/톡상담)**: Solapi 전담 — 17종 템플릿 (Phase 1.8)
- **복원수리 (계좌입금)**: Solapi 전담 6종 — 접수/비용안내/입금확인/출고/취소/만족도 (Phase 7)

---

## Phase 5: 오프라인 판매 ✅ 완료 (이카운트 → Phase ERP-A에서 제거)

**목적:** 오프라인 판매 기록 관리 (이카운트 연동은 Phase ERP-A에서 제거, TMS 단독 관리로 전환)

| # | 작업 | 상태 | 설명 |
|---|------|------|------|
| 5-1 | DB 마이그레이션 007 | ✅ | offline_sales + offline_sale_items |
| 5-2 | ~~이카운트 API 클라이언트~~ | ❌ | Phase ERP-A에서 삭제됨 (lib/ecount/ 제거) |
| 5-3 | API Routes + Hooks | ✅ | /api/sales CRUD + PATCH 취소/상태변경 + use-sales.ts |
| 5-4 | 판매관리 UI | ✅ | /sales 목록+모달 + /sales/new 입력(채널선택) + /sales/[id] 상세+취소 |
| 5-5 | NAV 추가 | ✅ | 사이드바+모바일 판매관리 메뉴 (Store 아이콘) |
| 5-6 | 판매 취소/상태변경 | ✅ | 취소(시리얼/재고/아임웹 역전) + 결제상태 변경(낙관적) (03-22) |
| 5-7 | 판매 채널 칩 | ✅ | offline/online/talk 칩 표시 + DB sale_channel 컬럼 (03-22) |
| 5-8 | 판매 상세 모달 | ✅ | 목록 클릭→모달 즉시 조회 + 인라인 액션 (03-22) |
| 5-9 | 탭/채널/기간 필터 | ✅ | 탭 바(전체/오늘/미수금/취소) + 채널 칩 + 기간 드롭다운 + 탭별 건수 (03-25) |
| 5-10 | PC 테이블 뷰 | ✅ | lg 이상 테이블, 모바일 카드 뷰 자동 분기 (03-25) |
| 5-11 | 임시 제품 직접 입력 | ✅ | 미등록 제품(빗/소모품) 이름+금액 직접 입력 (03-25) |
| 5-12 | 계약서 연결 | ✅ | offline_sales.contract_id + 신규 계약서 알림 배너 (03-25) |

> **참고:** 이카운트 연동 코드(lib/ecount/, /api/sales/ecount-sync 등)는 Phase ERP-A(2026-03-01)에서 완전 제거됨.
> 판매 VAT 자동계산, 시리얼 연결, 딜러 가격 자동적용은 Phase ERP-A/C에서 추가됨.

---

## Phase 6: 전자 계약서 ✅ 완료

**목적:** 매장 방문/상담 시 태블릿으로 전자 계약서 작성 + 서명 + 알림톡 발송

| # | 작업 | 상태 | 설명 |
|---|------|------|------|
| 6-1 | DB 마이그레이션 008 | ✅ | contracts + contract_items 테이블 |
| 6-2 | 서명 캔버스 | ✅ | 터치+마우스 지원, 고해상도 대응, base64 저장 |
| 6-3 | 계약서 작성/목록/상세 UI | ✅ | 제품 카드형 선택, 할부, 할인, 메모 |
| 6-4 | API Routes + Hooks | ✅ | /api/contracts CRUD + /api/contracts/notify |
| 6-5 | NAV 추가 | ✅ | 사이드바+모바일 계약서 메뉴 (FileSignature 아이콘) |
| 6-6 | **전자문서 리디자인** | ✅ | 종이 계약서 형태 UI (모바일/태블릿 전용) |
| 6-7 | 제품 선택 모달 | ✅ | 카테고리 탭 + 터치 타겟 리스트 |
| 6-8 | 구매자/판매자 서명 | ✅ | SignatureCanvas 2개 (buyer + seller) |
| 6-9 | DB 마이그레이션 017 | ✅ | delivery_method, deposit/balance, seller_signature, 매장정보 8컬럼 |
| 6-10 | 목록 탭 재편 | ✅ | 전체/신규계약/전환완료/취소 + 건수 뱃지 + PC 테이블뷰 (03-25) |
| 6-11 | 고객 필기 캔버스 | ✅ | HandwritingField 컴포넌트 — 성함/연락처/주소 S펜/터치 필기 (03-25) |
| 6-12 | 상담자 불러오기 | ✅ | 오늘 예약 고객 모달 → 정보 자동 기입 + consultation_id 연결 (03-25) |
| 6-13 | 이미지 자동 캡처 | ✅ | html2canvas → Supabase Storage 업로드 → image_url 저장 (03-25) |
| 6-14 | 상세 액션 재배치 | ✅ | 판매전환 / 판매전환+알림톡 버튼 병렬 + 이미지 열람 링크 (03-25) |
| 6-15 | DB 마이그레이션 031 | ✅ | consultation_id, handwriting 3컬럼, image_url (03-25) |

### 전자문서 UI 구조 (6-6→6-11 리뉴얼, 2026-03-25)
```
종이 계약서 형태 A4 레이아웃:
  헤더(MAMORU 구매계약서)
  → [상담자 불러오기] 버튼 + 매장명
  → 고객 필기 영역 (성함/연락처/주소 — HandwritingField 캔버스)
  → 필독/유의사항(법적 문구)
  → 결제방식(이체/카드/CMS+할부+선납/잔금)
  → 제품 테이블(모달 선택)
  → 날짜 + 서명 2개(구매자/판매자)
  → 입금 계좌 안내
  → 저장 시 html2canvas 캡처 → Supabase Storage 자동 업로드
```

### 잔여 작업
- [x] 계약서 이미지 자동 생성 (html2canvas + Supabase Storage) ✅ 03-25
- [x] 솔라피 계약서 알림톡 템플릿 등록 + 검수 승인 ✅ 2026-03-03
- [ ] 판매전환+알림톡 버튼 활성화 (계약서 이미지 링크 포함 템플릿 등록 필요)

- [x] 판매 상세→후기 요청 알림톡 + 리뷰 페이지 uid 연동 완료 ✅ 2026-04-09

> 통합 리뷰 시스템: `memory/REVIEW_SYSTEM_BRIEF.md` 참조 (별도 Phase)

---

## Phase 7: 복원수리 관리 🔧 코드 완료 / 운영 설정 대기

**목적:** 복원수리 접수→수거→검수→비용안내→입금→수리→출고 전 과정을 TMS에서 관리

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| 7-1 | Supabase 스키마 + types.ts | ✅ | repairs/repair_inspections/repair_history 3개 테이블 + RepairStatus ENUM |
| 7-2 | lib/repair/ 유틸리티 | ✅ | transitions (v3: 6단계 파이프라인, paid_at 분리), cost-calculator, inspection-text, sync |
| 7-3 | API Routes | ✅ | /api/repair/ (CRUD + 검수 + 출고 + 알림 + 동기화 + 리포트) 7개 라우트 |
| 7-4 | hooks/use-repairs.ts | ✅ | React Query 훅 8개 (목록/단건/상태변경/검수/출고/취소/알림/동기화) |
| 7-5 | 목록 UI + NAV | ✅ | /repairs 페이지, 상태 탭(6그룹: 신규접수/입고대기/작업중/출고/완료/취소), 경과일 표시 |
| 7-6 | 상세 UI + 검수 UI | ✅ | 2컬럼 레이아웃, 검수 체크리스트(7항목), 자동 문구, 비용 안내, 출고, 타임라인 |
| 7-7 | GAS 동기화 | ✅ | Code.gs doPost AS_CREATE → TMS sync webhook 추가, 상태변경 시 양방향 동기화 |
| 7-8 | 알림톡 6종 | ✅ | as_received/as_cost_notice/as_payment_confirmed/as_shipped/as_cancelled/as_satisfaction |
| 7-9 | Supabase SQL 실행 | 📋 | sql/phase7_repairs.sql → Supabase SQL Editor에서 실행 |
| 7-10 | GAS Script Properties 설정 | 📋 | TMS_REPAIR_SYNC_URL, TMS_BASE_URL, CRON_SECRET 설정 |
| 7-11 | 수리내역 페이지 TMS API 연동 | ✅ | page_as_report.html → GitHub Pages + TMS API(CORS) 구조로 구현 |
| 7-12 | E2E 테스트 | 📋 | 접수→검수→비용안내→입금→수리→출고 전체 플로우 검증 |
| 7-13 | PC 마스터-디테일 레이아웃 | ✅ | 좌측 목록 + 우측 상세 패널 (lg+), 모바일은 기존 페이지 이동 유지 |
| 7-14 | 대시보드 6단계 파이프라인 + paid_at 분리 | ✅ | 8→6단계, 입금확인 독립 플래그, 출고 2단계 분리 |
| 7-15 | R1 대시보드 탭 바 리모델 | ✅ | 고정 탭 바 6개(신규접수/수거접수필요/입고대기/진행중/출고대기/출고완료) + 인라인 액션 칩 + confirmed_at/packed_at |

### 상태 머신 (v3 — 6단계 파이프라인 + paid_at 독립)
```
파이프라인 6단계: 신규접수 → 입고대기 → 작업중 → 출고대기 → 출고완료 → 배송완료

방문수거: intake(신규접수) → pickup_scheduled(입고대기) → cost_notified(작업중)
         → repairing(작업중) → ready_to_ship(출고대기, 송장생성) → shipped(출고완료) → delivered → completed

직접발송: intake(신규접수) → cost_notified(작업중)
         → repairing(작업중) → ready_to_ship(출고대기, 송장생성) → shipped(출고완료) → delivered → completed

입금확인: paid_at 플래그 (파이프라인과 독립, 어느 상태에서든 입금확인 가능)
레거시 호환: picked_up, inspecting, ready_to_ship, payment_confirmed (기존 데이터 전이 가능)
```

### 운영 플로우
```
고객 접수 (page_form.html)
  → GAS doPost(AS_CREATE) → Sheets 저장 + Make 알림톡 + TMS sync
  → TMS /repairs 목록에 자동 반영

방문수거: [수거접수 완료] → [입고 & 비용안내](검수+비용+알림톡) → [작업 시작] → 송장생성(출고대기) → [출고완료](알림톡)
직접발송: [입고 & 비용안내](검수+비용+알림톡) → [작업 시작] → 송장생성(출고대기) → [출고완료](알림톡)
입금확인: 독립 버튼 (비용안내 이후 어느 단계에서든 가능, paid_at 설정 + 알림톡)
```

### 최근 완료 (2026-02-26) — R1~R4
- [x] **R7** 시리얼넘버/바코드 관리 (/products + /products/[id]/serials) + 단건/일괄 등록 + 상태 추적
- [x] **R7** DB 마이그레이션 009: product_serials + 판매/출고/계약서 연결
- [x] **R6** 전자 계약서 (/contracts 목록/작성/상세) + 서명 캔버스(터치+마우스) + 알림톡 발송 API
- [x] **R6** DB 마이그레이션 008: contracts + contract_items + 서명/PDF/상태 관리
- [x] **R5** 오프라인 판매 관리 (/sales 목록/입력/상세) + 이카운트 ERP API 클라이언트 + 판매전표 동기화
- [x] **R5** DB 마이그레이션 007: offline_sales + offline_sale_items + customers.ecount_customer_code
- [x] **R5** NAV에 판매관리 추가 (Store 아이콘, 사이드바+모바일)
- [x] **R4** 주문관리 결제칩(paid_at→결제완료/미납) + 배송메모 말줄임+호버 표시
- [x] **R3** 허브 대시보드: 판매 4단계+주/월금액, 상담 3수치+대응필요, 복원수리 3수치+주간요약
- [x] **R2** 상담 대시보드 6h기준 재설계 + 매장3탭/출장4탭 + 지도 PC고정+핀색상+양방향 + 달력 초록/보라
- [x] R1 대시보드 탭 바 리모델 (PipelineBar+UrgentList → 고정 탭 바 6개)
- [x] 탭별 인라인 액션: 신규접수[접수확인], 수거접수필요[수거접수완료], 입고대기[입고&비용안내], 진행중[내역서/입금/송장/포장 칩], 출고대기[출고완료], 출고완료[입금확인]
- [x] DB 마이그레이션 006: confirmed_at(접수확인), packed_at(포장완료) 컬럼 추가
- [x] 상태 머신: ready_to_ship 활성 승격 (repairing → ready_to_ship → shipped)
- [x] use-repair-tabs.ts: 탭별 Supabase 쿼리 + 카운트 통합 훅
- [x] RepairActionChips: 진행중 카드 전용 인라인 칩 바 (완료 시 초록 배경)

### 완료 (2026-02-25)
- [x] 대시보드 6단계 파이프라인 (신규접수/입고대기/작업중/출고대기/출고완료/배송완료)
- [x] payment_confirmed → paid_at 독립 플래그 분리 (Supabase 마이그레이션 완료)
- [x] 허브 카드: 접수처리/비용안내→수거접수 필요/작업중
- [x] 긴급리스트: 수거접수 필요(방문수거) + 직접발송 분리
- [x] 사이드바: 입금확인 독립 버튼 + 입금완료 배지 + 출고완료 버튼
- [x] 송장생성→ready_to_ship, 출고완료→shipped 2단계 분리
- [x] 목록 탭 재설계 (신규접수/입고대기/작업중/출고/완료/취소)
- [x] AS 폴더 구조 정리 — _gas/ 서브폴더 제거, consulting과 clasp 패턴 통일
- [x] deprecated 파일 삭제 (index.html 46KB, RepairReport.html 16KB)
- [x] .claspignore 생성 (page_*.html, iframe_*.html, icons/ 등 GAS 제외)

### 이전 완료 (2026-02-24)
- [x] 상태 머신 단순화 (12→9상태, 알림톡 수동발송 카드 제거)
- [x] 사이드바 통합 (비용+비용안내+액션+출고 → SidebarActionCard)
- [x] 접수정보 수량/주소 인라인 수정 + 비용 자동 재계산
- [x] 검수 폼 개선 (+ 버튼 추가 방식, 사진 촬영 UI)
- [x] 목록 날짜 표시 정리 (formatRelative 제거)
- [x] 수리내역 페이지 TMS API 연동 (page_as_report.html → GitHub Pages + TMS CORS API)

### 잔여 작업
- [x] GAS Script Properties 설정 (TMS_REPAIR_SYNC_URL 등) ✅ 2026-02-28
- [x] 솔라피 복원수리 템플릿 5종 등록/검수 승인 ✅ 2026-03-03
- [ ] Make Router에 복원수리 6종 분기 추가 (as_cancelled 포함)
- [ ] BC 메타데이터 제거 후 재검수 대기
- [ ] **주소 수정 시 다음 주소검색 API 연동** (롯데택배 송장 호환)
- [ ] 사진 업로드 Supabase Storage 연동 (버킷 생성 필요)
- [ ] 수리내역서 자동 생성 (Before/After 타임라인 웹카드)
- [ ] 사진 마킹 (photo-marker.tsx) — html2canvas 캡처 기능

---

## Phase ERP: 자체 ERP 전환 (이카운트 제거 + 도소매/매입/재고/회계) ✅ 완료 (2026-03-01)

**목적:** 이카운트 ERP 제거 → TMS 단일 시스템에서 판매·고객·제품·매입·재고·회계 직접 관리
**배경:** 이카운트 API 쓰기 전용(조회 불가), 사용 기능 5가지 미만, 간이사업자로 복잡한 ERP 불필요
**효과:** 이카운트 구독 비용 절약 + 이중 관리 해소 + 맞춤 UI/UX

### Phase A: 이카운트 제거 + 판매 강화 ✅
| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| A-1 | 이카운트 코드 제거 | ✅ | lib/ecount/ 6파일 삭제 + API/훅/페이지에서 이카운트 참조 제거 |
| A-2 | 판매 VAT 자동 계산 | ✅ | supply_amount/vat_amount 컬럼 + calcVAT() 유틸 + UI 표시 |
| A-3 | 판매 시 시리얼 연결 | ✅ | SerialPicker + 판매 생성 시 시리얼 status→sold 전환 |
| A-4 | 모바일 네비 개선 | ✅ | 5탭(대시보드/판매/상담/복원수리/더보기) + 바텀시트 |
- 커밋: `0beab17` — 22 파일, +462/-896줄
- DB: 011_remove_ecount.sql, 012_sales_vat.sql

### Phase B: 고객 관리 ✅
| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| B-1 | 고객 유형 시스템 | ✅ | customer_type (retail/online/dealer/supplier) + company_name + memo + outstanding_balance |
| B-2 | 고객 목록 페이지 | ✅ | /customers — 검색 + 유형 필터 + 판매액/미수금 표시 + 페이지네이션 |
| B-3 | 고객 상세 페이지 | ✅ | /customers/[id] — 인라인 편집 + 판매/계약/상담 관련내역 + 요약 카드 |
| B-4 | NAV 활성화 | ✅ | '고객' 메뉴 활성화 (Users 아이콘) |
- 커밋: `78fe991`
- DB: 013_customer_type.sql

### Phase C: 제품 관리 강화 ✅
| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| C-1 | 3단 가격 체계 | ✅ | price(소매) + price_dealer(도매) + price_purchase(매입가) |
| C-2 | 제품 등록/상세 | ✅ | /products/new + /products/[id] — 가격/매입처/아임웹매핑/바코드/설명 |
| C-3 | 딜러 가격 자동 적용 | ✅ | 판매 입력 시 customer_type=dealer → price_dealer 자동 적용 |
| C-4 | 아임웹 매핑 | ✅ | imweb_product_no 필드 (수동 매핑, API 제약으로 동기화 불가) |
- 커밋: `64b3b0b`
- DB: 014_product_prices.sql

### Phase D: 매입 관리 ✅
| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| D-1 | 발주 시스템 | ✅ | purchase_orders + items 테이블, PO-YYYYMMDD-NNN 자동 채번 |
| D-2 | 발주 목록/작성 | ✅ | /purchasing — 상태별 탭 + /purchasing/new 제품 선택형 작성 |
| D-3 | 발주 상세/액션 | ✅ | /purchasing/[id] — 발주확정/선납/입고/잔금/취소 상태 전환 |
| D-4 | 입고 시 재고 증가 | ✅ | received 전환 시 products.stock_quantity 자동 증가 |
| D-5 | NAV 추가 | ✅ | '매입관리' (Truck 아이콘) |
| D-6 | 매입품목 카탈로그 | ✅ | supplier_product_catalog 테이블 — 주문명/특징 + 제품 불러오기 (04-11) |
| D-7 | 부가세 3유형 | ✅ | 포함/별도/미적용 — calcVAT 확장 + purchase_orders.vat_type (04-11) |
| D-8 | 발주서 인쇄 | ✅ | POPrintModal — 주문품목+단가+수량+부가세 연동 인쇄 (04-11) |
- 커밋: `617957c`, `c17fb59`
- DB: 015_purchasing.sql, 061_supplier_catalog_vat.sql
- 상태 흐름: draft → ordered → deposit_paid → received → balance_paid | cancelled

### Phase D2: 납품관리 (B2B 딜러/아카데미) ✅ 신설 (2026-04-12)

**목적:** B2B 거래처(딜러/아카데미) 납품을 별도 모듈에서 관리 — 납품서 작성·확정·출고·정산

| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| D2-1 | DB 스키마 | ✅ | deliveries + delivery_items 테이블 (062 마이그레이션) |
| D2-2 | API Routes | ✅ | GET/POST /api/deliveries + GET/PATCH /api/deliveries/[id] |
| D2-3 | 상태 흐름 | ✅ | draft→confirmed(재고차감)→shipped(출고)→settled(정산) |
| D2-4 | 미수금 연동 | ✅ | 생성 시 증가, 정산/취소 시 차감 |
| D2-5 | UI 페이지 | ✅ | 마스터-디테일 + 납품서 작성 모달 + 상세 패널 |
| D2-6 | 납품서 인쇄 | ✅ | DLPrintModal (거래처+품목+부가세+증빙) |
| D2-7 | 부가세/증빙 | ✅ | 포함/별도/미적용 + 지출증빙/세금계산서/미적용 |
| D2-8 | 사이드바 | ✅ | 판매 그룹 내 '납품관리' 추가 |
- 채번: DL-YYYYMMDD-NNN
- 상태: draft → confirmed → shipped → settled | cancelled
| D2-9 | 복원수리 간편입력 | ✅ | 모달 모드 전환 — 거래처+수량+단가(8,000원)+결제 (04-13) |
| D2-10 | 대시보드 매출 연동 | ✅ | RPC 블록 내 deliveries 추가 쿼리 — 매출+B2B 수량 합산 (04-13) |
- 잔여: 출고완료 알림톡, 회계 보고서 연동

### Phase E: 재고 관리 강화 ✅
| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| E-1 | 창고 구분 | ✅ | product_serials.warehouse_zone (storage/display) |
| E-2 | 미입고 수량 | ✅ | v_pending_stock 뷰 (발주 진행 중 수량) |
| E-3 | 재고 대시보드 | ✅ | /inventory — 요약 카드(총재고/미입고/저재고/원가) + 카테고리 탭 + 저재고 필터 + 정렬 |
| E-4 | NAV 추가 | ✅ | '재고' (Boxes 아이콘) |
- 커밋: `bcd8cae`
- DB: 016_inventory.sql

### Phase F: 회계 리포트 ✅
| # | 작업 | 상태 | 구현 내용 |
|---|------|------|-----------|
| F-1 | 집계 API | ✅ | /api/reports/summary — 기간별 매출/매입/VAT/일별 추이 |
| F-2 | 엑셀 내보내기 | ✅ | /api/reports/export — xlsx 패키지로 매출/매입 엑셀 다운로드 |
| F-3 | 리포트 허브 | ✅ | /reports — 기간 프리셋 + 매출/매입/VAT 요약 카드 + 일별 바 차트 |
| F-4 | 거래내역서 | ✅ | /reports/transaction — 고객별 그룹핑 + @media print A4 + 서명란 |
| F-5 | NAV 추가 | ✅ | '회계' (BarChart3 아이콘) |
- 커밋: `dd9cedd` — 10 파일, +966줄
- DB 변경 없음 (기존 테이블 조회만)

### DB 마이그레이션 총괄 (Phase ERP)
| # | 파일 | 내용 |
|---|------|------|
| 011 | 011_remove_ecount.sql | ecount 기본값 제거 (컬럼 유지) |
| 012 | 012_sales_vat.sql | 판매 VAT 분리 (supply_amount, vat_amount) |
| 013 | 013_customer_type.sql | 고객 유형 + 메모 + 미수금 |
| 014 | 014_product_prices.sql | 도매가 + 매입가 + 매입처 + 아임웹 매핑 |
| 015 | 015_purchasing.sql | 발주 테이블 (purchase_orders + items) |
| 016 | 016_inventory.sql | 창고 구분 + 미입고 수량 뷰 |
| 017 | 017_contract_extend.sql | 계약서 전자문서 확장 (수령/선납/잔금/판매자서명/매장정보) |

### NAV 메뉴 (12개 — PC 사이드바)
```
대시보드 | 주문관리 | 상담관리 | 복원수리 | 판매관리 | 계약서 | 고객 | 제품 | 매입관리 | 재고 | 회계 | 설정
```

---

## Phase R (대규모 리모델) ✅ 코드 구현 완료 / 운영 연동 진행 중

> R1(복원수리) → R2(상담) → R3(허브) → R4(주문) → R5(오프라인판매+이카운트) → R6(계약서) → R7(시리얼)
> 코드+DB 구현: 2026-02-26 완료 (커밋 a29b989) | 58 파일, +4,744줄

| Phase | 내용 | 코드 | 운영 연동 |
|-------|------|------|-----------|
| R1 | 복원수리 대시보드 탭 바 리모델 (6탭+인라인 칩) | ✅ | 🔧 GAS 설정 + E2E |
| R2 | 상담관리 대시보드/탭 리모델 (지도+달력+양방향) | ✅ | ✅ |
| R3 | 허브 대시보드 리모델 (useHubStats+카드 확장) | ✅ | ✅ |
| R4 | 주문관리 온라인 강화 (결제칩, 메모란, 검색) | ✅ | ✅ |
| R5 | 오프라인 판매 + 취소/상태변경/채널칩/모달 (03-22 강화) | ✅ | ✅ 프로덕션 검증 완료 |
| R6 | 전자 계약서 (전자문서 리디자인 + 서명 2개 + 제품 모달) | ✅ | 🔧 PDF생성+솔라피 |
| R7 | 시리얼넘버/바코드 (단건/일괄 등록+상태 추적) | ✅ | 🔧 UI동작 확인 |

### 남은 운영 연동 작업 요약
1. ~~**이카운트**: 환경변수 → 세션 테스트~~ ✅ 프로덕션 검증 완료 (2026-02-28) → ❌ Phase ERP-A에서 제거
2. ~~**솔라피**: 복원수리 5종 + 계약서 1종 템플릿 등록 → 검수 제출~~ ✅ 23종 전체 검수 승인 (2026-03-03) → BC 메타데이터 제거 재검수 중
3. **Make**: 상담 17종 + 복원수리 6종 + 계약서 1종 Router 연결 완료 ✅ / 리뷰 템플릿 미연결
4. ~~**GAS**: Script Properties 설정~~ ✅ 완료 (2026-02-28)
5. **UI 확인**: Phase ERP 전체 모듈 동작 검증 (고객/제품/매입/재고/회계 포함)

---

## 문서: 모듈별 프로세스 흐름도 ✅ 완료 (2026-02-28)

각 모듈의 비즈니스 흐름 + 시스템 연동 + 구현 완료/미완료 현황 문서:

| 문서 | 모듈 | 상태 |
|------|------|------|
| `docs/TMS_FLOW_CONSULTATION.md` | 상담 (매장/출장/톡) | ✅ |
| `docs/TMS_FLOW_REPAIR.md` | 복원수리 (접수→출고) | ✅ |
| `docs/TMS_FLOW_SALES.md` | 판매 (오프라인+취소/채널/모달) | ✅ |
| `docs/TMS_FLOW_ORDERS.md` | 주문 (아임웹+롯데택배) | ✅ |

> 규칙: 모듈 작업 완료 후 해당 흐름도의 완료/미완료 섹션 반드시 업데이트

---

## 성능 최적화 (03-22) ✅ 완료

| # | 작업 | 상태 | 효과 |
|---|------|------|------|
| P-1 | Supabase 클라이언트 싱글턴화 | ✅ | 모든 hook 동일 인스턴스 재사용 |
| P-2 | QueryProvider gcTime 10분 + refetchOnWindowFocus 비활성 | ✅ | 탭 전환 시 불필요한 refetch 제거 |
| P-3 | 주문 탭 카운트 RPC 통합 (get_order_counts) | ✅ | 6개 쿼리 → 1개 |
| P-4 | 복원수리 탭 staleTime 20초 | ✅ | 10초마다 8개 → 20초마다 |
| P-5 | 대시보드 staleTime 30초 | ✅ | 15초 → 30초 (RPC 1회면 충분) |
| P-6 | 판매 API 재고/아임웹 순차→병렬 | ✅ | for loop 2N번 → Promise.all |
| P-7 | OrderRow + SaleRow React.memo | ✅ | 리스트 리렌더링 비용 절감 |
| P-8 | 송장 생성 낙관적 업데이트 | ✅ | 즉시 '배송중' 표시, 서버 후처리 |
| P-9 | DB: 027_perf_rpc (get_order_counts, get_repair_tab_counts) | ✅ | RPC 함수 2개 |

## UI 개선 (03-22) ✅ 완료

| # | 작업 | 상태 |
|---|------|------|
| U-1 | 판매 상세 모달 (목록 클릭→모달 조회+액션) | ✅ |
| U-2 | 판매 채널 칩 (오프라인/온라인/톡상담) | ✅ |
| U-3 | 제품 모바일 슬라이드 패널 (우측 시트) | ✅ |
| U-4 | SlidePanel 범용 컴포넌트 | ✅ |

## 시리얼 Lifecycle 강화 (03-24) ✅ 완료

| # | 작업 | 상태 | 설명 |
|---|------|------|------|
| SL-1 | warehouse_zone 3단계 | ✅ | raw(매입원본) / ready(마킹+포장) / display(진열) |
| SL-2 | 시리얼 조회 페이지 (/serials) | ✅ | 번호→제품/고객/판매/복원수리 즉시 확인 |
| SL-3 | zone PATCH API + 일괄 전환 | ✅ | 인라인/체크박스 일괄 zone 변경 |
| SL-4 | SerialPicker ready 필터 | ✅ | 판매 시 ready만 선택 가능 |
| SL-5 | 판매 취소 zone→ready 복원 | ✅ | 포장 상태 유지 |
| SL-6 | 재고 현황 3단계 zone 표시 | ✅ | raw/ready/display 수치 |
| SL-7 | NAV: CS 그룹 + 시리얼 조회 | ✅ | Search 아이콘 |
| SL-8 | DB: 028_serial_lifecycle | ✅ | storage→raw 마이그레이션 |

## 이카운트 데이터 이관 (03-24) ✅ 완료

| # | 작업 | 상태 | 설명 |
|---|------|------|------|
| EC-1 | CSV 파싱 강화 | ✅ | 큰따옴표 보호 + 에러 핸들링 |
| EC-2 | 분할 업로드 | ✅ | 30건씩 (Vercel 타임아웃 방지) |
| EC-3 | 일자 자동 정제 | ✅ | 2025-06-15-1 → 2025-06-15 |
| EC-4 | `모바일` 컬럼 인식 | ✅ | 기존 고객 전화번호 업데이트 |
| EC-5 | 고객 855명 이관 | ✅ | 딜러/아카데미/소매 자동 분류 |
| EC-6 | 판매 1,702건 이관 | ✅ | 금액 ×1000 SQL 보정 완료 |
| EC-7 | 판매 모달 최신 연락처 | ✅ | customer_id→customers.phone 실시간 |
| EC-8 | 연락처 하이픈 포맷 | ✅ | 010-1234-5678 형식 |

---

## 범례
- ✅ 완료
- 🔧 진행중 (코드 완료, 운영 설정/연동 대기)
- 📋 미착수
- ⏸️ 보류

---

## 작업 일지

### 2026-03-29 (TMS 전체 개선 Phase 0~3 + 발주 편집 + Vercel 서울 리전)
- **Phase 0**: ConfirmModal 공통 컴포넌트 + useIsLg 훅 + useEscapeKey 훅
- **Phase 1**: 확인 모달 24건 전체 적용 (복원수리9/상담3/판매2/주문3/매입5/제품1)
- **Phase 2**: PC 마스터-디테일 4개 모듈 (주문/고객/계약서/매입)
- **Phase 3**: 비활성 제품 토글 + 대시보드 매입·부자재 알림 + 기간 필터(주문/계약서/매입) + ESC 키
- **발주 편집**: draft 상태 품목 추가/수정/삭제 + 합계 재계산
- **시리얼 재임포트**: 이카운트 CSV 843건 연결 + 판매 상세 시리얼 순서 배분 + 조회 제품명 표시
- **제품 삭제**: 연관 데이터 체크 + 확인 모달 (시리얼/판매/계약서 있으면 거부)
- **B2B 판매**: 시리얼 없는 판매 시 raw_stock 보관창고 차감
- **Vercel 서울 리전(icn1)**: Washington D.C. → Seoul 전환 — API ~150ms 단축

### 2026-03-28 (복합결제 + SKU중복체크 + 시리얼재임포트 + 제품비활성화)
- **복합 결제 분리 입력**: '복합' 선택 시 카드/현금/이체 각 금액 입력 + 합계 검증
- **VAT 카드 기준 계산**: 카드 금액만 공급가액/부가세 분리 (현금/이체 비적용)
- **SKU 중복 체크**: 제품 등록 시 blur 실시간 검증 (API sku_check)
- **제품 비활성화 토글**: 상세 패널 EyeOff 아이콘으로 is_active 전환
- **시리얼 재임포트 API**: 이카운트 CSV → 기존 판매 건 매칭 → product_serials 생성
- DB 마이그레이션: 036_payment_detail, 037_serial_product_nullable

### 2026-03-27 (제품·창고·재고 전면 개편 + 재고 모델 재정립)
- **제품 페이지 전면 개편**: 카테고리 칩 탭 + 검색 + 컴팩트 3열 카드 + PC 3:2 모니터뷰
- **시리얼 빠른 등록 모달**: 시작번호 자동이어가기 + zone 선택(보관/준비/디스플레이) + raw_stock 차감
- **창고·재고 리모델**: 명칭 변경(재고→창고·재고) + 카테고리 칩 + 동적 요약 + 미사용 필터 + 도트 시각화 + 모니터뷰
- **zone 라벨 변경**: 매입원본→보관 / 판매준비→준비 / 진열→디스플레이 (6파일)
- **재고 모델 재정립**: `products.raw_stock` 컬럼 추가 — 보관창고=비시리얼 수량, 시리얼 생성=보관에서 차감
- **복원수리 비용 수동 편집**: 수리비/수거비 직접 입력 + 0원 무상처리 자동 입금완료
- **메뉴얼**: docs/MANUAL_PRODUCT.md (제품등록+시리얼+창고3단계+판매연동)
- DB 마이그레이션: 035_raw_stock
- 커밋: 8건

### 2026-03-26 (복원수리 IA 리모델)
- **page.tsx 리모델**: 상단 요약 카드 4개(신규/진행/미입금/3일경과) + isLg conditional rendering
- **repair-list.tsx 전면 리모델**: 7탭→6탭 파이프라인(useRepairTabData 활용) + 탭별 인라인 퀵 액션 + 특화 정보 칩 + 완료/취소 시각 구분
- 상담관리에서 검증된 패턴 적용: isLg 조건부 렌더링, 카운트 뱃지, 파이프라인 탭, 시각 구분

### 2026-03-06 (가이드 페이지 다크 모노크롬 + 수거비 개정)
- **page_guide.html 다크 모노크롬 전환**: 5탭→4탭(과정안내/소요시간/비용안내/포장방법), 접수방식 탭 제거→과정안내 서브탭 통합
- **page_as_guide.html 리디자인**: 알림톡 링크용 모바일 전용 4탭, 다크 모노크롬, CTA 제거(접수 완료 고객 대상)
- **비용안내 할인 표시**: 실제 수거+발송 비용 ~~8,000원~~ 취소선 + 할인 태그 (page_guide + page_as_guide 양쪽)
- **GAS 수거비 개정**: 방문수거 1자루 5,000→6,000원, 직접발송 1자루 3,000원 추가 (기존 무료→유료)
- GAS 배포: v283 (clasp push)

### 2026-03-05 (복원수리 알림톡 보강 + 접수폼 브랜드 전환)
- **Make 웹훅 URL 분기**: 상담(`MAKE_WEBHOOK_URL`) / 복원수리 상태변경(`MAKE_REPAIR_WEBHOOK_URL`) 분리
- **복원수리 알림톡 5종→6종**: `as_cancelled` (취소안내) 추가 — TMS 취소 시 자동 발송
- **courier 필드 추가**: 출고완료 알림톡 배송조회 버튼 활성화 (`롯데택배` 고정)
- **접수폼 브랜드 전환**: `page_form.html` Terracotta→Brand Guide v1.0 모노크롬 (CSS only)
- **접수폼 로고 교체**: 거북이 심볼 SVG (`vvvv12.svg`)
- 커밋: 79895a4, 24ba443, 6d8bebb, f8da8cd

### 2026-03-03 (솔라피 검수 승인 + BC 메타데이터 이슈 + Make Router)
- **솔라피 23종 전체 검수 승인**: 상담 17종 + 복원수리 5종 + 계약서 1종
- **BC 버튼 chatExtra 한글 불가 발견**: 3080 에러 — 한글 문자가 chatExtra에 포함되면 알림톡 발송 실패 (SMS 대체 발송)
  - 초기 시도: type_code (STORE/FIELD/TALK) 영문 코드 추가 → name(고객명)도 한글이므로 근본 해결 불가
  - 최종 결정: BC 버튼 메타데이터 전체 제거 + 해피톡 사전 입력 폼으로 고객 식별 대체
  - 솔라피 템플릿에서 메타데이터 삭제 후 재검수 요청 완료
- **TMS 일정변경(reschedule) 알림톡 버그 수정**:
  - change_request_link 누락 → extraData에 포함하도록 수정
  - template 자동 분기: store_visit → `rescheduled`, field_request → `field_rescheduled`
  - reschedule-modal.tsx에 consultationType, uniqueId props 추가
- **Make Router 연결 시작**: 상담 17종 분기 연결 진행 중 (수동 작업)
- GAS 배포: v287(type_code 추가), v288(전체 적용), v289(type_code 제거 — 최종)
- 커밋: 4c6e0d2, 7b94dce, 52dcd3d, cefdfc6, 729e037

### 2026-03-01 (계약서 전자문서 리디자인)
- **Phase 6-6~6-9**: 계약서 작성 페이지를 종이 계약서 형태 전자문서 UI로 완전 재작성
- 모바일/태블릿 전용 (PC는 "태블릿에서 작성하세요" 안내)
- 제품 선택 모달 (카테고리 탭 + 터치 타겟)
- 구매자/판매자 서명 캔버스 2개
- 법적 문구(유의사항, 반품/교환 조건) 하드코딩
- 수령방법(본사발송/직접수령), 결제방식(이체/카드/CMS), 선납/잔금 자동계산
- 매장 정보(직함/매장명/매장주소) 필드 추가
- 상세 페이지에 새 필드 표시 (매장정보, 수령방법, 선납/잔금, 판매자 서명)
- DB 마이그레이션 017: contracts 테이블 8개 컬럼 확장
- 커밋: `c7458ad` — 7 파일, +639/-213줄

### 2026-03-01 (간편진단 태블릿 PWA)
- `/diagnosis` 라우트 — 매장 태블릿 전용 가위 간편진단
- page_diag.html 13개 질문 + 조건부 분기를 Next.js PWA로 포팅
- 별도 manifest (mamoru-diagnosis), fullscreen + portrait
- 결과: 질문-답변 쌍 표 형태 표시 (상담사가 보고 안내)
- 파일: manifest-diagnosis.json, types.ts, data.ts, layout.tsx, page.tsx + middleware 수정

### 2026-03-01 (Phase ERP A~F: 자체 ERP 전환)
- **Phase A**: 이카운트 코드 완전 제거 (6파일 삭제) + 판매 VAT 자동계산 + 시리얼 연결 + 모바일 5탭 네비
- **Phase B**: 고객 관리 (목록/상세/유형필터 retail/online/dealer/supplier)
- **Phase C**: 제품 강화 (3단 가격: 소매/도매/매입가 + 아임웹 매핑 + 딜러 가격 자동적용)
- **Phase D**: 매입관리 (발주 작성/상세/상태전환 + 입고 시 재고 자동 증가)
- **Phase E**: 재고 현황 (창고 구분 storage/display + 미입고 수량 + 저재고 알림 + 원가 집계)
- **Phase F**: 회계 리포트 (매출/매입/VAT 집계 + 일별 추이 + 엑셀 내보내기 + 거래내역서 인쇄)
- DB 마이그레이션: 011~016 (6개)
- 커밋 6개: `0beab17` → `78fe991` → `64b3b0b` → `617957c` → `bcd8cae` → `dd9cedd`
- NAV 12개 완성: 대시보드/주문/상담/복원수리/판매/계약서/고객/제품/매입/재고/회계/설정

### 2026-02-26 (R1~R7 대규모 리모델)
- R1~R7 전체 코드 구현 + 빌드 통과
- DB 마이그레이션 006~009 Supabase SQL Editor에서 실행 완료
- 커밋: `a29b989` feat: TMS 대규모 리모델 R1~R7 전체 구현
- 58 파일 변경, +4,744줄 / -667줄
- 새 테이블: offline_sales, offline_sale_items, contracts, contract_items, product_serials
- 새 모듈: lib/ecount/ (이카운트 ERP API 클라이언트)
- 새 페이지: /sales(3), /contracts(3), /products(2)
- 새 컴포넌트: repair-tab-bar, repair-action-chips, 6개 탭, signature-canvas

### 2026-03-16 (Phase QOL: 체감 속도/UX 통합 개선)
- **P1 staleTime 명시**: 7개 목록 훅에 `staleTime: 30_000` 추가 (repairs/consultations/orders/sales/serials/contracts/customers)
- **P2 캐시 무효화 스코프 축소**: mutation 후 불필요한 hub-stats/dashboard-stats 즉시 무효화 제거 → staleTime(15s) 자연 갱신 위임
- **P3 페이지네이션 전체 건수**: orders/sales/contracts 페이지에 `총 N건` 텍스트 추가 (customers는 이미 있음)
- **P4 EmptyState 공통 컴포넌트**: `components/ui/empty-state.tsx` 생성, orders/sales/contracts/customers 4개 페이지 적용 (아이콘+텍스트 통일)
- **P5 repair-detail-panel 높이 유연화**: `h-[calc(100vh-130px)]` → `flex-1 overflow-y-auto` + 부모에 flex-col 적용
- **P6 Optimistic Update**: useUpdateRepairStatus/useUpdateRepairFields/useUpdateConsultationStatus에 onMutate 즉시 캐시 업데이트 + onError 롤백 + onSettled 재검증
- **P7 syncOrders 트랜잭션 보호**: order_items delete→insert를 upsert 패턴으로 전환 (데이터 유실 방지)
- **P11 모듈간 CTA 연결**: 상담(완료)→계약서 작성 CTA, 계약서(서명완료)→판매 등록 CTA 추가
- **P12 useHubStats RPC**: 14개 개별 쿼리를 1개 RPC 함수로 통합 (fallback 유지), DB 마이그레이션 018
- 파일: 15개 수정, 2개 신규 (empty-state.tsx, 018_hub_stats_rpc.sql)
