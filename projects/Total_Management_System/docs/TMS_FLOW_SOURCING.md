# 샘플 소싱 프로세스 흐름도

> 최종 업데이트: 2026-05-31 — 신규 모듈 `/sourcing` 구축 (Phase 1~3 + 업체정보·복제·매수·환율·바코드스캐너)

## 개요

**1688 신상 샘플을 들여와 "팔지 말지" 선별하는 도구.** 기존 매입관리(`/purchasing` — 아는 제품 재입고)와 별개. 사이드바: 상품·재고 → **샘플 소싱**.

### 도구 철학 (사장님 확정 2026-05-31)
- 역할 = **선별기**. 끝나는 지점 = **채택분 리스트 출력**. 등록(아임웹/TMS)은 사장님이 직접.
- ❌ **products INSERT·SKU 자동채번·판매가·발주수량 전부 없음** (의도적 제외). 등록은 아임웹 먼저 → TMS 동기화 유지.
- 소싱 = **회차/배치** (단일 매입처 헤더 없음). 출처는 품목별 정보로.
- **색상은 안 나눔** = 속성(아임웹 옵션). 나누는 기준 = "따로 *결정*해야 하나?" → YES=단위, NO=속성.

## 데이터 모델

```
sourcing_pos (소싱 회차)
  id, po_number(SRC-YYYYMMDD-NNN), supplier_name/url(미사용), order_date,
  exchange_rate(환율 — 라벨 한화가격에 적용), status(sourcing|done), memo(회차명)

sourcing_items (품목 = 샘플 1종 = 라벨 1장)
  id, po_id(FK CASCADE), sticker_no(SRC-...-001, QR 내용·UNIQUE),
  supplier_name(회사명), supplier_url(회사링크),   ← 업체 (099)
  vendor_url(품목링크), product_name, features_memo, unit_price(CNY), moq,
  inbound_photos(jsonb URL배열, 최대 5), inbound_memo,
  inspection_status(pending|matched|selected|rejected), selected_at, sort_order
```
- 마이그레이션: `098_sourcing.sql` + `099_sourcing_item_supplier.sql`
- RLS: 운영 표준(인증 전체 접근)

## 상태 머신 (inspection_status)

```
pending(대기) ──[폰/PC 입고매칭 완료]──> matched(매칭완료)
     │                                        │
     └──────────[선별]──────────┬─────────────┘
                                ▼
                  selected(채택) / rejected(탈락)
                                │
                          [되돌리기] → pending
```
- `selected` 전환 시 `selected_at` 자동 스탬프
- 탈락은 hard delete 아닌 상태 변경 (재발주 방지 기록)

## 전체 흐름

```
STEP 1 매입(품목 입력)
  소싱 회차 생성 → 품목마다: 회사명·회사링크 / 품목링크 / 품목명·단가·MOQ / 특징
  · 환율 상단 입력 → 단가(CNY) × 환율 = 한화(≈₩) 자동
  · 복제 버튼: 회사정보 유지 + 새 바코드 → 같은 업체 다른 품목 빠른 입력
        ↓
STEP 2 라벨 인쇄
  QR + 번호(#001) + 품목명 + 한화가격 · 품목별 매수(동일라벨 N장)
  → KM-106D 등 라벨프린터로 출력 → 샘플에 부착
        ↓ (중국에서 물건 도착)
STEP 3 입고매칭 (폰/PC 스캔)
  라벨 QR 스캔 → /sourcing/inbound/{itemId} → 실물 사진 촬영·업로드 + 메모 → 매칭완료
        ↓
STEP 3 선별
  사진 보고 채택/탈락 · 업체별 채택 현황("A 3/3 · C 2/5")으로 좋은 공장 선정
        ↓
STEP 4 선별 리스트 출력
  채택분(회사명+회사링크+품목명+품목링크+단가+특징+사진) → [복사] → 사장님이 직접 등록
```

## 라벨 (3렌더러 일관)

- **QR 내용**: `{origin}/sourcing/inbound/{itemId}` (스캔 → 입고매칭 직진입). `qrBaseUrl` prop.
- **표시**: QR + #번호 + 품목명 + ₩한화가격(단가×환율)
- **매수**: 품목별 `copies` (동일 라벨 N장 — 한 종류 제품 여러 실물에 부착)
- **출력 3경로** (`LabelPreview.tsx`):
  - **브라우저 인쇄** [라벨 N장 인쇄]: 새 창 `@page {size}mm` + window.print(), 매수만큼 반복. (KM-106D 등 범용)
  - **ZPL 저장** [.zpl]: 네이티브 canvas 1비트 `^GFA`(한글·QR 깨짐 방지, Tailwind v4 oklch가 html2canvas 깨뜨려서 직접 렌더), `^PQ{매수}`. **Zebra 전용**.
  - **CSV**: NiceLabel 데이터소스용 (copies 컬럼 포함).
- ⚠️ **싼 열전사+브라우저 정렬 핵심**: 윈도우 드라이버 용지(Stock) 크기를 라벨과 동일·여백0 등록 + 갭센서 캘리브레이션 필수. CSS만으론 구석에 찍힘. (KM-106D 검증됨)

## 스캔 흐름

| | 폰 | PC + 바코드 스캐너 |
|---|---|---|
| 스캔 | 카메라 QR | 스캐너(키보드 HID) → `ScanBox` URL에서 itemId 추출 |
| 진입 | `/sourcing/inbound/{id}` | 동일 |
| 사진 | 즉석 촬영(capture) | 파일 업로드 |
| 강점 | 도착 현장 검수 | 책상 빠른 호출 |
- 매칭 후에도 스캔하면 동일 품목 재진입 (QR=영구 식별자). 사진 있으면 표시, 없으면 추가.

## API / 페이지 맵

```
API   /api/sourcing            GET 목록 · POST 생성
      /api/sourcing/[id]       GET 상세 · PATCH 헤더 · DELETE
      /api/sourcing/[id]/items POST 품목추가(=복제)
      /api/sourcing/items/[itemId]        GET · PATCH(선별/입력/사진·메모) · DELETE
      /api/sourcing/items/[itemId]/photos POST 업로드 · DELETE (repair-photos 버킷 sourcing/ 경로)
훅    src/hooks/use-sourcing.ts (React Query)
페이지 /sourcing                  목록 + [새소싱] + ScanBox
      /sourcing/[id]             작업화면 (STEP 1~4)
      /sourcing/inbound/[itemId] 폰/PC 입고매칭
컴포넌트 components/sourcing/scan-box.tsx
        design-lab/_sections/sourcing-1688/LabelPreview.tsx (※ 향후 components/sourcing/로 이동 예정)
```

## 업체(공장) 선별 — IA

- 회사명 **free-text** + **복제 버튼**(정확 복사) → 회사명 그룹핑이 오타로 안 깨짐.
- STEP 3 **업체별 채택 현황** = `supplier_name` 그룹핑, 채택/전체/탈락 집계 → "이 공장에 좋은 게 많네 → 이 업체로".
- 향후 업체 수십 곳 + 업체 프로필 필요 시 → supplier 정규화(별도 테이블) 고려. 현재 규모엔 free-text가 가볍고 빠름.

## 의도적 제외 (혼동 방지)

| 항목 | 왜 없나 |
|------|---------|
| products 테이블 INSERT | 등록은 사장님 직접(아임웹 먼저→동기화). 유령 제품 방지 |
| SKU 자동채번·판매가 | 제품이 태어나는 아임웹에서 할 일 |
| 발주 수량/소계 | 현 단계 역할 = 수량 아닌 "선별" |
| 단일 매입처 헤더 | 샘플은 제품마다 다른 상점 → 품목별 회사정보로 |
