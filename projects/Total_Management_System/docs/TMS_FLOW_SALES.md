# 판매관리 프로세스 흐름도
> 최종 업데이트: 2026-02-28

---

## 1. 비즈니스 프로세스 흐름

### 오프라인 판매 등록
```
(매장 방문 고객)
  → (관리자) TMS /sales/new 판매 입력
    → 카테고리별 제품 선택 (BL/TH/LO/SL)
    → 수량/할인/결제방법 입력
    → 고객명 + 전화번호 입력
    → "판매 등록" → sale_number 자동 채번 (OS-YYYYMMDD-NNN)
  → /sales/[id] 상세 페이지
    → ecount_sync_status = 'pending'
    → "이카운트 동기화" 버튼 클릭
      → POST /api/sales/ecount-sync
      → 이카운트 SaveSale 호출 (판매전표 생성)
      → ecount_sync_status = 'synced' + ecount_slip_no 기록
```

### 결제 유형
```
card    — 카드 결제
cash    — 현금 결제
transfer — 계좌이체
mixed   — 복합 결제
```

### 결제 상태
```
paid    — 결제완료 (paid_amount = total_amount - discount_amount)
unpaid  — 미결제
partial — 부분결제
```

---

## 2. 시스템 연동 흐름

```
[관리자]
  │
  ▼
[TMS /sales/new] ──POST──→ [/api/sales]
                               │
                               ▼
                          [Supabase DB]
                          (offline_sales + offline_sale_items)
                               │
                               ▼
                          [/sales/[id]]
                               │
                        "이카운트 동기화" 버튼
                               │
                               ▼
                     [/api/sales/ecount-sync]
                               │
                    ┌──────────┴──────────┐
                    ▼                     ▼
             [이카운트 OAPI]        [Supabase 업데이트]
             SaveSale               ecount_sync_status
             (판매전표 생성)         ecount_slip_no
```

### 이카운트 연동 상세
| API | 엔드포인트 | 용도 |
|-----|------------|------|
| 로그인 | `/OAPI/V2/OAPILogin` | 세션 토큰 발급 (50분 캐시) |
| 판매입력 | `/OAPI/V2/Sale/SaveSale` | 판매전표 생성 |
| 거래처등록 | `/OAPI/V2/AccountBasic/SaveBasicCust` | 고객 → 이카운트 거래처 |
| 품목조회 | `/OAPI/V2/InventoryBasic/GetBasicProductsList` | 제품 목록 (786개) |
| 재고현황 | `/OAPI/V2/InventoryBalance/GetListInventoryBalanceStatus` | 재고 수량 (BASE_DATE 필수) |
| 품목등록 | `/OAPI/V2/InventoryBasic/SaveBasicProduct` | 신규 품목 등록 |

### 이카운트 환경변수
```
ECOUNT_COM_CODE=635735
ECOUNT_USER_ID=<user>
ECOUNT_API_CERT_KEY=<정식키, 1년>
ECOUNT_ZONE=AA
ECOUNT_TEST_MODE=false  (정식: oapi / 테스트: sboapi)
```

---

## 3. 구현 완료 ✅

### API Routes (2개)
| 엔드포인트 | 메서드 | 기능 |
|------------|--------|------|
| `/api/sales` | GET | 목록 조회 (검색/페이징) |
| `/api/sales` | POST | 판매 등록 (sale + items) |
| `/api/sales/ecount-sync` | POST | 이카운트 판매전표 동기화 |

### 페이지 (3개)
- `/sales` — 판매 목록 (결제상태/ERP상태 배지, 검색, 페이징)
- `/sales/new` — 판매 입력 (카드형 제품 선택 + 장바구니 + 결제)
- `/sales/[id]` — 판매 상세 (이카운트 동기화 버튼)

### 이카운트 API 클라이언트
- `lib/ecount/client.ts` — 인증/세션/도메인 분기 (oapi vs sboapi)
- `lib/ecount/sales.ts` — SaveSale (판매전표)
- `lib/ecount/customer.ts` — SaveBasicCust (거래처)
- `lib/ecount/inventory.ts` — 품목조회/등록/재고현황

### DB 테이블
- `offline_sales` — 판매 메인 (sale_number, 고객, 결제, ERP 동기화)
- `offline_sale_items` — 판매 상품 (제품, 수량, 단가)
- `customers.ecount_customer_code` — 고객별 이카운트 거래처 코드

### 프로덕션 검증 (2026-02-28 완료)
- 로그인 + 품목 786개 조회 + 재고 123개 조회 성공
- Vercel 환경변수 설정 완료
- 정식 인증키 (1년) 적용

---

## 4-A. 추가 구현 완료 ✅ (2026-03-01)

| 항목 | 구현 내용 |
|------|-----------|
| 고객 자동완성 | GET /api/customers/search — 이름/전화번호 ILIKE 검색, CustomerAutocomplete 공유 컴포넌트 |
| 고객 신규등록 + 이카운트 동시 | POST /api/customers — customers INSERT + saveCustomer(MM-NNN 채번) |
| 판매 저장 시 이카운트 자동 동기화 | POST /api/sales 내부에서 saveSale() 자동 호출, 실패 시 판매는 성공 유지 |
| 거래처 코드 자동 매핑 | 기존 고객 코드 없으면 판매 시 자동 채번+등록+UPDATE |
| 계약서 고객 자동완성 | contracts/new에도 CustomerAutocomplete 적용 (email/address 확장 필드) |

## 4-B. 미완료 ❌

| 항목 | 의존성 | 우선순위 |
|------|--------|----------|
| 온라인 주문 → 이카운트 매출전표 자동 연동 | 주문모듈 연동 | 중간 |
| 재고 연동 (판매 시 재고 자동 차감) | 재고 현황 API (검증완료) | 낮음 |
| 판매 수정/삭제 기능 | 없음 | 낮음 |

---

## 5. 핵심 파일 맵

### TMS API
| 파일 | 설명 |
|------|------|
| `app/src/app/api/sales/route.ts` | GET/POST 판매 |
| `app/src/app/api/sales/ecount-sync/route.ts` | 이카운트 동기화 |

### TMS UI
| 파일 | 설명 |
|------|------|
| `app/src/app/(dashboard)/sales/page.tsx` | 판매 목록 |
| `app/src/app/(dashboard)/sales/new/page.tsx` | 판매 입력 (카드형 UI) |
| `app/src/app/(dashboard)/sales/[id]/page.tsx` | 판매 상세 |

### TMS Lib
| 파일 | 설명 |
|------|------|
| `app/src/lib/ecount/client.ts` | 이카운트 OAPI 코어 (인증/세션/분기) |
| `app/src/lib/ecount/sales.ts` | SaveSale (판매전표) |
| `app/src/lib/ecount/customer.ts` | SaveBasicCust (거래처) |
| `app/src/lib/ecount/inventory.ts` | 품목조회/등록/재고현황 |
| `app/src/hooks/use-sales.ts` | React Query 훅 5개 |

### DB 마이그레이션
| 파일 | 설명 |
|------|------|
| `app/supabase/migrations/007_offline_sales.sql` | offline_sales + offline_sale_items |
