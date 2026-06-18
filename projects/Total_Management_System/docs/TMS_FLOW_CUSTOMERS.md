# 고객 관리 프로세스 흐름도

> 최종 수정: 2026-06-18 — **고객 병합(merge) 기능 추가** / 2026-04-30 B2B 카테고리 동적화 + 거래처별 단가(074) + 073 catalog + phone 매칭

---

## 2026-06-18 — 고객 병합(merge) + 중복 방지 정리

**배경**: 접수 자동매칭은 `phone_normalized`(숫자만, GENERATED 컬럼) 기준이라 하이픈 유무는 동일 고객으로 연동되지만, **다른/오타 전화나 전화 없이 수기 등록** 시 같은 사람이 중복 레코드로 생김(이름은 매칭에 안 씀).

- **병합 동작**: 고객 상세 패널 `[병합]` → 중복 고객 검색·다중선택 → 미리보기(판매·수리·미수) → 확인 → 실행.
- **이관 대상**: `merge_customers(p_primary, p_victims[])` RPC(단일 트랜잭션, 마이그 **106**)가 흡수 대상의 거래를 주 고객으로 이관.
  - 매출/청구 문서(offline_sales·deliveries·contracts·manual_invoices): `customer_id` + denormalized `customer_name/phone`를 **주 고객으로 통일**(매출·송장 한 이름).
  - 접수/주문(repairs·consultations·orders): `customer_id`만 이관(접수자 본명 기록 보존).
- **흡수 고객**: 삭제 X → `merged_into_id`/`merged_at` soft 보존, 목록·검색·자동완성에서 숨김(`.is('merged_into_id', null)`).
- **미수금**: RPC 후 `recalcOutstanding(primary)`(단일 출처) 재계산, 흡수본 0.
- 파일: API `app/api/customers/merge/route.ts`, 훅 `useMergeCustomers`, 모달 `components/customers/customer-merge-modal.tsx`, 버튼 `customer-detail-panel.tsx`.
- ⚠️ **마이그 106 SQL 먼저 실행 후 배포** (목록/검색이 `merged_into_id` 컬럼을 참조하므로).
- **중복 방지 (2026-06-18 보강)**: 수기 등록(`POST /api/customers`)도 `phone_normalized` 검사 → 같은 번호 기존 고객 있으면 **409 + existing 반환**. CustomerCreateModal이 경고 표시("이미 등록된 번호 — [기존고객]")하고 **[기존 고객 사용]/[그래도 새로 등록(force)]** 선택. (이전엔 수기 등록이 전화 매칭을 안 거쳐 같은 번호 중복 생성 — 곽경진 3중복 원인)
- **체인 평탄화 (마이그 107)**: A→B 병합 후 B→C 병합 시 A가 B(병합된 고객)를 가리키는 체인 방지 — 흡수 시 victim 을 가리키던 고객도 primary 로 재지정(A→C). ⚠️ 마이그 107 SQL 실행 필요.

## 2026-04-30 (심야 +3) — B2B 거래처 페이지에서 카테고리 직접 변경 가능

PartnerDetailPanel 편집 모드에 **"거래처 카테고리" 드롭다운** 추가. 이전엔 customer_type 변경하려면 `/customers/[id]` 고객 화면으로 이동해야 했음 (IA 위반).

- 위치: /거래처 → 거래처 선택 → ✏️ 편집 버튼 → 상단 "거래처 카테고리" 드롭다운
- 옵션: 동적 B2B 카테고리 (사장님 추가) + 매입처
- 저장 시 다른 카테고리 탭으로 자동 이동 (invalidateQueries)

상세: /suppliers/page.tsx PartnerDetailPanel `customer_type` form state

---

## 2026-04-30 (심야 +2) — B2B 카테고리 동적 관리 (074)

### 핵심
사장님이 설정에서 B2B 카테고리(딜러/아카데미/학교/공기관 등)를 자유롭게 추가/수정 가능.

### 위치: 설정 → 고객 관리 → "B2B 납품처 카테고리"
- 기본 (삭제 차단): dealer / academy
- 사장님 추가 가능 (예: school / public / hospital / 기타)

### 데이터 구조 (`system_settings.b2b.categories`)
```json
[
  { "key": "dealer", "label": "딜러", "icon": "Users", "display_order": 1, "is_active": true, "is_default": true },
  { "key": "school", "label": "학교", "icon": "School", "display_order": 3, "is_active": true, "is_default": false }
]
```

### 자동 반영
- /거래처 페이지 → 카테고리 탭 자동 노출
- 거래처 등록 시 `customer_type='school'` 등 자유 값 사용 가능 (DB가 TEXT라 가능)
- 색상: key 해시 기반 deterministic 자동

### 운영 매뉴얼 — 신규 카테고리 추가 후
1. ✅ B2B 카테고리 관리 (074로 추가) — UI 노출
2. ⚠️ 고객 유형 (`customer.types` 설정)에도 같은 key 추가 (사장님이)
3. ⚠️ 단가 그룹 관리 (`pricing.groups`)에도 등록 (옵션 — catalog.unit_price만 쓰면 생략 가능)

### catalog 단가 (074)
- `customer_product_catalog.unit_price` 컬럼 추가
- 거래처 상세 → 납품품목 탭에서 납품가 입력
- 판매 시 자동 적용 (가격 우선순위: catalog.unit_price → price_groups → product.price)

상세: `docs/TMS_FLOW_SALES.md` "2026-04-30 심야 +2" 섹션

---

## 2026-04-30 (심야 +1) — B2B 납품처 카탈로그 (073)

### 핵심
dealer/academy 거래처에 **납품품목 사전 등록** 가능. 각 납품처별 납품명(delivery_name) 별도 설정 → 판매 작성 시 자동 입력 → 송장 출력에 그 이름 박힘.

### 사장님 운영 흐름
1. /suppliers → 딜러 또는 아카데미 탭 → 거래처 선택 → "납품품목" 탭 클릭
2. "제품에서 불러오기" → 제품 선택 → catalog 항목 추가
3. 각 항목에 **납품명** + **특징** 입력 (저장)
4. 이후 그 거래처에 판매할 때 자동으로 납품명이 product_name으로 들어감

### 컴포넌트
- 신규 컴포넌트: `app/src/components/customers/customer-catalog-section.tsx`
- /suppliers 페이지 PartnerDetailPanel에서 customer_type='dealer'/'academy'일 때 노출

### Helper 인프라
- 마이그 073: `customer_product_catalog` 테이블
- API: `/api/customers/[id]/catalog`
- Hook: `useCustomerCatalog`, `useAddToCustomerCatalog`, `useUpdateCustomerCatalog`, `useRemoveFromCustomerCatalog`

상세: `docs/TMS_FLOW_SALES.md` "2026-04-30 심야 +1" 섹션

---

## 2026-04-30 (심야) — 고객 자동 매칭/생성 (phone 기반 SSOT)

### 핵심
신규 helper `app/src/lib/customer/match-or-create.ts`가 phone 기준 매칭/생성을 담당:
1. phone 정규화 (`replace(/\D/g, '')`) → phone_normalized 비교
2. 매칭됨 → 기존 customerId 반환 (`isNew: false`)
3. 매칭 X → 신규 INSERT → 신규 customerId (`isNew: true`)
4. phone 비어있음 → null (호출 측에서 customer_id NULL로 처리)

### 적용 라우트
- `/api/consultation/public/submit` — source='consultation'
- `/api/consultation/admin-create` — source='manual'
- `/api/repair/public/submit` — source='as'
- `/api/repair` POST — source='manual' (body.customer_id 없을 때만)

### Edge case
- 동명이인(같은 phone, 다른 이름): 가장 오래된 customer 매칭 (deterministic)
- 사후 분리: 사장님이 customer 상세에서 직접 분리/병합
- 이름 변경: phone 기준 매칭 → customers.name 자동 업데이트 X (사장님 직접)

### 백필 마이그 072
옛 NULL customer_id 데이터를 phone 매칭으로 1회성 채움. 사장님이 SQL Editor에서 1번 실행.

---

## 1. 고객 데이터 흐름

### 고객 생성 경로
| 경로 | 자동/수동 | customer_type |
|------|----------|--------------|
| 판매 입력 시 신규 고객 | 수동 (사장님 customer-autocomplete) | retail/dealer/academy |
| 상담 접수 (`public/submit`, `admin-create`) | **자동 (2026-04-30 구현됨)** | retail (default) |
| 복원수리 접수 (`public/submit`, POST) | **자동 (2026-04-30 구현됨)** | retail (default) |
| 아임웹 주문 동기화 | 자동 | online |
| CSV 임포트 | 수동 | 매핑 |
| 고객 직접 등록 | 수동 | 선택 |

### 고객 유형
| 유형 | 설명 | 가격 적용 |
|------|------|----------|
| retail | 일반 고객 | 소매가 |
| online | 아임웹 구매 고객 | 소매가 |
| dealer | 딜러 (B2B) | 딜러가 |
| academy | 아카데미 (B2B) | 아카데미가 |
| supplier | 매입처 | — |

---

## 2. 고객 상세 페이지 (`/customers/[id]`)

### 통합 타임라인
```
판매 (offline_sales) ──┐
계약서 (contracts) ────┤
상담 (consultations) ──┼──→ 시간순 통합 타임라인
복원수리 (repairs) ────┘
```

각 항목 클릭 → 해당 상세 페이지로 이동

### 거래 요약 카드 (4개)
- 총 판매건수 (취소 제외)
- 총 판매액 (취소 제외)
- 미수금 (outstanding_balance)
- 최근 거래일

### RFM 분류
| 등급 | 조건 | 뱃지 색상 |
|------|------|----------|
| VIP | 3회 이상 구매 or 50만원 이상 | 금색 |
| 일반 | VIP/휴면 아닌 경우 | 파랑 |
| 휴면 | 180일 이상 거래 없음 | 회색 |

---

## 3. 미수금 관리

### 미수금 발생
```
판매 등록 (payment_status = 'unpaid' or 'partial')
  → unpaid = (total - discount) - paid_amount
  → customer.outstanding_balance += unpaid
```

### 미수금 해소
```
결제상태 변경 (unpaid → paid)
  → customer.outstanding_balance -= unpaid
```

### 미수금 복원
```
판매 취소
  → customer.outstanding_balance -= unpaid (원래 미수금 차감)
```

### 에이징 (회계 리포트)
- 고객별 가장 오래된 미결제 판매일 기준
- 30/60/90일 구간 분류 → 색상 경고

---

## 4. 파일 매핑

| 파일 | 역할 |
|------|------|
| `app/api/customers/route.ts` | 목록 조회 + 신규 등록 |
| `app/api/customers/[id]/route.ts` | 상세 조회 (판매+계약+상담+수리 포함) + 수정 |
| `app/api/customers/search/route.ts` | 자동완성 검색 |
| `app/(dashboard)/customers/page.tsx` | 고객 목록 |
| `app/(dashboard)/customers/[id]/page.tsx` | 고객 상세 (타임라인+RFM+요약) |
| `hooks/use-customers.ts` | 훅 (CRUD + 검색) |

---

## 5. 연동 현황

### ✅ 구현
- [x] 판매 시 미수금 자동 반영
- [x] 판매 취소 시 미수금 복원
- [x] 결제상태 변경 시 미수금 조정
- [x] 통합 타임라인 (4개 모듈)
- [x] RFM 분류 (VIP/일반/휴면)
- [x] 최근 거래일 표시
- [x] 고객 자동완성 (판매 입력, 계약 작성)

### ❌ 미구현
- [ ] 상담/복원수리 접수 시 고객 자동 생성
- [ ] 복원수리 입금 시 미수금 연동
- [ ] 판매/수리/상담 상세에서 고객명 클릭 → 고객 상세 이동
