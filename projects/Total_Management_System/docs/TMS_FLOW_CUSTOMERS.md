# 고객 관리 프로세스 흐름도

> 최종 수정: 2026-04-30 (심야) — phone 기반 자동 매칭/생성 helper 신설 (`lib/customer/match-or-create.ts`)

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
