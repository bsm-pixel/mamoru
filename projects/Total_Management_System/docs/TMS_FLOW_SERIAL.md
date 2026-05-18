# 시리얼 관리 프로세스 흐름도

> 최종 업데이트: 2026-05-18 — **Phase C audit log + Phase B 양방향 교환** 도입. 시리얼 무결성 다층 안전망 구축.

시리얼은 MAMORU 비즈니스의 **정품 진위 추적 + 평생 복원수리 보증 + 교환/반품 흐름**의 핵심 식별자. 사장님이 가장 민감한 데이터.

---

## 1. 전체 흐름

```
[등록] 매입·자동생성 ── status: in_stock, warehouse_zone: raw
   │
   ↓ (창고 이동)
[준비/디스플레이] warehouse_zone: ready / display
   │
   ↓ (판매 등록)
[판매완료] status: sold, offline_sale_id, sale_item_id, sold_to_*, previous_zone(raw/ready/display 저장)
   │
   ├── (정상 흐름) → 복원수리 이력 누적
   │
   ├── (취소/반품) status: in_stock 으로 복원, warehouse_zone 복원(previous_zone), sold_* 초기화
   │
   └── (잘못 등록 발견) Phase A 단방향 이전 OR Phase B 양방향 교환

[모든 변경] → Phase C 트리거가 product_serial_audit_log 에 자동 캡처
```

---

## 2. 시리얼 변경 호출처 (전수)

| # | 경로 | 변경 사항 |
|---|---|---|
| 1 | `POST /api/sales` | 판매 생성 시 시리얼 → sold |
| 2 | `POST /api/sales/[id]` (rebuild_sale) | 판매 수정 시 sale_item_id 재매칭 (큐 FIFO) |
| 3 | `PATCH /api/sales/[id]` (cancel/return) | 취소/반품 시 in_stock 복원 |
| 4 | `POST /api/import/sales` | 외부 import |
| 5 | `POST /api/serials` | 단건 등록 (raw → ready) |
| 6 | `POST /api/serials/batch` | 일괄 등록 + 자동 번호 |
| 7 | `POST /api/serials/move` | 창고 간 이동 (ready ↔ display, → raw) |
| 8 | `RPC swap_serials` (Phase B) | 양방향 동시 교환 (트랜잭션) |

**→ 모든 경로 변경은 Phase C 트리거가 자동 캡처** (코드 추가 누락 위험 0)

---

## 3. 시리얼 필드 의미

| 필드 | 의미 | 판매 시 채워짐 | 취소 시 |
|---|---|---|---|
| `id` | UUID PK | — | — |
| `serial_number` | 표시 번호 (UNIQUE) | — | — |
| `barcode` | 바코드 (UNIQUE) | — | — |
| `product_id` | 제품 FK (037부터 nullable) | — | — |
| `status` | in_stock / reserved / sold / returned / defective | `sold` | `in_stock` 복원 |
| `warehouse_zone` | 물리적 위치 (raw / ready / display) | 유지 | `previous_zone`으로 복원 |
| `previous_zone` | 판매 전 zone 백업 (033) | 현재 zone 저장 | NULL 초기화 |
| `offline_sale_id` | 판매 건 FK | 채움 | NULL |
| `sale_item_id` | 판매 항목 FK (049) | **반드시 채움** | NULL |
| `contract_id` | 계약서 FK | (계약서 연동 시) | (계약서 삭제 시 NULL) |
| `sold_via` | offline / online / contract | 채움 | NULL |
| `sold_at` | 판매 타임스탬프 | 채움 | NULL |
| `sold_to_name` | 고객명 스냅숏 | 채움 | NULL |
| `sold_to_phone` | 고객 연락처 스냅숏 | 채움 | NULL |
| `verify_token` | 정품확인 토큰 (034) | — | — |
| `created_by` | 등록자 | — | — |

---

## 4. Phase A — 단방향 이전 (이전 판매에서 가져오기)

### 시나리오
사장님이 새 판매를 등록하는데 입력한 시리얼이 이미 다른 판매에 등록되어 있음 (실수로 같은 번호 사용 or 기존 판매가 잘못 등록됨).

### 흐름
```
1. /sales/new 또는 판매 수정 화면에서 시리얼 입력
2. /api/serials/check-duplicate?serial=XXX 호출
3. 다른 판매에 존재하면 SerialConflictDialog 모달 노출:
   - 기존 판매번호 / 고객명 / 제품명 / 판매일 / 상태 표시
   - "이전 판매에서 분리하여 가져옵니다" 경고
4. 사장님 [이전 판매에서 가져오기] 클릭 → allow_serial_transfer = true 플래그
5. /api/sales 또는 /api/sales/[id] 호출 시 플래그 전달
6. 서버에서 기존 판매 시리얼 → 새 판매로 UPDATE (offline_sale_id, sale_item_id, sold_to_* 모두)
```

### 핵심 파일
- `components/sales/serial-conflict-dialog.tsx` — Promise 기반 모달
- `components/sales/serial-picker.tsx` — confirmIfDuplicate() 호출
- `app/api/serials/check-duplicate/route.ts` — 중복 검사
- `app/api/sales/route.ts:235-291` — 서버 안전망 (allow_serial_transfer 검증)

### 가드
- 서버 안전망: 동의 플래그 없이는 강탈 거부 (`SERIAL_DUPLICATE_TRANSFER_BLOCKED`)

---

## 5. Phase B — 양방향 동시 교환 (Swap) ⭐ 2026-05-18 신규

### 시나리오
판매 A에 시리얼 X 등록, 판매 B에 시리얼 Y 등록 → 사장님이 "둘이 바뀌어야 함" 발견 → 한 번의 트랜잭션으로 X↔Y 동시 스왑.

(Phase A 두 번으로도 가능하지만 중간 "허공" 상태가 생김. Phase B는 원자적.)

### 흐름
```
1. 시리얼 조회 페이지에서 시리얼 X 검색
2. 판매 정보 카드 우측 [⟷ 교환] 버튼 클릭 (sold 상태에만 노출)
3. SerialSwapDialog 모달:
   3-1. 상대 시리얼 Y 번호 입력 + 엔터
   3-2. useSerialLookup으로 Y 정보 가져옴
   3-3. 좌(X "현재") / 우(Y "상대") 카드 + 가운데 ⟷ 화살표
   3-4. 클라이언트 사전 가드 검증 (5겹 중 4겹)
        - 같은 시리얼? → 빨간 경고
        - 같은 판매? → 빨간 경고
        - 다른 제품? → 빨간 경고
        - 둘 다 sold? → 빨간 경고
        - 모두 통과 → 회색 안내 + [교환합니다] 버튼 활성
4. [교환합니다] 클릭 → POST /api/serials/swap → RPC 호출
5. RPC 가드 5겹 최후 검증 + SELECT FOR UPDATE 두 행 lock
6. 5개 필드 양방향 UPDATE (offline_sale_id, sale_item_id, sold_to_name, sold_to_phone, sold_via)
7. 성공 → 토스트 + 모달 닫힘 + 시리얼 정보 자동 갱신
8. Phase C 트리거가 두 UPDATE를 audit_log에 자동 캡처
```

### 가드 5겹 (RPC `swap_serials()` 내부)
| # | 에러 코드 | 메시지 |
|---|---|---|
| 1 | `SAME_SERIAL` | 같은 시리얼끼리는 교환할 수 없습니다 |
| 2 | `SERIAL_NOT_FOUND_A/B` | 시리얼을 찾을 수 없습니다 |
| 3 | `NOT_SOLD` | 두 시리얼 모두 판매완료 상태여야 교환할 수 있습니다 |
| 4a | `NO_SALE` | 판매에 연결된 시리얼만 교환할 수 있습니다 |
| 4b | `SAME_SALE` | 같은 판매 안의 시리얼끼리는 교환이 무의미합니다 |
| 5 | `PRODUCT_MISMATCH` | 같은 제품의 시리얼끼리만 교환할 수 있습니다 |

### 스왑 대상 / 유지 대상
- **스왑 (5개)**: offline_sale_id, sale_item_id, sold_to_name, sold_to_phone, sold_via
- **유지 (필드별 이유)**:
  - `status` (둘 다 sold 유지)
  - `warehouse_zone` (시리얼은 물리적으로 안 움직임)
  - `sold_at` (각 시리얼의 원본 판매 시각 그대로)
  - `previous_zone` (원본)
  - `product_id` (가드 5번에서 일치 검증 — 변경 없음)

### 핵심 파일
- `supabase/migrations/087_serial_swap_rpc.sql` — RPC 함수
- `app/api/serials/swap/route.ts` — API
- `hooks/use-serial-lookup.ts` — useSwapSerials() 훅
- `components/serials/serial-swap-dialog.tsx` — 모달
- `app/(dashboard)/serials/page.tsx` — 진입점

### 안전성
- 단일 RPC 트랜잭션 → 실패 시 자동 ROLLBACK
- SELECT FOR UPDATE → 동시 스왑 race 차단
- SECURITY DEFINER + authenticated EXECUTE 권한
- 클라이언트 + 서버 양쪽 검증

---

## 6. Phase C — 변경 이력 ledger ⭐ 2026-05-18 신규

### 목적
시리얼의 모든 변경(누가/언제/무엇)을 자동 캡처. 향후 사고 시 추적 + Phase B 검증 도구 + 위변조 방지.

### 구조
- 테이블: `product_serial_audit_log`
- FK 없음 — `serial_id`는 일반 UUID + `serial_number` 스냅숏 (시리얼 삭제돼도 이력 보존)
- 추적 6개 필드: status / warehouse_zone / sale_item_id / offline_sale_id / contract_id / product_id (각 old/new)
- 메타: `changed_by` (auth.uid()), `changed_at`
- 인덱스 4개 (serial_id, serial_number, changed_at, changed_by)

### 트리거
```sql
CREATE TRIGGER trg_product_serial_audit_log
AFTER INSERT OR UPDATE OR DELETE ON product_serials
FOR EACH ROW EXECUTE FUNCTION log_product_serial_change();
```

- `INSERT`: new 값만 기록
- `UPDATE`: 추적 6개 필드 중 하나라도 바뀌었을 때만 기록 (`updated_at` 만 바뀐 경우 skip)
- `DELETE`: old 값만 기록
- `SECURITY DEFINER` 로 `auth.uid()` 자동 캡처

### RLS — append-only 강제
```sql
ALTER TABLE product_serial_audit_log ENABLE ROW LEVEL SECURITY;
CREATE POLICY psal_select_authenticated ON product_serial_audit_log
  FOR SELECT TO authenticated USING (true);
-- INSERT/UPDATE/DELETE 정책 없음 = 모두 거부
-- 트리거는 SECURITY DEFINER 라 RLS 우회 → 트리거만 INSERT 가능
-- → 사장님조차 ledger 조작 불가, 위변조 방지
```

### UI — 시리얼 조회 페이지 "이동 이력" 카드
- 최신순 100건 + profiles JOIN(changed_by_name)
- 각 행: 액션 배지(생성/변경/삭제) + 변경 요약 + 시각 · 변경자
- describeChange() 헬퍼가 사람이 읽기 쉬운 요약 생성
  - INSERT: "재고 / 준비"
  - UPDATE: "상태: 재고 → 판매완료 / 위치: 준비 → 디스플레이 / 판매 연결됨"
  - DELETE: "마지막 상태: 판매완료"

### 핵심 파일
- `supabase/migrations/086_product_serial_audit_log.sql` — 테이블 + 트리거 + RLS
- `app/api/serials/audit/route.ts` — GET endpoint
- `hooks/use-serial-lookup.ts` — useSerialAudit() 훅 + SerialAuditLog 타입
- `app/(dashboard)/serials/page.tsx` — 이동 이력 카드

---

## 7. 무결성 진단 SQL (운영 중 정기 점검)

```sql
-- 1) 깨진 시리얼 카운트 (항상 0이어야 정상)
SELECT COUNT(*) FROM product_serials
WHERE offline_sale_id IS NOT NULL
  AND (sale_item_id IS NULL OR product_id IS NULL);

-- 2) Phase C 트리거 작동 확인
SELECT count(*) FROM product_serial_audit_log;
SELECT trigger_name FROM information_schema.triggers
WHERE event_object_table = 'product_serials';

-- 3) Phase B RPC 등록 확인
SELECT count(*) FROM pg_proc WHERE proname = 'swap_serials';
SELECT has_function_privilege('authenticated', 'swap_serials(uuid, uuid)', 'EXECUTE');

-- 4) Phase C RLS 정책 확인 (SELECT 1개여야 함)
SELECT cmd, count(*) FROM pg_policies
WHERE tablename = 'product_serial_audit_log' GROUP BY cmd;

-- 5) audit_log 최근 활동 (최근 100건)
SELECT changed_at, action, serial_number,
       old_status, new_status,
       old_warehouse_zone, new_warehouse_zone,
       changed_by
FROM product_serial_audit_log
ORDER BY changed_at DESC
LIMIT 100;
```

---

## 8. 시리얼 수정 시 클로드 절대 룰 (memory/feedback_serial_integrity_strict.md)

1. `product_serials` 테이블 INSERT/UPDATE/DELETE 하는 모든 코드 경로 grep
2. 각 경로에서 `product_id` + `sale_item_id` 명시적으로 채우는지 확인
3. DELETE + INSERT 재구성 시 외래키(sale_item_id) 재매칭 — 큐 FIFO 패턴 필수
4. 변경 후 무결성 진단 SQL 0건 확인
5. Phase C ledger 자동 작동 — 추가 코드 0

**금지**: `filter(s => s.product_id === item.product_id)` 패턴 (다중 상품 시 같은 시리얼 반복 반환됨)

---

## 9. 마이그레이션 이력

| # | 파일 | 내용 |
|---|---|---|
| 009 | product_serials | 테이블 + updated_at 트리거 |
| 028 | serial_lifecycle | warehouse_zone 추가 (raw/ready/display) |
| 033 | serial_previous_zone | previous_zone 추가 (취소 시 복원용) |
| 034 | serial_verify_token | 정품확인 토큰 |
| 037 | serial_product_nullable | product_id nullable화 (임시 제품 허용) |
| 049 | serial_sale_item | sale_item_id 추가 |
| 082~085 | (사장님 SQL 직접 실행) | 무결성 응급 복구 — 93.92% → 0% |
| **086** | **product_serial_audit_log** | **Phase C — 이동 이력 ledger** |
| **087** | **serial_swap_rpc** | **Phase B — swap_serials RPC** |
