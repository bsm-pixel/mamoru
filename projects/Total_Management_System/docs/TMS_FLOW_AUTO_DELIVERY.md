# TMS 자동 배송완료 추적 시스템 — 통합 흐름도

> **2026-05-25 구축**. 사장님 비전: "송장 있는 모든 흐름 자동 추적, 송장 없는 건 자연 제외, 새 채널 추가 시 5분 안에 적용 가능".
> **새 테이블/채널에 자동 추적 기능 추가하려면 이 파일 그대로 보고 진행**.

---

## 핵심 원칙 (사장님 박제)

1. **송장 있는 운영 흐름 = 자동 추적 대상**
2. **송장 없는 흐름 (매장 직접 수령, 직접 전달 등) = 자동 추적 X, 사장님 수동 처리**
3. **재사용 패턴 — 새 테이블도 동일 구조로 5분 안에 추가 가능**
4. **회귀 위험 0 — 새 분기 격리, 기존 흐름 무영향**

---

## 1. 표준 흐름도

```
┌─────────────────────────────────────────────────────┐
│ 운영 시작: 송장 생성됨 (status = shipped/shipping)  │
│            invoice_number IS NOT NULL              │
└──────────────────┬──────────────────────────────────┘
                   ▼
┌─────────────────────────────────────────────────────┐
│ Vercel cron 자동 실행 (4시간마다)                    │
│ GET /api/cron/track-delivery                        │
│ Authorization: Bearer ${CRON_SECRET}                │
└──────────────────┬──────────────────────────────────┘
                   ▼
┌─────────────────────────────────────────────────────┐
│ 테이블별 분기 (현재 2종, 향후 확장)                  │
│  ├ [1] orders  → queryStatus(invoice)              │
│  ├ [2] repairs → queryTrackingStatus(invoice)      │
│  └ [3] offline_sales → (다음 구축 예정)             │
└──────────────────┬──────────────────────────────────┘
                   ▼
┌─────────────────────────────────────────────────────┐
│ ALPS 추적 API 호출                                  │
│ URL: ${LOTTE_TRACK_API_URL}?invNo=...&jobCustCd=... │
│ Headers: Authorization: IgtAK ${LOTTE_CLIENT_KEY}   │
└──────────────────┬──────────────────────────────────┘
                   ▼
┌─────────────────────────────────────────────────────┐
│ tracking[] 배열 godsStatCd 검사                      │
│  ├ '09' = CANCELLED (반품취소)                      │
│  ├ '41' = 배달완료 (기사 처리) ┐                    │
│  └ '45' = 인수자등록 (고객) ───┴→ DELIVERED         │
└──────────────────┬──────────────────────────────────┘
                   ▼
┌─────────────────────────────────────────────────────┐
│ DB 자동 전환                                        │
│  ├ status = 'delivered'                            │
│  ├ delivered_at = now()                            │
│  └ history 이력 박제 (선택, 테이블에 따라)          │
└──────────────────┬──────────────────────────────────┘
                   ▼
┌─────────────────────────────────────────────────────┐
│ TMS 카드 자동 갱신                                  │
│ 사장님 ALPS 직접 확인 0 → 운영 시간 절약            │
└─────────────────────────────────────────────────────┘
```

---

## 2. 현재 적용 상태 (2026-05-25)

| 채널 | 테이블 | 송장 컬럼 | 상태 컬럼 | delivered_at | 추적 함수 | 자동 추적 |
|------|--------|----------|----------|--------------|----------|----------|
| 아임웹 주문 | `orders` | `invoice_number` | `status: shipping → delivered` | ✅ | `queryStatus` | ✅ 작동 |
| 복원수리 | `repairs` | `invoice_number` | `status: shipped → delivered` | ✅ | `queryTrackingStatus` | ✅ 작동 |
| 합포장 출고 (repairs) | `repairs` (invoice 복사된 건) | ✅ | ✅ | ✅ | (위와 동일) | ✅ 작동 (자연 포함) |
| TMS 자체 판매 | `offline_sales` | `invoice_number` | ❌ **없음** | ❌ **없음** | (미적용) | ❌ **미구현** |

---

## 3. 자동 추적 미적용 케이스 (의도적 제외)

| 케이스 | 설명 | 처리 |
|--------|------|------|
| 매장 직접 수령 | 고객이 매장 방문해서 가위 가져감 | invoice_number NULL → 자동 추적 자연 제외. 사장님 수동 처리 |
| 직접 전달 (출장) | 사장님이 직접 배달 | 동일 |
| 합포장 출고 — 다른 주문 송장 사용 | repairs.invoice_number 에 판매건 송장 복사 | invoice_number 채워짐 → 자동 추적 대상 (이미 작동) |
| 매장 워크인 복원수리 | 향후 추가 예정 ([hidden-hugging-clarke.md](C:\Users\user\.claude\plans\hidden-hugging-clarke.md)) | 송장 없음 → 자동 추적 자연 제외 |

---

## 4. 새 테이블에 자동 추적 적용 — 5분 플레이북

### Step 1: 컬럼 확인 (1분)

필수 컬럼 3개:
```
invoice_number TEXT          (송장번호, NULL 허용)
status         TEXT          (또는 동등 상태값)
delivered_at   TIMESTAMPTZ   (배송완료 시각)
```

없으면 마이그레이션 추가:
```sql
ALTER TABLE <테이블>
  ADD COLUMN IF NOT EXISTS invoice_number TEXT,
  ADD COLUMN IF NOT EXISTS delivered_at   TIMESTAMPTZ;
```

인덱스 권장 (성능):
```sql
CREATE INDEX IF NOT EXISTS idx_<테이블>_tracking
  ON <테이블>(status, invoice_number)
  WHERE invoice_number IS NOT NULL;
```

### Step 2: cron 분기 추가 (2분)

[`api/cron/track-delivery/route.ts`](projects/Total_Management_System/app/src/app/api/cron/track-delivery/route.ts) 끝부분에 새 블록:

```typescript
// [N] <채널명> 추적 (<YYYY-MM-DD> 추가)
const { data: items } = await (supabase as any)
  .from('<테이블>')
  .select('id, <식별키>, invoice_number, ...')
  .eq('status', '<출고완료 상태>')
  .not('invoice_number', 'is', null)
  .limit(50);

let itemsDelivered = 0;
for (const item of items || []) {
  try {
    const result = await queryTrackingStatus(item.invoice_number);
    if (result.state === 'DELIVERED') {
      await (supabase as any)
        .from('<테이블>')
        .update({
          status: '<배송완료 상태>',
          delivered_at: new Date().toISOString(),
        })
        .eq('id', item.id);
      itemsDelivered++;
      console.log(`[track-delivery/<테이블>] ${item.<식별키>} → 배송완료`);
    }
  } catch (e) {
    console.error(`[track-delivery/<테이블>] ${item.<식별키>} 실패:`, e);
  }
}
```

응답에 추가:
```typescript
<channel_key>: { checked: items?.length || 0, delivered: itemsDelivered }
```

### Step 3: UI 분기 (선택, 2분)

해당 채널의 카드/사이드 패널에:
- `status='shipped'/'shipping'` 일 때 자동 추적 안내문
- 작은 fallback "수동 배송완료 처리" 텍스트 링크

예시: [`components/repairs/sidebar-action-card.tsx`](projects/Total_Management_System/app/src/components/repairs/sidebar-action-card.tsx) 참고

### Step 4: tsc 통과 + 커밋 + push

### Step 5: 매뉴얼 검증

```bash
curl -H "Authorization: Bearer $CRON_SECRET" \
  "https://app-eta-sandy-75.vercel.app/api/cron/track-delivery?debug=1"
```

응답에 새 채널 분기 결과 포함 확인.

---

## 5. ALPS 인프라 (공통)

### 환경변수 (Vercel Production)
| 변수 | 용도 |
|------|------|
| `LOTTE_TRACK_API_URL` | 추적 API endpoint (`cus/806/custmer-view-tracking`) |
| `LOTTE_CLIENT_KEY` | 인증 키 (215자) |
| `LOTTE_JOBCUSTCD` | 고객사 코드 (7자) |
| `CRON_SECRET` | cron 인증 |

### ALPS 응답 상태 코드 (2026-05-24 발견)
| 코드 | 의미 | 코드 처리 |
|------|------|----------|
| 02 | 출력 (운송장) | ACTIVE |
| 10 | 집하 | ACTIVE |
| 12 | 운송장등록 | ACTIVE |
| 20 | 셔틀발송/구간발송 | ACTIVE |
| 21 | 셔틀도착/구간도착 | ACTIVE |
| 40 | 배달전 | ACTIVE |
| **41** | **배달완료 (기사 처리)** | **DELIVERED** |
| **45** | **인수자등록 (고객 인증)** | **DELIVERED** ⭐ |
| 09 | 반품취소 | CANCELLED |

### 추적 함수 위치
- [`lib/lotte/alps-client.ts:queryTrackingStatus`](projects/Total_Management_System/app/src/lib/lotte/alps-client.ts) (repairs 용)
- [`lib/lotte/client.ts:queryStatus`](projects/Total_Management_System/app/src/lib/lotte/client.ts) (orders 용)

---

## 6. 회귀 위험 평가

| 영역 | 위험도 | 대응 |
|------|--------|------|
| 기존 cron 작동 | 🟢 0 | 새 분기는 try/catch 격리 |
| 기존 테이블 데이터 | 🟢 0 | 새 컬럼만 추가 (NULL 허용) |
| ALPS API rate limit | 🟡 낮음 | 50건/테이블 × 4시간 간격 → 충분 |
| 알림톡 중복 발송 | 🟢 0 | delivered_at 채워진 건은 다시 조회 안 됨 |
| 매장 직접 수령 흐름 | 🟢 0 | invoice_number NULL → 추적 대상 제외 |

---

## 7. 검증된 실제 동작 (사장님 직접 확인)

- 2026-05-25 01:01 — repairs 14건 일괄 자동 delivered 전환 (1.5초)
- 합포장 출고 (유지혜 5/14) 자동 매칭 + 자동 알림톡 ✅
- 강병현 송장 5/21 인수자등록 자동 감지 ✅
- 환경변수 갱신 후 즉시 모든 흐름 작동 ✅

---

## 8. 다음 작업 (사장님 비전 확장)

### A. offline_sales 자동 추적 추가 (사장님 결정 후)
- DB 마이그레이션: `offline_sales` 에 `delivered_at` 컬럼 추가 + status 컬럼 또는 동등 표현 검토
- 매장 직접 수령 vs 택배 발송 구분 패턴 결정
- 위 플레이북 그대로 적용

### B. 향후 추가 가능 채널
- contracts (전자 계약서) — 송장 발송 케이스 있다면
- 다른 외부 판매 채널

### C. 잔존 정리
- 진단 코드 (`?debug=1`) 제거 (정상 작동 확인 완료)
- [`lib/lotte/client.ts:24`](projects/Total_Management_System/app/src/lib/lotte/client.ts#L24) fallback URL `cus/714a` → `cus/806` 정정

---

## 9. 관련 문서

- [TMS_FLOW_REPAIR.md](TMS_FLOW_REPAIR.md) — 복원수리 흐름 전체
- [TMS_FLOW_ORDERS.md](TMS_FLOW_ORDERS.md) — 아임웹 주문 흐름
- [TMS_FLOW_SALES.md](TMS_FLOW_SALES.md) — 오프라인 판매 흐름
- 메모리: `reference_auto_delivery_tracking.md` — 동일 내용 메모리 박제
- 메모리: `reference_repair_merged_ship.md` — 합포장 출고 시스템

---

## 10. 사장님 강조 (2026-05-25)

> "이 흐름도 및 플로우를 꼭 관련 파일에 저장해두도록 하여라.
> 추후 배송 체크 기능 등을 여기에 넣자 하는경우에 이 파일을 그대로 확인해서 간편하게 삽입하거나 기능을 이용할 수 있도록 해야한다."

→ **이 파일이 그 SSOT (Single Source of Truth)**. 새 자동 추적 작업 시 무조건 이 파일부터 read.
