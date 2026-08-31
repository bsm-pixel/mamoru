# TMS 자동 배송완료 추적 시스템 — 통합 흐름도

> **2026-05-25 구축**. 사장님 비전: "송장 있는 모든 흐름 자동 추적, 송장 없는 건 자연 제외, 새 채널 추가 시 5분 안에 적용 가능".
> **2026-07-12 확장 (109)**: **집하(수거) 자동 감지** — 기사님이 수거 스캔하면 자동 출고완료 + B2C 출고 알림톡.
> **새 테이블/채널에 자동 추적 기능 추가하려면 이 파일 그대로 보고 진행**.

---

## 🆕 집하 자동 감지 (109, 2026-07-12)

기존엔 **배송완료(41/45)만** 봤다. 집하 코드(`10`)는 원래부터 ALPS 응답에 왔지만 코드가 버리고 있었다.
이제 **기사님 수거 스캔 시점**에 자동으로 출고 처리하고, **B2C 고객에게만** 출고 알림톡을 보낸다.

```
송장 생성 (12 운송장등록 — 아직 우리 창고)
   ↓  ⏳ 기사님 방문 대기
기사님 수거 스캔 → ALPS godsStatCd '10' (집하)
   ↓  cron (1시간마다)
[3-A] offline_sales   shipped_at 자동 기록 + shipped_source='alps_pickup'
        └→ B2C 만 sales_shipped 알림톡 → shipped_notified_at 기록
[3-C] returns(교환출고) exchange_out_invoice_number 집하 → exchange_out_notified_at CAS 선점  ← 136(2026-08-27)
        └→ B2C 만 sales_shipped 알림톡(품명=새제품+"(교환)") · 매장교환(송장없음)/B2B/배달완료 제외
[2-A] repairs         ready_to_ship → shipped
        └→ as_shipped 알림톡
   ↓
(기존 흐름) 배달완료 41/45 → delivered_at → 리뷰 알림톡(약속한 고객만)
```

### 판정 규칙 (🚨 절대 건드리지 말 것)

```ts
// alps-client.ts isPickedUpCode
if (code === '09') return false;         // 반품취소
return n === 10 || n >= 20;              // 10=집하 / 20↑=간선·배달
```
- **`>= 10` 으로 하면 안 된다.** `12`(운송장등록)는 **우리가 송장을 발급하는 순간** 찍히는 코드다.
  이걸 집하로 오판하면 **송장만 만든 전 건이 즉시 출고 처리 + 알림톡 발송**된다. (구현 중 테스트로 발견)
- `10`만 정확히 매칭해도 안 된다 — 크론이 집하~간선 구간을 놓치면 영영 출고 처리가 안 된다. 그래서 `20↑`도 포함.

### 알림톡 안 나가는 경우 (의도된 skip)

| 상황 | 출고 처리 | 알림톡 |
|---|---|---|
| B2B (dealer/academy) | ✅ | ❌ 거래처는 발송 X |
| 이미 배달완료된 건 | ✅ | ❌ **이미 받은 고객에게 "출고했습니다" 금지** (집하 시점을 놓친 케이스) |
| 전화번호 없음 | ✅ | ❌ |
| `notifications.sales_shipped` 토글 OFF | ✅ | ❌ |

### 중복 발송 방지
- **조건부 CAS**: `.eq('id', x).is('shipped_at', null)` — 0행이면 이미 선점됨 → 알림톡 skip. 크론 중복 실행/수동 버튼 동시 클릭에도 정확히 1회.
- `shipped_notified_at` 기록 → 발송 여부를 화면에서 확인 가능.

### 적체 방지
`sale_date >= now() - 30일` 필터 — "송장만 뽑고 영영 안 나가는 건"이 limit 50을 영구 점유해 신건이 밀리는 것을 막는다.

---

## 🆕 B2B 납품도 같은 정의로 통일 (110, 2026-07-14)

**사장님 지적**: 납품은 송장만 만들었는데 화면에 "출고완료"가 떴다. 기사님은 오지도 않았는데.

**원인**: [`api/lotte/book/route.ts`](projects/Total_Management_System/app/src/app/api/lotte/book/route.ts) 가 송장 발급과 동시에 `deliveries.status='shipped'` 를 강제했다.
B2C 판매는 송장번호만 넣고 출고는 집하가 채우는데, **B2B만 "송장 발급 = 출고"로 다르게 구현**돼 있었다.

```
납품확정 → (송장 발급) 출고대기 → (기사님 수거=집하) 출고완료 → (인수자등록) 배송완료
           status=confirmed        status=shipped              delivered_at
           tracking_number          shipped_source='alps_pickup'
```

- **크론 [4-A]**: `status='confirmed' AND tracking_number NOT NULL` → 집하 감지 → `shipped` + `shipped_date` + `shipped_source`
- **B2B는 출고 알림톡을 보내지 않는다** (상태만 정확히)
- **뱃지 규칙 단일출처**: [`lib/deliveries/status.ts getDeliveryStatusChip()`](projects/Total_Management_System/app/src/lib/deliveries/status.ts) — 상세·목록·판매 통합목록이 전부 이걸 쓴다

### 🚨 status 를 늦춰도 안전한 이유 (전수 확인)

| 소비처 | 조건 | 영향 |
|---|---|---|
| 매출 집계 (hub_stats RPC 077/078/080/088, reports/summary) | `status IN ('confirmed','shipped','settled')` | **0** — confirmed 이미 포함 |
| 미수금 (`lib/outstanding.ts`) | `payment_status != 'paid'` (status 무관) | 0 |
| [결제완료 처리] (`update_payment`) | status 가드 없음 | 0 — 출고 전에도 결제 가능 |
| `settle` 액션 | UI에서 이미 제거된 죽은 경로 | 0 |

⚠️ `deliveries.status` 에는 **CHECK 제약도 enum 도 없다**(062) → 코드가 유일한 규약. 값 추가/변경 시 위 표를 반드시 다시 확인할 것.

### 🐛 함께 고친 버그
수동 [출고 완료] 액션이 `tracking_number: body.tracking_number || null` 이라, **송장이 이미 있는 건에 이 버튼을 누르면 송장번호가 지워졌다.** 전엔 "송장 생성 = 즉시 shipped" 라 누를 일이 없어 안 드러났던 버그 → `|| dl.tracking_number` 로 수정.

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
│ Vercel cron 자동 실행 (1시간마다)                    │
│ GET /api/cron/track-delivery                        │
│ Authorization: Bearer ${CRON_SECRET}                │
└──────────────────┬──────────────────────────────────┘
                   ▼
┌─────────────────────────────────────────────────────┐
│ 테이블별 분기 (현재 4종)                            │
│  ├ [1] orders        → queryStatus(invoice)        │
│  ├ [2-A] repairs 집하 → ready_to_ship → shipped   │
│  ├ [2] repairs       → queryTrackingStatus(invoice)│
│  ├ [3-A] sales 집하  → shipped_at + B2C 알림톡     │
│  ├ [3] offline_sales → queryTrackingStatus(invoice)│
│  └ [4] deliveries(B2B)→ queryTrackingStatus(track#)│
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

## 2. 현재 적용 상태 (2026-06-01 갱신)

| 채널 | 테이블 | 송장 컬럼 | 상태 처리 | delivered_at | 추적 함수 | 자동 추적 |
|------|--------|----------|----------|--------------|----------|----------|
| 아임웹 주문 | `orders` | `invoice_number` | `status: shipping → delivered` | ✅ | `queryStatus` | ✅ 작동 |
| 복원수리 | `repairs` | `invoice_number` | `status: shipped → delivered` | ✅ | `queryTrackingStatus` | ✅ 작동 |
| 합포장 출고 (repairs) | `repairs` (invoice 복사된 건) | ✅ | ✅ | ✅ | (위와 동일) | ✅ 작동 (자연 포함) |
| TMS 자체 판매(B2C) | `offline_sales` | `invoice_number` | status 미변경, `delivered_at`만 | ✅ (091) | `queryTrackingStatus` | ✅ 작동 (2026-05-25) |
| **거래처 납품(B2B)** | **`deliveries`** | **`tracking_number`** | **status 미변경(정산용), `delivered_at`만** | **✅ (101)** | **`queryTrackingStatus`** | **✅ 작동 (2026-06-01)** |

> ※ offline_sales·deliveries 는 status enum 에 'delivered' 가 없어 **status 는 안 바꾸고 `delivered_at` 만** 세팅(UI 표시는 delivered_at 기준). deliveries 는 송장 컬럼이 `invoice_number` 가 아니라 **`tracking_number`** 인 점만 다름.

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
| ALPS API rate limit | 🟡 낮음 | 50건/테이블 × 1시간 간격 → 충분 |
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

### A. offline_sales(B2C) 자동 추적 — ✅ 완료 (2026-05-25, 마이그 091)
### A-2. deliveries(B2B 거래처 납품) 자동 추적 — ✅ 완료 (2026-06-01, 마이그 101)
- 송장 컬럼 = `tracking_number`(기존), `delivered_at`만 추가. status 미변경.
- 고객/거래처 배송완료 표시 = 날짜 + **시간**(`M월 d일 HH:mm`).

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

---

## 11. 🚚 택배사 변경 시 참고 (롯데 → 타사, 예: CJ대한통운)

> **작성 2026-08-31.** 사장님 질문("사무실 옮기고 택배사 바꾸면 어떻게 되나?")에 대한 전환 가이드.
> **지금 할 일 아님** — 실제 택배사 변경이 결정되면 이 섹션을 그대로 보고 진행한다.

### 결론
**"설정에서 클릭 한 번"은 아니지만, 바꾸기 좋게 격리돼 있다.** 택배사 API 로직이 **`lib/lotte/` 한 폴더(3파일 ~760줄)** 에 모여 있고, 나머지 51개 파일은 여기서 함수만 갖다 쓴다. 새 택배사 API만 확보되면 그 폴더를 재작성하는 게 작업의 90%다. **진짜 전제는 기술이 아니라 새 택배사와의 API 연동 계약.**

### 어댑터 인터페이스 (이 함수 시그니처·반환형만 유지하면 바깥 안 깨짐)
| 함수 (`lib/lotte/`) | 하는 일 | 비고 |
|---|---|---|
| `bookShipment` (alps-client) | 송장 발행(출고) | 정방향 매장→고객 |
| `bookReturnPickup` (alps-client) | 반품 수거접수 | 롯데 `ustRtgSctCd='02'` |
| `queryTrackingStatus` / `queryStatus` | 배송 추적 | **이미 택배사 중립 형태** `{state, pickedUp, pickedUpAt, delivered}` 로 정규화 반환 → 크론·화면은 롯데 코드(10=집하 등)를 몰라도 됨 |
| `cancelShipment` / `cancelPickup` / `cancel` | 송장/수거 취소 | 롯데는 반품취소 API 미지원(수동) |
| `getNextInvoice` (alps-client) | 송장번호 채번 | 롯데식 체크디지트(base11 %7) — ⚠️아래 주의 |

### 바꿀 곳 체크리스트
**1. 핵심 — `lib/lotte/` 어댑터 재작성 (작업의 90%)**
새 택배사는 API 주소·인증·요청/응답 형식이 완전히 다름 → 이 폴더를 새 API에 맞춰 다시 씀. 함수 이름·반환형 유지가 핵심.

**2. 곁다리 (find & replace 수준)**
- **환경변수**: `LOTTE_TRACK_API_URL` / `LOTTE_CLIENT_KEY` / `LOTTE_JOBCUSTCD`(§5) → 새 택배사 계약값
- **하드코딩 문자열** `'롯데택배'`: 라벨·알림톡·DB `courier_name` 표기 (courier_name 컬럼은 레코드별로 저장돼 있어 **과거 롯데 건과 신규 타사 건이 공존 가능**)
- **ZPL 라벨 양식**(송장 출력, `MANUAL_LABEL`): 택배사마다 라벨 규격 다름
- **아임웹 택배사 코드**: 배송중 back-sync(`shipImwebOrder`)·`push-invoice` 시 아임웹에 넘기는 택배사 식별 코드

**3. 구조적으로 손이 가는 지점 ⚠️**
- **송장번호 채번 방식이 다를 수 있음**: 롯데는 **우리가 번호를 미리 계산**(`getNextInvoice`, 체크디지트)하는 구조. CJ 등 다수는 **발행 시 택배사가 번호를 내려줌** → 단순 교체가 아니라 "미리 채번 → 발행" 흐름을 "발행 응답에서 번호 수신"으로 바꿔야 할 수 있음.
- **추적 상태 코드 매핑**: `isPickedUpCode`(§9 판정규칙, `n===10||n>=20`)는 롯데 코드 전용 → 새 택배사 코드표로 다시 매핑하되, **반환형은 `{pickedUp, delivered, state}` 그대로 유지**해야 크론([1]~[4], [3-C] 집하)이 무수정으로 돈다.

### 거의 안 바뀌는 것 (재사용)
- 크론 구조([1]~[4], [3-A/C] 집하 감지)·조건부 CAS·적체방지 30일 필터
- 51개 호출부 대부분(어댑터 시그니처 유지 시)
- 모든 화면·알림톡 트리거·상태 흐름

### 전제 조건 (진짜 관건)
- 새 택배사와 **API 연동 계약 + 개발자 키 발급** (CJ대한통운 API 있음, 단 계약·심사 필요)
- ⚠️ 일부 소형 계약은 **추적만 주고 자동 송장발행 API를 안 줌** → 그러면 송장 자동발행 자동화 범위가 줄어듦(수동 송장 입력으로 폴백)

### 예상 규모
반나절~며칠짜리 개발(새 택배사 API 문서 난이도에 좌우). "클릭 토글"은 아님.

### 권장 (미래대비 리팩터, 선택)
전환 가능성이 실제로 있다면 지금 `lib/lotte/` 를 **`lib/couriers/`(공통 `CourierAdapter` 인터페이스 + `lotte.ts`/`cj.ts` + 설정값으로 고르는 팩토리)** 로 한 겹 감싸두면, 나중엔 **설정에서 택배사 선택**·심지어 멀티 택배사까지 가능. 지금은 급하지 않음 → 이 문서 참고로만 남김.
