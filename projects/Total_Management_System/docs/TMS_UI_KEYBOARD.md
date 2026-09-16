# TMS 키보드 조작 표준 — ESC 닫기 · Enter 선택

> 2026-09-16 신설. 사장님 요청: *"검색해서 하단에 하나 뜨면 엔터시 바로입력되는 그런 기본적인 느낌적인것들… TMS 전역에서 그런부분 모두 적용해주고 또 모달뜨거나 그런거에서 ESC 누르면 모달창 닫히도록"*
>
> 관련: [TMS_UI_ACTION_PANEL.md](TMS_UI_ACTION_PANEL.md) · [TMS_FILE_GUIDE.md](TMS_FILE_GUIDE.md)

---

## 왜 문서로 남기는가 — 같은 버그가 반복된 이유

모달 배경 드래그 닫힘 버그(2026-09-15, 37곳)와 **원인이 같다.**

> 공용 유틸 없이 화면마다 따로 구현 → 한 곳을 고쳐도 나머지로 퍼지지 않는다.

엔터 선택은 원래 `sales/new` **한 곳에만** 있었다. 나머지 선택 화면은 전부 마우스 클릭이었다.
ESC 닫기는 `slide-panel.tsx` 한 곳에만 있었다.

→ 그래서 **공용 모듈 2개를 만들고**, 새 화면은 반드시 이걸 쓰도록 여기에 고정한다.

| 기능 | 공용 모듈 | 금지 |
|------|-----------|------|
| 모달 ESC 닫기 | `components/ui/esc-close.tsx` | 모달마다 `useEffect` 로 `keydown` 붙이기 |
| 검색 Enter 선택 | `hooks/use-enter-select.ts` | `onKeyDown` 에 `if (e.key === 'Enter')` 직접 쓰기 |
| 모달 배경 닫기 | `lib/ui/backdrop.ts` | 배경 `div` 에 `onClick={onClose}` 직접 쓰기 |

---

## 1. ESC 로 모달 닫기 — `<EscClose />`

### 쓰는 법 (한 줄)

```tsx
<div {...backdropClose(onClose)} className="fixed inset-0 z-50 …">
  <EscClose onClose={onClose} />
  <div onClick={(e) => e.stopPropagation()}> … 모달 본문 … </div>
</div>
```

배경 `div` **바로 안**에 넣는다. 렌더 결과는 `null` 이라 레이아웃에 영향이 없다.

### 설계 (왜 이렇게)

| 규칙 | 이유 |
|------|------|
| **모듈 전역 스택**으로 맨 위 모달만 닫는다 | 모달 위에 확인 모달이 뜬 상태에서 ESC 를 누르면 **둘 다** 닫혀버린다. 스택 최상단만 반응하게 해서 "확인 모달만 닫히고 원래 모달은 남는" 정상 동작을 만든다 |
| 한글 조합 중(`isComposing`) ESC 무시 | IME 조합 취소용 ESC 가 모달 닫기로 새면, 오타 고치려다 작성 내용이 날아간다 |
| `useEffect` 를 화면마다 붙이지 않는다 | 화면마다 붙이면 스택이 없어 위 문제가 그대로 재발한다 |

### 적용 범위
`backdropClose(...)` 를 쓰는 배경 **37곳 / 34파일** 전부. 단 `components/ui/slide-panel.tsx` 는 자체 ESC 처리가 이미 있어 제외(중복 방지).

---

## 2. 검색 결과에서 Enter 로 선택 — `useEnterSelect`

### 쓰는 법

```tsx
const pick = useEnterSelect(filtered, (p) => { addProduct(p); setSearch(''); });

<input value={search} onChange={…} onKeyDown={pick.onKeyDown} />

{filtered.map((p, i) => (
  <button key={p.id}
    onClick={() => addProduct(p)}
    onMouseEnter={() => pick.setActiveIdx(i)}
    className={pick.isActive(i) ? 'bg-neutral-100' : 'hover:bg-neutral-50'}>
    …
  </button>
))}
```

> ⚠️ **훅이므로 이른 return 보다 위**, 그리고 콜백이 참조하는 함수 **선언보다 아래**에 둔다.
> (react-compiler 린트가 "선언 전 접근"을 에러로 잡는다. 보통 메인 `return (` 직전이 정답)

### 동작 규칙

| 키 | 동작 |
|----|------|
| `↓` | 활성 행 아래로 (지정이 없으면 **첫 행**) |
| `↑` | 활성 행 위로 |
| `Enter` | ① 활성 행이 있으면 그것 ② 없고 결과가 **딱 1개**면 그것 |
| `Enter` (결과 2개 이상 + 지정 없음) | **아무 일도 안 한다** — 오선택 방지 |
| `Enter` (한글 조합 중) | 무시 — 조합 확정용 Enter 가 선택으로 새면 안 된다 |

### 활성 행은 렌더 중에 무효화한다

검색어를 좁혀 결과가 줄면 이전 활성 인덱스가 **엉뚱한 행**을 가리킨다.
`useEffect` 로 되돌리면 한 박자 늦게(추가 렌더로) 반영돼 **그 사이 Enter 가 엉뚱한 행을 고를 수 있다.**

```ts
const [rawIdx, setActiveIdx] = useState(-1);
const activeIdx = rawIdx < items.length ? rawIdx : -1;   // 렌더 중 계산
```

### 적용된 화면 (8곳)

| 화면 | 파일 | 고르는 것 |
|------|------|-----------|
| 납품 생성 | `components/deliveries/create-delivery-modal.tsx` | 제품 |
| 매입처 선택 | `components/ui/supplier-select.tsx` | 매입처 |
| 거래처 취급제품 등록 | `components/customers/customer-catalog-section.tsx` | 제품 |
| 발주 생성 | `app/(dashboard)/purchasing/new/page.tsx` | 제품 |
| 납품 상세 — 제품 추가 | `components/deliveries/delivery-detail-panel.tsx` | 제품 |
| 매입 상세 — 제품 추가 | `components/purchasing/purchase-detail-panel.tsx` | 제품 |
| 고객 병합 | `components/customers/customer-merge-modal.tsx` | 합칠 고객 |
| 소싱 — 기존 제품 연결 | `app/(dashboard)/sourcing/[id]/_components/register-link-modals.tsx` | 제품 |

`app/(dashboard)/sales/new/page.tsx` 는 원래부터 자체 구현(정확 SKU 스캔 로직 포함)이라 그대로 둔다.

### 덤으로 고친 것 — DOM 직접 필터 제거

납품 상세 / 매입 상세의 "제품 추가" 검색은 React 상태를 안 쓰고
`querySelectorAll('[data-product-row]').style.display` 로 **DOM 을 직접 숨기고 있었다.**
→ React 가 모르는 상태라 Enter 선택을 붙일 수 없고, 리렌더가 나면 필터가 풀린다.
상태 기반(`addSearch` + `addableProducts`)으로 바꿨고 **SKU 검색도 이참에 추가**했다.

---

## 3. 적용하지 않은 곳 — 목록 필터

아래는 "검색 → 하나 고르기"가 아니라 **목록을 좁히는** 화면이라 Enter 선택 대상이 아니다.
(Enter 를 누를 대상 자체가 없다 — 좁힌 뒤 행을 클릭해 상세를 연다)

`contracts` · `customers` · `deliveries` · `inventory` · `orders` · `products` · `returns` · `sales` · `serials` · `suppliers` · 상담 목록 · 복원수리 목록 등 약 14개 목록 페이지.

필요해지면 "필터 결과가 1건이면 Enter 로 그 행 열기"를 별도 표준으로 추가한다.

---

## 4. 검증 방법 (재현 절차)

로그인 화면 때문에 실제 화면을 헤드리스로 못 여는 문제 → **임시 devpreview + 미들웨어 우회**로 푼다.

1. `src/middleware.ts` 백업 후 matcher 에 `devpreview` 한 단어 추가 → `'/((?!devpreview|api|…'`
2. `src/app/devpreview/page.tsx` 에 검증용 화면 작성
   - 실제 패널을 띄울 땐 `queryClient.setQueryData(['delivery','dl1'], {...})` 로 목 데이터 주입
   - 🚨 날짜 필드(`delivery_date`·`order_date`)를 빼면 `RangeError: Invalid time value` 로 크래시한다
3. `npm run dev` → `puppeteer-core` 로 **실제 키 입력**(`page.keyboard.press`) 테스트
   - 🚨 Chrome `--headless --screenshot` 은 이 PC에서 0x5 로 거부됨 → 반드시 `puppeteer-core`
   - 🚨 모달이 열려 있으면 뒤쪽 버튼 클릭이 오버레이에 막힌다 → `page.evaluate(() => el.click())` 로 히트테스트 우회
4. **끝나면 반드시 되돌린다**: `middleware.ts` 원복 · `devpreview` 삭제 · dev 서버 종료 · `.next/dev` 삭제

이번 검증 결과: ESC/Enter 9건 + 상세 패널 15건 = **24건 전부 통과**.
