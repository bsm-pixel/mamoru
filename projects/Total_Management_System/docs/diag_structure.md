# MAMORU 간편진단 종합 레퍼런스

> **대상 파일:** `projects/consulting/page_diag.html`
> **최종 업데이트:** 2026-02-10
> **문서 목적:** 질문 흐름, 아이콘, 수정 가이드를 한 곳에서 관리

---

## 1. 질문 흐름도

### 1-A. 전체 분기 플로차트

```
Q_STAGE ─────────────────────────────────────────── [필수]
  │
Q_TYPE (복수선택) ──────────────────────────────── [필수]
  │
  ├── BL 또는 LO 선택 ──┐
  │   └── (IN|DE) ──── Q_FEEL ──────────────────── [조건부]
  │
  ├── BL 선택 ──────────┐
  │   └── (DE) ──────── Q_STYLE ─────────────────── [조건부]
  │   └── (DE) ──────── Q_HABIT ─────────────────── [조건부]
  │
  ├── TH 선택 ──────── Q_TH_RATIO ──────────────── [조건부]
  │   └──────────────── Q_TH_WHY ────────────────── [조건부]
  │
  ├── LO 선택 ──────── Q_LO_USE ────────────────── [조건부]
  │
  ├── SL 선택 ──────── Q_SL_WHY ────────────────── [조건부]
  │   └── SL_SAME ──── Q_SL_SAME_WHY ───────────── [조건부]
  │
Q_GENDER ────────────────────────────────────────── [필수]
  │
Q_FING ──────────────────────────────────────────── [필수]
  │
  ▼ 결과 화면
```

### 1-B. 조건(condition) 상세표

| # | 질문 ID | label | condition | 설명 |
|---|---------|-------|-----------|------|
| 1 | `Q_STAGE` | 경력 | `null` | 항상 표시 |
| 2 | `Q_TYPE` | 가위 종류 | `null` | 항상 표시 (복수선택) |
| 3 | `Q_FEEL` | 커트 느낌 | `(IN\|DE) && (BL\|LO)` | 인턴/디자이너 + 블런트 또는 장가위 |
| 4 | `Q_STYLE` | 커트 스타일 | `DE && BL` | 디자이너 + 블런트 |
| 5 | `Q_HABIT` | 커트 습관 | `DE && BL` | 디자이너 + 블런트 |
| 6 | `Q_TH_RATIO` | 틴닝 감모량 | `TH 포함` | 틴닝 선택 시 |
| 7 | `Q_TH_WHY` | 틴닝 구매목적 | `TH 포함` | 틴닝 선택 시 |
| 8 | `Q_LO_USE` | 장가위 용도 | `LO 포함` | 장가위 선택 시 |
| 9 | `Q_SL_WHY` | 슬라이싱 구매동기 | `SL 포함` | 슬라이싱 선택 시 |
| 10 | `Q_SL_SAME_WHY` | 불만족 이유 | `Q_SL_WHY === 'SL_SAME'` | Q_SL_WHY에서 불만족 선택 시 |
| 11 | `Q_GENDER` | 성별 | `null` | 항상 표시 |
| 12 | `Q_FING` | 손가락 굵기 | `null` | 항상 표시 |

### 1-C. 실제 경로 예시

| 조합 | 표시되는 질문 (순서대로) | 문항 수 |
|------|------------------------|---------|
| CE + BL | Q_STAGE → Q_TYPE → Q_GENDER → Q_FING | 4 |
| IN + BL + TH | Q_STAGE → Q_TYPE → Q_FEEL → Q_TH_RATIO → Q_TH_WHY → Q_GENDER → Q_FING | 7 |
| DE + BL + SL | Q_STAGE → Q_TYPE → Q_FEEL → Q_STYLE → Q_HABIT → Q_SL_WHY → Q_GENDER → Q_FING | 8 |
| DE + BL + TH + LO + SL(불만족) | Q_STAGE → Q_TYPE → Q_FEEL → Q_STYLE → Q_HABIT → Q_TH_RATIO → Q_TH_WHY → Q_LO_USE → Q_SL_WHY → Q_SL_SAME_WHY → Q_GENDER → Q_FING | 12 (최대) |

### 1-D. 각 질문별 상세

#### Q1. 경력 (`Q_STAGE`)

| 필드 | 값 |
|------|-----|
| label | `경력` |
| question | `현재 단계가 어떻게 되시나요?` |
| sub | `가장 가까운 상황을 선택해주세요` |
| multiple | `false` |
| hasGif | `false` |
| condition | `null` |

| 옵션 ID | label | desc |
|---------|-------|------|
| `CE` | 자격증 준비 & 취득 | 아직 현장 경험이 없어요 |
| `IN` | 인턴 & 스탭 | 현장에서 배우며 연습해요 |
| `DE` | 디자이너 | 손님을 직접 담당해요 |

---

#### Q2. 가위 종류 (`Q_TYPE`)

| 필드 | 값 |
|------|-----|
| label | `가위 종류` |
| question | `어떤 종류의 가위가 필요하신가요?` |
| sub | `여러 개 선택 가능해요` |
| multiple | **`true`** |
| hasGif | `false` |
| condition | `null` |

| 옵션 ID | label | desc |
|---------|-------|------|
| `BL` | 블런트 | 사용비중이 가장 높은 메인 커트가위 |
| `TH` | 틴닝 | 양감 및 질감을 책임지는 가위 |
| `LO` | 장가위 | 면을 다듬는 가위의 기초 / 싱글링 |
| `SL` | 슬라이싱 | 질감 테크닉 |

---

#### Q3. 커트 느낌 (`Q_FEEL`)

| 필드 | 값 |
|------|-----|
| label | `커트 느낌` |
| question | `선호하는 커트 느낌이 어떻게 되시나요?` |
| sub | `가위를 쓸 때 원하는 느낌을 선택해주세요` |
| multiple | `false` |
| hasGif | **`true`** |
| condition | `(IN\|DE) && (BL\|LO)` |

| 옵션 ID | label | desc |
|---------|-------|------|
| `FEEL_SOFT` | 부드러운 느낌 | 폭신한 커트감 |
| `FEEL_POWER` | 힘있고 강한 느낌 | 시원시원한 커트감 |
| `FEEL_NONE` | 아직 잘 모르겠어요 | 어떤것을 원하는지 감이안와요 |

---

#### Q4. 커트 스타일 (`Q_STYLE`)

| 필드 | 값 |
|------|-----|
| label | `커트 스타일` |
| question | `주로 어떤 스타일로 커트하시나요?` |
| sub | *(비어있음)* |
| multiple | `false` |
| hasGif | **`true`** |
| condition | `DE && BL` |

| 옵션 ID | label | desc |
|---------|-------|------|
| `St_GO` | 직진성있게 커트해요 | 닫으면서 뒤로 빼지 않아요 |
| `St_BACK` | 뒤로 빼면서 커트해요 | 뒤로 빼듯 조곤조곤 |
| `St_NONE` | 딱 기준이 없어요 | 상황에 따라 달라요 |

---

#### Q5. 커트 습관 (`Q_HABIT`)

| 필드 | 값 |
|------|-----|
| label | `커트 습관` |
| question | `분무를 하여 커트하는 WET커트 비중이 어떻게 되세요?` |
| sub | *(비어있음)* |
| multiple | `false` |
| hasGif | **`true`** |
| condition | `DE && BL` |

| 옵션 ID | label | desc |
|---------|-------|------|
| `HAB_WET` | WET 커트 위주 | 항상 분무하고 커트해요 |
| `HAB_DRY` | DRY 커트 위주 | 마른 모발 커트가 많아요 |
| `HAB_NONE` | 잘 모르겠어요 | 그때그때 다른것 같아요 |

---

#### Q6. 틴닝 감모량 (`Q_TH_RATIO`)

| 필드 | 값 |
|------|-----|
| label | `틴닝 감모량` |
| question | `원하는 감모량이 어떻게 되시나요?` |
| sub | `숱이 빠지는 정도를 선택해주세요` |
| multiple | `false` |
| hasGif | `false` |
| condition | `TH 포함` |

| 옵션 ID | label | desc |
|---------|-------|------|
| `TH_25` | 25% (메인 틴닝) | 가장 범용적인 감모량 |
| `TH_15` | 15% (적은 감모) | 질감틴닝 / 숱이 적은고객 대상으로 사용 |
| `TH_35` | 35% (많은 감모) | 빠른 양감조절 |

---

#### Q7. 틴닝 구매목적 (`Q_TH_WHY`)

| 필드 | 값 |
|------|-----|
| label | `틴닝 구매목적` |
| question | `틴닝가위 구매 목적이 어떻게 되시나요?` |
| sub | *(비어있음)* |
| multiple | `false` |
| hasGif | `false` |
| condition | `TH 포함` |

| 옵션 ID | label | desc |
|---------|-------|------|
| `TH_WHY_NEW` | 새로운 감모량 추가 | 없는 감모량을 채우고 싶어요 |
| `TH_WHY_SAME` | 기존 틴닝 교체 | 쓰던 틴닝이 나쁘진 않았어요 |
| `TH_WHY_UP` | 업그레이드 | 같은 감모량 더 좋은 제품으로 |

---

#### Q8. 장가위 용도 (`Q_LO_USE`)

| 필드 | 값 |
|------|-----|
| label | `장가위 용도` |
| question | `장가위 주 사용 용도가 어떻게 될까요?` |
| sub | *(비어있음)* |
| multiple | `false` |
| hasGif | `false` |
| condition | `LO 포함` |

| 옵션 ID | label | desc |
|---------|-------|------|
| `LO_BL` | 블런트 겸용 | 커트가위처럼도 쓸 거예요 |
| `LO_SING` | 싱글링 전용 | 싱글링 작업 전용이에요 |

---

#### Q9. 슬라이싱 구매동기 (`Q_SL_WHY`)

| 필드 | 값 |
|------|-----|
| label | `슬라이싱 구매동기` |
| question | `슬라이싱 가위 구매 동기가 어떻게 되실까요?` |
| sub | *(비어있음)* |
| multiple | `false` |
| hasGif | `false` |
| condition | `SL 포함` |

| 옵션 ID | label | desc |
|---------|-------|------|
| `SL_NEW` | 첫 구매 | 슬라이싱 가위가 처음이에요 |
| `SL_SAME` | 기존 제품 불만족 | 쓰던 슬라이싱이 마음에 안 들어요 |

---

#### Q10. 슬라이싱 불만족 이유 (`Q_SL_SAME_WHY`)

| 필드 | 값 |
|------|-----|
| label | `불만족 이유` |
| question | `기존 슬라이싱 가위가 어떤 점이 불만족스러우셨나요?` |
| sub | *(비어있음)* |
| multiple | `false` |
| hasGif | `false` |
| condition | `Q_SL_WHY === 'SL_SAME'` |

| 옵션 ID | label | desc |
|---------|-------|------|
| `SL_SAME_WHY_UNCOM` | 밀리기만 하고 안 잘려요 | 커트 자체가 안 됨 |
| `SL_SAME_WHY_UNCOM1` | 너무 많이 잘려나가요 | 모발이 과하게 잘림 |
| `SL_SAME_WHY_HAND` | 핸들이 불편해요 | 사용감은 괜찮은데... |

---

#### Q11. 성별 (`Q_GENDER`)

| 필드 | 값 |
|------|-----|
| label | `성별` |
| question | `성별은 어떻게 되시나요?` |
| sub | `손에 맞는 가위 추천을 위해 필요해요` |
| multiple | `false` |
| hasGif | `false` |
| condition | `null` |

| 옵션 ID | label | desc |
|---------|-------|------|
| `FM` | 여성 | *(없음)* |
| `M` | 남성 | *(없음)* |

---

#### Q12. 손가락 굵기 (`Q_FING`)

| 필드 | 값 |
|------|-----|
| label | `손가락 굵기` |
| question | `가위를 사용할 손가락 굵기가 어떤 편일까요?` |
| sub | `핏에 맞는 가위 추천을 위해 필요해요` |
| multiple | `false` |
| hasGif | `false` |
| condition | `null` |

| 옵션 ID | label | desc |
|---------|-------|------|
| `FING_NORMAL` | 평범한 굵기 | 보통이에요 |
| `FING_THICK` | 두꺼운 편 | 주변보다 두꺼워요 |

---

## 2. 아이콘 총정리표

### 2-A. 외부 파일 아이콘 (`./icons/` 폴더)

| 파일명 | 사용처 | 설명 |
|--------|--------|------|
| `Level_2.svg` | Q_STAGE → `IN` | 인턴/스탭 레벨 아이콘 |
| `level_3.svg` | Q_STAGE → `DE` | 디자이너 레벨 아이콘 |
| `thinning.svg` | Q_TYPE → `TH` | 틴닝 가위 아이콘 |
| `Slide.svg` | Q_TYPE → `SL` | 슬라이싱 가위 아이콘 |

### 2-B. 인라인 SVG 아이콘

| 질문 | 옵션 ID | 모양 설명 |
|------|---------|-----------|
| Q_STAGE | `CE` | 자격증/카드 형태 (얼굴+텍스트 라인) |
| Q_TYPE | `BL` | 가위 (X자 교차) |
| Q_TYPE | `LO` | 장가위 (긴 날, 피벗 포인트) |
| Q_FEEL | `FEEL_SOFT` | 구름 |
| Q_FEEL | `FEEL_POWER` | 번개 (채움) |
| Q_FEEL | `FEEL_NONE` | 물음표 |
| Q_STYLE | `St_GO` | 직선 화살표 → |
| Q_STYLE | `St_BACK` | 위로 꺾인 화살표 (U턴) |
| Q_STYLE | `St_NONE` | 양방향 회전 화살표 |
| Q_HABIT | `HAB_WET` | 물방울 (채움) |
| Q_HABIT | `HAB_DRY` | 태양 (방사형 라인) |
| Q_HABIT | `HAB_NONE` | 저울/균형 |
| Q_TH_RATIO | `TH_25` | 4칸 중 1칸 채움 |
| Q_TH_RATIO | `TH_15` | 4칸 중 반칸 채움 |
| Q_TH_RATIO | `TH_35` | 4칸 중 1.5칸 채움 |
| Q_TH_WHY | `TH_WHY_NEW` | + 기호 |
| Q_TH_WHY | `TH_WHY_SAME` | 회전 화살표 |
| Q_TH_WHY | `TH_WHY_UP` | 위쪽 화살표 |
| Q_LO_USE | `LO_BL` | 가위 (X자 교차) |
| Q_LO_USE | `LO_SING` | 빗 + 가위 |
| Q_SL_WHY | `SL_NEW` | 별 (채움) |
| Q_SL_WHY | `SL_SAME` | 슬픈 얼굴 |
| Q_SL_SAME_WHY | `SL_SAME_WHY_UNCOM` | X 표시 가위 |
| Q_SL_SAME_WHY | `SL_SAME_WHY_UNCOM1` | 폭발 가위 |
| Q_SL_SAME_WHY | `SL_SAME_WHY_HAND` | 손 + 물결 |
| Q_GENDER | `FM` | 여성 실루엣 (긴 머리) |
| Q_GENDER | `M` | 남성 실루엣 (짧은 머리) |
| Q_FING | `FING_NORMAL` | 보통 손가락 |
| Q_FING | `FING_THICK` | 굵은 손가락 |

### 2-C. 미디어 상태 (`hasGif: true` 질문만)

| 질문 | hasGif | lottieUrl | gifUrl | 현재 상태 |
|------|--------|-----------|--------|-----------|
| Q_FEEL | `true` | `''` (비어있음) | `''` | 플레이스홀더 표시 |
| Q_STYLE | `true` | `''` (비어있음) | `''` | 플레이스홀더 표시 |
| Q_HABIT | `true` | `''` (비어있음) | `''` | 플레이스홀더 표시 |

> 위 3개 질문은 옵션마다 `gifUrl`/`lottieUrl` 필드가 존재하지만 현재 비어있어 플레이스홀더가 표시됨.

---

## 3. 결과 화면 데이터

### 3-A. 경력별 인사말 (`messagesData`)

| 키 | 경력 | 메시지 |
|----|------|--------|
| `MSG_CE` | 자격증 준비 | 새로운 꿈을 위한 멋진 도전! 응원하겠습니다! 마모루 대표는 필기만 따고 중도하차 했답니다 ^-^ |
| `MSG_IN` | 인턴/스탭 | 실무속에서 눈치보랴 움직이느라 힘드시죠? 미용가위만큼은 든든할 수 있도록 돕겠습니다! |
| `MSG_DE` | 디자이너 | 많은 손님의 스타일과 자신감에 도움을 주시느라 고생많으십니다 이젠 미용가위의 부담속에서 한결 가벼우시도록 저희가 노력하겠습니다 |

### 3-B. `labelMap` (결과 태그 표시용)

```js
const labelMap = {
  CE: '자격증 준비', IN: '인턴/스탭', DE: '디자이너',
  BL: '블런트', TH: '틴닝', LO: '장가위', SL: '슬라이싱',
  FEEL_SOFT: '부드러운 커트감', FEEL_POWER: '힘있고 강한 느낌', FEEL_NONE: '아직 모름',
  St_GO: '직진성 커트', St_BACK: '스트로크 커트', St_NONE: '기준 없음',
  HAB_WET: 'WET 커트', HAB_DRY: 'DRY 커트', HAB_NONE: '반반',
  TH_25: '25%', TH_15: '15%', TH_35: '35%',
  TH_WHY_NEW: '새 감모량', TH_WHY_SAME: '교체', TH_WHY_UP: '업그레이드',
  LO_BL: '블런트 겸용', LO_SING: '싱글링 전용',
  SL_NEW: '첫 구매', SL_SAME: '불만족 교체',
  SL_SAME_WHY_UNCOM: '안 잘림', SL_SAME_WHY_UNCOM1: '과다 절삭', SL_SAME_WHY_HAND: '핸들 불편',
  FM: '여성', M: '남성',
  FING_NORMAL: '보통', FING_THICK: '두꺼움'
};
```

### 3-C. 결과 태그 표시 키 (`tagKeys`)

```js
const tagKeys = new Set(['Q_STAGE', 'Q_TYPE', 'Q_GENDER', 'Q_FING']);
```

> 결과 화면 상단에 태그로 표시되는 답변. 위 4개 키에 해당하는 답변만 태그 칩으로 렌더링됨.

### 3-D. 복원수리 섹션 (`buildRestorationSection`)

복원수리 안내가 표시되는 조건:

| 조건 | 변수 | 설명 |
|------|------|------|
| 틴닝 복원 | `showTH` | `TH 선택 && Q_TH_WHY === 'TH_WHY_SAME'` (기존 틴닝 교체) |
| 슬라이싱 복원 | `showSL` | `SL 선택 && Q_SL_WHY === 'SL_SAME' && (SL_SAME_WHY_UNCOM \|\| SL_SAME_WHY_UNCOM1)` |

> `showTH`도 `showSL`도 아니면 → 복원수리 섹션 미표시

---

## 4. 수정 가이드

### 4-A. 질문 추가/수정

`questionsData` 배열에 객체를 추가한다.

```js
{
  id: 'Q_NEW',              // 고유 ID (필수)
  label: '라벨',            // 상단 칩 텍스트
  question: '질문 텍스트',  // \n으로 줄바꿈
  sub: '보조 설명',         // 비워도 됨
  multiple: false,          // true = 복수선택
  hasGif: false,            // true = GIF/Lottie 카드형 레이아웃
  condition: null,          // null = 항상 표시, 함수 = 조건부
  options: [
    { id: 'OPT1', label: '옵션1', desc: '설명', icon: '...' }
  ]
}
```

**condition 작성법:**

```js
// 항상 표시
condition: null

// 특정 경력일 때만
condition: (answers) => answers.Q_STAGE === 'DE'

// 특정 가위 종류 포함 시
condition: (answers) => (answers.Q_TYPE || []).includes('TH')

// 복합 조건
condition: (answers) => {
  const stage = answers.Q_STAGE;
  const types = answers.Q_TYPE || [];
  return (stage === 'IN' || stage === 'DE') && types.includes('BL');
}
```

### 4-B. 선택지 추가/수정

`options` 배열 내 객체 구조:

```js
// 기본형 (hasGif: false)
{ id: 'OPT_ID', label: '표시 텍스트', desc: '보조 설명', icon: '...' }

// GIF형 (hasGif: true) — 추가 필드
{ id: 'OPT_ID', label: '표시 텍스트', desc: '보조 설명', icon: '...',
  gifUrl: '',      // GIF 이미지 URL (비어있으면 플레이스홀더)
  lottieUrl: ''    // Lottie JSON URL (gifUrl보다 우선)
}
```

### 4-C. 아이콘 교체

`renderIcon()` 함수가 자동 분기:

| icon 값 | 처리 |
|---------|------|
| `<svg ...>...</svg>` | 인라인 SVG 그대로 삽입 |
| `<img ...>` | img 태그 그대로 삽입 |
| `./icons/파일.svg` (경로) | `<img src="경로" class="diag-custom-icon">` 으로 변환 |
| `✂️` (이모지/텍스트) | 텍스트 그대로 삽입 |

**교체 예시:**
```js
// 방법1: 인라인 SVG
icon: '<svg viewBox="0 0 24 24" fill="#C9A962">...</svg>'

// 방법2: 외부 SVG 파일
icon: './icons/new_icon.svg'

// 방법3: 이모지
icon: '✂️'
```

### 4-D. 결과 메시지 수정

| 수정 대상 | 위치 | 설명 |
|-----------|------|------|
| 경력별 인사말 | `messagesData` 객체 | `MSG_CE`, `MSG_IN`, `MSG_DE` 키의 값 수정 |
| 가위별 그룹 블록 | `buildPersonalizedMessage()` | 내부 맵(`feelMap`, `styleMap`, `habitMap`, `thWhyMap`, `loMap`, `slMap`, `slWhyMap`) 수정 |
| 복원수리 안내 | `buildRestorationSection()` | 조건 분기 및 `descText` 수정 |

**가위별 그룹 추가 시:**
1. `buildPersonalizedMessage()` 내에 새 `if (types.includes('XX'))` 블록 추가
2. 내부 맵 객체에 옵션 ID → 메시지 매핑 추가

### 4-E. labelMap 동기화

새 옵션을 추가하면 **반드시** `labelMap`에도 해당 ID → 한국어 라벨을 추가해야 한다.

```js
// 예: 새 옵션 MY_OPT 추가 시
const labelMap = {
  // ... 기존 ...
  MY_OPT: '내 옵션 라벨',  // ← 추가
};
```

결과 태그에 새 질문의 답변도 표시하려면 `tagKeys` Set을 수정:

```js
const tagKeys = new Set(['Q_STAGE', 'Q_TYPE', 'Q_GENDER', 'Q_FING', 'Q_NEW']); // ← 추가
```

### 4-F. 복원수리 섹션 수정

`buildRestorationSection(answers)` 내부:

```js
// 틴닝 복원 조건 — 다른 Q_TH_WHY 값도 포함하려면:
const showTH = types.includes('TH') && answers.Q_TH_WHY === 'TH_WHY_SAME';
//                                        ↑ 여기 조건 수정

// 슬라이싱 복원 조건 — 핸들 불편도 포함하려면:
const showSL = types.includes('SL') && answers.Q_SL_WHY === 'SL_SAME'
  && (answers.Q_SL_SAME_WHY === 'SL_SAME_WHY_UNCOM'
      || answers.Q_SL_SAME_WHY === 'SL_SAME_WHY_UNCOM1');
//    ↑ || answers.Q_SL_SAME_WHY === 'SL_SAME_WHY_HAND' 추가
```

### 4-G. 절대 수정 금지 영역

| 함수/영역 | 이유 |
|-----------|------|
| `diagGoToConsult()` | URL 파라미터 + postMessage 통신 — 상담접수 페이지 연동 |
| `diagGoToRecommend()` | URL 파라미터 + postMessage + localStorage — 추천 페이지 연동 |
| `diagAnswers` 구조 | 상담접수/추천 페이지에서 JSON.parse로 읽음 |
| `localStorage.setItem('MAMORU_DIAGNOSIS', ...)` | 추천 페이지 데이터 수신용 |
| iframe 높이 IIFE | `postMessage`로 아임웹 부모 프레임에 높이 전달 |
| `CONSULT_PAGE_URL`, `RECOMMEND_PAGE_URL` | 외부 페이지 URL 상수 |

---

## 5. 주요 함수 라인 참조

| 함수/변수 | 라인 | 설명 |
|-----------|------|------|
| `questionsData` | L1117 | 질문 데이터 배열 |
| `messagesData` | L1296 | 경력별 인사말 |
| `buildPersonalizedMessage()` | L1303 | 결과 메시지 HTML 생성 |
| `buildRestorationSection()` | L1373 | 복원수리 안내 HTML 생성 |
| `labelMap` | L1419 | 옵션 ID → 한국어 매핑 |
| `renderIcon()` | L1443 | 아이콘 렌더링 분기 |
| `diagShowResult()` | L1655 | 결과 화면 렌더링 |
| `tagKeys` | L1666 | 결과 태그 표시 키 Set |
| `diagGoToConsult()` | L1715 | 상담접수 페이지 이동 |
| `diagGoToRecommend()` | L1725 | 추천 페이지 이동 |
| iframe 높이 IIFE | L1742 | 아임웹 높이 동기화 |
