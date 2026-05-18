# MAMORU 상품 상세 페이지 카탈로그

> 시각적 빌더 시스템의 콘텐츠 자산 라이브러리. 한 번 정리하면 100자루 양산까지 재사용.

---

## 구조

```
catalog/
├── cards/              ← 카드 옵션 카탈로그 (시각 선택용)
│   ├── blade_edge.json
│   ├── blade_design.json
│   ├── handle_grip.json
│   ├── handle_camel.json
│   ├── grade.json
│   ├── thinning_teeth.json    (Phase 1B)
│   ├── thinning_pattern.json  (Phase 1B)
│   ├── thinning_reduction.json (Phase 1B)
│   └── ...
└── copy_pool/          ← 카피 라이브러리 (텍스트 선택용)
    ├── hero_subtitle_blunt.json
    ├── about_body_blunt.json
    ├── for_you_match_blunt.json
    ├── for_you_miss_blunt.json
    ├── handle_description_pool.json
    └── ...
```

---

## 카드 JSON 스키마

각 카드 파일은 하나의 *선택 그룹*을 정의합니다. 빌더 페이지가 이 JSON을 읽어 카드 갤러리 UI를 자동 생성합니다.

```json
{
  "card_type": "blade_edge",
  "applies_to": ["blunt"],
  "label_ko": "BLADE EDGE",
  "label_subtitle_ko": "날 선 형태",
  "select_mode": "single",
  "options": [
    {
      "id": "F",
      "name_ko": "직선형",
      "name_en": "FORCE",
      "svg_inline": "<svg ...></svg>",
      "description_ko": "거친 커트감 · 모발 밀림 적음"
    }
  ]
}
```

**필드 정의**:
- `card_type` — 고유 식별자 (snake_case). spec.json의 `selections.{card_type}` 키와 매칭
- `applies_to` — 가위 종류 배열. `["blunt"]` / `["thinning"]` / `["blunt","long"]` (공용 카드는 여러 종류)
- `label_ko` — 빌더 UI에 표시될 카드 그룹 제목
- `label_subtitle_ko` — 부제 (— 날 선 형태 같은)
- `select_mode` — `"single"` (단일 선택) | `"multi"` (다중 선택)
- `options[]` — 선택지 배열. 각 옵션:
  - `id` — spec.json에 저장될 값 (예: "F")
  - `name_ko` — 화면 표시명 (예: "직선형")
  - `name_en` — 영문 라벨 (예: "FORCE")
  - `svg_inline` — 카드 위 SVG (currentColor 사용 권장, stroke-width 0.75)
  - `description_ko` — 카드 본문 설명

---

## 카피 풀 JSON 스키마

각 카피 풀 파일은 한 영역(Hero subtitle, About 본문 등)의 *교체 가능한 카피 후보*들을 모은 라이브러리입니다.

```json
{
  "copy_type": "hero_subtitle",
  "applies_to": ["blunt"],
  "label_ko": "Hero subtitle (감성 카피)",
  "select_mode": "single",
  "options": [
    {
      "id": "blunt_consistency",
      "text": "한 손에 머무는 무게,\n일관된 끝매김.",
      "tone": "안정",
      "matches": ["5.5인치", "F+S"]
    }
  ]
}
```

**필드 정의**:
- `copy_type` — 영역 식별자. spec.json의 `copy_selections.{copy_type}` 키와 매칭
- `applies_to` — 가위 종류 배열
- `select_mode` — `"single"` (Hero subtitle 같은) | `"multi"` (For You 매치 같은)
- `options[]`:
  - `id` — spec에 저장될 값
  - `text` — 실제 표시될 카피 (멀티라인은 `\n`)
  - `tone` — 선택 시 사장님이 톤 직관적으로 알 수 있게 (선택)
  - `matches` — 어떤 모델 특성에 어울리는지 (선택, 빌더가 추천 정렬에 활용)

---

## spec.json 스키마 (모델별 자료)

각 모델의 빌더 선택 결과를 영구 보관하는 파일. `projects/products/specs/{모델}.json`.

```json
{
  "model": "A2-55FS",
  "type": "blunt",
  "category_label": "BLUNT 5.5 INCH",
  "size_inch": 5.5,
  "weight_g": 58,
  "price_grade": "A",

  "selections": {
    "blade_edge": "F",
    "blade_design": "S",
    "handle_grip": "semi_offset",
    "handle_camel": "camel"
  },

  "copy_selections": {
    "hero_subtitle": "blunt_consistency",
    "for_you_match": ["blunt_80pct", "short_blade", "light_touch"],
    "for_you_miss": ["slide_50pct", "big_hand", "beginner"]
  },

  "custom_fields": {
    "grip_thumb": "15 × 21",
    "grip_ring": "15 × 18",
    "handle_description": "표준 사이즈에 가까워 대부분의 손에 자연스럽게 맞습니다."
  },

  "lineup": ["A2-45FS", "A2-65FS", "A2-55FC"],
  "images_folder": "A2-55FS",
  "created_at": "2026-05-17",
  "updated_at": "2026-05-17"
}
```

---

## 카탈로그 운영 룰

1. **신규 카피 추가 시**: 해당 `copy_pool/*.json` 의 `options[]` 배열에 새 객체 추가. `id`는 모든 옵션 중 고유해야 함.
2. **신규 카드 옵션 추가 시**: 해당 `cards/*.json` 의 `options[]` 배열에 추가. SVG는 inline, stroke-width 0.75 통일.
3. **종류 신설 시** (예: 신규 가위 카테고리): 새 카드·카피 파일 생성 후 `applies_to` 배열에 신규 종류 추가.
4. **카탈로그는 git 영구 보관** — 변경 이력 자동 추적.
5. **카탈로그 수정 시 클로드 호출 가능** — "카탈로그에 {이런 카피} 추가" 한 마디로 추가.

---

## 종류별 카드 매핑 (Phase 1A에서 블런트만 우선)

| 종류 | 적용 카드 |
|---|---|
| **blunt** (블런트) | blade_edge · blade_design · handle_grip · handle_camel · grade |
| **thinning** (틴닝) | thinning_teeth · thinning_pattern · thinning_reduction · handle_grip · handle_camel · grade |
| **long** (장가위) | blade_edge · blade_length · balance · handle_grip · handle_camel · grade |
| **dry** (드라이) | TBD (사장님과 정의) |

빌더 페이지가 종류 선택 시 위 매핑 보고 해당 카드만 UI에 노출.
