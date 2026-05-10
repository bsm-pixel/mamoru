# MAMORU Figma 마스터 명세서

> HTML(`v8_trendy.html`)의 디자인을 Figma로 미러링하기 위한 사양서.
> 사장님이 Figma에서 디자인 조정 → 클로드에게 알림 → HTML 갱신 (하이브리드 흐름).

---

## 핵심 원칙

- **HTML이 마스터**, Figma는 미러 (시각화/디자인 실험/매뉴얼/포트폴리오)
- 양산은 HTML 코드 paste 방식 (SEO + 텍스트 검색 ↑)
- Figma 변경 = 수동 동기화 (자동 반영 X)
- 두 도구 동기화 책임은 사장님 ↔ 클로드 협업

---

## Figma 파일 구조 (권장)

```
MAMORU 상세페이지 마스터.fig
├─ 📐 1. Foundations
│  ├─ Color Styles (10색)
│  ├─ Text Styles (8개)
│  └─ Spacing Tokens
├─ 🧩 2. Components (8섹션 마스터)
├─ 🎨 3. Master Page (CL1-4T 더미)
└─ 📦 4. Products (제품별 인스턴스)
```

---

## 1. Color Styles (Brand Guide v1.0)

Figma → Local Styles → Color에 다음 10개 등록:

| Style 이름 | HEX | 용도 |
|-----------|-----|------|
| `Void` | `#1A1A1A` | 메인 텍스트, CTA 다크 배경 |
| `Graphite` | `#2D2D2D` | 본문 진한 텍스트 |
| `Stone` | `#4A4A4A` | 보조 본문 텍스트 |
| `Warm Gray` | `#8A8580` | 캡션, 라벨, 보조 |
| `Mist` | `#B8B4AF` | 플레이스홀더 |
| `Sand` | `#D4D0CB` | 구분선, 경계선 |
| `Parchment` | `#EDEBE8` | 섹션 구분 배경 (Why MAMORU) |
| `Shell` | `#F5F3F0` | 카드 배경 (Spec Table) |
| `Cream` | `#FAF9F7` | 메인 배경, CTA 다크 텍스트 |
| `White` | `#FFFFFF` | 카드 내부 (Note 박스) |

**금기**: 골드/포인트 컬러 ❌. 모노크롬 10색만.

---

## 2. Text Styles

Figma → Local Styles → Text에 다음 8개 등록:

| Style 이름 | Family | Size (PC) | Weight | Line Height | Letter Spacing |
|-----------|--------|-----------|--------|-------------|---------------|
| `Display` | Outfit | 112 | 900 | 1.0 | -4% |
| `H1` | Outfit / Plus Jakarta | 64 | 800 | 1.1 | -3% |
| `H2` | Outfit / Plus Jakarta | 52 | 800 | 1.15 | -2% |
| `H3` | Plus Jakarta + Noto Sans KR | 36 | 600~700 | 1.4 | -1% |
| `Body Large` | Plus Jakarta + Noto Sans KR | 18 | 400 | 1.85 | 0 |
| `Body` | Plus Jakarta + Noto Sans KR | 16 | 400 | 1.7 | 0 |
| `Label` | Outfit | 13 | 700 | 1.5 | 25% |
| `Spec Value` | Plus Jakarta | 16 | 700 | 1.5 | 0 |

**모바일 대응**: Figma에서는 두 프레임(PC 840px / Mobile 375px) 별도 작성 권장. HTML의 `clamp()` 자동 스케일을 시각화.

---

## 3. Spacing Tokens

Figma Local Variables 또는 메모로 관리:

| 이름 | 값 (PC) | 용도 |
|------|--------|------|
| `xs` | 8px | 작은 gap |
| `sm` | 16px | 카드 간격 |
| `md` | 24px | 작은 padding |
| `lg` | 40px | 좌우 padding |
| `xl` | 64px | 카드 내부 padding |
| `2xl` | 96px | Hero 텍스트 영역 |
| `3xl` | 140px | 섹션 간 간격 |

---

## 4. 8섹션 컴포넌트 사양

### Section 01 — Hero
프레임: max-width 840px, vertical Auto Layout
```
├─ 라벨 (Label style, Warm Gray) — padding-top 64px, padding-x 40px
├─ IMG_hero (4:5 풀폭, 모서리 radius 0)
├─ 텍스트 영역 (padding 80px 40px 96px, vertical Auto Layout)
│  ├─ 제품명 (Display style, 1줄, white-space nowrap)
│  ├─ 한 줄 카피 (H3 light, max-width 520px)
│  └─ 사양 인라인 (horizontal Auto Layout, gap 40px, Spec Value style)
└─ IMG_front (3:2 풀폭)
```

### Section 02 — Detail Grid
프레임: padding 140px 40px (top) / 0 40px 140px (bottom)
```
├─ 라벨 (Label, margin-bottom 32px)
└─ Grid 2×2 (gap 16px)
   ├─ blade1.png (1:1)
   ├─ blade2.png (1:1)
   ├─ model.png (1:1)
   └─ bolt.png (1:1)
```

### Section 03 — In Action (Void 다크)
프레임: background Void, padding 140px 0
```
├─ 텍스트 영역 (padding 0 40px, padding-bottom 64px)
│  ├─ 라벨 (Label, Warm Gray on Void)
│  └─ H2 "한 컷, 한 손에 멈춥니다." (Cream)
└─ IMG_cut (16:9 풀폭, GIF/MP4)
```

### Section 04 — About
프레임: padding 140px 40px
```
├─ 라벨
├─ H2 "CL1-4T의 특성"
├─ IMG_handle (3:2)
└─ 본문 (Body Large, max-width 620px) × 2 단락
```

### Section 05 — Spec Table (Shell 배경)
프레임: background Shell, padding 140px 48px
```
├─ 라벨
├─ H2 "사양"
└─ 6 행 (vertical Auto Layout, 행 간 Sand 1px border)
   ├─ 좌: label (Body, Warm Gray)
   └─ 우: value (Spec Value, Void)
```

### Section 06 — Note
프레임: padding 140px 40px
```
├─ 라벨
├─ H2 "맞는 분 / 안 맞는 분"
├─ 박스 1 (White, border 1px Parchment, radius 12px, padding 48px, margin-bottom 24px)
│  ├─ "이런 분에게 맞습니다" (H3 small, Void)
│  └─ bullet 3개 (Body, Graphite, 행 간 Parchment 1px border)
└─ 박스 2 (Shell, radius 12px, padding 48px)
   ├─ "맞지 않을 수 있습니다" (H3 small, Warm Gray)
   └─ bullet 3개 (Body, Warm Gray, 행 간 Sand 1px border)
```

### Section 07 — Why MAMORU (Parchment)
프레임: background Parchment, padding 140px 48px
```
├─ 라벨
├─ 큰 인용부호 (Outfit 80px Black, Void)
├─ 인용문 (H3 large 36px, max-width 680px, 사장님 메시지)
├─ 사장님 정보 (horizontal Auto Layout, gap 18px)
│  ├─ 원형 아바타 (72×72px, Sand 배경)
│  └─ 이름 + 직함
└─ 본문 (border-top 1px Sand, padding-top 40px)
   └─ Body × 2 단락
```

### Section 08 — CTA (Void 다크)
프레임: background Void, padding 140px 48px
```
├─ H2 "안 사셔도 괜찮습니다." (Cream)
├─ 본문 (Body Large, Cream 75% opacity, max-width 580px)
├─ 버튼 (Cream 배경, Void 텍스트, padding 22px 28px, radius 8px, max-width 480px)
└─ 풋터 (border-top 1px Cream 15%, padding-top 40px)
   ├─ "CUT THE FAKE, KEEP THE REAL" (Label, Cream 40% opacity)
   └─ 풋터 텍스트 (Body small, Cream 30% opacity)
```

---

## 5. 작업 순서 (사장님이 Figma에서)

### 1단계 — Foundations 셋업 (10분)
1. 새 Figma 파일 생성: "MAMORU 상세페이지 마스터"
2. Color Styles 10개 추가 (위 표 참조)
3. Text Styles 8개 추가
4. Spacing Tokens는 Local Variables로 (선택)

### 2단계 — HTML import (5~10분, 시작점)
1. Figma Plugin **`html.to.design`** 설치 (Figma 우측 상단 → Plugins → Browse → 검색)
2. 플러그인 실행 → URL 입력:
   ```
   https://bsm-pixel.github.io/mamoru/projects/products/master/v10_trendy.html
   ```
3. Import → 자동 변환 → 디자인 정리 (간격/Auto Layout 매뉴얼 정돈 필요)

### 3단계 — Components 정제 (30~60분)
1. 각 섹션을 **Component**로 변환 (우클릭 → Create Component, 또는 `Ctrl/Cmd + Alt + K`)
2. **Auto Layout** 적용 (`Shift + A`)
3. Color Styles / Text Styles 매핑 (스타일 누락된 부분 클릭 → Local Styles 적용)

### 4단계 — Master Page 조립 (20분)
1. "Master Page (CL1-4T)" 페이지에 8개 컴포넌트 인스턴스 배치
2. 더미 콘텐츠로 완성도 확인

### 5단계 — Products Page (제품별 인스턴스) — 양산 시
1. 새 제품마다 Master Page 복제
2. 텍스트/이미지만 오버라이드 (Components라 디자인은 자동 동기화)
3. 디자인 변경 사항이 있으면 → Master Component 수정 → 모든 인스턴스 자동 반영

---

## 6. HTML 갱신 흐름 (사장님 → 클로드)

사장님이 Figma에서 디자인 조정한 경우:

### 옵션 A — 변경 텍스트 설명
사장님이 "[섹션명] [어느 부분] [어떻게 변경]" 알림 → 클로드 코드 갱신.
예: "Hero 제품명 폰트 크기 좀 줄여줘 / Note 박스 radius 키워줘"

### 옵션 B — Figma 스크린샷 + 설명
Figma에서 Frame 우클릭 → Copy as PNG → 클로드에게 첨부 + 짧은 설명. 클로드가 시각 비교 후 코드 갱신.

### 옵션 C — Figma 공유 링크 (가장 정확)
Figma 우측 상단 Share → Copy link (View 권한) → 클로드에게 링크. 단, 클로드가 Figma 직접 접근은 제한적이라 옵션 A/B가 더 빠름.

---

## 7. 양산 시 흐름 (Figma 미러 활용)

새 제품(예: CL1-70) 추가 시:

1. **사장님 (Figma)**: Master Page 복제 → CL1-70 인스턴스 → 사진/텍스트 오버라이드 → 디자인 검토
2. **사장님 (변수 시트)**: 클로드에게 변경 사항 답변 (제품명/사양/카피/Note 등)
3. **클로드 (HTML)**: 마스터 코드 복사 → 변수 갈아끼움 → `projects/products/product_detail/{SKU}/index.html` 생성 + push
4. **사장님 (아임웹)**: 새 상품 등록 → 본문에 코드 paste

→ Figma는 시각 검토 + 디자인 시안 보존 / HTML은 실 양산 = 두 흐름 동시 진행.

---

## 8. 주의사항

- Figma 변경 = HTML 자동 반영 ❌ (수동 동기화 필요)
- Figma JPG export로 양산 시 SEO/검색 ↓ (그래서 HTML 양산 권장)
- Figma는 **시각 검토 + 디자인 실험 + 사장님 직접 조정 + 매뉴얼**용
- 사진은 Figma에 import 시 원본 활용 (PNG 누끼 그대로)
- 모바일/PC 두 프레임 작업 권장 (HTML clamp() 시각화)

---

## 빠른 시작 체크리스트

- [ ] Figma 파일 생성 ("MAMORU 상세페이지 마스터")
- [ ] Color Styles 10개 등록
- [ ] Text Styles 8개 등록
- [ ] `html.to.design` 플러그인 설치
- [ ] v8_trendy URL import
- [ ] 8섹션을 Component로 변환
- [ ] Master Page 조립 (CL1-4T)
- [ ] 사장님 검토 → 디자인 조정 → 클로드에게 알림 → HTML 동기화

→ **첫 Figma 마스터 구축 = ~1.5~2시간**. 이후 새 제품마다 인스턴스 복제 = 10~20분.
