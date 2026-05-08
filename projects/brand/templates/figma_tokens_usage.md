# Figma 토큰 자동 import 가이드

> `figma_tokens.json` 파일을 **Tokens Studio for Figma** 플러그인으로 import하면
> Color/Typography/Spacing/Radius **모든 스타일이 자동 생성**됩니다 (사장님 수동 등록 X).

---

## 무엇이 자동 생성되나?

`figma_tokens.json` 1번 import = Figma에 다음이 모두 생성:

### Color Styles (10개)
- void / graphite / stone / warm-gray / mist / sand / parchment / shell / cream / white

### Text Styles (Typography Composition, 8개)
- display / h1 / h2 / h3 / body-large / body / spec-value / label
- 각 스타일은 폰트 패밀리 / 굵기 / 크기 / 라인 높이 / 자간 모두 포함

### Variables (Spacing/Radius/FontSize/Weight)
- spacing: xs / sm / md / lg / xl / 2xl / 3xl
- borderRadius: sm / md / lg / xl / full
- fontSize / fontWeights / letterSpacing 등 원자 토큰

---

## 사용 단계 (5분)

### 1. Tokens Studio for Figma 플러그인 설치
1. Figma 우측 상단 → **Plugins** → **Browse plugins in Community**
2. 검색: **`Tokens Studio for Figma`** (만든이: Jan Six)
3. **Save** 또는 **Install**

> 무료 플러그인. ~30만 명 사용 중인 디자인 토큰 표준 도구.

### 2. Figma 파일 생성
1. 새 Figma 파일 만들기: **"MAMORU 상세페이지 마스터"**
2. (캔버스 비어있는 상태)

### 3. JSON Import
1. 우측 상단 **Plugins** → **Tokens Studio for Figma** 실행
2. Tokens Studio 패널 우측 상단 **`Settings`** (톱니바퀴 아이콘)
3. **`Tools`** 탭 → **`Load from file/folder or preset`**
4. **`Load from file`** 선택 → `figma_tokens.json` 파일 업로드
5. **Save** 클릭

> 또는: Tokens Studio 패널 → 좌측 상단 **`+`** → **Load from JSON** → 파일 직접 paste

### 4. 토큰을 Figma Styles로 변환
1. Tokens Studio 패널 좌측 하단 **`Apply`** 또는 **`Create Styles`**
2. 옵션:
   - ✅ **Create styles** (Color + Typography를 Figma Local Styles로 자동 생성)
   - ✅ **Create variables** (Spacing/Radius 등을 Local Variables로)
3. **Apply** 클릭 → 자동 변환

### 5. 확인
1. Figma 우측 패널 → 빈 사각형 클릭 → Fill에서 **Local Styles** 확인 → 10개 색상 스타일 보임
2. 텍스트 추가 → 우측 패널 → Type에서 **Local Styles** 확인 → 8개 타이포그래피 스타일 보임
3. ✅ 완료 — **수동 등록 0**

---

## 다음 단계 (토큰 셋업 완료 후)

### A. HTML 자동 import로 시작점 만들기
1. Figma Plugin **`html.to.design`** 설치
2. URL 입력:
   ```
   https://bsm-pixel.github.io/mamoru/projects/brand/iframe_test/v8_trendy.html
   ```
3. Import → 디자인이 캔버스에 자동 변환됨

### B. import된 디자인 정리
1. 각 섹션을 **Component**로 변환 (`Ctrl/Cmd + Alt + K`)
2. **Auto Layout** 적용 (`Shift + A`)
3. 색상/폰트가 누락된 부분은 방금 등록한 Local Styles로 매핑

### C. 마스터 페이지 조립
- 8섹션 컴포넌트 인스턴스 배치
- 더미 콘텐츠 (CL1-4T)로 완성도 검토

---

## 토큰 업데이트 흐름 (향후)

Brand Guide나 디자인이 변경되면:

1. **클로드**: `figma_tokens.json` 갱신 + push
2. **사장님**: Tokens Studio에서 **`Pull from JSON`** 또는 **`Reload`**
3. **사장님**: **`Apply`** → Figma에 자동 반영 (모든 인스턴스 자동 갱신)

→ 디자인 시스템 단일 출처 (JSON 파일) 보장.

---

## 참고

- Tokens Studio 공식: https://tokens.studio/
- 플러그인: https://www.figma.com/community/plugin/843461159747178978/tokens-studio-for-figma
- html.to.design: https://www.figma.com/community/plugin/1159123024924461424/html-to-design

---

## 문제 발생 시

### "Load from file" 메뉴가 안 보임
→ 좌측 상단 `+` 버튼 클릭 → "Load from JSON" 선택 → `figma_tokens.json` 내용 통째로 paste

### Typography가 적용 안 됨
→ Apply 시 **`Create text styles`** 옵션이 체크되어 있는지 확인. 폰트 (Outfit / Plus Jakarta Sans / Noto Sans KR)는 미리 Google Fonts에서 활성화 필요.

### Color는 됐는데 Spacing이 안 보임
→ Spacing은 **Variables**로 생성됨 (Styles 아님). 우측 패널 → Local variables → Number variables에서 확인.

### Apply 후에도 변경 사항 반영 안 됨
→ Tokens Studio 패널 새로고침 (좌측 상단 `Reload` 또는 플러그인 재실행).
