# MAMORU 시스템 구축 — TODO

> 최종 수정: 2026-05-12 — TMS 2단계 매출 3분할 코드 완료 (2-A~2-E + 3-A), push·SQL 실행 대기
> **미완료 항목을 상단에, 완료 이력을 하단에 배치**

---

## 🟡 배포 대기 — TMS 2단계 매출 3분할 (2026-05-12, 커밋만 완료)

코드 6커밋 완료(`5c96d1b`~`e7142e3`), `tsc --noEmit` + `next build` 통과. **push 안 함** (Vercel 빌드 크레딧 — 사장님 승인 시).

- [ ] **push 승인** → Vercel 배포
- [ ] **사장님 Supabase SQL Editor 에서 실행**: `078_hub_stats_b2c_b2b_split.sql`, `079_customers_default_repair_price.sql`
   - 미실행이어도 화면 정상 동작(클라이언트 fallback). 078은 후처리 쿼리 최적화, 079는 거래처 단가 기능 필수.
- [ ] 배포 후 확인: 메인 대시보드 총매출 = B2C 제품 + B2B 제품 + 복원수리 전체 3분할 / 회계 리포트 탭별(전체·제품·복원수리) / 거래처 정보 화면(딜러·아카데미) 복원수리 기본 단가 입력 / 납품 "+B2B수리" 거래처 선택 시 단가 자동 / 판매 입력 "복원수리" 모드(마모루·타사 자루)

내역: 2-A useHubStats salesB2C/salesB2B / 2-B RPC 078 / 2-C 대시보드 KPI 3분할 / 2-D 회계 리포트 RS 집계 정확화(복원수리=A접수+B판매RS+C납품RS, 제품=B2C+B2B 납품포함, by_product RS 제외) / 2-E customers.default_repair_price / 3-A 판매 입력 복원수리 모드. 상세: `memory/project_tms_repair_revenue_split.md`

## 🔥 즉시 작업 가능 — v10 첫 아임웹 paste (2026-05-10)

### Step 1 — 아임웹 paste 검증 (사장님 직접, 15분)
- [ ] 미리보기 확인: https://bsm-pixel.github.io/mamoru/projects/products/master/v10_trendy.html
- [ ] 아임웹 상품 등록 (A2-55FS 또는 시제 모델)
- [ ] 본문 → `</>` 코드보기 → v10_trendy.html 본문 div paste
- [ ] PC/모바일 미리보기 확인
- [ ] 폰트 fallback 확인 (Google Fonts link strip 시 시스템 sans-serif)

### Step 2 — 사장님 자산 작업 (병렬 가능)
- [ ] **사진 PNG** — `product_detail/A2-55FS/images/` (hero/blade1/blade2/handle/back/bolt/model/cut)
- [ ] **공통 자산** — `shared/face.png`, `workshop.png`, `scissors-grip.svg`
- [ ] **라인업 사진** — `lineup/A2/A2-*.png` (5~6장)

### Step 3 — 데이터 답변 (사장님 → 클로드 갱신)
- [ ] GRIP SIZE 측정값 (A2-55FS 엄지부/약지부 가로×세로 mm)
- [ ] 핸들 특성 한 단락 (예: "표준에 가까워...")
- [ ] About / FOR YOU / Why MAMORU 본문 카피
- [ ] VOICES 후기 3개 (실 후기 또는 가공)
- [ ] SAME HANDLE 정확한 라인업 모델 목록

---

## 💡 양산 워크플로 (사장님 확정 방향 2026-05-09)

> **"틀(v10 마스터) 잡아놓고 → 이미지만 로컬 폴더에 넣고 → 그에 맞춰 자동화 느낌으로 클로드와 함께 제작"**

### 다음 모델 양산 시 사장님 액션 (간단)
1. 사진 7~8장 + SVG 아이콘 → `projects/products/product_detail/{모델}/images/` 또는 `icons/` 에 넣기
2. 클로드에게 `다음 모델 진행` 한 마디 + 모델명/변수 답변
3. 클로드가 v10 마스터를 갈아끼워 새 모델 HTML 생성 → push

### 클로드 자동 흐름
1. v10_trendy.html 마스터 복사 → `product_detail/{모델}/index.html`
2. 사장님 변수 시트 답변 받아서 placeholder 갈아끼움
3. `images/` 폴더 사진 자동 참조
4. SVG 아이콘 inline 적용 (currentColor)
5. push → GitHub Pages 자동 배포 → 아임웹 inline paste 준비 완료

---

## 🌅 다음 작업 시작 시 (사장님이 클로드에게 한 마디 입력 → 즉시 가이드)

**입력할 명령어 (자연어, 이 중 하나)**:
- `v10 이어서`
- `어제 작업 이어서`
- `상품 상세 v10 진행`
- `MAMORU 이어서`

→ 클로드가 받으면:
1. 본 TODO.md 상단 (이 섹션) 읽기
2. 우선순위 작업 안내
3. 사장님이 어디부터 시작할지 결정

### 🎯 우선순위 작업 (다음 시작 시 즉시 가능한 순)

#### A. 사장님 명세 확정 (클로드 직접 작업 X — 사장님 정보 필요)
- [ ] **CL → 정식 모델로 변경** (CL은 Clearance 임시)
  - 정식 마스터 모델명 결정 (예: A1-4T 또는 다른 코드)
  - hero/About/SPEC/PROFILE 모두 변경
- [ ] **핸들 분류 정확한 명칭** (현재 더미: A/B/C 핸들)
  - A 핸들 = ? / B 핸들 = ? / C 핸들 = ?
- [ ] **등급 분류 정확한 명칭** (현재 더미: R/A/E/S)
  - 영문/한글 정확한 명칭 알려주기
- [ ] **날등 형태 정확한 명칭** (현재 더미: S/K/B = Straight/Ken-form/Byeol-form)
- [ ] **본 모델의 핸들 그룹** (예: A2 핸들이면 SAME HANDLE 섹션 데이터 갱신)
- [ ] **본 모델의 BLADE 형태** (F/A/G 중 어느?)
- [ ] **본 모델의 BLADE BACK 형태** (S/K/B 중 어느?)

#### B. 사진 촬영 + 추가 (사장님 직접 작업)
- [ ] **back.png** — 뒷면 사선 전체샷 (모델명 보이는 각도, 3:2)
- [ ] **사장님 얼굴** — Why MAMORU 섹션 (원형 작게)
- [ ] **공방 사진** — 가위 점검 / 공방 내부 (3:2)
- [ ] **모델별 날부 2:3 세로 사진** (SAME HANDLE 그리드 5장 — A2-4T/6T/7A/55T/55SL)

#### C. 영상 (cut.gif) — 사장님 직접
- [ ] 블런트 커트 동작 3~5초 촬영 (1280px 이상)
- [ ] [ezgif.com](https://ezgif.com)에서 GIF 변환 + 압축 (3MB 이하)
- [ ] `cut.gif` 파일명으로 폴더에 추가

#### D. 본문 카피 채우기 (사장님 답변 → 클로드 갱신)
- [ ] 한 줄 카피 ("한 손에 머무는 무게, 일관된 끝매김" 유지 or 변경)
- [ ] About 본문 단락 2개 (CL1-4T 특성 — 사장님 직접 작성)
- [ ] FOR YOU? "이런 분에게" 3 bullet
- [ ] FOR YOU? "맞지 않을 수 있습니다" 3 bullet
- [ ] Why MAMORU 인용문 (사장님 직접 메시지)
- [ ] 사장님 자기소개 한 줄
- [ ] VOICES 후기 3개 (실제 상담 후기 사용 가능 시)

#### E. 사진 압축 (사장님 직접, 10분)
- [ ] [TinyPNG](https://tinypng.com)에 7장 드래그 → 압축본 다운로드 → 폴더 덮어쓰기 (28MB → ~10MB)

#### F. 아임웹 실 상품 검증 (사장님 직접, 15분)
- [ ] CL1-4T 또는 정식 모델로 아임웹 상품 등록
- [ ] 본문 → `</>` 코드보기 → v10 코드 paste → 저장
- [ ] PC + 모바일 미리보기 확인

#### G. Figma 셋업 (사장님 결정 — **A 옵션 확정**, 1~1.5시간)

> **방향 확정 (2026-05-09)**: 클로드가 자료 제공 + 사장님이 Figma에서 마무리 셋업.
> 클로드 자료(이미 완성): `figma_tokens.json` / `figma_master_spec.md` / `figma_tokens_usage.md` / `icons/*.svg`

**사장님 진행 단계** (순서대로):
1. [ ] Figma 새 파일 생성 ("MAMORU 상세페이지 마스터")
2. [ ] (필요 시) Outfit + Plus Jakarta Sans + Paperlogy 폰트 설치
3. [ ] **Tokens Studio for Figma 플러그인** 설치 → `figma_tokens.json` import → Apply (5분)
   - Color 10개 + Typography 8개 + Spacing/Radius Variables 자동 등록
4. [ ] **html.to.design 플러그인** 설치 → URL import (10분):
   ```
   https://bsm-pixel.github.io/mamoru/projects/products/master/v10_trendy.html
   ```
5. [ ] import된 디자인 정리 + 8섹션을 Component로 변환 (30~60분)
6. [ ] **SVG 아이콘 6개 import** (`projects/products/icons/blade-*.svg`) → 사장님 톤으로 개량
7. [ ] (선택) Master Page 조립 + Products Page (제품별 인스턴스)

**왜 A 옵션인가** (참고):
- 클로드가 100% 자동 생성은 한계 (Figma Plugin API 외부 접근 X)
- 클로드 자료 + 사장님 셋업 = 가장 빠른 1~1.5시간 완성
- 향후 디자인 조정/실험은 사장님이 Figma에서 자유롭게

---

## 📋 2026-05-09 진행 상황 (오늘 한 일 — PROFILE 영역 완전 재설계)

### ✅ HANDLE 영역 (5개 카드)
- PROFILE [3] HANDLE 5 카드 짧은 설명 추가 (commit `049abca`)
- CORE / DETAIL 그룹 위계 분리 (commit `6b1a5fe`)
- HANDLE SVG placeholder 가로 비율 96×40 + 박스 제거 (commit `1fdf0d2`)
- 사장님 핸들 일러스트 5종 inline 삽입 (commit `0e8fc94`) — 세미오프셋/스탠다드/오프셋/플랫/카멜
- 미선택 SVG opacity:0.35 (commit `95d7e59`)
- [3] HANDLE 라벨 제거 (영문 위계 단순화, commit `673f89d`)

### ✅ BLADE / BLADE BACK 영역 (6개 카드 — Priority 1 + 사장님 SVG)
- 시각적 위계 재설계 (commit `eddb55b`):
  - 알파벳 32~72px → 13~20px (코드 라벨 강등) → 다시 28~64px (시각 영웅 회복)
  - SVG 24~44px → width 100% (메인)
  - 한글명 11~14px → 15~22px → 13~18px (균형)
  - 미선택 카드 SVG opacity:0.35 일관 적용
  - 영문 라벨(FORCE/ALL-ROUND/GLIDE/SWORD/CONVEX/BEVELED) 제거
- 사장님 일러스트 SVG 6종 적용 (commit `b695c1e`):
  - F/A/G — 가로 viewBox 21.65×10
  - S/C/B — 정사각형 viewBox 24.32×24.32 (3D 입체)
- 알파벳 큼직 + 한글명 가로 한 줄 배치 (commit `501c2c9`)
- SVG 위 → 헤더(F 직선형 ✓) → 구분선 → 설명 (Editorial 캡션 패턴, commit `4d4c0ae`)

### ✅ 자산 보존
- 사장님 SVG 11개 원본 commit (`83f4edc`) — `projects/products/icons/`
  - BLADE/BLADE BACK 6개 (F/A/G/S/C/B)
  - HANDLE 5개 (handle-flat/camel/semi-offset/standard/offset)

### 🎯 디자인 톤 추가 확정
- BLADE/BLADE BACK 카드 = SVG 위 → 라벨 → 구분선 → 설명 (4단 흐름)
- 알파벳 + 한글명 = 한 줄 가로 정렬 (baseline align)
- 미선택 카드 SVG opacity:0.35 통일 룰 (HANDLE / BLADE / BLADE BACK 모두)
- CORE (가위의 본질) / DETAIL (잡는 느낌) 그룹 라벨로 위계 분리

---

## 📋 2026-05-08 진행 상황 (이전)

### ✅ 디자인 안정화 — v10 13섹션 풀 마스터
- v5~v9 정리 (8개 파일 삭제)
- 모델명 폰트 페이퍼로지 적용 (commit `563a7ab`)
- 사양 인라인 "일본 ATS-314" → "Blunt" (사장님 직접)
- 사진 영역 배경 Parchment → Shell (Brand Guide D-04 정식)
- Hero 메인 2 (전면 가로) 제거 + DETAIL 풀폭 3장 + 1:1 클로즈업 2장
- DETAIL 비율 1:1 → 3:2 + auto-fit minmax(360, 1fr)
- PROFILE 카드 SVG 우측 상단 + ✓ 체크만
- HANDLE → BLADE BACK (날등 형태) 전환
- SAME HANDLE 신설 (07번, SPEC 직후)
- VOICES 위치 이동 (08번, SAME HANDLE 직후)
- BRAND TRANSITION 신설 (Void 슬로건, VOICES 직후)
- LINEUP → GRADE 분리 (등급 카드만)
- SVG 아이콘 6개 별도 파일 export (`projects/products/icons/*.svg`)

### 🎯 디자인 톤 확정 (확정 사항)
- v10 = 13섹션 풀 마스터
- 흐름: 객관(Hero/Detail/Action/About/PROFILE/SPEC/SAME HANDLE) → 외부 시각(VOICES) → BRAND TRANSITION → 내부 정체성(VS/WHY/CARE/GRADE/CTA)
- 사진 톤: Shell 배경 + object-fit:contain
- 모델명 폰트: Paperlogy
- BLADE TYPE 카드: SVG 라인 일러스트 (사장님 Figma 개량 예정)

---

## 📋 2026-05-08 진행 순서 (이전)

### 현재 상태 (확정)
- ✅ HTML 마스터: `projects/products/master/v8_trendy.html` (8섹션, 사진 7장 적용)
- ✅ 양산 방식: inline HTML 직접 paste (검증 완료)
- ✅ 흐름맵: 8섹션 (Hero / Detail / In Action / About / Spec / For You / Why MAMORU / CTA)
- ✅ 디자인 톤: A 타이포 + B 사진 하이브리드
- ✅ 사진 7장 폴더: `projects/products/product_detail/CL1-4T/images/`
- ✅ Figma 하이브리드 (C 옵션) 결정
- ✅ Figma 토큰 JSON + 가이드: `projects/products/templates/figma_tokens.json` `figma_tokens_usage.md` `figma_master_spec.md`

### 🎯 Step 1 — v8 디자인 검증 (5~10분, 즉시)
- [ ] [v8_trendy 미리보기](https://bsm-pixel.github.io/mamoru/projects/products/master/v8_trendy.html) PC + 모바일에서 다시 보기
- [ ] 만족스러운지 결정:
  - ✅ OK → Step 2로
  - ⚠️ 부분 조정 → 어디 수정 원하는지 클로드에게 알림 → 코드 갱신 후 다시 검증
  - ❌ 큰 방향 변경 → v9 시안 (드물게)

### 🎯 Step 2 — 본문 카피 채우기 (30분, 사장님은 답변만)
v8 디자인 OK 후 진행. 클로드가 항목별 질문 → 사장님 답변 → HTML 갱신.
- [ ] 한 줄 카피 ("한 손에 머무는 무게, 일관된 끝매김" 유지 or 변경)
- [ ] 사양 정확한 값 (길이/무게/소재/베어링/곡률 — 더미가 정확한지 확인)
- [ ] About 본문 단락 (CL1-4T 특성 — 사장님 직접 작성)
- [ ] "이런 분에게 맞습니다" 3 bullet
- [ ] "맞지 않을 수 있습니다" 3 bullet
- [ ] Why MAMORU 인용문 (사장님 메시지)
- [ ] 사장님 자기소개 (한 줄)

### 🎯 Step 3 — 사진 압축 (10분, 사장님 직접)
- [ ] [TinyPNG](https://tinypng.com)에 7장 드래그
- [ ] 다운로드 → `projects/products/product_detail/CL1-4T/images/` 폴더에 덮어쓰기
- [ ] 결과: 28MB → ~10MB (모바일 로딩 ↑)
- [ ] 클로드에게 알리면 commit + push

### 🎯 Step 4 — 영상(cut.gif) 촬영 (별도 시간)
- [ ] 블런트 커트 동작 3~5초 촬영 (1280px 이상 권장)
- [ ] [ezgif.com](https://ezgif.com)에서 GIF 변환 + 압축 (3MB 이하)
- [ ] `cut.gif` 파일명으로 폴더에 추가
- [ ] 클로드에게 알림 → IN ACTION 섹션 placeholder 교체

### 🎯 Step 5 — 아임웹 실 상품 검증 (15분)
- [ ] CL1-4T 아임웹 상품 등록 (이미 있으면 skip)
- [ ] 본문 → `</>` 코드보기 → v8_trendy 코드 paste
- [ ] PC + 모바일 미리보기 확인
- [ ] 펼치기 동작 정상 확인
- [ ] ✅ 통과 → CL1-4T 시제품 1호 완성

---

## 🅱️ 병행 가능 작업 — Figma 셋업 (선택, 1~2시간)

급한 거 아니면 다음 제품 양산 전에 진행 권장. 디자인 조정 자유도 ↑.

- [ ] **Tokens Studio for Figma** 플러그인 설치
- [ ] (필요 시) Google Fonts에서 Outfit + Plus Jakarta Sans 시스템 설치
- [ ] Figma 새 파일 생성 ("MAMORU 상세페이지 마스터")
- [ ] `figma_tokens.json` import → Apply → Color/Typography/Spacing 자동 등록 (5분)
- [ ] **html.to.design** 플러그인 설치 → v8_trendy URL import → 자동 시작점 (5분)
- [ ] import된 디자인 정리 + 8섹션을 Component로 변환 (30~60분)
- [ ] Master Page 조립 (CL1-4T 더미)
- [ ] 향후 디자인 조정 시 → Figma에서 변경 → 클로드에게 알림 → HTML 동기화

가이드:
- `projects/products/templates/figma_tokens_usage.md` (토큰 import)
- `projects/products/templates/figma_master_spec.md` (마스터 사양)

---

## 🅲️ 시제품 1호 완성 후 (클로드 작업)

- [ ] v8_trendy → `projects/products/product_detail_template.html` 정식 마스터로 이동
- [ ] 변수 시트 양식 commit (`templates/spec_sheet_scissors.md`)
- [ ] 카피 가이드 commit (`templates/copy_brief.md` — Brand Guide B-04 보이스)
- [ ] 양산 매뉴얼 commit (`MANUAL_PRODUCT_DETAIL.md` — 사장님/직원이 보고 양산 가능)

---

## 🅳️ 양산 (Phase C, 추후 반복)

각 새 제품(CL1-70 등)마다:
1. 사장님: 변수 시트 채우기 (클로드 AskUserQuestion 패턴)
2. 사장님: 사진 7장 + 영상 촬영/압축
3. 사장님: `projects/products/product_detail/{SKU}/images/` 폴더에 추가
4. 클로드: 마스터 → `product_detail/{SKU}/index.html` 생성
5. 사장님: GitHub push (자동 GitHub Pages 배포)
6. 사장님: 아임웹 새 상품 등록 + 본문 paste
7. 1제품 ~30분 내 완성

---

## 🅴️ 추후 (별도 작업)

- [ ] 주변제품 마스터 추가 (빗 / 롤브러시 / 핀셋 — 다른 톤, 별도 마스터)
- [ ] 양산 매뉴얼 직원용 작성 (사진 촬영 가이드 / 변수 시트 작성법)

---

## 🗂️ 핵심 파일 위치 (이어서 작업 시 참조)

| 파일 | 용도 |
|------|------|
| `projects/products/master/v8_trendy.html` | HTML 마스터 (시안 단계) |
| `projects/products/product_detail/CL1-4T/images/` | CL1-4T 사진 폴더 (7장 + cut.gif 추후) |
| `projects/products/templates/figma_tokens.json` | Figma 토큰 (Tokens Studio import용) |
| `projects/products/templates/figma_tokens_usage.md` | 토큰 import 가이드 |
| `projects/products/templates/figma_master_spec.md` | Figma 마스터 사양서 (8섹션) |
| `memory/reference_imweb_product_detail_inline.md` | 아임웹 inline style 검증된 패턴 |
| `.claude/MAMORU-Complete-Brand-Guide-v1.0.md` | Brand Guide (D-03 상품 상세) |
| `C:\Users\user\.claude\plans\tms-stateful-valiant.md` | 플랜 원본 |

---

## ✅ 2026-05-07~08 작업 이력

### 양산 방식 검증 (commits `70640c0` `be46017` `20e7493` `1f50db7`)
- iframe → inline 전환 (script 차단 발견 → inline style + clamp() 정착)
- v3/v4 검증 — 펼치기 동작 + 4섹션 더미 흐름

### 디자인 컨셉 (commits `10727dd` `a1b9d8f`)
- v5 3개 컨셉 시안 (A 타이포 / B 사진 / C 장인 스토리)
- 사장님 결정: **A + B 하이브리드**

### v6/v7/v8 진화
- v6 하이브리드 (commit `7218592`) — Hero + 시장진실 결합 시안
- v6 풀 마스터 (commit `ad705d3`) — 11섹션
- v7 (commit `9d2ac85`) — 사진 6컷 + 영상 섹션
- v8 트렌디 (commit `371bc54`) — 8섹션 압축 + 라벨 영문화 + 영상 위로
- v8 fix (commit `c9aad32`) — 모델명 1줄 + 구매 버튼 제거 (CLAUDE.md 중복 CTA 룰 적용)
- v8 사진 7장 매핑 재구성 (commit `f6a1722`) — 사장님 앵글 정리 (사선/전면/날부2/모델명/볼트부/핸들부)
- 실 사진 7장 적용 (commit `b2a8389`) — 28MB

### Figma 하이브리드 (commits `834c94c` `07a714e`)
- Figma 마스터 사양서 commit
- Tokens Studio JSON + 사용 가이드 commit
- 사장님 결정: **C 옵션 (HTML 양산 + Figma 미러)**

---

## ✅ 2026-05-03 (저녁) 작업 이력

### 상담 — 수동 일정 확정 (079)
- [x] **우측 상세 패널 "수동 일정 확정" 버튼** (commit `652ab3a`) — DM/유선 협의된 일정 즉시 확정 + 알림톡 + 캘린더 자동
- [x] **알림톡 변수 매칭 fix** (commit `22ccc77`) — admin-create 대비 누락 필드(change_request_link 등) 보강 → 3109 SMS 대체 사고 fix
- [x] **MANUAL_CONSULTATION.md 직원 가이드 작성** — 단계별 사용법 + 일정 수동 등록과의 차이

### 창고재고 — 재고조사 인쇄
- [x] **재고조사 인쇄 기능** (commit `4f3f485`) — 화면 필터/정렬 그대로 + 카테고리 자동 그룹화 + 실측 빈 칸 + 비고 영역
- [x] **MANUAL_INVENTORY.md 직원 가이드 작성** — 5단계 사용법 (필터 → 인쇄 → 카운트 → 차이 발견 → 재고 조정)

### 메인 페이지 iframe (이전 작업)
- [x] **iframe_main_top init=0px 회귀** (commit `c7caab0`) — 빈 공간 사고 fix + REQUEST_HEIGHT 동기화 강화

---

## ✅ 2026-05-03 작업 이력

### 가이드 페이지 위계/중복 정리 (사장님 룰 적용)
- [x] **상담 가이드 CTA 3중 보장 + Brand Guide 정합성** (commits `108d292` `1b115c2` `439e377` `55051c3`) — Hero CTA + PC 탭 인라인 CTA + 모바일 floating + Brand Guide ADDENDUM § 5 PC 수치(680px) 정합 + Hero CTA 다크 배경 묻힘 fix
- [x] **iframe wrapper 부모 측 모바일 floating CTA** (commit `ec6ebfc`) — iframe 안 fixed가 iframe document 끝점 기준이라 작동 X. 부모 wrapper에 직접 fixed로 추가 (상담 + 복원수리 가이드 둘 다)
- [x] **모바일 CTA 중복 제거** (commit `298ebd2`) — page 내부 floating 영구 숨김 (wrapper로 단일화) + Hero secondary "Q&A 보기" HTML 제거 (탭과 중복) + Hero CTA 모바일 숨김 (wrapper로 단일 진입점)
- [x] **사장님 룰 박제 "전체 흐름 + UI/UX 위계 점검 필수"** — `memory/feedback_page_holistic_review.md` + MEMORY.md 인덱스
- [x] **페이지 작업 전 iframe 환경 식별 메모리** — `memory/reference_iframe_pages.md` + MEMORY.md 인덱스

### 가이드 페이지 콘텐츠 정합화
- [x] **과정안내 Step 1 위계 강화 + "또는" 구분자** (commits `aeb3128` `e2b9ad3`) — page_as_guide(다크) + page_guide(라이트) 두 가이드 동일 패턴. step-title 18px 굵게 + method-group 카드 강화 + "또는" divider
- [x] **"마모루 컨설팅" 탭 시장 문제점 3블록 제거** (commit `f412525`) — brand/page_intro와 70~80% 중복. 페이지 간 메시지 분담: brand=Why / consulting·as guide=How
- [x] **상담철칙 ↔ 다크 메시지 영역 swap** (commit `67da51e`) — 실무 원칙 먼저 → 다크 메시지로 결심 강화

### 메인 페이지 + 사장님 직접 수정
- [x] **메인 슬로건 "거짓 없는 본질"** (commit `85e9cad`) — Brand Guide "조용히 압도" 톤 + 영문 슬로건과 한국어 일관
- [x] **상담철칙 03 추가 + 텍스트 정제** (commit `1885869`) — 사장님 직접 수정 (01 desc / 02 타이틀 간결 / 03 신규 "고객님께 맞는 제품이 없다면")

---

## ✅ 2026-05-02 작업 이력

### 078 — 상담 달력 관리 + 휴무 SSOT 통합
- [x] **신규 화면 `/consultations/calendar`**: 4개월 달력(현재월 ~ +3) + 정기 휴무 요일 토글(7개) + 임시 휴무일 토글 모달
- [x] **신규 API**: `/api/consultation/blackouts` (closed_dates CRUD), `/api/consultation/settings` (consultation_settings PATCH)
- [x] **신규 hook**: `use-blackouts.ts` (조회/추가/삭제/요일 토글)
- [x] **SSOT 통합**: 설정 → 상담 설정의 "휴무 요일" + "특별 휴무일" UI 삭제 → "달력 관리로 이동" 링크로 대체
- [x] **사장님 룰 박제**: 막힘은 고객 셀프 예약에만 적용, 사장님 측 흐름(일정수동등록/시간제안)은 항상 유동 — `memory/feedback_consultation_blackout_rule.md`
- [x] **흐름도/매뉴얼/로드맵 갱신**: MANUAL_CONSULTATION + TMS_FLOW_CONSULTATION + TMS_SYSTEM_ARCHITECTURE + TMS_ROADMAP
- commits: `ca95025` (1차) + `b270439` (옵션 C 통합)

### 리뷰 모달 viewport 중앙 fix + 페이지별 iframe 환경 메모리
- [x] **리뷰 모달 fix**: `iframe_reviews.html` wrapper에 `MAMORU_REQUEST_VIEWPORT` 응답 코드 추가 (commit `f066717`)
- [x] **페이지별 iframe 환경 메모리 박제**: `memory/reference_iframe_pages.md` (13개 iframe + 8개 단독 페이지 표)
- [x] **사장님 룰 메모리 추가**: 페이지 작업 전 iframe 환경 식별 필수 + 상담 블랙아웃 룰

### 075/076 잔여 (Phase 3) 진행
- [x] **SSOT 점검 SQL** (customer.outstanding_balance ↔ 미수금 합계): 정합성 완벽 (0 rows) — commit `1de203c`
- [x] **사장님 SQL 실행**: 075/076 마이그레이션 완료

---

---

## 📅 2026-05-01 작업 이력

### 통합 수정 — 상담/캘린더/대시보드 KST 정합성
- ✅ 상담 취소/삭제 안전화 (consultation/[id] DELETE — 이력 침묵 제거 + errMsg 통일)
- ✅ 클라이언트 에러 직렬화 통일 (use-consultations 7개 mutation)
- ✅ 출장 흐름 캘린더 동기화 일관성 (cancel/resched/suggest 모두 await + errMsg)
- ✅ 입출금 카테고리 설정화 + 회계 흐름도 매뉴얼 보강
- ✅ 076 RPC KST timezone fix (대시보드 월 시작이 UTC라 4월 데이터 합산되던 버그)
- ✅ **클라이언트 toISOString UTC 버그 통합 fix**: useHubStats msd, useSalesStats today/weekStart 등
  - 원인: `new Date(...).toISOString().slice(0,10)`이 KST 5/1 자정 → '2026-04-30'으로 잘못 변환
  - 결과: 4/30자 deliveries 23만원이 5월 매출에 합산됨
  - 수정: `lib/utils/format.ts`에 `toLocalDateString(d)` 헬퍼 신설 + 모든 핵심 위치 교체

### 새 지침 (CLAUDE.md + memory)
- ✅ **Root-First 원칙**: CLAUDE.md 1.5섹션 신설 — 표면 패치 금지, Phase 1~5 절차 강제
- 사장님 강조: 한 번 수정으로 깔끔히 끝나야 함. 표면만 보고 끝내지 말고 뿌리부터 모든 경로 훑기

### 남은 잠재 이슈 (낮음, 입력 폼 default — 사용자 직접 선택 가능)
- [ ] `toISOString().slice(0, 10)` 패턴 78곳 (대부분 입력 폼 기본값) — 자정~09시 사이 새 입력 시 어제 날짜로 default 잡힐 수 있음. 별도 작업으로 분리 (회귀 위험 격리).

---

## 📅 예정된 작업 (날짜 도래 시 사장님이 "할일 뭐야?" 물으면 안내)

- [ ] **2026-05-07 이후 — GAS 코드 archive (Phase 3)**: 1주 모니터링 후 GAS 코드를 archive 폴더로 이동. 대상: `projects/consulting/Code.gs`, `supabase-sync.gs`, `Total_Management_System/gas/consultation-sync.gs`. 코드 보존 (삭제 X), Apps Script 프로젝트는 유지(트리거만 OFF). 사장님이 1주 동안 알림톡 정상 도착 확인 후 진행.

- [ ] **2026-06-15 이후 — 솔라피(Solapi) 직접 호출 마이그 검토**: 현재 TMS → Make webhook → Solapi 흐름을 TMS → Solapi 직접으로 단축. Make 비용 절감 + 발송 속도 개선 + 디버깅 용이. 작업 시간 ~반나절(코딩 4~5h + 사장님 비즈 콘솔에서 templateId 수집 1~2h). 진행 전 확인 필요: ① 카카오 비즈 콘솔 접근 ② 발신번호 Solapi 등록 ③ Make 월 사용량 / 비용 (ROI). 점진 마이그 안전장치: `NOTIFY_PROVIDER=make|solapi` 환경변수 토글. **GAS 폐기 안정화(약 1.5개월) 후 진행 권장 — 안 깨진 시스템 굳이 깨지 않음 원칙**.

---

## 🔴 진행 보류 / 추후 해결

- [ ] **회사소개 페이지(`page_intro`) 모바일 sticky 퀵네비**: iframe 환경에서 부모(아임웹) 코드위젯 + 자식(page_intro) postMessage 통신 패턴으로 구현 시도(커밋 `7b2260a` `6b0c404`). 모바일 스크롤 시 chip이 상단에 따라오지 않음 — 추후 정밀 진단 필요. 후보 원인: ① 아임웹 페이지 빌더의 위젯 컨테이너 `overflow: hidden` 가능성 ② 코드위젯 자체에 `position: sticky` 미적용 (페이지 빌더가 자체 wrapping) ③ iframe `mamoruIntroFrame`이 위젯 위에 있을 때 z-index/scroll 컨텍스트 충돌. 우선 다른 작업 처리 후 재검토.

---

## 📌 네이버 리뷰 운영 가이드 (월 1회 권장)

### 1단계: 네이버 리뷰 추출
상세: [projects/marketing/README.md](projects/marketing/README.md)
1. 크롬에서 네이버 스마트플레이스 관리자 → 리뷰 관리 → 스크롤 끝까지
2. F12 → Console 탭
3. `projects/marketing/naver_review_extract.js` 전체 복사 → 붙여넣기 → Enter
4. `naver_reviews_YYYY-MM-DD.zip` 자동 다운로드
5. 압축해제 후 폴더에서 `powershell -ExecutionPolicy Bypass -File .\download_images.ps1` 실행 (이미지 다운로드)

### 2단계: TMS 네이버 리뷰 등록
1. TMS → 리뷰 관리 → 네이버 리뷰 등록 페이지
2. 고객별 폴더 열기 (예: `001_2024-09-10_김_관/`)
3. 타입 선택 (상담/복원수리/제품구매)
4. **"📄 review.md 파일 선택"** → 작성자/작성일/방문일/본문 자동 입력
5. 사진 업로드 (photo_01.jpg, photo_02.jpg)
6. **"등록하기"** 클릭 → DB 저장

---

## 📌 미완료

- [ ] **Order API 전환**: 주문 취소 자동화 + 송장 수정/삭제
- [ ] **판매 출고 알림톡 외부 설정**: 솔라피 배송조회 버튼 변수명 확인 (솔라피 문의 대기중)

---

## ✅ 완료 — 04-30 심야 +3 (075)

- [x] **TMS 즉각 반영 풀 연동 + 사장님 보고 3건 fix** (2026-04-30 심야 +3) — 사장님 보고: ① 설정에서 추가한 경비 카테고리 실전 화면 미반영 ② 대시보드 "일정 재요청" 0건인데 1건 표시 ③ 복원수리 이번달 매출 표기 부정확. 진단 결과 3개 근본 패턴 발견(hard-coded 누락 / status 필터 버그 / cross-domain invalidation 끊김). 수정: ① `expenses/page.tsx` `useSetting('accounting.expense_categories')` 적용 ② `use-dashboard-stats.ts:154` + `migrations/075_hub_stats_rpc_v3.sql`에서 `pending_admin` 제거 ③ 복원수리 매출 정의 옵션 A "발생 기준"(사장님 합의) — A채널 paid_at 조건 제거, B채널 category='RS' + total_price>0 ④ `lib/query/invalidate-keys.ts` helper + 모든 sale/repair mutation에 `invalidateFinancialQueries(qc)` ⑤ staleTime sales-stats 60s→30s, products 5분→1분. 흐름도/매뉴얼 4종 갱신. 사장님이 SQL Editor에서 075 실행 완료. commit `20d4752`

---

## ✅ 완료 — 04-26 작업

### 오후 추가 작업

#### 메인 페이지 YouTube 섹션 신설
> Quick Nav ↔ "라인업 교체" 배너 사이에 영상 콘텐츠 허브 신설. lite-youtube 패턴 자체구현 (라이브러리 0).

- [x] **`page_main_top` [3.5] mm-videos 섹션 추가** — 모바일 가로 스크롤(snap, 75vw 카드 + max 300px) / PC 3열 grid(max-width 960~1080px, 좌측 정렬) / 모노크롬 ▶ Play 버튼 / 16:9 썸네일 ✅ commit `6d67c63`
- [x] **첫 영상 등록** — `1LZhDgEyrMA` 사장님 본인 제작 "미용가위에 가까워질 감각 — 3가지 감는 방법" ✅ commit `baebe0d`
- [x] **카드 추가 패턴 정착** — `data-yt="VIDEO_ID"` + 제목 한 줄로 운영. 빈 ID 자동 숨김 + 보일 카드 0개 시 섹션 자동 숨김. 썸네일/href 자동 채움 ✅
- [x] **모바일 영상 가로 스크롤 → 페이지 세로 스크롤 가로채기 버그 수정** — `touch-action: pan-x` + `overflow-y: hidden` + `overscroll-behavior-x: contain` 적용 ✅ commit `6a149fd`
- [x] **다음 카드 peek 첫 화면 노출** — 카드 폭 80vw→75vw, max 320→300px / 마스크 페이드 90%→94% 약화 (작은 폰 38px·큰 폰 60px peek) ✅ commit `e36243a`

#### 메인 영역 iframe 통신 강화 — 만성 잘림 현상 해결
> 사장님이 모바일 PWA(아임웹)에서 자주 겪던 "Trust Numbers 아래 콘텐츠 통째 사라짐" 근본 원인 해결.

- [x] **page_main_top / page_main_btm / page_main 3개 파일 `initIframeComm()` 강화** — 🔒 수정금지 마커 → ⚙️ 강화 적용 2026-04-26 (외부 인터페이스 100% 호환 유지) ✅ commit `d4c6f75`
- [x] **보수적 over-estimation 7중 안전망**: ① 4가지 측정값(body/document/lastEl) 최댓값 ② img load 이벤트 바인딩 ③ document.fonts.ready 후 재전송 ④ 다단 setTimeout (200/600/1500/3000/5000ms) ⑤ resize 디바운스 ⑥ REQUEST_HEIGHT 양방향 메시지 ⑦ ?debug=1 콘솔 로그 ✅
- [x] **확인 결과** — 일주일 모니터링 후 잘림 재발 0회 시 consulting/as 영역도 동일 패턴 일괄 적용 검토

#### 진단 페이지 (page_diag.html) Lottie 도입
> Q_FEEL / Q_STYLE / Q_HABIT 세 질문에 Lottie(.json) 모션 적용. Jitter 작업물 → ./icons/ 폴더 → 데이터 슬롯에 경로 입력 흐름.

- [x] **진단 SVG 아이콘 18종 정비** — 신규 18개(15/25/38/NEW12/blunt/change12/female/level2,3,_1/long/longcut/longsingl/male/nomal/slide1312/thick/up12) + 폐기 3개(Level_2/Slide/level_3) + thinning 업데이트 ✅ commit `84cbf89`
- [x] **Lottie 인프라 점검 — 이미 완비됨 확인** — `<dotlottie-player>` 라이브러리 head 로드, `renderGif()` lottieUrl→gifUrl→placeholder 분기, 데이터 슬롯 모두 준비. 추가 작업 0 ✅
- [x] **Q_FEEL/Q_STYLE/Q_HABIT 사용 가이드 코멘트 추가** — Q_FEEL 상단 통합 가이드(4단계 사용법 + 우선순위 + 캔버스 320×180 / 16:9 권장) + Q_STYLE/Q_HABIT 직전 reminder ✅ commit `7f7c79e`
- [x] **첫 Lottie 작업 등록** (사장님) — `style_go.json` / `style_back.json` 2개 ✅
- [ ] **남은 7개 옵션 Lottie 채우기** — feel_soft/feel_power/feel_none, style_none, habit_wet/habit_dry/habit_none (사장님 Jitter 작업 진행 중)

### 고객 대면 페이지 21개 Brand Guide v1.0 정합성 정비
> 5개 영역 / 5 commit / 모두 GitHub Pages 자동 배포 (Vercel 빌드 0회, 비용 0원)

- [x] **Phase A 브랜드 (1)** — `page_intro` 골드 팔레트 "히어로 한정 예외" 주석 명시 + PC 본문 16px + `navigateTo` → `window.mm` 네임스페이스 (인라인 핸들러 동시 변경) ✅ commit `4afa23c`
- [x] **Phase B 메인 (5)** — `page_main_btm` body color cream→void(잠재 가독성 폭탄 제거) + 네이버 칩 #03C75A → Sand+Stone 모노크롬 + 라벨 11px / `page_main` Pretendard 제거 + hero PC 16px + 카드 13px / `page_main_top` Pretendard 제거 / `page_label_first/lineup` Pretendard 제거 + kicker weight 700 + 폰트 로드 정정 ✅
- [x] **Phase C 상담 (7)** — `page_guide` `--mm-gold` → `--mm-ink` 28곳 + 더블 br 단락 분리 / `page_form` `--mm-gold/--gold-dark/--trust-gold` 31곳+ 치환, Terracotta 주석 9곳 정리, SVG `stroke="#C9A962"` → currentColor, --error 모노크롬, 본문 13px 6곳 / `page_diag` `.replace(/\n/g,'<br>')` → CSS `white-space: pre-line` / `page_suggest` 모달 br 분리, 비표준 #f0f9ff → Parchment+Stone / `page_change_request` 상태 카드 4종(error/success/cancel/rebook) #fef2f2/#f0fdf4/#fffbeb/#fef3c7 → Shell/Parchment+Void/Stone / `page_recommend` 이모지 ✨📦⭐ 제거 / `page_result` 동적 br → display:block ✅ commit `a96e634`
- [x] **Phase D 복원수리 (4)** — `page_as_report` 상태 배지 #e8f5e9/#fff3e0/#ffebee → 모노크롬 + 카카오 인앱 닫기 fallback iOS/안드로이드 분기 처리(memory/reference_kakao_inapp_close.md 패턴) + mamoru.kr 강제 이동 제거 / `page_form` --error #ef4444 → Void + max-width 720→680px / `page_guide` 탭 콘텐츠 700→680px / `page_as_guide` 더블 br → margin, Sand 강조 → Cream ✅ commit `252506e`
- [x] **Phase E 리뷰+정품확인 (3)** — `page_reviews` 베스트 카드 다크 박스(Void) → Shell+Void 라이트(가이드 "라이트 페이지 안 다크 박스 = 장식" 위반 해소) + 네이버 칩 모노크롬 + 라벨 11px / `page_review` 사진 호버 #E5E3E0 (가이드 외) → Sand / `verify/index` valid 배너 그라디언트 → 단색 Void + --danger 모노크롬 + SKU 11px + 타임존 Asia/Seoul ✅ commit `9c63f1d`
- [x] **잔존 마무리** — `page_main_btm` Pretendard 잔존분 + `page_form` --error #C44040 + .same-day-notice 빨강 배경 + SVG 폐기 골드 stroke 3곳 ✅ commit `1ea963a`
- [x] **검증 통계 (Before → After)**: 폐기 컬러 80곳 → 10곳 (88% 감소, 잔존은 모두 의도적) / Pretendard 5파일 → 0파일 (100% 제거) / 본문 12px 142곳 → 74곳 (라벨 11px 하한 안의 사용은 유지) ✅
- [x] **사용자 결정 사항** — 작업 범위 Critical+Major / 진행 순서 영역별 순차 / page_intro 다크+골드 히어로 한정 예외 유지 / 외부 브랜드 컬러(네이버 등) 모두 모노크롬화 ✅
- [x] **펜딩 항목 (별도 결정 필요)**: ① 메인 후기 외부 API `app-eta-sandy-75.vercel.app` 자체 도메인 이관 (도메인·DNS·환경변수 결정 필요) ② Polish 단계(진단 progress bar, Trust Number "할인=0" 카피, Masonry 동적 컬럼) — 플랜 파일: `C:\Users\user\.claude\plans\validated-spinning-tower.md` 참조

---

## ✅ 완료 — 04-24 작업

### 상담관리 일정 수동 등록 기능 (인스타DM/유선 접수용)
- [x] **API `/api/consultation/admin-create` 신규** — 관리자 전용 등록 (인증 · 검증 · 중복체크 · 지오코딩) ✅
- [x] **중복 체크 로직** — phone_normalized 기준 + 오늘 이후 + 같은 visit_date/time → 409 경고 ✅
- [x] **수기 접수 마킹** — `gas_raw.source = 'admin_manual'` 필드로 리포트 집계 구분 ✅
- [x] **출장 지오코딩** — 카카오 REST API로 주소→좌표 자동 변환 (지도 표시용) ✅
- [x] **CreateConsultationModal** — 타입 세그먼트 · 동적 필드 · 알림톡 체크박스 ✅
- [x] **DuplicateWarningModal** — 기존 상담 정보 카드 + 입력 복귀/기존 확인 옵션 ✅
- [x] **useCreateConsultation 훅** — AdminCreatePayload 타입 + 409 duplicate 정상 처리 ✅
- [x] **상담관리 페이지 상단 "일정수동등록" 버튼** — 탭 행 우측 CalendarPlus 아이콘 ✅
- [x] **알림톡 자동 발송** — confirmed/field_confirmed 템플릿 (notify=true 기본) ✅
- [x] **Google Calendar 자동 동기화** — after() 래퍼로 이벤트 생성 ✅
- [x] **리마인더 cron 자동 포함** — status=confirmed + visit_date/time 기반 별도 작업 없이 반영 ✅

### 아임웹 배너 모달 UX 정리
- [x] **우측 상단 X 버튼 제거** — 하단 "오늘 하루 보지 않기" / "닫기" 버튼만으로 통일 ✅
- [x] **TMS 미리보기 모달 닫기 버튼 기능 연결** — 실제 onClose 동작 ✅

### 상담 리마인더 방문주소 누락 복구
- [x] **send-reminders cron** — SELECT 에 address_road/detail 추가 ✅
- [x] **data.address 치환 변수 전달** — FIELD_REMIND_24H/2H 솔라피 템플릿의 #{address} 정상 ✅

### 푸시 알림 테스트 발송 패널 (어제 04-23 완료분 계속 운용)
- [x] **알림 구분** — 리뷰/상담 3종/출장 2종/복원수리/주문 9타입 ✅
- [x] **3건 실패 원인 확인** — 과거 등록 만료 토큰 · 고객 영향 없음 확인 ✅

---

## ✅ 완료 — 04-23 작업

### 아임웹 상품 동기화 사일런트 실패 버그 수정
- [x] **원인 파악** — products.sku UNIQUE 제약 + error 체크 누락으로 31건 조용히 실패 ✅
- [x] **product-sync.ts 재설계** — 매칭 우선순위 (imweb_product_no → sku → 신규) ✅
- [x] **`.single()` → `.maybeSingle()`** — no-match 시 에러 대신 null 반환 ✅
- [x] **update/insert error 반환값 체크** — 실패 수집해 정확한 집계 ✅
- [x] **동기화 결과 투명화 UI** — total_fetched / synced / created / updated / linked / failed 구분 표시 + 실패 내역 펼쳐보기 ✅

### 고객 행동 푸시 알림 Vercel 서버리스 사일런트 실패 수정
- [x] **reviews/submit** — after() 래퍼로 리뷰 작성 푸시 실행 보장 ✅
- [x] **consultation/public/confirm** — 푸시를 기존 after() 블록 안으로 이동 ✅
- [x] **consultation/public/resched** — 푸시를 after() 블록 안으로 이동 ✅

### 푸시 알림 테스트 발송 패널
- [x] **POST /api/push/test 신규** — 9가지 타입별 preset 테스트 발송 ✅
- [x] **설정 UI — 테스트 발송 그리드** — 각 버튼 클릭 시 실제 기기로 [테스트] 접두어 푸시 ✅
- [x] **진단 흐름 완비** — 기본 테스트(토글 무관) → 기기 연결 검증 → 타입별 토글 검증 ✅

### 푸시 알림 중복 표시 버그 수정 (tag 통일)
- [x] **원인 파악** — FCM SW 경로와 Realtime window.Notification 경로가 서로 다른 tag 사용 → 브라우저 dedup 실패 ✅
- [x] **send-push.ts** — FCM payload `data` 필드에 tag 추가 ✅
- [x] **firebase-messaging-sw.js** — `notification.tag` fallback 추가 ✅
- [x] **use-push-notifications.ts** — 하드코딩 'mamoru-push' 제거, `data.tag` 사용 ✅
- [x] **영향** — 리뷰뿐 아니라 모든 9가지 푸시 타입 자동 dedup ✅

### 상담 리마인더 템플릿 연결 오류 수정 (Make 설정)
- [x] **원인** — Make 시나리오 `FIELD_REMIND_24H` 모듈 내부 Solapi 템플릿이 '매장방문 리마인드' 내용으로 잘못 연결되어 있었음 ✅
- [x] **해결** — 사장님이 Make에서 올바른 '출장 리마인드' Solapi 템플릿으로 재연결 ✅
- [x] **시사점** — 알림톡 문제 발생 시 3곳 크로스 체크 원칙: ①TMS 코드 → ②Make 시나리오 → ③Solapi 템플릿 텍스트 ✅

### 메인 페이지(/main) 카피·디자인 리뉴얼
- [x] **히어로** — "그 기준을 정의한다" → "그 기준을 정의하다" (선언적 여운) ✅
- [x] **Trust 3번** — `0144 상술` 버그 → `ZERO 상술 · 권유` (중공형 외곽선 효과) ✅
- [x] **Trust 4번** — `자체진행 복원수리` → `자체진행 / 수냉식 복원수리` (compact 폰트 변형) ✅
- [x] **Trust 1번 라벨** — `대를 잇는 기술` → `대를 잇는 기술력` (축적 뉘앙스) ✅
- [x] **Quick Nav 2줄 분리** — 제품 카테고리 4개(미용가위/빗·브러시/가위집/정리용품) + 서비스 4개 ✅
- [x] **카테고리 URL 실제 매핑** — `/29` 공통 → `/44` `/sideproduct` `/37` `/38` (기존 제품보기 버그 해결) ✅
- [x] **슬로건 한글 카피** — `가짜를 잘라내고, 진짜를 지킨다` → `거짓 없는 본질` (4인 회의 채택) ✅
- [x] **철학 문장** — `판매하지 않는다. 안내할 뿐이다.` → `권유하지 않습니다. 안내할 뿐입니다.` (존댓말 통일) ✅
- [x] **복원수리 카드** — `첫 날의 커트감을` → `첫 만남의 그 느낌.` (개인 기억 소환) ✅
- [x] **컨설팅 카드** — `추천하지 않습니다` → `사용자가 기준이 되어야 합니다` (철학 명시) ✅
- [x] **도구·소모품 라벨** — `도구 · 소모품` → `주변제품` (아임웹 /45 카테고리명 통일) ✅

---

## ✅ 완료 — 04-22 작업

### 아임웹 배너/팝업 원격 관리 (Script API 구현 Phase 1)
- [x] **DB 마이그레이션 053** — imweb_banners 테이블 + Storage 버킷 ✅
- [x] **관리자 API 3종** — GET/PATCH banners + POST upload ✅
- [x] **공개 API 2종** — banner-config (JSON, CORS) + banner-widget.js (자기완결형) ✅
- [x] **Script API 자동 주입** — upsertMamoruWidget (GET→POST/PUT 분기) ✅
- [x] **TMS 설정 UI** — 이미지 업로드 + 토글 + 미리보기 + 자동/수동 설치 ✅
- [x] **MAMORU Brand Guide 준수 모달** — 모노크롬, Noto Sans KR, 여백 ✅
- [x] **쿠키 기반 "오늘 하루 보지 않기"** 동작 ✅
- [x] **노출 기간 필터** — starts_at / ends_at ✅
- [x] **아임웹 Footer Code 반영** — 공통 코드 footer_code.txt 업데이트 ✅
- [x] **실제 가동 확인** — 메인 페이지 모달 배너 정상 노출 ✅
- [x] **과거 종료일 경고 모달** (UX 개선) — 종료일 과거 입력 시 저장 전 경고 ✅

### 아임웹 배너 슬라이드 기능 (Phase 2)
- [x] **DB 마이그레이션 054** — images JSONB 컬럼 + 기존 데이터 자동 마이그레이션 ✅
- [x] **이미지 최대 5장** — 1장=정적, 2장+=자동 슬라이드 ✅
- [x] **5초 자동 전환** — 호버 시 일시정지, 수동 이동 시 타이머 리셋 ✅
- [x] **점(dot) 네비게이션** — 현재 슬라이드 시각적 강조 ✅
- [x] **데스크톱 좌우 화살표** — hover 시 표시 (pointer:fine 미디어 쿼리) ✅
- [x] **모바일 스와이프** — touchstart/touchend 40px 임계 ✅
- [x] **순환 루프** — 마지막 → 첫 번째 자동 ✅
- [x] **이미지별 개별 링크** — 슬라이드마다 고유 link_url 지원 ✅
- [x] **순서 변경** — ↑↓ 버튼 (DB 즉시 반영) ✅
- [x] **TMS 미리보기** — 실제 슬라이드 동작 시뮬레이션 ✅
- [x] **실제 가동 확인** — 다중 이미지 슬라이드 정상 노출 ✅

---

## ✅ 완료 — 04-21 작업

### Google Calendar 연동 (Phase 1 MVP)
- [x] **OAuth 2.0 Authorization Code Flow** — Workspace/Gmail 자동 판별 (id_token hd 필드) ✅
- [x] **상담 확정 자동 캘린더 등록** — 매장방문 즉시 / 출장 고객 시간 수락 ✅
- [x] **관리자 수동 일정 변경 → 캘린더 자동 UPDATE** — 같은 이벤트 ID 유지 ✅
- [x] **출장 확정건 "일정변경" 버튼 신설** — 매장방문과 동일 RescheduleModal 재사용 ✅
- [x] **재요청 상태 → ⏳ 접두어 + 노란색 업데이트** (이벤트 유지) ✅
- [x] **취소 / 보류 → 이벤트 삭제** ✅
- [x] **완료 → ✅ 접두어 + 회색 업데이트** (이력 보존) ✅
- [x] **suggested 상태는 캘린더 미표시** — 고객 미확정 후보 시간 제외 ✅
- [x] **설정 UI** — 연결 상태·재동기화·해제 버튼 + Workspace 태그 ✅
- [x] **Vercel `after()` 래퍼로 실시간 반영 보장** — 플로팅 Promise 문제 해결 ✅
- [x] **openid/email/profile 스코프 추가** — 연결 계정 자동 표시 ✅
- [x] **이벤트 색상 구분** — 매장 파랑 / 출장 녹색 / 재요청 노랑 / 완료 회색 ✅
- [x] **이벤트 설명 풍부화** — 고객명·연락처·주소·메모·TMS 링크·tel: 링크 ✅
- [x] **전체 재동기화 버튼** — 과거 60일 ~ 미래 180일 일괄 처리 ✅

### 복원수리 접수 MAKE 웹훅 버그 수정
- [x] **pickup_date 누락 버그 수정** — 방문수거 4종(문앞/카운터/직접전달) 모두 적용 ✅
- [x] **수거예정일 포맷 한글화** — `YYYY년 MM월 DD일 (X요일)` ✅
- [x] **TMS_FLOW_REPAIR.md 페이로드 필드표 업데이트** ✅

### 푸시 알림 회수 로직
- [x] **복원수리 삭제 시 OS 알림 자동 회수** — tag에 as_id 포함 + SW postMessage ✅
- [x] **Service Worker DISMISS 메시지 리스너 추가** ✅
- [x] **useDeleteRepair 성공 시 SW에 알림 회수 요청** ✅
- [x] **DELETE API에서 push_notifications 테이블 행 정리** ✅

---

## ✅ 완료 — 04-20 작업

### 관리자 푸시 알림 확장 (3개 → 8개)
- [x] **상담 접수 (매장/출장/톡)** — 기존 3개에 톡상담 추가 ✅
- [x] **복원수리 접수** — 기존 유지 ✅
- [x] **고객 리뷰 작성 ⭐** — 신규 (리뷰 submit 후 push) ✅
- [x] **아임웹 주문 접수 📦** — 신규 (sync 시 신규 주문만) ✅
- [x] **출장 일정 확정 ✅** — 신규 (고객이 시간 선택) ✅
- [x] **출장 일정 재요청 🔄** — 신규 (고객이 다른 시간 요청) ✅
- [x] **설정 UI 분리** — "📱 내 푸시 알림" + "💬 고객 알림톡 발송" 분리 ✅

### 네이버 리뷰 시스템 완성
- [x] **review.md 자동 입력 기능** — 파일 선택 → 작성자/날짜/본문 자동 파싱 ✅
- [x] **received_at 입력 필드 추가** — 상담일/접수일/구매일 (타입별 라벨) ✅
- [x] **product_group 자동 조회** — purchase 타입 제품명으로 매핑 ✅
- [x] **리뷰 API limit 50 → 200** — 네이버 리뷰 누락 방지 ✅
- [x] **page_reviews 초기 노출 6 → 12개** — 리뷰 풍부하게 ✅

### 버그 수정
- [x] **톡상담 취소 시 잘못된 알림톡 방지** — cancelled 템플릿(매장용) 발송 차단 ✅

---

## ✅ 완료 — 04-17 작업

### 발주 시스템 대폭 개선
- [x] **외화 발주 (USD/CNY)**: 통화 선택 + 환율 입력 → KRW 자동 환산 (선납/잔금 호환) ✅
- [x] **발주 부가세 유형 변경**: 확정 후에도 포함/별도/미적용 변경 가능 (total 자동 재계산) ✅
- [x] **발주 품목 수량 직접 입력**: 숫자 직접 타이핑 가능 (버튼 증감 병행) ✅
- [x] **발주 제품 검색 필터**: 제품명/SKU/제품군 실시간 검색 ✅
- [x] **발주 카탈로그 특징 표시**: 매입품목 features를 제품 카드에 표시 (위안화 원가 확인용) ✅

### B2B 거래처 기능
- [x] **수금 처리 모달**: 미결제 납품+판매 체크 선택 → 일괄 수금 (대시보드 연동) ✅
- [x] **매입품목 → 제품 매입처 자동 연결**: 카탈로그 추가 시 products.supplier_id 자동 설정 ✅

### 제품 관리 개선
- [x] **제품군 필터 칩 탭**: 카테고리 아래 제품군(A1/A2/R1) 2차 필터 ✅
- [x] **정렬/그룹핑 정합성 수정**: Map 기반 그룹핑 + 설정값 연동 (제품군순=그룹핑, 기타=플랫) ✅
- [x] **전체 탭 그룹핑 개선**: 제품군만 묶기 (카테고리 혼합 플랫) ✅

### 문서/출력 개선
- [x] **거래명세서 + 납품서 이미지 복사**: 클립보드 PNG 복사 버튼 ✅
- [x] **B2B 납품서 품목명 인라인 편집**: 미리보기 수정 + 납품명 자동 적용 ✅

---

## ✅ 완료 — 04-16 작업

- [x] **판매 출고완료 알림톡**: 송장생성/출고완료 2단계 분리 + 선택적 알림톡 발송 (Make+솔라피 연동) ✅
- [x] **제품 일괄 수정**: 테이블 형태 편집 (가격/카테고리/발주명/제품군/순서 한번에) ✅
- [x] **발주서 출력 특징 컬럼**: 매입품목 카탈로그 features 자동 매핑 ✅
- [x] **매입품목 카탈로그 출력**: 거래처별 전체 품목 리스트 인쇄 (컬럼 선택 토글 + 주문명 fallback) ✅
- [x] **발주 작성 매입품목 필터**: 매입처 선택 시 "매입품목만" 토글 + 재고 수량 표시 + IW-XX 숨김 ✅
- [x] **제품 통합 정렬**: product_group(대분류) → category(중분류) → sort_order → name ✅
- [x] **제품 화면 그룹핑**: 대분류(제품군) 헤더 + 중분류(카테고리) 라벨 2단 그룹핑 ✅
- [x] **거래명세서 품목명 편집**: 미리보기에서 인라인 수정 + 납품명 자동 적용 ✅
- [x] **매입관리 취소 건 삭제**: 취소 상태 발주 영구 삭제 가능 ✅
- [x] **정렬 설정 연동**: 설정 → 상품 정렬 기본값에 "제품군→카테고리→순서→이름" 옵션 ✅
- [x] **제품 수정 에러 수정**: supplier_id FK 위반 방지 + 에러 직렬화 개선 ✅
- [x] **API purchase_name 누락 수정**: PATCH allowed 배열에 purchase_name 추가 ✅

---

## 📌 아임웹 새 OpenAPI 활용 계획 (04-15 회의 결과)

> 새 OpenAPI(`openapi.imweb.me`) OAuth2 연동 완료. Script + Order + Product scope 활성화.

### 도메인 이전
- [ ] `mamoru.kr` 카페24 → 아임웹 또는 가비아로 도메인 이전
  - 카페24 도메인 만료일 확인 (만료 전 이전 필수)
  - 아임웹이 .kr 이전 지원하는지 확인 → 안 되면 가비아로
  - 이전 후 DNS에서 아임웹 쇼핑몰 + Vercel TMS 연결
  - TMS 서브도메인 검토 (예: `tms.mamoru.kr`)

### 우선순위 1 — 즉시 효과 (난이도 낮음)
- [ ] **주문 취소 → 아임웹 자동 반영**: TMS에서 취소 시 아임웹도 자동 취소 (현재 수동)
  - API: `주문 취소 접수 요청 (PATCH)`, `주문 섹션 취소 접수 요청`
- [ ] **송장 수정/삭제**: 잘못된 송장 입력 시 삭제 후 재등록 (현재 불가)
  - API: `주문 송장 수정 (PATCH)`, `주문 송장 삭제 (DEL)`
- [ ] **배송완료 → 아임웹 자동 전환**: 롯데택배 배송완료 감지 시 아임웹도 자동 상태 전환
  - API: `주문 배송 처리 (PATCH)`
  - 기존 크론(`track-delivery`)에 추가

### 우선순위 2 — 비즈니스 확장 (난이도 중간)
- [ ] **TMS→아임웹 상품 자동 등록**: TMS에서 신상품 등록 시 아임웹 자동 생성 (이중 등록 제거)
  - API: `상품 상세 수정`, `상품 가격 설정 수정`, `상품 상태 수정`
- [ ] **회원 정보 동기화**: 아임웹 회원 = TMS 고객 자동 매핑
  - API: Member-Info 카테고리

### 우선순위 3 — 추후 검토
- [ ] **쿠폰/적립금 연동**: VIP 고객 자동 쿠폰 발급 등
  - API: Promotion 카테고리
- [ ] **상품 동기화 크론 자동화**: 현재 수동 버튼 → 하루 1회 자동

---

## ✅ 완료 — 04-14~15 작업

### 아임웹 재고 연동 (04-15 완성)
- [x] 아임웹 개발자 센터 앱 등록 (MAMORU TMS) + OAuth2 연동
- [x] 새 OpenAPI (`PATCH /products/{prodNo}/stock-info`) 재고 수정 구현
- [x] delta(증감값) 방식 전면 전환 — 8곳 호출처 모두 수정
- [x] OAuth 콜백 API + DB 토큰 저장 + refreshToken 자동 갱신
- [x] Vercel 환경변수 4개 설정 (v2 + OpenAPI)

### 문서 서식 + B2B거래 개선 (04-15)
- [x] 문서 경칭 추가 — 개인 고객 "님" + B2B 거래처 "귀중" (6개 문서)
- [x] 저재고 알림 — 설정값 동기화 + 하드코딩 3 제거 (설정값 실제 반영)
- [x] B2B거래 ALPS 송장 생성 — 납품 상세에서 롯데택배 연동
- [x] B2B거래 정산완료 버튼/탭 제거 — 출고완료+결제완료가 최종 상태
- [x] 아임웹 OAuth scope 확장 — Script + Order 권한 추가 (재인증 완료)

### Vercel 빌드 분리 (04-14)
- [x] `vercel.json` ignoreCommand 설정 — TMS 변경 시에만 빌드, 페이지 push 스킵

### 리뷰 시스템 (04-14)
- [x] repair 후기 제출 불가 수정 (submit/info API offline_sales fallback)
- [x] subtype 전달 수정 (page_review.html submit body에 urlSubtype 추가)
- [x] 칩 표시 통일 (상담·출장, 복원수리+방문수거)
- [x] 즉시노출 ON/OFF 설정 저장 버그 수정 (API 형식 불일치)
- [x] 네이버 리뷰 등록 — 상담 방식 하위 선택 추가
- [x] 고객 후기 날짜 라벨 (상담/접수/구매) + YYMMDD 포맷
- [x] 기존 리뷰 subtype 마이그레이션 (offline→field_request/store_visit)

### B2B거래 (04-14)
- [x] 납품관리 → B2B거래 명칭 변경
- [x] 복원수리 납품 즉시 confirmed (대시보드 즉시 반영)
- [x] B2B수리 버튼 분리 (납품서 작성 옆)
- [x] 추가항목 직접 입력 + 배송비 원클릭 버튼
- [x] 모달 드래그 닫힘 방지
- [x] 거래처 업체명(company_name) 우선 표시 + 검색

### 기타 (04-14)
- [x] TypeScript 빌드 에러 근본 수정 (CustomerResult 타입 + unsafe 캐스팅 제거)
- [x] CLAUDE.md — tsc 체크 필수 + unsafe 캐스팅 금지 규칙 추가
- [x] 상담 가이드 다크카드 3종 분리

---

## ✅ 완료 — 롯데택배 ALPS 답변 수신 (04-13)

> 답변 요약: 동일 운송장 재사용 금지 (TMS 순차채번으로 자동 방지), SM APP 오류는 롯데 측 조치 완료

- [x] **1. 주소 불일치 원인**: 동일 운송장 취소→재접수 시 복수 데이터 섞임 → TMS 순차채번으로 방지됨 ✅
- [x] **2. 운송장 대역 관리**: 거래처 자체 관리 필요 → TMS DB(lotte_waybill_config)에서 관리 + 잔여 UI 구현 ✅
- [x] **3. SM APP 오류**: 대리점 기사 앱 오류 → 롯데 측 조치 요청 완료 ✅
- [x] **4. 테스트 방법**: 미사용 번호로 테스트 → TMS 순차채번이라 자동 해결 ✅

---

## 📌 내일 할 것 (04-11)

### 아임웹 반영 (사장님 수동)
- [ ] 카테고리바 v2.0 각 카테고리 페이지 코드위젯 교체 (5개 — active+배너 텍스트만 변경)
- [ ] 후기 페이지(/62) 코드위젯 → iframe_reviews.html 교체
- [ ] 아임웹 Header Code 상단 교체 (header_code_top.txt)
- [ ] 아임웹 Footer Code 교체 (footer_code.txt)
- [ ] 메인 이미지 삽입 — Before/After, 컨설팅, 도구 카드 (촬영분)

### 메인 페이지
- [ ] 아임웹 배치 최종 확인 — top → 이름표 → 기획전 → btm 순서
- [ ] 이름표 코드위젯 섹션 배경색 #FAF9F7 설정
- [ ] 모바일 퀵네비 스크롤 힌트 확인

### TMS
- [ ] 대시보드 할일 메모 테스트
- [ ] B2B 거래처 상세 테스트
- [ ] 매입관리 제품 검색 리스트 테스트
- [x] 준비표 인쇄 시리얼 표시 수정 (sale_item_id 매핑) ✅ 04-10
- [ ] 남은 하드코딩 교체 (수리비, 저재고 기준, 체류 경고일)
- [ ] 거래명세서 사업자 정보 → 설정 연동
- [ ] 사이드바 커스텀 → 설정값 연동

### 대기
- [ ] 롯데택배 답변 확인 시 코드 반영
- [ ] Vercel 빌드 비용 모니터링 (Ignored Build Step 정상 동작 확인)

---

## 📌 미완료 — 설정 탭 리뉴얼 (04-03 승인)

### Phase 0~2: 인프라 + UI + 96개 항목 ✅
- [x] system_settings 테이블 + warehouses + 반품 컬럼 (050)
- [x] GET/PATCH /api/settings API + useSettings 훅
- [x] 설정 페이지 10탭 전면 재작성

### Phase 3: 신규 기능
- [x] 3-A: 반품 기능 (API + Hook + 판매상세 반품버튼) ✅
- [x] 3-D: 푸시 알림 (FCM Admin SDK + Supabase Realtime 이중 구조) ✅
- [ ] 3-B: 검수 다중선택 UI (radio→checkbox) — 별도 작업
- [ ] 3-C: 배송추적 Cron 디버깅 — 별도 작업
- [ ] 3-E: 다중 창고 재고 이동 UI — DB 준비됨, UI는 별도

### Phase 4: 하드코딩 교체 (진행중)
- [x] 대시보드 KPI: localStorage→DB, 색상 설정 연동 ✅
- [x] 알림톡 마스터/개별 on/off DB 체크 ✅
- [x] 대시보드 카드 순서/숨김 설정 연동 ✅
- [ ] cost-calculator 수리비/배송비 → 설정값 교체
- [ ] 저재고 기준/체류 경고일 → 설정값 교체
- [ ] 사이드바 커스텀 → 설정값 연동

### 별도 작업
- [ ] 복원수리 고객 상태 페이지 (진행현황 타임라인)

---

## 📌 미완료 — 시뮬레이션 테스트

### 시나리오 A: 매장방문 상담 → 판매 → 계약서
- [x] 1. 고객이 매장방문 접수 (page_form.html) ✅
- [x] 2. 알림톡(confirmed) 수신 확인 ✅
- [ ] 3. TMS 상담관리 → 해당 건 확인
- [x] 4. 대면 상담 후 "상담 완료" 처리 ✅
- [x] 5. 리뷰 요청 알림톡 발송 확인 ✅
- [ ] 6. 계약서 작성 — 개편 필요 (입력방식/내용 수정)
- [ ] 7. 고객 서명 + 이미지 저장
- [ ] 8. 계약서 → 판매 전환 (contract_id 연결 + 계약내용 표시 필요)
- [ ] 9. 제품 선택 + 시리얼 선택 + 결제 → 판매 등록
- [ ] 10. 시리얼 sold 전환 + 재고 차감 확인

### 시나리오 B: 출장상담 → 판매
- [ ] 1~7 전체 미테스트

### 시나리오 C: 톡상담 → 판매
- [ ] 1~4 전체 미테스트

### 시나리오 D: B2B 딜러 납품 (납품관리 모듈 구축 완료)
- [x] 납품관리 모듈 구축 (DB+API+Hooks+UI) ✅ 04-12
- [ ] 1. 납품서 작성 → B2B 고객 선택 → 딜러가 적용 확인
- [ ] 2. 납품 확정 → 재고 차감 확인
- [ ] 3. 출고 완료 → 송장번호 기록
- [ ] 4. 정산 완료 → 미수금 감소 확인
- [ ] 5. 납품서 인쇄 확인
- [ ] 6. 취소 → 재고 복원 + 미수금 차감 확인

### 시나리오 E: 복원수리 전체 흐름
- [ ] 1. 고객 접수 → 알림톡(as_received)
- [ ] 2. TMS 입고 & 비용안내 → 알림톡(as_cost_notice)
- [ ] 3. 입금확인 → 알림톡(as_payment_confirmed)
- [ ] 4. 작업 완료 → 출고대기 → 송장생성 (ALPS)
- [ ] 5. 출고완료 → 알림톡(as_shipped)
- [ ] 6. 배송완료 → 리뷰 요청 알림톡

### 시나리오 F: 삭제/취소 테스트
- [ ] 1~5 전체 미테스트

---

## 📌 미완료 — 사용자 수동 작업

- [ ] Gmail 앱 비밀번호 생성 → Vercel GMAIL_USER + GMAIL_APP_PASSWORD 설정
- [ ] TMS에서 제품별 product_group 설정 (R4-58ST → "R4" 등)
- [ ] 아임웹 디자인 모드 → 상품 상세 하단에 리뷰 코드위젯 삽입 (1회)
- [ ] 네이버 리뷰 160개 CSV 일괄 등록 (수동)
- [ ] 베스트 리뷰 5~10개 선정 + 사진 첨부 (수동)
- [ ] GAS 비활성화 — 2주 병렬 운영 후 배포 보관처리

---

## 📌 미완료 — 기능 구현

### UI 동작 검증 (수동)
- [ ] /sales/new 판매입력 (딜러→딜러가, 아카데미→아카데미가 자동 적용)
- [ ] /contracts/new 전자문서 (제품 모달 + 서명 2개 + 수령방법 + 선납/잔금)
- [ ] /products/[id]/serials 시리얼 등록
- [ ] /customers 고객 목록/상세 (매입처 필터 제거 확인)
- [ ] /products/new 제품 등록 (4단 가격: 소매/딜러/아카데미/매입)
- [ ] /suppliers B2B 거래처 (딜러/아카데미/매입처 서브탭)
- [ ] /purchasing/new 발주 작성 → 입고 → 재고 증가
- [ ] /inventory 재고 현황 + 저재고 필터
- [ ] /reports 회계 리포트 + 엑셀 다운로드
- [ ] /reports/transaction 거래내역서 인쇄

### 상담관리 Phase 2
- [ ] 일괄 시간 제안 (복수 고객 → 공통 가능요일 → 한번에 제안)
- [ ] 지도 MarkerClusterer (지역별 클러스터링)
- [ ] "출장 계획" 뷰 (날짜 선택 → 해당 요일 가능 고객 + 동선 최적화)
- [ ] 고객 다른 일정 요청 → TMS 재요청 반영 테스트

### 판매관리
- [ ] 딜러 납품 시 딜러가 자동 적용 확인 + 제품별 납품가 등록 필수화
- [ ] 판매 모달에서 직접 계약서 작성 CTA

### 복원수리
- [ ] TMS 송장생성 (ALPS boxTypCd 수정 완료) → 정상 발급 재테스트
- [ ] 비용안내 → 알림톡 + 입금확인 → 출고 → 전체 E2E
- [ ] 복원수리: 주소 수정 시 다음 주소검색 API 연동
- [ ] 복원수리: 사진 마킹 (photo-marker.tsx — html2canvas)
- [x] 판매 수정 시 시리얼 추가/삭제/변경 (소급 연결 포함) ✅ 04-11
- [x] 임시제품에도 시리얼 입력 가능 ✅ 04-11
- [x] 판매 수정 시 sale_item_id 매핑 순서 보장 (insert().select()) ✅ 04-11

### B2B 거래처
- [x] 매입품목 카탈로그 (주문명/특징 + 제품 불러오기) ✅ 04-11
- [x] 부가세 3유형 (포함/별도/미적용) ✅ 04-11
- [x] 발주서 인쇄 (주문품목+단가+수량+부가세) ✅ 04-11
- [ ] 거래 조건 관리 UI (할인율, 결제 조건)
- [ ] 거래처별 판매 실적 요약

### 제품/재고/시리얼
- [ ] 이미지 배치 업로드
- [ ] 카테고리 관리 (현재 4개 고정)
- [ ] 다중 선택 재고 조정
- [ ] 재고 이동 이력 (zone 변경 로그)
- [ ] 바코드 스캔 기능 (카메라)
- [ ] 시리얼 조회 → 복원수리 접수 바로 연결
- [ ] QR 출력 프로그램 연동
- [x] 판매 후 시리얼 소급 연결 (판매 수정으로 해결) ✅ 04-11

### 매입/부자재
- [ ] 선납금/잔금 추적 UI 강화
- [ ] 입고 시 시리얼 자동 생성 옵션
- [ ] 부자재 재주문 알림

### 리뷰
- [x] 판매 상세→후기 요청 알림톡 발송 (review_url에 uid 포함) ✅ 04-09
- [x] 리뷰 info API — 판매 건(OS-*) offline_sales fallback 조회 ✅ 04-09
- [x] 후기 webhook consult_uid/as_uid 변수 매핑 추가 ✅ 04-09
- [ ] 네이버 리뷰 160개 CSV 등록 (수동)
- [ ] 베스트 리뷰 선정 UI 개선
- [ ] 리뷰 승인/거절 모달
- [ ] 통합 리뷰 시스템 설계 (memory/REVIEW_SYSTEM_BRIEF.md)

### 회계
- [ ] 기간별 매출 추이 차트
- [ ] COGS 동적 계산 (실제 매입가 반영)

### 납품관리
- [x] fix: 납품 확정/취소 시 raw_stock 연동 (재고 불일치 해결) ✅ 04-12
- [x] style: 모달 제품검색 드롭다운 absolute + 섹션 그루핑 ✅ 04-12
- [x] feat: 임시 제품 직접 입력 + 부분결제 선납금 입력/표시 ✅ 04-12
- [x] fix: 납품서 수량 직접입력 가능 (±버튼 + input 병행) ✅ 04-13
- [x] fix: 기본값 변경 — 부가세 미적용 / 증빙 미적용 / 미결제 ✅ 04-13
- [x] feat: 운송장 잔여 표시 UI + 100건 미만 대시보드 알림 ✅ 04-13
- [x] feat: 신규 리뷰 등록 시 대시보드 알림 (pending 상태) ✅ 04-13
- [x] style: 설정→주문배송 송장번호 현황 (잔여+프로그레스바+번호범위) ✅ 04-13
- [x] feat: 복원수리 간편입력 모드 (거래처+수량+단가+결제) ✅ 04-13
- [x] fix: 대시보드 매출에 납품 금액 합산 (RPC 블록 내 추가 쿼리) ✅ 04-13
- [x] fix: 복원수리 대시보드 B2B 수량+금액+총매출에 납품 RS 항목 합산 ✅ 04-13
- [ ] 출고완료 시 알림톡 발송 기능 (솔라피 템플릿 등록)
- [ ] 대시보드 매출 합산 (deliveries 테이블 쿼리 추가)
- [ ] 회계 보고서 연동 (reports/summary에 deliveries 집계)

### 알림/연동
- [ ] E2E 테스트 (상담 17종 + 복원수리 6종 + 계약서 1종)
- [ ] 계약서 알림톡 템플릿 등록 (솔라피 검수)
- [ ] 동기화 상태 대시보드 표시 (마지막 시간, 에러 로그)
- [ ] 복원수리 입금 시 고객 미수금 연동
- [ ] 세금계산서 ↔ 판매/매입 건 연결 (FK)

### 고객 페이지 — 메인 리뉴얼 (04-06 완료)
- [x] 메인 페이지 전면 리뉴얼 (7섹션 → 2분할 구조) ✅
- [x] iframe 2분할 (top/btm) — 아임웹 상품 위젯 사이 배치 ✅
- [x] Trust Numbers 브랜드 무기형 (2대/100%/0/직접) + 카운터 애니메이션 ✅
- [x] Quick Nav 재방문 고객 바로가기 6칩 ✅
- [x] 클리어런스 배너 "LINEUP CHANGE" 톤 재설계 ✅
- [x] Service Showcase (복원수리+컨설팅) 통합 ✅
- [x] 후기: 카드 내 서브타입 칩 + 네이버 출처 배지 ✅
- [x] Brand Statement blur-to-sharp reveal ✅
- [x] 상품 위젯 호버 Terracotta→모노크롬(Sand) ✅
- [x] AS 리포트 페이지 Brand Guide 모노크롬 리뉴얼 ✅
- [x] 미사용 page_dealer_confirm.html 삭제 ✅
- [x] 고객 대면 페이지 Brand Guide v1.0 100% 통일 달성 ✅
- [ ] 메인 Before/After 이미지 삽입 (사장님 촬영 후)
- [ ] 메인 복원수리 영상 삽입 (촬영 후)
- [ ] 메인 컨설팅 분위기 이미지 삽입
- [ ] 메인 가위 프리뷰 — 아임웹 상품 위젯 정리 (카테고리/순서)
- [ ] 메인 도구·소모품 카드 이미지 삽입

### 고객 페이지 — Brand Guide v1.0 리뉴얼
- [x] `projects/as/page_as_report.html` — 모노크롬 리뉴얼 완료 ✅
- [x] `projects/consulting/page_dealer_confirm.html` — 미사용 → 삭제 ✅

### 향후
- [ ] 바코드 스캔 재고 입출고
- [ ] QR 출력 프로그램 연동 (라벨 프린터 전용 프로그램 확인 후)

---

---

## ✅ 완료 이력

<details>
<summary>04-11~12 작업 (클릭하여 펼치기)</summary>

### 판매 시리얼 수정
- [x] feat: 판매 수정 시 시리얼 추가/삭제/변경 기능 ✅
- [x] fix: 임시제품에도 시리얼 입력 가능하도록 변경 ✅
- [x] fix: 판매 수정 시 sale_item_id 매핑 순서 보장 (insert().select()) ✅
- [x] style: 판매조회 PC 좌측 패널 w-2/5 (기간 필터 줄바꿈 방지) ✅

### B2B 매입처 관리
- [x] feat: 매입품목 카탈로그 (supplier_product_catalog 테이블 + API + UI) ✅
- [x] feat: 부가세 3유형 — 포함/별도/미적용 (calcVAT 확장 + 발주 생성/상세) ✅
- [x] feat: 발주서 인쇄 (POPrintModal + 주문품목+부가세 연동) ✅
- [x] DB: 061 마이그레이션 (카탈로그 테이블 + vat_type + default_vat_type) ✅

### 납품관리 모듈 신설
- [x] DB: 062 마이그레이션 (deliveries + delivery_items) ✅
- [x] API: GET/POST /api/deliveries + GET/PATCH /api/deliveries/[id] ✅
- [x] 상태 흐름: draft→confirmed(재고차감)→shipped(출고)→settled(정산) ✅
- [x] 미수금 연동: 생성 시 증가, 정산/취소 시 차감 ✅
- [x] UI: 마스터-디테일 페이지 + 납품서 작성 모달 + 상세 패널 ✅
- [x] 납품서 인쇄 (DLPrintModal) ✅
- [x] 부가세 3유형 + 증빙유형 + 결제상태 ✅
- [x] 사이드바 '판매' 그룹에 '납품관리' 추가 ✅

### 고객 대면 페이지 개선 (상담/복원수리/브랜드)
- [x] 상담 안내: 마모루 컨설팅 탭 리디자인 (사장님 원고 기반 8섹션) ✅
- [x] 상담 안내: 상담철칙 탭 제거 → 컨설팅 탭에 흡수 ✅
- [x] 상담 안내: 탭별 개선 (캐릭터 복원, 프로세스 이미지 placeholder, Tip 추가) ✅
- [x] 상담 안내: 중복 CTA 제거 + 히어로 CTA 제거 + 감성 메시지 색상 ✅
- [x] 상담 안내: 찾아오시는 길 네이버 지도 링크 추가 ✅
- [x] 복원수리 안내: PREVIEW 빈 영상 → '영상 준비중' placeholder (4:3) ✅
- [x] 복원수리 안내: 소요시간 제목 줄바꿈 개선 ✅
- [x] 브랜드 인트로: 모바일 퀵네비 칩 추가 (PC 숨김) ✅
- [x] 브랜드 인트로: 오시는길 네이버 지도 링크 교체 ✅
- [x] CLAUDE.md: 4인 전문가 역할 지침 강화 (시각적 계층 전략 + 중복CTA 금지 + 전체 영역 인식) ✅

### TMS 알림 + 설정 + 납품 연동
- [x] feat: 운송장 잔여 표시 API + 설정 UI (프로그레스바+번호범위) ✅
- [x] feat: 100건 미만 시 대시보드 경고 카드 ✅
- [x] feat: 신규 리뷰(pending) 대시보드 알림 카드 ✅
- [x] fix: 납품서 수량 직접입력 + 기본값 미적용/미적용/미결제 ✅
- [x] feat: 복원수리 간편입력 모드 (거래처+수량+단가 8,000원+결제상태) ✅
- [x] fix: 대시보드 매출에 납품 금액 합산 — RPC 블록 내 추가 쿼리 ✅
- [x] fix: 복원수리 대시보드 B2B 수량+금액+월간 총매출에 납품 RS 합산 ✅
- [x] feat: 리뷰 삭제 기능 + 자동노출 토글 + 3열 레이아웃 + 칩 개편 ✅
- [x] feat: 후기요청 모달 복원수리 subtype 추가 (직접방문/방문수거) ✅

### 창고 이동 모달 개편
- [x] feat: '시리얼 등록·창고배치' → '창고 이동' (출발지→도착지 패턴) ✅
- [x] feat: 보관→준비/디스플레이 시 시리얼 자동생성, 역이동 시 시리얼 삭제+raw_stock 복원 ✅
- [x] feat: POST /api/serials/move API 신설 (zone 이동 + 역이동) ✅
- [x] style: 불필요 버튼 제거 ('창고·재고', '판매 등록') ✅

### 납품관리 버그/개선 6건
- [x] fix: 납품 확정/취소 시 raw_stock 차감/복원 (재고 불일치 해결) ✅
- [x] style: 제품 검색 드롭다운 absolute (모달 높이 흔들림 방지) ✅
- [x] feat: 임시 제품 직접 입력 (품목명+금액) ✅
- [x] style: 부가세/증빙/결제 섹션 border 그루핑 ✅
- [x] feat: 부분결제 선납금 입력 + 목록 "X원 선납 / Y원" 표시 ✅
- [x] audit: B2C 판매 취소/반품 재고 복구 — 코드 검증 완료 (수정 불필요) ✅

### 상담 안내 페이지
- [x] style: 히어로 리뉴얼 — Chosen 가시성 + fade-in + 배경 gradient ✅
- [x] style: 상담철칙 디테일 — 넘버링+소제목+부연설명 + PC 2열 ✅
- [x] feat: '마모루 컨설팅' 탭 추가 — 시장 현실+약속+기술력+이미지 ✅

</details>

<details>
<summary>04-10 작업 (클릭하여 펼치기)</summary>

### 고객 관리 개선
- [x] fix: 고객 등록 시 주소/매장명/유형 저장 안 되던 버그 수정 (API snake_case 불일치) ✅
- [x] style: 프로필에 매장명·주소 항상 표시, 이메일 하단 이동 ✅
- [x] feat: 고객 태그 시스템 활성화 (DB 060 + 등록/상세/목록/판매 전체 연동 + 필터) ✅

### 시리얼 + 판매
- [x] fix: 준비표 시리얼 미표시 — 기존 등록 시리얼 sale_item_id 누락 수정 ✅
- [x] style: 시리얼 피커 UX 개선 — 재고 0이면 자동 직접입력 모드 + 자동생성 버튼 크게 ✅
- [x] fix: 판매조회 전체 탭에서 취소 건 제외 (취소 탭에서만 표시) ✅

### 후기 요청
- [x] fix: 리뷰 submit API에 판매 건(OS-*) offline_sales fallback 추가 ✅
- [x] docs: 판매→후기 요청 흐름 완료 반영 (흐름도+로드맵+TODO) ✅

</details>

<details>
<summary>04-07 작업 (클릭하여 펼치기)</summary>

### 메인 페이지 리뉴얼
- [x] body 배경 검정→크림 (카테고리 칩 아래 검정 제거) ✅
- [x] iframe 높이 계산 개선 (scrollHeight→마지막 섹션 기준) ✅
- [x] top/btm iframe 메시지 ID 분리 (높이 충돌 방지) ✅
- [x] 대표라인업→첫구매맞춤가위 + 평생쓰는라인업 이름표 분리 ✅
- [x] 이름표 별도 HTML+iframe 파일 (label_first, label_lineup) ✅
- [x] mm-products 하단패딩 축소 + fallback 높이 0px ✅
- [x] 모바일 세로 스크롤 방지 (quicknav + 도구카드 overflow-y:hidden) ✅
- [x] 퀵네비 모바일 우측 페이드+화살표(›) 스크롤 힌트 ✅
- [x] 이름표 모바일 간격 축소 (48→24px) ✅
- [x] 쇼핑기획전 슬라이드 공통 디자인 코드 (상품위젯과 별도) ✅
- [x] 카테고리바 Brand Guide v2.0 모노크롬 리디자인 ✅

### 비용 절감
- [x] Vercel Ignored Build Step 설정 — TMS 미변경 시 빌드 스킵 ✅

</details>

<details>
<summary>04-10 작업 (클릭하여 펼치기)</summary>

### 후기 페이지 리뉴얼
- [x] 모바일 2열 Masonry 레이아웃 (잡지 스타일, 가로 우선 배치) ✅
- [x] PC 4열 / 태블릿 2열 / 모바일 2열 반응형 ✅
- [x] 네이버 출처 연녹색 칩 (카드+모달 동일) ✅
- [x] 이미지 비율 16:9→4:5 세로형 (모바일) ✅
- [x] 텍스트 3줄 clamp + 미리보기 50자 ✅
- [x] 탭 버튼 줄바꿈 방지 (nowrap + 가로 스크롤) ✅
- [x] 모달 iframe 뷰포트 센터링 (MAMORU_REQUEST_VIEWPORT) ✅
- [x] iframe_reviews.html 래퍼 생성 + ImwebWidgetCode_reviews.html 삭제 ✅

### 전 페이지 공통
- [x] -webkit-tap-highlight-color: transparent 전 page_*.html 일괄 적용 ✅
- [x] ADDENDUM_IMWEB.md 전면 동기화 (GAS→Vercel, 파일명 통일, 폰트 하한선, tap-highlight 규칙) ✅

### 정리
- [x] GAS 시절 구버전 4개 삭제 (consulting/_gas/) ✅

</details>

<details>
<summary>04-06 작업 (클릭하여 펼치기)</summary>

### 쇼핑몰 메인 페이지 대규모 리뉴얼
- [x] page_main.html → page_main_top.html + page_main_btm.html 2분할 ✅
- [x] iframe 래퍼 2개 (iframe_main_top/btm.html) — 아임웹 상품 위젯 사이 배치 ✅
- [x] Trust Numbers: 브랜드 무기형 (2대/100%/0/직접) + 카운터 애니메이션 ✅
- [x] Quick Nav: 재방문 고객 바로가기 6칩 ("제품 보기" 등) ✅
- [x] 클리어런스 배너: "LINEUP CHANGE" 톤 (display:none 가이드 포함) ✅
- [x] Service Showcase: 복원수리(Before/After+영상 플레이스) + 컨설팅 통합 ✅
- [x] 후기: 카드 내 서브타입 칩(직접방문/출장/톡 등) + 네이버 출처 배지 ✅
- [x] Brand Statement: blur-to-sharp reveal + "가짜를 잘라내고, 진짜를 지킨다" ✅
- [x] Marquee/Finder/Dual CTA 제거 → 브랜드 쇼룸 구조 전환 ✅
- [x] 상품 위젯 호버: Terracotta(#D4613E) → 모노크롬 Sand(#D4D0CB) ✅
- [x] Ken Burns 히어로 이미지 줌 효과 ✅
- [x] 이미지/영상 플레이스 15+ 슬롯 확보 ✅

### 출고 준비표 인쇄
- [x] PrepSheetModal 컴포넌트 (테이블형 — 고객명/제품명(시리얼)/수량/메모) ✅
- [x] 판매 조회: "준비표 뽑기" 토글 → 체크박스 복수 선택 → 인쇄 ✅
- [x] 판매 상세: 거래명세서/준비표/수정 3버튼 — 단건 인쇄 ✅
- [x] 메모란 인쇄 전 수정 가능 ✅

### 제품 페이지 개선
- [x] 제품 카드 슬림화 (3행→2행, SKU 제거, 3열 유지) ✅
- [x] 재고 미사용 제품: 재고 자리 빈칸 (0 미표시) ✅
- [x] 재고 관리 토글 (ON/OFF → stock_quantity -1 ↔ 0) ✅
- [x] 카테고리 탭 표시 설정 (설정에서 체크/해제) ✅
- [x] SUP(부자재) 제품 페이지에서 숨김 ✅

### 부자재 페이지
- [x] 1열 리스트 → 3열 카드 그리드 ✅
- [x] 카드 클릭 → 수정 모달 ✅
- [x] 상세설명 2줄 말줄임 ✅

### 판매 입력 3열 레이아웃
- [x] B안: [제품선택] [고객+결제] [장바구니+판매등록] ✅
- [x] 검색바 아래 라인 상단 정렬 ✅
- [x] 고객 메모 고정 영역 (min-h 48px) ✅
- [x] 고객 검색 시 memo 필드 연동 (API+인터페이스 수정) ✅
- [x] 제품 테이블 SKU 컬럼 제거 + 높이 확장 ✅

### B2B 거래처
- [x] 마스터-디테일 2열 (PC: 좌 목록 + 우 상세 패널) ✅
- [x] 클릭 → 회사정보/담당자/메모/미수금 표시 ✅
- [x] 인라인 편집 (수정 버튼) ✅

### 대시보드 할일 메모
- [x] dashboard_todos 테이블 + API (GET/POST/DELETE) ✅
- [x] 할일 위젯 (입력+체크+완료확인모달+삭제) ✅
- [x] 설정 카드 표시/숨김 연동 ✅

### 매입관리
- [x] 편집 모드 제품 추가: 칩 12개 → 검색+스크롤 리스트 ✅

### 하드코딩 교체
- [x] 카테고리 라벨: 제품목록/수정/재고 → useSetting 연동 ✅
- [x] 배송비: 방문수거 1개 5000→6000원 + 직접발송 추가 ✅
- [x] 경비 카테고리: DEFAULT_EXPENSE_CATEGORIES 통합 ✅
- [x] setting-defaults.ts: 공유 기본값 중앙 관리 ✅
- [x] 고객명/연락처 수정 → 판매 건 자동 동기화 ✅

</details>

<details>
<summary>04-05 작업 (클릭하여 펼치기)</summary>

### 판매관리 대규모 리모델링
- [x] PC 마스터-디테일 2컬럼 (좌480px 목록 + 우 상세패널) ✅
- [x] 기본 탭 today (전체내역 부담 제거) ✅
- [x] SaleDetailPanel 추출 (모달→인라인 재사용) + 메모 인라인편집 ✅
- [x] 모바일 SlidePanel 전환 ✅
- [x] 결제상태 3버튼 (결제완료/부분결제/미결제) + 입금액 입력 ✅

### 시리얼 & 정품인증 시스템
- [x] 시리얼 포맷: 순차 8자리 숫자 (이지캐드 연번 호환) ✅
- [x] verify_token 해시 토큰 (QR에 시리얼 노출 방지) ✅
- [x] 정품확인 공개 API (/api/verify/[token]) ✅
- [x] 정품확인 페이지 (projects/verify/index.html, 브랜드 가이드) ✅
- [x] 시리얼 일괄생성 UI — 시작번호 입력 + 생성 범위 미리보기 ✅
- [x] 시리얼 중복 사전 체크 + 구체적 에러 안내 ✅
- [x] display zone 시리얼 판매 허용 + SerialPicker [진열] 배지 ✅
- [x] previous_zone 저장/복원 (판매취소 시 원래 zone 복원) ✅
- [x] DB: 033_serial_previous_zone + 034_serial_verify_token ✅

### 제품 등록 개선
- [x] 제품 등록: 풀페이지 → 우측 패널 create 모드 (페이지 전환 없음) ✅
- [x] 복제 기능: 상세에서 복제 버튼 → duplicate 모드 (이름만 비움) ✅
- [x] SKU 자동 채번 (/api/products/next-sku, 카테고리별 마지막+1) ✅
- [x] 카테고리 변경 시 SKU 자동 재채번 ✅

### 레이아웃 개선
- [x] 제품·창고·판매 — 사이드바 높이 맞춤 (고정높이→flex-1 min-h-0) ✅
- [x] 제품 관리 — 상단 고정 + 하단 내부 스크롤 (여백 스크롤 제거) ✅

### 알림톡 & 푸시 수정
- [x] 알림톡 웹훅 DB 우선 + 환경변수 fallback (설정 페이지 연동) ✅
- [x] 알림톡 웹훅 3분기 — 상담/AS접수/AS상태변경 별도 시나리오 분리 ✅
- [x] 복원수리 접수 알림톡 — 주소 + pickup_address_text 웹훅 payload 누락 수정 ✅
- [x] 푸시 알림 — 웹훅 성공과 무관하게 독립 발송으로 변경 ✅
- [x] FCM 토큰 등록 흐름 연결 (requestPushToken → /api/push/subscribe) ✅
- [x] Firebase 서버 환경변수 3개 Vercel 등록 (PROJECT_ID/CLIENT_EMAIL/PRIVATE_KEY) ✅

### 기타
- [x] 노트북 git user → bsm-pixel 통일 (Vercel 배포 트리거 정상화) ✅
- [x] CLAUDE.md — 배포 프로토콜 추가 (GitHub Pages 즉시 / Vercel 확인 후) ✅
- [x] xlsx + firebase-admin + nodemailer 패키지 의존성 추가 ✅

</details>

<details>
<summary>04-04 작업 (클릭하여 펼치기)</summary>

### 설정 탭 리뉴얼 (Phase 0~4)
- [x] DB: system_settings + warehouses + push_subscriptions + push_notifications 테이블 ✅
- [x] API: GET/PATCH /api/settings + /api/push/subscribe ✅
- [x] Hook: useSettings + useSetting + useUpdateSettings + usePushNotifications ✅
- [x] 설정 페이지 10탭 전면 재작성 (대시보드/주문배송/상담/복원수리/판매/고객/상품재고/회계/알림연동/시스템) ✅
- [x] 대시보드 카드 순서/숨김 — 2열 그리드 프리뷰 + 드래그 순서변경 + 실제 반영 ✅
- [x] KPI 목표 localStorage→DB 이관 + 색상 임계치 설정 연동 ✅
- [x] 알림톡 마스터/개별 on/off DB 체크 (make-webhook.ts) ✅

### ALPS 송장 생성 수정
- [x] alps-client.ts 완전 재작성 — GAS 로직 100% 재현 ✅
- [x] 성공 판정: rtnCd === '0000' → 'S' (rtn_list[0]) ✅
- [x] 헤더: Accept + X-Idempotency-Key + X-Correlation-Id + charset=utf-8 ✅
- [x] ordNo + pickReqYmd 필드 추가 ✅
- [x] 환경변수 .trim() 전체 적용 ✅
- [x] 환경변수 LOTTE_JOB_CUST_CD || LOTTE_JOBCUSTCD 양쪽 읽기 ✅
- [x] ordSct: '1' (일반출고) 적용 ✅
- [x] 복원수리 상품명: '[MAMORU] 복원수리' ✅
- [x] 발송인/수취인 빈값 검증 + 명확한 에러 메시지 ✅

### 반품 기능
- [x] API: PATCH /api/sales/[id] action:'return' — 재고/시리얼 복귀 ✅
- [x] Hook: useReturnSale ✅
- [x] UI: 판매상세 반품 처리 버튼 + 사유 입력 모달 ✅

### B2B 납품명
- [x] DB: products에 dealer_name, academy_name 컬럼 추가 (052) ✅
- [x] 제품 수정(PC 패널): 딜러 납품명 / 아카데미 납품명 입력 필드 ✅
- [x] 판매 입력: 고객 유형별 납품명 자동 표시 + 저장 ✅

### 거래명세서 개선
- [x] 품명에서 SKU(IW번호) 제거 ✅
- [x] 서명란 → 공급자(마모루 사업장) + 공급받는자(고객) 정보 ✅
- [x] 하단 MAMORU 로고 ✅

### 푸시 알림
- [x] Firebase 프로젝트 생성 + 서비스 계정 키 발급 ✅
- [x] FCM Admin SDK 서버 발송 (send-push.ts) ✅
- [x] Supabase Realtime 구독 (use-push-notifications.ts) — 알림음 + 브라우저 알림 ✅
- [x] Service Worker (firebase-messaging-sw.js) 백그라운드 알림 ✅
- [x] 상담접수/복원수리접수 시 자동 푸시 (make-webhook.ts 연동) ✅

### 기타
- [x] 고객명/연락처 수정 시 판매 건 자동 동기화 ✅
- [x] 대시보드 매출 0원 수정 — RPC에 sales + monthRepairAmount 추가 (051) ✅
- [x] 카테고리 표시명 편집 (코드 고정, 라벨만 수정 가능) ✅
- [x] 롯데택배 ALPS 담당자 문의 발송 (6건) ✅

</details>

<details>
<summary>04-03 작업 (클릭하여 펼치기)</summary>

### 판매 시리얼 + 송장 + 수정 강화
- [x] 시리얼 직접 입력 (판매 시 미등록 시리얼 자동 생성+sold) ✅
- [x] 직접입력 품목에도 시리얼 입력 가능 ✅
- [x] 시리얼 자동 번호 생성 (마지막 +1, 중복 방지) ✅
- [x] 시리얼→판매항목 정확 매칭 (sale_item_id 컬럼 + 3단계 매칭 로직) ✅
- [x] 판매→택배발송 송장생성 (ALPS 직접 호출 + 품목명 자동) ✅
- [x] ALPS 필드명 전체 수정 (rcvNm→acperNm 등 GAS 동일) ✅
- [x] 고객 등록 공통 모달 (다음 주소검색 + 전체 필드) ✅
- [x] 판매 입력에 '신규 고객 등록' 버튼 + 모달 연결 ✅
- [x] 고객 수정에 다음 주소검색 (우편번호+도로명) ✅
- [x] 판매 수정 시 기존 시리얼 보존 (치명적 버그 수정) ✅
- [x] 판매 수정 모달 — 결제상태/채널/메모/입금액 편집 추가 ✅
- [x] 판매 목록 금액: paid_amount→total_amount (미결제 0원 수정) ✅
- [x] 판매 조회 기본탭 전체 + 건수뱃지 축소 + 날짜범위 선택 ✅
- [x] DB: customers.postcode + offline_sales.invoice_number/shipped_at ✅
- [x] DB: product_serials.sale_item_id ✅

</details>

<details>
<summary>04-02 작업 (클릭하여 펼치기)</summary>

### 판매관리 IA 개편 + 기능 추가
- [x] P1-1: 사이드바 메뉴 분리 (판매 입력 + 판매 조회) ✅
- [x] P1-2: 판매 조회 통계 카드 (주간/월간/미수금) ✅
- [x] P1-3: 판매 입력 제품 선택 카드→테이블 목록/검색형 전환 ✅
- [x] 판매일 선택 추가 (과거 날짜 입력 가능) ✅
- [x] 거래명세서 A4 모달 (인쇄 + 이미지 저장 PNG) ✅
- [x] 판매 수정 — 금액/할인/결제방법/날짜 편집 ✅
- [x] 판매 수정 — 제품 추가/삭제 (내부 취소→재생성) ✅
- [x] 고객 상세 거래 요약 — 최근 거래일 + 취소건 제외 ✅
- [x] 수량 직접 입력 (판매 입력 + 수정 모달) ✅
- [x] 로그인 리다이렉트 무한 루프 수정 (태블릿) ✅

### 회계 + 복원수리 매출 통합
- [x] 복원수리 페이지 상단 매출 카드 (접수+판매 합산) ✅
- [x] 회계 API 복원수리 매출 집계 (paid_at 기준) ✅
- [x] 회계 UI 탭 바 [전체 매출/상품 판매/복원수리] ✅

### 로드맵 Tier 1~3 전체
- [x] 고객 통합 타임라인 ✅
- [x] 경비 등록 (DB+API+페이지) ✅
- [x] 재고 이력 API ✅
- [x] 매출채권 에이징 (30/60/90일) ✅
- [x] 월별 손익계산서 ✅
- [x] RFM 고객 분석 (VIP/일반/휴면) ✅
- [x] 입출금 관리 (캐시플로우) ✅
- [x] 고정 경비 자동 등록 ✅
- [x] 세금계산서 발행 관리 ✅
- [x] 제품 수명주기 분석 API ✅
- [x] 복원수리 작업 일지 ✅
- [x] 대시보드 KPI 게이지 ✅
- [x] B2B 거래처별 매출 ✅
- [x] 고객 퀵뷰 모달 ✅
- [x] 경비/입출금/세금계산서 PATCH API ✅
- [x] 프로세스 문서 3개 신규 + 2개 업데이트 ✅

</details>

<details>
<summary>04-01 작업 (클릭하여 펼치기)</summary>

### GAS→Vercel 전환 + E2E 테스트
- [x] 알림톡 Make 시나리오 ON + 솔라피 Rate Limit 확인 ✅
- [x] GITHUB_PAGES https:// 롤백 ✅
- [x] 출장 슬롯 차단: confirmed + suggested 버퍼 적용 ✅
- [x] 카카오 Geocoder 좌표 자동 세팅 ✅
- [x] 계약서 CTA: confirmed/in_progress + 톡상담 제외 ✅
- [x] ALPS boxTypCd 추가 ✅
- [x] 입금확인 skip_notify 분리 ✅
- [x] 상담/복원수리 삭제 기능 ✅
- [x] 리마인더 중복 발송 방지 + KST 타임존 ✅
- [x] 복원수리 주소 필드명 수정 ✅
- [x] 에러 typeof 체크 ✅
- [x] 복원수리 접수 Gmail 알림 ✅
- [x] 삭제 시 repair_photos 정리 ✅
- [x] 달력 suggested 표시 ✅
- [x] 판매관리 6건 수정 (시리얼잠금/계약연결/재고복원/에러/검증) ✅
- [x] 마스터 아키텍처 문서 생성 ✅

</details>

<details>
<summary>03월 작업 (클릭하여 펼치기)</summary>

### 03-31~04-01: GAS→Vercel 전면 이전
- [x] 상담/복원수리 Vercel API 직접 처리 ✅
- [x] ALPS 직접 호출 클라이언트 ✅
- [x] 리마인더 Cron ✅
- [x] 고객 페이지 JS 교체 ✅
- [x] Google Calendar 제거 ✅
- [x] Vercel Pro 업그레이드 ✅

### 03-29~30: TMS 전체 개선
- [x] ConfirmModal 24건 적용 ✅
- [x] PC 마스터-디테일 (주문/고객/계약/매입) ✅
- [x] 대시보드 강화 ✅
- [x] 재고 동기화 정합성 ✅
- [x] 복원수리 뱃지 필터 ✅

### 03-27~28: 판매/제품/시리얼 전면 개편
- [x] 판매 PC 마스터-디테일 ✅
- [x] 복합 결제 + VAT ✅
- [x] 시리얼 정품인증 ✅
- [x] 제품/창고 전면 리모델 ✅
- [x] 재고 모델 재정립 ✅

### 03-25~26: 상담/복원수리 IA 대규모 개선
- [x] 상담 6탭+3열+지도 통합 ✅
- [x] 복원수리 6탭 파이프라인 ✅
- [x] 계약서 필기 캔버스 ✅
- [x] 주문 슬라이드 패널 ✅

### 03-21~24: 기초 인프라
- [x] 아임웹 상품/재고 동기화 ✅
- [x] 이카운트 데이터 이관 ✅
- [x] 시리얼 Lifecycle ✅
- [x] 거래처 구조 재설계 ✅
- [x] 성능 최적화 ✅

### 03-16: 초기 구현
- [x] 기본 CRUD + 최적화 ✅

</details>

---

## 범례
- [x] 완료
- [ ] 미완료
- 📌 미완료 (상단), ✅ 완료 이력 (하단 접기)
