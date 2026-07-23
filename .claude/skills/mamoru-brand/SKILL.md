---
name: mamoru-brand
description: >-
  MAMORU 브랜드·디자인 기준(모노크롬 팔레트, 타이포, 톤, 시각적 위계)으로 고객 대면 페이지·시안·카피를
  만들거나 검토할 때 로드. Use when creating or reviewing any MAMORU customer-facing page, design
  mockup, hero, product detail, or copy — or when checking brand / tone / color / font / hierarchy
  / spacing compliance.
when_to_use: >-
  브랜드, 디자인, 시안, 페이지, 톤, 카피, 컬러, 팔레트, 폰트, 위계, 여백, 아임웹, 히어로, 상세페이지,
  brand check, design review, visual hierarchy, color palette, typography, spacing
---

# MAMORU 브랜드·디자인 기준 (작업/검토용 요약)

> 전체 기준은 `.claude/MAMORU-Complete-Brand-Guide-v1.0.md` (SSOT, 685줄).
> 이 파일은 **작업하거나 검토할 때 바로 쓰는 핵심 요약**이다. 애매하면 전체 가이드를 읽는다.

## 0. 한 문장
**"조용히 압도하는 전문가"** — 장식 없이, 여백과 위계로 압도한다. **CUT THE FAKE, KEEP THE REAL.**
판매하지 않는다. 안내할 뿐이다. (할인·구매권유 금지 = 브랜드 불변)

## 1. 컬러 — 모노크롬 (액센트 색 금지)
```css
--void:#1A1A1A  --graphite:#2D2D2D  --stone:#4A4A4A  --warm-gray:#8A8580
--mist:#B8B4AF  --sand:#D4D0CB  --parchment:#EDEBE8  --shell:#F5F3F0
--cream:#FAF9F7  --white:#FFFFFF
```
- 배경 = **Cream(#FAF9F7) 라이트 기본**, 다크(Void)는 오프닝/클로징·몰입 띠만.
- 🚫 **구 Terracotta/Gold 폐기** (v1.0에서 삭제). 유채색 액센트 쓰지 말 것.
- 상태색(성공/경고)은 기능상 최소한만(파랑/앰버), 브랜드 표면엔 안 씀.

## 2. 타이포
- 영문/숫자 = **Outfit**(700~900) · 본문 = **Plus Jakarta Sans + Noto Sans KR / Pretendard**.
- 한글·영문 폰트는 **span 분리**(스택 폴백 X). 상품 모델명은 Paperlogy/Pretendard 규칙 참조.
- 스케일: 모바일 본문 13~14, PC 16, 라벨 하한 11px. 플루이드 `clamp()` (제목 `clamp(30,9vw,52)`, 본문 `clamp(14,3.8vw,16)`). `<br>` 남용 금지.

## 3. 시각적 위계 (가장 중요 — "1초 테스트")
- 정보 가치 **1→2→3 순위**를 명확히. "모두 강조 = 아무것도 강조 안 함."
- **문제↔답 대비**: 문제는 흐리게·작게 / 답은 선명·크게·Bold. 3축 대비 + 거대 인덱스 + hairline.
- **여백 = 구조**(빈 공간이 아니라 그룹핑·시선유도 프레임). 요소 간격이 정보 위계와 일치.
- 화면을 1초 보고 눈 감았을 때 핵심 메시지/CTA가 잔상으로 남아야 함.

## 4. 톤 & 카피
- 말투: 먼저 묻고, 사실을 알려주고, 판단은 고객에게. 소리 안 지르는데 가장 무서운 사람.
- 피할 것: 장사꾼 냄새·촌스러움·올드함·장난기·시끄러움.
- 보조 카피: "돈이 안 아까운 곳, 시간이 아깝지 않은 곳" / "판매하지 않는다, 안내할 뿐이다."

## 5. 구조 패턴 (나열 금지 — 뼈대부터)
"트렌디 적용" 요청이 오면 스킨만 칠하지 말고 **콘텐츠 구조 재구성 먼저** 제안:
P1 문제→답 / P2 에디토리얼 인덱스 / P3 증거 우선. 모션은 순수 CSS 스크롤리빌(`animation-timeline:view()`), JS 0, 폴백.
- 아임웹 전 페이지 = **Page Kit 표준**(`projects/_design_lab/_page_kit.html`): 풀블리드 `.mm-band` > 중앙 `.mm-inner`(760/와이드1100), 변수 `--void/--cream` 통일.
- 모바일 좌우 여백 20px(히어로 제외 콘텐츠·배너). 가로영역 100% 확장이 기본.

## 6. 금기 (디자인 제출 전 자문)
- 유채색 액센트 / 장식 / 촌스러운 그림자·그라데이션 금지.
- 중복 CTA 금지(상위 공통 CTA 있으면 컴포넌트 내부 CTA 불필요).
- 🚨 **TMS 내부도구는 예외** — 가이드 100% 추종 X(최신 트렌드+효율, 컬러감만). **고객 대면만 100% 준수**.

## 체크리스트 (페이지/시안 제출 전)
- [ ] 모노크롬만 썼나 (유채색 액센트 0)
- [ ] 1초 테스트 통과 (핵심 잔상)
- [ ] 위계 3순위 명확 (문제↔답 대비)
- [ ] 타이포 스케일·clamp·11px 하한 준수
- [ ] 여백이 그룹핑과 일치
- [ ] 모바일 좌우 20px / 실기기 폭 실측 (헤드리스만으로 판정 금지)
- [ ] 중복 CTA 없음
- [ ] 브랜드 금기(할인·구매권유·장사꾼톤) 없음

## 더 깊이
- 전체 기준: `.claude/MAMORU-Complete-Brand-Guide-v1.0.md`
- 페르소나(냉정·솔직 평가): memory `persona_product_design.md`
- 페이지 표준: memory `reference_mamoru_page_kit`, `.claude/ADDENDUM_IMWEB.md`
