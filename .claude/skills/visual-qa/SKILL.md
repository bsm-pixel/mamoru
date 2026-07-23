---
name: visual-qa
description: >-
  고객 대면 화면(페이지·모달·카탈로그·카드)을 만들거나 고친 뒤 "됐다" 하기 전에 PC/모바일 실제 렌더를
  검증하는 절차. Use whenever a customer-facing UI change was made — verify it actually renders on
  mobile AND desktop before claiming done, and avoid headless-viewport false judgments.
when_to_use: >-
  모바일, 폰에서, 렌더, 스크린샷, 캡처, 반응형, 비주얼 QA, 화면 확인, 깨짐, 여백, 1:1, 비율, 열수, 컬럼
---

# 시각 검증 (고객 화면 "됐다" 하기 전)

> 왜 있나: 로직·코드만 맞다고 "됐습니다" 했다가 **모바일에서 깨져** 여러 번 재수정한 이력.
> 규칙: **눈으로 본 것만 "됐다"고 말한다. 안 봤으면 "안 봤다"고 말한다.**

## 0. 무조건 둘 다
화면 변경이면 **PC 폭 + 모바일 폭 둘 다** 실제로 띄워 확인. 바꾼 플로우는 끝까지(담기→합계→접수 등).

## 1. 🚨 헤드리스 모바일 오판 함정 (가장 자주 틀림)
- `chrome --headless --window-size=390` **만으로 모바일 열수/비율 판정 금지.** vw·`auto-fit`·미디어쿼리가 실제 폰과 다르게 나옴(과거 2번 헛짚음).
- 신뢰할 실측: **puppeteer `page.setViewport({width, height, isMobile:true})` + `getComputedStyle`** 로 실제 계산값 확인.
- 헤드리스 스크린샷은 "대략 배치" 확인용. **열 개수·비율은 실측 수치로** 판단.
- 렌더 레시피(대략 확인): `chrome --headless=new --no-sandbox --user-data-dir=<tmp> --window-size=W,H --virtual-time-budget=3000 --screenshot=out.png "file:///..."` → 캡처 Read.

## 2. 모바일 레이아웃 체크
- **좌우 여백 20px** (히어로 아닌 콘텐츠·배너). 가로영역 100% 확장 기본.
- **컬럼 붕괴·가로넘침**: A=`flex-col md:flex-row`, B=`flex-wrap`+`w-full sm:flex-1`, C=라벨 `shrink-0 whitespace-nowrap`, 전역 `html{overflow-x:hidden}`.
- inline 그리드는 **`repeat(auto-fit, minmax(clamp(...),1fr))`** (고정 `repeat(N)` 금지). ~360px 폭 역산해 컬럼 수 확인.
- 이미지 비율은 **`aspect-ratio`(폭 기준)** 로 — `height:vh`+`object-fit:cover`는 비율을 깬다(크롭). 임베드 iframe은 vh가 콘텐츠 전체높이라 불안정 → `min(…, px)` 상한.

## 3. 자산·라이브 검증
- `src`/`href` 가 실제 파일과 매칭되나 (깨진 이미지 = 고객 노출).
- 배포 후 `curl`로 라이브 콘텐츠 실측. 이미지 교체는 Fastly 엣지캐시 함정 → `git hash-object`(로컬)↔`curl|git hash-object`(라이브) 대조.

## 4. 1초 테스트 (제출 직전)
화면 1초 보고 눈 감았을 때 핵심 메시지·CTA가 잔상으로 남나. 위계·여백은 `mamoru-brand` 기준.

## 보고할 때
"확인함: PC/모바일 `<무엇을>` 봤고 `<결과>`. **확인 못 함**: `<헤드리스 한계로 실측 못 한 것>`." — 못 본 건 숨기지 말 것.
관련 memory: `reference_headless_mobile_viewport`, `reference_mobile_column_collapse`, `feedback_visual_qa_mobile_first`.
