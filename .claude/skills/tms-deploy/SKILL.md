---
name: tms-deploy
description: >-
  TMS(Vercel) 또는 고객 페이지(GitHub Pages)를 배포하기 전 검증 체크리스트와 승인 절차를 로드.
  Use when the user says 배포/배포하자/올려/push/deploy, or before pushing any code — to run
  tsc/build/lint, trace all consumers, verify actual behavior, then ask for approval once.
when_to_use: >-
  배포, 배포하자, 배포해줘, 올려, 푸시, push, deploy, 빌드, tsc, 커밋하고 배포, 반영, 라이브 반영
---

# TMS·페이지 배포 프로토콜 (밀어붙이기 전 필수)

> 목적: "됐습니다"를 실측 없이 말하지 않기. 배포 후 고객·사장님이 바로 보므로.

## 어디에 배포하나 (먼저 구분)
| 대상 | 방식 | 배포 규칙 |
|---|---|---|
| **고객 페이지**(consulting/as/main/event/stock_sale/products…) | GitHub Pages | **묻지 말고 즉시 커밋+푸시** (단, 검증은 동일하게) |
| **TMS**(Total_Management_System/app) | Vercel | **검증 완주 → 배포 직전 딱 1회 승인** → 승인 시 push |

## A. 전수 추적 (사소한 것 하나도)
- 바꾼 필드/컴포넌트/함수/타입/상태값의 **모든 소비처를 grep**. "내가 고친 화면"만 보지 말 것.
- 같은 패턴이 다른 파일에도 있는지 grep → 있으면 같이 고치거나 왜 안 고치는지 명시.
- 데이터 출처 끝까지: 화면값 → 훅 → API/RPC → DB 컬럼. 중간 덮어쓰기/캐시/타임존 없나.
- 엣지: null·빈문자열·0건·구데이터·권한없음·에러응답·긴텍스트·모바일폭.

## B. 기계 검증 (전부 통과)
1. `cd projects/Total_Management_System/app && npx tsc --noEmit` → **에러 0**
2. `npm run build` → **성공** (tsc가 못 잡는 라우트/서버컴포넌트 에러를 여기서 잡음. 생략 금지)
3. `npm run lint` → **신규 에러 0**
- 🔒 재고/시리얼 건드렸으면 `git diff | grep -E "stock_quantity|raw_stock"` 로 의도 외 변경 0 확인.

## C. 실제 구동 검증
- 로직만 맞다고 "됐습니다" 금지. **화면을 실제로 띄워** 바꾼 플로우를 끝까지(입력→저장→재조회→표시).
- 화면 변경이면 **PC + 모바일 폭 둘 다**. 헤드리스 `--window-size`만으론 모바일 열수 판정 금지(puppeteer isMobile 실측).
- 못 한 검증은 **숨기지 말고 "이건 확인 못 했다"** 명시 + 대체 검증 제시.

## D. 커밋 (푸시는 아직)
- 한국어 메시지. 흐름도(`docs/TMS_FLOW_*.md`)·로드맵 갱신 필요 시 같은 커밋에.
- 트레일러: `Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>`

## E. 배포 직전 — TMS는 딱 1회만 질문 (증거와 함께)
> "TMS 검증 완료
>  · 전수추적: `<필드>` 소비처 N곳 grep → 영향 [있음/없음]
>  · tsc 0 / build 성공 / lint 0
>  · 실제 구동: `<화면·플로우>` 돌려 `<확인한 것>` (PC/모바일)
>  · 엣지: null·0건·구데이터 확인
>  · ⚠️ 확인 못 한 것: `<있으면 명시>`
>  커밋 `<sha>` 준비됨. **배포할까요?**"
- 승인 → `git push`. 승인 전엔 **절대 push 안 함.**
- 배포 후: `gh api repos/bsm-pixel/mamoru/commits/<sha>/check-runs` 로 build/deploy success 확인.

## ❌ 금지
- 검증 도중 "이거 해도 될까요?" 끊어 묻기 (승인 후엔 무중단 완주, 배포 직전에만 1회)
- tsc만 돌리고 build·구동 생략 / 실측 없이 "~일 겁니다" 단정 / 실패 숨기고 "됐습니다"
- `common_code/header_code.txt` 등 공유 공통코드 손대기
