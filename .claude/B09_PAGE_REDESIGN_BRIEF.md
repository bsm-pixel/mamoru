> 2026-10-23 B-09 개편(22개 페이지)에 실제로 쓴 작업 지침서 원본. 다음 개편·시안 작업 때 에이전트에게 그대로 넘긴다.
> ⚠️ 아래 캡처 도구 경로(scratchpad/pp)는 당시 세션 임시 폴더 — 새 세션에선 memory reference_headless_mobile_viewport 의 설치법대로 puppeteer-core를 임시 폴더에 다시 설치해 쓸 것.

# B-09 시안 제작 지침 (공통) — 반드시 끝까지 읽고 따를 것

저장소: `C:\Users\user\Desktop\mamoru` (GitHub Pages = https://page.mamoru.kr/<경로>). 고객 페이지는 `projects/<폴더>/*.html`.

## 0. 절대 규칙
- **원본 페이지는 절대 수정하지 않는다.** 시안 파일만 새로 만든다: `projects/_design_lab/<이름>_b09.html`.
- **커밋·푸시 금지.** `projects/_design_lab/redesign.html`(현황판)도 건드리지 않는다.
- 다른 파일·공통코드(`projects/common_code/*`) 수정 금지.
- 문구(글자)·id·class(스크립트가 쓰는 것)·data-*·form name/필드·버튼 onclick·`<script>` 로직 **100% 유지**. 디자인만 바꾼다. (예외: 구절 줄바꿈용 `<span class="b-nw">` 감싸기는 허용 — 단어 변경 없이)
- 한국어로 보고.

## 1. 디자인 기준 (먼저 읽기)
1. `C:\Users\user\Desktop\mamoru\.claude\MAMORU-Complete-Brand-Guide-v1.0.md` 의 **B-06·B-07·B-08·B-09** 섹션 (B-09 = 개성 레이어, 최우선).
2. 사장님 피드백 누적: `C:\Users\user\.claude\projects\c--Users-user-Desktop-mamoru\memory\feedback_b09_distinctive_approved.md` (전부 반영 필수).
3. **완성 예시 2개 (사장님 OK·반영된 것) — 반드시 열어서 패턴을 그대로 따를 것:**
   - `projects/_design_lab/as_guide_b09.html` (복원수리 안내) — 파일 맨 아래쪽 `<style id="b09">` 블록과 맨 위 안내 띠, 히어로 `b-hero` 마크업
   - `projects/_design_lab/consult_guide_b09.html` (상담 안내) — 같은 구조

### 사장님 피드백 요약 (모두 적용)
- 모노크롬, 그림자·그라데이션 금지, 모서리 0/2/4px(대부분 2px). 새 글꼴 금지(Outfit + Pretendard만).
- **흑백 필터는 히어로 시그니처 사진에만.** 일러스트·아이콘·작업/후기 사진은 **원래 색 유지**.
- 시그니처 = 흑백 확대 사진 위 **hairline 정밀 주석(점·선·짧은 라벨)**. 주석은 **실제 부위/물건을 정확히** 짚을 것(사진을 Read로 확대해 좌표 산정). 라벨은 사실만(페이지에 있는 내용·사장님 용어). 확신 없으면 주석 넣지 말고 보고서에 "위치 확인 필요"로 적기.
- 사진 위 주석은 사진을 **원본 비율 그대로**(aspect-ratio = 원본 가로/세로) 둬야 좌표가 안 틀어짐. 정사각형 아니면 circle 대신 비율 보정 `ellipse`.
- 거대 영문 단어(Outfit 900)가 사진 아래 끝을 살짝 덮는 겹침 1곳. 단어는 페이지 내용에서(예: RESTORE, CHOSEN).
- 제목 굵기 대비(얇은 300 ↔ 굵은 900), 숫자 스펙 3개(페이지에 있는 사실만), 거대 번호 인덱스(01·02·03, Outfit 300, 흐린 색).
- **문제→답**: 문제는 취소선·흐리게, 답은 크고 굵게.
- 탭: 둥근 알약 → 글자 탭 + 굵은 밑줄 2px. **안내 라벨을 탭과 같은 모양·위치로 두지 말 것**(머리말 + 세로 구분선 + 가는 선).
- 왼쪽 정렬로 바꿀 때 **외톨이 요소·빈 공간 금지**: 아이콘이 홀로 뜨면 [거대 번호 | 글 | 일러스트] 3칸, 아이콘 여러 개는 카드 폭 균등, 1개면 글 오른쪽.
- **모바일 줄바꿈**: ≤767px에서 좁은 2단 금지(1단 행). 번호 칸 때문에 글이 좁으면 번호를 글 위로. 구절 중간 끊김은 `<span class="b-nw">`(white-space:nowrap)로 묶기. `word-break:keep-all`. **360px에서 확인 필수.**
- **CTA 중복 금지**: 원래 구조의 CTA(모바일 하단 고정 바 등)를 존중. 모바일 히어로에 버튼 추가 금지(원래 없으면). PC 전용 버튼은 유지.
- PC: 페이지 폭이 원래 제한(예 760px)이면 그 안에서. 폭 제한이 없으면 히어로 등은 max-width 1100 안에.
- 페이지에 없는 CSS 변수(--mist 등)를 쓰면 b09 블록에서 정의.

### 접수폼·작성폼 페이지 (가볍게)
- 입력 편의가 최우선. 입력칸·버튼 크기 줄이지 말 것. 큰 히어로 사진/주석 넣지 말 것.
- 적용 범위: 제목 위계(얇↔굵), 가는 선 구분, 모서리 2px, 그림자 제거, 라벨/안내문 정리, 단계 번호 거대화 정도.

## 2. 작업 방법 (예시와 동일하게)
1. 원본 HTML을 읽고 구조·스크립트가 참조하는 class/id 파악(`querySelector`, `getElementById`, `classList` grep).
2. 원본을 복사해 시안 파일 생성:
   - 상대경로 자산(`src="images/...`, `url('images/...`, `src="videos/...` 등)을 `../<원본폴더>/...`로 바꿈(시안이 `_design_lab`에 있으므로). 절대 URL은 그대로.
   - `<body>` 바로 뒤에 **안내 띠**(예시 파일 맨 위 div 그대로 복사, 문구만 해당 페이지로): "🧪 작업실 시안 · B 개편안(B-09) — <페이지명>. 문구·버튼은 같고 디자인만 다릅니다" + `A 지금 페이지 보기 ↗`(https://page.mamoru.kr/projects/<원본경로>) + `현황판`(redesign.html) 링크. 특이사항 있으면 한 줄 추가.
   - `<title>`을 `시안 B-09 · <페이지명>`으로, `<meta name="robots" content="noindex">` 추가.
   - `</head>` 바로 앞에 B-09 덧입힘 블록: Outfit 웨이트 link + `<style id="b09"> ... </style>` (주석으로 콘셉트·시그니처 명시). 기존 스타일보다 뒤에 와야 함. 필요시 `!important`.
   - 히어로가 있는 안내·소개 페이지만 히어로 마크업 교체(스크립트가 참조하는 class 유지).
3. 원본에서 실제 반영할 때 "같은 변경을 원본에 그대로 옮길 수 있도록" 변경은 **(a) `<style id="b09">` 블록 (b) 히어로 마크업 (c) b-nw span** 세 종류로 한정. 나머지 마크업은 건드리지 말 것.

## 3. 검증 (필수 — 결과 보고에 포함)
- 캡처 도구: Node `C:/Users/user/AppData/Local/Temp/claude/c--Users-user-Desktop-mamoru/a8bc1b03-f69b-43b4-91b4-7668e8caaee8/scratchpad/pp/node_modules/puppeteer-core`, Chrome `C:/Program Files/Google/Chrome/Application/chrome.exe`, 실행 인자 `--allow-file-access-from-files`. 파일은 `file:///C:/Users/user/Desktop/mamoru/projects/_design_lab/<이름>_b09.html` 로 연다.
- 뷰포트 3종: 360(isMobile true), 375(isMobile true), 1280(PC). 탭이 있으면 탭마다. `.fade-in`/reveal 류는 캡처 전 `visible` 등 클래스 강제 부여, `prefers-reduced-motion: reduce` 에뮬레이트. 안내 띠(sticky)는 캡처 때 `position:static`으로.
- `document.documentElement.scrollWidth`가 뷰포트 폭과 같아야 함(가로 넘침 0).
- 캡처 PNG를 **Read로 직접 보고** 문제(겹침·잘림·구절 중간 끊김·외톨이 요소·대비 부족)를 고칠 것. 길면 PIL(`py -c "from PIL import Image..."`)로 잘라 보기. 파이썬은 `py` 명령.
- 캡처 파일은 scratchpad 폴더 안 `b09_<이름>_*.png`로 저장.

## 4. 보고 형식
- 만든 파일 경로, 페이지별 콘셉트·시그니처, 바꾼 것 요약(5줄 이내), 원본 대비 변경 종류(style 블록/히어로/b-nw),
- 검증 결과(360/375/1280, 가로 넘침, 탭별), 스스로 고친 문제,
- 사장님 확인이 필요한 점(주석 위치·사진 선택·원래 CTA 구조 등) — 솔직하게.

## 5. 추가 규칙 (2026-10-23 사장님 확정 — 반드시 적용)
- **브랜드명 워드마크**: 제목·로고·디스플레이 크기의 "MAMORU" 글자는 Outfit 타이핑 금지 → `projects/brand/img/mamoru-wordmark.svg` 내용을 인라인 `<svg class="mm-wordmark ..." aria-hidden="true">`로 넣고 감싸는 요소에 `aria-label`로 원래 글자. 9.28:1로 길다 → 모바일은 폭 기준(`width:min(100%,Npx);height:auto`), 한 줄 안에 섞을 땐 `height:.74em;width:auto`. 문장 속 MAMORU·작은 eyebrow 라벨은 대상 아님. 이미 SVG 로고가 있으면 그대로 둠.
- **이미 OK된 완성 예시 추가**: `projects/_design_lab/brand_intro_b09.html`, `review_write_b09.html`(작성폼 가볍게), `event_address_b09.html`(입력폼 가볍게), `review_event_b09.html` — 폼 페이지는 뒤 두 개 패턴을 따를 것.
- **토큰/데이터가 있어야 보이는 페이지**(수리내역서·정품확인·일정 변경·제안·결과 등): 파일·스크립트는 고치지 말고, 캡처할 때만 puppeteer `page.setRequestInterception(true)`로 API 응답을 **가짜 데이터로 대답**해 정상 화면을 띄워 확인(응답 형태는 스크립트에서 읽는 필드를 보고 맞출 것). URL 파라미터가 필요하면 `?token=demo` 식으로 붙여 열기. 오류·만료 화면 상태도 1장 캡처해 깨지지 않는지 확인. 실제 API에 POST(저장·접수) 보내지 말 것 — 제출 버튼은 누르지 않는다.
- **안내 띠에 "실제 링크로 열면 고객 데이터가 보입니다 — 시안은 빈/데모 상태" 같은 특이사항 1줄** 추가(해당 페이지만).
- 파일이 **아임웹에 직접 붙여넣는 코드**(iframe으로 불러오는 게 아니라 `<style>`부터 통째로 붙이는 조각)인지 확인해 보고서에 적을 것.
