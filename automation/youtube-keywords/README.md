# 주간 유튜브 검색어 자동 추출 (YouTube Data API → Notion)

매주 월요일 오전 9시 30분(KST), GitHub Actions가 유튜브에서 미용가위·미용사 실무 검색어의
상위 영상을 조사해 "소형 채널도 올라갈 수 있는 빈자리"를 골라 Notion `🔍 주간 검색어 로그` DB에
새 항목(채널=유튜브)으로 기록한다.

> 코드에는 어떤 비밀키도 들어있지 않다. 전부 GitHub Secrets(환경변수)에서 읽는다.

## 기록되는 것
| 섹션 | 내용 |
|---|---|
| 💎 가위 관점 빈자리 | 상위10 채널 구독자 중앙값 5만 미만 or 2년 넘은 영상 절반 이상 → 다음 영상 후보 |
| 🟡 보통 | 그 사이 — 썸네일·첫 30초로 승부 |
| 🔥 대형 채널 판 | 상위가 구독자 20만↑ 채널 최신 영상 → 장기전·시리즈용 |
| 🔎 실제로 치는 표현 | 유튜브 자동완성 (제목·설명에 그대로 쓸 말) |
| 👀 미용사가 보는 것 | ②③ 실무 검색어 상위에 자주 뜨는 채널 10개의 최근 3개월 인기 제목 |

## 사장님이 할 일 (1단계 · 2분)
`github.com/bsm-pixel/mamoru` → **Settings → Secrets and variables → Actions → New repository secret**

| 이름 | 값 |
|---|---|
| `YOUTUBE_API_KEY` | Google Cloud `mamoru-content-bot` 프로젝트의 API 키 (youtube-bot) |

`NOTION_TOKEN`, `NOTION_DB_ID`는 네이버 봇 것을 그대로 같이 쓴다 (이미 등록돼 있음).

## 테스트
GitHub → **Actions** 탭 → `주간 유튜브 검색어 자동 추출` → **Run workflow**
→ 초록불이면 Notion DB에 "○월 ○주차 · 유튜브" 항목이 생긴다.
→ 빨간불이면 로그 캡처해서 Claude에게.

## 조사 범위 바꾸기
`main.py`의 `SEEDS` (① 가위 직접 / ② 실무 기법 / ③ 고민·증상) 목록만 고치면 된다.
할당량: 하루 10,000유닛 중 약 4,500 사용 (검색 1회 = 100유닛). 씨앗을 크게 늘리면 초과할 수 있음.
