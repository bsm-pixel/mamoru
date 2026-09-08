# 주간 검색어 자동 추출 (네이버 검색광고 API → Notion)

매주 월요일 오전 9시(KST), GitHub Actions가 네이버 검색광고 API로 미용가위 관련
검색어 + 월간검색량을 뽑아 Notion `🔍 주간 검색어 로그` DB에 새 항목으로 기록한다.

> 코드에는 어떤 비밀키도 들어있지 않다. 전부 GitHub Secrets(환경변수)에서 읽는다.

## 사장님이 할 일 (딱 2단계 · 각 2분)

### 1) Notion 통합 토큰 발급 + DB 공유
1. https://www.notion.so/my-integrations 접속 → **새 API 통합(New integration)**
2. 이름 `마모루 검색어봇`, 유형 **Internal**, 워크스페이스 선택 → 저장
3. **Internal Integration Secret** 복사 (이게 `NOTION_TOKEN`, `secret_...` 또는 `ntn_...`)
4. Notion에서 **🔍 주간 검색어 로그** DB 열기 → 우상단 `···` → **연결(Connections)** → `마모루 검색어봇` 추가
   (이걸 해야 봇이 이 DB에 쓸 수 있음)

### 2) GitHub 비밀값 5개 등록
`github.com/bsm-pixel/mamoru` → **Settings → Secrets and variables → Actions → New repository secret**
아래 5개를 각각 등록:

| 이름 | 값 |
|------|-----|
| `NAVER_API_KEY` | 네이버 액세스라이선스 (01000000...) |
| `NAVER_SECRET_KEY` | 네이버 비밀키 (AQAAA...) |
| `NAVER_CUSTOMER_ID` | `4498405` |
| `NOTION_TOKEN` | 위 1)에서 복사한 통합 토큰 |
| `NOTION_DB_ID` | `4c94f972-7936-4396-8586-516eb6a10dea` |

## 테스트 (월요일 안 기다리고 지금 확인)
GitHub → **Actions** 탭 → `주간 검색어 자동 추출` → **Run workflow** 버튼 클릭
→ 초록불이면 성공. Notion DB에 새 항목이 생겼는지 확인.
→ 빨간불이면 로그 캡처해서 Claude에게 보여주면 고쳐줌.

## 조사 범위 바꾸기
`main.py`의 `SEED_KEYWORDS` 목록만 수정하면 된다. 제목·표현(창의적 카피)은
자동화가 만들지 않는다 — 콘텐츠 만들 때 Claude에게 "이 검색어로 제목 뽑아줘" 요청.
