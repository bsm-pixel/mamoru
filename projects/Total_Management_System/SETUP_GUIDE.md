# MAMORU TMS 셋업 가이드

> 개발자가 아닌 사용자를 위한 단계별 안내서입니다.
> 각 단계를 순서대로 진행하세요. 완료 후 Claude에게 "N단계 완료"라고 말하면 됩니다.

---

## 전체 흐름 요약

```
1단계: Supabase 계정 생성 + 프로젝트 생성     (10분)
2단계: 데이터베이스 테이블 생성                 (5분)
3단계: 사용자 계정 생성                        (3분)
4단계: 환경변수 파일 작성                      (5분)
5단계: 로컬에서 실행 + 확인                    (3분)
```

---

## 1단계: Supabase 프로젝트 생성

### 1-1. 회원가입
1. https://supabase.com 접속
2. **Start your project** 클릭
3. GitHub 계정으로 로그인 (또는 이메일 가입)

### 1-2. 새 프로젝트 만들기
1. **New project** 버튼 클릭
2. 아래 정보 입력:
   - **Name:** `mamoru-tms`
   - **Database Password:** 원하는 비밀번호 입력 → **반드시 메모해둘 것**
   - **Region:** `Northeast Asia (Seoul)` 선택
3. **Create new project** 클릭
4. 약 2분 기다리면 프로젝트 생성 완료

### 1-3. API 키 복사 (3개 필요)
1. 좌측 메뉴에서 ⚙️ **Project Settings** 클릭
2. **API** 탭 클릭
3. 아래 3개 값을 메모장에 복사:

| 항목 | 위치 | 용도 |
|------|------|------|
| **Project URL** | 상단 `URL` | `NEXT_PUBLIC_SUPABASE_URL` |
| **anon public** | API Keys 섹션 | `NEXT_PUBLIC_SUPABASE_ANON_KEY` |
| **service_role** | API Keys 섹션 (Reveal 클릭) | `SUPABASE_SERVICE_ROLE_KEY` |

> ⚠️ **service_role 키는 절대 외부에 공유하지 마세요**

### Claude에게: "1단계 완료. URL은 https://xxx.supabase.co 입니다"

---

## 2단계: 데이터베이스 테이블 생성

### 2-1. SQL 에디터 열기
1. Supabase 대시보드 좌측 메뉴에서 **SQL Editor** 클릭 (</> 아이콘)
2. **New query** 클릭

### 2-2. 첫 번째 SQL 실행 (테이블 생성)
1. 이 파일의 전체 내용을 복사:
   ```
   projects/Total_Management_System/app/supabase/migrations/001_initial_schema.sql
   ```
2. SQL Editor에 붙여넣기
3. **Run** 버튼 클릭 (또는 Ctrl+Enter)
4. 하단에 `Success` 메시지 확인

### 2-3. 두 번째 SQL 실행 (제품 데이터)
1. **New query** 다시 클릭
2. 이 파일의 전체 내용을 복사:
   ```
   projects/Total_Management_System/app/supabase/migrations/002_seed_products.sql
   ```
3. SQL Editor에 붙여넣기
4. **Run** 클릭
5. `Success` 확인

### 확인 방법
- 좌측 **Table Editor** 클릭 → `products` 테이블 클릭 → 12개 제품이 보이면 성공

### Claude에게: "2단계 완료"

---

## 3단계: 로그인 사용자 생성

1. Supabase 대시보드 좌측 **Authentication** 클릭
2. **Users** 탭 확인
3. **Add user** → **Create new user** 클릭
4. 입력:
   - **Email:** `admin@mamoru.kr` (또는 원하는 이메일)
   - **Password:** 원하는 비밀번호
   - ✅ **Auto Confirm User** 체크
5. **Create user** 클릭

### Claude에게: "3단계 완료"

---

## 4단계: 환경변수 파일 작성

### 4-1. 파일 생성
1. VS Code에서 이 폴더를 열기:
   ```
   projects/Total_Management_System/app/
   ```
2. `.env.local.example` 파일을 복사해서 `.env.local` 이름으로 저장

### 4-2. 값 채우기
`.env.local` 파일을 열고 아래처럼 실제 값을 넣기:

```env
# ─── Supabase (1단계에서 복사한 값) ───
NEXT_PUBLIC_SUPABASE_URL=여기에_Project_URL_붙여넣기
NEXT_PUBLIC_SUPABASE_ANON_KEY=여기에_anon_public_키_붙여넣기
SUPABASE_SERVICE_ROLE_KEY=여기에_service_role_키_붙여넣기

# ─── 아임웹 API (아임웹 관리자에서 복사) ───
IMWEB_API_KEY=아임웹_API키
IMWEB_API_SECRET=아임웹_시크릿키

# ─── 롯데택배 (기존 GAS Script Properties에서 복사) ───
LOTTE_API_URL=기존_LOTTE_API_URL_PROD_값
LOTTE_CANCEL_API_URL=기존_LOTTE_CANCEL_API_URL_PROD_값
LOTTE_TRACK_API_URL=https://apigw.llogis.com:10100/api/pid/cus/714a/custmer-view-tracking
LOTTE_CLIENT_KEY=기존_LOTTE_CLIENT_KEY_PROD_값
LOTTE_JOBCUSTCD=기존_LOTTE_JOBCUSTCD_PROD_값
LOTTE_SENDER_NAME=기존_발송자이름
LOTTE_SENDER_TEL=기존_발송자전화
LOTTE_SENDER_ZIP=기존_발송자우편번호
LOTTE_SENDER_ADDR=기존_발송자주소
LOTTE_DEFAULT_FARE=03

# ─── Cron (아무 랜덤 문자열) ───
CRON_SECRET=mamoru-tms-cron-2026
```

### 아임웹 API 키 찾는 법
1. 아임웹 관리자 로그인
2. **환경설정** → **외부 서비스 연동** → **API 키**
3. API Key, Secret Key 복사

### 롯데택배 키 찾는 법
1. Google Apps Script 열기 (AS 백엔드)
2. **프로젝트 설정** → **스크립트 속성**
3. `LOTTE_API_URL_PROD`, `LOTTE_CLIENT_KEY_PROD`, `LOTTE_JOBCUSTCD_PROD` 등 복사

> 💡 아임웹/롯데 키를 지금 못 넣어도 괜찮습니다.
> Supabase 3개만 넣으면 로그인 + 대시보드 화면은 확인 가능합니다.

### Claude에게: "4단계 완료" (또는 "Supabase만 넣었고 나머지는 나중에")

---

## 5단계: 로컬 실행 + 확인

### 5-1. 터미널에서 실행
VS Code 터미널 (또는 Claude Code)에서:

```bash
cd projects/Total_Management_System/app
npm run dev
```

### 5-2. 브라우저에서 확인
1. 브라우저에서 열기: **http://localhost:3000**
2. 자동으로 로그인 페이지로 이동됨
3. 3단계에서 만든 이메일/비밀번호로 로그인
4. 대시보드 화면이 나오면 성공!

### 5-3. 확인 체크리스트

| 확인 항목 | 방법 | 예상 결과 |
|-----------|------|-----------|
| 로그인 | 이메일+비밀번호 입력 | 대시보드로 이동 |
| 대시보드 | /dashboard 접속 | 통계 카드 4개 표시 (전부 0) |
| 사이드바 | PC에서 좌측 | MAMORU 로고 + 메뉴 3개 |
| 모바일 내비 | 브라우저 창 좁히기 (375px) | 하단 탭 3개 |
| 주문 목록 | 좌측 메뉴 "주문관리" | 빈 목록 (동기화 전) |
| 동기화 | "아임웹 동기화" 버튼 | 아임웹 키 있으면 주문 불러옴 |
| 주문 상세 | 주문 행 클릭 | 상세 정보 + 송장 생성 버튼 |
| 설정 | "설정" 메뉴 | 동기화 이력 + 환경변수 안내 |

### 모바일 확인
- 브라우저 F12 → 상단 📱 아이콘 클릭 → iPhone SE 또는 Galaxy S8 선택
- 하단 탭 내비게이션이 보이면 정상

---

## 문제가 생겼을 때

| 증상 | 원인 | 해결 |
|------|------|------|
| 로그인 후 바로 다시 로그인 화면 | Supabase URL/Key 틀림 | `.env.local` 값 재확인 |
| "SHEET_NOT_FOUND" 에러 | DB 마이그레이션 안 함 | 2단계 다시 실행 |
| 동기화 버튼 눌러도 0건 | 아임웹 API 키 없음 | 4단계에서 아임웹 키 입력 |
| 송장 생성 실패 | 롯데택배 키 없음 | 4단계에서 롯데 키 입력 |
| `npm run dev` 에러 | 패키지 미설치 | `npm install` 먼저 실행 |

### Claude에게 도움 요청 시
"5단계에서 이런 에러가 나옵니다: [에러 메시지]" 식으로 알려주세요.

---

## 이후 단계 (Phase 2+)

Phase 1 완료 후 진행 가능한 작업:
- **Phase 2:** 상담 + AS 데이터 통합 (기존 GAS 데이터 연동)
- **Phase 3:** 이카운트 재고 연동
- **Phase 4:** 고객 리뷰 시스템
- **Phase 5:** 태블릿 계약서 + 전자서명
- **Phase 6:** 역할/권한 + 분석 대시보드

Claude에게: "Phase 2 시작" 이라고 말하면 됩니다.
