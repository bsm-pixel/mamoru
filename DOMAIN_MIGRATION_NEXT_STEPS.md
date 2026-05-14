# 🌐 도메인 전환 — 사장님 내일 작업 안내문

> 작성: 2026-05-14 (퇴근 직전)
> 클로드가 어젯밤 자동 완료한 작업 + 사장님이 출근하셔서 마무리할 작업 요약

---

## ✅ 어젯밤 클로드가 완료한 작업 (사장님 확인 불필요)

1. **GitHub Pages CNAME 파일 추가** — `page.mamoru.kr`
2. **TMS 코드 7개 파일 갱신** — `GITHUB_PAGES` 상수 + `REVIEW_FORM_BASE` 모두 `page.mamoru.kr` 로
3. **GAS Code.gs 갱신** — `GITHUB_PAGES_BASE/CONSULT` 모두 새 도메인
4. **아임웹 코드위젯 iframe HTML 13개 갱신** — `src=` + `postMessage origin` 모두 새 도메인
5. **상품 상세 v10_trendy.html (마스터+백업)** — inline 이미지 6개씩 새 도메인
6. **페이지 3개 아임웹 도메인 교체** — `mamoruscissors63682.imweb.me` → `mamoru.kr`
   - `consulting/page_form.html` (출장상담 링크 /58)
   - `consulting/page_diag.html` (IMWEB_DOMAIN 상수)
   - `consulting/page_change_request.html` (STORE_BOOKING_URL /52)
7. **공통 preconnect 2개 갱신**
8. **TypeScript 컴파일 에러 0건 확인** + 커밋 + push 완료 (Vercel 자동 빌드 진행 중)
9. **메모리 카탈로그 갱신** — `reference_solapi_templates.md` 새 도메인 반영

→ 커밋: `749c0af` (30 파일, +702/-51)

---

## 🌅 사장님 출근 후 작업 (예상 15~20분)

### Step 1️⃣ — NS 전파 확인 (1분)

브라우저에서 시도:
- `https://mamoru.kr` 접속 → **아임웹 쇼핑몰 사이트** 보여야 정상 ✅
  - 아직 카페24 쇼핑몰 보이면 → 전파 진행 중. 30분~1시간 후 재시도

### Step 2️⃣ — 아임웹 SSL 발급 확인 (1분)

- 아임웹 어드민 → 도메인·SSL → SSL 신청 상태 확인
- `https://mamoru.kr` 접속 시 **자물쇠 아이콘 정상**이면 SSL 발급 완료
- 안 되어 있으면 → 아임웹 SSL 신청 버튼 클릭 (Let's Encrypt 자동 발급, 수십 분)

### Step 3️⃣ — 아임웹 대표 도메인 전환 (3분)

NS 전파 + SSL 발급 모두 OK 확인 후:

1. 아임웹 어드민 → 사이트 관리 → 도메인 설정 (이전 화면)
2. **mamoru.kr 라디오 버튼 ◉ 클릭** → 대표 도메인 전환
3. 저장
4. 다시 `https://mamoru.kr` 접속 → 아임웹 정상 표시 확인

### Step 4️⃣ — 카페24 쇼핑몰 연결 해제 (2분) ⚠️ 마지막 정리

1. 카페24 쇼핑몰 어드민 → mamoru.kr 설정 화면
2. **"연결 해제"** 버튼 클릭
3. 확인 → 카페24 쇼핑몰과 mamoru.kr 매핑 끊김 (카페24 쇼핑몰 데이터는 보존됨)

> 이 작업은 카페24 쇼핑몰을 격리하는 마지막 단계입니다. 이전에 사장님이 봤던 그 "연결 해제" 버튼.

### Step 5️⃣ — 솔라피 템플릿 URL 2개 갱신 (5분 + 재검수 1~3 영업일)

솔라피 콘솔 → 알림톡 템플릿 → 다음 2개 본문/버튼 URL 갱신:

**1. `purchase_review_request` (구매 후기)**
- 옛 URL: `https://bsm-pixel.github.io/mamoru/projects/reviews/page_review.html?uid=#{order_uid}&type=purchase&name=#{name}`
- 새 URL: `https://page.mamoru.kr/projects/reviews/page_review.html?uid=#{order_uid}&type=purchase&name=#{name}`

**2. `as_review_request` (복원수리 후기)**
- 옛 URL: `https://bsm-pixel.github.io/mamoru/projects/reviews/page_review.html?uid=#{as_uid}&type=as`
- 새 URL: `https://page.mamoru.kr/projects/reviews/page_review.html?uid=#{as_uid}&type=repair`
  - ⚠️ `type=as` → `type=repair` 같이 변경 (코드 표준 통일)

**갱신 후 재검수 신청** → 카카오 검수 1~3 영업일 대기

> 검수 기간 동안 옛 템플릿이 계속 발송 가능 → 운영 무중단 (GitHub Pages 자동 redirect로 옛 URL 도 작동)

### Step 6️⃣ — 26종 템플릿 풀 URL 추가 점검 (10분)

사장님 시간 되실 때 솔라피 콘솔에서 26개 템플릿 본문/버튼 URL을 한 번 훑어보세요. **`bsm-pixel.github.io` 또는 `mamoruscissors63682.imweb.me`** 가 박혀 있는 게 보이면 추가로 알려주세요 — 같이 갱신·재검수 묶음에 넣겠습니다.

---

## 🧪 검증 체크리스트

전환 완료 후 다음 흐름을 한 번씩 테스트해 보세요:

- [ ] `https://mamoru.kr` 접속 → 아임웹 쇼핑몰 정상
- [ ] `https://mamoru.kr/52` → 매장방문 접수 페이지 정상
- [ ] `https://mamoru.kr/58` → 출장상담 페이지 정상
- [ ] `https://page.mamoru.kr/projects/consulting/page_change_request.html?uid=테스트` → 일정변경 페이지 정상
- [ ] `bsm@mamoru.kr` 으로 외부 메일 발송·수신 정상
- [ ] 사장님 본인 번호로 상담 알림톡 테스트 발송 (TMS에서 가능) → 카톡 버튼 클릭 → 새 도메인 정상 로드
- [ ] 옛 알림톡 (이미 발송된) 의 `bsm-pixel.github.io` 링크도 정상 작동 (GitHub Pages 자동 redirect)

---

## 🆘 문제 발생 시

다음 상황은 즉시 클로드에게 알려주세요:

| 증상 | 원인 추정 | 클로드가 도울 일 |
|---|---|---|
| `mamoru.kr` 접속 시 아임웹 안 보임 | NS 전파 지연 또는 아임웹 도메인 인식 못 함 | PowerShell DNS 조회 + 아임웹 설정 가이드 |
| `https` 자물쇠 깨짐 | SSL 발급 지연 | 아임웹 SSL 재신청 가이드 |
| 카톡 알림톡 버튼 클릭 시 404 | TMS 코드 미반영 또는 Vercel 빌드 실패 | Vercel 빌드 로그 확인 + 재배포 |
| 메일 끊김 (bsm@mamoru.kr) | MX 레코드 미반영 | 아임웹 DNS 어드민 MX 재확인 |
| 솔라피 재검수 거절 | 본문 정책 위반 | 본문 텍스트 조정 가이드 |

---

## 📊 작업 완료 후 메모리 갱신 필요

전환 완료되면 클로드에게 **"도메인 전환 완료"** 라고 알려주세요. 다음 메모리 갱신 진행:
- `MEMORY.md` 도메인 매핑 섹션 갱신
- 옛 URL 호환 안내 제거 (안정화 후)
- `reference_iframe_pages.md` 갱신
- `TMS_ROADMAP.md` 작업 완료 기록

---

## 📞 카페24 자동연장 (선택)

카페24 호스팅 어드민에서 **자동연장 카드 등록** 안 하셨으면:
- 이번처럼 만료 D-10 직전 발견 사고 방지
- 1년 22,000원 자동 결제
- 카페24 어드민 → 도메인 관리 → mamoru.kr → "카드등록"

선택사항이지만 권장.

---

[🎨 UX아키텍트 | ⚙️ 수석 개발자 | 🎯 잡스 | 🏢 COO 합의 완료]
