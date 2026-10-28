# 출장 맡김 · 함께 출고 알림톡 등록안 (2026-10-27)

> 흐름: [TMS_FLOW_REPAIR.md](TMS_FLOW_REPAIR.md) 2026-10-27 절 · 마이그 165
> 가동 순서: ① 솔라피 템플릿 2개 등록·검수 → ② Make 새 시나리오(아래 분기 2개) → ③ TMS 설정 「Make 웹훅 URL (출장 맡김·함께 출고)」에 URL 입력 = 가동
> URL 이 비어 있는 동안: 맡김 확인은 나가지 않고, 함께 출고는 지금처럼 2통(복원수리 출고 + 판매 출고)으로 나간다.

🚨 솔라피 변수명 = TMS 가 보내는 키 그대로 (대소문자·밑줄까지). 키를 바꾸면 빈칸으로 나간다.

---

## ① 출장 맡김 확인 — `as_field_received` · 템플릿 이름 `TMS-출장맡김확인`

- **Make 분기 조건**: `template` = `as_field_received` (event `AS_FIELD_RECEIVED`)
- **카테고리**: 서비스이용 > 신청접수 (004002) · **이미지형**
- **배너**: `RECEIVED` / 복원수리 접수 / 가위를 잘 받았습니다 — 파일 `as_field_received.png`

**본문** (2026-10-28 승인 템플릿 톤으로 다듬음)
```
#{name}님, 안녕하세요
맡겨 주신 가위를 잘 받았습니다

✂️ 맡기신 가위
#{qty}자루 (마모루 #{qty_mamoru} · 타사 #{qty_other})

• 접수번호 : #{as_id}
• 받은 날 : #{received_date}
• 수리비 : #{pay_status}

가위마다 접수번호 꼬리표를 달아 따로 관리합니다
수리가 끝나면 수리내역서와 함께 보내 드립니다
```

**버튼**: 웹링크 [접수 내역 조회] → `https://page.mamoru.kr/projects/as/page_as_report.html?uid=#{as_uid}`
(수리 전에는 페이지에 "작업 준비 중" 안내, 내역서를 쓰면 가위별 내역으로 채워진다)

**검수 요청 메모·Meshy 아이콘 문구**: 등록 안내 페이지(아래) 그대로

---

## ② 제품·가위 함께 출고 — `combined_shipped` · 템플릿 이름 `TMS-함께출고`

- **Make 분기 조건**: `template` = `combined_shipped` (event `COMBINED_SHIPPED`)
- **카테고리**: 배송 > 상태안내 (006001) · **이미지형**
- **배너**: `SHIPPED TOGETHER` / 제품·가위 함께 출고 / 한 상자로 보내 드렸습니다 — 파일 `shipped_combined.png`
- **언제**: 같은 롯데 송장에 B2C 판매 제품과 복원수리 가위가 함께 있고, 기사님 집하 스캔이 감지됐을 때 (크론, 자동)

**본문**
```
#{name}님, 안녕하세요
주문하신 제품과 맡기신 가위를 함께 보내 드렸습니다

📦 함께 보낸 것
• 제품 : #{product_name}
• 복원수리 가위 : #{qty}자루 (#{as_id})

• 택배사 : #{courier}
• 송장번호 : #{tracking}

가위별 작업 내용은 수리내역서에서 확인하실 수 있습니다
```

**버튼**
1. 웹링크 [수리내역 조회] (기존 「복원수리 출고」와 같은 이름) → `https://page.mamoru.kr/projects/as/page_as_report.html?uid=#{as_uid}`
2. 버튼 종류 **「배송조회」** — 링크 없음 (카카오 자동 배송조회)

---

## Make 시나리오 — `07 출장 맡김·함께 출고` (id 10056000 · 웹훅 id 4469625) — 2026-10-28 구축, **꺼짐**
- 웹훅 → 라우터 → 분기 2개(`TMS-출장맡김확인`=as_field_received · `TMS-함께출고`=combined_shipped) → 솔라피 알림톡. 변수 매핑 완료, **templateId 만 비어 있음**.
- 가동: 검수 승인 → templateId 2개 입력 → 시험 발송 → 시나리오 켜기 → TMS 설정에 웹훅 URL 입력.
- 등록 안내(복사 버튼·배너 PNG): `C:\마모루드라이브\ClaudeWork\알림톡등록_출장맡김\등록안내.html`
- 새 시나리오마다 `Store incomplete executions` = **Yes** 확인 (솔라피 순단 1회로 시나리오가 꺼지는 사고 예방).
- 웹훅 URL 을 TMS 설정 → 알림·연동 → 「Make 웹훅 URL (출장 맡김·함께 출고)」에 넣고 저장.
