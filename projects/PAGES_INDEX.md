# 📄 고객 페이지 편집·배포 가이드 (page.mamoru.kr)

> 최종 업데이트: 2026-08-21
> 베이스 URL: `https://page.mamoru.kr/` · 모든 고객 페이지는 `projects/` 아래에 있습니다.

---

## ✅ 편집 → 배포 (이게 전부)

1. **로컬에서 파일을 연다** (아래 표의 "파일" 경로 — `projects/...html`)
2. **화면에 보이는 문구를 그대로 고친다** (텍스트만 바꾸면 됨)
3. 저장 후 **`git push`** → **GitHub Pages가 1~2분 안에 자동 반영** (별도 빌드·비용 없음)

> 🚫 **폴더·파일 이름은 바꾸지 말 것** — 라이브 URL이 **알림톡·QR 라벨·아임웹 코드위젯에 박제**돼 있어서, 이름을 바꾸면 그 링크들이 전부 깨집니다. **내용(문구)만** 고치세요.
>
> ✏️ 파일 안에 `✏️` 표시가 있으면 거기가 "고쳐도 되는 곳"입니다.

---

## 🛡️ 정품인증 (신규, 2026-08)

| 페이지 | 파일 | 라이브 URL | 문구 고치는 곳 |
|--------|------|-----------|---------------|
| 정품 확인 | `projects/verify/index.html` | `/projects/verify/?t={token}` | `renderValid()` 안의 "정품 확인 / 이 제품은…", 혜택(복원수리 50%) 문구 |
| 가위 자가진단 | `projects/verify/self-check.html` | `/projects/verify/self-check.html` | 증상 6개(`<div class="checks">`), 결과 3단계 문구(`run` 클릭 스크립트의 `body=`), 접수 링크(`REPAIR_INTAKE`/`REPAIR_GUIDE`) |

> QR 라벨은 TMS 시리얼 관리에서 출력(고객 페이지 아님). QR = 위 정품확인 URL.

## 💬 상담 (매장방문 / 출장 / 톡상담)

| 페이지 | 파일 | 라이브 URL | 용도 |
|--------|------|-----------|------|
| 상담 안내 | `projects/consulting/page_guide.html` | 아임웹 `/61` | 매장방문/출장/톡상담 안내 |
| 상담 접수 | `projects/consulting/page_form.html` | 코드위젯 | 접수 폼 |
| 간편 진단 | `projects/consulting/page_diag.html` | 코드위젯 | 어떤 가위가 맞나 진단 |
| 맞춤 추천 | `projects/consulting/page_recommend.html` | 코드위젯 | 진단 결과 → 추천 |
| 출장 일정 제안 | `projects/consulting/page_suggest.html` | `/projects/consulting/page_suggest.html?t={token}` | 고객이 출장 슬롯 선택 |
| 예약 변경/취소 | `projects/consulting/page_change_request.html` | `...?uid={unique_id}` | 셀프 변경/취소 |
| 알림 결과 | `projects/consulting/page_result.html` | `...?title=&msg=` | 알림톡 버튼 결과 표시 |

## 🔧 복원수리

| 페이지 | 파일 | 라이브 URL | 용도 |
|--------|------|-----------|------|
| 복원수리 안내 | `projects/as/page_guide.html` | 아임웹 `/60` | 서비스 안내(풀페이지) |
| 복원수리 접수 | `projects/as/page_form.html` | 아임웹 `/53` | 통합 접수 폼 |
| 수리내역 조회 | `projects/as/page_as_report.html` | `...?uid={as_id}` | 고객 수리현황(TMS API) |

## ⭐ 후기 / 🏠 메인

| 페이지 | 파일 | 라이브 URL | 용도 |
|--------|------|-----------|------|
| 후기 작성 | `projects/reviews/page_review.html` | `...?uid={id}&type={type}` | 리뷰 작성 폼 |
| 고객 후기 목록 | `projects/reviews/page_reviews.html` | 코드위젯 | 리뷰 표시 |
| 메인 홈 | `projects/main/page_main.html` | 아임웹 `/` | 메인 페이지 |
| 리뷰 이벤트 | `projects/reviews/page_review_event.html` | 코드위젯 | 매월 후기 이벤트 |

---

## 🔗 URL 빠른 복사

```
# 정품인증
https://page.mamoru.kr/projects/verify/?t={token}
https://page.mamoru.kr/projects/verify/self-check.html
# 상담
https://page.mamoru.kr/projects/consulting/page_suggest.html?t={token}
https://page.mamoru.kr/projects/consulting/page_change_request.html?uid={unique_id}
# 복원수리
https://page.mamoru.kr/projects/as/page_form.html
https://page.mamoru.kr/projects/as/page_as_report.html?uid={as_id}
# 후기
https://page.mamoru.kr/projects/reviews/page_review.html?uid={id}&type={type}
```

---

## ⚠️ 주의

- **아임웹 iframe 삽입 페이지**(안내/진단/추천/후기목록/메인)는 아임웹 안에서 열려야 정상. 단독 URL 직접 접근은 레이아웃이 다를 수 있음.
- **공통코드**(`common_code/header_code.txt` 등)는 여러 페이지가 공유하므로 **손대지 말 것**.
- 문구가 아니라 **기능·레이아웃**을 바꾸려면 Claude에게 요청 (깨질 위험 있음).
