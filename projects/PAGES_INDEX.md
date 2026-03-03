# 단독 페이지 인덱스 (GitHub Pages)
> 최종 업데이트: 2026-03-03
> 베이스 URL: `https://bsm-pixel.github.io/mamoru/`

---

## 고객 대면 — 단독 페이지 (알림톡/링크 직접 접근)

| 페이지 | 경로 | 용도 | 파라미터 |
|--------|------|------|----------|
| 출장 일정 제안 | `projects/consulting/page_suggest.html` | 출장 시간 제안 캘린더 — 고객이 슬롯 선택/재요청 | `?t={token}` |
| 예약 변경/취소 | `projects/consulting/page_change_request.html` | 확정된 예약 셀프서비스 변경/취소 | `?uid={unique_id}` |
| 알림 결과 | `projects/consulting/page_result.html` | 알림톡 버튼 → 결과 메시지 표시 (동적) | `?title=&msg=` |
| 딜러 일정 확정 | `projects/consulting/page_dealer_confirm.html` | 출장 딜러 대면 확정 | — |
| 리스케줄 리다이렉트 | `projects/consulting/page_reschedule.html` | 레거시 링크 호환 → page_suggest.html 자동 이동 | `?t={token}` |
| 복원수리 접수 | `projects/as/page_form.html` | 통합 접수 폼 (마모루+타사) | — |
| 복원수리 안내 | `projects/as/page_as_guide.html` | 복원수리 서비스 안내 (모달형) | — |
| 수리내역 조회 | `projects/as/page_as_report.html` | 고객 수리 현황 조회 (TMS API 연동) | `?uid={as_id}` |
| 후기 작성 | `projects/reviews/page_review.html` | 리뷰 작성 폼 (복원수리/상담 등) | `?uid={id}&type={as\|consult}` |
| 브랜드 소개 | `projects/brand/page_intro.html` | 브랜드 히어로 랜딩 | — |

---

## 아임웹 iframe 삽입 페이지 (코드위젯)

| 페이지 | 경로 | 아임웹 경로 | 용도 |
|--------|------|-------------|------|
| 메인 홈 | `projects/main/page_main.html` | `/` | 메인 페이지 (PC+모바일, Lottie) |
| 상담 안내 | `projects/consulting/page_guide.html` | `/61` | 매장방문/출장/톡상담 안내 |
| 간편 진단 | `projects/consulting/page_diag.html` | 코드위젯 | 13개 질문 조건부 분기 |
| 상담 접수 | `projects/consulting/page_form.html` | 코드위젯 | 상담 접수 폼 (매장/출장) |
| 맞춤 추천 | `projects/consulting/page_recommend.html` | 코드위젯 | 진단 결과 → 맞춤 추천 |
| 복원수리 안내 | `projects/as/page_guide.html` | `/60` | 복원수리 서비스 안내 (풀페이지) |
| 고객 후기 | `projects/reviews/page_reviews.html` | 코드위젯 | 리뷰 목록 표시 |

---

## 전체 URL 목록 (빠른 복사)

### 고객 대면 (단독)

```
# [출장 일정 제안] 관리자가 제안한 출장 시간을 고객이 캘린더에서 선택/재요청
https://bsm-pixel.github.io/mamoru/projects/consulting/page_suggest.html?t={token}

# [예약 변경/취소] 확정된 매장방문/출장 예약을 고객이 직접 변경 또는 취소
https://bsm-pixel.github.io/mamoru/projects/consulting/page_change_request.html?uid={unique_id}

# [알림 결과] 알림톡 버튼 클릭 → 동적 메시지 표시 (제목/본문 파라미터)
https://bsm-pixel.github.io/mamoru/projects/consulting/page_result.html?title={제목}&msg={메시지}

# [딜러 일정 확정] 출장 딜러가 고객 대면 시 일정 확정 처리
https://bsm-pixel.github.io/mamoru/projects/consulting/page_dealer_confirm.html

# [리스케줄 리다이렉트] 레거시 알림톡 링크 호환 → page_suggest.html로 자동 이동
https://bsm-pixel.github.io/mamoru/projects/consulting/page_reschedule.html?t={token}
```

```
# [복원수리 접수] 마모루+타사가위 통합 접수 폼 (수량/진행방식/주소/사진)
https://bsm-pixel.github.io/mamoru/projects/as/page_form.html

# [복원수리 안내] 복원수리 서비스 소개 페이지 (모달형, 알림톡 링크용)
https://bsm-pixel.github.io/mamoru/projects/as/page_as_guide.html

# [수리내역 조회] 고객이 접수번호로 수리 현황/검수결과 확인 (TMS API 연동)
https://bsm-pixel.github.io/mamoru/projects/as/page_as_report.html?uid={as_id}
```

```
# [후기 작성] 복원수리/상담 완료 후 고객 리뷰 작성 폼
https://bsm-pixel.github.io/mamoru/projects/reviews/page_review.html?uid={id}&type={type}

# [브랜드 소개] MAMORU 브랜드 히어로 랜딩 페이지
https://bsm-pixel.github.io/mamoru/projects/brand/page_intro.html
```

### 아임웹 iframe (직접 접근은 비정상)

```
# [메인 홈] 아임웹 메인페이지에 삽입되는 히어로+감성 레이아웃 (PC+모바일, Lottie)
https://bsm-pixel.github.io/mamoru/projects/main/page_main.html

# [상담 안내] 매장방문/출장/톡상담 서비스 안내 (아임웹 /61)
https://bsm-pixel.github.io/mamoru/projects/consulting/page_guide.html

# [간편 진단] 13개 질문 조건부 분기 가위 진단 퀴즈 (코드위젯)
https://bsm-pixel.github.io/mamoru/projects/consulting/page_diag.html

# [상담 접수] 매장방문/출장 상담 접수 폼 (코드위젯)
https://bsm-pixel.github.io/mamoru/projects/consulting/page_form.html

# [맞춤 추천] 간편진단 결과 → 고객 맞춤 제품 추천 (코드위젯)
https://bsm-pixel.github.io/mamoru/projects/consulting/page_recommend.html

# [복원수리 안내] 복원수리 서비스 안내 풀페이지 (아임웹 /60)
https://bsm-pixel.github.io/mamoru/projects/as/page_guide.html

# [고객 후기] 리뷰 목록 표시 (코드위젯)
https://bsm-pixel.github.io/mamoru/projects/reviews/page_reviews.html
```

---

## 참고

- **GitHub Pages 배포**: `git push` → 자동 반영 (별도 빌드 없음)
- **아임웹 iframe**: `initIframeComm()` — ResizeObserver 우선, setInterval fallback
- **알림톡 링크**: GAS `GITHUB_PAGES_CONSULT` 변수에 프로토콜 미포함 (`bsm-pixel.github.io/mamoru/projects/consulting`)
  → 솔라피 템플릿 WL 버튼에서 `https://#{change_request_link}` 형태로 사용
