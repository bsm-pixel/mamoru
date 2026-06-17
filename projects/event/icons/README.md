# EVENT 접수페이지 아이콘

이 폴더에 SVG를 넣으면 `page_form.html`이 자동으로 표시합니다.
(파일이 없으면 임시로 ✕ 가 보입니다. 올리고 push하면 즉시 반영 — GitHub Pages)

## 파일명 (정확히)

| 용도 | 파일명 |
|------|--------|
| 오른손 | `hand-right.svg` |
| 왼손 | `hand-left.svg` |
| 종류 — 블런트 | `cat-blunt.svg` |
| 종류 — 틴닝 | `cat-thinning.svg` |
| 종류 — 장가위 | `cat-long.svg` |
| 종류 — DRY | `cat-dry.svg` |
| 가격 5만 | `price-50000.svg` |
| 가격 10만 | `price-100000.svg` |
| 가격 20만 | `price-200000.svg` |

## 규칙
- **가격 파일명 = 실제 금액(원) 숫자** 그대로. 예) 7만원 품목이면 `price-70000.svg`.
- 권장: 정사각 SVG (손/종류 ~64×64, 가격 ~64×64). 표시는 CSS가 각각 30/22/26px로 맞춤.
- 경로: `projects/event/icons/<파일명>` → 서빙 URL `page.mamoru.kr/projects/event/icons/<파일명>`
