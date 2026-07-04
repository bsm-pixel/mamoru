# 🎨 디자인 작업실 (projects/_design_lab)

사장님 + 클로드가 **고객 페이지(HTML) 디자인을 빠르게 검토·수정**하는 공간.
`projects/` 바로 밑에 둬서 **딱 보이고**, 파일을 **브라우저로 바로 열어** 확인합니다(빌드 불필요).

> 마케팅 · 프로덕트 페이지 · 리뷰 · 복원수리 안내 등 **모든 standalone HTML 페이지 작업**을 여기서.

## 쓰는 법
1. 이 폴더의 `*.html` 파일을 **브라우저로 직접 열기** (더블클릭 = `file://`)
   - 또는 VS Code "Live Server"로 열면 저장 시 자동 새로고침
2. 클로드가 수정하거나 사장님이 직접 수정 → **브라우저 새로고침(F5)** 하면 반영
3. 확정되면 실제 페이지(`projects/as/…`, `projects/main/…` 등)에 옮기고, 여기 작업 파일은 정리

## 🗺️ 페이지 허브 (전체 화면 지도)
- **`index.html`** — 모든 고객 페이지 + 관리자(TMS) 화면을 **한눈에 보고 라이브로 바로 여는 허브**. 검색 필터 포함.
  - 브라우저로 열기 or `https://page.mamoru.kr/projects/_design_lab/index.html`
  - "그 페이지 어디였지?" 헤맬 때 여기부터. (낡은 `projects/PAGES_INDEX.md` 대체)

## 현재 작업 파일
| 파일 | 내용 |
|------|------|
| `index.html` | **페이지 허브** — 전체 화면 인덱스(고객+TMS) + 검색 |
| `_page_kit.html` | 아임웹 전 페이지 통합 표준 — 복사 원본 |
| `as_principle.html` | 복원수리 안내 — '원칙' 카피 개편 (문제→답). A안 1열(모바일) / B안 지그재그 비교 |

## 규칙 (IA — 위치 헷갈리지 않게)
- **HTML 페이지 검토 = 여기(`projects/_design_lab/`)**. 브라우저로 바로 열림.
- **TMS(관리자 앱) 화면 검토 = `/design-lab`** (Next.js 앱 안, `app-eta-sandy-75.vercel.app/design-lab`). React라 앱 안에 있어야 함.
- 즉 "**고객 HTML 페이지 → 작업실 / 관리자 React 화면 → TMS 디자인랩**" 두 갈래. 헷갈리면 이 표만 보면 됨.
