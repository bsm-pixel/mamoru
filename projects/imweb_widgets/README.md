# 🧩 MAMORU 아임웹 커스텀 위젯 모음

아임웹 **디자인모드 → 커스텀 위젯**(HTML/CSS/JS 3탭 + 패널)에 붙여넣어 쓰는 위젯들입니다.
각 위젯 폴더 안에 `widget.html` / `widget.css` / `widget.js`(3탭에 각각 붙여넣기) + `README.md`(사용설명서)가 있습니다.

> 만드는 법: 아임웹 디자인모드 → 커스텀 위젯 새로 만들기 → **HTML 탭에 `widget.html`, CSS 탭에 `widget.css`, JavaScript 탭에 `widget.js`** 붙여넣기 → 우측 패널에서 사진·문구 등 입력 → 미리보기 확인 → 게시 → 사이트 연동.

---

## 📋 위젯 카탈로그

| # | 위젯 | 한줄 가치 | 상태 |
|---|------|----------|------|
| 01 | 복원 Before/After 슬라이더 | 자체복원 기술을 시각적으로 압도 | ✅ 완료 |
| 02 | 한정세일 카운트다운 | 마감 임박 긴박감 → 전환↑ | ✅ 완료 |
| 05 | 정적 누적 카운터 | 신뢰·규모감 (복원 N건) | ⏳ |
| 06 | 후기 캐러셀 | 사회적 증거 | ⏳ |
| 11 | 가위 추천 미니진단 | 인터랙티브 전환 도구 | ⏳ |
| 03 | 가위 스펙 비교표 | 구매 결정 가속 | ⏳ |
| 04 | 등급별 견적 계산기 | B2B 리드 질↑ | ⏳ |
| 12 | 360° 회전 뷰어 | 제품 디테일 신뢰 | ⏳ |
| 13 | 마감/컬러 옵션 미리보기 | 구매 확신 | ⏳ |
| 07 | 가위 관리법 가이드 | 전문성·체류↑ | ⏳ |
| 08 | 시술별 가위 선택 가이드 | 진입장벽↓ | ⏳ |
| 09 | 미용가위 용어사전 | 교육·SEO | ⏳ |
| 10 | 영업시간 배지 + 카톡 | 상담 진입 명확화 | ⏳ |
| 14 | 스크롤 절단 히어로 | 브랜드 첫인상 | ⏳ |
| 15 | 딜러/아카데미 게이트 | B2B 정보 분리 | ⏳ |
| 16 | 오늘의 추천 모델 | 재방문 유도 | ⏳ |
| T2 | 복원 인터랙티브 타임라인 | 과정 투명성=신뢰 | ⏳ |
| T3 | 가위 해부도 핫스팟 | 전문성·교육 | ⏳ |
| T6 | 라이트박스 사례 갤러리 | 복원 사례 몰입 | ⏳ |
| T1 | 스크롤 스토리텔링 | 브랜드 철학 몰입 | ⏳ |
| T5 | 무한 마퀴 띠 | 핵심 키워드 흐름 | ⏳ |
| T4 | 3D 틸트 제품 카드 | 프리미엄 인터랙션 | ⏳ |
| T7 | 마우스 스포트라이트 히어로 | 임팩트 | ⏳ |
| T8 | 타이핑/스크램블 헤드라인 | 슬로건 임팩트 | ⏳ |
| T9 | 읽기 진행바 + 섹션 네비 | 긴 페이지 UX | ⏳ |
| T10 | 그립/손크기 시뮬레이터 | 전환 | ⏳ |

---

## ⚙️ 공통 규격 (아임웹 커스텀 위젯)

### 🚫 금지 (= "유효하지 않은 코드"로 저장 거부)
- `fetch` / `XMLHttpRequest` (외부 네트워크 요청) → **라이브 데이터 연동 불가**
- `<iframe>` 삽입 → 인라인 유튜브 재생 불가
- 외부 라이브러리 / CDN → **바닐라 JS만**
- `localStorage` / 쿠키 → 방문 간 기억 불가
- 아임웹 내부데이터(회원·상품)·쇼핑·위젯 간 통신

### 🎛️ 패널 설정값 (Handlebars annotation)
HTML 탭에 주석으로 선언 → `{{변수명}}`으로 사용. **{{변수}}는 HTML/CSS/JS 3탭 모두 치환**되지만, JS에서는 안전하게 **HTML의 `data-*` 속성에 심어 읽는 방식**을 표준으로 한다.

```html
{{!-- @name title @type outlined-textfield @default "제목" @label "제목" --}}
<h2>{{title}}</h2>
```

**@type 토큰 (검증됨)**

| 입력 종류 | @type 토큰 | 비고 |
|-----------|-----------|------|
| 텍스트 한 줄 | `outlined-textfield` | 최대 100자 (숫자형도 가능) |
| 줄글(서식) | `text-editor` | 최대 2,000자 |
| 드롭다운 | `select` | 옵션 최대 10 |
| 옵션버튼/탭 | `segmented-control` | 옵션 최대 4·각 8자 |
| ON/OFF | `switch` | |
| 색상 | `color-picker` | |
| **이미지 업로드** | **`image`** | ★ 100MB·8000px 이하 (image-uploader 아님) |
| 날짜 | `date-picker` | 미래만 |
| 시간 | `time-picker` | |
| 반복 항목 | `item` | `{{#each name}}…{{/each}}`, 부모1·자식20·속성20 |

부가 속성: `@default` `@label` `@placeholder` `@values 값1,값2,값3` `@valueNames 라벨1,라벨2,라벨3` (값에 공백 없이).

### 📐 제약
- **탭당 1만 자** (HTML/CSS/JS 각각)
- 12컬럼 100% 너비 기본 + 모바일 반응형(clamp 권장)
- 전역 스타일 격리 → CSS에 **브랜드 색 hex 직접** (공통 CSS 상속 안 됨, 폰트만 미선언시 상속)

---

## 🎨 브랜드 토큰 (모든 위젯 CSS에 직접 사용)

```
/* 모노크롬 팔레트 (포인트컬러·그라데이션·장식 금지) */
void #1A1A1A · graphite #2D2D2D · stone #4A4A4A
warm-gray #8A8580 · mist #B8B4AF · sand #D4D0CB
parchment #EDEBE8 · shell #F5F3F0 · cream #FAF9F7 · white #FFFFFF

/* 폰트 */
디스플레이: 'Outfit','Plus Jakarta Sans','Noto Sans KR',sans-serif  (700/800/900)
본문:      'Plus Jakarta Sans','Pretendard','Noto Sans KR',-apple-system,sans-serif

/* 반경 */ 8 / 12 / 16 / 20 px
/* 이즈 */ cubic-bezier(.4,0,.2,1)
/* 타이포 clamp */ 제목 clamp(22px,6vw,30px) · 본문 clamp(14px,3.8vw,16px) · 라벨 11~12px
/* 모션 */ translateY(24px)+opacity, 진입 시 1회 은은하게
```
> ⚠️ Outfit이 사이트에 로드 안 돼 있으면 시스템 폰트로 폴백됩니다(브랜드상 허용). 정확히 맞추려면 아임웹 SEO > body code에 Google Fonts 링크 삽입.

---

## 🧱 새 위젯 만들 때
`_TEMPLATE/` 폴더를 복사해 `NN_이름/`으로 만들고 3탭 + README를 채운다. 빌드 표준:
- HTML: 주석박스 → 패널 annotation → 시맨틱 마크업 + `data-*`
- CSS: 브랜드 hex·clamp·여백·subtle 모션
- JS: 바닐라 IIFE, `querySelectorAll` 전 인스턴스 init, `data-*` 읽기, 금지 API 0
- IA 체크: 올바른 @type·합리적 기본값·빈상태·터치+키보드·반응형
