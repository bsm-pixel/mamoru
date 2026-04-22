# TMS 아임웹 배너/팝업 원격 관리 흐름

> **최종 업데이트**: 2026-04-22 (Phase 2 슬라이드 기능 완료)
> **상태**: Phase 1 + Phase 2 가동 중
> **위치**: TMS 설정 → 알림·연동 탭 → 📢 아임웹 배너/팝업 관리

---

## 1. 목적

사장님이 **아임웹 관리자 페이지를 건드리지 않고** TMS에서 메인 모달 배너를 즉시 관리:
- 추석/연휴 배송지연 공지
- 이벤트·프로모션 배너
- 매장 휴무 안내
- 사이트 점검 안내

---

## 2. 아키텍처 — 하이브리드 방식

### 2-1. 데이터 흐름
```
사장님 TMS UI에서 저장
   ↓ PATCH /api/imweb/banners
Supabase imweb_banners 테이블
   ↓
고객이 아임웹 접속
   ↓
Footer Code에 주입된 <script src=".../banner-widget.js" defer>
   ↓
widget.js가 DOMContentLoaded 후 /api/imweb/banner-config 호출
   ↓
enabled=true + 기간 내 → 모달 DOM 생성
```

### 2-2. 핵심 원칙
- **SSOT**: TMS Supabase DB가 진실의 원천
- **아임웹 의존 최소**: 초기 1회 script 태그만 주입, 이후 아임웹 건드리지 않음
- **캐시 전략**: widget.js (1시간) / config JSON (60초)
- **MAMORU Brand Guide 준수**: 모노크롬 팔레트, Noto Sans KR, 여백

---

## 3. 파일 구조

```
app/src/
├── lib/imweb/script-api.ts                # 아임웹 Script API 클라이언트 (upsertMamoruWidget)
│
├── app/api/imweb/
│   ├── banners/route.ts                   # GET/PATCH (관리자, 인증 필수)
│   ├── banners/upload/route.ts            # POST 이미지 업로드 (2MB 제한)
│   ├── banner-config/route.ts             # GET 공개 JSON (CORS, 60s 캐시)
│   ├── banner-widget.js/route.ts          # GET 공개 JS (Content-Type: application/javascript)
│   └── script-install/route.ts            # POST 자동 설치 (Script API)
│
├── components/settings/banner-settings.tsx  # TMS 설정 UI
└── hooks/use-banner.ts                     # react-query 훅

supabase/migrations/053_imweb_banners.sql   # DB 스키마
Supabase Storage: imweb-banners (public)    # 이미지 CDN
```

---

## 4. DB 스키마

```sql
imweb_banners
├── id TEXT PRIMARY KEY           -- 'main_modal' (MVP 고정)
├── enabled BOOLEAN               -- 노출 토글
├── title / description            -- 텍스트 (모든 슬라이드 공통)
├── images JSONB                   -- Phase 2: [{url, path, link_url?}, ...] 최대 5장
├── image_url / image_path         -- legacy (첫 번째 이미지 shortcut, backwards-compat)
├── link_url                       -- legacy (첫 번째 이미지 link_url)
├── starts_at / ends_at            -- 노출 기간 (NULL = 무기한)
├── dismiss_cookie_hours INT       -- '오늘 하루 보지 않기' 재노출 대기
└── updated_at / updated_by
```

### images JSONB 구조 (Phase 2)
```json
[
  { "url": "https://...", "path": "main_modal/xxx.png", "link_url": "https://..." },
  { "url": "https://...", "path": "main_modal/yyy.png", "link_url": "" },
  ...
]
```
- 배열 순서 = 슬라이드 순서
- 최대 5장 (애플리케이션 레이어 검증)
- 이미지 1장 = 정적 배너 / 2장+ = 자동 슬라이드

---

## 5. 아임웹 설치 (초기 1회)

### 옵션 A — 수동 (현재 사용 중)
**아임웹 관리자 → 환경설정 → Footer Code** 맨 아래:
```html
<!-- MAMORU 배너 위젯 — TMS에서 원격 관리 -->
<script src="https://app-eta-sandy-75.vercel.app/api/imweb/banner-widget.js" defer></script>
```
→ 공통 코드 파일: [`projects/common_code/footer_code.txt`](../../../common_code/footer_code.txt)

### 옵션 B — 자동 (Script API)
TMS 설정 UI에서 **unitCode 입력 + "자동 설치"** 버튼
- 내부 동작: GET /script → 있으면 PUT, 없으면 POST
- OAuth scope: `script:read script:write` 필요

---

## 6. 운영 시나리오

### 시나리오 A — "사이트 공사중" 긴급 배너
1. TMS → 설정 → 알림·연동 → 📢 아임웹 배너/팝업 관리
2. 이미지 업로드 (jpg/png/webp, ≤2MB)
3. 제목: "사이트 공사중"
4. 노출 토글 On → 저장
5. 1분 이내 아임웹 메인페이지 모달 노출

### 시나리오 B — 추석 배송지연 (기간 지정)
1. 시작일시: 추석 전주 월요일 09:00
2. 종료일시: 추석 다음주 월요일 09:00
3. 토글 On + 저장
4. 기간 내에만 자동 노출

### 시나리오 C — 프로모션 + 클릭 이동
1. 이미지 업로드
2. 클릭 링크: `https://mamoru.kr/event/spring`
3. 고객이 배너 클릭 시 해당 페이지 새 탭 오픈

---

## 7. UX 보호 장치

| 장치 | 설명 |
|------|------|
| 저장 전 미리보기 | 고객 시점으로 확인 가능 |
| **과거 종료일 경고 모달** | ends_at이 현재보다 과거면 저장 직전 경고 (2026-04-22 추가) |
| 이미지 크기 제한 | 2MB (업로드 시 검증) |
| MIME 검증 | jpg/png/webp만 |
| 쿠키 기반 닫기 | "오늘 하루 보지 않기" — 재방문 시 사용자 경험 보호 |
| XSS 방지 | widget.js에서 escapeHtml |
| 중복 로딩 방지 | `window.__MAMORU_BANNER_LOADED` 가드 |

---

## 8. Phase 2 슬라이드 기능 (2026-04-22 완료)

### 스펙
| 항목 | 값 |
|------|-----|
| 최대 이미지 수 | **5장** |
| 1장 모드 | 정적 배너 (기존과 동일) |
| 2장+ 모드 | **자동 슬라이드** |
| 자동 전환 간격 | **5초 고정** |
| 루프 | 마지막 → 첫 번째 순환 |
| 호버 동작 | 자동 전환 일시정지 |
| 수동 이동 | 점 네비 클릭 / 화살표 / 스와이프 → 타이머 리셋 |

### 인터랙션
| 인터랙션 | 방식 |
|---------|------|
| **점(dot) 네비게이션** | 하단 가운데, 현재 슬라이드는 길쭉한 검정색 |
| **데스크톱 화살표** | 좌우 < > 버튼 (마우스 hover 시) — `@media (hover:hover) and (pointer:fine)` |
| **모바일 스와이프** | touchstart/touchend 40px 임계값 |
| **이미지별 개별 링크** | 각 슬라이드 클릭 시 고유 link_url 새 탭 열림 |

### TMS UI
- 이미지 그리드 (썸네일 + 번호 뱃지 + 링크 입력)
- 위/아래 화살표로 순서 변경 (DB 즉시 반영)
- 개별 이미지 삭제 버튼
- "+ 이미지 추가" 버튼 (5장 도달 시 disabled)
- 2장+ 시 "🎠 슬라이드 모드 · 5초 자동 전환" 배지 표시
- 미리보기 모달에서 실제 슬라이드 동작 시뮬레이션

### 제목/설명 정책
- **모든 슬라이드에 동일한 제목/설명**이 하단 본문 영역에 표시
- 이유: 슬라이드별 텍스트 오버레이는 브랜드 톤 관리 복잡 → 이미지 내 텍스트 권장

---

## 9. 추가 확장 로드맵 (미구현)

- 복수 배너 슬롯 (main_modal + top_strip + side_notice)
- 노출 통계 (노출수/클릭수/닫기율)
- 특정 페이지(메인/카테고리)별 분기
- 스케줄 자동화 (추석 D-7 자동 활성화 등)
- 슬라이드 전환 속도 커스터마이징 (현재 5초 고정)

---

## 9. 장애 대응

| 증상 | 원인 | 해결 |
|------|------|------|
| 배너 안 뜸 (enabled=false 응답) | 토글 Off / 기간 지남 / 미저장 | TMS 재확인 |
| widget.js 로드 실패 | Footer Code 주입 누락 | 아임웹 환경설정 → Footer Code 확인 |
| 이미지 표시 안 됨 | Storage 버킷 public 미설정 | Supabase Dashboard → Storage → imweb-banners → public |
| 저장 후 반영 지연 | config 60초 캐시 | 1분 대기 or 시크릿 창 새로고침 |
| 자동 설치 시 scope 부족 | OAuth에 `script:write` 없음 | 아임웹 OAuth 재연결 |
