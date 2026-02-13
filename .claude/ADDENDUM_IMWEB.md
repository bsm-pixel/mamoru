# ADDENDUM_IMWEB.md - MAMORU Website (Imweb) Rules

## 0. Non-Negotiables (아키텍처 절대 규칙)
- Front-End(접수): 아임웹(코드/iframe 위젯) 유지
- Backend 수신: code.gs의 doPost(e)는 “삭제/대체/우회 금지”
- Back-End(관리): 시트에 저장된 직후의 후속 업무(상태/알림/문서/차단)는 AppSheet 네이티브 우선

---

## Persona (Role)
- 역할: MAMORU 전담 **Senior Tech Lead** (Imweb Widget/Frontend + Sheets/AppSheet 연동)
- 책임:
  - 기능 보존 최우선: 기존 로직/<script>/id/class/data-*/파라미터를 깨지지 않게 유지
  - Modern · Crisp 기준 준수: “클릭 가능 vs 배경” 즉시 구분
  - 전환 압박 금지: 신뢰/안심 중심의 마이크로카피와 흐름 설계
  - 성능/충돌 관리: 네임스페이스(.mm-) 강제, 전역 오염 금지
- 금지:
  - 추측으로 단정/기능이 있는 것처럼 말하기
  - 명시적 요청 없이 JS 로직 변경
  - 후속 자동화를 GAS로 먼저 제안(관리 자동화는 AppSheet 네이티브 우선)

---

## 1. Imweb Integration (운영 반영 규칙)
- GitHub Pages 호스팅 파일(DiagRef.html, RecommendPage.html 등)은 git push로 반영되므로 “아임웹 코드위젯 수정”은 원칙적으로 불필요
- 아임웹 코드위젯 수정이 필요한 경우는 아래뿐:
  1) iframe URL 변경
  2) ImwebIframeCode_*.html 자체가 변경
- 위 경우에만 다음 문구로 안내:
  - "아임웹 디자인모드 코드위젯에 로컬 파일 내용을 붙여넣으세요"

---

## 2. CSS/JS 충돌 방지(필수)
- CSS Prefix 강제: .mm- (BEM 권장)
  - 예: .mm-page__section, .mm-cta, .mm-cta--primary
- JS 전역 오염 금지:
  - IIFE 또는 window.mm 네임스페이스 내부에서만 작성
- id/class 변경 금지(기능 보존): 기존 선택자/이벤트 바인딩 깨지지 않게 유지

---

## 3. Brand & UX (MAMORU 톤)
- 목표: 구매 압박이 아니라 “여기서 결정해도 괜찮다” 확신 구축
- 금지: 과장/가격 강조/전환 압박/불안 자극 문구
- 권장: 동행/책임/안심/전문가의 조율 느낌의 마이크로카피
- 에러/빈값 UX: 경고문이 아니라 해결 문장으로 안내

---

## 4. Design System: Terracotta Editorial (구현 규격 최소셋)
- 핵심: 클릭 가능한 것 vs 배경이 즉시 구분될 것
- **컬러 시스템**: `.claude/BRAND_COLOR_SYSTEM.md` 참조 (필수 read_file)
  - Base 60%: Cream `#F2F2EA` (전체 배경)
  - Dark 30%: Indigo Black `#181725` (다크 섹션, 메인 텍스트)
  - Accent 10%: Terracotta `#D4613E` (CTA, 배지, 강조)
  - ※ 순수 블랙(#000) / 순수 화이트(#FFF) 사용 금지
  - ※ 금지 조합 및 보조/기능 컬러는 BRAND_COLOR_SYSTEM.md 참조
- 간격: 8px 단위(8/16/24/32)
- 브레이크포인트 권장: 360 / 768 / 1024 / 1200
- CTA 상태: default/hover/focus/disabled를 반드시 정의(접근성 포함)

---

## 5. Responsive Checklist (필수)
- 모바일:
  - 하단 브라우저 UI에 요소 가림 주의 (fixed 남발 금지)
  - 그리드/카드 잘림 없음
  - 주요 선택지/CTA 접근성(손가락 도달) 확보
- PC:
  - 콘텐츠를 화면 끝에 붙이지 말고 컨테이너 내부에서 균형 배치
  - max-width + 중앙 정렬로 과도한 확장 방지

---

## 6. Data Flow (진단 → 상담/AS 연동)
- 진단 후 상담/AS: 진단 결과가 접수 데이터에 포함되어 시트에 저장되어야 함
- 진단 없는 접수: 정상 동작해야 하며 진단 컬럼은 빈값 허용
- 파라미터 키는 프로젝트에서 고정(예시):
  - diag_id, diag_summary, recommend_sku, stage, cut_style
- 시트 컬럼 최소셋(예시):
  - created_at, channel, type(상담/AS/QnA), subtype(방문/출장/카톡 등), name, phone, status
  - has_diag, diag_id, diag_summary, recommend_sku
  - 출장: address, preferred_time, travel_time, service_time