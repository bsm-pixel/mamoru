# 📚 MAMORU 매뉴얼·문서 인덱스 (전체 위치 지도)

> "그 문서 어디 있더라?" 방지용 단일 인덱스. 알림톡·흐름도·기능 매뉴얼이 3곳(루트 docs / TMS docs / consulting / memory)에 흩어져 있어 여기서 한 번에 찾는다.
> 최종 정리: 2026-07-31

---

## 🔔 알림톡 (솔라피 + Make) — 카테고리별

| 카테고리 | 문서 | 내용 |
|---|---|---|
| **전체 카탈로그(SSOT)** | `memory/reference_solapi_templates.md` | 운영 26 + 직접방문 5 = **31종** 트리거·변수·웹훅·코드위치. **알림톡 작업 전 무조건 먼저 read** |
| **복원수리 매장방문(직접방문)** | [docs/AS_VISIT_ALIMTALK_TEMPLATES.md](AS_VISIT_ALIMTALK_TEMPLATES.md) | 접수완료·리마인드24·리마인드2·변경·취소 **5종** 양식+Make 필터+버튼+변수. 셀프 변경/취소 페이지 연동 |
| **이벤트(EVENT)** | [docs/EVENT_ALIMTALK_SETUP.md](EVENT_ALIMTALK_SETUP.md) | 접수완료·입금완료(+선택 입금안내) 양식+Make+웹훅 생성법+발송(sales_shipped) |
| **주문/판매(orders)** | `projects/Total_Management_System/docs/SOLAPI_TEMPLATES_ORDERS.md` | 주문/판매 계열 알림톡 |
| **상담·출장·톡 + 변경/취소** | `projects/consulting/FLOW_change_request.md` | 상담(매장/출장/톡) 전체 흐름 + 각 템플릿(field_confirmed/remind/cancelled/suggest 등) + 고객 셀프 변경/취소 페이지 |

**공통 규칙**
- 알림톡 변수 = TMS→Make payload **영문 키와 일치**. 미발송 진단은 [feedback_alimtalk_diagnosis_first]: ①솔라피 검수 ②Make 분기 ③토글.
- Make **웹훅 3개**: consultation(`MAKE_WEBHOOK_URL`, 상담+이벤트+판매출고+재고판매) / as_received(`MAKE_AS_RECEIVED_WEBHOOK_URL`, 복원수리 접수) / repair(`MAKE_REPAIR_WEBHOOK_URL`, 복원수리 상태변경).

---

## 🧭 TMS 기능별 사용 매뉴얼 — `projects/Total_Management_System/docs/`

| 문서 | 다루는 것 |
|---|---|
| `MANUAL_SALES.md` | 판매관리 (B2C·B2B 통합) |
| `MANUAL_REPAIR.md` | 복원수리 접수→수거→입고→수리→출고 |
| `MANUAL_CONSULTATION.md` | 상담(매장/출장/톡) 관리 |
| `MANUAL_EVENT.md` | 이벤트 접수→입금→판매전환 |
| `STOCK_SALE_MANUAL.md` | 재고판매(LS) |
| `MANUAL_ORDER.md` | 아임웹 주문 관리 |
| `MANUAL_CUSTOMER.md` | 고객(활동명·병합) |
| `MANUAL_CONTRACT.md` | 계약서 |
| `MANUAL_INVENTORY.md` · `MANUAL_SERIAL.md` | 재고·시리얼 무결성 |
| `MANUAL_PRODUCT.md` | 제품 |
| `MANUAL_SOURCING.md` | 샘플 소싱(1688) |
| `MANUAL_LABEL.md` | 라벨(제브라) |
| `MANUAL_DASHBOARD.md` | 대시보드 |

---

## 🔀 TMS 흐름도(데이터 흐름) — `projects/Total_Management_System/docs/`

> 코드 수정 시 **같은 커밋에 갱신** [feedback_flow_doc_update].

`TMS_FLOW_SALES` · `TMS_FLOW_REPAIR` · `TMS_FLOW_CONSULTATION` · `TMS_FLOW_EVENT` · `TMS_FLOW_ORDERS` · `TMS_FLOW_CUSTOMERS` · `TMS_FLOW_INVENTORY` · `TMS_FLOW_SERIAL` · `TMS_FLOW_SOURCING` · `TMS_FLOW_ACCOUNTING` · `TMS_FLOW_AUTO_DELIVERY`(배송추적·집하) · `TMS_FLOW_IMWEB_BANNER`

---

## 🏗️ TMS 구조·운영 참고 — `projects/Total_Management_System/docs/`

| 문서 | 내용 |
|---|---|
| `TMS_FILE_GUIDE.md` | 파일 구조 안내 |
| `TMS_PROCESS_MAP.md` · `TMS_INTEGRATION_FLOW.md` | 전체 프로세스·연동 지도 |
| `TMS_SYSTEM_ARCHITECTURE.md` | 시스템 아키텍처 |
| `TMS_ROADMAP.md` | 로드맵 (루트 `TODO.md`와 함께 유지) |
| `TEST_CHECKLIST.md` | 테스트 체크리스트 |
| `imweb-settings-checklist.md` · `imweb-code/` | 아임웹 설정·코드위젯 |
| 롯데 PDF 2종 | 롯데 주소정제·주문처리 API 가이드 (`projects/Total_Management_System/*.pdf`) |

---

## 📣 마케팅 · 기타

| 문서 | 내용 |
|---|---|
| `projects/marketing/릴스_ManyChat_세팅가이드.md` | 릴스→ManyChat DM 자동화 |
| `projects/_design_lab/index.html` | 전체 고객페이지·TMS 화면 인덱스(라이브 링크) |
| 루트 `TODO.md` | 진행중·대기 작업 단일 누적 |
| 루트 `📁_폴더안내_한글.md` · `projects/PAGES_INDEX.md` | 폴더·페이지 안내 |

---

## 🗂️ 큰 그림 (어디를 먼저 볼까)

- **알림톡 만들/고칠 때** → `memory/reference_solapi_templates.md`(카탈로그) → 해당 카테고리 문서(위 🔔표)
- **기능 사용법** → TMS `MANUAL_*`
- **데이터가 어떻게 흐르나** → TMS `TMS_FLOW_*`
- **어떤 파일이 뭐 하나** → `TMS_FILE_GUIDE.md` · `TMS_PROCESS_MAP.md`
- **화면/페이지 찾기** → `projects/_design_lab/index.html`
