# TMS 자동 점검 (2026-11-15~)

> 화면을 고칠 때마다 핵심 흐름이 깨지지 않았는지 몇 초 만에 확인한다. 운영 자료·알림톡·외부 서비스에 닿지 않는다.

## 실행

```
cd projects/Total_Management_System/app
npm run check
```

## 점검하는 것

| 파일 | 내용 |
|---|---|
| `tests/flows/sale-cancel.test.ts` | 판매 등록 → 취소 → 재고·고객 미수금·매출이 등록 전으로 |
| `tests/flows/inventory-adjust.test.ts` | 재고 조정 +/-, 재고보다 많이 줄이기 차단, 로그인 없으면 거부 |
| `tests/flows/purchasing.test.ts` | 발주 생성(금액·부가세) → 취소해야만 삭제 → 재고 불변 |
| `tests/flows/customers.test.ts` | 고객 등록·같은 전화 차단·수정, 병합 요청값 |
| `tests/flows/repair-status.test.ts` | 복원수리 상태 변경·이력, 건너뛰기 차단, 외부 호출 0 |
| `tests/unit/rules.test.ts` | 금액 공식, 복원수리 상태 규칙 |

## 원리

- `tests/helpers/fake-db.ts` — 수파베이스 조회 문법을 메모리에서 흉내 낸 가짜 저장소. 지원 안 하는 문법은 오류를 낸다(거짓 통과 방지).
- `tests/helpers/setup.ts` — 수파베이스 접속을 가짜로 바꾸고, 알림톡·푸시는 "보내려 했다"만 기록, 그 밖의 외부 호출(fetch)은 전부 차단.

## 못 잡는 것 (알고 쓰기)

- 데이터베이스 안에서 도는 계산: 고객 병합(`merge_customers`), 시리얼 맞바꾸기, 집계 함수, 트리거, 행 보안 규칙
- 화면 모양·버튼 동작 (빌드·눈 확인으로 봄)
- 실제 알림톡 도착, 롯데·아임웹 응답

## 같이 쓰는 점검

- `py scripts/anon-probe.py` — 공개 키만으로 읽히는 표가 생겼는지 (0이어야 정상)
