# N7. 에디토리얼 섹션 헤더

## 기능 / 가치
거대한 인덱스 번호 + 키커 + 제목 + hairline로 구성된 에디토리얼 섹션 구분자. 페이지에 잡지 같은 위계·리듬을 부여(Brand Guide B-08).

## 고객 경험
"01 / SECTION / 섹션 제목" 이 큰 숫자와 함께 진입 시 은은하게 떠오름.

## 아임웹 삽입 방법
1. 커스텀 위젯 → HTML/CSS/JS 3탭 붙여넣기
2. 섹션마다 인덱스·키커·제목만 바꿔 반복 사용 → 게시

## 패널 설정 항목
| 항목 | @type | 기본값 | 설명 |
|---|---|---|---|
| 인덱스 | outlined-textfield | 01 | 큰 번호 |
| 키커(영문) | outlined-textfield | SECTION | |
| 제목 | outlined-textfield | 섹션 제목 | |
| 정렬 | segmented-control | 왼쪽 | 왼쪽/가운데 |

## 주의사항
- 여러 섹션 위에 반복 배치하면 일관된 위계가 생김. fetch·iframe 미사용.
