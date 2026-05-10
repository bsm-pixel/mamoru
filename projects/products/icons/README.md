# MAMORU 라인 일러스트 아이콘 (SVG)

v10 PROFILE 섹션의 BLADE / BLADE BACK 카드에서 사용 중인 SVG 아이콘 6개.

## 파일 목록

### BLADE — 날 부분 형태 (블런트/장가위/드라이 공통)
| 파일 | 의미 | 패스 |
|------|------|------|
| `blade-f.svg` | F (Force) — 직선형 | 두 직선 V자 |
| `blade-a.svg` | A (All-round) — 밸런스형 | 두 날 살짝 곡선 (Q 28,8 / 28,24) |
| `blade-g.svg` | G (Glide) — 둥근형 | 두 날 부드러운 곡선 (Q 22,14 / 22,18) |

### BLADE BACK — 날등 형태 (블런트/장가위/드라이 공통)
| 파일 | 의미 | 패스 |
|------|------|------|
| `blade-back-s.svg` | S (Straight) — 직선형 날등 | 위/아래 직선 평행 |
| `blade-back-k.svg` | K (Ken-form) — 검형 날등 | 위 곡선 (Q 30,11 → 44,6) + 아래 직선 |
| `blade-back-b.svg` | B (Byeol-form) — 별형 날등 | 위 꺾임 (L 24,4 → 44,11) + 아래 직선 |

## 사양

- **viewBox**: 48 × 32 (가위 비율)
- **stroke**: currentColor (CSS color로 동적 변경 가능)
- **stroke-width**: 1.6
- **stroke-linecap**: round
- **fill**: none

## Figma에서 사용

1. Figma에서 **File → Place image** 또는 SVG 파일 드래그
2. import된 SVG를 **Component**로 변환
3. 라인 굵기/색상 변경 가능 (사장님 톤으로 개량)
4. 추후 사장님 개량본을 v10 코드의 inline SVG와 교체

## 향후 확장

추가 아이콘 후보:
- 핸들 형태 (A 핸들 / B 핸들 / C 핸들) — 두 고리 + 연결선
- 무게감 (저울 마커)
- 가위 종류 카테고리 (블런트/장가위/틴닝/슬라이싱) — 시각 구분용
- 등급 (R/A/E/S) — 거대 타이포 대신 시각 아이콘

사장님 명세 후 추가 가능.
