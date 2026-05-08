# CL1-4T 이미지 폴더

사장님이 채팅에 보여주신 사진 6컷을 아래 파일명으로 저장해주세요.

## 파일명 가이드 (v8_trendy 매핑)

| 파일명 | 사장님 사진 | 위치 (HTML) | 비율 |
|--------|-----------|------------|------|
| `hero.jpg` | #2 — 전체 사선 | Hero 풀폭 메인 | 4:5 |
| `handle.jpg` | #1 — 손잡이 디테일 | 디테일 그리드 ① | 1:1 |
| `bearing.jpg` | #3 — 베어링 클로즈업 | 디테일 그리드 ② | 1:1 |
| `front.jpg` | #4 — 정면 단축 | 디테일 그리드 ③ | 1:1 |
| `vertical.jpg` | #5 — 세로 디테일 | 디테일 그리드 ④ | 1:1 |
| `alt.jpg` | #6 — 반대 사선 | "이 가위에 대하여" 큰 사진 | 3:2 |
| `cut.gif` (추후) | 실제 커트 영상 | "IN ACTION" 섹션 | 16:9 |

## 추가 권장 사양

- **포맷**: JPG (사진), GIF/MP4 (영상)
- **해상도**: 사장님 보유 원본 그대로 OK (3000~3700px)
- **용량 한도**:
  - 단일 사진 500KB 이하 권장 (TinyPNG 또는 Squoosh로 압축)
  - GIF 3MB 이하 / MP4 5MB 이하

## 진행 순서

1. Windows 탐색기에서 이 폴더(`projects/brand/product_detail/CL1-4T/images/`) 열기
2. 사장님 사진 6장을 위 파일명으로 복사
3. 클로드에게 "넣었어" 알리기
4. 클로드가 v8_trendy.html의 placeholder를 `<img src="...">`로 교체 + push
5. GitHub Pages 배포 완료 후 미리보기에서 실제 사진 확인

## TinyPNG / Squoosh로 압축 방법 (선택)

원본 그대로 GitHub에 올려도 되지만 — 모바일 로딩 ↑하려면 압축 권장:

- **TinyPNG**: https://tinypng.com — JPG/PNG 드래그 → 자동 압축 → 다운로드
- **Squoosh**: https://squoosh.app — 더 정밀 조정 가능 (품질 80~85% 권장)

압축 시 화질 거의 유지 + 용량 50~70% 감소.
