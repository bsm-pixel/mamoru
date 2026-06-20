# 복원수리 가이드 영상 (videos/)

이 폴더의 `.mp4`는 복원수리 안내 페이지에서 **상대경로 `videos/파일명.mp4`** 로 삽입됩니다.
새 영상은 **이 폴더에 정확한 파일명으로 넣고 커밋·푸시**하면 그 자리에 자동 표시됩니다(코드 수정 불필요, 단 신규 자리는 placeholder 교체 필요 — 아래 참고).

> 서빙 경로: `page.mamoru.kr/projects/as/videos/파일명.mp4` (GitHub Pages)
> 코드에서 영상 찾기: 페이지 HTML에서 **`[VIDEO]` 검색** → 모든 영상 위치가 주석으로 표시됨

## 영상 ↔ 코드 매칭표

| 파일 | 쓰이는 페이지 | 위치(섹션) | 코드 마커 |
|------|--------------|-----------|-----------|
| `복원수리안내히어로영상.mp4` | page_guide.html | 최상단 히어로(자동재생 루프) | `[VIDEO] 히어로 영상` |
| `전후영상1.mp4` | page_guide.html | **Before & After**(결과 영상) 1번 | `[VIDEO] 전후영상 1` |
| `전후영상2.mp4` | page_guide.html | **Before & After**(결과 영상) 2번 | `[VIDEO] 전후영상 2` |
| `전후영상3.mp4` | page_guide.html | **Before & After**(결과 영상) 3번 | `[VIDEO] 전후영상 3` |
| `전후영상6.mp4` | page_guide.html | **Before & After**(결과 영상, "복원수리 후 모발테스트") | `[VIDEO] 전후영상 6` |
| (없음) 전후영상4 | page_guide.html | Before & After 4번 = **"영상 준비중" placeholder** | `[VIDEO] 전후영상 4` |
| (없음) 전후영상5 | page_guide.html | Before & After 5번 = **"영상 준비중" placeholder** | `[VIDEO] 전후영상 5` |
| (없음) 작업영상1~3 | page_guide.html | **PREVIEW**(작업 과정 클립) = **"영상 준비중" placeholder** | `[VIDEO] 작업영상 1~3` |
| `고무줄.mp4` | page_guide.html · page_as_guide.html | 포장 안내 Step 1(고무줄 고정, 강조) | `[VIDEO] 포장 Step 1 고무줄` |

> **구조(2026-06-20 개편)**: **PREVIEW = 작업 과정 클립**(과정 티저, `작업영상N.mp4`) / **Before & After = 결과 증거**(전후 영상 + 전후 사진 통합).

> page_guide.html(아임웹 삽입용 메인)과 page_as_guide.html(별도 안내)이 **같은 videos/ 폴더 공유**.

## 영상 추가/교체하는 법

**기존 영상 교체**: 같은 파일명으로 이 폴더에 덮어쓰기 → 커밋·푸시. 끝.

**새 자리(전후영상 4·5)에 영상 넣기**:
1. mp4를 이 폴더에 `전후영상4.mp4` / `전후영상5.mp4` 로 넣는다.
2. page_guide.html 에서 `[VIDEO] 전후영상 4` 검색 → 그 아래 **"영상 준비중" placeholder div**를
   아래 형태로 교체:
   ```html
   <video autoplay muted loop playsinline src="videos/전후영상4.mp4"></video>
   ```
   (1·2·3번 코드 그대로 복사해서 숫자만 바꾸면 됨)
3. 커밋·푸시.

## 규칙
- **파일명은 한글 그대로** 가능(현재 방식 유지). 단 띄어쓰기·특수문자는 피한다.
- 자동재생 영상은 `autoplay muted loop playsinline` 4속성 필수(모바일 자동재생 조건).
- 용량 큰 원본은 압축 후 올린다(모바일 로딩 — 가능하면 1080p·H.264·~10MB 이하 권장).
