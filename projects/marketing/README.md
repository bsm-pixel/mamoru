# MAMORU — 마케팅 도구 모음

고객 데이터 수집, 분석, 백업을 위한 자동화 스크립트 저장소.

---

## 📂 구성

| 파일 | 용도 |
|------|------|
| [naver_review_extract.js](naver_review_extract.js) | 네이버 스마트플레이스 리뷰 일괄 추출 |

---

## 🔍 네이버 리뷰 일괄 추출

### 용도
네이버 스마트플레이스의 리뷰를 메타데이터/사진/영상 포함하여 통째로 백업.
- 베스트 리뷰 선정 작업
- TMS 리뷰 시스템 이관
- 브랜드 후기 아카이브

### 사용법

#### 1단계: 네이버 페이지 준비
1. 크롬에서 [스마트플레이스 관리자](https://new.smartplace.naver.com) 로그인
2. **리뷰 관리** 메뉴 이동
3. **스크롤 끝까지 내리기** — 모든 리뷰가 DOM에 로드될 때까지
   - "더보기" 버튼이 있으면 전부 클릭하여 펼쳐두기

#### 2단계: Console에서 스크립트 실행
1. **F12** 키 → **Console** 탭
2. [`naver_review_extract.js`](naver_review_extract.js) 파일을 **VS Code에서 열어 Ctrl+A → Ctrl+C** (전체 복사)
3. Console에 **Ctrl+V** 후 **Enter**
   - ⚠ 네이버가 "붙여넣기 허용?" 경고 띄우면 `allow pasting` 입력 → Enter → 다시 붙여넣기
4. 자동 실행 → `naver_reviews_YYYY-MM-DD.zip` 다운로드됨

Console에 이렇게 표시되면 성공:
```
✅ 파싱 완료: 167건
📦 ZIP 파일 생성 중...
✅ 완료! naver_reviews_2026-04-18.zip 다운로드됨
⚠ CORS로 실패 128건 → download_images.ps1 실행하세요
```

#### 3단계: 이미지 다운로드 (PowerShell)
브라우저는 CORS 정책 때문에 네이버 이미지를 직접 받을 수 없음. PowerShell로 우회:

1. ZIP을 원하는 폴더에 **압축해제**
2. 압축 푼 폴더 안에서 **빈 공간 Shift+우클릭** → **"여기에서 PowerShell 창 열기"**
3. 실행:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\download_images.ps1
   ```
4. 진행률이 이렇게 표시됨:
   ```
   Downloading 128 files...
     [OK]   [1/128] 001_2024-09-10_김_관/photo_01.jpg
     [OK]   [2/128] 001_2024-09-10_김_관/photo_02.jpg
     ...
   Done!
     Success: 128 / Failed: 0
   ```

### 결과물

```
naver_reviews_2026-04-18/
├── reviews.json            # 전체 리뷰 메타데이터 (통합)
├── _failed.json            # 다운로드 실패 URL 상세 (디버그용)
├── urls.tsv                # PowerShell이 읽는 경로+URL 목록
├── download_images.ps1     # PowerShell 다운로드 스크립트
├── _README.txt             # ZIP 내 사용 안내
│
├── 001_2024-09-10_김_관/   # 고객별 폴더
│   ├── review.md           # 본문+태그(이모지)+메타데이터 (마크다운)
│   ├── photo_01.jpg
│   └── photo_02.jpg
│
├── 002_2024-09-08_이_희/
│   ├── review.md
│   └── photo_01.jpg
...
```

#### review.md 내용 예시
```markdown
# 김*관 고객 리뷰

## 메타데이터
- **작성일**: 2024-09-10
- **방문일**: 2024-09-10 17:00
- **상태**: 완료
- **업체**: 마모루 미용가위
- **서비스**: 가위 컨설팅상담 (마모루 사무실 방문)
- **출처**: 네이버 스마트플레이스

## 키워드
- 😊 품질이 좋아요
- 🔍 A/S가 세심해요
- 💗 친절해요

## 리뷰 본문
자격증을 따고 첫가위 구매를 마모루에서 상담받고...

## 사진
- photo_01.jpg
- photo_02.jpg
```

---

## 🛠 기술 구조

### 왜 이런 방식?
1. **브라우저 Console 스크립트** — 네이버 로그인 세션을 그대로 활용 (API 토큰 불필요)
2. **JSZip** — 폴더 구조 그대로 ZIP 생성 (CDN 동적 로드)
3. **CORS 우회** — 이미지 URL을 TSV로 저장 → PowerShell로 다운로드
   - `Invoke-WebRequest`는 CORS 제약 없음
   - URL 목록은 별도 파일(urls.tsv)로 분리 → PS 스크립트 본문에 `&` 파싱 문제 없음
4. **PowerShell 5.1 호환** — 스크립트 완전 ASCII + UTF-8 BOM

### 네이버 DOM 난독화 대응
네이버는 클래스명이 무작위 해시(`_abcd_1ef2`)로 생성됨. 구조 기반 셀렉터 사용:
- `li[class*="review"]` 같은 속성 선택자 + 내부 텍스트 조건 필터링
- 사이트 업데이트 시 셀렉터 갱신 필요할 수 있음

### 추출 데이터
- 리뷰어 마스킹 이름 (김*관)
- 작성일 / 방문일 / 방문시간
- 상태 (완료/취소/노쇼)
- 업체명 · 서비스명
- 본문 (자동 펼침 적용)
- 태그 칩 (이모지 포함)
- 사진 URL (원본 해상도 시도)
- 영상 URL

---

## 🚀 향후 확장 가능 영역

같은 패턴(Console 스크립트 + PowerShell 보조)으로 구현 가능:
- **인스타그램 DM/댓글** 백업
- **카카오톡 비즈니스 채널** 상담 내역 추출
- **유튜브** 자사 채널 댓글 수집
- **경쟁사 리뷰** 모니터링
- **블로그 브랜드 언급** 스크래핑
- **네이버 쇼핑 가격** 주기 수집

필요 시 요청하면 맞춤 스크립트 작성 가능.
