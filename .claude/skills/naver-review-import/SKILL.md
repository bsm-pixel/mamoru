---
name: naver-review-import
description: >-
  네이버 스마트플레이스 리뷰를 TMS 리뷰로 일괄 등록(글·작성일·방문일·사진). 내려받은 폴더를 읽어
  상담/복원수리/제품구매를 본문으로 판정하고 중복 없이 등록한다.
  Use when the user asks to import/register Naver reviews into TMS, or mentions the downloaded
  Naver review folder/zip.
when_to_use: >-
  네이버 리뷰, 네이버리뷰 가져와, 네이버 리뷰 등록, 스마트플레이스 리뷰, 리뷰 끌어오기, 리뷰 일괄 등록,
  naver review, naver_reviews zip
---

# 네이버 리뷰 → TMS 일괄 등록

사장님이 "네이버 리뷰 가져와"라고 하면 아래 순서로 끝까지 진행한다. 사장님이 직접 하는 일은 **1번(내려받기)뿐**이다.

## 0. 전체 그림

| 단계 | 누가 | 도구 |
|---|---|---|
| 1. 네이버에서 내려받기 | 사장님 (로그인된 브라우저 필요 — 네이버는 플레이스 리뷰 공식 API가 없다) | `projects/marketing/naver_review_extract.js` |
| 2. 압축 풀기 + 사진 받기 | Claude | `download_images.ps1` (폴더 안에 들어 있음) |
| 3. 계획 만들기 | Claude | `node scripts/naver-review-import.cjs plan "<폴더>"` |
| 4. 종류 판정 | Claude (본문을 직접 읽는다) | `<폴더>/_tms_overrides.json` |
| 5. 등록 | Claude | `node scripts/naver-review-import.cjs apply "<폴더>"` |
| 6. 확인·보고 | Claude | DB 건수 + 라이브 리뷰 페이지 |

## 1. 폴더 찾기

- 사장님이 위치를 말하면 그곳. 아니면 `Downloads` / 바탕화면에서 가장 최근 `naver_reviews_*.zip` 또는 `네이버리뷰*` 폴더.
- zip이면 같은 위치에 푼다. `reviews.json` 과 `NNN_날짜_이름/` 폴더들이 있어야 한다.
- 사진: `reviews.json` 의 photos 수보다 폴더의 `photo_*.jpg` 가 적으면 폴더에서
  `powershell -ExecutionPolicy Bypass -File .\download_images.ps1` 실행 (브라우저는 CORS 때문에 사진을 못 받는다).
- 사장님이 아직 안 내려받았으면 `projects/marketing/README.md` 의 3단계(스마트플레이스 리뷰 화면 → F12 콘솔 → 스크립트 붙여넣기)를 안내한다.

## 2. plan 실행

```
node scripts/naver-review-import.cjs plan "<폴더>"
```

- DB에는 아무것도 쓰지 않는다. `<폴더>/_tms_plan.json` 이 생긴다.
- 이미 등록된 건(같은 작성일 + 같은 본문)과 본문 없는 건은 자동으로 `skip`.
- 각 건에 규칙 기반 추정(`type`·`subtype`·`guess_sure`)이 들어 있다. **추정은 참고만** — 4번에서 본문을 읽고 확정한다.

## 3. 종류 판정 기준 (가장 중요 — 사장님 요구: 상담·복원수리를 정확히 나눌 것)

네이버 파일의 서비스명은 비어 있다. **등록 예정 건의 본문을 전부 읽고** 판정한다. 기준은 사장님이 직접 넣은 기존 네이버 리뷰의 분류 방식:

| 판정 | type / subtype | 본문 단서 |
|---|---|---|
| 방문 상담 (기본값) | `consult` / `store_visit` | 방문해서 설명·추천받고 구매. **"상담받고 구매"는 제품구매가 아니라 상담이다** |
| 출장 상담 | `consult` / `field_request` | 사장님이 고객 매장으로 감 — "출장", "와주셔서", "매장에 직접 방문해주셔서", "방문상담 해주셔서" |
| 복원수리 — 방문 | `repair` / `direct_visit` | 방문 목적이 수리·복원·연마·AS·점검 ("수리받으러", "AS 받으러 재방문") |
| 복원수리 — 택배 | `repair` / `parcel_pickup` | "택배", "멀리서 맡긴", "보내고 받았다" |
| 복원수리 — 구분 불명 | `repair` / `restoration` | 수리인 건 분명한데 방문인지 택배인지 알 수 없음 |
| 제품구매 | `purchase` / null | 방문·상담 없이 제품만 배송받은 것이 분명할 때만 (드묾) |

- 수리와 구매가 섞인 글: **방문 목적**이 무엇이었는지로 정한다. "수리하러 갔다가 가위도 샀다" = 복원수리, "가위 사러 갔는데 쓰던 가위도 봐주셨다" = 상담.
- **사장님 답글이 본문으로 수집된 건은 제외**한다. 고객이 글 없이 사진만 올리면 추출 스크립트가 사장님 답글을 본문으로 잡는다.
  "안녕하세요 마모루입니다", "~하겠습니다", "방문주셔서 감사했습니다", "믿고 맡겨주셔서" 같은 사장님 말투면 `skip`.
- 사장님이 예전에 수동으로 넣은 건은 본문을 다듬어 넣어서 자동 중복 감지가 안 될 수 있다. 가장 오래된 날짜대에서 같은 이름·같은 날짜의 기존 건이 있는지 한 번 확인한다.

판정은 `<폴더>/_tms_overrides.json` 에 적는다 (index 가 키):

```json
{
  "3":   { "type": "repair", "subtype": "direct_visit" },
  "55":  { "type": "consult", "subtype": "field_request" },
  "149": { "skip": true, "skip_reason": "고객 글 없음(사장님 답글만 수집됨)" }
}
```

적은 뒤 plan 을 다시 실행하면 반영된다 (확신 낮음 0 이 되어야 한다).

## 4. apply 실행

```
node scripts/naver-review-import.cjs apply "<폴더>"
```

- 사진을 저장소 `review-photos/reviews/naver/` 에 올리고 리뷰를 등록한다.
- 저장 규칙: 별점 5(네이버엔 별점이 없다), 승인 상태(바로 공개), `source='naver'`, `created_at`=네이버 작성일, `meta.received_at`=방문일,
  `source_id`=`naver-<작성일>-<index>`.
- **등록 즉시 고객 리뷰 페이지에 공개된다.** 그래서 판정(3번)을 끝낸 뒤에만 실행한다.
- 다시 실행해도 중복 등록되지 않는다.

## 5. 확인하고 보고

- apply 를 한 번 더 실행해 "등록 0" 인지 확인(중복 방지).
- DB: `reviews` 에서 `source='naver'` 건수, 종류별 분포.
- 사진 URL 하나를 실제로 열어 200 인지 확인.
- 사장님께 보고: 등록 건수 / 종류별 건수 / 제외한 건과 이유 / 판정이 애매했던 건 목록(index·한 줄 요약) — 사장님이 TMS 리뷰 화면에서 고칠 수 있게.

## 되돌리기

잘못 넣었으면 `source='naver'` 이고 `meta.imported_at` 이 그날인 건을 `status='hidden'` 으로 바꾼다(삭제하지 않는다).

## 한계

- 영상은 등록하지 않는다(리뷰 화면이 사진만 표시).
- 이름은 네이버가 가린 형태 그대로(`정*진`). 이름이 비어 있으면 `네이버 고객`.
- 1번(내려받기)을 Claude 크롬 확장으로 대신할 수는 있지만, 네이버 화면이 바뀌면 깨지기 쉬워 기본 절차에 넣지 않았다.
