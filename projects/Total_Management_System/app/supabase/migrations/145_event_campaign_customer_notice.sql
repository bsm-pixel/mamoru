-- 145: EVENT 캠페인 고객 안내 문구 — 범용 알림톡(EVENT_신청완료)의 #{event_notice} 변수
-- 이벤트별 진행 안내(예: "보내실 가위는 신청 후 3일 안에 발송해 주세요", 주소 회신 요청 등).
-- 비어 있으면 코드가 기본 문구로 채워 발송한다(lib/event/campaign-notify.ts). 기존 데이터 영향 없음.
ALTER TABLE event_campaigns ADD COLUMN IF NOT EXISTS customer_notice TEXT;
