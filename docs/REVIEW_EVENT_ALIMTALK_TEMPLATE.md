# 리뷰 이벤트 알림톡 템플릿 (솔라피 콘솔 등록용)

> 리뷰 이벤트 **당첨자 안내**용. TMS 「리뷰 이벤트 관리」에서 [당첨 발표] 시 발송(코드 배선은 후속 — 현재 미연결).
> 톤 = MAMORU "조용히 압도하는 전문가" — 절제·담백·존댓말, **마침표 생략**, 항목 앞 검은 점 `•`, 이모지는 한 곳만.
> 검수 1~3영업일. 변수명 = 코드가 보낼 JSON 필드명과 100% 일치시킬 것.

---

## ① review_event_won — 리뷰 이벤트 당첨 안내 (발표 시 자동) ★필수
**변수**: `#{name}` `#{event_month}`(예: 8월) `#{rank}`(예: 1등) `#{prize}`(상품명)
- 강조표기 제목: `리뷰 이벤트에 당첨되셨습니다` / 보조문구: `MAMORU`
- 내용:
```
#{name}님, 안녕하세요
#{event_month} 리뷰 이벤트 #{rank}에 당첨되셨습니다
남겨주신 진심 어린 후기 감사합니다

🎁 당첨 상품
#{prize}

• 상품을 보내드릴 수 있도록
  받으실 주소·연락처를 남겨 주세요
• 안내 후 7일간 회신이 없으면
  상품은 다음 달로 이월됩니다
```
- 부가정보: `고객센터 · 평일 09:00 ~ 17:00`
- 버튼: **[주소·연락처 전달]** — BC(봇전환/채팅상담), extra `#{name}_won` (50바이트 가드, 한글 1자=3바이트)

---

## (선택) review_event_thanks — 미당첨 참여 감사
> 보통 미발송. 참여 독려 캠페인 때만 등록. 변수 `#{name}` `#{event_month}`.
- 강조표기 제목: `이번 달 후기 감사합니다` / 보조문구: `MAMORU`
- 내용(요지): `#{name}님, #{event_month} 리뷰 이벤트에 함께해 주셔서 감사합니다 / 이번엔 아쉽게 비껴갔지만 다음 달에도 진짜 후기를 기다립니다`

---

## Make 시나리오
- 웹훅 필터 `{{1.template}}` Equal to `review_event_won` → 매핑 `name, event_month, rank, prize`.
- 발송 주체: TMS [당첨 발표] 시 announced 당첨자에게(코드 후속). 지금은 **솔라피 등록·검수만** 먼저 진행 가능.

## 가동 전 체크리스트
- [ ] 솔라피: `review_event_won` 등록 + 카카오 검수
- [ ] (후속) TMS 발송 배선 — 발표 시 announced 당첨자 phone 으로 웹훅. 마스킹 아닌 실 연락처 사용(발송용)
- [ ] 토글 키(예: `notifications.review_event_won`) 설계

관련: [reference_solapi_templates] · `projects/Total_Management_System/docs/TMS_FLOW_REVIEW_EVENT.md` · [project_review_event]
