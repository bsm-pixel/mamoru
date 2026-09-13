---
name: notion-task-schedule
description: >-
  노션 할 일 등록 · 구글 캘린더 일정 등록 · 프로젝트 연결을 claude.ai(Cowork)와 똑같은 방식으로 처리.
  Use when the user asks to add/register a to-do, task, deadline, schedule, appointment or calendar event,
  or when the conversation implies "this needs to be done later" (→ auto-create a to-do).
when_to_use: >-
  할 일, 할일, 할일 등록, 할 일로 넣어줘, 내일 할 일, 투두, 일정, 일정 등록, 일정 넣어줘, 캘린더, 달력에 넣어줘,
  캘린더 추가, 약속, 미팅, 예약, 마감, 언제까지, 해야 해, 해야 함, 나중에 해야, 프로젝트 연결, 업무/개인
---

# 노션 할 일 · 일정 등록

## 🔒 규칙 원본(SSOT) = 노션 페이지 1개

**📋 MAMORU 할일·일정 등록 규칙**
https://app.notion.com/p/3daf1558c01a81c49b38f73cb1047b21

- **등록하기 전에 매번 이 페이지를 `notion-fetch`로 읽고 그대로 따른다.** 기억·이 파일 요약으로 대신하지 않는다 (규칙이 수시로 바뀜)
- claude.ai(Cowork)도 같은 페이지를 기준으로 쓴다 → 모든 창이 같은 방식
- **규칙을 바꾸라는 요청은 이 노션 페이지를 고친다.** 이 SKILL.md에 규칙 내용을 옮겨 적지 않는다 (두 곳이 되면 어긋남). 변경 이력 칸에 날짜와 내용을 남긴다
- 페이지를 못 읽으면(커넥터 오류 등) **등록하지 말고** 사용자에게 알린다

## 도구
- 노션: `mcp__claude_ai_Notion__*` (fetch · query-data-sources · create-pages · update-page)
- 구글 캘린더: `mcp__claude_ai_Google_Calendar__*` (create_event 등) — deferred면 `ToolSearch`로 로드

## 절대 틀리면 안 되는 것 (원본 읽은 뒤 재확인용 — 원본이 우선)
- 일정 = **구글 캘린더**(업무 `bsm@mamoru.kr` / 🏠개인 캘린더). 노션 「일정 캘린더」 DB 사용 금지
- 마모루 **출장방문 · 방문예약** 캘린더는 조회만 (시스템 자동 기록)
- 할 일 등록 시 `상태`=`할 일`, `구분`(업무/개인) 반드시 채움, 모르는 값은 비움
- 제목에 `[솔라피]` 같은 대괄호 말머리 금지 · 대표 직접 작업은 `👤` 접두
- 묻지 않고 등록 → 표 1개 + `⚠️ 확인 필요`로 보고
