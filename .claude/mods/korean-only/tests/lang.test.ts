import { describe, expect, test } from 'claude-code/testing'

import { englishHeavy, shouldRewrite } from '../hooks/lang'

// 2026-11-26 오탐 사례 — 보조 에이전트(웹 문서 읽기)의 영어 보고
const SUBAGENT_REPORT = 'PARTIAL RESULT. Only the SOLAPI blog post could be read. The other pages the caller named did not yield the requested details. Most of the requested items are not stated in what I could read. Requirement: Node.js 16 or higher. The blog does not give a separate Claude Desktop vs Claude Code snippet.'

describe('재작성 판정 (본 답변만)', () => {
  test('보조 에이전트의 영어 보고는 재작성하지 않는다', async () => {
    expect(shouldRewrite({ reason: 'answer', agentId: 'a41a70d51cede7ba6' }, SUBAGENT_REPORT)).toBe(false)
  })
  test('본 대화의 영어 답변은 여전히 재작성한다', async () => {
    expect(shouldRewrite({ reason: 'answer' }, SUBAGENT_REPORT)).toBe(true)
  })
  test('본 대화의 한국어 답변은 통과', async () => {
    expect(shouldRewrite({ reason: 'answer' }, '설치와 등록까지 끝났습니다. 남은 것은 사장님이 직접 하실 키 입력과 앱 재시작입니다.')).toBe(false)
  })
  test('중단된 답변은 보지 않는다', async () => {
    expect(shouldRewrite({ reason: 'aborted' }, SUBAGENT_REPORT)).toBe(false)
  })
})

describe('영어 답변 판정', () => {
  test('영어 위주 답변은 잡는다', async () => {
    const en = 'I sent one actual tag-shaped test label to the printer. Please check these three things and send a photo. Did the text land only inside the white pads? Once you confirm the position and layout, I will set it up in TMS.'
    expect(englishHeavy(en)).toBe(true)
  })
  test('한국어 답변은 통과', async () => {
    const ko = '꼬리표 시험 라벨을 1장 보냈습니다. 흰 패드 안에만 찍혔는지 확인해 주세요. 위치가 맞으면 TMS에 템플릿으로 넣고 수량별 원클릭 출력을 만들겠습니다.'
    expect(englishHeavy(ko)).toBe(false)
  })
  test('코드·경로·주소가 많아도 설명이 한국어면 통과', async () => {
    const mixed = '수정 위치는 아래입니다.\n```ts\nexport const register: Register = on => { on("turn.complete", async ($, e, next) => next(e)) }\nconst veryLongEnglishIdentifierName = somethingElseHere\n```\n파일 `C:/Users/user/Desktop/mamoru/projects/Total_Management_System/app/src/lib/label/templates.ts` 와 https://code.claude.com/docs/en/plugins/mods/overview.md 참고.'
    expect(englishHeavy(mixed)).toBe(false)
  })
})
