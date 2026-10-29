import { describe, expect, test } from 'claude-code/testing'

import { countMalformed, markDone, markUndone, parse } from '../hooks/parse'

const SAMPLE = `# 작업 이어하기

> 설명

## [진행중] B-09 12~22번 개편
- id: 20261023-2140
- 멈춘 시각: 2026-10-23 21:40
- 왜: 고객 페이지 22개를 새 디자인으로 — 브랜드 통일
- 다음: 시험 접수 확인
- 끝난 것: 시안 11개

## [완료] 반품 부분처리
- id: 20261022-1810
- 완료 시각: 2026-10-22 20:00

## [진행중] 제목만 있는 작업
- 다음: 확인
`

describe('작업 목록 파일 읽기·완료 처리', () => {
  test('진행중·완료 항목과 필드를 읽는다', async () => {
    const list = parse(SAMPLE)
    expect(list.length).toBe(3)
    expect(list[0]).toEqual({ key: '20261023-2140', status: '진행중', title: 'B-09 12~22번 개편', when: '2026-10-23 21:40', next: '시험 접수 확인', why: '고객 페이지 22개를 새 디자인으로 — 브랜드 통일', doneAt: '' })
    expect(list[1].doneAt).toBe('2026-10-22 20:00')
    expect(list[1].status).toBe('완료')
    expect(list[2].key).toBe('제목만 있는 작업')
  })

  test('id로 찾은 항목만 [완료]로 바꾸고 기록은 남긴다', async () => {
    const out = markDone(SAMPLE, '20261023-2140', '2026-10-25 10:00')
    const list = parse(out)
    expect(list[0].status).toBe('완료')
    expect(list[2].status).toBe('진행중')
    expect(out).toContain('- 완료 시각: 2026-10-25 10:00')
    expect(out).toContain('- 끝난 것: 시안 11개')
  })

  test('id 없는 항목은 제목으로 찾고, 없는 key는 그대로 둔다', async () => {
    expect(parse(markDone(SAMPLE, '제목만 있는 작업', 'x'))[2].status).toBe('완료')
    expect(markDone(SAMPLE, '없는-id', 'x')).toBe(SAMPLE)
  })
})

describe('되돌리기·형식 점검', () => {
  test('완료를 되돌리면 진행중으로 돌아가고 완료 시각 줄이 지워진다', async () => {
    const done = markDone(SAMPLE, '20261023-2140', '2026-10-25 10:00')
    const back = markUndone(done, '20261023-2140')
    expect(parse(back)[0].status).toBe('진행중')
    expect(back).not.toContain('2026-10-25 10:00')
    expect(parse(back)[1].status).toBe('완료') // 다른 완료 항목은 그대로
  })
  test('형식이 안 맞는 머리줄을 센다', async () => {
    expect(countMalformed(SAMPLE)).toBe(0)
    expect(countMalformed(SAMPLE + '\n## 대괄호 없는 제목\n- 다음: x\n')).toBe(1)
    expect(parse(SAMPLE + '\n## 대괄호 없는 제목\n- 다음: x\n').length).toBe(3) // 엉뚱한 항목에 붙지 않음
  })
})

describe('패널', () => {
  test('파일이 없으면 빈 안내를 그린다', async ($, on) => {
    on('fs.exists', () => ({ value: false }) as never)
    on('fs.write', () => ({ value: undefined }) as never)
    on('command.register', () => ({ value: undefined }) as never)
    on('ui.open', () => ({ value: { isPlaced: true } }) as never)
    on('clock.every', () => ({ value: undefined }) as never)
    on('fs.list', () => ({ value: [] }) as never)
    on('session.start', (_$, e) => e as never)
    await $.session.start({ cwd: 'C:/' } as never)
    for (const surface of ['vscode', 'terminal'] as const) {
      const ui = await $.ui.mount({ plugin: 'work-resume', surface, component: 'Pane', requestId: 'work-resume', props: {} } as never)
      expect(await ui.find({ type: 'Text', text: /멈춘 작업이 없습니다/ })).toBeDefined()
      await ui.unmount()
    }
  })

  test('진행중 항목만 버튼과 함께 그린다', async ($, on) => {
    on('fs.exists', () => ({ value: true }) as never)
    on('fs.read', () => ({ value: SAMPLE }) as never)
    on('command.register', () => ({ value: undefined }) as never)
    on('ui.open', () => ({ value: { isPlaced: true } }) as never)
    on('clock.every', () => ({ value: undefined }) as never)
    on('fs.list', () => ({ value: [] }) as never)
    on('session.start', (_$, e) => e as never)
    await $.session.start({ cwd: 'C:/' } as never)
    for (const surface of ['vscode', 'terminal'] as const) {
      const ui = await $.ui.mount({ plugin: 'work-resume', surface, component: 'Pane', requestId: 'work-resume', props: {} } as never)
      expect(await ui.find({ key: 'resume-20261023-2140' })).toBeDefined()
      expect(await ui.find({ key: 'done-20261023-2140' })).toBeDefined()
      expect(await ui.find({ key: 'resume-20261022-1810' })).toBeUndefined()
      expect(await ui.find({ type: 'Text', text: /최근 완료/ })).toBeDefined()
      expect(await ui.find({ key: 'undo-20261022-1810' })).toBeDefined() // 완료 항목 되돌리기 버튼
      expect(await ui.find({ type: 'Text', text: /진행중 2건/ })).toBeDefined()
      await ui.unmount()
    }
  })
})
