import { atom, read, update } from 'claude-code'
import type { EngineInterface, Register } from 'claude-code'

import type { ResumeItem } from '../types'
import { countMalformed, kst, markDone, markUndone, parse } from './parse'

// 멈춘 작업 목록 — 구글 드라이브 폴더라 PC 2대가 자동 동기화된다 (git 불필요)
const DIR = 'C:/마모루드라이브/ClaudeWork'
const FILE = `${DIR}/작업이어하기.md`
const PANE = 'work-resume'

const items = atom({ plugin: 'work-resume', key: 'items' } as const, [] as ResumeItem[])
const error = atom({ plugin: 'work-resume', key: 'error' } as const, '')
const warn = atom({ plugin: 'work-resume', key: 'warn' } as const, '')
const checkedAt = atom({ plugin: 'work-resume', key: 'checkedAt' } as const, '')
const sent = atom({ plugin: 'work-resume', key: 'sent' } as const, '')
const confirm = atom({ plugin: 'work-resume', key: 'confirm' } as const, '')

const TEMPLATE = `# 작업 이어하기

> 하다가 멈춘 작업 목록. 클로드 코드 패널(work-resume mod)이 이 파일을 읽어 보여준다.
> 형식: "## [진행중] 제목" 또는 "## [완료] 제목" + 아래 "- 항목: 값" 줄. 새 작업은 이 설명 바로 아래(맨 위)에 추가.

`

const PAUSE_WORDS = /(멈추자|멈출게|멈춰\s*줘|멈춤|중단하자|그만하자|여기까지\s*하자|나중에\s*이어)/
const RESUME_WORDS = /(이어서|이어하|이어 하|하던\s*(거|것|작업))/

const pauseGuide = () => `[작업 이어하기 규칙 — work-resume mod]
사장님 말이 "지금 작업을 여기서 멈추자"는 뜻일 때만 따른다(아니면 무시).
${FILE} 파일(없으면 만든다)의 설명 줄 바로 아래, 기존 항목들보다 위에 지금 작업을 아래 형식으로 추가한다. 기존 항목은 건드리지 않는다.

## [진행중] <작업 제목 — 15자 안팎, 한국어>
- id: <YYYYMMDD-HHMM, 지금 KST>
- 멈춘 시각: <YYYY-MM-DD HH:MM KST>
- 왜: <어느 화면·업무의 무슨 일인지 + 끝나면 사장님께 뭐가 좋아지는지, 전문용어 없이 한 줄>
- 다음: <바로 다음에 할 일 한 줄>
- 끝난 것: <한두 줄>
- 남은 것: <한두 줄>
- 관련 파일·커밋: <경로, 커밋 sha>
- 주의: <승인 대기·검증 못 한 것 등, 없으면 생략>

저장한 뒤 "작업 목록에 저장했습니다" 한 줄 + 미커밋 변경이 있으면 커밋할지 묻는다.`

const resumeGuide = () => `[작업 이어하기 목록 — work-resume mod]
하다가 멈춘 작업 기록 = ${FILE}. 사장님이 예전에 멈춘 작업을 이어가자는 뜻이면, 이 파일에서 해당 [진행중] 항목을 찾아 끝난 것·남은 것·주의를 읽고 이어서 진행한다. 이어간 작업이 끝나면 그 항목 머리줄을 [완료]로 바꾸고 "- 완료 시각:" 줄을 덧붙인다.`

// 파일을 읽어 목록을 새로 그린다. 없으면 빈 틀을 만든다. 형식 오류·드라이브 충돌 사본도 함께 점검.
async function load($: EngineInterface) {
  try {
    const has = await $.fs.exists(FILE)
    if (!has) await $.fs.write(FILE, TEMPLATE)
    const text = has ? ((await $.fs.read(FILE)) as string) : TEMPLATE
    const notes: string[] = []
    const bad = countMalformed(text)
    if (bad > 0) notes.push(`형식이 안 맞는 항목 ${bad}개는 목록에 안 보입니다("## [진행중] 제목" 형식)`)
    // 구글 드라이브가 PC 2대 동시 수정 때 만드는 충돌 사본(예: 작업이어하기 (1).md)
    const copies = (await $.fs.list(DIR).catch(() => []))
      .map(f => f.name)
      .filter(n => /^작업이어하기.+\.md$/.test(n) && n !== '작업이어하기.md')
    if (copies.length > 0) notes.push(`드라이브 충돌 사본이 있습니다: ${copies.join(', ')} — 클로드에게 합쳐 달라고 하세요`)
    await update($, items, () => parse(text))
    await update($, warn, () => notes.join(' / '))
    await update($, error, () => '')
    await update($, checkedAt, () => kst(Date.now()).slice(11))
  } catch (err) {
    await update($, error, () => `작업 목록 파일을 읽지 못했습니다: ${String(err).slice(0, 160)}`)
  }
}

export const register: Register = on => {
  on('session.start', async ($, e, next) => {
    await $.command.register({ name: 'pause', description: '지금 작업을 "이어할 작업" 목록에 저장하고 멈춤' })
    await $.command.register({ name: 'resume-list', description: '이어할 작업 패널 열기' })
    await load($)
    void $.ui.open({ id: PANE, title: '이어할 작업' })
    // 다른 PC에서 드라이브로 넘어온 변경도 보이게 1분마다 다시 읽음
    $.clock.every(60_000, () => void load($))
    return next(e)
  })

  on('command.run', { command: 'pause' }, async $ => {
    void $.prompt.submit({ text: `여기서 멈추자. 지금 작업을 이어할 작업 목록에 저장해줘.\n\n${pauseGuide()}` })
    return { text: '지금 작업을 목록에 저장하도록 요청했습니다.' }
  })

  on('command.run', { command: 'resume-list' }, async $ => {
    await load($)
    await $.ui.open({ id: PANE, title: '이어할 작업' })
    return { text: '이어할 작업 패널을 열었습니다.' }
  })

  // "여기서 멈추자" / "그거 이어서 하자" 같은 말에 규칙을 몰래 붙여 클로드에게 전달
  on('prompt.submit', async ($, e, next) => {
    if (e.origin?.kind === 'plugin') return next(e)
    const add: string[] = []
    if (PAUSE_WORDS.test(e.text)) add.push(pauseGuide())
    if (RESUME_WORDS.test(e.text)) add.push(resumeGuide())
    if (add.length === 0) return next(e)
    return next({ ...e, context: [...(e.context ?? []), ...add] })
  })

  // 클로드가 파일을 고쳤을 수 있으니 턴이 끝날 때마다 다시 읽음
  on('turn.complete', async ($, e, next) => {
    const done = await next(e)
    await load($)
    return done
  })

  on('ui.render', { component: 'Pane', requestId: PANE }, async ($, e) => {
    const { Box, Text, Button } = $.ui.resolve(e)
    const list = await read($, items)
    const problem = await read($, error)
    const note = await read($, warn)
    const at = await read($, checkedAt)
    const sentKey = await read($, sent)
    const confirmKey = await read($, confirm)
    // 최근에 멈춘 작업이 위로
    const open = list.filter(one => one.status === '진행중').sort((a, b) => b.when.localeCompare(a.when))
    const done = list.filter(one => one.status === '완료').sort((a, b) => b.doneAt.localeCompare(a.doneAt))

    // 이어하기: 클로드에게 "이 항목 기록 읽고 이어서" 요청을 보낸다. 5초 안 재클릭은 무시(중복 전송 방지)
    const resume = (one: ResumeItem) => async () => {
      if ((await read($, sent)) === one.key) return
      await update($, sent, () => one.key)
      $.clock.after(5000, () => void update($, sent, cur => (cur === one.key ? '' : cur)))
      await $.prompt.submit({
        text: `작업 이어하기: "${one.title}" (id ${one.key}). ${FILE} 에서 이 항목 기록(끝난 것·남은 것·주의)을 읽고 이어서 진행해줘. 진행하면서 큰 단계가 끝날 때마다 이 항목의 "다음·끝난 것·남은 것"을 갱신해줘.`,
      })
    }
    // 완료: 실수 방지로 두 번 눌러야 완료(5초 안). 되돌리기는 아래 "최근 완료"에서
    const finish = (one: ResumeItem) => async () => {
      if ((await read($, confirm)) !== one.key) {
        await update($, confirm, () => one.key)
        $.clock.after(5000, () => void update($, confirm, cur => (cur === one.key ? '' : cur)))
        return
      }
      await update($, confirm, () => '')
      try {
        const text = (await $.fs.read(FILE)) as string
        await $.fs.write(FILE, markDone(text, one.key, kst(Date.now())))
        $.ui.toast(`완료 처리: ${one.title}`)
      } catch (err) {
        $.ui.toast(`완료 처리 실패: ${String(err).slice(0, 120)}`)
      }
      await load($)
    }
    const undo = (one: ResumeItem) => async () => {
      try {
        const text = (await $.fs.read(FILE)) as string
        await $.fs.write(FILE, markUndone(text, one.key))
        $.ui.toast(`다시 진행중으로: ${one.title}`)
      } catch (err) {
        $.ui.toast(`되돌리기 실패: ${String(err).slice(0, 120)}`)
      }
      await load($)
    }

    return (
      <Box flexDirection="column">
        <Text dimColor>
          진행중 {open.length}건{at !== '' ? ` · 마지막 확인 ${at}` : ''}
        </Text>
        {problem !== '' && <Text>⚠ {problem}</Text>}
        {note !== '' && <Text>⚠ {note}</Text>}
        {problem === '' && open.length === 0 && (
          <Text dimColor>멈춘 작업이 없습니다. 작업 중에 "여기서 멈추자" 또는 /pause 라고 하면 여기에 쌓입니다.</Text>
        )}
        {open.map((one, i) => (
          <Box key={one.key} flexDirection="column" marginTop={1}>
            <Text bold>
              {i + 1}. {one.title}
            </Text>
            {one.why !== '' && <Text>왜: {one.why}</Text>}
            {one.next !== '' && <Text>다음: {one.next}</Text>}
            {one.when !== '' && <Text dimColor>멈춘 시각 {one.when}</Text>}
            <Box flexDirection="row">
              <Button
                key={`resume-${one.key}`}
                label={sentKey === one.key ? '보냈습니다 ✓' : '이어하기'}
                variant="primary"
                onPress={resume(one)}
              />
              <Button
                key={`done-${one.key}`}
                label={confirmKey === one.key ? '한 번 더 누르면 완료' : '완료'}
                dimColor={confirmKey !== one.key}
                onPress={finish(one)}
              />
            </Box>
          </Box>
        ))}
        {done.length > 0 && (
          <Box flexDirection="column" marginTop={1}>
            <Text dimColor>최근 완료 (잘못 눌렀으면 되돌리기)</Text>
            {done.slice(0, 3).map(one => (
              <Box key={`d-${one.key}`} flexDirection="row">
                <Text dimColor>
                  {one.title}
                  {one.doneAt !== '' ? ` · ${one.doneAt}` : ''}
                </Text>
                <Button key={`undo-${one.key}`} label="되돌리기" dimColor onPress={undo(one)} />
              </Box>
            ))}
          </Box>
        )}
      </Box>
    )
  })
}
