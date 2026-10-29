import type { ResumeItem } from '../types'

// KST 시각 문자열 (YYYY-MM-DD HH:MM)
export const kst = (ms: number) => new Date(ms + 9 * 3600 * 1000).toISOString().slice(0, 16).replace('T', ' ')

const HEAD = /^##\s*\[(진행중|완료)\]\s*(.+?)\s*$/

export function parse(text: string): ResumeItem[] {
  const out: ResumeItem[] = []
  let cur: ResumeItem | undefined
  for (const line of text.split(/\r?\n/)) {
    const head = line.match(HEAD)
    if (head) {
      cur = { key: '', status: head[1] === '완료' ? '완료' : '진행중', title: head[2] ?? '', when: '', next: '', why: '', doneAt: '' }
      out.push(cur)
      continue
    }
    if (/^##\s/.test(line)) {
      cur = undefined // 형식이 안 맞는 머리줄 — 아래 필드를 엉뚱한 항목에 붙이지 않는다
      continue
    }
    if (!cur) continue
    const field = line.match(/^-\s*(id|멈춘 시각|다음|완료 시각|왜)\s*:\s*(.*?)\s*$/)
    if (!field) continue
    const value = field[2] ?? ''
    if (field[1] === 'id') cur.key = value
    else if (field[1] === '왜') cur.why = value
    else if (field[1] === '멈춘 시각') cur.when = value
    else if (field[1] === '완료 시각') cur.doneAt = value
    else cur.next = value
  }
  return out.map(one => ({ ...one, key: one.key || one.title }))
}

/** "## " 로 시작하지만 [진행중]/[완료] 형식이 아닌 머리줄 수 — 패널에 경고로 띄운다 */
export function countMalformed(text: string): number {
  return text.split(/\r?\n/).filter(line => /^##\s/.test(line) && !HEAD.test(line)).length
}

/** key(id 또는 제목)로 지정 상태의 항목 머리줄 위치를 찾는다 */
function findHead(lines: string[], key: string, status: '진행중' | '완료'): number {
  const want = new RegExp(`^##\\s*\\[${status}\\]\\s*(.+?)\\s*$`)
  let block = -1
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] ?? ''
    if (/^##\s/.test(line)) {
      const head = line.match(want)
      block = head ? i : -1
      if (head && head[1] === key) return i // id 없는 항목은 제목이 key
      continue
    }
    const id = line.match(/^-\s*id\s*:\s*(.*?)\s*$/)
    if (block >= 0 && id && id[1] === key) return block
  }
  return -1
}

// 해당 항목의 머리줄을 [완료]로 바꾸고 완료 시각을 남긴다 (기록은 지우지 않음 — 이력 보존)
export function markDone(text: string, key: string, at: string): string {
  const lines = text.split(/\r?\n/)
  const headAt = findHead(lines, key, '진행중')
  if (headAt < 0) return text
  lines[headAt] = (lines[headAt] ?? '').replace('[진행중]', '[완료]')
  lines.splice(headAt + 1, 0, `- 완료 시각: ${at}`)
  return lines.join('\n')
}

// 잘못 누른 [완료] 되돌리기 — 머리줄을 [진행중]으로, 그 항목의 "완료 시각" 줄은 지운다
export function markUndone(text: string, key: string): string {
  const lines = text.split(/\r?\n/)
  const headAt = findHead(lines, key, '완료')
  if (headAt < 0) return text
  lines[headAt] = (lines[headAt] ?? '').replace('[완료]', '[진행중]')
  for (let i = headAt + 1; i < lines.length && !/^##\s/.test(lines[i] ?? ''); i++) {
    if (/^-\s*완료 시각\s*:/.test(lines[i] ?? '')) {
      lines.splice(i, 1)
      break
    }
  }
  return lines.join('\n')
}
