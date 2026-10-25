import type { ResumeItem } from '../types'

// KST 시각 문자열 (YYYY-MM-DD HH:MM)
export const kst = (ms: number) => new Date(ms + 9 * 3600 * 1000).toISOString().slice(0, 16).replace('T', ' ')

export function parse(text: string): ResumeItem[] {
  const out: ResumeItem[] = []
  let cur: ResumeItem | undefined
  for (const line of text.split(/\r?\n/)) {
    const head = line.match(/^##\s*\[(진행중|완료)\]\s*(.+?)\s*$/)
    if (head) {
      cur = { key: '', status: head[1] === '완료' ? '완료' : '진행중', title: head[2] ?? '', when: '', next: '' }
      out.push(cur)
      continue
    }
    if (!cur) continue
    const field = line.match(/^-\s*(id|멈춘 시각|다음)\s*:\s*(.*?)\s*$/)
    if (!field) continue
    const value = field[2] ?? ''
    if (field[1] === 'id') cur.key = value
    else if (field[1] === '멈춘 시각') cur.when = value
    else cur.next = value
  }
  return out.map(one => ({ ...one, key: one.key || one.title }))
}

// 해당 항목의 머리줄을 [완료]로 바꾸고 완료 시각을 남긴다 (기록은 지우지 않음 — 이력 보존)
export function markDone(text: string, key: string, at: string): string {
  const lines = text.split(/\r?\n/)
  let block = -1 // 지금 보고 있는 [진행중] 머리줄 위치
  let headAt = -1 // 찾은 항목의 머리줄 위치
  lines.forEach((line, i) => {
    if (headAt >= 0) return
    if (/^##\s/.test(line)) {
      const head = line.match(/^##\s*\[진행중\]\s*(.+?)\s*$/)
      block = head ? i : -1
      if (head && head[1] === key) headAt = i // id 없는 항목은 제목이 key
      return
    }
    const id = line.match(/^-\s*id\s*:\s*(.*?)\s*$/)
    if (block >= 0 && id && id[1] === key) headAt = block
  })
  if (headAt < 0) return text
  lines[headAt] = (lines[headAt] ?? '').replace('[진행중]', '[완료]')
  lines.splice(headAt + 1, 0, `- 완료 시각: ${at}`)
  return lines.join('\n')
}
