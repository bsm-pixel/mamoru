// 답변이 영어 위주인지 판정 — 코드·경로·주소·표 구분선은 빼고 순수 문장만 센다
export function englishHeavy(text: string): boolean {
  const prose = text
    .replace(/```[\s\S]*?```/g, ' ') // 코드 블록
    .replace(/`[^`]*`/g, ' ') // 인라인 코드·경로
    .replace(/https?:\/\/\S+/g, ' ') // 주소
    .replace(/\[[^\]]*\]\([^)]*\)/g, ' ') // 마크다운 링크
    .replace(/[A-Za-z]:[\/][^\s)]*/g, ' ') // 윈도우 경로
  const hangul = (prose.match(/[가-힣]/g) ?? []).length
  const latin = (prose.match(/[A-Za-z]/g) ?? []).length
  if (latin < 60) return false // 짧은 영어(메뉴명·약어)는 허용
  return hangul / (hangul + latin) < 0.5
}

// 재작성을 걸지 판정 — 사장님 화면에 보이는 본 대화의 답변만 본다.
// 2026-11-26 보조 에이전트(웹 문서 읽기 등)가 끝날 때도 같은 이벤트가 와서, 그 영어 보고 때문에
// 한국어로 이미 쓴 본 답변을 다시 쓰게 하던 오탐을 막는다 (agentId 있음 = 보조 에이전트).
export function shouldRewrite(e: { reason: string; agentId?: string }, text: string): boolean {
  if (e.agentId !== undefined) return false
  return e.reason === 'answer' && englishHeavy(text)
}
