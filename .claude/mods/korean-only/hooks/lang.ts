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
  if (latin < 120) return false // 짧은 영어(메뉴명·약어)는 허용
  return hangul / (hangul + latin) < 0.35
}
