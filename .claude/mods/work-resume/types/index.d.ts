export type ResumeItem = {
  key: string
  status: '진행중' | '완료'
  title: string
  when: string
  next: string
  doneAt: string
}

declare module 'claude-code' {
  interface PluginState {
    'work-resume': {
      items: ResumeItem[]
      error: string
      /** 형식 경고·드라이브 충돌 사본 안내 */
      warn: string
      /** 마지막으로 파일을 읽은 시각 (KST HH:MM) */
      checkedAt: string
      /** 이어하기를 막 누른 항목 key — 5초간 중복 전송 방지 */
      sent: string
      /** 완료 확인 대기 중인 항목 key — 한 번 더 눌러야 완료 */
      confirm: string
    }
  }
}
