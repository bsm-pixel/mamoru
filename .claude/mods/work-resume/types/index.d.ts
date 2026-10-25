export type ResumeItem = {
  key: string
  status: '진행중' | '완료'
  title: string
  when: string
  next: string
}

declare module 'claude-code' {
  interface PluginState {
    'work-resume': { items: ResumeItem[]; error: string }
  }
}
