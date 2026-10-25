import type { Register } from 'claude-code'

import { englishHeavy } from './lang'

const RULE = `[답변 언어 — 절대 규칙 (사장님 지시, korean-only mod)]
사장님(사용자)에게 보이는 모든 글은 처음부터 끝까지 한국어로 쓴다: 최종 답변, 도구 사용 사이의 진행 메시지, 표, 목록, 제목, 체크리스트, 기술 설명 전부.
도구 결과·문서·로그가 영어여도 설명은 한국어로 옮겨 쓴다. 영어는 화면에 실제로 보이는 메뉴명·명령어·코드·파일 경로를 괄호나 코드 표기로 덧붙일 때만 허용.
보내기 전에 영어 문장이 섞였는지 스스로 확인한다.`

const REWRITE = '방금 답변이 영어 위주로 쓰였습니다. 같은 내용을 빠짐없이 한국어로 다시 써 주세요. (korean-only mod 자동 요청)'

let isRewriting = false // 재작성 요청이 무한 반복되지 않게 1회만

export const register: Register = on => {
  // 매 요청마다 시스템 프롬프트 끝에 한국어 규칙을 붙인다
  on('prompt.compose', async ($, e, next) => {
    const composed = await next(e)
    return { ...composed, sections: [...composed.sections, { id: 'korean-only:rule', text: RULE, scope: 'session' as const }] }
  })

  // 답변이 끝나면 영어 위주인지 검사 → 경고 + 한국어 재작성 자동 요청
  on('turn.complete', async ($, e, next) => {
    const done = await next(e)
    if (isRewriting) {
      isRewriting = false
      return done
    }
    if (e.reason === 'answer' && englishHeavy(done.text)) {
      isRewriting = true
      $.ui.toast('⚠ 영어 답변 감지 — 한국어로 다시 쓰게 했습니다')
      void $.prompt.submit({ text: REWRITE })
    }
    return done
  })
}
