# MAMORU 클로드 코드 mods

클로드 코드 화면·동작을 바꾸는 추가 장치. 모든 대화에서 자동으로 켜지려면 `~/.claude/settings.json`의 `env`에 아래 줄을 PC마다 1회 등록한다.

```json
"env": { "CLAUDE_CODE_PLUGIN_DIRS": "C:/Users/user/Desktop/mamoru/.claude/mods/korean-only;C:/Users/user/Desktop/mamoru/.claude/mods/work-resume" }
```

| mod | 하는 일 |
|---|---|
| `korean-only` | 매 요청마다 "답변은 한국어" 규칙을 시스템 프롬프트에 주입 + 답변이 영어 위주면 경고 후 한국어로 자동 재작성 |
| `work-resume` | 「이어할 작업」 패널 — 멈춘 작업 목록(`C:\마모루드라이브\ClaudeWork\작업이어하기.md`) + 이어하기(5초 중복 방지)·완료(두 번 눌러 확정)·최근 완료 되돌리기, 형식 오류·드라이브 충돌 사본 경고, `/pause` |

검사: `claude plugin validate <폴더>` · `claude plugin test <폴더>` (VS Code 확장에 들어 있는 claude.exe 2.1.301 이상으로 실행)
