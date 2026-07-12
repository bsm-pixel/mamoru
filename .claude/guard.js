/**
 * PreToolUse 안전 가드 — 되돌릴 수 없는 명령만 하드 차단.
 *
 * 설계 의도:
 *  - 사장님이 계획 승인 후 클릭 없이 자율 수행하게 하되(팝업 0), 사고성 파괴 명령만 막는다.
 *  - "언급"과 "실행"을 구분한다: 따옴표 안 문자열(커밋 메시지·echo·grep 패턴 등)은 검사에서 제외.
 *    (예: git commit -m "rm -rf 금지" → 차단 안 함 / cd x && rm -rf y → 차단)
 *  - 이 PC엔 jq가 없으므로 node로 stdin JSON을 파싱한다.
 *
 * 보안 경계가 아니라 "사고 방지 안전망"이다. bash -c "..." 처럼 따옴표로 감싸 우회하는 건 막지 않는다.
 */
let raw = '';
process.stdin.on('data', d => (raw += d));
process.stdin.on('end', () => {
  let cmd = '';
  try {
    cmd = (JSON.parse(raw).tool_input || {}).command || '';
  } catch (e) {
    cmd = '';
  }

  // 따옴표로 감싼 구간 제거 → 실제 실행되는 "맨 명령"만 남긴다
  const Q = String.fromCharCode(34);
  const bare = cmd
    .replace(/'[^']*'/g, ' ')
    .replace(new RegExp(Q + '[^' + Q + ']*' + Q, 'g'), ' ');

  const DANGER = [
    /(^|[;&|(]|\s)rm\s+-[a-zA-Z]*r[a-zA-Z]*f/i,      // rm -rf / -fr
    /(^|[;&|(]|\s)rm\s+-[a-zA-Z]*f[a-zA-Z]*r/i,
    /(^|[;&|(]|\s)sudo\s+rm/i,
    /git\s+push\s[^;|&]*--force/i,                    // --force, --force-with-lease
    /git\s+push\s+([^;|&]*\s)?-f(\s|$)/i,             // -f (push 직후 / 인자 뒤 둘 다)
    /git\s+reset\s+--hard/i,
    /git\s+clean\s+-[a-zA-Z]*f/i,
    /git\s+filter-branch/i,
    /(drop|truncate)\s+(table|database|schema)/i,     // DB 파괴
    /(^|[;&|(]|\s)mkfs/i,
    /(^|[;&|(]|\s)dd\s+if=/i,
  ];

  if (DANGER.some(re => re.test(bare))) {
    process.stdout.write(JSON.stringify({
      hookSpecificOutput: {
        hookEventName: 'PreToolUse',
        permissionDecision: 'deny',
        permissionDecisionReason:
          '되돌릴 수 없는 명령이라 자동 차단했습니다 (rm -rf · force push · reset --hard · git clean -f · DB drop/truncate 등). ' +
          '정말 필요하면 사장님이 직접 실행하시거나, 저에게 이유를 알려주세요. ' +
          '규칙 위치: .claude/guard.js',
      },
    }));
  }
});