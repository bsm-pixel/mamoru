# 작업 폴더 → 구글 드라이브 미러 (2026-09-24)
#
#   원본(진짜)  : C:\Users\user\Desktop\mamoru        ← 여기서만 작업·편집한다
#   미러(사본)  : C:\마모루드라이브\ClaudeWork\mamoru  ← 열람·백업용. 여기서 편집하지 말 것(다음 미러 때 덮어씀)
#
# 왜 한 방향인가: 양쪽에서 고치면 두 PC 사이에서 충돌이 난다. 진실은 git, 드라이브는 사본.
# 제외: .git(GitHub 에 있음) · node_modules·.next(언제든 다시 생성) · __pycache__
# 포함: .env.local 등 설정 파일 (2026-09-24 사장님 결정 — 새 PC 에서 바로 작업하기 위함)
#
# 실행: 커밋할 때마다 자동(.githooks/post-commit) · 수동은  powershell -File scripts\mirror-to-drive.ps1

$ErrorActionPreference = 'Stop'
$src = 'C:\Users\user\Desktop\mamoru'
$dst = 'C:\마모루드라이브\ClaudeWork\mamoru'

# /MIR 는 대상에만 있는 파일을 지운다 → 엉뚱한 폴더를 지우지 않도록 경로를 고정 확인
if ($dst -notlike '*\ClaudeWork\mamoru') { Write-Host '미러 대상 경로가 예상과 다릅니다. 중단.'; exit 1 }
if (-not (Test-Path 'C:\마모루드라이브')) { Write-Host '구글 드라이브 미러 폴더가 없습니다(동기화 꺼짐?). 건너뜁니다.'; exit 0 }
if (-not (Test-Path $dst)) { New-Item -ItemType Directory -Path $dst -Force | Out-Null }

$sw = [Diagnostics.Stopwatch]::StartNew()
# /MIR 미러 · /XD 폴더 제외 · /R:1 /W:1 재시도 최소 · /NFL /NDL 파일목록 생략(로그 폭주 방지)
robocopy $src $dst /MIR /XD '.git' 'node_modules' '.next' '__pycache__' '.vercel' /XF 'desktop.ini' '_마지막미러.txt' `
  /R:1 /W:1 /NFL /NDL /NJH /NP | Out-Null
$code = $LASTEXITCODE
$sw.Stop()

# robocopy: 0~7 정상(0=변경없음, 1=복사됨, 2=추가정리, 3=둘다) / 8 이상만 실패
if ($code -ge 8) {
  Write-Host ("미러 실패 (robocopy {0})" -f $code)
  exit 0   # 커밋을 막지 않는다
}
$stamp = "마지막 미러: {0} ({1:N1}초, robocopy {2})" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $sw.Elapsed.TotalSeconds, $code
Set-Content -Path (Join-Path $dst '_마지막미러.txt') -Value $stamp -Encoding utf8
Write-Host $stamp
exit 0
