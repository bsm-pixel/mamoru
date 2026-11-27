# 제브라 ZT231 주소 따라가기 (2026-11-27)
#
# 왜: 공유기가 재부팅 등으로 프린터 주소를 바꿔 주면, ZebraDesigner가 쓰는 Windows 프린터 연결(ZDesigner ZT231)은
#     옛 주소로 계속 보낸다. 그 주소를 다른 기기(삼성 프린터)가 받으면 라벨이 그쪽으로 나간다(2026-11-27 실제 사고).
# 하는 일: 지금 연결 주소에 제브라(일련번호로 확인)가 있으면 아무것도 안 한다.
#          없으면 같은 망에서 9100 포트가 열린 기기 중 일련번호가 맞는 제브라를 찾아 연결 주소만 바꾸고 인쇄 서비스를 다시 켠다.
# 건드리지 않는 것: TMS 라벨 인쇄(Browser Print는 제브라를 네트워크에서 직접 찾음) · 삼성 등 다른 프린터 · 공유기 · 프린터 본체 설정.
# 설치: install.ps1 (관리자) — 작업 스케줄러 「MAMORU 제브라 주소 따라가기」, 켜질 때 + 10분마다, SYSTEM 계정.
# 되돌리기: uninstall.ps1 (관리자)

$Serial  = 'T3J260604332'                     # 제브라 ZT231 일련번호(관리 화면 제목)
$PortKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Monitors\ZDesigner Port Monitor\Ports\LAN_ZDesigner_ZT231-203dpi_ZPL'
$Dir     = 'C:\ProgramData\MamoruZebraWatch'
$Log     = Join-Path $Dir 'log.txt'

function Write-WatchLog($m) {
  if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Path $Dir -Force | Out-Null }
  Add-Content -Path $Log -Value ("{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m) -Encoding UTF8
  $lines = Get-Content $Log -Encoding UTF8
  if ($lines.Count -gt 300) { $lines | Select-Object -Last 200 | Set-Content $Log -Encoding UTF8 }
}

function Test-IsZebra($ip) {
  # 9100 포트가 열려 있고, 관리 화면에 일련번호가 보이면 제브라
  $c = New-Object System.Net.Sockets.TcpClient
  try {
    $t = $c.ConnectAsync($ip, 9100)
    if (-not $t.Wait(1500) -or -not $c.Connected) { return $false }
  } catch { return $false } finally { $c.Dispose() }
  try {
    $r = Invoke-WebRequest -Uri "http://$ip/" -UseBasicParsing -TimeoutSec 5
    return ($r.Content -match [regex]::Escape($Serial))
  } catch { return $false }
}

try {
  if (-not (Test-Path $PortKey)) { Write-WatchLog '제브라 연결(LAN_ZDesigner_ZT231) 없음 — 아무것도 안 함'; exit 0 }
  $current = (Get-ItemProperty $PortKey).IPAddress
  if ($current -and (Test-IsZebra $current)) { exit 0 }   # 정상 — 기록도 남기지 않는다

  # 같은 망(기본 게이트웨이가 있는 연결) 안에서 찾기
  $cfg = Get-NetIPConfiguration | Where-Object { $_.IPv4DefaultGateway } | Select-Object -First 1
  if (-not $cfg) { Write-WatchLog '네트워크 연결 없음 — 다음 회차에'; exit 0 }
  $myIp = $cfg.IPv4Address.IPAddress
  $prefix = ($myIp -split '\.')[0..2] -join '.'

  $probes = foreach ($i in 1..254) {
    $ip = "$prefix.$i"
    if ($ip -eq $myIp) { continue }
    $c = New-Object System.Net.Sockets.TcpClient
    [pscustomobject]@{ ip = $ip; c = $c; t = $c.ConnectAsync($ip, 9100) }
  }
  Start-Sleep -Milliseconds 2000
  $open = @()
  foreach ($p in $probes) { if ($p.t.Status -eq 'RanToCompletion' -and $p.c.Connected) { $open += $p.ip }; $p.c.Dispose() }

  $found = $null
  foreach ($ip in $open) { if (Test-IsZebra $ip) { $found = $ip; break } }

  if (-not $found) { Write-WatchLog "제브라를 못 찾음(지금 연결 주소 $current · 9100 열린 기기: $($open -join ', ')) — 꺼져 있을 수 있음, 그대로 둠"; exit 0 }
  if ($found -eq $current) { exit 0 }

  Set-ItemProperty -Path $PortKey -Name IPAddress -Value $found
  Restart-Service -Name Spooler -Force
  Write-WatchLog "제브라 주소 바뀜: $current → $found (ZebraDesigner 연결 고침 · 인쇄 서비스 다시 켬)"
} catch {
  Write-WatchLog ("오류: " + $_.Exception.Message)
}
