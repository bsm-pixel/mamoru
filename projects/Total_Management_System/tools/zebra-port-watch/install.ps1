# 제브라 주소 따라가기 설치 (관리자 PowerShell) — 2026-11-27
# 스크립트를 C:\ProgramData\MamoruZebraWatch 로 복사하고, 작업 스케줄러에 켜질 때 + 10분마다(SYSTEM) 등록한다.
$ErrorActionPreference = 'Stop'
$Dir  = 'C:\ProgramData\MamoruZebraWatch'
$Name = 'MAMORU 제브라 주소 따라가기'
New-Item -ItemType Directory -Path $Dir -Force | Out-Null
Copy-Item -Path (Join-Path $PSScriptRoot 'zebra-port-watch.ps1') -Destination (Join-Path $Dir 'zebra-port-watch.ps1') -Force

$action   = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Dir\zebra-port-watch.ps1`""
$atStart  = New-ScheduledTaskTrigger -AtStartup
$every10  = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 10)
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -MultipleInstances IgnoreNew
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName $Name -Action $action -Trigger @($atStart, $every10) -Settings $settings -Principal $principal -Description '공유기가 제브라 주소를 바꾸면 ZebraDesigner 연결 주소를 따라 고친다. TMS 라벨 인쇄·다른 프린터는 건드리지 않음. 저장소 projects/Total_Management_System/tools/zebra-port-watch' -Force | Out-Null
Start-ScheduledTask -TaskName $Name
'설치 완료'
