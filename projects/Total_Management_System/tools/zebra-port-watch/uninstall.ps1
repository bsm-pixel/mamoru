# 제브라 주소 따라가기 끄기 (관리자 PowerShell) — 작업 스케줄러 등록만 지운다. 기록(C:\ProgramData\MamoruZebraWatch\log.txt)은 남긴다.
Unregister-ScheduledTask -TaskName 'MAMORU 제브라 주소 따라가기' -Confirm:$false
'꺼짐'
