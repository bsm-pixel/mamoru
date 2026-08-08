@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 참석 응답을 가져오는 중...
node tools\rsvp_export.mjs
if errorlevel 1 (
  echo.
  echo [!] 실패했습니다. 위 메시지를 확인하세요.
  pause
  exit /b 1
)
echo.
echo 브라우저로 응답 목록을 엽니다...
start "" "BSMKHJ\_rsvp_responses.html"
timeout /t 2 >nul
