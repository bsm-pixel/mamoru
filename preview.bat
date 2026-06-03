@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 실시간 미리보기 서버를 시작합니다... (이 창을 닫으면 종료)
node tools\preview-server.js
pause
