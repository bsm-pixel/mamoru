@echo off
rem MAMORU TMS daily DB backup (run by Windows Task Scheduler) - 2026-11-15
cd /d "%~dp0.."
set PYTHONIOENCODING=utf-8
py scripts\backup-db.py >> "%~dp0backup-db.log" 2>&1
