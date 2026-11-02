@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0create_scheduled_task.ps1"
pause
