@echo off
rem Started by Windows Task Scheduler. Output goes to scheduler.log.
chcp 65001 >nul
cd /d "%~dp0"
python catprep.py auto --if-new >> "%~dp0scheduler.log" 2>&1
