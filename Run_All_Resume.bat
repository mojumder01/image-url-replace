@echo off
chcp 65001 >nul
cd /d "%~dp0"
python catprep.py auto --resume
echo.
pause
