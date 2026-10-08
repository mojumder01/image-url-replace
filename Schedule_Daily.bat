@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Runs once a day (config.toml [auto] daily_time) if there is a new mapping file.
python catprep.py schedule --daily
echo.
pause
