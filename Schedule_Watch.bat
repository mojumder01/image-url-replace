@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Checks the Data folder every few minutes (config.toml [auto] watch_minutes).
echo A new mapping file is processed automatically.
python catprep.py schedule --watch
echo.
pause
