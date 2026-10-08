@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo ================================================================
echo   CATALOGUE PREP - FINISH  (Claude files -^> upload files)
echo ================================================================
echo.
python catprep.py finish
echo.
pause
