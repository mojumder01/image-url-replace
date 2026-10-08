@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo ================================================================
echo   CATALOGUE PREP - FULLY AUTOMATIC RUN
echo   newest mapping file in Data -^> cleaned -^> upload files
echo ================================================================
echo.
python catprep.py auto
echo.
pause
