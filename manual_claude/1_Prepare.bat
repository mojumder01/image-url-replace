@echo off
chcp 65001 >nul
cd /d "%~dp0.."
echo.
echo ================================================================
echo   CATALOGUE PREP - PREPARE  (mapping file -^> parts for Claude)
echo   If it stopped half-way, run 1_Prepare_Resume.bat
echo ================================================================
echo.
python catprep.py prepare
echo.
pause
