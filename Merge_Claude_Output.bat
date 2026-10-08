@echo off
chcp 65001 >nul
cd /d "e:\FreeBuff\Image URL work agent"

echo.
echo ========================================
echo   STAGE 1 of 2: Merging Claude Cleaned Output
echo ========================================
echo.
python agent.py --auto --step 9
if errorlevel 1 goto :step9_error

echo.
echo ========================================
echo   STAGE 2 of 2: Building Final Upload-Ready Files
echo   Content update, official template, split
echo ========================================
echo.
python agent.py --auto --step 6
if errorlevel 1 goto :step6_error

echo.
echo ========================================
echo   DONE! Upload-ready files are in the update folder.
echo   D:\Image URL change\...\update\
echo ========================================
echo.
pause
exit /b 0

:step9_error
echo.
echo [ERROR] Merge Step 9 failed - dekho upore ki error eshechhe.
echo Step 6 ar chalano hocche na.
echo.
pause
exit /b 1

:step6_error
echo.
echo [ERROR] Step 6 failed - dekho upore ki error eshechhe.
echo.
pause
exit /b 1
