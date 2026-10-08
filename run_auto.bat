@echo off
chcp 65001 >nul
cd /d "e:\FreeBuff\Image URL work agent"
echo.
echo ========================================
echo   Catalogue Prep Agent - AUTO MODE
echo ========================================
echo.
python agent.py --auto
echo.
pause
