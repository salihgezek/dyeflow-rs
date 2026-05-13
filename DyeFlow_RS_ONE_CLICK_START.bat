@echo off
title DyeFlow RS Web v45 - Premium Report Compare Engine
cd /d "%~dp0"

echo.
echo ============================================
echo   DyeFlow RS v45 Premium Report Engine
echo ============================================
echo.

py --version >nul 2>&1
if errorlevel 1 (
    echo Python bulunamadi. Lutfen once Python kurun:
    echo https://www.python.org/downloads/
    pause
    exit /b
)

echo Gerekli paketler kontrol ediliyor...
py -m pip install -r requirements.txt

echo.
echo Tarayici aciliyor...
start "" "http://127.0.0.1:8002"

echo.
echo Program calisiyor.
echo Bu siyah pencere acik kaldigi surece uygulama calisir.
echo.

py -m uvicorn main:app --host 127.0.0.1 --port 8002
pause
