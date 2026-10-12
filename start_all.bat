@echo off
chcp 65001 >nul
title DotaArena 2026 - Server & CCTV Streams
echo ============================================================
echo   Zapusk DotaArena 2026 i CCTV kamer (Windows)
echo ============================================================
echo.

cd /d "%~dp0"

REM 1. Zapusk go2rtc
if exist "..\tools\go2rtc.exe" (
    echo [*] Zapusk shlyuza kamer go2rtc...
    start /b "" "..\tools\go2rtc.exe" -config "go2rtc.yaml" >nul 2>&1
) else if exist "go2rtc.exe" (
    echo [*] Zapusk shlyuza kamer go2rtc...
    start /b "" "go2rtc.exe" -config "go2rtc.yaml" >nul 2>&1
)

REM 2. Otkrytie sayta v brauzere
echo [*] Otkrytie sayta v brauzere...
start http://localhost:8000/

echo.
echo ============================================================
echo   VSE SERVERA ZAPUSHCHENY:
echo   - Sayt i Baza Dannykh: http://localhost:8000
echo   - Shlyuz kamer go2rtc: http://127.0.0.1:1984
echo ============================================================
echo Dlya ostanovki servera zakroyte eto okno ili nazhmite Ctrl+C.
echo.

py server.py
pause
