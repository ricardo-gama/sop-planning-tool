@echo off
REM ============================================================
REM  SOP Planning — one-click launcher
REM  Double-click this file (or a shortcut to it) to start the
REM  dev server and open Excel with the add-in loaded.
REM ============================================================

cd /d "%~dp0"

echo Starting SOP Planning dev server...
echo (Keep this window open while you work — closing it stops the server)
echo.

start "SOP Planning - dev server" cmd /k "npm run dev-server"

echo Waiting for the dev server to be ready...
timeout /t 8 /nobreak >nul

echo Launching Excel with the add-in...
call npm start

pause