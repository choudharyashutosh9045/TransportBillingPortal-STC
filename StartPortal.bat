@echo off
title STC Billing Portal - Auto Start
cd /d C:\Users\Onmove\Desktop\TransportBillingPortal

echo ==========================================
echo   STARTING STC BILLING PORTAL...
echo ==========================================
echo.

REM Activate venv
call venv\Scripts\activate

REM Start Flask in new window
start "Flask Server" cmd /k "cd /d C:\Users\Onmove\Desktop\TransportBillingPortal && call venv\Scripts\activate && python app.py"

REM Wait 2 seconds
timeout /t 2 /nobreak >nul

REM Start Ngrok in new window
start "Ngrok Tunnel" cmd /k "ngrok http 5000"

REM Wait 2 seconds
timeout /t 2 /nobreak >nul

REM Open local portal
start http://127.0.0.1:5000

echo.
echo DONE! Flask + Ngrok started.
echo Close windows to stop.
pause
