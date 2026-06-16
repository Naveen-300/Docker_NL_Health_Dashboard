@echo off
title AI Docker NL Dashboard Launcher

:: Start Docker Desktop if not running
tasklist /FI "IMAGENAME eq Docker Desktop.exe" | find /I "Docker Desktop.exe" >nul
if errorlevel 1 (
    echo Starting Docker Desktop...
    start "" "C:\Program Files\Docker\Docker\Docker Desktop.exe"
    timeout /t 20 /nobreak >nul
)

:: Wait until Docker is ready
echo Waiting for Docker...
:waitdocker
timeout /t 3 /nobreak >nul
docker ps >nul 2>&1
if errorlevel 1 goto waitdocker

echo Docker is ready!

:: Kill any old instance on port 8501
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :8501 ^| findstr LISTENING') do taskkill /PID %%a /F >nul 2>&1

:: Start the app
echo Starting Dashboard...
start "" C:\Users\navee\AppData\Local\Programs\Python\Python313\python.exe -m streamlit run "%~dp0app.py" --theme.base light --theme.backgroundColor "#f5f7fa" --theme.secondaryBackgroundColor "#ffffff" --theme.textColor "#2c3e50" --theme.primaryColor "#27ae60"

:: Wait for app to start
timeout /t 6 /nobreak >nul

:: Open browser
start "" http://localhost:8501
echo Dashboard is running at http://localhost:8501
