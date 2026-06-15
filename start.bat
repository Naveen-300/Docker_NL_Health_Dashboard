@echo off
title AI Docker NL Dashboard

echo Starting Docker Desktop...
start "" "C:\Program Files\Docker\Docker\Docker Desktop.exe"

echo Waiting for Docker to be ready...
:wait
timeout /t 3 /nobreak >nul
docker ps >nul 2>&1
if errorlevel 1 goto wait

echo Docker is running!
echo Starting Dashboard...

cd /d "%~dp0"

start "" C:\Users\navee\AppData\Local\Programs\Python\Python313\python.exe -m streamlit run app.py --server.port 8501 --theme.base light --theme.backgroundColor "#f5f7fa" --theme.secondaryBackgroundColor "#ffffff" --theme.textColor "#2c3e50" --theme.primaryColor "#27ae60"

echo Waiting for app to start...
timeout /t 6 /nobreak >nul

start "" http://localhost:8501
