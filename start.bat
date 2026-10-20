@echo off
cd /d "%~dp0"
if not exist node_modules (
  echo Installing dependencies...
  call npm install --omit=dev
)
rem server.js creates the localhost certificate and asks Windows to trust it on first run.
node server.js
pause
