@echo off
REM Startet das CDE-Dashboard in der eigenen .venv (Port 5100, nicht 5000 wie die bess-pipeline).
cd /d "%~dp0\.."
if not exist ".venv\Scripts\python.exe" (
  echo .venv fehlt. Bitte zuerst scripts\setup.bat ausfuehren.
  exit /b 1
)
call .venv\Scripts\activate
python run_server.py
