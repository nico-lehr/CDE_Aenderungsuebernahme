@echo off
REM Einmalig: eigene Python-3.11-Umgebung anlegen und Pakete installieren.
REM Hinweis: Im Firmennetz kann pip an Zertifikaten scheitern. NICHT mit --trusted-host umgehen,
REM sondern Fehlermeldung melden.
cd /d "%~dp0\.."
py -3.11 -m venv .venv
call .venv\Scripts\activate
python -m pip install -r requirements.txt
pip install -e .
echo.
echo Fertig. Test: pytest    Start: scripts\start_dashboard.bat
