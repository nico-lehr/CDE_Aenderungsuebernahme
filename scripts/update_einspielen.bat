@echo off
REM Spielt eine von GitHub heruntergeladene ZIP-Datei ueber deinen Projektordner ein.
REM Deine eigenen Daten bleiben unberuehrt: data\, .venv\, .claude\, .git\ und alle *.local.yaml.
REM Es wird NICHTS geloescht. Vorher wird eine Sicherung des Codes angelegt (in data\update_sicherung).
REM
REM Aufruf:  scripts\update_einspielen.bat "C:\Pfad\CDE_Aenderungsuebernahme-main.zip" ["C:\Pfad\zum\Projektordner"]
REM Ohne zweiten Wert gilt der Ordner, in dem dieses Skript liegt (eine Ebene hoeher).
REM Erster Lauf: bitte zuerst den ganzen Projektordner einmal von Hand kopieren (zur Sicherheit).

if /I not "%~1"=="--kopie" (
    copy /Y "%~f0" "%TEMP%\cde_update_lauf.bat" >nul
    call "%TEMP%\cde_update_lauf.bat" --kopie "%~1" "%~2" "%~dp0.."
    exit /b %errorlevel%
)
REM Ab hier laeuft die Kopie im TEMP-Ordner (damit dieses Skript sich selbst aktualisieren kann).
setlocal
set "ZIP=%~2"
set "ZIEL=%~3"
if "%ZIEL%"=="" set "ZIEL=%~4"
for %%I in ("%ZIEL%") do set "ZIEL=%%~fI"

if "%ZIP%"=="" (
    echo Bitte die ZIP-Datei angeben. Beispiel: scripts\update_einspielen.bat "C:\Downloads\CDE_Aenderungsuebernahme-main.zip"
    exit /b 2
)
if not exist "%ZIP%" (
    echo Die ZIP-Datei wurde nicht gefunden: %ZIP%
    exit /b 2
)
if not exist "%ZIEL%\README.md" (
    echo Der Zielordner sieht nicht wie das Projekt aus ^(README.md fehlt^): %ZIEL%
    exit /b 2
)
if not exist "%ZIEL%\src" (
    echo Der Zielordner sieht nicht wie das Projekt aus ^(Ordner src fehlt^): %ZIEL%
    exit /b 2
)

for /f %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "ZEIT=%%T"
set "TMP_ENTPACKT=%TEMP%\cde_update_%ZEIT%"
set "SICHERUNG=%ZIEL%\data\update_sicherung\%ZEIT%"

echo.
echo 1/4  Entpacke die ZIP-Datei ...
powershell -NoProfile -Command "Expand-Archive -LiteralPath '%ZIP%' -DestinationPath '%TMP_ENTPACKT%' -Force"
if errorlevel 1 (
    echo Das Entpacken ist fehlgeschlagen.
    exit /b 3
)
set "QUELLE="
for /d %%D in ("%TMP_ENTPACKT%\*") do if not defined QUELLE set "QUELLE=%%~fD"
if not defined QUELLE (
    echo In der ZIP-Datei wurde kein Projektordner gefunden.
    exit /b 3
)
if not exist "%QUELLE%\README.md" (
    echo Die ZIP-Datei enthaelt nicht das Projekt ^(README.md fehlt^).
    exit /b 3
)

echo 2/4  Lege eine Sicherung des bisherigen Codes an: %SICHERUNG%
for %%O in (src config docs scripts tests frontend) do (
    if exist "%ZIEL%\%%O" robocopy "%ZIEL%\%%O" "%SICHERUNG%\%%O" /E /XD __pycache__ /NFL /NDL /NJH /NJS /NP >nul
)
for %%F in (README.md CLAUDE.md requirements.txt pyproject.toml run_server.py .gitignore) do (
    if exist "%ZIEL%\%%F" copy /Y "%ZIEL%\%%F" "%SICHERUNG%\" >nul
)

echo 3/4  Spiele die neuen Dateien ein ^(nichts wird geloescht^) ...
robocopy "%QUELLE%" "%ZIEL%" /E /XD data .venv .claude .git __pycache__ .pytest_cache /XF *.local.yaml /NDL /NJH /NJS /NP
if errorlevel 8 (
    echo Beim Kopieren ist ein Fehler aufgetreten ^(Code %errorlevel%^). Die Sicherung liegt in %SICHERUNG%.
    exit /b 4
)

echo 4/4  Raeume auf ...
rmdir /S /Q "%TMP_ENTPACKT%" 2>nul

echo.
echo Fertig. Deine Daten ^(data\, .venv\, *.local.yaml^) wurden nicht veraendert.
echo Naechste Schritte:
echo   - Hat sich requirements.txt geaendert:  scripts\setup.bat
echo   - Tests:                               .venv\Scripts\python -m pytest -q
echo Eine Sicherung des alten Codes liegt in: %SICHERUNG%
exit /b 0
