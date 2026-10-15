@echo off
REM Einmalig nach dem Verschieben: Regeln dorthin kopieren, wo Claude Code sie erwartet (.claude\).
cd /d "%~dp0\.."
if not exist .claude mkdir .claude
copy /Y docs\regeln\entwicklungsregeln.md .claude\entwicklungsregeln.md
copy /Y docs\regeln\design-guideline.md .claude\design-guideline.md
echo Fertig.
