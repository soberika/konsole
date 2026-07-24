@echo off
REM =====================================================================
REM  Abgleich starten - kein Active Directory, keine Installation noetig.
REM
REM  NUTZUNG:
REM    * Doppelklick: nimmt SOLL (xlsx/csv) und IST (txt) aus diesem Ordner.
REM    * ODER beide Listen auf diese Datei ZIEHEN (Drag and Drop).
REM =====================================================================
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Abgleich.ps1" %*
echo.
echo Fertig. Fenster kann geschlossen werden.
pause
