@echo off
rem ============================================================
rem  Runs both test suites. ASCII-only, keep CRLF line endings.
rem ============================================================
setlocal
set "DIR=%~dp0"

echo.
echo ================ CORE LOGIC TESTS ================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%DIR%_test-core.ps1"
if errorlevel 1 goto failed

echo.
echo ================ GUI TESTS ================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%DIR%_test-gui.ps1"
if errorlevel 1 goto failed

echo.
echo All suites finished. Check the "FAILURES: 0" lines above.
echo.
pause
exit /b 0

:failed
echo.
echo A suite reported failures. See the output above.
echo.
pause
exit /b 1
