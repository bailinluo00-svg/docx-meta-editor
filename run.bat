@echo off
rem ============================================================
rem  DocxMeta Editor - GUI launcher
rem
rem  Keep this file ASCII-only, and keep it saved with CRLF line
rem  endings. cmd.exe parses .bat files line by line and rewrites
rem  its own read position by byte offset, so a file with LF-only
rem  endings and multi-byte characters can be mis-parsed.
rem
rem  This file deliberately contains NO non-ASCII characters, so it
rem  survives any copy/transfer that mangles encoding. The GUI and
rem  messages are fully in Chinese (those live in the .ps1 files,
rem  which carry a UTF-8 BOM).
rem ============================================================

setlocal
set "SCRIPT=%~dp0DocxMetaGui.ps1"

if not exist "%SCRIPT%" (
    echo.
    echo [ERROR] DocxMetaGui.ps1 was not found next to this launcher.
    echo         Expected: "%SCRIPT%"
    echo         Keep run.bat and both .ps1 files in the same folder.
    echo.
    pause
    exit /b 1
)

where powershell.exe >nul 2>nul
if errorlevel 1 (
    echo.
    echo [ERROR] powershell.exe not found on PATH.
    echo         This tool needs Windows PowerShell 5.1 ^(built into Windows 7+^).
    echo.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT%"

if errorlevel 1 (
    echo.
    echo The tool exited with an error. See the message above.
    echo.
    pause
)

endlocal
