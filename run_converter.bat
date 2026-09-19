@echo off
setlocal DisableDelayedExpansion
rem The drive is configured at the top of convertpubtopdf.ps1.
rem The script writes a timestamped conversion_log_*.txt beside these files.
rem Use the current user session: Store-installed Publisher may not activate elevated.
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File "%~dp0convertpubtopdf.ps1"
set "CONVERTER_EXIT_CODE=%ERRORLEVEL%"

pause
exit /b %CONVERTER_EXIT_CODE%
