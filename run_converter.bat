@echo off
setlocal DisableDelayedExpansion
rem The drive is configured at the top of convertpubtopdf.ps1.
rem The script writes a timestamped conversion_log_*.txt beside these files.
rem Use the current user session: Store-installed Publisher may not activate elevated.
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File "%~dp0convertpubtopdf.ps1"
set "CONVERTER_EXIT_CODE=%ERRORLEVEL%"

echo.
if "%CONVERTER_EXIT_CODE%"=="0" (
    echo Conversion scan completed. See the totals above.
) else if "%CONVERTER_EXIT_CODE%"=="2" (
    echo Conversion scan completed. See the totals above; additional details are in the log.
) else (
    echo Conversion could not finish. See console output and the log if created.
)
pause
exit /b %CONVERTER_EXIT_CODE%
