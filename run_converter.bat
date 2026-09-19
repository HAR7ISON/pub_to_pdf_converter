@echo off
setlocal DisableDelayedExpansion
rem The drive is configured at the top of convertpubtopdf.ps1.
rem The script writes a timestamped conversion_log_*.txt beside these files.
rem Pass the path through the environment so spaces and apostrophes are safe.
set "CONVERTER_SCRIPT=%~dp0convertpubtopdf.ps1"

rem If already elevated, run directly without another UAC prompt.
powershell.exe -NoLogo -NoProfile -NonInteractive -Command "if (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { exit 0 } else { exit 1 }"
if not errorlevel 1 goto run_elevated

echo Requesting administrator privileges. Please approve the Windows UAC prompt.
rem Wait for the elevated converter and preserve its exit code.
powershell.exe -NoLogo -NoProfile -NonInteractive -Command "try { $arguments = '-NoLogo -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File ' + [char]34 + $env:CONVERTER_SCRIPT + [char]34; $process = Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $arguments -Verb RunAs -Wait -PassThru -ErrorAction Stop; exit $process.ExitCode } catch { Write-Host ('Administrator launch was cancelled or failed: ' + $_.Exception.Message); exit 1 }"
set "CONVERTER_EXIT_CODE=%ERRORLEVEL%"
goto report_result

:run_elevated
powershell.exe -NoLogo -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File "%~dp0convertpubtopdf.ps1"
set "CONVERTER_EXIT_CODE=%ERRORLEVEL%"

:report_result
echo.
if "%CONVERTER_EXIT_CODE%"=="0" (
    echo Conversion scan completed. See the log beside this batch file.
) else (
    echo Conversion failed to start or reported errors. See console output and the log if created.
)
pause
exit /b %CONVERTER_EXIT_CODE%
