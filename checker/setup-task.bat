@echo off
REM ============================================================
REM Setup Windows Task Scheduler to run validation every hour
REM Run this script as Administrator
REM ============================================================

REM --- Edit these paths ---
SET CHECKER_DIR=%~dp0
SET SHARED_DRIVE=Z:\path\to\shared\drive

echo Creating scheduled task: ExcelValidationChecker
echo Checker dir: %CHECKER_DIR%
echo Target dir:  %SHARED_DRIVE%
echo Schedule:    Every 1 hour
echo.

schtasks /create ^
  /tn "ExcelValidationChecker" ^
  /tr "node \"%CHECKER_DIR%check.js\" \"%SHARED_DRIVE%\" \"%CHECKER_DIR%reports\"" ^
  /sc hourly ^
  /mo 1 ^
  /f

if %errorlevel% equ 0 (
  echo.
  echo Task created successfully!
  echo To run it now:  schtasks /run /tn "ExcelValidationChecker"
  echo To delete it:   schtasks /delete /tn "ExcelValidationChecker" /f
) else (
  echo.
  echo Failed to create task. Make sure you run this as Administrator.
)

pause
