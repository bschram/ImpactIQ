@echo off
setlocal EnableExtensions EnableDelayedExpansion
REM =====================================================================================================================
REM  ImpactIQ launcher (Windows). Runs ImpactIQ.ps1 from the folder this file is in, with the same interactive
REM  experience as "Final PS Script" (environment dialog, sign-in, pickers) unless told otherwise.
REM
REM    ImpactIQ.bat                       interactive run (dialogs, browser sign-in)
REM    ImpactIQ.bat --no-prompts          no dialogs: scope and environment come from Config\ImpactIQ.Settings.json,
REM                                       environment variables or the extra parameters below
REM    ImpactIQ.bat --resume              pick up an unfinished run even if it is older than 3 days (-Resume Always)
REM    ImpactIQ.bat --fresh               wipe today's state and backup folders first (-Force)
REM    ImpactIQ.bat --environment USGov   cloud (Public, USGov/GCC, USGovHigh/GCCHigh, USGovMil/DoD, China, Germany)
REM    ImpactIQ.bat --settings <file>     another settings file than Config\ImpactIQ.Settings.json
REM    ImpactIQ.bat --help                this text
REM
REM  Anything else on the command line is passed to ImpactIQ.ps1 unchanged, e.g.
REM    ImpactIQ.bat --no-prompts -WorkspaceName "Finance*" -IncludeMyWorkspace -TimeBudgetMinutes 55
REM  Exit code = ImpactIQ's (0 ok, 2 item failures, 3 paused by the time budget, 1 fatal). Logs: Logs\ImpactIQ_*.log
REM =====================================================================================================================
set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "ENTRY=%SCRIPT_DIR%\ImpactIQ.ps1"
set "PS_ARGS=-AuthMode Interactive"
set "PAUSE_AT_END=1"
set "EXTRA="

:parse
if "%~1"=="" goto run
if /i "%~1"=="--help"        goto help
if /i "%~1"=="-h"            goto help
if /i "%~1"=="/?"            goto help
if /i "%~1"=="--no-prompts"  ( set "PS_ARGS=-NonInteractive" & set "PAUSE_AT_END=0" & shift & goto parse )
if /i "%~1"=="--resume"      ( set "EXTRA=!EXTRA! -Resume Always" & shift & goto parse )
if /i "%~1"=="--fresh"       ( set "EXTRA=!EXTRA! -Force" & shift & goto parse )
if /i "%~1"=="--environment" ( set "EXTRA=!EXTRA! -Environment %~2" & shift & shift & goto parse )
if /i "%~1"=="--settings"    ( set "EXTRA=!EXTRA! -SettingsPath "%~2"" & shift & shift & goto parse )
set "EXTRA=!EXTRA! %1"
shift
goto parse

:help
echo.
echo ImpactIQ launcher options:
echo   --no-prompts          run without dialogs (scope/environment from Config\ImpactIQ.Settings.json, env vars or parameters)
echo   --resume              resume an unfinished run of any age (-Resume Always)
echo   --fresh               wipe today's state and backups first (-Force)
echo   --environment NAME    Public, USGov/GCC, USGovHigh/GCCHigh, USGovMil/DoD, China, Germany
echo   --settings FILE       use another settings file
echo   --help                this text
echo   anything else is passed to ImpactIQ.ps1 unchanged (see README, "Command-line reference")
echo.
exit /b 0

:run
if not exist "%ENTRY%" (
  echo [ERROR] ImpactIQ.ps1 not found next to this file: "%ENTRY%"
  echo         Download the whole repository into one folder and run ImpactIQ.bat from there.
  if "%PAUSE_AT_END%"=="1" pause
  exit /b 1
)
echo [INFO] ImpactIQ - running "%ENTRY%" %PS_ARGS%!EXTRA!
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%ENTRY%" -BaseFolder "%SCRIPT_DIR%" %PS_ARGS%!EXTRA!
set "RC=%ERRORLEVEL%"
if "%RC%"=="0" ( echo [DONE] ImpactIQ finished. Open "Power BI Governance Model.pbit" and point Base Directory at "%SCRIPT_DIR%\Outputs". ) else (
  if "%RC%"=="2" ( echo [DONE] ImpactIQ finished with item failures - see the Failures sheet; the next run retries them. ) else (
    if "%RC%"=="3" ( echo [PAUSED] Time budget reached - run ImpactIQ.bat again to resume. ) else ( echo [FAILED] ImpactIQ exit code %RC% - see the newest file in "%SCRIPT_DIR%\Logs". )
  )
)
if "%PAUSE_AT_END%"=="1" pause
exit /b %RC%
