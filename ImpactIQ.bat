@echo off
setlocal enableextensions enabledelayedexpansion
REM ======= Config =======
set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR:~-1%"=="\" set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"

set "TARGET_DIR=%SCRIPT_DIR%"
set "SETTINGS_JSON=%TARGET_DIR%\Config\ImpactIQ.Settings.json"

REM Optional override of target directory from settings file
if exist "%SETTINGS_JSON%" (
  for /f "usebackq delims=" %%I in (`powershell -NoProfile -Command "$p=[IO.Path]::GetFullPath('%SETTINGS_JSON%'); if(Test-Path $p){ try { $s=Get-Content -Raw -Path $p | ConvertFrom-Json; if($s.BaseFolderPath -and -not [string]::IsNullOrWhiteSpace([string]$s.BaseFolderPath)){ [IO.Path]::GetFullPath([string]$s.BaseFolderPath) } } catch {} }"`) do set "TARGET_DIR=%%I"
)

set "LOG_DIR=%TARGET_DIR%\Config\Logs"
set "SETTINGS_JSON=%TARGET_DIR%\Config\ImpactIQ.Settings.json"

set "CLI_HELP=false"
set "CLI_NO_PROMPTS=false"
set "CLI_WORKBOOK_MODE="
set "CLI_RESUME_MODE="

:parse_args
if "%~1"=="" goto args_done

if /i "%~1"=="--help" (
  set "CLI_HELP=true"
) else if /i "%~1"=="-h" (
  set "CLI_HELP=true"
) else if /i "%~1"=="/?" (
  set "CLI_HELP=true"
) else if /i "%~1"=="--no-prompts" (
  set "CLI_NO_PROMPTS=true"
  echo [INFO] Non-interactive mode enabled.
) else if /i "%~1"=="--overwrite" (
  set "CLI_WORKBOOK_MODE=overwrite"
  echo [INFO] Workbook mode set to overwrite.
) else if /i "%~1"=="--append" (
  set "CLI_WORKBOOK_MODE=append"
  echo [INFO] Workbook mode set to append.
) else if /i "%~1"=="--resume" (
  set "CLI_RESUME_MODE=resume"
  echo [INFO] Resume mode set to resume.
) else if /i "%~1"=="--fresh" (
  set "CLI_RESUME_MODE=fresh"
  echo [INFO] Resume mode set to fresh.
) else (
  echo [WARNING] Unknown argument ignored: %~1
)

shift
goto parse_args

:args_done

if /i "%CLI_HELP%"=="true" (
  echo.
  echo ImpactIQ launcher options:
  echo   --help, -h, /?      Show this help and exit.
  echo   --no-prompts         Run non-interactive using ImpactIQ.Settings.json values.
  echo   --overwrite          Force workbook mode to overwrite existing detail files.
  echo   --append             Force workbook mode to append to existing detail files.
  echo   --resume             Resume from existing checkpoint when found.
  echo   --fresh              Start fresh when checkpoint is found.
  echo.
  echo Settings file:
  echo   %SETTINGS_JSON%
  echo.
  echo Notes:
  echo   - BaseFolderPath in settings controls where ImpactIQ reads/writes files.
  echo   - Notifications are configured in ImpactIQ.Settings.json, not ImpactIQ.Notify.env.
  exit /b 0
)

REM ======= Ensure target directory exists =======
if not exist "%TARGET_DIR%" (
  echo [INFO] Target directory does not exist. Creating: "%TARGET_DIR%"
  mkdir "%TARGET_DIR%" >nul 2>&1
  if errorlevel 1 (
    echo [ERROR] Failed to create target directory: "%TARGET_DIR%"
    pause
    exit /b 1
  )
)

echo [INFO] Running Impact IQ from: "%TARGET_DIR%"
set "IMPACTIQ_SETTINGS_PATH=%SETTINGS_JSON%"
set "IMPACTIQ_BASE_FOLDER=%TARGET_DIR%"

if /i "%CLI_NO_PROMPTS%"=="true" (
  set "IMPACTIQ_NO_PROMPTS=true"
)

if not "%CLI_WORKBOOK_MODE%"=="" (
  set "IMPACTIQ_WORKBOOK_MODE=%CLI_WORKBOOK_MODE%"
)

if not "%CLI_RESUME_MODE%"=="" (
  set "IMPACTIQ_RESUME_MODE=%CLI_RESUME_MODE%"
)

REM ======= Ensure log directory exists =======
if not exist "%LOG_DIR%" (
  mkdir "%LOG_DIR%" >nul 2>&1
)

for /f %%i in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "RUN_STAMP=%%i"
set "RUN_LOG=%LOG_DIR%\ImpactIQ_Run_%RUN_STAMP%.log"
set "TRANSCRIPT_LOG=%LOG_DIR%\ImpactIQ_Transcript_%RUN_STAMP%.txt"

echo [INFO] Run log: "%RUN_LOG%"
echo [INFO] Transcript: "%TRANSCRIPT_LOG%"

REM ======= Locate Final Script in repo root =======
set "FINAL_TXT=%TARGET_DIR%\Final PS Script.txt"
set "PBIT_FILE=%TARGET_DIR%\Power BI Governance Model.pbit"

REM ======= Run PowerShell script inline =======
if exist "%FINAL_TXT%" (
  echo [INFO] Running Final Script: "%FINAL_TXT%"
  echo [INFO] Please wait for the PowerShell script to complete...
  cd /d "%TARGET_DIR%"
  
  REM Execute PowerShell script content inline with transcript + stream logging
  powershell -ExecutionPolicy Bypass -NoProfile -Command "& { $ErrorActionPreference='Stop'; Set-Location '%TARGET_DIR%'; try { Start-Transcript -Path '%TRANSCRIPT_LOG%' -Force | Out-Null } catch { Write-Warning ('Failed to start transcript: ' + $_.Exception.Message) }; try { $content = Get-Content '%FINAL_TXT%' -Raw; Invoke-Expression $content } catch { Write-Error $_; exit 1 } finally { try { Stop-Transcript | Out-Null } catch {} } } *>&1 | Tee-Object -FilePath '%RUN_LOG%'"

  set "PS_EXIT=%ERRORLEVEL%"
  if not "!PS_EXIT!"=="0" (
    echo [ERROR] PowerShell script exited with code !PS_EXIT!.
    echo [ERROR] Check logs:
    echo [ERROR]   %RUN_LOG%
    echo [ERROR]   %TRANSCRIPT_LOG%
    pause
    exit /b !PS_EXIT!
  )
  
  echo [INFO] PowerShell script finished.
) else (
  echo [ERROR] Final PS Script not found: "%FINAL_TXT%"
  echo [ERROR] Please ensure the Impact IQ files are installed.
  pause
  exit /b 1
)

REM ======= Open Power BI Template =======
if exist "%PBIT_FILE%" (
  echo [INFO] Opening Power BI template: "%PBIT_FILE%"
  start "" "%PBIT_FILE%"
) else (
  echo [WARNING] Power BI template file not found: "%PBIT_FILE%"
)

echo [INFO] Execution complete.
echo [INFO] Logs saved to:
echo [INFO]   %RUN_LOG%
echo [INFO]   %TRANSCRIPT_LOG%
exit /b 0