# =====================================================================================================================
#  ImpactIQ - interactive launcher ("Final PS Script")
# =====================================================================================================================
#
#  WHAT THIS IS
#    The one-click way to run ImpactIQ by hand, exactly like v2: save this file as "Final PS Script.ps1" next to
#    ImpactIQ.ps1 and run it (or paste the whole text into a PowerShell window opened in that folder). It starts
#    ImpactIQ.ps1 with the interactive defaults - the environment dialog, the browser sign-in and the run-mode /
#    workspace / report / model pickers all appear as before - and the four workbooks plus the Model / Report /
#    Dataflow Backups folders land in the folder ImpactIQ runs from.
#
#  WHERE THE DATA GOES
#    By default everything stays next to ImpactIQ.ps1 (the folder you downloaded the repository into). To keep the
#    backups or the workbooks somewhere else, set IMPACTIQ_BACKUP_FOLDER and/or IMPACTIQ_OUTPUT_FOLDER before
#    running this launcher, or add -BackupFolder / -OutputFolder to the ImpactIQ.ps1 call below. The old fixed
#    location C:\Power BI Backups is no longer assumed; it is only used when this text is pasted into a console
#    that is not in the ImpactIQ folder and ImpactIQ.ps1 still lives there.
#
#  WHAT CHANGED IN v3
#    The 3,700-line script moved into ImpactIQ.ps1 and Config\Modules\ImpactIQ.*.ps1. Every run is checkpointed
#    under State\runs\<date>\ - if a run is interrupted, simply start it again and it resumes where it stopped
#    (finished models / reports / dataflows are not downloaded twice). The console ends with a per-stage summary,
#    the four output paths and the log file (Logs\ImpactIQ_<timestamp>.log).
#
#  AUTOMATION (Azure DevOps, Task Scheduler, any headless host)
#    Do NOT use this launcher; call ImpactIQ.ps1 directly with -NonInteractive and the scope parameters, e.g.
#      powershell -NoProfile -ExecutionPolicy Bypass -File "C:\ImpactIQ\ImpactIQ.ps1" -NonInteractive `
#          -Environment USGov -AuthMode DeviceCode -AllWorkspaces -IncludeMyWorkspace
#    Exit codes: 0 = ok, 2 = finished with item failures (see the Failures sheet), 3 = paused (-TimeBudgetMinutes
#    reached; run it again to resume), 1 = fatal (see the log).
#    See README.md (Getting started), docs\Automation.md, docs\Auth-Options.md, docs\Azure-DevOps.md and
#    docs\Headless-and-Resume.md.
#
#  NOTE FOR PASTING
#    $PSScriptRoot is empty when the text is pasted, so the folder is resolved in this order: IMPACTIQ_BASE_FOLDER,
#    the current PowerShell location, C:\Power BI Backups (legacy) - whichever contains ImpactIQ.ps1. When the file
#    is saved as a .ps1 next to ImpactIQ.ps1, that folder always wins.
# =====================================================================================================================

# Allow the scripts to run in THIS process only (no machine-wide policy change, no admin rights needed).
try {
    if ((Get-ExecutionPolicy -Scope Process) -ne 'Bypass') {
        Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
    }
}
catch {
    Write-Host "[WARN] Could not set the process execution policy to Bypass ($($_.Exception.Message)). If ImpactIQ.ps1 refuses to run, start PowerShell with: powershell -ExecutionPolicy Bypass" -ForegroundColor Yellow
}

# Resolve the ImpactIQ folder: next to this file when it was saved as a .ps1, else IMPACTIQ_BASE_FOLDER, else the
# current PowerShell location, else the legacy C:\Power BI Backups - the first one that contains ImpactIQ.ps1 wins.
$impactIqBaseFolder = $null
$impactIqCandidates = New-Object System.Collections.Generic.List[string]
if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $impactIqCandidates.Add($PSScriptRoot) }
if (-not [string]::IsNullOrWhiteSpace($env:IMPACTIQ_BASE_FOLDER)) { $impactIqCandidates.Add($env:IMPACTIQ_BASE_FOLDER) }
$impactIqCandidates.Add((Get-Location).Path)
$impactIqCandidates.Add('C:\Power BI Backups')
foreach ($impactIqCandidate in $impactIqCandidates) {
    if (Test-Path -LiteralPath (Join-Path $impactIqCandidate 'ImpactIQ.ps1') -PathType Leaf) { $impactIqBaseFolder = $impactIqCandidate; break }
}

if ([string]::IsNullOrWhiteSpace($impactIqBaseFolder)) {
    Write-Host "[ERROR] ImpactIQ.ps1 was not found in any of: $($impactIqCandidates -join '; ')" -ForegroundColor Red
    Write-Host "        Download the repository (all files, including the Config folder) into one folder, open PowerShell in that folder (or set IMPACTIQ_BASE_FOLDER to it) and run this again. See README.md." -ForegroundColor Red
    return
}
$impactIqEntry = Join-Path $impactIqBaseFolder 'ImpactIQ.ps1'

Write-Host "[INFO] ImpactIQ - starting the interactive run from '$impactIqBaseFolder' ..." -ForegroundColor Cyan
Write-Host "[INFO] You will be asked for the Power BI environment, then to sign in, then what to run against." -ForegroundColor Cyan

$impactIqExitCode = 1
try {
    # -AuthMode Interactive = browser sign-in through the Power BI PowerShell module (installed for the current user if
    # missing). All other settings keep their defaults; add parameters here if you want them every time, e.g.
    # -Environment USGov, -IncludeMyWorkspace, -SkipToolUpdate, -ModelDetailMethod Dax.
    & $impactIqEntry -BaseFolder $impactIqBaseFolder -AuthMode Interactive
    $impactIqExitCode = $LASTEXITCODE
    if ($null -eq $impactIqExitCode) { $impactIqExitCode = 0 }
}
catch {
    $impactIqExitCode = 1
    Write-Host "[ERROR] ImpactIQ stopped: $($_.Exception.Message)" -ForegroundColor Red
}

switch ($impactIqExitCode) {
    0 {
        Write-Host ''
        Write-Host "[DONE] ImpactIQ finished. Open 'Power BI Governance Model.pbit' in '$impactIqBaseFolder', let it refresh and save it as .pbix." -ForegroundColor Green
    }
    2 {
        Write-Host ''
        Write-Host "[DONE] ImpactIQ finished with some item failures - the workbooks were still produced. Check the 'Failures' sheet in 'Power BI Environment Detail.xlsx' and the newest file in '$impactIqBaseFolder\Logs'. Running this launcher again retries only what failed." -ForegroundColor Yellow
    }
    3 {
        Write-Host ''
        Write-Host "[PAUSED] ImpactIQ stopped because its time budget was reached - partial workbooks were produced. Run this launcher again to resume; everything already finished is skipped." -ForegroundColor Yellow
    }
    default {
        Write-Host ''
        Write-Host "[FAILED] ImpactIQ did not complete (exit code $impactIqExitCode). See the newest log in '$impactIqBaseFolder\Logs'. Running this launcher again resumes the run." -ForegroundColor Red
    }
}
