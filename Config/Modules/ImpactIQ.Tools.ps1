# ImpactIQ.Tools.ps1 - external tool acquisition (Tabular Editor 2 portable, pbi-tools) and the process runner.
#
# Contract: brief section 2.5. Initialize-IQTools ports the monolith's bootstrap (Final PS Script.txt lines 34-160:
# download latest pbi-tools + TE2 portable, extract, TE2 preflight) with try/catch, -SkipToolUpdate / offline
# fallback and a version stamp under State\tools\ (FileVersion vs release tag decides, so a lost stamp never forces
# a re-download). Invoke-IQProcess / Invoke-IQProcessBatch / Invoke-IQTabularEditor give every external process a
# timeout, captured output under tool-logs\, a process-tree kill on timeout and cleanup of the pool on interruption.
# Tool folders are replaced only when no instance runs from them (no process is ever killed for an update); the
# WER DontShowUI setting is changed for headless runs only and restored (Restore-IQWerSetting / PowerShell.Exiting).
#
# Windows PowerShell 5.1 and PowerShell 7 compatible. Loaded by dot-sourcing, so $script:IQ is the shared context.
# Cross-module functions used (brief section 2): Write-IQLog, Get-IQSafeKey.
# Tabular Editor / pbi-tools only run on Windows ($script:IQ.IsWindows); everything else works on Linux.

function Get-IQRedactedArgumentList {
    <#
    .SYNOPSIS
    Masks secrets (Password=..., bearer tokens) in a command-line string before it is logged (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$ArgumentList
    )
    if ([string]::IsNullOrEmpty($ArgumentList)) { return '' }
    $redacted = $ArgumentList -replace '(?i)(Password=)[^;"]+', '$1***'
    $redacted = $redacted -replace '(?i)(Bearer\s+)[A-Za-z0-9\-_\.]+', '$1***'
    $redacted = $redacted -replace '(?i)(access_token=)[^&"\s]+', '$1***'
    return $redacted
}

function Get-IQToolLogFolder {
    <#
    .SYNOPSIS
    Returns (and creates) the tool-logs folder for a stage: <RunPath>\tool-logs\<Stage> or <LogsPath>\tool-logs\<Stage> (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$Stage = 'Tools'
    )
    $root = $null
    if ($script:IQ) {
        if ($script:IQ.ContainsKey('RunPath') -and -not [string]::IsNullOrWhiteSpace([string]$script:IQ.RunPath)) { $root = [string]$script:IQ.RunPath }
        elseif ($script:IQ.ContainsKey('LogsPath') -and -not [string]::IsNullOrWhiteSpace([string]$script:IQ.LogsPath)) { $root = [string]$script:IQ.LogsPath }
        elseif ($script:IQ.ContainsKey('BaseFolder') -and -not [string]::IsNullOrWhiteSpace([string]$script:IQ.BaseFolder)) { $root = Join-Path $script:IQ.BaseFolder 'Logs' }
    }
    if (-not $root) { $root = Join-Path ([System.IO.Path]::GetTempPath()) 'ImpactIQ' }
    $safeStage = 'Tools'
    if (-not [string]::IsNullOrWhiteSpace($Stage)) { $safeStage = Get-IQSafeKey -Value $Stage }
    $folder = Join-Path (Join-Path $root 'tool-logs') $safeStage
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    return $folder
}

function Get-IQToolTempFolder {
    <#
    .SYNOPSIS
    Returns (and creates) the short scratch folder used for temp scripts and extraction: <BaseFolder>\Config\Temp (private).
    #>
    [CmdletBinding()]
    param()
    $folder = $null
    if ($script:IQ -and $script:IQ.ContainsKey('ConfigFolder') -and -not [string]::IsNullOrWhiteSpace([string]$script:IQ.ConfigFolder)) {
        $folder = Join-Path $script:IQ.ConfigFolder 'Temp'
    }
    else {
        $folder = Join-Path ([System.IO.Path]::GetTempPath()) 'ImpactIQ'
    }
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    return $folder
}

function Stop-IQProcessTree {
    <#
    .SYNOPSIS
    Kills a process and its children: taskkill /T /F on Windows, Process.Kill(true)/Kill() elsewhere (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Diagnostics.Process]$Process
    )
    $processId = $null
    try { $processId = $Process.Id } catch { $processId = $null }
    $onWindows = $false
    if ($script:IQ -and $script:IQ.ContainsKey('IsWindows')) { $onWindows = [bool]$script:IQ.IsWindows } else { $onWindows = ($env:OS -eq 'Windows_NT') }
    if ($onWindows -and $processId) {
        try {
            $taskkill = Join-Path $env:SystemRoot 'System32\taskkill.exe'
            if (-not (Test-Path -LiteralPath $taskkill)) { $taskkill = 'taskkill.exe' }
            & $taskkill /PID $processId /T /F 2>&1 | Out-Null
        }
        catch {
            Write-IQLog -Level Debug -Stage Tools -Message ("taskkill failed for PID {0}: {1}" -f $processId, $_.Exception.Message)
        }
    }
    try {
        if (-not $Process.HasExited) {
            try { $Process.Kill($true) }     # PowerShell 7 / .NET Core: kill the whole tree
            catch { $Process.Kill() }        # .NET Framework has no Kill(bool) overload
        }
    }
    catch {
        Write-IQLog -Level Debug -Stage Tools -Message ("Kill failed for PID {0}: {1}" -f $processId, $_.Exception.Message)
    }
    try { $Process.WaitForExit(5000) | Out-Null } catch { Write-IQLog -Level Debug -Stage Tools -Message "WaitForExit after kill failed for PID $processId" }
}

function Start-IQProcessJob {
    <#
    .SYNOPSIS
    Starts one external process with redirected output files and returns its tracking record (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$ArgumentList,
        [Parameter(Mandatory = $false)]
        [string]$WorkingDirectory,
        [Parameter(Mandatory = $false)]
        [int]$TimeoutMinutes = 20,
        [Parameter(Mandatory = $false)]
        [string]$LogName,
        [Parameter(Mandatory = $false)]
        [string]$Stage = 'Tools',
        [Parameter(Mandatory = $false)]
        [string]$ItemKey,
        [Parameter(Mandatory = $false)]
        [string]$Item
    )
    if ($TimeoutMinutes -lt 1) { $TimeoutMinutes = 1 }
    if ([string]::IsNullOrWhiteSpace($LogName)) {
        $LogName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath) + '_' + (Get-Date -Format 'yyyyMMdd_HHmmssfff')
    }
    $job = @{
        ItemKey = $ItemKey; Item = $Item; FilePath = $FilePath; ArgumentList = $ArgumentList; WorkingDirectory = $WorkingDirectory
        LogName = $LogName; Stage = $Stage; OutFile = $null; ErrFile = $null; TimeoutMs = [int64]$TimeoutMinutes * 60000
        Process = $null; Started = $null; StartError = $null; TimedOut = $false; Completed = $false; Result = $null; ExitCodeUnknown = $false
    }
    $redacted = Get-IQRedactedArgumentList -ArgumentList $ArgumentList
    Write-IQLog -Level Debug -Stage $Stage -Item $Item -Message ("Starting: {0} {1}" -f $FilePath, $redacted)
    try {
        # The log-folder preparation lives inside the try so a disk-full / permission / path-length problem becomes a
        # StartError result (the pool keeps running) instead of an exception thrown out of Invoke-IQProcessBatch.
        $logFolder = Get-IQToolLogFolder -Stage $Stage
        $safeName = Get-IQSafeKey -Value $LogName
        $job.OutFile = Join-Path $logFolder ($safeName + '.out.txt')
        $job.ErrFile = Join-Path $logFolder ($safeName + '.err.txt')
        foreach ($file in @($job.OutFile, $job.ErrFile)) { if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file -Force -ErrorAction Stop } }
        if (-not (Test-Path -LiteralPath $FilePath)) { throw "Executable not found: $FilePath" }
        $startParams = @{ FilePath = $FilePath; PassThru = $true; NoNewWindow = $true; RedirectStandardOutput = $job.OutFile; RedirectStandardError = $job.ErrFile; ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace($ArgumentList)) { $startParams.ArgumentList = $ArgumentList }
        if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) {
            if (-not (Test-Path -LiteralPath $WorkingDirectory)) { New-Item -Path $WorkingDirectory -ItemType Directory -Force | Out-Null }
            $startParams.WorkingDirectory = $WorkingDirectory
        }
        $process = Start-Process @startParams
        # Touch the handle so ExitCode is available after exit on Windows PowerShell 5.1. When that fails the exit code
        # will be $null later; say so in the run log instead of silently reporting the process as successful.
        try { $null = $process.Handle }
        catch {
            $processId = '?'
            try { $processId = [string]$process.Id } catch { $processId = '?' }
            Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message ("Could not cache the process handle of {0} (PID {1}): {2}. Its exit code may be unavailable and the job will then be treated as failed." -f [System.IO.Path]::GetFileName($FilePath), $processId, $_.Exception.Message)
        }
        $job.Process = $process
        $job.Started = [datetime]::UtcNow
    }
    catch {
        $job.StartError = $_.Exception.Message
        Write-IQLog -Level Error -Stage $Stage -Item $Item -Message ("Could not start {0}: {1}" -f $FilePath, $_.Exception.Message)
    }
    return $job
}

function Complete-IQProcessJob {
    <#
    .SYNOPSIS
    Reads the redirected output files of a finished (or failed-to-start) job and builds the result hashtable (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Job
    )
    $exitCode = -1
    $duration = 0
    $exitUnknown = $false
    if ($Job.Process) {
        # Start-Process -PassThru on Windows PowerShell 5.1 leaves ExitCode $null when the handle was not cached before
        # the process exited; [int]$null is 0, which would report an unknown (possibly failed) exit as success.
        $rawExit = $null
        try { $rawExit = $Job.Process.ExitCode } catch { $rawExit = $null }
        if ($null -eq $rawExit) {
            $exitCode = -1
            if (-not $Job.TimedOut) {
                $exitUnknown = $true
                $note = 'exit code unavailable (process handle not cached)'
                if ($Job.StartError) { $Job.StartError = [string]$Job.StartError + '; ' + $note } else { $Job.StartError = $note }
            }
        }
        else {
            try { $exitCode = [int]$rawExit } catch { $exitCode = -1 }
        }
        if ($Job.Started) { $duration = [math]::Round(([datetime]::UtcNow - $Job.Started).TotalSeconds, 1) }
        try { $Job.Process.Dispose() } catch { $null = $null }
    }
    if ($Job.TimedOut) { $exitCode = -1 }
    $Job.ExitCodeUnknown = $exitUnknown
    $stdOut = ''
    $stdErr = ''
    try { if ($Job.OutFile -and (Test-Path -LiteralPath $Job.OutFile)) { $stdOut = [string](Get-Content -LiteralPath $Job.OutFile -Raw -ErrorAction SilentlyContinue) } } catch { $stdOut = '' }
    try { if ($Job.ErrFile -and (Test-Path -LiteralPath $Job.ErrFile)) { $stdErr = [string](Get-Content -LiteralPath $Job.ErrFile -Raw -ErrorAction SilentlyContinue) } } catch { $stdErr = '' }
    if ($null -eq $stdOut) { $stdOut = '' }
    if ($null -eq $stdErr) { $stdErr = '' }
    if ($Job.StartError) {
        if ($stdErr) { $stdErr = $stdErr.TrimEnd() + [Environment]::NewLine }
        $stdErr += 'Start failure: ' + $Job.StartError
        try { if ($Job.ErrFile) { Set-Content -LiteralPath $Job.ErrFile -Value $stdErr -Encoding UTF8 -ErrorAction SilentlyContinue } } catch { $null = $null }
    }
    $result = @{
        ExitCode = $exitCode; TimedOut = [bool]$Job.TimedOut; StdOut = $stdOut; StdErr = $stdErr
        OutFile = $Job.OutFile; ErrFile = $Job.ErrFile; DurationSec = $duration; StartError = $Job.StartError; FilePath = $Job.FilePath
    }
    $Job.Result = $result
    $Job.Completed = $true
    $name = [System.IO.Path]::GetFileName($Job.FilePath)
    if ($Job.TimedOut) {
        Write-IQLog -Level Warn -Stage $Job.Stage -Item $Job.Item -Message ("{0} timed out after {1} min and was killed (log: {2})" -f $name, [int]($Job.TimeoutMs / 60000), $Job.OutFile)
    }
    elseif ($exitUnknown) {
        Write-IQLog -Level Warn -Stage $Job.Stage -Item $Job.Item -Message ("{0} finished after {1} s but its exit code is unavailable (process handle not cached); treating it as failed (log: {2})" -f $name, $duration, $Job.OutFile)
    }
    elseif ($Job.StartError) {
        # already logged as Error in Start-IQProcessJob
    }
    elseif ($exitCode -ne 0) {
        $firstError = ''
        if ($stdErr) { $firstError = (($stdErr -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 1) }
        Write-IQLog -Level Warn -Stage $Job.Stage -Item $Job.Item -Message ("{0} exited with code {1} after {2} s. {3} (log: {4})" -f $name, $exitCode, $duration, $firstError, $Job.ErrFile)
    }
    else {
        Write-IQLog -Level Debug -Stage $Job.Stage -Item $Job.Item -Message ("{0} exited with code 0 after {1} s" -f $name, $duration)
    }
    return $result
}

function Invoke-IQProcess {
    <#
    .SYNOPSIS
    Runs an external process with a timeout and captured stdout/stderr files; never throws.
    .DESCRIPTION
    Uses Start-Process -PassThru -NoNewWindow with -RedirectStandardOutput/-RedirectStandardError into
    tool-logs\<Stage>\<LogName>.out.txt/.err.txt, waits with WaitForExit(ms), kills the process tree on timeout
    (taskkill /T /F on Windows) and reads the files back. Returns
    @{ ExitCode; TimedOut; StdOut; StdErr; OutFile; ErrFile; DurationSec; StartError }.
    Anything matching Password=[^;"]+ is redacted in the logged argument string.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$ArgumentList,
        [Parameter(Mandatory = $false)]
        [string]$WorkingDirectory,
        [Parameter(Mandatory = $false)]
        [int]$TimeoutMinutes = 20,
        [Parameter(Mandatory = $false)]
        [string]$LogName,
        [Parameter(Mandatory = $false)]
        [string]$Stage = 'Tools',
        [Parameter(Mandatory = $false)]
        [string]$Item
    )
    try {
        $job = Start-IQProcessJob -FilePath $FilePath -ArgumentList $ArgumentList -WorkingDirectory $WorkingDirectory -TimeoutMinutes $TimeoutMinutes -LogName $LogName -Stage $Stage -Item $Item
        if ($job.Process) {
            $exited = $false
            try { $exited = $job.Process.WaitForExit([int]$job.TimeoutMs) } catch { $exited = $job.Process.HasExited }
            if (-not $exited) {
                $job.TimedOut = $true
                Stop-IQProcessTree -Process $job.Process
            }
            else {
                # A second WaitForExit() without a timeout flushes the redirected streams before reading the files.
                try { $job.Process.WaitForExit() } catch { $null = $null }
            }
        }
        return (Complete-IQProcessJob -Job $job)
    }
    catch {
        Write-IQLog -Level Error -Stage $Stage -Item $Item -Message ("Invoke-IQProcess failed for {0}: {1}" -f $FilePath, $_.Exception.Message)
        return @{ ExitCode = -1; TimedOut = $false; StdOut = ''; StdErr = $_.Exception.Message; OutFile = $null; ErrFile = $null; DurationSec = 0; StartError = $_.Exception.Message; FilePath = $FilePath }
    }
}

function Invoke-IQProcessBatch {
    <#
    .SYNOPSIS
    Runs a list of process jobs through a simple pool (start up to MaxParallel, poll HasExited every 500 ms).
    .DESCRIPTION
    -Jobs is an array of @{ ItemKey; Item; FilePath; ArgumentList; WorkingDirectory; LogName }. Returns an array of
    @{ ItemKey; Item; Result = <Invoke-IQProcess result> } in job order. Each job gets the same timeout; a timed-out
    process tree is killed. -OnJobComplete (optional scriptblock, receives the entry as its first argument) lets the
    caller checkpoint the moment a job finishes; it runs in this function's dynamic scope, so the internal variables
    are prefixed 'iq' to keep out of the caller's way. Progress is logged every 10 completed jobs. Never throws for
    per-job failures.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [array]$Jobs,
        [Parameter(Mandatory = $false)]
        [int]$MaxParallel = 2,
        [Parameter(Mandatory = $false)]
        [int]$TimeoutMinutes = 20,
        [Parameter(Mandatory = $false)]
        [string]$Stage = 'Tools',
        [Parameter(Mandatory = $false)]
        [scriptblock]$OnJobComplete
    )
    if ($MaxParallel -lt 1) { $MaxParallel = 1 }
    $iqJobList = @($Jobs | Where-Object { $null -ne $_ })
    $iqTotal = $iqJobList.Count
    $iqEntries = New-Object System.Collections.Generic.List[object]
    if ($iqTotal -eq 0) { return @() }
    Write-IQLog -Level Info -Stage $Stage -Message "Running $iqTotal process job(s) with up to $MaxParallel in parallel (timeout $TimeoutMinutes min each)"

    $iqRunning = New-Object System.Collections.Generic.List[object]
    $iqNextIndex = 0
    $iqCompleted = 0
    $iqFailed = 0

    $iqFinish = {
        param($iqEntry)
        $iqEntry.Result = Complete-IQProcessJob -Job $iqEntry.Job
        $iqEntry.Remove('Job')
    }

    try {
        while ($iqNextIndex -lt $iqTotal -or $iqRunning.Count -gt 0) {
            # Fill the pool.
            while ($iqRunning.Count -lt $MaxParallel -and $iqNextIndex -lt $iqTotal) {
                $iqSpec = $iqJobList[$iqNextIndex]
                $iqNextIndex++
                $iqItemKey = [string]$iqSpec.ItemKey
                $iqItem = [string]$iqSpec.Item
                $iqLogName = [string]$iqSpec.LogName
                if ([string]::IsNullOrWhiteSpace($iqLogName)) { $iqLogName = $iqItemKey }
                $iqJob = Start-IQProcessJob -FilePath ([string]$iqSpec.FilePath) -ArgumentList ([string]$iqSpec.ArgumentList) -WorkingDirectory ([string]$iqSpec.WorkingDirectory) -TimeoutMinutes $TimeoutMinutes -LogName $iqLogName -Stage $Stage -ItemKey $iqItemKey -Item $iqItem
                $iqEntry = @{ ItemKey = $iqItemKey; Item = $iqItem; Index = $iqNextIndex - 1; Result = $null; Job = $iqJob }
                $iqEntries.Add($iqEntry)
                if ($iqJob.Process) {
                    $iqRunning.Add($iqEntry)
                }
                else {
                    & $iqFinish $iqEntry
                    $iqCompleted++
                    $iqFailed++
                    if ($OnJobComplete) { try { & $OnJobComplete $iqEntry } catch { Write-IQLog -Level Warn -Stage $Stage -Item $iqItem -Message ("OnJobComplete failed: " + $_.Exception.Message) } }
                }
            }
            if ($iqRunning.Count -eq 0) { continue }

            Start-Sleep -Milliseconds 500
            $iqNow = [datetime]::UtcNow
            $iqDone = @()
            foreach ($iqEntry in $iqRunning) {
                $iqJob = $iqEntry.Job
                $iqHasExited = $false
                try { $iqHasExited = $iqJob.Process.HasExited } catch { $iqHasExited = $true }
                if ($iqHasExited) {
                    try { $iqJob.Process.WaitForExit() } catch { $null = $null }
                    $iqDone += $iqEntry
                }
                elseif (($iqNow - $iqJob.Started).TotalMilliseconds -gt $iqJob.TimeoutMs) {
                    $iqJob.TimedOut = $true
                    Stop-IQProcessTree -Process $iqJob.Process
                    $iqDone += $iqEntry
                }
            }
            foreach ($iqEntry in $iqDone) {
                [void]$iqRunning.Remove($iqEntry)
                & $iqFinish $iqEntry
                $iqCompleted++
                if ($iqEntry.Result.TimedOut -or $iqEntry.Result.ExitCode -ne 0) { $iqFailed++ }
                if ($OnJobComplete) { try { & $OnJobComplete $iqEntry } catch { Write-IQLog -Level Warn -Stage $Stage -Item $iqEntry.Item -Message ("OnJobComplete failed: " + $_.Exception.Message) } }
                if (($iqCompleted % 10) -eq 0 -or $iqCompleted -eq $iqTotal) {
                    Write-IQLog -Level Info -Stage $Stage -Message ("Process batch progress: {0}/{1} done, {2} failed, {3} running" -f $iqCompleted, $iqTotal, $iqFailed, $iqRunning.Count)
                }
            }
        }
    }
    finally {
        # Anything still in the pool here was left behind by an exception (or a Ctrl+C / pipeline cancel). Without this
        # the orphaned Tabular Editor processes would keep writing their outputs (with no timeout) after the stage has
        # recorded a failure and while a resumed run starts new instances for the same items.
        foreach ($iqEntry in $iqRunning.ToArray()) {
            try {
                if ($iqEntry.Job.Process -and -not $iqEntry.Job.Process.HasExited) {
                    Write-IQLog -Level Warn -Stage $Stage -Item $iqEntry.Item -Message ("Process batch interrupted; killing the running {0} (PID {1})" -f [System.IO.Path]::GetFileName($iqEntry.Job.FilePath), $iqEntry.Job.Process.Id)
                    $iqEntry.Job.TimedOut = $true
                    Stop-IQProcessTree -Process $iqEntry.Job.Process
                }
            }
            catch { Write-IQLog -Level Debug -Stage $Stage -Item $iqEntry.Item -Message ("Cleanup of an interrupted process job failed: " + $_.Exception.Message) }
            if ($iqEntry.ContainsKey('Job') -and $iqEntry.Job -and -not $iqEntry.Job.Completed) {
                try { & $iqFinish $iqEntry } catch { Write-IQLog -Level Debug -Stage $Stage -Item $iqEntry.Item -Message ("Could not finalise an interrupted process job: " + $_.Exception.Message) }
            }
        }
        $iqRunning.Clear()
    }
    # $iqEntries was filled in start order, which is job order.
    $iqResults = @()
    foreach ($iqEntry in $iqEntries) {
        $iqResults += @{ ItemKey = $iqEntry.ItemKey; Item = $iqEntry.Item; Result = $iqEntry.Result }
    }
    return , $iqResults
}

function Select-IQTabularEditorError {
    <#
    .SYNOPSIS
    Returns the lines of Tabular Editor output that look like script/compile errors (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Output
    )
    if ([string]::IsNullOrWhiteSpace($Output)) { return @() }
    $lines = @($Output -split "`r?`n")
    $errors = @()
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*(Error\b|Script compilation error|Script error|Unhandled exception|Exception:)') {
            $errors += $lines[$i]
            if ($i + 1 -lt $lines.Count -and $lines[$i + 1].Trim()) { $errors += $lines[$i + 1] }
        }
    }
    return @($errors)
}

function Invoke-IQTabularEditor {
    <#
    .SYNOPSIS
    Runs Tabular Editor 2 (portable CLI) via Invoke-IQProcess and classifies the outcome; never throws.
    .DESCRIPTION
    Uses $script:IQ.Tools.TabularEditorPath with WorkingDirectory = BaseFolder (override with -WorkingDirectory).
    Returns the Invoke-IQProcess result plus Success (bool), ErrorLines (string[]) and FailureReason.
    Failure = ExitCode -ne 0, TimedOut, or stdout containing lines starting with "Error" (script errors); the first
    20 stdout lines are logged at Warn on failure.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$ArgumentList,
        [Parameter(Mandatory = $false)]
        [string]$WorkingDirectory,
        [Parameter(Mandatory = $false)]
        [int]$TimeoutMinutes = 20,
        [Parameter(Mandatory = $false)]
        [string]$LogName = 'tabular-editor',
        [Parameter(Mandatory = $false)]
        [string]$Stage = 'Tools',
        [Parameter(Mandatory = $false)]
        [string]$Item
    )
    $tePath = $null
    if ($script:IQ -and $script:IQ.Tools) { $tePath = [string]$script:IQ.Tools.TabularEditorPath }
    if ([string]::IsNullOrWhiteSpace($WorkingDirectory) -and $script:IQ) { $WorkingDirectory = [string]$script:IQ.BaseFolder }
    if ([string]::IsNullOrWhiteSpace($tePath) -or -not (Test-Path -LiteralPath $tePath)) {
        Write-IQLog -Level Error -Stage $Stage -Item $Item -Message "Tabular Editor is not available (path: '$tePath'); cannot run: $(Get-IQRedactedArgumentList -ArgumentList $ArgumentList)"
        return @{ ExitCode = -1; TimedOut = $false; StdOut = ''; StdErr = 'Tabular Editor not available'; OutFile = $null; ErrFile = $null; DurationSec = 0; Success = $false; ErrorLines = @('Tabular Editor not available'); FailureReason = 'missing-exe' }
    }
    if ($TimeoutMinutes -lt 1) { $TimeoutMinutes = 1 }
    $result = Invoke-IQProcess -FilePath $tePath -ArgumentList $ArgumentList -WorkingDirectory $WorkingDirectory -TimeoutMinutes $TimeoutMinutes -LogName $LogName -Stage $Stage -Item $Item
    $errorLines = @(Select-IQTabularEditorError -Output $result.StdOut)
    if ($result.StdErr -and $result.StdErr.Trim()) { $errorLines += @(($result.StdErr -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 5) }
    $reason = $null
    if ($result.TimedOut) { $reason = 'timeout' }
    elseif ($result.StartError) { $reason = 'start-failure' }
    elseif ($result.ExitCode -ne 0) { $reason = "exit $($result.ExitCode)" }
    elseif (@(Select-IQTabularEditorError -Output $result.StdOut).Count -gt 0) { $reason = 'script-error' }
    $result.Success = ($null -eq $reason)
    $result.ErrorLines = @($errorLines)
    $result.FailureReason = $reason
    if (-not $result.Success) {
        $head = @(($result.StdOut -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 20)
        $summary = ''
        if ($head.Count -gt 0) { $summary = ' Output: ' + ($head -join ' | ') }
        Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message ("Tabular Editor failed ({0}) after {1} s.{2}" -f $reason, $result.DurationSec, $summary)
    }
    return $result
}

function Get-IQTagFromRedirect {
    <#
    .SYNOPSIS
    Extracts the release tag from a GitHub ".../releases/latest" redirect Location (".../releases/tag/<tag>"); $null otherwise (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Location
    )
    if ([string]::IsNullOrWhiteSpace($Location)) { return $null }
    $trimmed = $Location.Trim().TrimEnd('/')
    if ($trimmed -notmatch '(?i)/releases/tag/[^/]+$') { return $null }
    $tag = ($trimmed -split '/')[-1]
    try { $tag = [uri]::UnescapeDataString($tag) } catch { $null = $null }
    if ([string]::IsNullOrWhiteSpace($tag)) { return $null }
    return $tag
}

function Get-IQGitHubLatestTag {
    <#
    .SYNOPSIS
    Resolves the latest release tag of a GitHub repo through a rate-limit-free HEAD redirect probe; $null on failure (private).
    .DESCRIPTION
    .NET Framework hands the 302 back from GetResponse(); the .NET Core HttpWebRequest shim (PowerShell 7) throws a
    WebException for it whose Response still carries the Location header, so the redirect is read from both places.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Repo
    )
    try {
        $request = [System.Net.HttpWebRequest]::Create("https://github.com/$Repo/releases/latest")
        $request.Method = 'HEAD'
        $request.AllowAutoRedirect = $false
        $request.Timeout = 30000
        $request.UserAgent = 'ImpactIQ/3.0'
        $response = $request.GetResponse()
        try {
            $location = [string]$response.Headers['Location']
        }
        finally { $response.Close() }
        return (Get-IQTagFromRedirect -Location $location)
    }
    catch {
        $location = $null
        try {
            $resp = $null
            if ($_.Exception.PSObject.Properties['Response']) { $resp = $_.Exception.Response }
            if (-not $resp -and $_.Exception.InnerException -and $_.Exception.InnerException.PSObject.Properties['Response']) { $resp = $_.Exception.InnerException.Response }
            if ($resp -and $resp.Headers) { $location = [string]$resp.Headers['Location'] }
            if ($resp) { try { $resp.Close() } catch { $null = $null } }
        }
        catch { $location = $null }
        $tag = Get-IQTagFromRedirect -Location $location
        if ($tag) { return $tag }
        Write-IQLog -Level Debug -Stage Tools -Message ("Latest-tag probe failed for {0}: {1}" -f $Repo, $_.Exception.Message)
    }
    return $null
}

function Test-IQToolVersionCurrent {
    <#
    .SYNOPSIS
    True when the installed tool matches the latest release tag: the stamp recorded that tag, or the exe FileVersion starts with the tag's numeric parts (private).
    .DESCRIPTION
    Compares version segments ("2.26.0" matches FileVersion "2.26.0.0", not "2.260.0"); a leading "v" and any
    pre-release suffix are ignored. With an empty tag nothing can be compared and $false is returned.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$InstalledVersion,
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Tag,
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$StampTag
    )
    if ([string]::IsNullOrWhiteSpace($Tag)) { return $false }
    if (-not [string]::IsNullOrWhiteSpace($StampTag) -and ($StampTag.Trim() -eq $Tag.Trim())) { return $true }
    if ([string]::IsNullOrWhiteSpace($InstalledVersion)) { return $false }
    $tagNumeric = ($Tag.Trim() -replace '^[vV]', '') -replace '[^0-9\.].*$', ''
    $verNumeric = ($InstalledVersion.Trim() -replace '^[vV]', '') -replace '[^0-9\.].*$', ''
    $tagParts = @($tagNumeric -split '\.' | Where-Object { $_ -ne '' })
    $verParts = @($verNumeric -split '\.' | Where-Object { $_ -ne '' })
    if ($tagParts.Count -eq 0 -or $verParts.Count -lt $tagParts.Count) { return $false }
    for ($i = 0; $i -lt $tagParts.Count; $i++) {
        if ([int64]$tagParts[$i] -ne [int64]$verParts[$i]) { return $false }
    }
    return $true
}

function Select-IQPbiToolsAsset {
    <#
    .SYNOPSIS
    Picks the pbi-tools Desktop-edition zip (pbi-tools.<version>.zip) from a GitHub release asset list; $null when absent (private).
    .DESCRIPTION
    Releases also ship pbi-tools.core.<version>.<rid>.zip (cross-platform build without the PBIX model extraction that
    Report Backup needs). The choice must not depend on the API's asset order, so the Desktop name pattern is matched
    explicitly and the .core. builds are excluded; nothing is installed when no Desktop asset exists.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [array]$Assets
    )
    foreach ($asset in @($Assets | Where-Object { $null -ne $_ })) {
        $name = ''
        try { $name = [string]$asset.name } catch { $name = '' }
        if ($name -match '^pbi-tools\.\d[\w\.\-]*\.zip$' -and $name -notmatch '(?i)\.core\.') { return $asset }
    }
    return $null
}

function Get-IQToolVersionStampPath {
    <#
    .SYNOPSIS
    Path of the tool version stamp: State\tools\tool-versions.json (created), so the impactiq-state pipeline artifact carries it (private).
    .DESCRIPTION
    Config\ is a fresh git checkout on every pipeline run, so a stamp there never survives; State\ is persisted.
    Falls back to Config\tool-versions.json when no State path is known.
    #>
    [CmdletBinding()]
    param()
    $root = $null
    if ($script:IQ -and $script:IQ.ContainsKey('StatePath') -and -not [string]::IsNullOrWhiteSpace([string]$script:IQ.StatePath)) { $root = Join-Path ([string]$script:IQ.StatePath) 'tools' }
    elseif ($script:IQ -and $script:IQ.ContainsKey('ConfigFolder') -and -not [string]::IsNullOrWhiteSpace([string]$script:IQ.ConfigFolder)) { $root = [string]$script:IQ.ConfigFolder }
    else { $root = Join-Path ([System.IO.Path]::GetTempPath()) 'ImpactIQ' }
    try { if (-not (Test-Path -LiteralPath $root)) { New-Item -Path $root -ItemType Directory -Force | Out-Null } } catch { $null = $null }
    return (Join-Path $root 'tool-versions.json')
}

function Get-IQToolVersionStamp {
    <#
    .SYNOPSIS
    Reads the tool-versions.json stamp ({ pbiTools; tabularEditor }) or returns an empty hashtable (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    $stamp = @{ pbiTools = $null; tabularEditor = $null }
    try {
        if (Test-Path -LiteralPath $Path) {
            $json = ConvertFrom-Json -InputObject ([string](Get-Content -LiteralPath $Path -Raw))
            if ($json.PSObject.Properties['pbiTools']) { $stamp.pbiTools = [string]$json.pbiTools }
            if ($json.PSObject.Properties['tabularEditor']) { $stamp.tabularEditor = [string]$json.tabularEditor }
        }
    }
    catch { Write-IQLog -Level Debug -Stage Tools -Message ("tool-versions.json unreadable: " + $_.Exception.Message) }
    return $stamp
}

function Save-IQToolVersionStamp {
    <#
    .SYNOPSIS
    Writes the tool-versions.json stamp (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [hashtable]$Stamp
    )
    try {
        $object = [pscustomobject]@{ pbiTools = $Stamp.pbiTools; tabularEditor = $Stamp.tabularEditor; updatedUtc = [datetime]::UtcNow.ToString('o') }
        $json = ConvertTo-Json -InputObject $object -Depth 20
        [System.IO.File]::WriteAllText($Path, $json, (New-Object System.Text.UTF8Encoding($false)))
    }
    catch { Write-IQLog -Level Debug -Stage Tools -Message ("Could not write tool-versions.json: " + $_.Exception.Message) }
}

function Invoke-IQToolDownload {
    <#
    .SYNOPSIS
    Downloads a URL to <ZipPath> via a .partial file and verifies it is a non-empty zip; returns $true/$false (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url,
        [Parameter(Mandatory = $true)]
        [string]$ZipPath
    )
    $partial = $ZipPath + '.partial'
    $ProgressPreference = 'SilentlyContinue'   # Audit C1-07
    try {
        if (Test-Path -LiteralPath $partial) { Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue }
        Invoke-WebRequest -Uri $Url -OutFile $partial -UseBasicParsing -TimeoutSec 600 -UserAgent 'ImpactIQ/3.0' -ErrorAction Stop
        if (-not (Test-Path -LiteralPath $partial) -or (Get-Item -LiteralPath $partial).Length -lt 1024) { throw "downloaded file is missing or too small" }
        # Verify the zip opens before replacing the previous download (C1-08).
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        $archive = [System.IO.Compression.ZipFile]::OpenRead($partial)
        try { if ($archive.Entries.Count -lt 1) { throw "zip has no entries" } } finally { $archive.Dispose() }
        Move-Item -LiteralPath $partial -Destination $ZipPath -Force
        return $true
    }
    catch {
        Write-IQLog -Level Warn -Stage Tools -Message ("Download failed for {0}: {1}" -f $Url, $_.Exception.Message)
        if (Test-Path -LiteralPath $partial) { Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue }
        return $false
    }
}

function Get-IQToolLockingProcess {
    <#
    .SYNOPSIS
    Returns the running processes whose executable IS the given exe path (same file, not merely the same name) (private).
    .DESCRIPTION
    Used before a tool folder is replaced. Only instances started from that folder lock its files; a process with the
    same name elsewhere (the user's interactive Tabular Editor 2 installation, another ImpactIQ checkout on the same
    agent) is unrelated and must be left alone. Processes whose path cannot be read (other users' sessions) are
    ignored. Returns an empty array when nothing matches.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath
    )
    $found = @()
    $name = [System.IO.Path]::GetFileNameWithoutExtension($ExePath)
    $full = $ExePath
    try { $full = [System.IO.Path]::GetFullPath($ExePath) } catch { $full = $ExePath }
    try {
        foreach ($process in @(Get-Process -Name $name -ErrorAction SilentlyContinue)) {
            if ($null -eq $process) { continue }
            $path = $null
            try { $path = [string]$process.Path } catch { $path = $null }
            if ([string]::IsNullOrWhiteSpace($path)) { continue }
            $same = $false
            try { $same = ([System.IO.Path]::GetFullPath($path) -ieq $full) } catch { $same = $false }
            if ($same) { $found += $process }
        }
    }
    catch { Write-IQLog -Level Debug -Stage Tools -Message ("Process scan for {0} failed: {1}" -f $name, $_.Exception.Message) }
    return @($found)
}

function Install-IQToolFromZip {
    <#
    .SYNOPSIS
    Extracts a tool zip into a temp folder, verifies the expected exe, swaps it into place and removes the zip (private).
    .DESCRIPTION
    Ports monolith lines 63-121 (extract to %TEMP%, then copy over the tool folder) with audit fix C1-09: the
    extracted folder replaces the old one (old renamed to .old and removed on success, restored on failure) so
    stale files do not accumulate. When an instance is running from the target folder the exe is locked: the
    existing copy is kept and reported (no process is killed - a same-named process may be the user's own Tabular
    Editor session or another ImpactIQ run). Returns $true when the tool folder was updated.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ZipPath,
        [Parameter(Mandatory = $true)]
        [string]$TargetFolder,
        [Parameter(Mandatory = $true)]
        [string]$ExeName,
        [Parameter(Mandatory = $true)]
        [string]$DisplayName
    )
    if (-not (Test-Path -LiteralPath $ZipPath)) { return $false }
    $tempExtractPath = Join-Path (Get-IQToolTempFolder) (($DisplayName -replace '[^A-Za-z0-9]', '') + '_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $oldFolder = $TargetFolder.TrimEnd('\', '/') + '.old'
    try {
        New-Item -Path $tempExtractPath -ItemType Directory -Force | Out-Null
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $tempExtractPath)
        # Some archives wrap everything in a single top-level folder.
        $exe = Get-ChildItem -LiteralPath $tempExtractPath -Filter $ExeName -Recurse -File -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $exe) { throw "$ExeName not found inside $ZipPath" }
        $sourceFolder = $exe.DirectoryName

        # C1-09: an exe that is running from the target folder cannot be replaced. Instances running from THIS folder
        # are the only ones that lock it; they are reported and the existing copy is kept (never killed: the same
        # process name may be the user's interactive Tabular Editor or another ImpactIQ run on this host).
        $locking = @(Get-IQToolLockingProcess -ExePath (Join-Path $TargetFolder $ExeName))
        if ($locking.Count -gt 0) {
            $processIds = @($locking | ForEach-Object { try { [string]$_.Id } catch { '?' } }) -join ', '
            Write-IQLog -Level Warn -Stage Tools -Message ("Could not install {0} from {1}: {2} is running from {3} (PID {4}). Keeping the existing copy; the update is retried on the next run." -f $DisplayName, $ZipPath, $ExeName, $TargetFolder, $processIds)
            return $false
        }

        if (Test-Path -LiteralPath $oldFolder) { Remove-Item -LiteralPath $oldFolder -Recurse -Force -ErrorAction SilentlyContinue }
        $hadExisting = Test-Path -LiteralPath $TargetFolder
        if ($hadExisting) { Rename-Item -LiteralPath $TargetFolder -NewName ([System.IO.Path]::GetFileName($oldFolder)) -Force -ErrorAction Stop }
        try {
            New-Item -Path $TargetFolder -ItemType Directory -Force | Out-Null
            # -LiteralPath enumeration: a BaseFolder containing [ or ] would otherwise be expanded as a wildcard.
            Get-ChildItem -LiteralPath $sourceFolder -Force -ErrorAction Stop | Copy-Item -Destination $TargetFolder -Recurse -Force -ErrorAction Stop
            if (-not (Test-Path -LiteralPath (Join-Path $TargetFolder $ExeName))) { throw "$ExeName missing after copy" }
        }
        catch {
            # Restore the previous folder.
            if (Test-Path -LiteralPath $TargetFolder) { Remove-Item -LiteralPath $TargetFolder -Recurse -Force -ErrorAction SilentlyContinue }
            if ($hadExisting -and (Test-Path -LiteralPath $oldFolder)) { Rename-Item -LiteralPath $oldFolder -NewName ([System.IO.Path]::GetFileName($TargetFolder)) -Force -ErrorAction SilentlyContinue }
            throw
        }
        if (Test-Path -LiteralPath $oldFolder) { Remove-Item -LiteralPath $oldFolder -Recurse -Force -ErrorAction SilentlyContinue }
        Remove-Item -LiteralPath $ZipPath -Force -ErrorAction SilentlyContinue
        Write-IQLog -Level Success -Stage Tools -Message "Updated $DisplayName at $TargetFolder"
        return $true
    }
    catch {
        Write-IQLog -Level Warn -Stage Tools -Message ("Could not install {0} from {1}: {2}. Using the existing copy if present." -f $DisplayName, $ZipPath, $_.Exception.Message)
        return $false
    }
    finally {
        if (Test-Path -LiteralPath $tempExtractPath) { Remove-Item -LiteralPath $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Get-IQFileVersion {
    <#
    .SYNOPSIS
    Returns the FileVersion of an executable or $null (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    try {
        if (Test-Path -LiteralPath $Path) {
            $version = (Get-Item -LiteralPath $Path).VersionInfo.FileVersion
            if ($version) { return [string]$version }
        }
    }
    catch { return $null }
    return $null
}

function Set-IQWerDontShowUI {
    <#
    .SYNOPSIS
    Sets HKCU WER DontShowUI=1 for the duration of a headless run, remembering the previous value (private).
    .DESCRIPTION
    Audit C1-10: a crashing Tabular Editor / pbi-tools must not block a headless agent with a Windows Error Reporting
    dialog. The previous value (or its absence) is stored in $script:IQ.Tools.WerPrevious / WerHadValue and put back
    by Restore-IQWerSetting; a PowerShell.Exiting engine event restores it as well when the caller never gets there.
    Interactive runs never touch the setting (Initialize-IQTools only calls this when $script:IQ.Interactive is false).
    #>
    [CmdletBinding()]
    param()
    $werKey = 'HKCU:\Software\Microsoft\Windows\Windows Error Reporting'
    try {
        $previous = $null
        $hadValue = $false
        if (Test-Path -LiteralPath $werKey) {
            $item = Get-ItemProperty -LiteralPath $werKey -Name 'DontShowUI' -ErrorAction SilentlyContinue
            if ($null -ne $item -and $item.PSObject.Properties['DontShowUI']) { $previous = $item.DontShowUI; $hadValue = $true }
        }
        else { New-Item -Path $werKey -Force | Out-Null }
        if ($hadValue -and ([int]$previous -eq 1)) {
            Write-IQLog -Level Debug -Stage Tools -Message 'WER DontShowUI is already 1; nothing to change.'
            return
        }
        New-ItemProperty -LiteralPath $werKey -Name 'DontShowUI' -Value 1 -PropertyType DWord -Force | Out-Null
        if ($script:IQ -and $script:IQ.Tools) {
            $script:IQ.Tools.WerPrevious = $previous
            $script:IQ.Tools.WerHadValue = $hadValue
            $script:IQ.Tools.WerChanged = $true
        }
        Write-IQLog -Level Debug -Stage Tools -Message ("WER DontShowUI set to 1 for this headless run (previous: {0}); it is restored at the end of the run." -f $(if ($hadValue) { [string]$previous } else { 'not set' }))
        if (-not $script:IQWerExitHookRegistered) {
            $data = @{ Previous = $previous; HadValue = $hadValue; Key = $werKey }
            Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -MessageData $data -Action {
                $d = $Event.MessageData
                try {
                    if ($d.HadValue) { Set-ItemProperty -LiteralPath $d.Key -Name 'DontShowUI' -Value ([int]$d.Previous) -ErrorAction SilentlyContinue }
                    else { Remove-ItemProperty -LiteralPath $d.Key -Name 'DontShowUI' -ErrorAction SilentlyContinue }
                }
                catch { $null = $null }
            } | Out-Null
            $script:IQWerExitHookRegistered = $true
        }
    }
    catch { Write-IQLog -Level Debug -Stage Tools -Message ("Could not set WER DontShowUI: " + $_.Exception.Message) }
}

function Restore-IQWerSetting {
    <#
    .SYNOPSIS
    Puts the HKCU WER DontShowUI value back to what it was before Set-IQWerDontShowUI changed it; no-op otherwise.
    #>
    [CmdletBinding()]
    param()
    if (-not ($script:IQ -and $script:IQ.Tools)) { return }
    $changed = $false
    try { if ($script:IQ.Tools.ContainsKey('WerChanged')) { $changed = [bool]$script:IQ.Tools.WerChanged } } catch { $changed = $false }
    if (-not $changed) { return }
    $werKey = 'HKCU:\Software\Microsoft\Windows\Windows Error Reporting'
    try {
        $hadValue = $false
        try { if ($script:IQ.Tools.ContainsKey('WerHadValue')) { $hadValue = [bool]$script:IQ.Tools.WerHadValue } } catch { $hadValue = $false }
        if ($hadValue) { Set-ItemProperty -LiteralPath $werKey -Name 'DontShowUI' -Value ([int]$script:IQ.Tools.WerPrevious) -ErrorAction Stop }
        else { Remove-ItemProperty -LiteralPath $werKey -Name 'DontShowUI' -ErrorAction Stop }
        $script:IQ.Tools.WerChanged = $false
        Write-IQLog -Level Debug -Stage Tools -Message 'WER DontShowUI restored to its previous value.'
    }
    catch { Write-IQLog -Level Debug -Stage Tools -Message ("Could not restore WER DontShowUI: " + $_.Exception.Message) }
}

function Get-IQPbiDesktopDetection {
    <#
    .SYNOPSIS
    Detects a Power BI Desktop installation: parsed "pbi-tools info" JSON first, then well-known paths and the registry (private).
    .DESCRIPTION
    Returns @{ Found = <bool>; Location = <string or $null>; Source = 'pbi-tools' | 'path' | 'registry' | $null; Edition = <pbi-tools edition or $null> }.
    The info JSON reports installs as pbiInstalls[].location / effectivePbiInstallDir, which need not contain the
    token "PBIDesktop", so the output is parsed instead of substring-matched (a false negative would make Report
    Backup disable PBIX model extraction for the run after a few unrelated failures). Path/registry probes cover the
    64-bit and 32-bit Program Files, the Store alias and HKLM\SOFTWARE\Microsoft\Microsoft Power BI Desktop.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$ProbeOutput
    )
    $result = @{ Found = $false; Location = $null; Source = $null; Edition = $null }
    if (-not [string]::IsNullOrWhiteSpace($ProbeOutput)) {
        $info = $null
        try { $info = ConvertFrom-Json -InputObject $ProbeOutput -ErrorAction Stop }
        catch {
            # Tolerate banner lines around the JSON document.
            $start = $ProbeOutput.IndexOf('{')
            $end = $ProbeOutput.LastIndexOf('}')
            if ($start -ge 0 -and $end -gt $start) {
                try { $info = ConvertFrom-Json -InputObject $ProbeOutput.Substring($start, $end - $start + 1) -ErrorAction Stop } catch { $info = $null }
            }
            if ($null -eq $info) { Write-IQLog -Level Debug -Stage Tools -Message ('pbi-tools info output is not JSON: ' + $_.Exception.Message) }
        }
        if ($null -ne $info) {
            try { if ($info.PSObject.Properties['edition'] -and $info.edition) { $result.Edition = [string]$info.edition } } catch { $null = $null }
            try {
                $installs = @()
                if ($info.PSObject.Properties['pbiInstalls'] -and $null -ne $info.pbiInstalls) { $installs = @($info.pbiInstalls | Where-Object { $null -ne $_ }) }
                if ($installs.Count -gt 0) {
                    $result.Found = $true
                    $result.Source = 'pbi-tools'
                    $first = $installs[0]
                    if ($first -is [string]) { $result.Location = $first }
                    elseif ($first.PSObject.Properties['location'] -and $first.location) { $result.Location = [string]$first.location }
                }
                elseif ($info.PSObject.Properties['effectivePbiInstallDir'] -and -not [string]::IsNullOrWhiteSpace([string]$info.effectivePbiInstallDir)) {
                    $result.Found = $true
                    $result.Source = 'pbi-tools'
                    $result.Location = [string]$info.effectivePbiInstallDir
                }
            }
            catch { Write-IQLog -Level Debug -Stage Tools -Message ('pbi-tools info JSON could not be evaluated: ' + $_.Exception.Message) }
        }
    }
    if ($result.Found) { return $result }

    $desktopPaths = @()
    $programFiles64 = $env:ProgramW6432
    if ([string]::IsNullOrWhiteSpace($programFiles64)) { $programFiles64 = $env:ProgramFiles }
    $programFiles86 = ${env:ProgramFiles(x86)}
    foreach ($root in @($programFiles64, $env:ProgramFiles, $programFiles86)) {
        if (-not [string]::IsNullOrWhiteSpace($root)) { $desktopPaths += (Join-Path $root 'Microsoft Power BI Desktop\bin\PBIDesktop.exe') }
    }
    if ($env:LOCALAPPDATA) { $desktopPaths += (Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\PBIDesktop.exe') }
    foreach ($candidate in @($desktopPaths | Select-Object -Unique)) {
        try {
            if (Test-Path -LiteralPath $candidate) {
                $result.Found = $true; $result.Source = 'path'; $result.Location = Split-Path -Path $candidate -Parent
                return $result
            }
        }
        catch { $null = $null }
    }

    $onWindows = ($env:OS -eq 'Windows_NT')
    if ($script:IQ -and $script:IQ.ContainsKey('IsWindows')) { $onWindows = [bool]$script:IQ.IsWindows }
    if ($onWindows) {
        foreach ($regKey in @('HKLM:\SOFTWARE\Microsoft\Microsoft Power BI Desktop', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Microsoft Power BI Desktop')) {
            try {
                if (-not (Test-Path -LiteralPath $regKey)) { continue }
                $item = Get-ItemProperty -LiteralPath $regKey -ErrorAction SilentlyContinue
                $location = $null
                if ($null -ne $item) {
                    foreach ($prop in @('InstallLocation', 'InstallDir', 'Path')) { if ($item.PSObject.Properties[$prop] -and $item.$prop) { $location = [string]$item.$prop; break } }
                }
                if (-not [string]::IsNullOrWhiteSpace($location) -and (Test-Path -LiteralPath $location)) {
                    $result.Found = $true; $result.Source = 'registry'; $result.Location = $location
                    return $result
                }
            }
            catch { $null = $null }
        }
    }
    return $result
}

function Initialize-IQTools {
    <#
    .SYNOPSIS
    Ensures Tabular Editor 2 portable and pbi-tools exist under Config\ and runs their preflight checks; never throws.
    .DESCRIPTION
    Ports monolith lines 34-160: download the latest pbi-tools zip (GitHub releases API) and Tabular Editor 2
    portable zip, extract them into Config\PBI Tools and Config\TabularEditor, then run the TE2 preflight
    (Blank Model.bim + a trivial script, 2-minute timeout). Every step is wrapped in try/catch; with -SkipToolUpdate,
    IMPACTIQ_OFFLINE=1, or when downloads fail, the binaries already present are used. Sets $script:IQ.Tools
    (TabularEditorPath, PbiToolsPath, TabularEditorWorks, PbiToolsWorks, versions, preflight reasons).
    On non-Windows both *Works flags are $false and nothing is downloaded or executed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [switch]$SkipToolUpdate
    )
    if (-not $script:IQ) { throw "ImpactIQ context missing; call Initialize-IQContext before Initialize-IQTools." }
    $configFolder = [string]$script:IQ.ConfigFolder
    if ([string]::IsNullOrWhiteSpace($configFolder)) { $configFolder = Join-Path $script:IQ.BaseFolder 'Config' }
    $teFolder = Join-Path $configFolder 'TabularEditor'
    $pbiFolder = Join-Path $configFolder 'PBI Tools'
    $tePath = Join-Path $teFolder 'TabularEditor.exe'
    $pbiPath = Join-Path $pbiFolder 'pbi-tools.exe'
    $blankModelPath = Join-Path $configFolder 'Blank Model.bim'
    # The stamp lives under State\ (persisted by the impactiq-state pipeline artifact); Config\ is a fresh checkout per run.
    $stampPath = Get-IQToolVersionStampPath
    $legacyStampPath = Join-Path $configFolder 'tool-versions.json'

    $tools = @{
        TabularEditorPath = $tePath; PbiToolsPath = $pbiPath; TabularEditorWorks = $false; PbiToolsWorks = $false
        TabularEditorVersion = $null; PbiToolsVersion = $null; PbiToolsTag = $null; PbiToolsEdition = $null
        TabularEditorPreflight = $null; PbiToolsPreflight = $null; PbiDesktopFound = $false; PbiDesktopLocation = $null; ToolUpdateSkipped = $false
        WerChanged = $false; WerPrevious = $null; WerHadValue = $false
    }
    $script:IQ.Tools = $tools

    $onWindows = $false
    if ($script:IQ.ContainsKey('IsWindows')) { $onWindows = [bool]$script:IQ.IsWindows } else { $onWindows = ($env:OS -eq 'Windows_NT') }
    if (-not $onWindows) {
        $tools.TabularEditorVersion = Get-IQFileVersion -Path $tePath
        $tools.PbiToolsVersion = Get-IQFileVersion -Path $pbiPath
        $tools.TabularEditorPreflight = 'not-windows'
        $tools.PbiToolsPreflight = 'not-windows'
        $tools.ToolUpdateSkipped = $true
        Write-IQLog -Level Info -Stage Tools -Message "Not running on Windows: Tabular Editor and pbi-tools are unavailable (Model Detail will use the DAX fallback)."
        return $tools
    }

    # Audit C1-06: TLS 1.2 before the GitHub downloads, OR-ed into the existing protocol set.
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { $null = $null }

    $offline = ($env:IMPACTIQ_OFFLINE -eq '1')
    $skipUpdate = [bool]$SkipToolUpdate
    if (-not $skipUpdate -and $script:IQ.Options -and $script:IQ.Options.ContainsKey('SkipToolUpdate') -and $script:IQ.Options.SkipToolUpdate) { $skipUpdate = $true }
    if ($skipUpdate -or $offline) {
        $tools.ToolUpdateSkipped = $true
        Write-IQLog -Level Info -Stage Tools -Message "Tool update skipped (SkipToolUpdate/IMPACTIQ_OFFLINE); using the binaries present under $configFolder"
    }
    else {
        $stamp = Get-IQToolVersionStamp -Path $stampPath
        if (-not $stamp.pbiTools -and -not $stamp.tabularEditor -and (Test-Path -LiteralPath $legacyStampPath)) { $stamp = Get-IQToolVersionStamp -Path $legacyStampPath }
        $newStamp = @{ pbiTools = $stamp.pbiTools; tabularEditor = $stamp.tabularEditor }
        # Zips written by the download branch: the manual-zip scan below must not install them a second time when
        # the install failed (locked exe) and the zip is still there.
        $handledZips = @()

        # ---- pbi-tools (monolith lines 34-49, 63-91) ----
        try {
            Write-IQLog -Level Info -Stage Tools -Message "Fetching latest pbi-tools release..."
            $pbiRelease = Invoke-RestMethod -Uri 'https://api.github.com/repos/pbi-tools/pbi-tools/releases/latest' -UseBasicParsing -TimeoutSec 60 -UserAgent 'ImpactIQ/3.0' -ErrorAction Stop
            $pbiAssets = @()
            try { if ($pbiRelease.PSObject.Properties['assets']) { $pbiAssets = @($pbiRelease.assets) } } catch { $pbiAssets = @() }
            $pbiAsset = Select-IQPbiToolsAsset -Assets $pbiAssets
            $pbiTag = [string]$pbiRelease.tag_name
            $pbiInstalledVersion = Get-IQFileVersion -Path $pbiPath
            $pbiExists = Test-Path -LiteralPath $pbiPath
            if ($pbiAsset -and $pbiAsset.browser_download_url) {
                # Up to date when the persisted stamp OR the exe's own FileVersion matches the release tag, so the
                # decision does not depend on a stamp file surviving between pipeline runs (C1-08).
                if ($pbiExists -and (Test-IQToolVersionCurrent -InstalledVersion $pbiInstalledVersion -Tag $pbiTag -StampTag $stamp.pbiTools)) {
                    Write-IQLog -Level Info -Stage Tools -Message "pbi-tools $pbiTag is already installed (version $pbiInstalledVersion); skipping download"
                    $newStamp.pbiTools = $pbiTag
                }
                elseif ($pbiExists -and [string]::IsNullOrWhiteSpace($pbiTag)) {
                    Write-IQLog -Level Info -Stage Tools -Message "Could not determine the latest pbi-tools release tag; keeping the installed copy (version $pbiInstalledVersion)"
                }
                else {
                    $pbiZip = Join-Path $configFolder 'PBI Tools.zip'
                    $handledZips += $pbiZip
                    Write-IQLog -Level Info -Stage Tools -Message ("Selected pbi-tools release asset '{0}' (Desktop edition) of {1}" -f [string]$pbiAsset.name, $pbiTag)
                    if (Invoke-IQToolDownload -Url ([string]$pbiAsset.browser_download_url) -ZipPath $pbiZip) {
                        Write-IQLog -Level Info -Stage Tools -Message "Downloaded pbi-tools $pbiTag ($([string]$pbiAsset.name)) to $pbiZip"
                        if (Install-IQToolFromZip -ZipPath $pbiZip -TargetFolder $pbiFolder -ExeName 'pbi-tools.exe' -DisplayName 'PBI Tools') { $newStamp.pbiTools = $pbiTag }
                    }
                }
            }
            else {
                Write-IQLog -Level Warn -Stage Tools -Message ("Could not find the pbi-tools Desktop-edition zip (pbi-tools.<version>.zip) among the {0} asset(s) of release {1}; using the existing copy if present." -f $pbiAssets.Count, $pbiTag)
            }
        }
        catch {
            Write-IQLog -Level Warn -Stage Tools -Message ("pbi-tools update failed ({0}); using the existing copy if present." -f $_.Exception.Message)
        }

        # ---- Tabular Editor 2 portable (monolith lines 55-61, 94-121) ----
        try {
            $teTag = Get-IQGitHubLatestTag -Repo 'TabularEditor/TabularEditor'
            $teInstalledVersion = Get-IQFileVersion -Path $tePath
            $teExists = Test-Path -LiteralPath $tePath
            $teUpToDate = $false
            if ($teExists -and (Test-IQToolVersionCurrent -InstalledVersion $teInstalledVersion -Tag $teTag -StampTag $stamp.tabularEditor)) { $teUpToDate = $true }
            if ($teUpToDate) {
                Write-IQLog -Level Info -Stage Tools -Message "Tabular Editor 2 $teTag is already installed (version $teInstalledVersion); skipping download"
                $newStamp.tabularEditor = $teTag
            }
            elseif ($teExists -and [string]::IsNullOrWhiteSpace($teTag)) {
                # Unknown latest tag (probe failed): an installed copy is kept rather than re-downloaded on every run.
                Write-IQLog -Level Info -Stage Tools -Message "Could not determine the latest Tabular Editor 2 release tag; keeping the installed copy (version $teInstalledVersion)"
            }
            else {
                Write-IQLog -Level Info -Stage Tools -Message "Downloading latest Tabular Editor 2 Portable..."
                $teZip = Join-Path $configFolder 'TabularEditor.zip'
                $handledZips += $teZip
                $teUrl = 'https://github.com/TabularEditor/TabularEditor/releases/latest/download/TabularEditor.Portable.zip'
                if (Invoke-IQToolDownload -Url $teUrl -ZipPath $teZip) {
                    Write-IQLog -Level Info -Stage Tools -Message "Downloaded Tabular Editor 2 Portable to $teZip"
                    if (Install-IQToolFromZip -ZipPath $teZip -TargetFolder $teFolder -ExeName 'TabularEditor.exe' -DisplayName 'Tabular Editor') {
                        if ($teTag) { $newStamp.tabularEditor = $teTag } else { $newStamp.tabularEditor = Get-IQFileVersion -Path $tePath }
                    }
                }
            }
        }
        catch {
            Write-IQLog -Level Warn -Stage Tools -Message ("Tabular Editor update failed ({0}); using the existing copy if present." -f $_.Exception.Message)
        }

        # Also honour zips dropped manually into Config\ (monolith patterns "PBI Tools*.zip" / "TabularEditor*.zip").
        # -LiteralPath: a BaseFolder containing [ or ] must not be expanded as a wildcard; the download branch's own
        # zips are skipped (a failed install would otherwise be repeated in the same run).
        try {
            foreach ($spec in @(@{ Pattern = 'PBI Tools*.zip'; Folder = $pbiFolder; Exe = 'pbi-tools.exe'; Name = 'PBI Tools' }, @{ Pattern = 'TabularEditor*.zip'; Folder = $teFolder; Exe = 'TabularEditor.exe'; Name = 'Tabular Editor' })) {
                $zip = Get-ChildItem -LiteralPath $configFolder -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -like $spec.Pattern -and ($handledZips -notcontains $_.FullName) } | Select-Object -First 1
                if ($zip) {
                    Write-IQLog -Level Info -Stage Tools -Message ("Installing {0} from the manually provided zip {1}" -f $spec.Name, $zip.FullName)
                    Install-IQToolFromZip -ZipPath $zip.FullName -TargetFolder $spec.Folder -ExeName $spec.Exe -DisplayName $spec.Name | Out-Null
                }
            }
        }
        catch { Write-IQLog -Level Debug -Stage Tools -Message ("Manual zip scan failed: " + $_.Exception.Message) }

        Save-IQToolVersionStamp -Path $stampPath -Stamp $newStamp
        $tools.PbiToolsTag = $newStamp.pbiTools
    }

    $tools.TabularEditorVersion = Get-IQFileVersion -Path $tePath
    $tools.PbiToolsVersion = Get-IQFileVersion -Path $pbiPath

    # Audit C1-10: keep a crashing TE2/pbi-tools from blocking a headless agent with a WER dialog (no admin needed).
    # Headless runs only - an interactive laptop keeps its setting; the previous value is restored by Restore-IQWerSetting
    # / the PowerShell.Exiting hook registered in Set-IQWerDontShowUI.
    $interactive = $false
    try { if ($script:IQ.ContainsKey('Interactive')) { $interactive = [bool]$script:IQ.Interactive } } catch { $interactive = $false }
    if (-not $interactive) { Set-IQWerDontShowUI }
    else { Write-IQLog -Level Debug -Stage Tools -Message 'Interactive session: the WER DontShowUI setting is left unchanged.' }

    # ---- Tabular Editor preflight (monolith lines 123-150) ----
    try {
        if (-not (Test-Path -LiteralPath $tePath)) {
            $tools.TabularEditorPreflight = 'missing-exe'
        }
        elseif (-not (Test-Path -LiteralPath $blankModelPath)) {
            $tools.TabularEditorPreflight = 'missing-bim'
        }
        else {
            $teTestScriptPath = Join-Path (Get-IQToolTempFolder) ('TabularEditor_Preflight_' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.cs')
            [System.IO.File]::WriteAllText($teTestScriptPath, 'Info("Tabular Editor preflight OK");', (New-Object System.Text.UTF8Encoding($false)))
            try {
                $teTestArgs = "`"$blankModelPath`" -S `"$teTestScriptPath`""
                $preflight = Invoke-IQProcess -FilePath $tePath -ArgumentList $teTestArgs -WorkingDirectory $script:IQ.BaseFolder -TimeoutMinutes 2 -LogName 'te2-preflight' -Stage 'Tools' -Item 'Tabular Editor preflight'
                if ($preflight.TimedOut) { $tools.TabularEditorPreflight = 'timeout' }
                elseif ($preflight.StartError) { $tools.TabularEditorPreflight = 'start-failure' }
                elseif ($preflight.ExitCode -ne 0) { $tools.TabularEditorPreflight = "exit $($preflight.ExitCode)" }
                else { $tools.TabularEditorPreflight = 'ok'; $tools.TabularEditorWorks = $true }
            }
            finally {
                Remove-Item -LiteralPath $teTestScriptPath -Force -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
        $tools.TabularEditorPreflight = 'exception: ' + $_.Exception.Message
    }
    if ($tools.TabularEditorWorks) {
        Write-IQLog -Level Success -Stage Tools -Message ("Tabular Editor preflight OK (version {0})" -f $tools.TabularEditorVersion)
    }
    else {
        Write-IQLog -Level Error -Stage Tools -Message ("Tabular Editor is not able to run ({0}) - Model Detail will fall back to DAX and Report Detail will be skipped" -f $tools.TabularEditorPreflight)
    }

    # ---- pbi-tools preflight (audit C1-18) ----
    try {
        if (-not (Test-Path -LiteralPath $pbiPath)) {
            $tools.PbiToolsPreflight = 'missing-exe'
        }
        else {
            $probe = Invoke-IQProcess -FilePath $pbiPath -ArgumentList 'info' -WorkingDirectory $script:IQ.BaseFolder -TimeoutMinutes 2 -LogName 'pbi-tools-preflight' -Stage 'Tools' -Item 'pbi-tools preflight'
            if ($probe.TimedOut) { $tools.PbiToolsPreflight = 'timeout' }
            elseif ($probe.StartError) { $tools.PbiToolsPreflight = 'start-failure' }
            elseif ($probe.ExitCode -ne 0) { $tools.PbiToolsPreflight = "exit $($probe.ExitCode)" }
            else { $tools.PbiToolsPreflight = 'ok'; $tools.PbiToolsWorks = $true }
            # Parsed "pbi-tools info" JSON (pbiInstalls / effectivePbiInstallDir) first, then path and registry probes.
            $detection = Get-IQPbiDesktopDetection -ProbeOutput ([string]$probe.StdOut)
            $tools.PbiDesktopFound = [bool]$detection.Found
            $tools.PbiDesktopLocation = $detection.Location
            $tools.PbiToolsEdition = $detection.Edition
            if ($detection.Found) { Write-IQLog -Level Debug -Stage Tools -Message ("Power BI Desktop detected via {0}: {1}" -f $detection.Source, $detection.Location) }
        }
    }
    catch {
        $tools.PbiToolsPreflight = 'exception: ' + $_.Exception.Message
    }
    if ($tools.PbiToolsWorks) {
        Write-IQLog -Level Success -Stage Tools -Message ("pbi-tools preflight OK (version {0}; edition {1}; Power BI Desktop detected: {2})" -f $tools.PbiToolsVersion, $tools.PbiToolsEdition, $tools.PbiDesktopFound)
        if ($tools.PbiToolsEdition -and ([string]$tools.PbiToolsEdition -match '(?i)core')) { Write-IQLog -Level Warn -Stage Tools -Message ("The installed pbi-tools is the '{0}' edition; PBIX model extraction needs the Desktop edition (pbi-tools.<version>.zip)." -f $tools.PbiToolsEdition) }
        if (-not $tools.PbiDesktopFound) { Write-IQLog -Level Warn -Stage Tools -Message "Power BI Desktop was not detected; pbi-tools may not be able to extract models from Pro PBIX files on this host (model detail then falls back to DAX)." }
    }
    else {
        Write-IQLog -Level Warn -Stage Tools -Message ("pbi-tools is not able to run ({0}) - Pro workspace models will use the DAX fallback for Model Detail" -f $tools.PbiToolsPreflight)
    }
    return $tools
}
