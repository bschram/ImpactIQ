# ImpactIQ.Tools.ps1 - external tool acquisition (Tabular Editor 2 portable, pbi-tools) and the process runner.
#
# Contract: brief section 2.5. Initialize-IQTools ports the monolith's bootstrap (Final PS Script.txt lines 34-160:
# download latest pbi-tools + TE2 portable, extract, TE2 preflight) with try/catch, -SkipToolUpdate / offline
# fallback and version stamps. Invoke-IQProcess / Invoke-IQProcessBatch / Invoke-IQTabularEditor give every
# external process a timeout, captured output under tool-logs\, and a process-tree kill on timeout.
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
    $logFolder = Get-IQToolLogFolder -Stage $Stage
    if ([string]::IsNullOrWhiteSpace($LogName)) {
        $LogName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath) + '_' + (Get-Date -Format 'yyyyMMdd_HHmmssfff')
    }
    $safeName = Get-IQSafeKey -Value $LogName
    $outFile = Join-Path $logFolder ($safeName + '.out.txt')
    $errFile = Join-Path $logFolder ($safeName + '.err.txt')
    foreach ($file in @($outFile, $errFile)) { if (Test-Path -LiteralPath $file) { Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue } }

    $job = @{
        ItemKey = $ItemKey; Item = $Item; FilePath = $FilePath; ArgumentList = $ArgumentList; WorkingDirectory = $WorkingDirectory
        LogName = $LogName; Stage = $Stage; OutFile = $outFile; ErrFile = $errFile; TimeoutMs = [int64]$TimeoutMinutes * 60000
        Process = $null; Started = $null; StartError = $null; TimedOut = $false; Completed = $false; Result = $null
    }
    $redacted = Get-IQRedactedArgumentList -ArgumentList $ArgumentList
    Write-IQLog -Level Debug -Stage $Stage -Item $Item -Message ("Starting: {0} {1}" -f $FilePath, $redacted)
    try {
        if (-not (Test-Path -LiteralPath $FilePath)) { throw "Executable not found: $FilePath" }
        $startParams = @{ FilePath = $FilePath; PassThru = $true; NoNewWindow = $true; RedirectStandardOutput = $outFile; RedirectStandardError = $errFile; ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace($ArgumentList)) { $startParams.ArgumentList = $ArgumentList }
        if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) {
            if (-not (Test-Path -LiteralPath $WorkingDirectory)) { New-Item -Path $WorkingDirectory -ItemType Directory -Force | Out-Null }
            $startParams.WorkingDirectory = $WorkingDirectory
        }
        $process = Start-Process @startParams
        # Touch the handle so ExitCode is available after exit on Windows PowerShell 5.1.
        try { $null = $process.Handle } catch { $null = $null }
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
    if ($Job.Process) {
        try { $exitCode = [int]$Job.Process.ExitCode } catch { $exitCode = -1 }
        if ($Job.Started) { $duration = [math]::Round(([datetime]::UtcNow - $Job.Started).TotalSeconds, 1) }
        try { $Job.Process.Dispose() } catch { $null = $null }
    }
    if ($Job.TimedOut) { $exitCode = -1 }
    $stdOut = ''
    $stdErr = ''
    try { if (Test-Path -LiteralPath $Job.OutFile) { $stdOut = [string](Get-Content -LiteralPath $Job.OutFile -Raw -ErrorAction SilentlyContinue) } } catch { $stdOut = '' }
    try { if (Test-Path -LiteralPath $Job.ErrFile) { $stdErr = [string](Get-Content -LiteralPath $Job.ErrFile -Raw -ErrorAction SilentlyContinue) } } catch { $stdErr = '' }
    if ($null -eq $stdOut) { $stdOut = '' }
    if ($null -eq $stdErr) { $stdErr = '' }
    if ($Job.StartError) {
        if ($stdErr) { $stdErr = $stdErr.TrimEnd() + [Environment]::NewLine }
        $stdErr += 'Start failure: ' + $Job.StartError
        try { Set-Content -LiteralPath $Job.ErrFile -Value $stdErr -Encoding UTF8 -ErrorAction SilentlyContinue } catch { $null = $null }
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

function Get-IQGitHubLatestTag {
    <#
    .SYNOPSIS
    Resolves the latest release tag of a GitHub repo through a rate-limit-free HEAD redirect probe; $null on failure (private).
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
        if ($location) { return (($location.TrimEnd('/') -split '/')[-1]) }
    }
    catch {
        Write-IQLog -Level Debug -Stage Tools -Message ("Latest-tag probe failed for {0}: {1}" -f $Repo, $_.Exception.Message)
    }
    return $null
}

function Get-IQToolVersionStamp {
    <#
    .SYNOPSIS
    Reads Config\tool-versions.json ({ pbiTools; tabularEditor }) or returns an empty hashtable (private).
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
    Writes Config\tool-versions.json (private).
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

function Install-IQToolFromZip {
    <#
    .SYNOPSIS
    Extracts a tool zip into a temp folder, verifies the expected exe, swaps it into place and removes the zip (private).
    .DESCRIPTION
    Ports monolith lines 63-121 (extract to %TEMP%, then copy over the tool folder) with audit fix C1-09: the
    extracted folder replaces the old one (old renamed to .old and removed on success, restored on failure) so
    stale files do not accumulate; a locked exe is reported instead of silently producing a mixed-version folder.
    Returns $true when the tool folder was updated.
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

        # Stop stray instances that would lock the exe (C1-09).
        try { Get-Process -Name ([System.IO.Path]::GetFileNameWithoutExtension($ExeName)) -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue } catch { $null = $null }

        if (Test-Path -LiteralPath $oldFolder) { Remove-Item -LiteralPath $oldFolder -Recurse -Force -ErrorAction SilentlyContinue }
        $hadExisting = Test-Path -LiteralPath $TargetFolder
        if ($hadExisting) { Rename-Item -LiteralPath $TargetFolder -NewName ([System.IO.Path]::GetFileName($oldFolder)) -Force -ErrorAction Stop }
        try {
            New-Item -Path $TargetFolder -ItemType Directory -Force | Out-Null
            Copy-Item -Path (Join-Path $sourceFolder '*') -Destination $TargetFolder -Recurse -Force -ErrorAction Stop
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
    $stampPath = Join-Path $configFolder 'tool-versions.json'

    $tools = @{
        TabularEditorPath = $tePath; PbiToolsPath = $pbiPath; TabularEditorWorks = $false; PbiToolsWorks = $false
        TabularEditorVersion = $null; PbiToolsVersion = $null; PbiToolsTag = $null
        TabularEditorPreflight = $null; PbiToolsPreflight = $null; PbiDesktopFound = $false; ToolUpdateSkipped = $false
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
        $newStamp = @{ pbiTools = $stamp.pbiTools; tabularEditor = $stamp.tabularEditor }

        # ---- pbi-tools (monolith lines 34-49, 63-91) ----
        try {
            Write-IQLog -Level Info -Stage Tools -Message "Fetching latest pbi-tools release..."
            $pbiRelease = Invoke-RestMethod -Uri 'https://api.github.com/repos/pbi-tools/pbi-tools/releases/latest' -UseBasicParsing -TimeoutSec 60 -UserAgent 'ImpactIQ/3.0' -ErrorAction Stop
            $pbiAsset = $pbiRelease.assets | Where-Object { $_.name -like '*.zip' -and $_.name -notlike '*Desktop*' } | Select-Object -First 1
            $pbiTag = [string]$pbiRelease.tag_name
            if ($pbiAsset -and $pbiAsset.browser_download_url) {
                if ($pbiTag -and $stamp.pbiTools -eq $pbiTag -and (Test-Path -LiteralPath $pbiPath)) {
                    Write-IQLog -Level Info -Stage Tools -Message "pbi-tools $pbiTag is already installed; skipping download"
                }
                else {
                    $pbiZip = Join-Path $configFolder 'PBI Tools.zip'
                    if (Invoke-IQToolDownload -Url ([string]$pbiAsset.browser_download_url) -ZipPath $pbiZip) {
                        Write-IQLog -Level Info -Stage Tools -Message "Downloaded pbi-tools $pbiTag to $pbiZip"
                        if (Install-IQToolFromZip -ZipPath $pbiZip -TargetFolder $pbiFolder -ExeName 'pbi-tools.exe' -DisplayName 'PBI Tools') { $newStamp.pbiTools = $pbiTag }
                    }
                }
            }
            else {
                Write-IQLog -Level Warn -Stage Tools -Message "Could not find a valid pbi-tools zip asset in the latest release; using the existing copy if present."
            }
        }
        catch {
            Write-IQLog -Level Warn -Stage Tools -Message ("pbi-tools update failed ({0}); using the existing copy if present." -f $_.Exception.Message)
        }

        # ---- Tabular Editor 2 portable (monolith lines 55-61, 94-121) ----
        try {
            $teTag = Get-IQGitHubLatestTag -Repo 'TabularEditor/TabularEditor'
            $teInstalledVersion = Get-IQFileVersion -Path $tePath
            $teUpToDate = $false
            if ($teTag -and (Test-Path -LiteralPath $tePath)) {
                if ($stamp.tabularEditor -eq $teTag) { $teUpToDate = $true }
                elseif ($teInstalledVersion -and ($teInstalledVersion -replace '^v', '').StartsWith(($teTag -replace '^v', ''))) { $teUpToDate = $true }
            }
            if ($teUpToDate) {
                Write-IQLog -Level Info -Stage Tools -Message "Tabular Editor 2 $teTag is already installed; skipping download"
                $newStamp.tabularEditor = $teTag
            }
            else {
                Write-IQLog -Level Info -Stage Tools -Message "Downloading latest Tabular Editor 2 Portable..."
                $teZip = Join-Path $configFolder 'TabularEditor.zip'
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
        try {
            foreach ($spec in @(@{ Pattern = 'PBI Tools*.zip'; Folder = $pbiFolder; Exe = 'pbi-tools.exe'; Name = 'PBI Tools' }, @{ Pattern = 'TabularEditor*.zip'; Folder = $teFolder; Exe = 'TabularEditor.exe'; Name = 'Tabular Editor' })) {
                $zip = Get-ChildItem -Path $configFolder -Filter $spec.Pattern -File -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($zip) { Install-IQToolFromZip -ZipPath $zip.FullName -TargetFolder $spec.Folder -ExeName $spec.Exe -DisplayName $spec.Name | Out-Null }
            }
        }
        catch { Write-IQLog -Level Debug -Stage Tools -Message ("Manual zip scan failed: " + $_.Exception.Message) }

        Save-IQToolVersionStamp -Path $stampPath -Stamp $newStamp
        $tools.PbiToolsTag = $newStamp.pbiTools
    }

    $tools.TabularEditorVersion = Get-IQFileVersion -Path $tePath
    $tools.PbiToolsVersion = Get-IQFileVersion -Path $pbiPath

    # Audit C1-10: keep a crashing TE2/pbi-tools from blocking a headless agent with a WER dialog (no admin needed).
    try {
        $werKey = 'HKCU:\Software\Microsoft\Windows\Windows Error Reporting'
        if (-not (Test-Path -LiteralPath $werKey)) { New-Item -Path $werKey -Force | Out-Null }
        New-ItemProperty -Path $werKey -Name 'DontShowUI' -Value 1 -PropertyType DWord -Force | Out-Null
    }
    catch { Write-IQLog -Level Debug -Stage Tools -Message ("Could not set WER DontShowUI: " + $_.Exception.Message) }

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
            $desktopPaths = @()
            if ($env:ProgramFiles) { $desktopPaths += (Join-Path $env:ProgramFiles 'Microsoft Power BI Desktop\bin\PBIDesktop.exe') }
            if ($env:LOCALAPPDATA) { $desktopPaths += (Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\PBIDesktop.exe') }
            foreach ($candidate in $desktopPaths) { if (Test-Path -LiteralPath $candidate) { $tools.PbiDesktopFound = $true } }
            if (-not $tools.PbiDesktopFound -and $probe.StdOut -match '(?i)pbiInstalls' -and $probe.StdOut -match '(?i)PBIDesktop') { $tools.PbiDesktopFound = $true }
        }
    }
    catch {
        $tools.PbiToolsPreflight = 'exception: ' + $_.Exception.Message
    }
    if ($tools.PbiToolsWorks) {
        Write-IQLog -Level Success -Stage Tools -Message ("pbi-tools preflight OK (version {0}; Power BI Desktop detected: {1})" -f $tools.PbiToolsVersion, $tools.PbiDesktopFound)
        if (-not $tools.PbiDesktopFound) { Write-IQLog -Level Warn -Stage Tools -Message "Power BI Desktop was not detected; pbi-tools may not be able to extract models from Pro PBIX files on this host (model detail then falls back to DAX)." }
    }
    else {
        Write-IQLog -Level Warn -Stage Tools -Message ("pbi-tools is not able to run ({0}) - Pro workspace models will use the DAX fallback for Model Detail" -f $tools.PbiToolsPreflight)
    }
    return $tools
}
