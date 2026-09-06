#Requires -Version 5.1
<#
.SYNOPSIS
    ImpactIQ v3 - State module: run manifest, resume decision, stage runner, per-item checkpoints, inventory files.

.DESCRIPTION
    State root: <BaseFolder>\State\
        runs\<RunId>\manifest.json                       run manifest (brief section 2.4)
        runs\<RunId>\inventory\workspaces.json|global.json|ws-<safeWorkspaceId>.json
        runs\<RunId>\done\<Stage>\<safeItemKey>.json     per-item checkpoint
        runs\<RunId>\extracts\...                        intermediate JSON
        runs\<RunId>\tool-logs\<Stage>\<safeItemKey>.out.txt / .err.txt
    Run status values: Running | Completed | CompletedWithErrors | Paused (time budget, brief -TimeBudgetMinutes, or a
    stage left Interrupted/unfinished by a stage-filtered resume) | Failed | Cancelled. -Resume Auto resumes the
    manifest for <RunId> whenever it is not Completed; without -RunId it also adopts the newest Running/Paused/Failed
    run (never a run that finished on its own as CompletedWithErrors - that is retried only by naming its RunId).
    Timestamps loaded from disk are always ISO-8601 strings (PowerShell 7's ConvertFrom-Json turns them into
    [datetime]; Read-IQManifestFile converts them back so sheets/logs are culture-independent on both hosts).
    The in-memory manifest ($script:IQ.Manifest) is a nested [ordered] hashtable (a manifest loaded from disk is
    converted with ConvertTo-IQHashtable) so modules can read/write it uniformly: $IQ.Manifest.stages['Inventory'].status.
    Requires ImpactIQ.Common.ps1 to be dot-sourced first. Windows PowerShell 5.1 compatible.
#>

$script:IQStageOrder = @('Inventory', 'ModelBackup', 'ReportBackup', 'ReportDetail', 'ModelDetail', 'Dataflows', 'Extras', 'Assemble')

function Get-IQRunsRoot {
    <#
    .SYNOPSIS
        Returns <BaseFolder>\State\runs (created on demand).
    #>
    [CmdletBinding()]
    param()
    if (-not $script:IQ) { throw 'ImpactIQ context not initialised. Call Initialize-IQContext first.' }
    $root = Join-Path $script:IQ.StatePath 'runs'
    if (-not (Test-Path -LiteralPath $root)) { New-Item -ItemType Directory -Path $root -Force | Out-Null }
    return $root
}

function Test-IQRunIdValue {
    <#
    .SYNOPSIS
        Validates a RunId is a plain folder name (no separators, not empty, not "." / "..").
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$RunId)
    if ([string]::IsNullOrWhiteSpace($RunId)) { return $false }
    if ($RunId -match '[\\/:*?"<>|]') { return $false }
    if ($RunId -eq '.' -or $RunId -eq '..') { return $false }
    return $true
}

function Get-IQRunDate {
    <#
    .SYNOPSIS
        Today's "yyyy-MM-dd" (local time), or the injected test clock's date when Options.NowUtc is set.
    #>
    [CmdletBinding()]
    param()
    if ($script:IQ -and $script:IQ.Options -and $script:IQ.Options.Contains('NowUtc') -and $null -ne $script:IQ.Options['NowUtc'] -and
        -not ($script:IQ.Options['NowUtc'] -is [string] -and [string]::IsNullOrWhiteSpace($script:IQ.Options['NowUtc']))) {
        return (Get-IQNowUtc).ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return (Get-Date).ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-IQUtcStamp {
    <#
    .SYNOPSIS
        Current (or injected) UTC time as an ISO-8601 round-trip string.
    #>
    [CmdletBinding()]
    param()
    return (Get-IQNowUtc).ToString('o', [System.Globalization.CultureInfo]::InvariantCulture)
}

function ConvertTo-IQManifestOptionSet {
    <#
    .SYNOPSIS
        Copies the effective entry-point options into a JSON-friendly ordered hashtable with every secret removed.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Options)
    $out = [ordered]@{}
    if ($null -eq $Options) { return $out }
    $secretPattern = '(?i)(password|secret|credential|tokencachekey|pbi_token|fabric_token|accesstoken|refreshtoken)'
    foreach ($key in @($Options.Keys | Sort-Object)) {
        $k = [string]$key
        if ($k -match $secretPattern) { continue }
        if ($k -eq 'NowUtc') { continue }
        $v = $Options[$key]
        if ($null -eq $v) { $out[$k] = $null; continue }
        if ($v -is [System.Management.Automation.PSCredential]) { continue }
        if ($v -is [System.Security.SecureString]) { continue }
        if ($k -match '(?i)webhook') {
            # An incoming-webhook URL (Teams/Slack) is the credential: keep only the host so the manifest (published as a
            # pipeline artifact) shows where the device-code message went without the secret path.
            $out[$k] = ConvertTo-IQRedactedUrl -Url ([string]$v)
            continue
        }
        if ($v -is [System.Management.Automation.SwitchParameter]) { $out[$k] = [bool]$v; continue }
        if ($v -is [string] -or $v -is [bool] -or $v.GetType().IsPrimitive -or $v -is [decimal]) { $out[$k] = $v; continue }
        if ($v -is [datetime]) { $out[$k] = $v.ToString('o'); continue }
        if ($v -is [System.Collections.IDictionary]) { $out[$k] = ConvertTo-IQHashtable -InputObject $v; continue }
        if ($v -is [System.Collections.IEnumerable]) { $out[$k] = @($v | ForEach-Object { [string]$_ }); continue }
        $out[$k] = [string]$v
    }
    return $out
}

function ConvertTo-IQRedactedUrl {
    <#
    .SYNOPSIS
        "https://host/***" for a URL whose path carries a secret (webhooks); "***" for any other non-empty text; '' stays ''.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return '' }
    if ($Url -match '^(?<prefix>[A-Za-z][A-Za-z0-9+.-]*://[^/?#]+)') { return ($Matches['prefix'] + '/***') }
    return '***'
}

function ConvertTo-IQTimestampText {
    <#
    .SYNOPSIS
        A manifest timestamp as the ISO-8601 round-trip string ('o', UTC) whether it is stored as text or as [datetime]; '' for $null.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [datetime]) {
        $d = [datetime]$Value
        if ($d.Kind -eq [System.DateTimeKind]::Unspecified) { $d = [datetime]::SpecifyKind($d, [System.DateTimeKind]::Utc) }
        return $d.ToUniversalTime().ToString('o', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return [string]$Value
}

function ConvertTo-IQManifestTimestampText {
    <#
    .SYNOPSIS
        Walks a manifest loaded from disk and replaces every [datetime] leaf (PowerShell 7 ConvertFrom-Json) with its ISO-8601 UTC string.
    .DESCRIPTION
        Dictionaries and PSCustomObjects are updated in place; arrays are rebuilt. Returns the (possibly replaced) value
        so scalar leaves can be assigned back: $h[$k] = ConvertTo-IQManifestTimestampText -InputObject $h[$k].
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$InputObject)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [datetime]) { return (ConvertTo-IQTimestampText -Value $InputObject) }
    if ($InputObject -is [string] -or $InputObject.GetType().IsPrimitive -or $InputObject -is [decimal]) { return $InputObject }
    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($k in @($InputObject.Keys)) { $InputObject[$k] = ConvertTo-IQManifestTimestampText -InputObject $InputObject[$k] }
        return $InputObject
    }
    if ($InputObject -is [System.Collections.IEnumerable]) {
        $list = @()
        foreach ($i in $InputObject) { $list += , (ConvertTo-IQManifestTimestampText -InputObject $i) }
        return , $list
    }
    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        foreach ($p in @($InputObject.PSObject.Properties)) {
            if ($p.IsSettable) { $p.Value = ConvertTo-IQManifestTimestampText -InputObject $p.Value }
        }
        return $InputObject
    }
    return $InputObject
}

function New-IQManifest {
    <#
    .SYNOPSIS
        Builds a fresh run manifest (schema of brief section 2.4) as an ordered hashtable.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$RunId)
    $now = Get-IQUtcStamp
    $authText = ''
    if ($script:IQ.Auth -and $script:IQ.Auth['Description']) { $authText = [string]$script:IQ.Auth['Description'] }
    elseif (Get-Command -Name Get-IQAuthDescription -ErrorAction SilentlyContinue) {
        try { $authText = [string](Get-IQAuthDescription) } catch { $authText = '' }
    }
    $envName = ''
    if ($script:IQ.Environment) { $envName = [string]$script:IQ.Environment }

    $scope = [ordered]@{ runMode = 'Workspaces'; workspaceIds = @(); reportIds = @(); datasetIds = @(); includeMyWorkspace = $false }
    if ($script:IQ.Options) {
        if ($script:IQ.Options['RunMode']) { $scope.runMode = [string]$script:IQ.Options['RunMode'] }
        if ($script:IQ.Options['IncludeMyWorkspace']) { $scope.includeMyWorkspace = [bool]$script:IQ.Options['IncludeMyWorkspace'] }
    }

    $manifest = [ordered]@{
        schemaVersion = 1
        runId         = $RunId
        status        = 'Running'
        startedUtc    = $now
        updatedUtc    = $now
        endedUtc      = $null
        host          = [ordered]@{
            machine       = [Environment]::MachineName
            user          = [Environment]::UserName
            psVersion     = $PSVersionTable.PSVersion.ToString()
            isAzureDevOps = [bool]$script:IQ.IsAzureDevOps
        }
        auth          = $authText
        environment   = $envName
        options       = (ConvertTo-IQManifestOptionSet -Options $script:IQ.Options)
        scope         = $scope
        stages        = [ordered]@{}
        failures      = @()
        outputs       = [ordered]@{ environmentWorkbook = $null; reportWorkbook = $null; modelWorkbook = $null; dataflowWorkbook = $null }
    }
    return $manifest
}

function Read-IQManifestFile {
    <#
    .SYNOPSIS
        Reads a manifest.json into an ordered hashtable ($null when missing or unreadable).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)
    try {
        $obj = ConvertFrom-IQJsonFile -Path $Path
        if ($null -eq $obj) { return $null }
        $h = ConvertTo-IQHashtable -InputObject $obj
        if (-not ($h -is [System.Collections.IDictionary])) { return $null }
        # PowerShell 7 parses the 'o'-format strings into [datetime]; keep the manifest text-only on both hosts.
        ConvertTo-IQManifestTimestampText -InputObject $h | Out-Null
        if (-not $h.Contains('stages') -or $null -eq $h['stages']) { $h['stages'] = [ordered]@{} }
        if (-not $h.Contains('failures') -or $null -eq $h['failures']) { $h['failures'] = @() }
        $h['failures'] = @($h['failures'])
        if (-not $h.Contains('outputs') -or $null -eq $h['outputs']) {
            $h['outputs'] = [ordered]@{ environmentWorkbook = $null; reportWorkbook = $null; modelWorkbook = $null; dataflowWorkbook = $null }
        }
        if (-not $h.Contains('scope') -or $null -eq $h['scope']) {
            $h['scope'] = [ordered]@{ runMode = 'Workspaces'; workspaceIds = @(); reportIds = @(); datasetIds = @(); includeMyWorkspace = $false }
        }
        return $h
    }
    catch {
        Write-IQLog -Level Warn -Message "Could not read manifest '$Path': $($_.Exception.Message)"
        return $null
    }
}

function ConvertTo-IQDateTimeUtc {
    <#
    .SYNOPSIS
        Parses an ISO timestamp from a manifest into a UTC [datetime]; $null when empty/invalid.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $d = [datetime]::MinValue
    if ([datetime]::TryParse($text, [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal, [ref]$d)) {
        return [datetime]::SpecifyKind($d, [System.DateTimeKind]::Utc)
    }
    return $null
}

function Find-IQResumableRun {
    <#
    .SYNOPSIS
        Newest unfinished run (status Running/Paused/Failed, started within MaxAgeDays) that a start without -RunId may adopt; returns @{RunId; Manifest} or $null.
    .DESCRIPTION
        Only auto-generated (yyyy-MM-dd) RunIds are considered, so an ad-hoc "-RunId smoke" run is never picked up by a
        scheduled start, and only runs recorded for the current environment (a USGov pipeline never resumes a Public
        test run with its foreign scope). A run that finished on its own as CompletedWithErrors is not adopted: the next
        scheduled start produces a fresh dated snapshot, and the failed items are retried by naming that RunId
        (-RunId <old>, rule 1 of Initialize-IQRun).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][int]$MaxAgeDays = 3,
        [Parameter(Mandatory = $false)][string]$ExcludeRunId
    )
    $root = Get-IQRunsRoot
    $now = Get-IQNowUtc
    $best = $null
    $bestStart = [datetime]::MinValue
    $currentEnv = ''
    if ($script:IQ.Environment) { $currentEnv = [string]$script:IQ.Environment }
    foreach ($dir in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
        if ($ExcludeRunId -and $dir.Name -eq $ExcludeRunId) { continue }
        if ($dir.Name -notmatch '^\d{4}-\d{2}-\d{2}$') {
            Write-IQLog -Level Debug -Message "Resume candidate '$($dir.Name)' skipped: not an auto-generated (yyyy-MM-dd) RunId - resume it explicitly with -RunId."
            continue
        }
        $mPath = Join-Path $dir.FullName 'manifest.json'
        if (-not (Test-Path -LiteralPath $mPath)) { continue }
        $m = Read-IQManifestFile -Path $mPath
        if ($null -eq $m) { continue }
        if ([string]$m['status'] -notin @('Running', 'Paused', 'Failed')) { continue }
        $runEnv = [string]$m['environment']
        if ($currentEnv -and $runEnv -and $runEnv -ne $currentEnv) {
            Write-IQLog -Level Info -Message "Resume candidate '$($dir.Name)' (status $($m['status'])) skipped: it was recorded for environment '$runEnv', this run targets '$currentEnv'."
            continue
        }
        $started = ConvertTo-IQDateTimeUtc -Value $m['startedUtc']
        if ($null -eq $started) { continue }
        $age = $now - $started
        if ($age.TotalDays -gt $MaxAgeDays -or $age.TotalDays -lt -1) { continue }
        if ($started -gt $bestStart) { $bestStart = $started; $best = @{ RunId = $dir.Name; Manifest = $m } }
    }
    return $best
}

function Clear-IQRunFolders {
    <#
    .SYNOPSIS
        Deletes "Model Backups\<RunId>", "Report Backups\<RunId>", "Dataflow Backups\<RunId>" and (unless -KeepState) "State\runs\<RunId>".
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RunId,
        [Parameter(Mandatory = $false)][switch]$KeepState
    )
    if (-not (Test-IQRunIdValue -RunId $RunId)) { throw "Invalid RunId '$RunId' (must be a plain folder name)." }
    if (-not $script:IQ) { throw 'ImpactIQ context not initialised. Call Initialize-IQContext first.' }
    $targets = @(
        (Join-Path $script:IQ.Paths.ModelBackups $RunId),
        (Join-Path $script:IQ.Paths.ReportBackups $RunId),
        (Join-Path $script:IQ.Paths.DataflowBackups $RunId)
    )
    if (-not $KeepState) { $targets += (Join-Path (Get-IQRunsRoot) $RunId) }
    foreach ($t in $targets) {
        if (Test-Path -LiteralPath $t) {
            Write-IQLog -Level Info -Message "Clearing existing folder for run '$RunId': $t"
            Remove-Item -LiteralPath $t -Recurse -Force -ErrorAction SilentlyContinue
            if (Test-Path -LiteralPath $t) {
                # Second attempt after a short pause (Windows file handles from a previous process).
                Start-Sleep -Milliseconds 500
                Remove-Item -LiteralPath $t -Recurse -Force -ErrorAction SilentlyContinue
            }
            if (Test-Path -LiteralPath $t) { Write-IQLog -Level Warn -Message "Could not fully remove '$t' (files may be locked)." }
        }
    }
}

function Initialize-IQRun {
    <#
    .SYNOPSIS
        Decides the effective RunId and whether this is a resume (brief section 5.2), loads/creates the manifest and writes it.
    .DESCRIPTION
        Effective RunId: -RunId if given, else today's yyyy-MM-dd (Options.NowUtc overrides the clock for tests).
        -Force / -ResumePolicy Never => fresh run (run state and the three backup folders for <RunId> are cleared).
        Always => resume <RunId> when its manifest exists (any status), else a new manifest is started for <RunId>
        WITHOUT clearing anything (the entry point implies Always for "-Stages Assemble" etc.; files already in the
        backup folders for <RunId> are exactly what those stages need).
        Auto => (1) manifest for <RunId> exists and status <> Completed (Running, Paused, Failed, ...) => resume it;
        (2) else, when -RunId was not given, the newest Running/Paused/Failed run (dated RunId, same environment)
        started within -ResumeMaxAgeDays is resumed (yesterday's pipeline run died or paused - see
        Find-IQResumableRun); (3) else fresh: a Completed (or unreadable) manifest for <RunId> is archived as
        manifest.<timestamp>.json, the run's done/inventory/extracts/tool-logs folders and the backup folders for
        <RunId> are cleared - legacy re-run semantics.
        Sets $script:IQ.RunId, RunPath, IsResume, Manifest, RunPaths.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$RunId,
        [Parameter(Mandatory = $false)][ValidateSet('Auto', 'Always', 'Never')][string]$ResumePolicy = 'Auto',
        [Parameter(Mandatory = $false)][int]$ResumeMaxAgeDays = 3,
        [Parameter(Mandatory = $false)][switch]$Force
    )
    if (-not $script:IQ) { throw 'ImpactIQ context not initialised. Call Initialize-IQContext first.' }

    $runIdGiven = -not [string]::IsNullOrWhiteSpace($RunId)
    $effectiveRunId = $RunId
    if (-not $runIdGiven) { $effectiveRunId = Get-IQRunDate }
    $effectiveRunId = $effectiveRunId.Trim()
    if (-not (Test-IQRunIdValue -RunId $effectiveRunId)) { throw "Invalid RunId '$effectiveRunId' (must be a plain folder name without path separators)." }

    $runsRoot = Get-IQRunsRoot
    $runPath = Join-Path $runsRoot $effectiveRunId
    $manifestPath = Join-Path $runPath 'manifest.json'
    $existing = $null
    if (Test-Path -LiteralPath $manifestPath) { $existing = Read-IQManifestFile -Path $manifestPath }

    $isResume = $false
    $manifest = $null
    $reason = ''

    if ($Force) {
        $reason = '-Force: fresh run'
        Clear-IQRunFolders -RunId $effectiveRunId
    }
    elseif ($ResumePolicy -eq 'Never') {
        $reason = '-Resume Never: fresh run'
        Clear-IQRunFolders -RunId $effectiveRunId
    }
    elseif ($ResumePolicy -eq 'Always') {
        if ($null -ne $existing) {
            $isResume = $true
            $manifest = $existing
            $reason = "-Resume Always: resuming run '$effectiveRunId' (status $($existing['status']))"
        }
        else {
            # No clearing here: "-Resume Always -Stages Assemble" (implied by the entry point whenever Inventory is not
            # in the stage list) must never delete the backups it is about to assemble, even when the State folder was
            # not restored or the manifest is unreadable.
            $reason = "-Resume Always: no readable manifest for '$effectiveRunId', starting a new manifest (existing files in the run folders are kept)"
            Write-IQLog -Level Warn -Message "-Resume Always: no readable manifest for '$effectiveRunId'; starting a new manifest but keeping any existing backup and state folders for it."
        }
    }
    else {
        # Auto
        if ($null -ne $existing -and [string]$existing['status'] -ne 'Completed') {
            $isResume = $true
            $manifest = $existing
            $reason = "-Resume Auto: manifest for '$effectiveRunId' has status '$($existing['status'])', resuming"
        }
        else {
            $candidate = $null
            if (-not $runIdGiven) { $candidate = Find-IQResumableRun -MaxAgeDays $ResumeMaxAgeDays -ExcludeRunId $effectiveRunId }
            if ($null -ne $candidate) {
                $isResume = $true
                $manifest = $candidate.Manifest
                $reason = "-Resume Auto: resuming unfinished run '$($candidate.RunId)' (status $($candidate.Manifest['status']), started $($candidate.Manifest['startedUtc'])) instead of starting '$effectiveRunId'"
                $effectiveRunId = $candidate.RunId
                $runPath = Join-Path $runsRoot $effectiveRunId
                $manifestPath = Join-Path $runPath 'manifest.json'
            }
            else {
                $reason = "-Resume Auto: fresh run '$effectiveRunId'"
                if (Test-Path -LiteralPath $manifestPath) {
                    $stamp = (Get-IQNowUtc).ToString('yyyyMMddHHmmss', [System.Globalization.CultureInfo]::InvariantCulture)
                    $archive = Join-Path $runPath ('manifest.' + $stamp + '.json')
                    $n = 1
                    while (Test-Path -LiteralPath $archive) { $archive = Join-Path $runPath ('manifest.' + $stamp + '-' + $n + '.json'); $n++ }
                    Move-Item -LiteralPath $manifestPath -Destination $archive -Force
                    if ($null -ne $existing) { $reason = $reason + " (previous Completed manifest archived as $(Split-Path -Leaf $archive))" }
                    else { $reason = $reason + " (unreadable manifest archived as $(Split-Path -Leaf $archive))" }
                }
                if (Test-Path -LiteralPath $runPath) {
                    # Drop any previous checkpoints/inventory for this RunId - whether its manifest was Completed,
                    # unreadable or missing - so a fresh run never honours stale checkpoints or ws-*.json files.
                    foreach ($sub in @('done', 'inventory', 'extracts', 'tool-logs')) {
                        $p = Join-Path $runPath $sub
                        if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue }
                    }
                }
                Clear-IQRunFolders -RunId $effectiveRunId -KeepState
            }
        }
    }

    if ($null -eq $manifest) { $manifest = New-IQManifest -RunId $effectiveRunId }
    else {
        $manifest['runId'] = $effectiveRunId
        $manifest['status'] = 'Running'
        $manifest['endedUtc'] = $null
        if (-not $manifest.Contains('resumes') -or $null -eq $manifest['resumes']) { $manifest['resumes'] = @() }
        $manifest['resumes'] = @($manifest['resumes']) + @(Get-IQUtcStamp)
        $manifest['resumeCount'] = @($manifest['resumes']).Count
        # Stages interrupted mid-flight are re-run; their counters are recomputed from checkpoints by Invoke-IQStage.
        foreach ($stageName in @($manifest['stages'].Keys)) {
            $st = $manifest['stages'][$stageName]
            if ($st -is [System.Collections.IDictionary] -and [string]$st['status'] -eq 'Running') { $st['status'] = 'Interrupted' }
        }
        # Refresh provenance for this attempt (auth may differ from the original attempt).
        $fresh = New-IQManifest -RunId $effectiveRunId
        $manifest['host'] = $fresh['host']
        if (-not [string]::IsNullOrEmpty([string]$fresh['auth'])) { $manifest['auth'] = $fresh['auth'] }
        if (-not [string]::IsNullOrEmpty([string]$fresh['environment'])) { $manifest['environment'] = $fresh['environment'] }
        $manifest['options'] = $fresh['options']
    }

    # Folders for this run.
    foreach ($sub in @('inventory', 'done', 'extracts', 'tool-logs')) {
        $p = Join-Path $runPath $sub
        if (-not (Test-Path -LiteralPath $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
    }
    $runPaths = @{
        Run             = $runPath
        Manifest        = $manifestPath
        Inventory       = (Join-Path $runPath 'inventory')
        Done            = (Join-Path $runPath 'done')
        Extracts        = (Join-Path $runPath 'extracts')
        ToolLogs        = (Join-Path $runPath 'tool-logs')
        ModelBackups    = (Join-Path $script:IQ.Paths.ModelBackups $effectiveRunId)
        ReportBackups   = (Join-Path $script:IQ.Paths.ReportBackups $effectiveRunId)
        DataflowBackups = (Join-Path $script:IQ.Paths.DataflowBackups $effectiveRunId)
    }
    foreach ($p in @($runPaths.ModelBackups, $runPaths.ReportBackups, $runPaths.DataflowBackups)) {
        if (-not (Test-Path -LiteralPath $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
    }

    $script:IQ.RunId = $effectiveRunId
    $script:IQ.RunPath = $runPath
    $script:IQ.IsResume = $isResume
    $script:IQ.Manifest = $manifest
    $script:IQ.RunPaths = $runPaths

    Save-IQManifest
    if ($isResume) { Write-IQLog -Level Info -Message ("Run '{0}' RESUMED. {1}" -f $effectiveRunId, $reason) }
    else { Write-IQLog -Level Info -Message ("Run '{0}' started. {1}" -f $effectiveRunId, $reason) }
    Write-IQLog -Level Debug -Message ("Run state: {0}" -f $runPath)
    return $manifest
}

function Save-IQManifest {
    <#
    .SYNOPSIS
        Atomically writes $script:IQ.Manifest to State\runs\<RunId>\manifest.json (updates updatedUtc). Never throws.
    #>
    [CmdletBinding()]
    param()
    if (-not $script:IQ -or $null -eq $script:IQ.Manifest -or [string]::IsNullOrEmpty($script:IQ.RunPath)) {
        Write-IQLog -Level Debug -Message 'Save-IQManifest called before Initialize-IQRun; nothing saved.'
        return
    }
    try {
        $script:IQ.Manifest['updatedUtc'] = Get-IQUtcStamp
        ConvertTo-IQJsonFile -Object $script:IQ.Manifest -Path (Join-Path $script:IQ.RunPath 'manifest.json')
    }
    catch {
        Write-IQLog -Level Warn -Message "Could not save manifest: $($_.Exception.Message)"
    }
}

function Get-IQStageEntry {
    <#
    .SYNOPSIS
        Returns (creating when missing) the manifest entry for a stage.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Name)
    if ($null -eq $script:IQ.Manifest) { throw 'No run manifest. Call Initialize-IQRun first.' }
    $stages = $script:IQ.Manifest['stages']
    if ($null -eq $stages) { $stages = [ordered]@{}; $script:IQ.Manifest['stages'] = $stages }
    if (-not $stages.Contains($Name)) {
        $stages[$Name] = [ordered]@{ status = 'Pending'; startedUtc = $null; endedUtc = $null; itemsDone = 0; itemsFailed = 0; error = $null }
    }
    $entry = $stages[$Name]
    foreach ($k in @('status', 'startedUtc', 'endedUtc', 'itemsDone', 'itemsFailed', 'error')) {
        if (-not $entry.Contains($k)) { $entry[$k] = $null }
    }
    if ($null -eq $entry['itemsDone']) { $entry['itemsDone'] = 0 }
    if ($null -eq $entry['itemsFailed']) { $entry['itemsFailed'] = 0 }
    return $entry
}

function Get-IQDoneFolder {
    <#
    .SYNOPSIS
        Returns State\runs\<RunId>\done\<Stage> (created when -Create).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Stage,
        [Parameter(Mandatory = $false)][switch]$Create
    )
    if (-not $script:IQ -or [string]::IsNullOrEmpty($script:IQ.RunPath)) { throw 'No active run. Call Initialize-IQRun first.' }
    $folder = Join-Path (Join-Path $script:IQ.RunPath 'done') $Stage
    if ($Create -and -not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    return $folder
}

function Update-IQStageCounter {
    <#
    .SYNOPSIS
        Recomputes itemsDone/itemsFailed for a stage from its checkpoint files (authoritative after a resume).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Stage)
    $entry = Get-IQStageEntry -Name $Stage
    $folder = Get-IQDoneFolder -Stage $Stage
    $done = 0
    $failed = 0
    if (Test-Path -LiteralPath $folder) {
        foreach ($f in @(Get-ChildItem -LiteralPath $folder -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
            try {
                $cp = ConvertFrom-IQJsonFile -Path $f.FullName
                $st = [string](Get-IQMemberValue -Object $cp -Name 'status')
                if ($st -eq 'Failed') { $failed++ } elseif ($st -in @('Succeeded', 'Skipped')) { $done++ }
            }
            catch { $null = $_.Exception }
        }
    }
    $entry['itemsDone'] = $done
    $entry['itemsFailed'] = $failed
    return $entry
}

function Invoke-IQStage {
    <#
    .SYNOPSIS
        Runs one stage body with manifest bookkeeping; skips stages already Completed in a resumed run (Assemble always re-runs).
    .DESCRIPTION
        Marks the stage Running, runs -Body in try/catch, then marks Completed / CompletedWithErrors (item failures
        recorded during the stage) / Failed (exception) / Paused (the body stopped because Test-IQTimeBudget set
        $script:IQ.BudgetExceeded). Once the budget is exhausted every later stage except Assemble is marked Paused
        without running (Assemble still runs so partial workbooks exist); Paused stages are re-run on resume. -Fatal
        rethrows the exception; otherwise the error is logged and later stages still run. A Completed stage is only
        skipped when every one of its checkpoints is still valid (Test-IQItemDone: outputs present) - when backup files
        were deleted or were not restored on this agent the stage runs again and re-creates just the missing items.
        "Inventory" is re-run on resume when Options.RefreshInventory is set (its checkpoints, ws-*.json / global.json
        files and Inventory failure entries are cleared first). Returns the stage status.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$Body,
        [Parameter(Mandatory = $false)][switch]$Fatal
    )
    if ($null -eq $script:IQ.Manifest) { throw 'No run manifest. Call Initialize-IQRun first.' }
    $entry = Get-IQStageEntry -Name $Name
    $refreshInventory = ($Name -eq 'Inventory' -and $script:IQ.Options -and [bool]$script:IQ.Options['RefreshInventory'])

    if ($script:IQ.IsResume -and [string]$entry['status'] -eq 'Completed' -and $Name -ne 'Assemble' -and -not $refreshInventory) {
        $stale = Get-IQStaleCheckpointCount -Stage $Name
        if ($stale -eq 0) {
            Write-IQLog -Level Info -Stage $Name -Message "Stage already Completed in this run - skipping (resume)."
            return 'Completed'
        }
        Write-IQLog -Level Warn -Stage $Name -Message ("Stage was Completed in this run but {0} checkpoint(s) have missing outputs - re-running the stage (items whose outputs still exist are skipped)." -f $stale)
    }
    if ($Name -ne 'Assemble' -and (Test-IQTimeBudget -Stage $Name)) {
        Write-IQLog -Level Warn -Stage $Name -Message 'Time budget exhausted - stage not started (marked Paused; it runs on the next start, which resumes this run).'
        $entry['status'] = 'Paused'
        $entry['error'] = $null
        Save-IQManifest
        return 'Paused'
    }
    if ($refreshInventory -and $script:IQ.IsResume) {
        Write-IQLog -Level Info -Stage $Name -Message '-RefreshInventory: clearing Inventory checkpoints, ws-*.json / global.json files and Inventory failure entries so every workspace is collected again.'
        $doneFolder = Get-IQDoneFolder -Stage $Name
        if (Test-Path -LiteralPath $doneFolder) { Remove-Item -LiteralPath $doneFolder -Recurse -Force -ErrorAction SilentlyContinue }
        # A workspace removed from the tenant/scope since the original run must not survive as a stale ws-<id>.json
        # (later stages and Assemble read every ws-*.json); workspaces.json (the persisted scope) is kept.
        $invFolder = Join-Path $script:IQ.RunPath 'inventory'
        if (Test-Path -LiteralPath $invFolder) {
            foreach ($f in @(Get-ChildItem -LiteralPath $invFolder -Filter 'ws-*.json' -File -ErrorAction SilentlyContinue)) {
                Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
            }
            $globalFile = Join-Path $invFolder 'global.json'
            if (Test-Path -LiteralPath $globalFile) { Remove-Item -LiteralPath $globalFile -Force -ErrorAction SilentlyContinue }
        }
        $script:IQ.Manifest['failures'] = @(@($script:IQ.Manifest['failures']) | Where-Object { $null -ne $_ -and [string](Get-IQMemberValue -Object $_ -Name 'stage') -ne $Name })
    }

    $previousStage = $script:IQ.CurrentStage
    $script:IQ.CurrentStage = $Name
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Update-IQStageCounter -Stage $Name | Out-Null
    $entry['status'] = 'Running'
    $entry['startedUtc'] = Get-IQUtcStamp
    $entry['endedUtc'] = $null
    $entry['error'] = $null
    Save-IQManifest
    Write-IQLog -Level Info -Stage $Name -Message '=== Stage started ==='

    $status = 'Completed'
    $caught = $null
    try {
        & $Body | Out-Null
    }
    catch {
        $caught = $_
        $status = 'Failed'
    }

    $sw.Stop()
    $script:IQ.CurrentStage = $previousStage
    Update-IQStageCounter -Stage $Name | Out-Null
    if ($status -ne 'Failed' -and [int]$entry['itemsFailed'] -gt 0) { $status = 'CompletedWithErrors' }
    if ($status -ne 'Failed' -and $Name -ne 'Assemble' -and $script:IQ.ContainsKey('BudgetExceeded') -and [bool]$script:IQ['BudgetExceeded']) { $status = 'Paused' }
    $entry['status'] = $status
    $entry['endedUtc'] = Get-IQUtcStamp
    $entry['durationSeconds'] = [math]::Round($sw.Elapsed.TotalSeconds, 1)
    if ($null -ne $caught) { $entry['error'] = ConvertTo-IQRedactedText -Text $caught.Exception.Message }
    Save-IQManifest

    $summary = ("=== Stage {0}: {1} item(s) done, {2} failed, {3:N1} s ===" -f $status, $entry['itemsDone'], $entry['itemsFailed'], $sw.Elapsed.TotalSeconds)
    switch ($status) {
        'Completed' { Write-IQLog -Level Success -Stage $Name -Message $summary }
        'CompletedWithErrors' { Write-IQLog -Level Warn -Stage $Name -Message $summary }
        'Paused' { Write-IQLog -Level Warn -Stage $Name -Message ($summary + ' (time budget reached - the remaining items run on the next start)') }
        default {
            Write-IQLog -Level Error -Stage $Name -Message ("Stage failed: {0}" -f $caught.Exception.Message) -Exception $caught.Exception
            if ($caught.ScriptStackTrace) { Write-IQLog -Level Debug -Stage $Name -Message $caught.ScriptStackTrace }
            Write-IQLog -Level Error -Stage $Name -Message $summary
        }
    }
    if ($status -eq 'Failed' -and $Fatal) { throw $caught }
    return $status
}

function Get-IQItemCheckpoint {
    <#
    .SYNOPSIS
        Returns the checkpoint object for an item (done\<Stage>\<safeKey>.json) or $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Stage,
        [Parameter(Mandatory = $true)][string]$ItemKey
    )
    $path = Join-Path (Get-IQDoneFolder -Stage $Stage) ((Get-IQSafeKey -Value $ItemKey) + '.json')
    try { return (ConvertFrom-IQJsonFile -Path $path) } catch { return $null }
}

function Test-IQItemDone {
    <#
    .SYNOPSIS
        $true when the item's checkpoint is Succeeded/Skipped and every path in its outputs still exists.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Stage,
        [Parameter(Mandatory = $true)][string]$ItemKey
    )
    $cp = Get-IQItemCheckpoint -Stage $Stage -ItemKey $ItemKey
    if ($null -eq $cp) { return $false }
    return (Test-IQCheckpointValid -Checkpoint $cp -Stage $Stage -ItemKey $ItemKey)
}

function Test-IQCheckpointValid {
    <#
    .SYNOPSIS
        $true when a parsed checkpoint has status Succeeded/Skipped and every path in its outputs still exists (logs the first missing output at Debug).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Checkpoint,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Stage,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$ItemKey
    )
    if ($null -eq $Checkpoint) { return $false }
    $status = [string](Get-IQMemberValue -Object $Checkpoint -Name 'status')
    if ($status -notin @('Succeeded', 'Skipped')) { return $false }
    foreach ($o in @(Get-IQMemberValue -Object $Checkpoint -Name 'outputs')) {
        if ($null -eq $o) { continue }
        $p = [string]$o
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        if (-not (Test-Path -LiteralPath $p)) {
            $logArgs = @{ Level = 'Debug'; Message = "Checkpoint invalidated: output missing '$p'." }
            if (-not [string]::IsNullOrWhiteSpace($Stage)) { $logArgs['Stage'] = $Stage }
            if (-not [string]::IsNullOrWhiteSpace($ItemKey)) { $logArgs['Item'] = $ItemKey }
            Write-IQLog @logArgs
            return $false
        }
    }
    return $true
}

function Get-IQStaleCheckpointCount {
    <#
    .SYNOPSIS
        Number of checkpoint files in done\<Stage> that Test-IQItemDone would reject (status not Succeeded/Skipped, or an output missing).
    .DESCRIPTION
        Used before a Completed stage is skipped on resume: on an agent that restored only State\runs (never the backup
        folders), or after an operator deleted a backup folder, the stage must run again to re-create the missing files.
        One directory scan per stage; unreadable checkpoint files count as stale.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Stage)
    $folder = Get-IQDoneFolder -Stage $Stage
    $stale = 0
    if (-not (Test-Path -LiteralPath $folder)) { return $stale }
    foreach ($f in @(Get-ChildItem -LiteralPath $folder -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
        $cp = $null
        try { $cp = ConvertFrom-IQJsonFile -Path $f.FullName } catch { $cp = $null }
        if ($null -eq $cp) { $stale++; continue }
        $key = [string](Get-IQMemberValue -Object $cp -Name 'itemKey')
        if (-not (Test-IQCheckpointValid -Checkpoint $cp -Stage $Stage -ItemKey $key)) { $stale++ }
    }
    return $stale
}

function Remove-IQFailureEntry {
    <#
    .SYNOPSIS
        Removes manifest.failures entries for a stage/itemKey; returns how many were removed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Stage,
        [Parameter(Mandatory = $true)][string]$ItemKey
    )
    $kept = @()
    $removed = 0
    foreach ($f in @($script:IQ.Manifest['failures'])) {
        if ($null -eq $f) { continue }
        if ([string](Get-IQMemberValue -Object $f -Name 'stage') -eq $Stage -and [string](Get-IQMemberValue -Object $f -Name 'itemKey') -eq $ItemKey) { $removed++; continue }
        $kept += , $f
    }
    $script:IQ.Manifest['failures'] = $kept
    return $removed
}

function Set-IQItemDone {
    <#
    .SYNOPSIS
        Writes an item checkpoint (done\<Stage>\<safeKey>.json) and updates the manifest counters / failures list.
    .DESCRIPTION
        Status Succeeded/Skipped: any earlier failure entry for the item is removed (retry succeeded). Status Failed:
        an entry {stage,itemKey,item,message,timeUtc} is appended to manifest.failures. Counters are adjusted from the
        previous checkpoint status so re-recording an item never double-counts. Returns the checkpoint hashtable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Stage,
        [Parameter(Mandatory = $true)][string]$ItemKey,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Item,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$Outputs,
        [Parameter(Mandatory = $false)][ValidateSet('Succeeded', 'Skipped', 'Failed')][string]$Status = 'Succeeded',
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Message,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Method,
        [Parameter(Mandatory = $false)][AllowNull()][hashtable]$Data
    )
    if ($null -eq $script:IQ.Manifest) { throw 'No run manifest. Call Initialize-IQRun first.' }
    $folder = Get-IQDoneFolder -Stage $Stage -Create
    $safeKey = Get-IQSafeKey -Value $ItemKey
    $path = Join-Path $folder ($safeKey + '.json')

    $previousStatus = $null
    $attempt = 1
    $previous = $null
    try { $previous = ConvertFrom-IQJsonFile -Path $path } catch { $previous = $null }
    if ($null -ne $previous) {
        $previousStatus = [string](Get-IQMemberValue -Object $previous -Name 'status')
        $prevAttempt = Get-IQMemberValue -Object $previous -Name 'attempt'
        if ($null -ne $prevAttempt) { $attempt = [int]$prevAttempt + 1 }
    }

    $outputList = @()
    foreach ($o in @($Outputs)) { if (-not [string]::IsNullOrWhiteSpace([string]$o)) { $outputList += [string]$o } }
    $cleanMessage = ConvertTo-IQRedactedText -Text $Message
    $dataOut = $null
    if ($null -ne $Data) { $dataOut = ConvertTo-IQHashtable -InputObject $Data }

    $checkpoint = [ordered]@{
        stage    = $Stage
        itemKey  = $ItemKey
        item     = $Item
        status   = $Status
        message  = $cleanMessage
        method   = $Method
        outputs  = $outputList
        data     = $dataOut
        timeUtc  = Get-IQUtcStamp
        attempt  = $attempt
        runId    = $script:IQ.RunId
    }
    ConvertTo-IQJsonFile -Object $checkpoint -Path $path

    # Manifest bookkeeping.
    $entry = Get-IQStageEntry -Name $Stage
    if ([string]$entry['status'] -eq 'Pending') { $entry['status'] = 'Running' }
    $wasDone = ($previousStatus -in @('Succeeded', 'Skipped'))
    $wasFailed = ($previousStatus -eq 'Failed')
    if ($wasDone) { $entry['itemsDone'] = [Math]::Max(0, [int]$entry['itemsDone'] - 1) }
    if ($wasFailed) { $entry['itemsFailed'] = [Math]::Max(0, [int]$entry['itemsFailed'] - 1) }
    Remove-IQFailureEntry -Stage $Stage -ItemKey $ItemKey | Out-Null

    if ($Status -eq 'Failed') {
        $entry['itemsFailed'] = [int]$entry['itemsFailed'] + 1
        $script:IQ.Manifest['failures'] = @($script:IQ.Manifest['failures']) + @(
            [ordered]@{ stage = $Stage; itemKey = $ItemKey; item = $Item; message = $cleanMessage; timeUtc = $checkpoint.timeUtc }
        )
        Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message ("FAILED: {0}" -f $cleanMessage)
    }
    else {
        $entry['itemsDone'] = [int]$entry['itemsDone'] + 1
        $note = $Status
        if ($Method) { $note = $note + ' via ' + $Method }
        if ($cleanMessage) { $note = $note + ' - ' + $cleanMessage }
        Write-IQLog -Level Debug -Stage $Stage -Item $Item -Message $note
    }
    Save-IQManifest
    return $checkpoint
}

function Get-IQInventoryPath {
    <#
    .SYNOPSIS
        Path of State\runs\<RunId>\inventory\<safeName>.json.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Name)
    if (-not $script:IQ -or [string]::IsNullOrEmpty($script:IQ.RunPath)) { throw 'No active run. Call Initialize-IQRun first.' }
    $folder = Join-Path $script:IQ.RunPath 'inventory'
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    return (Join-Path $folder ((Get-IQSafeKey -Value $Name) + '.json'))
}

function Save-IQInventory {
    <#
    .SYNOPSIS
        Writes an inventory object to State\runs\<RunId>\inventory\<Name>.json (Name: workspaces, global, ws-<id>, ...).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][AllowNull()]$Object
    )
    $path = Get-IQInventoryPath -Name $Name
    ConvertTo-IQJsonFile -Object $Object -Path $path
    Write-IQLog -Level Debug -Message ("Inventory saved: {0}" -f $path)
    return $path
}

function Get-IQInventory {
    <#
    .SYNOPSIS
        Reads an inventory file (see Save-IQInventory); $null when missing. A saved array is emitted element by element (wrap in @( )).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Name)
    $path = Get-IQInventoryPath -Name $Name
    return (ConvertFrom-IQJsonFile -Path $path)
}

function Get-IQAllWorkspaceInventories {
    <#
    .SYNOPSIS
        The parsed ws-*.json inventory objects for the current run (sorted by file name); wrap in @( ) for an array.
    #>
    [CmdletBinding()]
    param()
    if (-not $script:IQ -or [string]::IsNullOrEmpty($script:IQ.RunPath)) { throw 'No active run. Call Initialize-IQRun first.' }
    $folder = Join-Path $script:IQ.RunPath 'inventory'
    $result = @()
    if (Test-Path -LiteralPath $folder) {
        foreach ($f in @(Get-ChildItem -LiteralPath $folder -Filter 'ws-*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
            try {
                $obj = ConvertFrom-IQJsonFile -Path $f.FullName
                if ($null -ne $obj) { $result += , $obj }
            }
            catch { Write-IQLog -Level Warn -Message "Could not read inventory file '$($f.FullName)': $($_.Exception.Message)" }
        }
    }
    return $result
}

function Get-IQRunStatus {
    <#
    .SYNOPSIS
        Derives the run status from the stage states: Paused (any stage Paused, or Interrupted/Running - i.e. unfinished), CompletedWithErrors (any failures / Failed stage), else Completed.
    .DESCRIPTION
        Derived from the stages only: a time budget that trips inside Assemble (which always runs last) leaves every
        stage Completed, so the run is Completed and no extra resume is forced. A stage a crash left Interrupted that
        a stage-filtered resume (-Stages Assemble) did not re-run keeps the run Paused (exit code 3, resumable) so the
        missing items are never reported as success or dropped from the next start's resume.
    #>
    [CmdletBinding()]
    param()
    $m = $script:IQ.Manifest
    if ($null -eq $m) { return 'Failed' }
    $withErrors = $false
    $paused = $false
    if (@($m['failures']).Count -gt 0) { $withErrors = $true }
    if (@(Get-IQUnfinishedStageName -Manifest $m).Count -gt 0) { $paused = $true }
    foreach ($name in @($m['stages'].Keys)) {
        $st = [string]$m['stages'][$name]['status']
        if ($st -in @('Failed', 'CompletedWithErrors')) { $withErrors = $true }
    }
    if ($paused) { return 'Paused' }
    if ($withErrors) { return 'CompletedWithErrors' }
    return 'Completed'
}

function Get-IQUnfinishedStageName {
    <#
    .SYNOPSIS
        Names of the stages whose status is Paused, Interrupted or Running (started but not finished); wrap in @( ).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Manifest)
    if ($null -eq $Manifest) { $Manifest = $script:IQ.Manifest }
    $names = @()
    if ($null -eq $Manifest) { return $names }
    $stages = Get-IQMemberValue -Object $Manifest -Name 'stages'
    if (-not ($stages -is [System.Collections.IDictionary])) { return $names }
    foreach ($name in @($stages.Keys)) {
        $st = [string](Get-IQMemberValue -Object $stages[$name] -Name 'status')
        if ($st -in @('Paused', 'Interrupted', 'Running')) { $names += [string]$name }
    }
    return $names
}

function Complete-IQRun {
    <#
    .SYNOPSIS
        Sets endedUtc and the final status (Completed | CompletedWithErrors | Paused | Failed; derived when omitted) and saves the manifest.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Status)
    if ($null -eq $script:IQ.Manifest) { Write-IQLog -Level Debug -Message 'Complete-IQRun: no manifest to complete.'; return $null }
    if ([string]::IsNullOrWhiteSpace($Status)) { $Status = Get-IQRunStatus }
    $script:IQ.Manifest['status'] = $Status
    $script:IQ.Manifest['endedUtc'] = Get-IQUtcStamp
    Save-IQManifest
    $level = 'Success'
    if ($Status -in @('CompletedWithErrors', 'Paused')) { $level = 'Warn' } elseif ($Status -ne 'Completed') { $level = 'Error' }
    Write-IQLog -Level $level -Message ("Run '{0}' finished with status {1} ({2} failure(s))." -f $script:IQ.RunId, $Status, @($script:IQ.Manifest['failures']).Count)
    if ($Status -eq 'Paused') {
        $unfinished = @(Get-IQUnfinishedStageName)
        $why = 'time budget'
        if ($unfinished.Count -gt 0) { $why = ('unfinished stage(s): ' + ($unfinished -join ', ')) }
        Write-IQLog -Level Warn -Message ("Run '{0}' is PAUSED ({1}): start ImpactIQ again to resume it - completed items are skipped." -f $script:IQ.RunId, $why)
    }
    return $Status
}

function Get-IQRunSummary {
    <#
    .SYNOPSIS
        Rows for the RunSummary sheet / console table: one "(Run)" row followed by one row per stage in pipeline order (wrap in @( )).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Manifest)
    if ($null -eq $Manifest) { $Manifest = $script:IQ.Manifest }
    $rows = @()
    if ($null -eq $Manifest) { return $rows }

    $runId = [string](Get-IQMemberValue -Object $Manifest -Name 'runId')
    $hostInfo = Get-IQMemberValue -Object $Manifest -Name 'host'
    $scope = Get-IQMemberValue -Object $Manifest -Name 'scope'
    $failures = @(Get-IQMemberValue -Object $Manifest -Name 'failures')
    $stages = Get-IQMemberValue -Object $Manifest -Name 'stages'
    $environment = [string](Get-IQMemberValue -Object $Manifest -Name 'environment')
    $auth = [string](Get-IQMemberValue -Object $Manifest -Name 'auth')
    $resumeCount = Get-IQMemberValue -Object $Manifest -Name 'resumeCount'
    if ($null -eq $resumeCount) { $resumeCount = 0 }

    $stageNames = @()
    if ($stages -is [System.Collections.IDictionary]) { $stageNames = @($stages.Keys) }
    elseif ($null -ne $stages) { $stageNames = @($stages.PSObject.Properties | ForEach-Object { $_.Name }) }
    $ordered = @($script:IQStageOrder | Where-Object { $_ -in $stageNames }) + @($stageNames | Where-Object { $_ -notin $script:IQStageOrder })

    $totalDone = 0
    $totalFailed = 0
    $stageRows = @()
    foreach ($name in $ordered) {
        $st = Get-IQMemberValue -Object $stages -Name $name
        $done = Get-IQMemberValue -Object $st -Name 'itemsDone'
        $failed = Get-IQMemberValue -Object $st -Name 'itemsFailed'
        if ($null -eq $done) { $done = 0 }
        if ($null -eq $failed) { $failed = 0 }
        $totalDone += [int]$done
        $totalFailed += [int]$failed
        $dur = Get-IQMemberValue -Object $st -Name 'durationSeconds'
        if ($null -eq $dur) {
            $s = ConvertTo-IQDateTimeUtc -Value (Get-IQMemberValue -Object $st -Name 'startedUtc')
            $e = ConvertTo-IQDateTimeUtc -Value (Get-IQMemberValue -Object $st -Name 'endedUtc')
            if ($null -ne $s -and $null -ne $e) { $dur = [math]::Round(($e - $s).TotalSeconds, 1) }
        }
        $stageRows += [PSCustomObject]@{
            RunId           = $runId
            Stage           = $name
            Status          = [string](Get-IQMemberValue -Object $st -Name 'status')
            StartedUtc      = ConvertTo-IQTimestampText -Value (Get-IQMemberValue -Object $st -Name 'startedUtc')
            EndedUtc        = ConvertTo-IQTimestampText -Value (Get-IQMemberValue -Object $st -Name 'endedUtc')
            DurationSeconds = $dur
            ItemsDone       = [int]$done
            ItemsFailed     = [int]$failed
            Error           = [string](Get-IQMemberValue -Object $st -Name 'error')
            Environment     = $environment
            Auth            = $auth
            RunMode         = [string](Get-IQMemberValue -Object $scope -Name 'runMode')
            Machine         = [string](Get-IQMemberValue -Object $hostInfo -Name 'machine')
            User            = [string](Get-IQMemberValue -Object $hostInfo -Name 'user')
            PSVersion       = [string](Get-IQMemberValue -Object $hostInfo -Name 'psVersion')
            IsAzureDevOps   = [bool](Get-IQMemberValue -Object $hostInfo -Name 'isAzureDevOps')
            ResumeCount     = [int]$resumeCount
        }
    }

    $runStart = ConvertTo-IQDateTimeUtc -Value (Get-IQMemberValue -Object $Manifest -Name 'startedUtc')
    $runEnd = ConvertTo-IQDateTimeUtc -Value (Get-IQMemberValue -Object $Manifest -Name 'endedUtc')
    if ($null -eq $runEnd) { $runEnd = ConvertTo-IQDateTimeUtc -Value (Get-IQMemberValue -Object $Manifest -Name 'updatedUtc') }
    $runDur = $null
    if ($null -ne $runStart -and $null -ne $runEnd) { $runDur = [math]::Round(($runEnd - $runStart).TotalSeconds, 1) }
    $runError = ''
    if ($failures.Count -gt 0) { $runError = ('{0} item failure(s)' -f $failures.Count) }

    $rows += [PSCustomObject]@{
        RunId           = $runId
        Stage           = '(Run)'
        Status          = [string](Get-IQMemberValue -Object $Manifest -Name 'status')
        StartedUtc      = ConvertTo-IQTimestampText -Value (Get-IQMemberValue -Object $Manifest -Name 'startedUtc')
        EndedUtc        = ConvertTo-IQTimestampText -Value (Get-IQMemberValue -Object $Manifest -Name 'endedUtc')
        DurationSeconds = $runDur
        ItemsDone       = $totalDone
        ItemsFailed     = $totalFailed
        Error           = $runError
        Environment     = $environment
        Auth            = $auth
        RunMode         = [string](Get-IQMemberValue -Object $scope -Name 'runMode')
        Machine         = [string](Get-IQMemberValue -Object $hostInfo -Name 'machine')
        User            = [string](Get-IQMemberValue -Object $hostInfo -Name 'user')
        PSVersion       = [string](Get-IQMemberValue -Object $hostInfo -Name 'psVersion')
        IsAzureDevOps   = [bool](Get-IQMemberValue -Object $hostInfo -Name 'isAzureDevOps')
        ResumeCount     = [int]$resumeCount
    }
    $rows += $stageRows
    return $rows
}
