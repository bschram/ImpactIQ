# ImpactIQ.Models.ps1 - ModelBackup and ModelDetail stages.
#
# Contract: brief sections 2.7, 7.1, 7.2 and 13 (+ task models-revise).
#   Invoke-IQModelBackupStage  - TE2 XMLA export of every dedicated-capacity dataset in scope to
#                                "<Model Backups>\<RunId>\<CleanWs> ~ <CleanModel>.bim" (monolith lines 2545-2620), run through
#                                Invoke-IQProcessBatch with a per-job rename script (Method XMLA). When TE2 is unavailable or an
#                                export fails and a Fabric token exists, the TMSL definition is fetched with Fabric
#                                semanticModels/{id}/getDefinition?format=TMSL and its model.bim saved under the same name
#                                (Method FabricDefinition). Pro workspaces: Fabric best effort, else Skipped (ReportBackup/PBIX).
#   Invoke-IQModelDetailStage  - "<CleanWs> ~ <CleanModel>.csv" / "_MD.csv" per dataset, via the two Tabular Editor csx scripts
#                                (monolith lines 3140-3290), the built-in TMSL parser (ImpactIQ.Bim.ps1) and/or the DAX
#                                INFO.VIEW.* fallback (Get-IQModelDetailViaDax). ModelDetailMethod: Auto | TabularEditor | Bim |
#                                Dax | Both (Auto/Both order: TabularEditor -> Bim -> Dax).
#
# Windows PowerShell 5.1 and PowerShell 7 compatible. Tabular Editor paths are Windows-only and guarded by
# $script:IQ.IsWindows / $script:IQ.Tools.TabularEditorWorks; on other hosts the Fabric, Bim and DAX paths still run.
#
# Cross-module functions used (brief section 2): Write-IQLog, Get-IQCleanName, Get-IQSafeKey, Get-IQToken,
# Get-IQDateFolder, Invoke-IQProcessBatch, Invoke-IQFabricLro, Test-IQItemDone, Set-IQItemDone, Get-IQItemCheckpoint,
# ConvertFrom-IQJsonFile, Get-IQSelectedDatasets, Get-IQSelectedWorkspaces, Get-IQModelBackupFileName,
# Get-IQModelDetailViaDax, Get-IQDaxModelAsOfDate, Export-IQModelDetailFromBim (ImpactIQ.Bim.ps1, loaded here when
# ImpactIQ.ps1 has not dot-sourced it yet). Private helpers are prefixed *-IQModel* and are not part of the contract.

if (-not (Get-Command -Name 'Export-IQModelDetailFromBim' -ErrorAction SilentlyContinue)) {
    $iqModelsBimModule = Join-Path $PSScriptRoot 'ImpactIQ.Bim.ps1'
    if (Test-Path -LiteralPath $iqModelsBimModule) { . $iqModelsBimModule }
}


function Get-IQModelOption {
    <#
    .SYNOPSIS
    Reads an entry-point option from $script:IQ.Options with a default (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $false)][AllowNull()]$Default
    )
    try {
        if ($script:IQ -and $script:IQ.Options -and $script:IQ.Options.ContainsKey($Name)) {
            $v = $script:IQ.Options[$Name]
            if ($null -ne $v -and -not ($v -is [string] -and [string]::IsNullOrWhiteSpace($v))) { return $v }
        }
    }
    catch { $null = $null }
    return $Default
}

function Get-IQModelMember {
    <#
    .SYNOPSIS
    Reads a named member from a hashtable/dictionary or an object property; $null when absent (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -ne $prop) { return $prop.Value }
    return $null
}

function Get-IQModelRunFolder {
    <#
    .SYNOPSIS
    Returns (and creates) "<Model Backups>\<RunId>" for the current run (private).
    #>
    [CmdletBinding()]
    param()
    $folder = $null
    if ($script:IQ.ContainsKey('RunPaths') -and $script:IQ.RunPaths -and $script:IQ.RunPaths.ModelBackups) { $folder = [string]$script:IQ.RunPaths.ModelBackups }
    else {
        $runId = [string]$script:IQ.RunId
        if ([string]::IsNullOrWhiteSpace($runId)) { $runId = Get-Date -Format 'yyyy-MM-dd' }
        $folder = Join-Path ([string]$script:IQ.Paths.ModelBackups) $runId
    }
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    return $folder
}

function Get-IQModelTempFolder {
    <#
    .SYNOPSIS
    Returns (and creates) the short scratch folder for per-job Tabular Editor scripts: <BaseFolder>\Config\Temp (private).
    #>
    [CmdletBinding()]
    param()
    $folder = $null
    if ($script:IQ.Paths -and $script:IQ.Paths.TempExtract) { $folder = [string]$script:IQ.Paths.TempExtract }
    elseif ($script:IQ.ConfigFolder) { $folder = Join-Path ([string]$script:IQ.ConfigFolder) 'Temp' }
    else { $folder = Join-Path ([System.IO.Path]::GetTempPath()) 'ImpactIQ' }
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    return $folder
}

function Test-IQModelTabularEditorAvailable {
    <#
    .SYNOPSIS
    True when Tabular Editor 2 passed its preflight on this Windows host (private).
    #>
    [CmdletBinding()]
    param()
    if (-not $script:IQ.IsWindows) { return $false }
    if (-not $script:IQ.Tools) { return $false }
    $works = $false
    try { $works = [bool]$script:IQ.Tools.TabularEditorWorks } catch { $works = $false }
    if (-not $works) { return $false }
    $path = [string]$script:IQ.Tools.TabularEditorPath
    return (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path))
}

function Get-IQModelTabularEditorReason {
    <#
    .SYNOPSIS
    One-line explanation of why Tabular Editor cannot be used (private).
    #>
    [CmdletBinding()]
    param()
    if (-not $script:IQ.IsWindows) { return 'Tabular Editor requires Windows; this host is not Windows' }
    $reason = $null
    try { if ($script:IQ.Tools -and $script:IQ.Tools.ContainsKey('TabularEditorPreflight')) { $reason = [string]$script:IQ.Tools.TabularEditorPreflight } } catch { $reason = $null }
    if ($reason) { return ('Tabular Editor preflight failed: ' + $reason) }
    return 'Tabular Editor is not available (preflight did not pass)'
}

function ConvertTo-IQModelBool {
    <#
    .SYNOPSIS
    Converts an API/JSON value (bool, "true"/"false", 0/1) to [bool]; $null when unknown (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { return $Value }
    $text = ([string]$Value).Trim()
    if ($text -ieq 'true' -or $text -eq '1') { return $true }
    if ($text -ieq 'false' -or $text -eq '0' -or $text -eq '') { return $false }
    try { return [System.Convert]::ToBoolean($Value) } catch { return $null }
}

function Get-IQModelWorkList {
    <#
    .SYNOPSIS
    Builds the per-dataset work list for both model stages from the inventory (scope-filtered datasets + workspaces) (private).
    .DESCRIPTION
    Each entry: @{ Key (DatasetId); Item; Dataset; DatasetId; DatasetName; WorkspaceId; WorkspaceName; IsDedicated;
    IsPseudoWorkspace; NoAccess; BaseName; BimPath; CsvPath; MdPath }. BaseName is "<CleanWs> ~ <CleanModel>"; when two datasets
    in scope sanitise to the same name a " (<8-char id>)" suffix is appended to the later ones (audit C6-10/C8-06) so
    backups never overwrite each other. Datasets of the "Shared Reports (No Workspace Access)" pseudo workspace are
    flagged NoAccess; "My Workspace" is never dedicated (no XMLA endpoint).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$RunFolder)
    $datasets = @()
    try { $datasets = @(Get-IQSelectedDatasets | Where-Object { $null -ne $_ }) }
    catch {
        Write-IQLog -Level Error -Message ("Could not read the selected datasets from the inventory: " + $_.Exception.Message) -Exception $_.Exception
        return @()
    }
    $workspaceDedicated = @{}
    try {
        foreach ($ws in @(Get-IQSelectedWorkspaces)) {
            if ($null -eq $ws) { continue }
            $id = [string](Get-IQModelMember -Object $ws -Name 'WorkspaceId')
            if ($id -eq '') { continue }
            $workspaceDedicated[$id] = ConvertTo-IQModelBool -Value (Get-IQModelMember -Object $ws -Name 'WorkspaceIsOnDedicatedCapacity')
        }
    }
    catch { Write-IQLog -Level Debug -Message ("Get-IQSelectedWorkspaces failed; using dataset-level capacity flags: " + $_.Exception.Message) }

    $work = New-Object System.Collections.Generic.List[object]
    $usedNames = @{}
    $seenIds = @{}
    foreach ($ds in $datasets) {
        $datasetId = [string](Get-IQModelMember -Object $ds -Name 'DatasetId')
        if ([string]::IsNullOrWhiteSpace($datasetId)) { continue }
        if ($seenIds.ContainsKey($datasetId)) { continue }
        $seenIds[$datasetId] = $true
        $workspaceId = [string](Get-IQModelMember -Object $ds -Name 'WorkspaceId')
        $workspaceName = [string](Get-IQModelMember -Object $ds -Name 'WorkspaceName')
        $datasetName = [string](Get-IQModelMember -Object $ds -Name 'DatasetName')
        $isPseudo = -not (Test-IQModelGuid -Value $workspaceId)
        $noAccess = ($workspaceId -eq 'Shared Reports (No Workspace Access)' -or $workspaceName -eq 'Shared Reports (No Workspace Access)')
        $dedicated = $null
        if ($workspaceDedicated.ContainsKey($workspaceId)) { $dedicated = $workspaceDedicated[$workspaceId] }
        if ($null -eq $dedicated) { $dedicated = ConvertTo-IQModelBool -Value (Get-IQModelMember -Object $ds -Name 'WorkspaceIsOnDedicatedCapacity') }
        if ($null -eq $dedicated) {
            Write-IQLog -Level Debug -Item $datasetName -Message "Capacity type of workspace '$workspaceName' unknown; treating as Pro (no XMLA export)"
            $dedicated = $false
        }
        if ($isPseudo) { $dedicated = $false }

        $baseName = Get-IQModelBackupFileName -WorkspaceName $workspaceName -DatasetName $datasetName
        if ($usedNames.ContainsKey($baseName.ToLowerInvariant())) {
            $suffix = $datasetId
            if ($suffix.Length -gt 8) { $suffix = $suffix.Substring(0, 8) }
            $unique = $baseName + ' (' + $suffix + ')'
            Write-IQLog -Level Warn -Item $baseName -Message "Another dataset in scope sanitises to the same file name; using '$unique' for dataset $datasetId"
            $baseName = $unique
        }
        $usedNames[$baseName.ToLowerInvariant()] = $true

        $work.Add(@{
                Key               = $datasetId
                Item              = $baseName
                Dataset           = $ds
                DatasetId         = $datasetId
                DatasetName       = $datasetName
                WorkspaceId       = $workspaceId
                WorkspaceName     = $workspaceName
                IsDedicated       = [bool]$dedicated
                IsPseudoWorkspace = $isPseudo
                NoAccess          = $noAccess
                BaseName          = $baseName
                BimPath           = (Join-Path $RunFolder ($baseName + '.bim'))
                CsvPath           = (Join-Path $RunFolder ($baseName + '.csv'))
                MdPath            = (Join-Path $RunFolder ($baseName + '_MD.csv'))
            })
    }
    return $work.ToArray()
}

function Test-IQModelGuid {
    <#
    .SYNOPSIS
    True when the value looks like a GUID (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value.Trim() -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')
}

function Test-IQModelFileHasContent {
    <#
    .SYNOPSIS
    True when the file exists and is larger than zero bytes; zero-byte leftovers are deleted (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    $length = 0
    try { $length = (Get-Item -LiteralPath $Path).Length } catch { $length = 0 }
    if ($length -gt 0) { return $true }
    try { Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue } catch { $null = $null }
    return $false
}

function Select-IQModelProcessError {
    <#
    .SYNOPSIS
    First error-looking lines (max 5) from a Tabular Editor process result, for checkpoint messages (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Result)
    if ($null -eq $Result) { return @() }
    $lines = @()
    $stdOut = [string](Get-IQModelMember -Object $Result -Name 'StdOut')
    $stdErr = [string](Get-IQModelMember -Object $Result -Name 'StdErr')
    if ($stdOut) {
        $all = @($stdOut -split "`r?`n")
        for ($i = 0; $i -lt $all.Count; $i++) {
            if ($all[$i] -match '^\s*(Error\b|Script compilation error|Script error|Unhandled exception|Exception:|Could not|Unable to|Failed)') {
                $lines += $all[$i].Trim()
                if ($i + 1 -lt $all.Count -and $all[$i + 1].Trim()) { $lines += $all[$i + 1].Trim() }
            }
        }
    }
    if ($stdErr) { $lines += @(($stdErr -split "`r?`n") | Where-Object { $_.Trim() } | ForEach-Object { $_.Trim() }) }
    if ($lines.Count -eq 0 -and $stdOut) { $lines = @(($stdOut -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -Last 3) }
    return @($lines | Select-Object -First 5)
}

function ConvertTo-IQModelResultList {
    <#
    .SYNOPSIS
    Normalises the Invoke-IQProcessBatch return value (which may arrive as a single nested array) into a flat array of entries (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Results)
    if ($null -eq $Results) { return @() }
    $list = @()
    foreach ($r in @($Results)) {
        if ($null -eq $r) { continue }
        if ($r -is [System.Collections.IDictionary]) { $list += , $r }
        elseif ($r -is [System.Collections.IEnumerable] -and -not ($r -is [string])) { foreach ($inner in $r) { if ($inner -is [System.Collections.IDictionary]) { $list += , $inner } } }
        else { $list += , $r }
    }
    return $list
}

function Get-IQModelResultSummary {
    <#
    .SYNOPSIS
    Short human-readable summary of an Invoke-IQProcess result (exit code / timeout / first error) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Result)
    if ($null -eq $Result) { return 'no process result' }
    $parts = @()
    if (Get-IQModelMember -Object $Result -Name 'TimedOut') { $parts += 'timed out' }
    $startError = Get-IQModelMember -Object $Result -Name 'StartError'
    if ($startError) { $parts += ('start failure: ' + $startError) }
    $exit = Get-IQModelMember -Object $Result -Name 'ExitCode'
    if ($null -ne $exit) { $parts += ('exit code ' + $exit) }
    $errorLines = @(Select-IQModelProcessError -Result $Result)
    if ($errorLines.Count -gt 0) { $parts += ($errorLines -join ' | ') }
    $logFile = Get-IQModelMember -Object $Result -Name 'OutFile'
    if ($logFile) { $parts += ('log: ' + $logFile) }
    return ($parts -join '; ')
}

function ConvertTo-IQModelCSharpString {
    <#
    .SYNOPSIS
    Escapes a value for use inside a C# string literal in a Tabular Editor script (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Value)
    if ($null -eq $Value) { return '' }
    return $Value.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n')
}

function New-IQModelRenameScript {
    <#
    .SYNOPSIS
    Writes the per-job Tabular Editor C# script that sets Model.Database.Name to the backup base name (private).
    .DESCRIPTION
    Same script as the monolith (line 2600) but one file per job, TabularEditor_RenameModel_<safeKey>.cs under
    Config\Temp, so parallel exports never share a file. Returns the script path.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)][string]$BaseName
    )
    $folder = Get-IQModelTempFolder
    $path = Join-Path $folder ('TabularEditor_RenameModel_' + (Get-IQSafeKey -Value $Key) + '.cs')
    $content = 'Model.Database.Name = "' + (ConvertTo-IQModelCSharpString -Value $BaseName) + '";'
    $encoding = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($path, $content + [Environment]::NewLine, $encoding)
    return $path
}

function Complete-IQModelBackupJob {
    <#
    .SYNOPSIS
    Checkpoints one finished XMLA export job: Succeeded when the .bim exists and is non-empty, else Failed (private).
    .DESCRIPTION
    With -DeferFailure a missing/empty .bim is not checkpointed: the failure summary is stored on the work entry
    (XmlaFailure) so the stage can try the Fabric getDefinition fallback after the batch and record one combined outcome.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Entry,
        [Parameter(Mandatory = $true)][hashtable]$WorkMap,
        [Parameter(Mandatory = $false)][switch]$DeferFailure
    )
    $key = [string](Get-IQModelMember -Object $Entry -Name 'ItemKey')
    if (-not $WorkMap.ContainsKey($key)) { return }
    $w = $WorkMap[$key]
    $result = Get-IQModelMember -Object $Entry -Name 'Result'
    $bim = [string]$w.BimPath
    if (Test-IQModelFileHasContent -Path $bim) {
        $note = ''
        $exit = Get-IQModelMember -Object $result -Name 'ExitCode'
        if ($null -ne $exit -and [int]$exit -ne 0) { $note = ('Tabular Editor exit code ' + $exit + ' but the .bim was written') }
        $size = 0
        try { $size = (Get-Item -LiteralPath $bim).Length } catch { $size = 0 }
        Set-IQItemDone -Stage 'ModelBackup' -ItemKey $key -Item $w.Item -Outputs @($bim) -Method 'XMLA' -Message $note -Data @{ BaseName = $w.BaseName; BimPath = $bim; WorkspaceId = $w.WorkspaceId; DatasetId = $w.DatasetId; SizeBytes = $size } | Out-Null
        Write-IQLog -Level Success -Stage 'ModelBackup' -Item $w.Item -Message ("Exported {0} ({1:N0} bytes)" -f $bim, $size)
        return
    }
    $message = 'XMLA export produced no .bim: ' + (Get-IQModelResultSummary -Result $result)
    if ($DeferFailure) {
        $w['XmlaFailure'] = $message
        Write-IQLog -Level Warn -Stage 'ModelBackup' -Item $w.Item -Message ($message + '; trying the Fabric getDefinition fallback')
        return
    }
    Set-IQItemDone -Stage 'ModelBackup' -ItemKey $key -Item $w.Item -Status Failed -Method 'XMLA' -Message $message -Data @{ BaseName = $w.BaseName; WorkspaceId = $w.WorkspaceId; DatasetId = $w.DatasetId } | Out-Null
}

function New-IQModelFabricState {
    <#
    .SYNOPSIS
    Decides once per stage whether the Fabric semantic-model getDefinition fallback can be attempted (private).
    .DESCRIPTION
    Returns @{ Enabled; Reason; ConsecutiveFailures; MaxConsecutiveFailures }. Enabled when a Fabric token can be minted
    (Get-IQToken -Resource Fabric returns a value; the Auth module returns $null when the provider cannot). Environments whose
    Fabric endpoint is unverified (GCC, GCC High, DoD) are still attempted - the first failures disable the fallback for the
    rest of the stage (circuit breaker) so a tenant without Fabric costs at most a few calls.
    #>
    [CmdletBinding()]
    param()
    $state = @{ Enabled = $false; Reason = ''; ConsecutiveFailures = 0; MaxConsecutiveFailures = 3 }
    $token = $null
    try { $token = Get-IQToken -Resource Fabric }
    catch { $state.Reason = 'Fabric token unavailable: ' + $_.Exception.Message; return $state }
    if ([string]::IsNullOrWhiteSpace([string]$token)) { $state.Reason = 'Fabric token unavailable for this sign-in mode'; return $state }
    try {
        if ($script:IQ.Endpoints -and $script:IQ.Endpoints.ContainsKey('FabricVerified') -and -not [bool]$script:IQ.Endpoints.FabricVerified) {
            Write-IQLog -Level Debug -Stage 'ModelBackup' -Message ("Fabric endpoint {0} is unverified for environment {1}; getDefinition will be attempted and disabled after {2} consecutive failures" -f $script:IQ.Endpoints.FabricApiPrefix, $script:IQ.Environment, $state.MaxConsecutiveFailures)
        }
    }
    catch { $null = $null }
    $state.Enabled = $true
    return $state
}

function Save-IQModelDefinitionFromFabric {
    <#
    .SYNOPSIS
    Downloads a semantic model's TMSL definition through the Fabric API and saves its model.bim part as the backup .bim (private).
    .DESCRIPTION
    POST workspaces/{ws}/semanticModels/{id}/getDefinition?format=TMSL via Invoke-IQFabricLro (202 polling handled there),
    decodes the InlineBase64 "model.bim" part, checks it is TMSL JSON with a "model" object and writes it atomically to
    Work.BimPath. Needs read+write permission on the model (workspace Contributor+); blocked for encrypted sensitivity
    labels. Never throws; returns @{ Success; BimPath; Message; SizeBytes }. Updates the circuit-breaker counters in -FabricState.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Work,
        [Parameter(Mandatory = $true)][hashtable]$FabricState,
        [Parameter(Mandatory = $false)][string]$Stage = 'ModelBackup'
    )
    $result = @{ Success = $false; BimPath = [string]$Work.BimPath; Message = ''; SizeBytes = 0 }
    if (-not $FabricState.Enabled) { $result.Message = 'Fabric getDefinition not available: ' + $FabricState.Reason; return $result }
    if (-not (Test-IQModelGuid -Value $Work.WorkspaceId)) { $result.Message = 'Fabric getDefinition needs a real workspace id (pseudo workspace)'; return $result }
    $timeout = 10
    try { $timeout = [int](Get-IQModelOption -Name 'DefinitionTimeoutMinutes' -Default 10) } catch { $timeout = 10 }
    if ($timeout -lt 1) { $timeout = 1 }
    $path = 'workspaces/' + $Work.WorkspaceId + '/semanticModels/' + $Work.DatasetId + '/getDefinition'
    Write-IQLog -Level Info -Stage $Stage -Item $Work.Item -Message 'Requesting the TMSL definition via Fabric getDefinition'
    $definition = $null
    try { $definition = Invoke-IQFabricLro -Method POST -Path $path -Query @{ format = 'TMSL' } -TimeoutMinutes $timeout -Stage $Stage }
    catch {
        $result.Message = 'Fabric getDefinition failed: ' + $_.Exception.Message
        Write-IQLog -Level Warn -Stage $Stage -Item $Work.Item -Message $result.Message
    }
    if ($null -eq $definition -and $result.Message -eq '') { $result.Message = 'Fabric getDefinition returned no definition (see previous warning; needs read+write on the model, unsupported for encrypted labels / this workspace type)' }
    if ($null -eq $definition) {
        $FabricState.ConsecutiveFailures = [int]$FabricState.ConsecutiveFailures + 1
        if ($FabricState.ConsecutiveFailures -ge $FabricState.MaxConsecutiveFailures) {
            $FabricState.Enabled = $false
            $FabricState.Reason = ('disabled after {0} consecutive getDefinition failures (last: {1})' -f $FabricState.ConsecutiveFailures, $result.Message)
            Write-IQLog -Level Warn -Stage $Stage -Message ('Fabric getDefinition fallback ' + $FabricState.Reason)
        }
        return $result
    }
    try {
        $parts = @()
        $def = Get-IQModelMember -Object $definition -Name 'definition'
        if ($null -ne $def) { $parts = @(Get-IQModelMember -Object $def -Name 'parts') }
        $parts = @($parts | Where-Object { $null -ne $_ })
        $bimPart = $null
        foreach ($p in $parts) {
            $partPath = [string](Get-IQModelMember -Object $p -Name 'path')
            if ($partPath -match '(?i)(^|[\\/])model\.bim$') { $bimPart = $p; break }
        }
        if ($null -eq $bimPart) { foreach ($p in $parts) { if ([string](Get-IQModelMember -Object $p -Name 'path') -match '(?i)\.bim$') { $bimPart = $p; break } } }
        if ($null -eq $bimPart) { throw ('no model.bim part in the definition (' + $parts.Count + ' parts: ' + (@($parts | ForEach-Object { [string](Get-IQModelMember -Object $_ -Name 'path') }) -join ', ') + ')') }
        $payloadType = [string](Get-IQModelMember -Object $bimPart -Name 'payloadType')
        $payload = [string](Get-IQModelMember -Object $bimPart -Name 'payload')
        $bytes = $null
        if ($payloadType -ieq 'InlineBase64' -or [string]::IsNullOrEmpty($payloadType)) { $bytes = [System.Convert]::FromBase64String($payload) }
        else { $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload) }
        if ($null -eq $bytes -or $bytes.Length -eq 0) { throw 'model.bim part is empty' }
        $text = [System.Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF)
        if ($text -notmatch '"model"\s*:') { throw 'model.bim part is not TMSL JSON (no "model" object)' }
        $folder = Split-Path -Path $result.BimPath -Parent
        if ($folder -and -not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        $tmp = $result.BimPath + '.tmp'
        [System.IO.File]::WriteAllBytes($tmp, $bytes)
        Move-Item -LiteralPath $tmp -Destination $result.BimPath -Force
        $result.SizeBytes = $bytes.Length
        $result.Success = $true
        $result.Message = ('TMSL definition exported via Fabric getDefinition ({0:N0} bytes)' -f $bytes.Length)
        $FabricState.ConsecutiveFailures = 0
        Write-IQLog -Level Success -Stage $Stage -Item $Work.Item -Message ("Saved {0} ({1:N0} bytes) from Fabric getDefinition" -f $result.BimPath, $bytes.Length)
    }
    catch {
        $result.Message = 'Fabric getDefinition definition could not be saved: ' + $_.Exception.Message
        Write-IQLog -Level Warn -Stage $Stage -Item $Work.Item -Message $result.Message
    }
    return $result
}

function Complete-IQModelBackupViaFabric {
    <#
    .SYNOPSIS
    Runs the Fabric getDefinition fallback for one dataset and writes its ModelBackup checkpoint (private).
    .DESCRIPTION
    Success: Succeeded with -Method FabricDefinition and the .bim as output. Failure: the -FallbackStatus (Failed for
    dedicated models whose XMLA export was impossible or failed, Skipped for Pro models that ReportBackup covers) with
    -FallbackMessage plus the Fabric reason. Returns the checkpoint status string.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Work,
        [Parameter(Mandatory = $true)][hashtable]$FabricState,
        [Parameter(Mandatory = $true)][ValidateSet('Failed', 'Skipped')][string]$FallbackStatus,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$FallbackMessage,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$FallbackMethod,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Prefix
    )
    $stage = 'ModelBackup'
    $data = @{ BaseName = $Work.BaseName; WorkspaceId = $Work.WorkspaceId; DatasetId = $Work.DatasetId }
    $fabric = Save-IQModelDefinitionFromFabric -Work $Work -FabricState $FabricState -Stage $stage
    if ($fabric.Success) {
        $message = $fabric.Message
        if ($Prefix) { $message = $Prefix + '; ' + $message }
        $data['BimPath'] = $fabric.BimPath
        $data['SizeBytes'] = $fabric.SizeBytes
        Set-IQItemDone -Stage $stage -ItemKey $Work.Key -Item $Work.Item -Outputs @($fabric.BimPath) -Method 'FabricDefinition' -Message $message -Data $data | Out-Null
        return 'Succeeded'
    }
    $message = $FallbackMessage
    if ($fabric.Message) { $message = $message + '; ' + $fabric.Message }
    if ($FallbackStatus -eq 'Failed') { $message = $message + '; no .bim exported (ModelDetail can still use DAX)' }
    Set-IQItemDone -Stage $stage -ItemKey $Work.Key -Item $Work.Item -Status $FallbackStatus -Method $FallbackMethod -Message $message -Data $data | Out-Null
    return $FallbackStatus
}

function Invoke-IQModelBackupStage {
    <#
    .SYNOPSIS
    ModelBackup stage body: Tabular Editor XMLA export of every dedicated-capacity dataset in scope (itemKey = DatasetId), with a Fabric getDefinition fallback.
    .DESCRIPTION
    Port of monolith lines 2545-2620. Connection string "Provider=MSOLAP;Data Source=<XmlaPrefix>/v1.0/myorg/<url-encoded
    workspace>;Password=<token>", positional database = dataset name, -S <per-job rename script> -B "<Model Backups>\<RunId>\
    <CleanWs> ~ <CleanModel>.bim". Jobs run through Invoke-IQProcessBatch in chunks of MaxParallelExtracts so each chunk's
    token is fetched immediately before its processes start (brief section 4.4). Every dataset is checkpointed the moment
    its process finishes (Method XMLA). When Tabular Editor is unavailable (non-Windows host, failed preflight) or an
    export fails, and a Fabric token is available, the TMSL definition is fetched with Fabric
    workspaces/{ws}/semanticModels/{id}/getDefinition?format=TMSL and its model.bim saved under the same name (Method
    FabricDefinition). Pro / My Workspace datasets are recorded as Skipped ("model extracted from PBIX in ReportBackup") -
    Pro datasets in real workspaces first try the Fabric definition too (best effort, unverified on shared capacity).
    Datasets already done in a resumed run are skipped. Returns @{ Total; Done; Skipped; Failed; AlreadyDone; ViaXmla; ViaFabric }.
    #>
    [CmdletBinding()]
    param()
    $stage = 'ModelBackup'
    $summary = @{ Total = 0; Done = 0; Skipped = 0; Failed = 0; AlreadyDone = 0; ViaXmla = 0; ViaFabric = 0; BudgetStop = $false }
    $runFolder = Get-IQModelRunFolder
    $work = @(Get-IQModelWorkList -RunFolder $runFolder)
    $summary.Total = $work.Count
    Write-IQLog -Level Info -Stage $stage -Message ("Model backup: {0} dataset(s) in scope; folder {1}" -f $work.Count, $runFolder)
    if ($work.Count -eq 0) { return $summary }

    $teAvailable = Test-IQModelTabularEditorAvailable
    $teReason = $null
    if (-not $teAvailable) { $teReason = Get-IQModelTabularEditorReason }
    $fabricState = New-IQModelFabricState
    if ($fabricState.Enabled) { Write-IQLog -Level Info -Stage $stage -Message 'Fabric getDefinition fallback available for models Tabular Editor cannot export' }
    else { Write-IQLog -Level Debug -Stage $stage -Message ('Fabric getDefinition fallback not available: ' + $fabricState.Reason) }
    $pending = New-Object System.Collections.Generic.List[object]
    $proMessage = 'Pro workspace - model extracted from PBIX in ReportBackup'

    $budgetStop = $false
    foreach ($w in $work) {
        if (Test-IQItemDone -Stage $stage -ItemKey $w.Key) {
            $summary.AlreadyDone++
            Write-IQLog -Level Debug -Stage $stage -Item $w.Item -Message 'Already done (checkpoint); skipping'
            continue
        }
        if (Test-IQTimeBudget -Stage $stage -Item $w.Item) { $budgetStop = $true; break }
        if ($w.NoAccess) {
            Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $w.Item -Status Skipped -Message 'No workspace access (shared report) - model cannot be exported' | Out-Null
            $summary.Skipped++
            continue
        }
        # Resume: a complete .bim from an earlier attempt of this run (crash before its checkpoint) is reused (audit X1-09).
        if ($script:IQ.IsResume -and (Test-IQModelFileHasContent -Path $w.BimPath)) {
            Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $w.Item -Outputs @($w.BimPath) -Method 'XMLA' -Message 'Existing .bim from an earlier attempt of this run reused' -Data @{ BaseName = $w.BaseName; BimPath = $w.BimPath; WorkspaceId = $w.WorkspaceId; DatasetId = $w.DatasetId } | Out-Null
            $summary.Done++
            continue
        }
        if (-not $w.IsDedicated) {
            if ($fabricState.Enabled -and -not $w.IsPseudoWorkspace) {
                $status = Complete-IQModelBackupViaFabric -Work $w -FabricState $fabricState -FallbackStatus Skipped -FallbackMessage $proMessage -FallbackMethod 'PBIX' -Prefix 'Pro workspace'
                if ($status -eq 'Succeeded') { $summary.Done++; $summary.ViaFabric++ } else { $summary.Skipped++ }
            }
            else {
                Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $w.Item -Status Skipped -Method 'PBIX' -Message $proMessage -Data @{ BaseName = $w.BaseName; WorkspaceId = $w.WorkspaceId; DatasetId = $w.DatasetId } | Out-Null
                $summary.Skipped++
            }
            continue
        }
        $blocker = $null
        if (-not $teAvailable) { $blocker = $teReason }
        elseif ($w.WorkspaceName -match '["\;]' -or $w.DatasetName -match '"') { $blocker = 'Workspace or dataset name contains a quote or semicolon that cannot be passed in the XMLA connection string' }
        if ($blocker) {
            if ($fabricState.Enabled) {
                $status = Complete-IQModelBackupViaFabric -Work $w -FabricState $fabricState -FallbackStatus Failed -FallbackMessage $blocker -FallbackMethod 'XMLA' -Prefix ('XMLA export not possible (' + $blocker + ')')
                if ($status -eq 'Succeeded') { $summary.Done++; $summary.ViaFabric++ } else { $summary.Failed++ }
            }
            else {
                Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $w.Item -Status Failed -Method 'XMLA' -Message ($blocker + '; no .bim exported (ModelDetail can still use DAX)') -Data @{ BaseName = $w.BaseName; WorkspaceId = $w.WorkspaceId; DatasetId = $w.DatasetId } | Out-Null
                $summary.Failed++
            }
            continue
        }
        $pending.Add($w)
    }
    if ($budgetStop) {
        Write-IQLog -Level Warn -Stage $stage -Message ("Time budget reached: model backup stopped after {0} export(s); the remaining models are exported on the next start." -f ($summary.Done + $summary.ViaFabric))
        $summary.BudgetStop = $true
        return $summary
    }
    if ($pending.Count -eq 0) {
        Write-IQLog -Level Info -Stage $stage -Message ("Model backup: nothing to export via XMLA ({0} already done, {1} via Fabric, {2} skipped, {3} failed)" -f $summary.AlreadyDone, $summary.ViaFabric, $summary.Skipped, $summary.Failed)
        return $summary
    }

    $maxParallel = 2
    try { $maxParallel = [int](Get-IQModelOption -Name 'MaxParallelExtracts' -Default 2) } catch { $maxParallel = 2 }
    if ($maxParallel -lt 1) { $maxParallel = 1 }
    $timeout = 20
    try { $timeout = [int](Get-IQModelOption -Name 'ToolTimeoutMinutes' -Default 20) } catch { $timeout = 20 }
    if ($timeout -lt 1) { $timeout = 1 }
    $tePath = [string]$script:IQ.Tools.TabularEditorPath
    $xmlaPrefix = [string]$script:IQ.Endpoints.XmlaPrefix
    if ([string]::IsNullOrWhiteSpace($xmlaPrefix)) { throw 'Endpoints.XmlaPrefix is not set; call Set-IQEnvironment / Initialize-IQContext with an Environment first.' }
    Write-IQLog -Level Info -Stage $stage -Message ("Exporting {0} model(s) via Tabular Editor XMLA, {1} in parallel, {2} min timeout each" -f $pending.Count, $maxParallel, $timeout)

    $workMap = @{}
    foreach ($w in $pending) { $workMap[[string]$w.Key] = $w }
    $deferFailure = [bool]$fabricState.Enabled
    $onDone = { param($iqEntry) Complete-IQModelBackupJob -Entry $iqEntry -WorkMap $workMap -DeferFailure:$deferFailure }.GetNewClosure()

    $index = 0
    while ($index -lt $pending.Count) {
        if (Test-IQTimeBudget -Stage $stage) {
            Write-IQLog -Level Warn -Stage $stage -Message ("Time budget reached: {0} model(s) not exported via XMLA yet - they are exported on the next start." -f ($pending.Count - $index))
            $summary.BudgetStop = $true
            break
        }
        $chunk = @()
        $last = [Math]::Min($index + $maxParallel, $pending.Count) - 1
        for ($i = $index; $i -le $last; $i++) { $chunk += $pending[$i] }
        $index = $last + 1

        $jobs = @()
        $scripts = @()
        foreach ($w in $chunk) {
            $token = $null
            try { $token = Get-IQToken -Resource PowerBI }
            catch {
                Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $w.Item -Status Failed -Method 'XMLA' -Message ('Could not obtain a Power BI token: ' + $_.Exception.Message) | Out-Null
                $summary.Failed++
                continue
            }
            if ([string]::IsNullOrWhiteSpace($token)) {
                Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $w.Item -Status Failed -Method 'XMLA' -Message 'No Power BI token available' | Out-Null
                $summary.Failed++
                continue
            }
            $scriptPath = New-IQModelRenameScript -Key $w.Key -BaseName $w.BaseName
            $scripts += $scriptPath
            if (Test-Path -LiteralPath $w.BimPath) { Remove-Item -LiteralPath $w.BimPath -Force -ErrorAction SilentlyContinue }
            $encodedWorkspace = [System.Uri]::EscapeDataString([string]$w.WorkspaceName)
            $dataSource = ('{0}/v1.0/myorg/{1}' -f $xmlaPrefix.TrimEnd('/'), $encodedWorkspace)
            $arguments = ('"Provider=MSOLAP;Data Source={0};Password={1}" "{2}" -S "{3}" -B "{4}"' -f $dataSource, $token, $w.DatasetName, $scriptPath, $w.BimPath)
            Write-IQLog -Level Info -Stage $stage -Item $w.Item -Message ('Exporting ' + $w.BaseName)
            $jobs += @{ ItemKey = $w.Key; Item = $w.Item; FilePath = $tePath; ArgumentList = $arguments; WorkingDirectory = [string]$script:IQ.BaseFolder; LogName = ('xmla-' + (Get-IQSafeKey -Value $w.Key)) }
        }
        if ($jobs.Count -gt 0) {
            $results = Invoke-IQProcessBatch -Jobs $jobs -MaxParallel $maxParallel -TimeoutMinutes $timeout -Stage $stage -OnJobComplete $onDone
            $results = ConvertTo-IQModelResultList -Results $results
            foreach ($r in $results) {
                $key = [string](Get-IQModelMember -Object $r -Name 'ItemKey')
                if (-not (Get-IQItemCheckpoint -Stage $stage -ItemKey $key)) {
                    $w = $null
                    if ($workMap.ContainsKey($key)) { $w = $workMap[$key] }
                    if ($null -ne $w -and $w.ContainsKey('XmlaFailure') -and $w.XmlaFailure) {
                        # Deferred XMLA failure: Fabric getDefinition fallback, then one combined checkpoint.
                        $status = Complete-IQModelBackupViaFabric -Work $w -FabricState $fabricState -FallbackStatus Failed -FallbackMessage ([string]$w.XmlaFailure) -FallbackMethod 'XMLA' -Prefix ([string]$w.XmlaFailure)
                        if ($status -eq 'Succeeded') { $summary.ViaFabric++ }
                    }
                    else {
                        # Safety net: a job the callback did not checkpoint (callback error) is checkpointed here.
                        Complete-IQModelBackupJob -Entry $r -WorkMap $workMap
                    }
                }
                $cp = Get-IQItemCheckpoint -Stage $stage -ItemKey $key
                if ($cp -and [string](Get-IQModelMember -Object $cp -Name 'status') -eq 'Succeeded') {
                    $summary.Done++
                    if ([string](Get-IQModelMember -Object $cp -Name 'method') -eq 'XMLA') { $summary.ViaXmla++ }
                }
                else { $summary.Failed++ }
            }
        }
        foreach ($s in $scripts) { try { Remove-Item -LiteralPath $s -Force -ErrorAction SilentlyContinue } catch { $null = $null } }
    }
    Write-IQLog -Level Info -Stage $stage -Message ("Model backup finished: {0} exported ({1} XMLA, {2} Fabric), {3} skipped, {4} failed, {5} already done" -f $summary.Done, $summary.ViaXmla, $summary.ViaFabric, $summary.Skipped, $summary.Failed, $summary.AlreadyDone)
    return $summary
}

function Get-IQModelBimPath {
    <#
    .SYNOPSIS
    Finds the .bim for a dataset: the run folder, the ModelBackup checkpoint outputs, then any ReportBackup checkpoint for the dataset (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Work)
    if (Test-IQModelFileHasContent -Path $Work.BimPath) { return [string]$Work.BimPath }
    $cp = Get-IQItemCheckpoint -Stage 'ModelBackup' -ItemKey $Work.Key
    if ($cp) {
        foreach ($o in @(Get-IQModelMember -Object $cp -Name 'outputs')) {
            if ([string]$o -like '*.bim' -and (Test-IQModelFileHasContent -Path ([string]$o))) { return [string]$o }
        }
    }
    # Pro models: the ReportBackup stage writes the .bim (pbi-tools) and records the dataset in its checkpoint data.
    try {
        $doneRoot = $null
        if ($script:IQ.ContainsKey('RunPaths') -and $script:IQ.RunPaths -and $script:IQ.RunPaths.Done) { $doneRoot = [string]$script:IQ.RunPaths.Done }
        elseif ($script:IQ.RunPath) { $doneRoot = Join-Path ([string]$script:IQ.RunPath) 'done' }
        if ($doneRoot) {
            $folder = Join-Path $doneRoot 'ReportBackup'
            if (Test-Path -LiteralPath $folder) {
                foreach ($file in @(Get-ChildItem -LiteralPath $folder -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
                    $rc = $null
                    try { $rc = ConvertFrom-IQJsonFile -Path $file.FullName } catch { $rc = $null }
                    if ($null -eq $rc) { continue }
                    $data = Get-IQModelMember -Object $rc -Name 'data'
                    $dsId = [string](Get-IQModelMember -Object $data -Name 'DatasetId')
                    if ($dsId -ne '' -and $dsId -ieq $Work.DatasetId) {
                        foreach ($o in @(Get-IQModelMember -Object $rc -Name 'outputs')) {
                            if ([string]$o -like '*.bim' -and (Test-IQModelFileHasContent -Path ([string]$o))) { return [string]$o }
                        }
                        $bp = [string](Get-IQModelMember -Object $data -Name 'BimPath')
                        if ($bp -and (Test-IQModelFileHasContent -Path $bp)) { return $bp }
                    }
                }
            }
        }
    }
    catch { Write-IQLog -Level Debug -Stage 'ModelDetail' -Item $Work.Item -Message ("ReportBackup checkpoint scan failed: " + $_.Exception.Message) }
    return $null
}

function Get-IQModelDatabaseNameFromBim {
    <#
    .SYNOPSIS
    Reads the database "name" of a .bim file (what the csx scripts use for the CSV file name); $null on failure (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$BimPath)
    try {
        $json = [System.IO.File]::ReadAllText($BimPath)
        $obj = $json | ConvertFrom-Json -ErrorAction Stop
        $name = Get-IQModelMember -Object $obj -Name 'name'
        if ($null -ne $name -and [string]$name -ne '') { return [string]$name }
    }
    catch { return $null }
    return $null
}

function Find-IQModelDetailCsv {
    <#
    .SYNOPSIS
    Locates the CSVs the csx scripts wrote for a model and moves them into the run folder when they landed elsewhere (private).
    .DESCRIPTION
    The csx scripts write "<Model.Database.Name>.csv" / "_MD.csv" into the newest yyyy-MM-dd folder under "Model Backups"
    relative to the working directory (BaseFolder), which is the run folder when RunId is today's date. Candidates checked:
    the run folder, Get-IQDateFolder, and "Model Backups\<today>"; names checked: the base name and the .bim database name.
    Returns @{ Csv; Md } with $null for anything not found.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Work,
        [Parameter(Mandatory = $true)][string]$RunFolder,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$BimPath
    )
    $names = @([string]$Work.BaseName)
    if ($BimPath) {
        $dbName = Get-IQModelDatabaseNameFromBim -BimPath $BimPath
        if ($dbName -and $names -notcontains $dbName) { $names += $dbName }
    }
    $folders = @($RunFolder)
    $latest = $null
    try { $latest = Get-IQDateFolder -Root ([string]$script:IQ.Paths.ModelBackups) } catch { $latest = $null }
    if ($latest -and $folders -notcontains $latest) { $folders += $latest }
    $today = Join-Path ([string]$script:IQ.Paths.ModelBackups) (Get-Date -Format 'yyyy-MM-dd')
    if ($folders -notcontains $today) { $folders += $today }

    $found = @{ Csv = $null; Md = $null }
    foreach ($suffix in @('.csv', '_MD.csv')) {
        $slot = 'Csv'
        if ($suffix -eq '_MD.csv') { $slot = 'Md' }
        $target = Join-Path $RunFolder ($Work.BaseName + $suffix)
        foreach ($folder in $folders) {
            if (-not (Test-Path -LiteralPath $folder)) { continue }
            foreach ($n in $names) {
                $candidate = Join-Path $folder ($n + $suffix)
                if (-not (Test-IQModelFileHasContent -Path $candidate)) { continue }
                if ($candidate -ne $target) {
                    try {
                        Move-Item -LiteralPath $candidate -Destination $target -Force
                        Write-IQLog -Level Debug -Stage 'ModelDetail' -Item $Work.Item -Message ("Moved {0} -> {1}" -f $candidate, $target)
                    }
                    catch {
                        Write-IQLog -Level Warn -Stage 'ModelDetail' -Item $Work.Item -Message ("Could not move {0} into the run folder: {1}" -f $candidate, $_.Exception.Message)
                        $target = $candidate
                    }
                }
                $found[$slot] = $target
                break
            }
            if ($found[$slot]) { break }
        }
    }
    return $found
}

function Invoke-IQModelDetailTabularEditor {
    <#
    .SYNOPSIS
    Runs the two csx scripts for a set of models through Invoke-IQProcessBatch and returns per-key outcomes (private).
    .DESCRIPTION
    Two jobs per model ("<bim>" -S "Model Detail Extract Script.csx" and "<bim>" -S "Measure Dependency Extract Script.csx"),
    WorkingDirectory = BaseFolder (the scripts locate "Model Backups\<latest date>" from the CWD). Returns a hashtable
    keyed by dataset id: @{ Success; Csv; Md; Message }.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Candidates,
        [Parameter(Mandatory = $true)][string]$RunFolder
    )
    $stage = 'ModelDetail'
    $outcomes = @{}
    if ($Candidates.Count -eq 0) { return $outcomes }
    $configFolder = [string]$script:IQ.ConfigFolder
    $script3 = Join-Path $configFolder 'Model Detail Extract Script.csx'
    $script4 = Join-Path $configFolder 'Measure Dependency Extract Script.csx'
    foreach ($s in @($script3, $script4)) {
        if (-not (Test-Path -LiteralPath $s)) {
            foreach ($c in $Candidates) { $outcomes[[string]$c.Work.Key] = @{ Success = $false; Csv = $null; Md = $null; Message = "Extract script missing: $s" } }
            Write-IQLog -Level Error -Stage $stage -Message "Extract script missing: $s"
            return $outcomes
        }
    }
    $maxParallel = 2
    try { $maxParallel = [int](Get-IQModelOption -Name 'MaxParallelExtracts' -Default 2) } catch { $maxParallel = 2 }
    if ($maxParallel -lt 1) { $maxParallel = 1 }
    $timeout = 20
    try { $timeout = [int](Get-IQModelOption -Name 'ToolTimeoutMinutes' -Default 20) } catch { $timeout = 20 }
    if ($timeout -lt 1) { $timeout = 1 }
    $tePath = [string]$script:IQ.Tools.TabularEditorPath

    $jobs = @()
    foreach ($c in $Candidates) {
        $w = $c.Work
        $bim = [string]$c.BimPath
        $safe = Get-IQSafeKey -Value $w.Key
        # Remove stale outputs so a leftover from another attempt is never mistaken for this run's result.
        foreach ($stale in @($w.CsvPath, $w.MdPath)) { if (Test-Path -LiteralPath $stale) { Remove-Item -LiteralPath $stale -Force -ErrorAction SilentlyContinue } }
        $jobs += @{ ItemKey = ($w.Key + '|detail'); Item = $w.Item; FilePath = $tePath; ArgumentList = ('"{0}" -S "{1}"' -f $bim, $script3); WorkingDirectory = [string]$script:IQ.BaseFolder; LogName = ('detail-' + $safe) }
        $jobs += @{ ItemKey = ($w.Key + '|md'); Item = $w.Item; FilePath = $tePath; ArgumentList = ('"{0}" -S "{1}"' -f $bim, $script4); WorkingDirectory = [string]$script:IQ.BaseFolder; LogName = ('md-' + $safe) }
        Write-IQLog -Level Info -Stage $stage -Item $w.Item -Message ('Extracting model detail via Tabular Editor from ' + $bim)
    }
    Write-IQLog -Level Info -Stage $stage -Message ("Running {0} Tabular Editor script job(s) for {1} model(s), {2} in parallel, {3} min timeout each" -f $jobs.Count, $Candidates.Count, $maxParallel, $timeout)
    # The csx scripts honour IMPACTIQ_BASE / IMPACTIQ_DATE_FOLDER / IMPACTIQ_REPORT_DATE (audit X2-H1) so they write into
    # THIS run's folder instead of the newest dated folder; the CWD-scanning behaviour stays their fallback.
    $previousEnv = @{ IMPACTIQ_BASE = $env:IMPACTIQ_BASE; IMPACTIQ_DATE_FOLDER = $env:IMPACTIQ_DATE_FOLDER; IMPACTIQ_REPORT_DATE = $env:IMPACTIQ_REPORT_DATE }
    $results = $null
    try {
        $env:IMPACTIQ_BASE = [string]$script:IQ.BaseFolder
        $env:IMPACTIQ_DATE_FOLDER = [string]$RunFolder
        $env:IMPACTIQ_REPORT_DATE = [string](Get-IQDaxModelAsOfDate)
        $results = Invoke-IQProcessBatch -Jobs $jobs -MaxParallel $maxParallel -TimeoutMinutes $timeout -Stage $stage
    }
    finally {
        foreach ($name in @($previousEnv.Keys)) { Set-Item -Path ('Env:' + $name) -Value $previousEnv[$name] -ErrorAction SilentlyContinue }
    }
    $results = ConvertTo-IQModelResultList -Results $results
    $byKey = @{}
    foreach ($r in $results) { $byKey[[string](Get-IQModelMember -Object $r -Name 'ItemKey')] = Get-IQModelMember -Object $r -Name 'Result' }

    foreach ($c in $Candidates) {
        $w = $c.Work
        $detailResult = $byKey[($w.Key + '|detail')]
        $mdResult = $byKey[($w.Key + '|md')]
        $files = Find-IQModelDetailCsv -Work $w -RunFolder $RunFolder -BimPath $c.BimPath
        $problems = @()
        if (-not $files.Csv) { $problems += ('no <model>.csv produced (' + (Get-IQModelResultSummary -Result $detailResult) + ')') }
        if (-not $files.Md) { $problems += ('no <model>_MD.csv produced (' + (Get-IQModelResultSummary -Result $mdResult) + ')') }
        if ($problems.Count -eq 0) {
            $outcomes[[string]$w.Key] = @{ Success = $true; Csv = $files.Csv; Md = $files.Md; Message = 'Model detail extracted via Tabular Editor scripts' }
        }
        else {
            $outcomes[[string]$w.Key] = @{ Success = $false; Csv = $files.Csv; Md = $files.Md; Message = ($problems -join '; ') }
        }
    }
    return $outcomes
}

function Get-IQModelDetailPlan {
    <#
    .SYNOPSIS
    Ordered list of extraction methods to try for one dataset given ModelDetailMethod, Tabular Editor availability and the .bim (private).
    .DESCRIPTION
    Auto / Both: TabularEditor (when TE2 works, a .bim exists and its database name equals the base name so the csx
    scripts write the expected file) -> Bim (when a .bim exists) -> Dax. TabularEditor: TE2 only. Bim: parser only.
    Dax: DAX only. Returns @{ Steps = @(...); Why = <reason TE2 / Bim were left out> }.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][bool]$TabularEditorAvailable,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$BimPath,
        [Parameter(Mandatory = $true)][string]$BaseName
    )
    $why = @()
    $hasBim = -not [string]::IsNullOrWhiteSpace($BimPath)
    $teEligible = $false
    if (-not $TabularEditorAvailable) { $why += (Get-IQModelTabularEditorReason) }
    elseif (-not $hasBim) { $why += 'no .bim available for this dataset' }
    else {
        $dbName = Get-IQModelDatabaseNameFromBim -BimPath $BimPath
        if ($null -ne $dbName -and $dbName -ne $BaseName) { $why += ("the .bim database name '{0}' differs from the file name (Fabric/pbi-tools export); the csx scripts would misname the CSV, using the built-in parser" -f $dbName) }
        else { $teEligible = $true }
    }
    $steps = @()
    switch ($Method) {
        'TabularEditor' { if ($teEligible) { $steps = @('TabularEditor') } }
        'Dax' { $steps = @('Dax') }
        'Bim' { if ($hasBim) { $steps = @('Bim') } elseif ($why -notcontains 'no .bim available for this dataset') { $why += 'no .bim available for this dataset' } }
        default {
            if ($teEligible) { $steps += 'TabularEditor' }
            if ($hasBim) { $steps += 'Bim' }
            $steps += 'Dax'
        }
    }
    return @{ Steps = @($steps); Why = ($why -join '; ') }
}

function Invoke-IQModelDetailStage {
    <#
    .SYNOPSIS
    ModelDetail stage body: "<CleanWs> ~ <CleanModel>.csv" and "_MD.csv" per dataset in scope (itemKey = DatasetId).
    .DESCRIPTION
    Options.ModelDetailMethod: TabularEditor = the two csx scripts against the .bim (batch, parallel; no .bim / no TE2 =
    Failed); Bim = the built-in TMSL parser (Export-IQModelDetailFromBim) against the .bim (no .bim = Failed); Dax =
    Get-IQModelDetailViaDax (INFO.VIEW.* over executeQueries); Auto = Tabular Editor when it works and a .bim it can name
    correctly exists, else the Bim parser when any .bim exists, else DAX; Both = Tabular Editor first and, when it fails or
    cannot run, the Bim parser then DAX. The method that produced the files is recorded in the checkpoint
    (-Method TabularEditor|Bim|Dax) with the two CSV paths as outputs; the message keeps the reasons earlier methods were
    not used. Returns @{ Total; Done; Failed; Skipped; AlreadyDone; ViaTabularEditor; ViaBim; ViaDax }.
    #>
    [CmdletBinding()]
    param()
    $stage = 'ModelDetail'
    $summary = @{ Total = 0; Done = 0; Failed = 0; Skipped = 0; AlreadyDone = 0; ViaTabularEditor = 0; ViaBim = 0; ViaDax = 0; BudgetStop = $false }
    $runFolder = Get-IQModelRunFolder
    $method = [string](Get-IQModelOption -Name 'ModelDetailMethod' -Default 'Auto')
    if ($method -notin @('Auto', 'TabularEditor', 'Dax', 'Both', 'Bim')) {
        Write-IQLog -Level Warn -Stage $stage -Message "Unknown ModelDetailMethod '$method'; using Auto"
        $method = 'Auto'
    }
    $work = @(Get-IQModelWorkList -RunFolder $runFolder)
    $summary.Total = $work.Count
    $teAvailable = Test-IQModelTabularEditorAvailable
    Write-IQLog -Level Info -Stage $stage -Message ("Model detail: {0} dataset(s) in scope, method {1}, Tabular Editor available: {2}" -f $work.Count, $method, $teAvailable)
    if ($work.Count -eq 0) { return $summary }
    if (-not $teAvailable -and $method -notin @('Dax', 'Bim')) { Write-IQLog -Level Info -Stage $stage -Message (Get-IQModelTabularEditorReason) }
    $asOfDate = Get-IQDaxModelAsOfDate

    $plans = @{}
    $bimByKey = @{}
    $noteByKey = @{}
    $teCandidates = @()
    $budgetStop = $false
    foreach ($w in $work) {
        if (Test-IQItemDone -Stage $stage -ItemKey $w.Key) {
            $summary.AlreadyDone++
            Write-IQLog -Level Debug -Stage $stage -Item $w.Item -Message 'Already done (checkpoint); skipping'
            continue
        }
        if (Test-IQTimeBudget -Stage $stage -Item $w.Item) { $budgetStop = $true; break }
        if ($w.NoAccess) {
            Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $w.Item -Status Skipped -Message 'No workspace access (shared report) - model detail unavailable' | Out-Null
            $summary.Skipped++
            continue
        }
        $bim = $null
        if ($method -ne 'Dax') { $bim = Get-IQModelBimPath -Work $w }
        $plan = Get-IQModelDetailPlan -Method $method -TabularEditorAvailable $teAvailable -BimPath $bim -BaseName $w.BaseName
        if ($plan.Steps.Count -eq 0) {
            $failMethod = $method
            $reason = $plan.Why
            if ([string]::IsNullOrWhiteSpace($reason)) { $reason = 'no extraction method available' }
            Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $w.Item -Status Failed -Method $failMethod -Message ($method + ' method requested but ' + $reason) -Data @{ BaseName = $w.BaseName; DatasetId = $w.DatasetId; WorkspaceId = $w.WorkspaceId } | Out-Null
            $summary.Failed++
            continue
        }
        $plans[[string]$w.Key] = New-Object System.Collections.Generic.List[string]
        foreach ($s in $plan.Steps) { $plans[[string]$w.Key].Add($s) }
        $bimByKey[[string]$w.Key] = $bim
        $noteByKey[[string]$w.Key] = New-Object System.Collections.Generic.List[string]
        if ($plan.Why -and $plan.Steps[0] -ne 'TabularEditor' -and $method -ne 'Dax' -and $method -ne 'Bim') { $noteByKey[[string]$w.Key].Add('Tabular Editor not used: ' + $plan.Why) }
        if ($plan.Steps[0] -eq 'TabularEditor') { $teCandidates += @{ Work = $w; BimPath = $bim } }
    }

    # 1. Tabular Editor batch (all csx jobs in one pool).
    $queue = New-Object System.Collections.Generic.List[object]
    if ($teCandidates.Count -gt 0 -and -not $budgetStop -and (Test-IQTimeBudget -Stage $stage)) { $budgetStop = $true }
    if ($teCandidates.Count -gt 0 -and -not $budgetStop) {
        $outcomes = Invoke-IQModelDetailTabularEditor -Candidates $teCandidates -RunFolder $runFolder
        foreach ($c in $teCandidates) {
            $w = $c.Work
            $key = [string]$w.Key
            $o = $outcomes[$key]
            if ($null -eq $o) { $o = @{ Success = $false; Csv = $null; Md = $null; Message = 'no outcome recorded' } }
            $plans[$key].RemoveAt(0)
            if ($o.Success) {
                $message = [string]$o.Message
                Set-IQItemDone -Stage $stage -ItemKey $key -Item $w.Item -Outputs @($o.Csv, $o.Md) -Method 'TabularEditor' -Message $message -Data @{ BaseName = $w.BaseName; BimPath = $c.BimPath; DatasetId = $w.DatasetId; WorkspaceId = $w.WorkspaceId } | Out-Null
                $summary.Done++
                $summary.ViaTabularEditor++
                Write-IQLog -Level Success -Stage $stage -Item $w.Item -Message ('Model detail extracted via Tabular Editor: ' + $o.Csv)
                $plans.Remove($key)
                continue
            }
            $noteByKey[$key].Add('Tabular Editor failed: ' + $o.Message)
            if ($plans[$key].Count -gt 0) {
                Write-IQLog -Level Warn -Stage $stage -Item $w.Item -Message ('Tabular Editor extraction failed (' + $o.Message + '); falling back to ' + $plans[$key][0])
            }
            else {
                Set-IQItemDone -Stage $stage -ItemKey $key -Item $w.Item -Status Failed -Method 'TabularEditor' -Message ($noteByKey[$key] -join '; ') -Data @{ BaseName = $w.BaseName; BimPath = $c.BimPath; DatasetId = $w.DatasetId; WorkspaceId = $w.WorkspaceId } | Out-Null
                $summary.Failed++
                $plans.Remove($key)
            }
        }
    }
    foreach ($w in $work) { if ($plans.ContainsKey([string]$w.Key)) { $queue.Add($w) } }

    # 2. Remaining steps per dataset, in plan order: Bim parser, then DAX.
    foreach ($w in $queue) {
        if ($budgetStop -or (Test-IQTimeBudget -Stage $stage -Item $w.Item)) { $budgetStop = $true; break }
        $key = [string]$w.Key
        $steps = $plans[$key]
        $notes = $noteByKey[$key]
        $done = $false
        while (-not $done -and $steps.Count -gt 0) {
            $step = $steps[0]
            $steps.RemoveAt(0)
            $r = $null
            $noteText = ''
            if ($notes.Count -gt 0) { $noteText = ' (' + ($notes -join '; ') + ')' }
            $modelId = $w.BaseName
            if ($w.IsDedicated) { $modelId = $w.DatasetId }
            switch ($step) {
                'Bim' {
                    Write-IQLog -Level Info -Stage $stage -Item $w.Item -Message ('Extracting model detail from the .bim with the built-in parser' + $noteText)
                    try { $r = Export-IQModelDetailFromBim -BimPath $bimByKey[$key] -OutputFolder $runFolder -ModelName $w.BaseName -ModelId $modelId -AsOfDate $asOfDate -Stage $stage }
                    catch { $r = @{ Success = $false; Message = ('Bim parser threw: ' + $_.Exception.Message); Outputs = @() } }
                }
                'Dax' {
                    Write-IQLog -Level Info -Stage $stage -Item $w.Item -Message ('Extracting model detail via DAX INFO.VIEW.*' + $noteText)
                    try { $r = Get-IQModelDetailViaDax -Dataset $w.Dataset -OutputFolder $runFolder -IsDedicated $w.IsDedicated -BaseName $w.BaseName -ModelAsOfDate $asOfDate -Stage $stage }
                    catch {
                        $r = @{ Success = $false; Message = ('DAX extraction threw: ' + $_.Exception.Message); Outputs = @() }
                        Write-IQLog -Level Error -Stage $stage -Item $w.Item -Message $r.Message -Exception $_.Exception
                    }
                }
                default { $r = @{ Success = $false; Message = ('unknown step ' + $step); Outputs = @() } }
            }
            if ($r -and $r.Success) {
                $message = [string]$r.Message
                if ($notes.Count -gt 0) { $message = $message + ' [' + ($notes -join '; ') + ']' }
                $data = @{ BaseName = $w.BaseName; DatasetId = $w.DatasetId; WorkspaceId = $w.WorkspaceId; RowCount = $r.RowCount; DependencyRowCount = $r.DependencyRowCount }
                if ($step -eq 'Bim') { $data['BimPath'] = $bimByKey[$key] }
                if ($step -eq 'Dax' -and $r.ContainsKey('Unavailable')) { $data['UnavailableViaRest'] = @($r.Unavailable); $data['DependencySource'] = $r.DependencySource; $data['ExpressionsMasked'] = $r.ExpressionsMasked }
                Set-IQItemDone -Stage $stage -ItemKey $key -Item $w.Item -Outputs @($r.Outputs) -Method $step -Message $message -Data $data | Out-Null
                $summary.Done++
                if ($step -eq 'Bim') { $summary.ViaBim++ } else { $summary.ViaDax++ }
                $done = $true
            }
            else {
                $failMessage = $step + ' extraction failed'
                if ($r -and $r.Message) { $failMessage = [string]$r.Message }
                $notes.Add($step + ' failed: ' + $failMessage)
                if ($steps.Count -gt 0) { Write-IQLog -Level Warn -Stage $stage -Item $w.Item -Message ($step + ' extraction failed (' + $failMessage + '); falling back to ' + $steps[0]) }
                else {
                    Set-IQItemDone -Stage $stage -ItemKey $key -Item $w.Item -Status Failed -Method $step -Message ($notes -join '; ') -Data @{ BaseName = $w.BaseName; DatasetId = $w.DatasetId; WorkspaceId = $w.WorkspaceId } | Out-Null
                    $summary.Failed++
                    $done = $true
                }
            }
        }
    }

    if ($budgetStop) {
        $summary.BudgetStop = $true
        Write-IQLog -Level Warn -Stage $stage -Message 'Time budget reached: model detail stopped; the remaining models are processed on the next start.'
    }
    Write-IQLog -Level Info -Stage $stage -Message ("Model detail finished: {0} done ({1} Tabular Editor, {2} Bim parser, {3} DAX), {4} failed, {5} skipped, {6} already done" -f $summary.Done, $summary.ViaTabularEditor, $summary.ViaBim, $summary.ViaDax, $summary.Failed, $summary.Skipped, $summary.AlreadyDone)
    return $summary
}
