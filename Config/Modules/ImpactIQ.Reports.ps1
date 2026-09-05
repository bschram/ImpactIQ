#Requires -Version 5.1
<#
    ImpactIQ v3 - ReportBackup and ReportDetail stages (ImpactIQ.Reports.ps1)

    Ports the Report Backup section of the original monolith ("Final PS Script.txt" lines 2660-3045) and the
    Report Detail extraction (lines 3046-3135):
      - Export-ReportUsingAPI            -> Export-IQReportUsingApi (Invoke-IQDownload, GUID-safe routes, .partial + move)
      - Export-ReportDefinitionAsPbix    -> Export-IQReportDefinitionAsPbix (Invoke-IQFabricLro; PBIR staging, legacy Layout
                                            re-encoded to UTF-16 LE, synthesized root Connections file - behaviour kept verbatim)
      - Get-FreeDriveLetter + subst      -> Get-IQFreeDriveLetter, used ONLY when the short extract path is still longer than
                                            200 characters (brief 8.1); the normal path is <BaseFolder>\Config\Temp\x-<8-char hash>
      - pbi-tools extract / generate-bim -> Invoke-IQReportModelExtract (Pro workspaces, IncludeModel exports only), then the
                                            Tabular Editor rename of Model.Database.Name/ID and the move into Model Backups\<RunId>
      - the two Report Detail csx runs   -> Invoke-IQReportDetailStage (PBIR script first, then the classic Layout script, both
                                            against Config\Blank Model.bim, plus the leftover VOL folder cleanup)

    Every report is checkpointed the moment it is done (Set-IQItemDone -Stage ReportBackup -ItemKey <ReportId>, outputs =
    exported file + .bim when produced) and skipped on re-run. Nothing here uses Read-Host, WinForms, globals or
    Write-Host; all shared state lives in $script:IQ. Windows PowerShell 5.1 and PowerShell 7 compatible; pbi-tools,
    Tabular Editor and subst are Windows-only and guarded by $script:IQ.IsWindows / $script:IQ.Tools.*Works.

    Cross-module functions used (brief section 2): Write-IQLog, Get-IQCleanName, Get-IQSafeKey, Get-IQDateFolder,
    ConvertTo-IQJsonFile, ConvertFrom-IQJsonFile, Invoke-IQFabricLro, Invoke-IQDownload, Test-IQItemDone, Set-IQItemDone,
    Get-IQItemCheckpoint, Invoke-IQProcess, Invoke-IQTabularEditor, Get-IQSelectedReports, Get-IQSelectedWorkspaces,
    Get-IQSelectedDatasets, Get-IQReportsWithSensitivityLabel. Private helpers are prefixed *-IQReport* and are not part of
    the contract.

    Audit items honoured: C7-01 (paginated detection via ReportType / ReportWebUrl, RDL branch reachable), C7-02 (group-less
    routes for pseudo workspaces, no Fabric fallback there), C7-03 (all output through Write-IQLog, per-item checkpoints),
    C7-04/C7-13 (definition export verified, staging cleaned in finally), C7-05 (pbi-tools needs Power BI Desktop: extraction
    is attempted once and disabled for the rest of the run when pbi-tools reports a Desktop/msmdsrv problem; the report backup
    itself always stays a success), C7-06 (subst only as long-path fallback),
    C7-07 (retry/backoff and bounded LRO polling come from the Http module), C7-08/C7-09 (no folder wipe, .partial downloads,
    resume via checkpoints, .bim moved immediately per report, existing .bim reused), C7-10 (name collisions get an id suffix,
    empty names get a fallback), C7-11 (pbi-tools only on real IncludeModel PBIX files), C7-12 (exit codes and tool output
    captured), C7-14 (ReportExports.txt with per-report export facts), C7-15 (optional Options.ReportDefinitionFormat),
    C7-17 (no $args, C# escaping), C8-01/C8-02 (TE2 exit codes, timeouts, captured output), C8-08 (stale TXT files and
    unzip folders removed before the PBIR script runs; classic script only after the PBIR script succeeded),
    X2-H1/X3-H2/X4-04 (IMPACTIQ_BASE / IMPACTIQ_DATE_FOLDER environment for the csx, run folder made the newest dated
    folder via a junction when needed), X3-D1/X3-DG1/X4-20 (<name>.meta.json sidecar per report with ids and method).
#>

# =====================================================================================================================
# Private helpers
# =====================================================================================================================

function Get-IQReportOption {
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

function Get-IQReportMember {
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

function ConvertTo-IQReportBool {
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

function Test-IQReportGuid {
    <#
    .SYNOPSIS
    True when the value looks like a GUID (real workspace ids; pseudo workspaces such as "My Workspace" are not) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value.Trim() -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')
}

function Test-IQReportFileHasContent {
    <#
    .SYNOPSIS
    True when the file exists and is larger than zero bytes; zero-byte leftovers are deleted (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    $length = 0
    try { $length = (Get-Item -LiteralPath $Path).Length } catch { $length = 0 }
    if ($length -gt 0) { return $true }
    try { Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue } catch { $null = $null }
    return $false
}

function Get-IQReportFileSize {
    <#
    .SYNOPSIS
    File size in bytes, 0 when missing (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return 0 }
    try { if (Test-Path -LiteralPath $Path -PathType Leaf) { return [int64](Get-Item -LiteralPath $Path).Length } } catch { $null = $null }
    return 0
}

function Get-IQReportRunFolder {
    <#
    .SYNOPSIS
    Returns (and creates) "<Report Backups>\<RunId>" for the current run (private).
    #>
    [CmdletBinding()]
    param()
    $folder = $null
    if ($script:IQ.ContainsKey('RunPaths') -and $script:IQ.RunPaths -and $script:IQ.RunPaths.ReportBackups) { $folder = [string]$script:IQ.RunPaths.ReportBackups }
    else {
        $runId = [string]$script:IQ.RunId
        if ([string]::IsNullOrWhiteSpace($runId)) { $runId = Get-Date -Format 'yyyy-MM-dd' }
        $folder = Join-Path ([string]$script:IQ.Paths.ReportBackups) $runId
    }
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    return $folder
}

function Get-IQReportModelFolder {
    <#
    .SYNOPSIS
    Returns (and creates) "<Model Backups>\<RunId>", the destination of Pro-workspace .bim files (private).
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

function Get-IQReportTempFolder {
    <#
    .SYNOPSIS
    Returns (and creates) the short scratch root <BaseFolder>\Config\Temp used for staging and pbi-tools extraction (private).
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

function Get-IQReportShortKey {
    <#
    .SYNOPSIS
    Deterministic 8-character hex hash of a value (SHA-1), used for short temp folder names (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Value)
    if ($null -eq $Value) { $Value = '' }
    $sha = $null
    try {
        $sha = [System.Security.Cryptography.SHA1]::Create()
        $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Value))
        $hex = [System.BitConverter]::ToString($bytes).Replace('-', '').ToLowerInvariant()
        return $hex.Substring(0, 8)
    }
    finally {
        if ($null -ne $sha) { $sha.Dispose() }
    }
}

function ConvertTo-IQReportCSharpString {
    <#
    .SYNOPSIS
    Escapes a value for use inside a C# string literal in a Tabular Editor script (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Value)
    if ($null -eq $Value) { return '' }
    return $Value.Replace('\', '\\').Replace('"', '\"').Replace("`r", '\r').Replace("`n", '\n')
}

function ConvertTo-IQReportTsvField {
    <#
    .SYNOPSIS
    Makes a value safe for a tab-separated row (tabs, CR and LF replaced by spaces) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    return $text.Replace("`r`n", ' ').Replace("`r", ' ').Replace("`n", ' ').Replace("`t", ' ')
}

function Remove-IQReportPath {
    <#
    .SYNOPSIS
    Deletes a file or folder (recursively) without ever throwing (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    try {
        if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop }
    }
    catch {
        Write-IQLog -Level Debug -Message ("Could not remove '{0}': {1}" -f $Path, $_.Exception.Message)
    }
}

function Test-IQReportTabularEditorAvailable {
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

function Get-IQReportTabularEditorReason {
    <#
    .SYNOPSIS
    One-line explanation of why Tabular Editor cannot be used (private).
    #>
    [CmdletBinding()]
    param()
    if (-not $script:IQ.IsWindows) { return 'Tabular Editor requires Windows; this host is not Windows' }
    $reason = $null
    try { if ($script:IQ.Tools -and $script:IQ.Tools.ContainsKey('TabularEditorPreflight')) { $reason = [string]$script:IQ.Tools.TabularEditorPreflight } } catch { $reason = $null }
    if ($reason -and $reason -notmatch '^(ok|true)$') { return ('Tabular Editor preflight failed: ' + $reason) }
    return 'Tabular Editor is not available (preflight did not pass)'
}

function Test-IQReportPbiToolsAvailable {
    <#
    .SYNOPSIS
    Returns @{ Available = <bool>; Reason = <string> } for pbi-tools model extraction on this host (private).
    .DESCRIPTION
    pbi-tools (Desktop edition) needs Windows, a working executable (Initialize-IQTools probe) and a Power BI Desktop
    installation to read the embedded model of an IncludeModel PBIX (audit C7-05). When unavailable the report backup
    itself still succeeds and ModelDetail falls back to DAX for that Pro model.
    #>
    [CmdletBinding()]
    param()
    if (-not $script:IQ.IsWindows) { return @{ Available = $false; Reason = 'pbi-tools requires Windows; this host is not Windows' } }
    if (-not $script:IQ.Tools) { return @{ Available = $false; Reason = 'tools were not initialised' } }
    $path = [string]$script:IQ.Tools.PbiToolsPath
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path)) { return @{ Available = $false; Reason = "pbi-tools executable not found ('$path')" } }
    $works = $false
    try { $works = [bool]$script:IQ.Tools.PbiToolsWorks } catch { $works = $false }
    if (-not $works) {
        $probe = ''
        try { if ($script:IQ.Tools.ContainsKey('PbiToolsPreflight')) { $probe = [string]$script:IQ.Tools.PbiToolsPreflight } } catch { $probe = '' }
        if ($probe) { return @{ Available = $false; Reason = ('pbi-tools probe failed: ' + $probe) } }
        return @{ Available = $false; Reason = 'pbi-tools probe did not pass' }
    }
    $disabled = ''
    try { if ($script:IQ.Tools.ContainsKey('PbiToolsExtractDisabled')) { $disabled = [string]$script:IQ.Tools.PbiToolsExtractDisabled } } catch { $disabled = '' }
    if (-not [string]::IsNullOrWhiteSpace($disabled)) { return @{ Available = $false; Reason = ('model extraction disabled for the rest of this run: ' + $disabled) } }
    # Power BI Desktop detection (Initialize-IQTools) is a heuristic: when it is negative we still try the first PBIX and
    # disable extraction for the rest of the run only when pbi-tools itself reports a Desktop/msmdsrv problem (C7-05).
    $desktopWarning = ''
    try {
        if ($script:IQ.Tools.ContainsKey('PbiDesktopFound')) {
            $desktop = ConvertTo-IQReportBool -Value $script:IQ.Tools.PbiDesktopFound
            if ($desktop -eq $false) { $desktopWarning = 'Power BI Desktop was not detected on this host; pbi-tools needs it to read the embedded model of a PBIX. Extraction is attempted once and disabled for the run if pbi-tools reports that problem (ModelDetail then falls back to DAX)' }
        }
    }
    catch { $desktopWarning = '' }
    return @{ Available = $true; Reason = ''; Warning = $desktopWarning }
}

function Disable-IQReportModelExtract {
    <#
    .SYNOPSIS
    Disables pbi-tools model extraction for the rest of the run when a failure points at a missing Power BI Desktop / AS engine (private).
    .DESCRIPTION
    Returns $true when extraction was disabled. Triggers: the tool output mentions Power BI Desktop / msmdsrv / PBIDesktop,
    or Power BI Desktop was not detected by Initialize-IQTools and the very first extraction failed. Later reports then
    skip the pbi-tools round-trip in seconds instead of minutes (audit C7-05); their report backup still succeeds.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Output,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Reason,
        [Parameter(Mandatory = $false)][string]$Stage = 'ReportBackup',
        [Parameter(Mandatory = $false)][string]$Item
    )
    if (-not $script:IQ.Tools) { return $false }
    $mentionsDesktop = (-not [string]::IsNullOrWhiteSpace($Output)) -and ($Output -match '(?i)Power ?BI ?Desktop|msmdsrv|PBIDesktop|Analysis Services instance|no .*desktop .*install')
    $desktopMissing = $false
    try { if ($script:IQ.Tools.ContainsKey('PbiDesktopFound')) { $desktopMissing = ((ConvertTo-IQReportBool -Value $script:IQ.Tools.PbiDesktopFound) -eq $false) } } catch { $desktopMissing = $false }
    $firstAttempt = $true
    try { if ($script:IQ.Tools.ContainsKey('PbiToolsExtractAttempts')) { $firstAttempt = ([int]$script:IQ.Tools.PbiToolsExtractAttempts -le 1) } } catch { $firstAttempt = $true }
    if (-not ($mentionsDesktop -or ($desktopMissing -and $firstAttempt))) { return $false }
    $why = 'pbi-tools cannot read embedded models on this host'
    if ($mentionsDesktop) { $why += ' (its output mentions Power BI Desktop / msmdsrv)' }
    elseif ($desktopMissing) { $why += ' (Power BI Desktop not detected and the first extraction failed)' }
    if (-not [string]::IsNullOrWhiteSpace($Reason)) { $why += ': ' + $Reason }
    $script:IQ.Tools.PbiToolsExtractDisabled = $why
    Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message ('Model extraction from PBIX files disabled for the rest of this run - ' + $why + '. Report backups continue; ModelDetail falls back to DAX for Pro models.')
    return $true
}

# =====================================================================================================================
# Work list (monolith lines 2874-2906: names, extension, per-report facts)
# =====================================================================================================================

function Get-IQReportWorkList {
    <#
    .SYNOPSIS
    Builds the per-report work list from the inventory: names, file paths, capacity flags, paginated detection, sensitivity labels (private).
    .DESCRIPTION
    One entry per report in scope (Get-IQSelectedReports): @{ Key (ReportId); Item ("<CleanWs> ~ <CleanReport>"); ReportId;
    ReportName; WorkspaceId; WorkspaceName; DatasetId; DatasetName; DatasetWorkspaceId; DatasetWorkspaceName; IsDedicated;
    IsPseudoWorkspace; NoAccess; IsPaginated; HasSensitivityLabel; ReportType; FileName; FilePath; MetaPath; ModelBaseName;
    BimPath; ShortKey }. Names use Get-IQCleanName (the exact monolith sanitiser). Two reports that sanitise to the same file
    name get a " (<8-char id>)" suffix on the later one (audit C7-10); an empty report name becomes "Report <8-char id>".
    The .bim of a Pro report is named after its DATASET ("<CleanDatasetWs> ~ <CleanDataset>.bim", audit C8-03/X1-08) so
    ModelDetail finds one model per dataset regardless of report names.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RunFolder,
        [Parameter(Mandatory = $true)][string]$ModelFolder
    )
    $reports = @()
    try { $reports = @(Get-IQSelectedReports | Where-Object { $null -ne $_ }) }
    catch {
        Write-IQLog -Level Error -Message ("Could not read the selected reports from the inventory: " + $_.Exception.Message) -Exception $_.Exception
        return @()
    }
    $wsById = @{}
    try {
        foreach ($ws in @(Get-IQSelectedWorkspaces)) {
            if ($null -eq $ws) { continue }
            $id = [string](Get-IQReportMember -Object $ws -Name 'WorkspaceId')
            if ($id -eq '') { continue }
            $wsById[$id.ToLowerInvariant()] = $ws
        }
    }
    catch { Write-IQLog -Level Debug -Message ("Get-IQSelectedWorkspaces failed; using report-level capacity flags: " + $_.Exception.Message) }
    $dsById = @{}
    try {
        foreach ($ds in @(Get-IQSelectedDatasets)) {
            if ($null -eq $ds) { continue }
            $id = [string](Get-IQReportMember -Object $ds -Name 'DatasetId')
            if ($id -eq '') { continue }
            if (-not $dsById.ContainsKey($id.ToLowerInvariant())) { $dsById[$id.ToLowerInvariant()] = $ds }
        }
    }
    catch { Write-IQLog -Level Debug -Message ("Get-IQSelectedDatasets failed; using report-level dataset names: " + $_.Exception.Message) }
    $labels = @{}
    try { $labels = Get-IQReportsWithSensitivityLabel; if ($null -eq $labels) { $labels = @{} } }
    catch { Write-IQLog -Level Debug -Message ("Get-IQReportsWithSensitivityLabel failed; assuming no labels: " + $_.Exception.Message); $labels = @{} }

    $work = New-Object System.Collections.Generic.List[object]
    $usedNames = @{}
    $seenIds = @{}
    foreach ($report in $reports) {
        $reportId = [string](Get-IQReportMember -Object $report -Name 'ReportId')
        if ([string]::IsNullOrWhiteSpace($reportId)) { continue }
        if ($seenIds.ContainsKey($reportId.ToLowerInvariant())) { continue }
        $seenIds[$reportId.ToLowerInvariant()] = $true
        $reportName = [string](Get-IQReportMember -Object $report -Name 'ReportName')
        $workspaceId = [string](Get-IQReportMember -Object $report -Name 'WorkspaceId')
        $workspaceName = [string](Get-IQReportMember -Object $report -Name 'WorkspaceName')
        $ws = $null
        if ($workspaceId -ne '' -and $wsById.ContainsKey($workspaceId.ToLowerInvariant())) { $ws = $wsById[$workspaceId.ToLowerInvariant()] }
        if ($null -ne $ws -and [string]::IsNullOrWhiteSpace($workspaceName)) { $workspaceName = [string](Get-IQReportMember -Object $ws -Name 'WorkspaceName') }
        $isPseudo = -not (Test-IQReportGuid -Value $workspaceId)
        $noAccess = ($workspaceId -eq 'Shared Reports (No Workspace Access)' -or $workspaceName -eq 'Shared Reports (No Workspace Access)')
        $dedicated = $null
        if ($null -ne $ws) { $dedicated = ConvertTo-IQReportBool -Value (Get-IQReportMember -Object $ws -Name 'WorkspaceIsOnDedicatedCapacity') }
        if ($null -eq $dedicated) { $dedicated = ConvertTo-IQReportBool -Value (Get-IQReportMember -Object $report -Name 'WorkspaceIsOnDedicatedCapacity') }
        if ($null -eq $dedicated) {
            Write-IQLog -Level Debug -Item $reportName -Message "Capacity type of workspace '$workspaceName' unknown; treating as Pro"
            $dedicated = $false
        }
        if ($isPseudo) { $dedicated = $false }

        # Clean names exactly like the monolith (lines 2884-2896), with fallbacks for names that sanitise to nothing.
        $shortId = $reportId
        if ($shortId.Length -gt 8) { $shortId = $shortId.Substring(0, 8) }
        $cleanWorkspaceName = Get-IQCleanName -Name $workspaceName
        if ([string]::IsNullOrWhiteSpace($cleanWorkspaceName)) { $cleanWorkspaceName = 'Workspace' }
        $cleanReportName = Get-IQCleanName -Name $reportName
        if ([string]::IsNullOrWhiteSpace($cleanReportName)) { $cleanReportName = 'Report ' + $shortId }

        # Paginated detection (audit C7-01): the renamed row carries ReportType / ReportWebUrl, never WebUrl.
        $reportType = [string](Get-IQReportMember -Object $report -Name 'ReportType')
        $webUrl = [string](Get-IQReportMember -Object $report -Name 'ReportWebUrl')
        $isPaginated = ($reportType -ieq 'PaginatedReport') -or ($webUrl -like '*/rdlreports/*')
        $extension = 'pbix'
        if ($isPaginated) { $extension = 'rdl' }

        $baseName = $cleanWorkspaceName + ' ~ ' + $cleanReportName
        $nameKey = ($baseName + '.' + $extension).ToLowerInvariant()
        if ($usedNames.ContainsKey($nameKey)) {
            $unique = $baseName + ' (' + $shortId + ')'
            Write-IQLog -Level Warn -Item $baseName -Message "Another report in scope sanitises to the same file name; using '$unique' for report $reportId"
            $baseName = $unique
            $nameKey = ($baseName + '.' + $extension).ToLowerInvariant()
        }
        $usedNames[$nameKey] = $true
        $fileName = $baseName + '.' + $extension

        # Dataset facts (for the Pro .bim name and the sidecar).
        $datasetId = [string](Get-IQReportMember -Object $report -Name 'DatasetId')
        $datasetName = [string](Get-IQReportMember -Object $report -Name 'DatasetName')
        $datasetWorkspaceId = [string](Get-IQReportMember -Object $report -Name 'DatasetWorkspaceId')
        if ([string]::IsNullOrWhiteSpace($datasetWorkspaceId)) { $datasetWorkspaceId = $workspaceId }
        $datasetWorkspaceName = $null
        $dsRow = $null
        if ($datasetId -ne '' -and $dsById.ContainsKey($datasetId.ToLowerInvariant())) { $dsRow = $dsById[$datasetId.ToLowerInvariant()] }
        if ($null -ne $dsRow) {
            $n = [string](Get-IQReportMember -Object $dsRow -Name 'DatasetName')
            if (-not [string]::IsNullOrWhiteSpace($n)) { $datasetName = $n }
            $dwid = [string](Get-IQReportMember -Object $dsRow -Name 'WorkspaceId')
            if (-not [string]::IsNullOrWhiteSpace($dwid)) { $datasetWorkspaceId = $dwid }
            $datasetWorkspaceName = [string](Get-IQReportMember -Object $dsRow -Name 'WorkspaceName')
        }
        if ([string]::IsNullOrWhiteSpace($datasetWorkspaceName) -and $datasetWorkspaceId -ne '' -and $wsById.ContainsKey($datasetWorkspaceId.ToLowerInvariant())) {
            $datasetWorkspaceName = [string](Get-IQReportMember -Object $wsById[$datasetWorkspaceId.ToLowerInvariant()] -Name 'WorkspaceName')
        }
        if ([string]::IsNullOrWhiteSpace($datasetWorkspaceName)) { $datasetWorkspaceName = $workspaceName }
        $modelBaseName = $baseName
        if (-not [string]::IsNullOrWhiteSpace($datasetId) -and -not [string]::IsNullOrWhiteSpace($datasetName) -and $datasetName -ne 'Unknown Dataset') {
            $cleanDsWs = Get-IQCleanName -Name $datasetWorkspaceName
            if ([string]::IsNullOrWhiteSpace($cleanDsWs)) { $cleanDsWs = 'Workspace' }
            $cleanDs = Get-IQCleanName -Name $datasetName
            if ([string]::IsNullOrWhiteSpace($cleanDs)) { $cleanDs = 'Model ' + $datasetId.Substring(0, [Math]::Min(8, $datasetId.Length)) }
            $modelBaseName = $cleanDsWs + ' ~ ' + $cleanDs
        }

        $hasLabel = $false
        try { if ($labels -is [System.Collections.IDictionary] -and $labels.Contains($reportId)) { $hasLabel = $true } } catch { $hasLabel = $false }

        $work.Add(@{
                Key                  = $reportId
                Item                 = $baseName
                Report               = $report
                ReportId             = $reportId
                ReportName           = $reportName
                ReportType           = $reportType
                WorkspaceId          = $workspaceId
                WorkspaceName        = $workspaceName
                DatasetId            = $datasetId
                DatasetName          = $datasetName
                DatasetWorkspaceId   = $datasetWorkspaceId
                DatasetWorkspaceName = $datasetWorkspaceName
                IsDedicated          = [bool]$dedicated
                IsPseudoWorkspace    = $isPseudo
                NoAccess             = $noAccess
                IsPaginated          = $isPaginated
                HasSensitivityLabel  = $hasLabel
                Extension            = $extension
                FileName             = $fileName
                FilePath             = (Join-Path $RunFolder $fileName)
                MetaPath             = (Join-Path $RunFolder ($baseName + '.meta.json'))
                ModelBaseName        = $modelBaseName
                BimPath              = (Join-Path $ModelFolder ($modelBaseName + '.bim'))
                ShortKey             = (Get-IQReportShortKey -Value $reportId)
            })
    }
    return $work.ToArray()
}

# =====================================================================================================================
# Export API (monolith 2668-2691) and getDefinition fallback (monolith 2693-2859)
# =====================================================================================================================

function Export-IQReportUsingApi {
    <#
    .SYNOPSIS
    Downloads a report through the Power BI Export API (IncludeModel | LiveConnect) to OutFilePath; returns $true on success.
    .DESCRIPTION
    Port of the monolith's Export-ReportUsingAPI. GUID workspace ids use groups/{ws}/reports/{id}/Export, pseudo workspaces
    ("My Workspace") use the group-less reports/{id}/Export route (audit C7-02). The body is streamed to "<file>.partial"
    through Invoke-IQDownload (retry/backoff, zero-byte files deleted) and moved into place only after a non-empty file was
    written (audit C7-08), so a killed run never leaves a truncated file that looks complete.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$WorkspaceId,
        [Parameter(Mandatory = $true)][string]$ReportId,
        [Parameter(Mandatory = $true)][string]$OutFilePath,
        [Parameter(Mandatory = $true)][ValidateSet('IncludeModel', 'LiveConnect')][string]$DownloadType,
        [Parameter(Mandatory = $false)][string]$Item,
        [Parameter(Mandatory = $false)][string]$Stage = 'ReportBackup'
    )
    $path = $null
    if (Test-IQReportGuid -Value $WorkspaceId) { $path = 'groups/' + $WorkspaceId + '/reports/' + $ReportId + '/Export?downloadType=' + $DownloadType }
    else { $path = 'reports/' + $ReportId + '/Export?downloadType=' + $DownloadType }
    $partial = $OutFilePath + '.partial'
    try {
        Remove-IQReportPath -Path $partial
        $timeout = 600
        try { $timeout = [int](Get-IQReportOption -Name 'DownloadTimeoutSec' -Default 600) } catch { $timeout = 600 }
        if ($timeout -lt 30) { $timeout = 30 }
        $ok = Invoke-IQDownload -Url $path -OutFile $partial -Api PowerBI -TimeoutSec $timeout -Stage $Stage
        if ($ok -and (Test-IQReportFileHasContent -Path $partial)) {
            Remove-IQReportPath -Path $OutFilePath
            Move-Item -LiteralPath $partial -Destination $OutFilePath -Force -ErrorAction Stop
            Write-IQLog -Level Debug -Stage $Stage -Item $Item -Message ("Export ({0}) wrote {1:N0} bytes" -f $DownloadType, (Get-IQReportFileSize -Path $OutFilePath))
            return $true
        }
        Write-IQLog -Level Debug -Stage $Stage -Item $Item -Message ("Export ({0}) produced no file" -f $DownloadType)
        Remove-IQReportPath -Path $partial
        return $false
    }
    catch {
        Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message ("Export ({0}) failed: {1}" -f $DownloadType, $_.Exception.Message)
        Remove-IQReportPath -Path $partial
        return $false
    }
}

function Export-IQReportDefinitionAsPbix {
    <#
    .SYNOPSIS
    Exports a report's definition through the Fabric getDefinition API, stages the parts like a PBIX and zips them to OutFilePath.
    .DESCRIPTION
    Port of the monolith's Export-ReportDefinitionAsPbix (lines 2693-2859) with the same on-disk result: parts nested under
    "Report\", a legacy report.json turned into "Report\Layout" re-encoded as UTF-16 LE without BOM, a synthesized root
    "Connections" file (EntityDataSource / pbiServiceLive / RemoteArtifacts) from definition.pbir, everything zipped and
    renamed .pbix. NOTE: a zipped PBIR definition renamed to .pbix is NOT an openable Power BI Desktop file by design.
    Long-running operations are handled by Invoke-IQFabricLro. Staging happens under <BaseFolder>\Config\Temp\def-<hash>
    and is always cleaned up (audit C7-04/C7-13). Returns @{ Success; Format ('PBIR'|'Legacy'|''); Message; PartCount;
    PageCount; VisualCount; DatasetReferenceType; DatasetIdFromPbir }.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$WorkspaceId,
        [Parameter(Mandatory = $true)][string]$ReportId,
        [Parameter(Mandatory = $true)][string]$OutFilePath,
        [Parameter(Mandatory = $false)][string]$Item,
        [Parameter(Mandatory = $false)][string]$Stage = 'ReportBackup'
    )
    $result = @{ Success = $false; Format = ''; Message = ''; PartCount = 0; PageCount = 0; VisualCount = 0; DatasetReferenceType = ''; DatasetIdFromPbir = '' }
    try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue } catch { $null = $null }

    $lroTimeout = 10
    try { $lroTimeout = [int](Get-IQReportOption -Name 'DefinitionTimeoutMinutes' -Default 10) } catch { $lroTimeout = 10 }
    if ($lroTimeout -lt 1) { $lroTimeout = 1 }
    $definitionPath = 'workspaces/' + $WorkspaceId + '/reports/' + $ReportId + '/getDefinition'

    # Optional explicit format (audit C7-15): Options.ReportDefinitionFormat = PBIR | PBIR-Legacy. Default = monolith
    # behaviour (no format parameter; the service returns the stored format). A rejected format falls back to no format.
    $definition = $null
    $requestedFormat = [string](Get-IQReportOption -Name 'ReportDefinitionFormat' -Default '')
    try {
        if (-not [string]::IsNullOrWhiteSpace($requestedFormat)) {
            $definition = Invoke-IQFabricLro -Method POST -Path $definitionPath -Query @{ format = $requestedFormat } -TimeoutMinutes $lroTimeout -Stage $Stage
            if ($null -eq $definition) { Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message "getDefinition with format=$requestedFormat was not accepted; retrying without a format" }
        }
        if ($null -eq $definition) {
            $definition = Invoke-IQFabricLro -Method POST -Path $definitionPath -TimeoutMinutes $lroTimeout -Stage $Stage
        }
    }
    catch {
        $result.Message = 'getDefinition failed: ' + $_.Exception.Message
        Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message $result.Message
        return $result
    }
    if ($null -eq $definition) {
        $result.Message = 'getDefinition returned no definition (see previous warnings)'
        return $result
    }
    $parts = @()
    try {
        $def = Get-IQReportMember -Object $definition -Name 'definition'
        $parts = @(Get-IQReportMember -Object $def -Name 'parts')
        $parts = @($parts | Where-Object { $null -ne $_ })
    }
    catch { $parts = @() }
    if ($parts.Count -eq 0) {
        $result.Message = 'No definition parts returned by getDefinition'
        Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message $result.Message
        return $result
    }
    $result.PartCount = $parts.Count

    $tempRoot = Get-IQReportTempFolder
    $shortKey = Get-IQReportShortKey -Value $ReportId
    $stageDir = Join-Path $tempRoot ('def-' + $shortKey)
    $tempZip = Join-Path $tempRoot ('def-' + $shortKey + '.zip')
    $separator = [System.IO.Path]::DirectorySeparatorChar
    try {
        Remove-IQReportPath -Path $stageDir
        Remove-IQReportPath -Path $tempZip
        New-Item -ItemType Directory -Path $stageDir -Force | Out-Null

        # Stage the parts. getDefinition returns part paths WITHOUT a "Report/" prefix (e.g. "definition/report.json"),
        # but the PBIX/PBIR layout expected by the Report Detail extractor is "Report/definition/...". Nest everything
        # under a "Report" folder so the zipped .pbix matches that structure.
        $reportRoot = Join-Path $stageDir 'Report'
        New-Item -ItemType Directory -Path $reportRoot -Force | Out-Null
        $rootFull = [System.IO.Path]::GetFullPath($reportRoot).TrimEnd('\', '/') + $separator
        $pageCount = 0
        $visualCount = 0
        foreach ($part in $parts) {
            $partRelative = [string](Get-IQReportMember -Object $part -Name 'path')
            if ([string]::IsNullOrWhiteSpace($partRelative)) { continue }
            $normalised = $partRelative.Replace('/', $separator).Replace('\', $separator).TrimStart($separator)
            $partPath = Join-Path $reportRoot $normalised
            $partFull = [System.IO.Path]::GetFullPath($partPath)
            if (-not $partFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
                Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message "Skipping definition part with an unsafe path: $partRelative"
                continue
            }
            $partDir = Split-Path -Path $partFull -Parent
            if (-not (Test-Path -LiteralPath $partDir)) { New-Item -ItemType Directory -Path $partDir -Force | Out-Null }
            $payloadType = [string](Get-IQReportMember -Object $part -Name 'payloadType')
            $payload = Get-IQReportMember -Object $part -Name 'payload'
            if ($payloadType -ieq 'InlineBase64') {
                [System.IO.File]::WriteAllBytes($partFull, [System.Convert]::FromBase64String([string]$payload))
            }
            else {
                [System.IO.File]::WriteAllText($partFull, [string]$payload)
            }
            if ($normalised -match '(^|[\\/])page\.json$') { $pageCount++ }
            elseif ($normalised -match '(^|[\\/])visual\.json$') { $visualCount++ }
        }
        $result.PageCount = $pageCount
        $result.VisualCount = $visualCount

        # If getDefinition returned the legacy (non-PBIR) layout, there is NO "definition" folder - instead a root-level
        # "report.json" carries the report layout. Rename it to "Layout" (no extension) AND re-encode it as UTF-16 LE
        # (no BOM) to match the classic PBIX layout the downstream extractor expects. getDefinition returns report.json
        # as UTF-8, but a real PBIX "Report/Layout" is UTF-16 LE; without this conversion the parser reads garbage.
        $defFolder = Join-Path $reportRoot 'definition'
        if (Test-Path -LiteralPath $defFolder -PathType Container) {
            $result.Format = 'PBIR'
        }
        else {
            $reportJson = Join-Path $reportRoot 'report.json'
            if (Test-Path -LiteralPath $reportJson -PathType Leaf) {
                $result.Format = 'Legacy'
                $layoutPath = Join-Path $reportRoot 'Layout'
                if (Test-Path -LiteralPath $layoutPath) { Remove-Item -LiteralPath $layoutPath -Force }
                $layoutText = [System.IO.File]::ReadAllText($reportJson, [System.Text.Encoding]::UTF8)
                $utf16NoBom = New-Object System.Text.UnicodeEncoding($false, $false)  # LE, no BOM
                [System.IO.File]::WriteAllText($layoutPath, $layoutText, $utf16NoBom)
                Remove-Item -LiteralPath $reportJson -Force
            }
        }

        # getDefinition does NOT return a "Connections" file, but the downstream Report Detail extractor expects one at
        # the ROOT (alongside the "Report" folder) to populate the Connections sheet + ModelID/ReportID. Synthesize it
        # from definition.pbir's datasetReference so the layout matches a PBIX.
        $pbirPath = Join-Path $reportRoot 'definition.pbir'
        if (Test-Path -LiteralPath $pbirPath -PathType Leaf) {
            try {
                $pbir = [System.IO.File]::ReadAllText($pbirPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json
                $dsRef = Get-IQReportMember -Object $pbir -Name 'datasetReference'
                $connString = $null
                $datasetId = ''
                $byConnection = Get-IQReportMember -Object $dsRef -Name 'byConnection'
                $byPath = Get-IQReportMember -Object $dsRef -Name 'byPath'
                if ($null -ne $byConnection -and (Get-IQReportMember -Object $byConnection -Name 'connectionString')) {
                    $connString = [string](Get-IQReportMember -Object $byConnection -Name 'connectionString')
                    $result.DatasetReferenceType = 'byConnection'
                    $m = [regex]::Match($connString, 'semanticmodelid=([0-9a-fA-F-]+)')
                    if ($m.Success) { $datasetId = $m.Groups[1].Value }
                }
                elseif ($null -ne $byPath -and (Get-IQReportMember -Object $byPath -Name 'path')) {
                    $connString = 'byPath:' + [string](Get-IQReportMember -Object $byPath -Name 'path')
                    $result.DatasetReferenceType = 'byPath'
                }
                $result.DatasetIdFromPbir = $datasetId
                if ($connString) {
                    $connObj = [pscustomobject]@{
                        Version         = '3.0'
                        Connections     = @(
                            [pscustomobject]@{
                                Name             = 'EntityDataSource'
                                ConnectionString = $connString
                                ConnectionType   = 'pbiServiceLive'
                            }
                        )
                        RemoteArtifacts = @(
                            [pscustomobject]@{
                                DatasetId = $datasetId
                                ReportId  = $ReportId
                            }
                        )
                    }
                    # Root-level file literally named "Connections" (no extension) to match PBIX; UTF-8 without BOM.
                    $connJson = ConvertTo-Json -InputObject $connObj -Depth 20 -Compress
                    [System.IO.File]::WriteAllText((Join-Path $stageDir 'Connections'), $connJson)
                }
            }
            catch {
                Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message ("Could not synthesize Connections file: " + $_.Exception.Message)
            }
        }

        # Zip the staged folder, then rename the archive to .pbix
        Remove-IQReportPath -Path $OutFilePath
        [System.IO.Compression.ZipFile]::CreateFromDirectory($stageDir, $tempZip)
        Move-Item -LiteralPath $tempZip -Destination $OutFilePath -Force -ErrorAction Stop

        $result.Success = Test-IQReportFileHasContent -Path $OutFilePath
        if ($result.Success) {
            Write-IQLog -Level Debug -Stage $Stage -Item $Item -Message ("getDefinition export ({0}, {1} parts) wrote {2:N0} bytes" -f $result.Format, $parts.Count, (Get-IQReportFileSize -Path $OutFilePath))
        }
        else { $result.Message = 'getDefinition export produced an empty archive' }
    }
    catch {
        $result.Success = $false
        $result.Message = 'getDefinition staging failed: ' + $_.Exception.Message
        Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message $result.Message
        Remove-IQReportPath -Path $OutFilePath
    }
    finally {
        Remove-IQReportPath -Path $stageDir
        Remove-IQReportPath -Path $tempZip
    }
    return $result
}

# =====================================================================================================================
# Model extraction for Pro workspaces (monolith 2956-3020): pbi-tools + subst fallback + TE2 rename + move
# =====================================================================================================================

function Get-IQFreeDriveLetter {
    <#
    .SYNOPSIS
    Returns the first free drive letter from Z: down to D: (monolith Get-FreeDriveLetter); $null when none is free or not on Windows.
    #>
    [CmdletBinding()]
    param()
    if (-not $script:IQ.IsWindows) { return $null }
    try {
        $used = @(Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | ForEach-Object { [string]$_.Name })
        $logical = @()
        try { $logical = @([System.IO.Directory]::GetLogicalDrives() | ForEach-Object { ([string]$_).Substring(0, 1) }) } catch { $logical = @() }
        foreach ($code in 90..68) {   # Z..Y..X.. down to D
            $letter = [string][char]$code
            if ($used -contains $letter) { continue }
            if ($logical -contains $letter) { continue }
            return ($letter + ':')
        }
    }
    catch { Write-IQLog -Level Debug -Message ("Get-IQFreeDriveLetter failed: " + $_.Exception.Message) }
    return $null
}

function New-IQReportSubstMapping {
    <#
    .SYNOPSIS
    Maps a free drive letter to a folder with subst (Windows only); returns the letter ("Z:") or $null when it did not work (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Folder)
    if (-not $script:IQ.IsWindows) { return $null }
    $letter = Get-IQFreeDriveLetter
    if (-not $letter) { Write-IQLog -Level Debug -Message 'No free drive letter for subst'; return $null }
    try {
        $cmd = Join-Path $env:SystemRoot 'System32\cmd.exe'
        if (-not (Test-Path -LiteralPath $cmd)) { $cmd = 'cmd.exe' }
        & $cmd /c "subst $letter `"$Folder`"" 2>&1 | Out-Null
        Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Out-Null   # refresh the drive list
        if (Test-Path -LiteralPath ($letter + '\')) {
            Write-IQLog -Level Debug -Message "subst $letter -> $Folder (long-path fallback)"
            return $letter
        }
        & $cmd /c "subst $letter /D" 2>&1 | Out-Null
    }
    catch { Write-IQLog -Level Debug -Message ("subst failed: " + $_.Exception.Message) }
    return $null
}

function Remove-IQReportSubstMapping {
    <#
    .SYNOPSIS
    Removes a subst drive mapping created by New-IQReportSubstMapping; never throws (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Letter)
    if ([string]::IsNullOrWhiteSpace($Letter) -or -not $script:IQ.IsWindows) { return }
    try {
        $cmd = Join-Path $env:SystemRoot 'System32\cmd.exe'
        if (-not (Test-Path -LiteralPath $cmd)) { $cmd = 'cmd.exe' }
        & $cmd /c "subst $Letter /D" 2>&1 | Out-Null
    }
    catch { Write-IQLog -Level Debug -Message ("subst /D failed for {0}: {1}" -f $Letter, $_.Exception.Message) }
}

function Get-IQReportProcessSummary {
    <#
    .SYNOPSIS
    Short human-readable summary of an Invoke-IQProcess result (exit code / timeout / first error lines / log path) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Result)
    if ($null -eq $Result) { return 'no process result' }
    $parts = @()
    if (Get-IQReportMember -Object $Result -Name 'TimedOut') { $parts += 'timed out' }
    $startError = Get-IQReportMember -Object $Result -Name 'StartError'
    if ($startError) { $parts += ('start failure: ' + $startError) }
    $exit = Get-IQReportMember -Object $Result -Name 'ExitCode'
    if ($null -ne $exit) { $parts += ('exit code ' + $exit) }
    $lines = @()
    $stdOut = [string](Get-IQReportMember -Object $Result -Name 'StdOut')
    $stdErr = [string](Get-IQReportMember -Object $Result -Name 'StdErr')
    if ($stdOut) {
        $all = @($stdOut -split "`r?`n")
        for ($i = 0; $i -lt $all.Count; $i++) {
            if ($all[$i] -match '^\s*(Error\b|Script compilation error|Script error|Unhandled exception|Exception:|Could not|Unable to|Failed|fail)') {
                $lines += $all[$i].Trim()
                if ($i + 1 -lt $all.Count -and $all[$i + 1].Trim()) { $lines += $all[$i + 1].Trim() }
            }
        }
    }
    if ($stdErr) { $lines += @(($stdErr -split "`r?`n") | Where-Object { $_.Trim() } | ForEach-Object { $_.Trim() }) }
    if ($lines.Count -eq 0 -and $stdOut) { $lines = @(($stdOut -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -Last 3) }
    $lines = @($lines | Select-Object -First 5)
    if ($lines.Count -gt 0) { $parts += ($lines -join ' | ') }
    $logFile = Get-IQReportMember -Object $Result -Name 'OutFile'
    if ($logFile) { $parts += ('log: ' + $logFile) }
    return ($parts -join '; ')
}

function Invoke-IQReportModelExtract {
    <#
    .SYNOPSIS
    Extracts the embedded model of a Pro-workspace IncludeModel PBIX with pbi-tools and stores it as "<Model Backups>\<RunId>\<CleanDatasetWs> ~ <CleanDataset>.bim".
    .DESCRIPTION
    Port of monolith lines 2956-3020: pbi-tools "extract <pbix> -extractFolder <folder> -modelSerialization Raw", then
    "generate-bim <folder> -transforms RemovePBIDataSourceVersion", then Tabular Editor sets Model.Database.Name/ID to the
    file base name (-S rename script -B destination), then the .bim lands in Model Backups\<RunId>. Differences: the extract
    folder is the short path <BaseFolder>\Config\Temp\x-<8-char hash> and subst is used only when that path is still longer
    than 200 characters AND the mapping succeeds (audit C7-06); exit codes and output are captured (C7-12); an existing .bim
    for the dataset is reused so one model per dataset is extracted regardless of how many reports share it (C7-09/C8-03);
    every failure is logged and returned instead of thrown - the report backup itself stays a success and ModelDetail
    falls back to DAX for that model. Returns @{ Success; BimPath; Message; Reused; Renamed }.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Work,
        [Parameter(Mandatory = $false)][string]$Stage = 'ReportBackup'
    )
    $out = @{ Success = $false; BimPath = $null; Message = ''; Reused = $false; Renamed = $false }
    $item = [string]$Work.Item
    $destination = [string]$Work.BimPath
    if (Test-IQReportFileHasContent -Path $destination) {
        $out.Success = $true; $out.BimPath = $destination; $out.Reused = $true
        $out.Message = 'existing .bim for this dataset reused'
        Write-IQLog -Level Debug -Stage $Stage -Item $item -Message ("Model already extracted for dataset {0}: {1}" -f $Work.DatasetId, $destination)
        return $out
    }
    $availability = Test-IQReportPbiToolsAvailable
    if (-not $availability.Available) {
        $out.Message = 'model not extracted: ' + $availability.Reason
        Write-IQLog -Level Warn -Stage $Stage -Item $item -Message $out.Message
        return $out
    }
    if (-not (Test-IQReportFileHasContent -Path ([string]$Work.FilePath))) {
        $out.Message = 'model not extracted: the PBIX file is missing or empty'
        Write-IQLog -Level Warn -Stage $Stage -Item $item -Message $out.Message
        return $out
    }

    $timeout = 20
    try { $timeout = [int](Get-IQReportOption -Name 'ToolTimeoutMinutes' -Default 20) } catch { $timeout = 20 }
    if ($timeout -lt 1) { $timeout = 1 }
    $pbiTools = [string]$script:IQ.Tools.PbiToolsPath
    $safeKey = Get-IQSafeKey -Value ([string]$Work.Key)
    $tempRoot = Get-IQReportTempFolder
    $extractFolder = Join-Path $tempRoot ('x-' + [string]$Work.ShortKey)
    $substLetter = $null
    $renameScript = $null
    try {
        Remove-IQReportPath -Path $extractFolder
        New-Item -ItemType Directory -Path $extractFolder -Force | Out-Null
        $target = $extractFolder.TrimEnd('\', '/')
        if ($script:IQ.IsWindows -and $target.Length -gt 200) {
            $substLetter = New-IQReportSubstMapping -Folder $target
            if ($substLetter) { $target = $substLetter }   # drive-relative "Z:" resolves to the root for a fresh process (as the monolith)
            else { Write-IQLog -Level Warn -Stage $Stage -Item $item -Message "Extract path is $($target.Length) characters and subst is unavailable; pbi-tools may hit MAX_PATH" }
        }

        Write-IQLog -Level Info -Stage $Stage -Item $item -Message ('Extracting model from ' + [string]$Work.FileName)
        $attempts = 0
        try { if ($script:IQ.Tools.ContainsKey('PbiToolsExtractAttempts')) { $attempts = [int]$script:IQ.Tools.PbiToolsExtractAttempts } } catch { $attempts = 0 }
        $script:IQ.Tools.PbiToolsExtractAttempts = $attempts + 1
        $extractArgs = ('extract "{0}" -extractFolder "{1}" -modelSerialization Raw' -f [string]$Work.FilePath, $target)
        $r1 = Invoke-IQProcess -FilePath $pbiTools -ArgumentList $extractArgs -WorkingDirectory $tempRoot -TimeoutMinutes $timeout -LogName ('pbitools-extract-' + $safeKey) -Stage $Stage -Item $item
        if ($r1.TimedOut -or $r1.StartError -or [int]$r1.ExitCode -ne 0) {
            $out.Message = 'pbi-tools extract failed: ' + (Get-IQReportProcessSummary -Result $r1)
            Write-IQLog -Level Warn -Stage $Stage -Item $item -Message $out.Message
            Disable-IQReportModelExtract -Output ([string]$r1.StdOut + "`n" + [string]$r1.StdErr) -Reason $out.Message -Stage $Stage -Item $item | Out-Null
            return $out
        }
        $bimArgs = ('generate-bim "{0}" -transforms RemovePBIDataSourceVersion' -f $target)
        $r2 = Invoke-IQProcess -FilePath $pbiTools -ArgumentList $bimArgs -WorkingDirectory $tempRoot -TimeoutMinutes $timeout -LogName ('pbitools-generate-bim-' + $safeKey) -Stage $Stage -Item $item
        if ($r2.TimedOut -or $r2.StartError -or [int]$r2.ExitCode -ne 0) {
            Write-IQLog -Level Warn -Stage $Stage -Item $item -Message ('pbi-tools generate-bim reported: ' + (Get-IQReportProcessSummary -Result $r2))
        }
        $bimFiles = @(Get-ChildItem -LiteralPath $extractFolder -Filter '*.bim' -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Length -gt 0 } | Sort-Object FullName)
        if ($bimFiles.Count -eq 0) {
            $out.Message = 'pbi-tools produced no .bim (the PBIX has no embedded model, or generate-bim failed: ' + (Get-IQReportProcessSummary -Result $r2) + ')'
            Write-IQLog -Level Warn -Stage $Stage -Item $item -Message $out.Message
            Disable-IQReportModelExtract -Output ([string]$r1.StdOut + "`n" + [string]$r1.StdErr + "`n" + [string]$r2.StdOut + "`n" + [string]$r2.StdErr) -Reason $out.Message -Stage $Stage -Item $item | Out-Null
            return $out
        }
        $sourceBim = $bimFiles[0].FullName
        if ($bimFiles.Count -gt 1) { Write-IQLog -Level Debug -Stage $Stage -Item $item -Message ("{0} .bim files generated; using {1}" -f $bimFiles.Count, $sourceBim) }

        # Rename Model.Database.Name/ID to the file base name (monolith 3004-3016) and write straight to the destination.
        $modelName = [string]$Work.ModelBaseName
        $destinationDir = Split-Path -Path $destination -Parent
        if ($destinationDir -and -not (Test-Path -LiteralPath $destinationDir)) { New-Item -ItemType Directory -Path $destinationDir -Force | Out-Null }
        Remove-IQReportPath -Path $destination
        if (Test-IQReportTabularEditorAvailable) {
            $renameScript = Join-Path $tempRoot ('TabularEditor_RenameProModel_' + $safeKey + '.cs')
            $escaped = ConvertTo-IQReportCSharpString -Value $modelName
            $content = 'Model.Database.Name = "' + $escaped + '";' + [Environment]::NewLine + 'Model.Database.ID = "' + $escaped + '";' + [Environment]::NewLine
            [System.IO.File]::WriteAllText($renameScript, $content, (New-Object System.Text.UTF8Encoding($true)))
            $teArgs = ('"{0}" -S "{1}" -B "{2}"' -f $sourceBim, $renameScript, $destination)
            $te = Invoke-IQTabularEditor -ArgumentList $teArgs -TimeoutMinutes $timeout -LogName ('te-rename-' + $safeKey) -Stage $Stage -Item $item
            if ($te.Success -and (Test-IQReportFileHasContent -Path $destination)) {
                $out.Renamed = $true
            }
            else {
                Write-IQLog -Level Warn -Stage $Stage -Item $item -Message ('Tabular Editor could not rename the model (' + [string]$te.FailureReason + '); keeping the .bim as generated by pbi-tools')
                Remove-IQReportPath -Path $destination
            }
        }
        else {
            Write-IQLog -Level Warn -Stage $Stage -Item $item -Message ('Tabular Editor unavailable; Model.Database.Name of ' + $modelName + '.bim is left as generated by pbi-tools')
        }
        if (-not (Test-IQReportFileHasContent -Path $destination)) {
            Move-Item -LiteralPath $sourceBim -Destination $destination -Force -ErrorAction Stop
        }
        if (Test-IQReportFileHasContent -Path $destination) {
            $out.Success = $true
            $out.BimPath = $destination
            $out.Message = 'model extracted with pbi-tools'
            if (-not $out.Renamed) { $out.Message += ' (not renamed by Tabular Editor)' }
            Write-IQLog -Level Success -Stage $Stage -Item $item -Message ("Model saved to {0} ({1:N0} bytes)" -f $destination, (Get-IQReportFileSize -Path $destination))
        }
        else {
            $out.Message = 'the extracted .bim could not be moved to ' + $destination
            Write-IQLog -Level Warn -Stage $Stage -Item $item -Message $out.Message
        }
    }
    catch {
        $out.Success = $false
        $out.Message = 'model extraction failed: ' + $_.Exception.Message
        Write-IQLog -Level Warn -Stage $Stage -Item $item -Message $out.Message
    }
    finally {
        Remove-IQReportSubstMapping -Letter $substLetter
        Remove-IQReportPath -Path $extractFolder
        Remove-IQReportPath -Path $renameScript
    }
    return $out
}

# =====================================================================================================================
# ReportExports.txt (audit C7-14): per-report export facts rebuilt from the checkpoints
# =====================================================================================================================

function Get-IQReportCheckpointList {
    <#
    .SYNOPSIS
    Reads every ReportBackup checkpoint of the current run (done\ReportBackup\*.json) (private).
    #>
    [CmdletBinding()]
    param()
    $list = @()
    $doneRoot = $null
    try {
        if ($script:IQ.ContainsKey('RunPaths') -and $script:IQ.RunPaths -and $script:IQ.RunPaths.Done) { $doneRoot = [string]$script:IQ.RunPaths.Done }
        elseif ($script:IQ.RunPath) { $doneRoot = Join-Path ([string]$script:IQ.RunPath) 'done' }
    }
    catch { $doneRoot = $null }
    if (-not $doneRoot) { return @() }
    $folder = Join-Path $doneRoot 'ReportBackup'
    if (-not (Test-Path -LiteralPath $folder)) { return @() }
    foreach ($file in @(Get-ChildItem -LiteralPath $folder -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        $cp = $null
        try { $cp = ConvertFrom-IQJsonFile -Path $file.FullName } catch { $cp = $null }
        if ($null -ne $cp) { $list += , $cp }
    }
    return @($list)
}

function Write-IQReportExportSummary {
    <#
    .SYNOPSIS
    Writes "<Report Backups>\<RunId>\ReportExports.txt" (tab-separated, header row) from the ReportBackup checkpoints; never throws.
    .DESCRIPTION
    The file becomes a "ReportExports" sheet of Report Detail.xlsx through the Assemble stage's *.txt loop (additive sheet).
    Rebuilt from the checkpoints every time, so re-runs never duplicate rows. UTF-8 without BOM, CRLF, like the csx output.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$RunFolder)
    try {
        if ([string]::IsNullOrWhiteSpace($RunFolder)) { $RunFolder = Get-IQReportRunFolder }
        $checkpoints = @(Get-IQReportCheckpointList)
        $reportDate = [string]$script:IQ.RunId
        if ($reportDate -notmatch '^\d{4}-\d{2}-\d{2}$') {
            $started = $null
            try { $started = [string](Get-IQReportMember -Object $script:IQ.Manifest -Name 'startedUtc') } catch { $started = $null }
            $parsed = [datetime]::MinValue
            if ($started -and [datetime]::TryParse($started, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal, [ref]$parsed)) { $reportDate = $parsed.ToString('yyyy-MM-dd') }
            else { $reportDate = (Get-Date).ToString('yyyy-MM-dd') }
        }
        $columns = @('ReportName', 'ReportID', 'ModelID', 'WorkspaceID', 'WorkspaceName', 'ReportDisplayName', 'ReportType', 'FileName', 'ExportMethod', 'Status', 'Message', 'FileSizeBytes', 'DurationSec', 'DefinitionFormat', 'ModelExtract', 'BimPath', 'ReportDate')
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append(($columns -join "`t")).Append("`r`n")
        $rows = @()
        foreach ($cp in $checkpoints) {
            $data = Get-IQReportMember -Object $cp -Name 'data'
            $rows += , [pscustomobject]@{
                ReportName        = [string](Get-IQReportMember -Object $cp -Name 'item')
                ReportID          = [string](Get-IQReportMember -Object $cp -Name 'itemKey')
                ModelID           = [string](Get-IQReportMember -Object $data -Name 'DatasetId')
                WorkspaceID       = [string](Get-IQReportMember -Object $data -Name 'WorkspaceId')
                WorkspaceName     = [string](Get-IQReportMember -Object $data -Name 'WorkspaceName')
                ReportDisplayName = [string](Get-IQReportMember -Object $data -Name 'ReportName')
                ReportType        = [string](Get-IQReportMember -Object $data -Name 'ReportType')
                FileName          = [string](Get-IQReportMember -Object $data -Name 'FileName')
                ExportMethod      = [string](Get-IQReportMember -Object $cp -Name 'method')
                Status            = [string](Get-IQReportMember -Object $cp -Name 'status')
                Message           = [string](Get-IQReportMember -Object $cp -Name 'message')
                FileSizeBytes     = [string](Get-IQReportMember -Object $data -Name 'SizeBytes')
                DurationSec       = [string](Get-IQReportMember -Object $data -Name 'DurationSec')
                DefinitionFormat  = [string](Get-IQReportMember -Object $data -Name 'DefinitionFormat')
                ModelExtract      = [string](Get-IQReportMember -Object $data -Name 'ModelExtract')
                BimPath           = [string](Get-IQReportMember -Object $data -Name 'BimPath')
                ReportDate        = $reportDate
            }
        }
        foreach ($row in @($rows | Sort-Object ReportName)) {
            $values = @()
            foreach ($c in $columns) { $values += (ConvertTo-IQReportTsvField -Value $row.$c) }
            [void]$sb.Append(($values -join "`t")).Append("`r`n")
        }
        $path = Join-Path $RunFolder 'ReportExports.txt'
        $tmp = $path + '.tmp'
        [System.IO.File]::WriteAllText($tmp, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
        Move-Item -LiteralPath $tmp -Destination $path -Force -ErrorAction Stop
        Write-IQLog -Level Debug -Message ("ReportExports.txt written with {0} row(s)" -f $rows.Count)
        return $path
    }
    catch {
        Write-IQLog -Level Warn -Message ("Could not write ReportExports.txt: " + $_.Exception.Message)
        return $null
    }
}

function Write-IQReportMetaFile {
    <#
    .SYNOPSIS
    Writes the "<name>.meta.json" sidecar next to an exported report (ids, names, export method) (private).
    .DESCRIPTION
    Lets a resumed run reuse a complete download whose checkpoint was never written, and gives the csx extractors a
    place to read ReportId/DatasetId/WorkspaceId for embedded-model PBIX files (audit X3-D1/DG-1, X4-20).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Work,
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$DefinitionFormat
    )
    try {
        $meta = [ordered]@{
            ReportId           = [string]$Work.ReportId
            ReportName         = [string]$Work.ReportName
            ReportType         = [string]$Work.ReportType
            WorkspaceId        = [string]$Work.WorkspaceId
            WorkspaceName      = [string]$Work.WorkspaceName
            DatasetId          = [string]$Work.DatasetId
            DatasetName        = [string]$Work.DatasetName
            DatasetWorkspaceId = [string]$Work.DatasetWorkspaceId
            FileName           = [string]$Work.FileName
            ExportMethod       = $Method
            DefinitionFormat   = [string]$DefinitionFormat
            IsPaginated        = [bool]$Work.IsPaginated
            RunId              = [string]$script:IQ.RunId
            ExportedUtc        = [datetime]::UtcNow.ToString('o')
        }
        ConvertTo-IQJsonFile -Object $meta -Path ([string]$Work.MetaPath)
    }
    catch { Write-IQLog -Level Debug -Item ([string]$Work.Item) -Message ("Could not write the meta sidecar: " + $_.Exception.Message) }
}

function Get-IQReportReusableExport {
    <#
    .SYNOPSIS
    On a resumed run, returns the meta sidecar of a complete earlier download for this report (crash before its checkpoint), else $null (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Work)
    if (-not $script:IQ.IsResume) { return $null }
    if (-not (Test-IQReportFileHasContent -Path ([string]$Work.FilePath))) { return $null }
    $meta = $null
    try { $meta = ConvertFrom-IQJsonFile -Path ([string]$Work.MetaPath) } catch { $meta = $null }
    if ($null -eq $meta) { return $null }
    if ([string](Get-IQReportMember -Object $meta -Name 'ReportId') -ne [string]$Work.ReportId) { return $null }
    $method = [string](Get-IQReportMember -Object $meta -Name 'ExportMethod')
    if ($method -notin @('IncludeModel', 'LiveConnect', 'getDefinition')) { return $null }
    return $meta
}

function Remove-IQReportLocalhostFolder {
    <#
    .SYNOPSIS
    Removes the "localhost" folders pbi-tools / Tabular Editor leave behind under Config\ and the base folder (monolith 3025-3026) (private).
    #>
    [CmdletBinding()]
    param()
    try {
        Remove-IQReportPath -Path (Join-Path ([string]$script:IQ.ConfigFolder) 'localhost')
        Remove-IQReportPath -Path (Join-Path ([string]$script:IQ.BaseFolder) 'localhost')
    }
    catch { $null = $null }
}

# =====================================================================================================================
# ReportBackup stage (brief 8.1)
# =====================================================================================================================

function Invoke-IQReportBackupStage {
    <#
    .SYNOPSIS
    ReportBackup stage body: exports every report in scope to "<Report Backups>\<RunId>" and extracts Pro-workspace models (itemKey = ReportId).
    .DESCRIPTION
    Port of monolith lines 2660-3045 with per-report checkpoints. Strategy per report (unchanged): paginated -> Export API
    LiveConnect (.rdl); Pro workspace -> IncludeModel (real PBIX, model extracted with pbi-tools) with getDefinition as the
    fallback; dedicated capacity -> LiveConnect first and getDefinition as the fallback, or getDefinition FIRST when the
    report carries a sensitivity label (Get-IQReportsWithSensitivityLabel). Pseudo workspaces ("My Workspace") use the
    group-less export route and never the Fabric fallback. Each report is checkpointed the moment it is done
    (Set-IQItemDone -Stage ReportBackup: outputs = exported file + .bim when produced; -Data carries DatasetId/BimPath for
    ModelDetail) and skipped on re-run. Returns @{ Total; Done; Skipped; Failed; AlreadyDone; ModelsExtracted }.
    #>
    [CmdletBinding()]
    param()
    $stage = 'ReportBackup'
    $summary = @{ Total = 0; Done = 0; Skipped = 0; Failed = 0; AlreadyDone = 0; ModelsExtracted = 0 }
    $runFolder = Get-IQReportRunFolder
    $modelFolder = Get-IQReportModelFolder
    $work = @(Get-IQReportWorkList -RunFolder $runFolder -ModelFolder $modelFolder)
    $summary.Total = $work.Count
    Write-IQLog -Level Info -Stage $stage -Message ("Report backup: {0} report(s) in scope; folder {1}" -f $work.Count, $runFolder)
    if ($work.Count -eq 0) {
        Write-IQReportExportSummary -RunFolder $runFolder | Out-Null
        return $summary
    }
    $pbiToolsState = Test-IQReportPbiToolsAvailable
    $proCount = @($work | Where-Object { -not $_.IsDedicated -and -not $_.IsPaginated -and -not $_.NoAccess }).Count
    if ($proCount -gt 0 -and -not $pbiToolsState.Available) {
        Write-IQLog -Level Warn -Stage $stage -Message ("{0} Pro-workspace report(s) in scope but models cannot be extracted from PBIX files: {1}. ModelDetail will use DAX for those models." -f $proCount, $pbiToolsState.Reason)
    }
    elseif ($proCount -gt 0 -and $pbiToolsState.Warning) {
        Write-IQLog -Level Warn -Stage $stage -Message ("{0} Pro-workspace report(s) in scope. {1}" -f $proCount, $pbiToolsState.Warning)
    }

    $index = 0
    foreach ($w in $work) {
        $index++
        if ($index % 10 -eq 0) { Write-IQLog -Level Info -Stage $stage -Message ("Progress: {0}/{1} reports" -f $index, $work.Count) }
        $item = [string]$w.Item
        try {
            if (Test-IQItemDone -Stage $stage -ItemKey $w.Key) {
                $summary.AlreadyDone++
                Write-IQLog -Level Debug -Stage $stage -Item $item -Message 'Already done (checkpoint); skipping'
                continue
            }
            if ($w.NoAccess) {
                Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $item -Status Skipped -Message 'No workspace access (shared report) - cannot be exported' -Data @{ ReportId = $w.ReportId; ReportName = $w.ReportName; WorkspaceId = $w.WorkspaceId; WorkspaceName = $w.WorkspaceName; DatasetId = $w.DatasetId; ReportType = $w.ReportType; FileName = $w.FileName } | Out-Null
                $summary.Skipped++
                continue
            }

            $started = [datetime]::UtcNow
            $method = $null
            $exported = $false
            $definition = $null
            $notes = @()
            $isGuid = Test-IQReportGuid -Value $w.WorkspaceId

            $reusable = Get-IQReportReusableExport -Work $w
            if ($null -ne $reusable) {
                $exported = $true
                $method = [string](Get-IQReportMember -Object $reusable -Name 'ExportMethod')
                $notes += 'existing download from an earlier attempt of this run reused'
                Write-IQLog -Level Info -Stage $stage -Item $item -Message ('Reusing ' + $w.FileName + ' (' + $method + ') from an earlier attempt')
            }
            else {
                Write-IQLog -Level Info -Stage $stage -Item $item -Message ('Exporting ' + $item)
                Remove-IQReportPath -Path $w.FilePath
                Remove-IQReportPath -Path $w.MetaPath
                if ($w.IsPaginated) {
                    # Paginated (RDL) report -> LiveConnect export (the Export API returns the RDL); no Fabric fallback.
                    $exported = Export-IQReportUsingApi -WorkspaceId $w.WorkspaceId -ReportId $w.ReportId -OutFilePath $w.FilePath -DownloadType LiveConnect -Item $item -Stage $stage
                    $method = 'LiveConnect'
                    if (-not $exported) { $notes += 'LiveConnect export failed for the paginated report' }
                }
                elseif (-not $w.IsDedicated) {
                    # Pro Workspace -> IncludeModel (real .pbix, used for model extraction below)
                    $exported = Export-IQReportUsingApi -WorkspaceId $w.WorkspaceId -ReportId $w.ReportId -OutFilePath $w.FilePath -DownloadType IncludeModel -Item $item -Stage $stage
                    $method = 'IncludeModel'
                    if (-not $exported) {
                        if ($isGuid) {
                            Write-IQLog -Level Warn -Stage $stage -Item $item -Message 'IncludeModel export failed; falling back to getDefinition.'
                            $definition = Export-IQReportDefinitionAsPbix -WorkspaceId $w.WorkspaceId -ReportId $w.ReportId -OutFilePath $w.FilePath -Item $item -Stage $stage
                            $exported = [bool]$definition.Success
                            $method = 'getDefinition'
                            if (-not $exported) { $notes += ('IncludeModel export failed; getDefinition also failed: ' + $definition.Message) }
                        }
                        else { $notes += 'IncludeModel export failed (no Fabric fallback for a pseudo workspace)' }
                    }
                }
                else {
                    # Dedicated capacity PBIX report. A report with a sensitivity label (only visible via the Fabric List
                    # Items API) prefers getDefinition FIRST so the protected definition is captured directly, falling
                    # back to LiveConnect; otherwise LiveConnect first with getDefinition as the fallback.
                    if ($w.HasSensitivityLabel -and $isGuid) {
                        $definition = Export-IQReportDefinitionAsPbix -WorkspaceId $w.WorkspaceId -ReportId $w.ReportId -OutFilePath $w.FilePath -Item $item -Stage $stage
                        $exported = [bool]$definition.Success
                        $method = 'getDefinition'
                        if (-not $exported) {
                            Write-IQLog -Level Warn -Stage $stage -Item $item -Message 'getDefinition export failed for the labelled report; falling back to LiveConnect.'
                            $exported = Export-IQReportUsingApi -WorkspaceId $w.WorkspaceId -ReportId $w.ReportId -OutFilePath $w.FilePath -DownloadType LiveConnect -Item $item -Stage $stage
                            $method = 'LiveConnect'
                            if (-not $exported) { $notes += ('getDefinition failed (' + $definition.Message + '); LiveConnect export also failed') }
                        }
                    }
                    else {
                        $exported = Export-IQReportUsingApi -WorkspaceId $w.WorkspaceId -ReportId $w.ReportId -OutFilePath $w.FilePath -DownloadType LiveConnect -Item $item -Stage $stage
                        $method = 'LiveConnect'
                        if (-not $exported) {
                            if ($isGuid) {
                                Write-IQLog -Level Warn -Stage $stage -Item $item -Message 'LiveConnect export failed; falling back to getDefinition.'
                                $definition = Export-IQReportDefinitionAsPbix -WorkspaceId $w.WorkspaceId -ReportId $w.ReportId -OutFilePath $w.FilePath -Item $item -Stage $stage
                                $exported = [bool]$definition.Success
                                $method = 'getDefinition'
                                if (-not $exported) { $notes += ('LiveConnect export failed; getDefinition also failed: ' + $definition.Message) }
                            }
                            else { $notes += 'LiveConnect export failed (no Fabric fallback for a pseudo workspace)' }
                        }
                    }
                }
            }

            $definitionFormat = ''
            if ($null -ne $definition -and $definition.Success) { $definitionFormat = [string]$definition.Format }
            elseif ($null -ne $reusable) { $definitionFormat = [string](Get-IQReportMember -Object $reusable -Name 'DefinitionFormat') }
            $baseData = @{
                ReportId = $w.ReportId; ReportName = $w.ReportName; ReportType = $w.ReportType; WorkspaceId = $w.WorkspaceId; WorkspaceName = $w.WorkspaceName
                DatasetId = $w.DatasetId; DatasetName = $w.DatasetName; DatasetWorkspaceId = $w.DatasetWorkspaceId; FileName = $w.FileName; FilePath = $w.FilePath
                IsPaginated = [bool]$w.IsPaginated; IsDedicated = [bool]$w.IsDedicated; HasSensitivityLabel = [bool]$w.HasSensitivityLabel; DefinitionFormat = $definitionFormat
            }
            if (-not $exported -or -not (Test-IQReportFileHasContent -Path $w.FilePath)) {
                Remove-IQReportPath -Path $w.FilePath
                Remove-IQReportPath -Path $w.MetaPath
                $message = 'Export failed'
                if ($notes.Count -gt 0) { $message = $notes -join '; ' }
                $baseData.DurationSec = [math]::Round(([datetime]::UtcNow - $started).TotalSeconds, 1)
                Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $item -Status Failed -Method $method -Message $message -Data $baseData | Out-Null
                $summary.Failed++
                continue
            }
            if ($null -eq $reusable) { Write-IQReportMetaFile -Work $w -Method $method -DefinitionFormat $definitionFormat }

            # Model extraction only for Pro workspaces AND only from a real IncludeModel PBIX (audit C7-11).
            $outputs = @([string]$w.FilePath)
            $modelExtract = ''
            $bimPath = $null
            if (-not $w.IsDedicated -and $method -eq 'IncludeModel' -and -not $w.IsPaginated) {
                $mx = Invoke-IQReportModelExtract -Work $w -Stage $stage
                $modelExtract = [string]$mx.Message
                if ($mx.Success) {
                    $bimPath = [string]$mx.BimPath
                    $outputs += $bimPath
                    if (-not $mx.Reused) { $summary.ModelsExtracted++ }
                }
            }
            elseif (-not $w.IsDedicated -and -not $w.IsPaginated) { $modelExtract = 'not attempted (export method ' + $method + ' does not contain the model)' }
            $baseData.BimPath = $bimPath
            $baseData.ModelBaseName = $w.ModelBaseName
            $baseData.ModelExtract = $modelExtract
            $baseData.SizeBytes = Get-IQReportFileSize -Path $w.FilePath
            $baseData.DurationSec = [math]::Round(([datetime]::UtcNow - $started).TotalSeconds, 1)
            if ($null -ne $definition -and $definition.Success) {
                $baseData.DefinitionParts = [int]$definition.PartCount
                $baseData.DefinitionPages = [int]$definition.PageCount
                $baseData.DefinitionVisuals = [int]$definition.VisualCount
                $baseData.DatasetReferenceType = [string]$definition.DatasetReferenceType
                $baseData.DatasetIdFromPbir = [string]$definition.DatasetIdFromPbir
            }
            $message = ''
            if ($notes.Count -gt 0) { $message = $notes -join '; ' }
            Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $item -Outputs $outputs -Method $method -Message $message -Data $baseData | Out-Null
            $summary.Done++
            Write-IQLog -Level Success -Stage $stage -Item $item -Message ("Exported {0} via {1} ({2:N0} bytes)" -f $w.FileName, $method, $baseData.SizeBytes)
        }
        catch {
            Write-IQLog -Level Error -Stage $stage -Item $item -Message ("Unexpected error: " + $_.Exception.Message) -Exception $_.Exception
            try {
                Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $item -Status Failed -Message ('Unexpected error: ' + $_.Exception.Message) -Data @{ ReportId = $w.ReportId; ReportName = $w.ReportName; WorkspaceId = $w.WorkspaceId; WorkspaceName = $w.WorkspaceName; DatasetId = $w.DatasetId; ReportType = $w.ReportType; FileName = $w.FileName } | Out-Null
            }
            catch { Write-IQLog -Level Error -Stage $stage -Item $item -Message ("Could not record the failure: " + $_.Exception.Message) }
            $summary.Failed++
        }
    }

    Write-IQReportExportSummary -RunFolder $runFolder | Out-Null
    Remove-IQReportLocalhostFolder
    Write-IQLog -Level Info -Stage $stage -Message ("Report backup finished: {0} exported ({1} models extracted), {2} skipped, {3} failed, {4} already done" -f $summary.Done, $summary.ModelsExtracted, $summary.Skipped, $summary.Failed, $summary.AlreadyDone)
    return $summary
}

# =====================================================================================================================
# ReportDetail stage (brief 8.2): the two csx runs + leftover VOL folder cleanup (monolith 3046-3135)
# =====================================================================================================================

function Get-IQReportDetailTxtName {
    <#
    .SYNOPSIS
    The eleven tab-separated files the Report Detail csx scripts write into the dated folder (private).
    #>
    [CmdletBinding()]
    param()
    return @('CustomVisuals.txt', 'ReportFilters.txt', 'PageFilters.txt', 'VisualFilters.txt', 'VisualObjects.txt', 'Visuals.txt', 'Bookmarks.txt', 'Pages.txt', 'Connections.txt', 'VisualInteractions.txt', 'ReportLevelMeasures.txt')
}

function Remove-IQReportDetailSubfolder {
    <#
    .SYNOPSIS
    Deletes every sub-folder of the run folder (the csx unzip folders / leftover VOL folders, monolith 3122-3134) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$RunFolder)
    $removed = 0
    if (-not (Test-Path -LiteralPath $RunFolder)) { return 0 }
    foreach ($sub in @(Get-ChildItem -LiteralPath $RunFolder -Directory -ErrorAction SilentlyContinue)) {
        try { Remove-Item -LiteralPath $sub.FullName -Recurse -Force -ErrorAction Stop; $removed++ }
        catch { Write-IQLog -Level Warn -Stage 'ReportDetail' -Message ("Could not remove leftover folder '{0}': {1}" -f $sub.FullName, $_.Exception.Message) }
    }
    return $removed
}

function Test-IQReportSamePath {
    <#
    .SYNOPSIS
    True when two paths point at the same location (normalised, case-insensitive) (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$A,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$B
    )
    if ([string]::IsNullOrWhiteSpace($A) -or [string]::IsNullOrWhiteSpace($B)) { return $false }
    try {
        $fa = [System.IO.Path]::GetFullPath($A).TrimEnd('\', '/')
        $fb = [System.IO.Path]::GetFullPath($B).TrimEnd('\', '/')
        return [string]::Equals($fa, $fb, [System.StringComparison]::OrdinalIgnoreCase)
    }
    catch { return $false }
}

function New-IQReportDetailWorkingFolder {
    <#
    .SYNOPSIS
    Builds a scratch working directory whose only "Report Backups\<date>" folder is a link to the run folder (private).
    .DESCRIPTION
    The csx scripts locate their input as "<CWD>\Report Backups\<newest yyyy-MM-dd folder>". When the run folder is not
    the newest dated folder (custom -RunId, or a newer folder exists on disk), a junction (Windows, no admin needed) or a
    symbolic link (other hosts) inside <Config>\Temp\rd-<hash>\Report Backups\<date> makes the run folder the one and
    only candidate. Returns @{ WorkingDirectory; LinkPath; Root } or $null when the link could not be created.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RunFolder,
        [Parameter(Mandatory = $true)][string]$DateName
    )
    $root = Join-Path (Get-IQReportTempFolder) ('rd-' + (Get-IQReportShortKey -Value ([string]$script:IQ.RunId)))
    $backups = Join-Path $root 'Report Backups'
    $link = Join-Path $backups $DateName
    try {
        Remove-IQReportDetailWorkingFolder -Info @{ WorkingDirectory = $root; LinkPath = $link; Root = $root }
        New-Item -ItemType Directory -Path $backups -Force | Out-Null
        # Junction first on Windows (no admin / developer mode needed), symbolic link otherwise or as the fallback.
        $linkTypes = @('SymbolicLink')
        if ($script:IQ.IsWindows) { $linkTypes = @('Junction', 'SymbolicLink') }
        $linkType = $null
        $lastError = $null
        foreach ($candidate in $linkTypes) {
            try {
                New-Item -ItemType $candidate -Path $link -Value $RunFolder -ErrorAction Stop | Out-Null
                if (Test-Path -LiteralPath $link) { $linkType = $candidate; break }
            }
            catch { $lastError = $_.Exception.Message; Remove-IQReportDetailWorkingFolder -Info @{ WorkingDirectory = $root; LinkPath = $link; Root = $root }; New-Item -ItemType Directory -Path $backups -Force | Out-Null }
        }
        if (-not $linkType) { throw ("link '{0}' was not created ({1})" -f $link, $lastError) }
        Write-IQLog -Level Debug -Stage 'ReportDetail' -Message ("Working folder {0}: '{1}' -> '{2}' ({3})" -f $root, $link, $RunFolder, $linkType)
        return @{ WorkingDirectory = $root; LinkPath = $link; Root = $root }
    }
    catch {
        Write-IQLog -Level Warn -Stage 'ReportDetail' -Message ("Could not create the working folder link '{0}' -> '{1}': {2}" -f $link, $RunFolder, $_.Exception.Message)
        Remove-IQReportDetailWorkingFolder -Info @{ WorkingDirectory = $root; LinkPath = $link; Root = $root }
        return $null
    }
}

function Remove-IQReportDetailWorkingFolder {
    <#
    .SYNOPSIS
    Removes the link (without touching its target) and the scratch working directory created by New-IQReportDetailWorkingFolder (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Info)
    if ($null -eq $Info) { return }
    $link = [string](Get-IQReportMember -Object $Info -Name 'LinkPath')
    $root = [string](Get-IQReportMember -Object $Info -Name 'Root')
    if ($link -and (Test-Path -LiteralPath $link)) {
        $done = $false
        try { [System.IO.Directory]::Delete($link, $false); $done = $true } catch { $done = $false }
        if (-not $done) { try { (Get-Item -LiteralPath $link -Force).Delete(); $done = $true } catch { $done = $false } }
        if (-not $done -and $script:IQ.IsWindows) {
            try {
                $cmd = Join-Path $env:SystemRoot 'System32\cmd.exe'
                & $cmd /c "rmdir `"$link`"" 2>&1 | Out-Null
                $done = -not (Test-Path -LiteralPath $link)
            }
            catch { $done = $false }
        }
        if (-not $done) {
            Write-IQLog -Level Warn -Stage 'ReportDetail' -Message "Could not remove the link '$link'; leaving the working folder in place"
            return
        }
    }
    if ($root -and (Test-Path -LiteralPath $root)) { Remove-IQReportPath -Path $root }
}

function Invoke-IQReportDetailStage {
    <#
    .SYNOPSIS
    ReportDetail stage body: runs the PBIR and classic Report Detail csx scripts once against Blank Model.bim (itemKey = all).
    .DESCRIPTION
    Port of monolith lines 3046-3135. "Report Detail Extract Script-PBIR.csx" (unzips every .pbix/.pbit in the newest dated
    folder under <CWD>\Report Backups and writes the eleven tab-separated TXT files with headers) runs first, then
    "Report Detail Extract Script.csx" (appends the classic-Layout rows), both through Invoke-IQTabularEditor with
    WorkingDirectory = BaseFolder and a timeout of ToolTimeoutMinutes * 3. When the run folder is not the newest dated
    folder a scratch working directory with a junction makes it so. Stale TXT files and leftover unzip folders are removed
    before the run (audit C8-08/X4-13); the classic script only runs after the PBIR script succeeded; leftover VOL
    sub-folders are deleted afterwards. Failure -> checkpoint Failed (stage CompletedWithErrors); Assemble still builds
    the workbook from whatever *.txt exist. Returns @{ Status; PbixCount; TxtFiles; Message }.
    #>
    [CmdletBinding()]
    param()
    $stage = 'ReportDetail'
    $key = 'all'
    $runFolder = Get-IQReportRunFolder
    $summary = @{ Status = 'Pending'; PbixCount = 0; TxtFiles = @(); Message = '' }
    if (Test-IQItemDone -Stage $stage -ItemKey $key) {
        Write-IQLog -Level Info -Stage $stage -Message 'Report detail already extracted (checkpoint); skipping'
        $summary.Status = 'AlreadyDone'
        return $summary
    }
    Write-IQReportExportSummary -RunFolder $runFolder | Out-Null   # keep ReportExports.txt current even if ReportBackup crashed
    $pbixFiles = @(Get-ChildItem -LiteralPath $runFolder -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ieq '.pbix' -or $_.Extension -ieq '.pbit' })
    $summary.PbixCount = $pbixFiles.Count
    Write-IQLog -Level Info -Stage $stage -Message ("Report detail extraction: {0} PBIX/PBIT file(s) in {1}" -f $pbixFiles.Count, $runFolder)

    if (-not (Test-IQReportTabularEditorAvailable)) {
        $reason = Get-IQReportTabularEditorReason
        if ($pbixFiles.Count -eq 0) {
            Set-IQItemDone -Stage $stage -ItemKey $key -Item 'Report Detail' -Status Skipped -Message ('No PBIX files to process; ' + $reason) -Data @{ PbixCount = 0 } | Out-Null
            $summary.Status = 'Skipped'
        }
        else {
            Set-IQItemDone -Stage $stage -ItemKey $key -Item 'Report Detail' -Status Failed -Message ($reason + '; the Report Detail csx scripts cannot run') -Data @{ PbixCount = $pbixFiles.Count } | Out-Null
            $summary.Status = 'Failed'
        }
        $summary.Message = $reason
        return $summary
    }

    $configFolder = [string]$script:IQ.ConfigFolder
    $blankModel = Join-Path $configFolder 'Blank Model.bim'
    $script1 = Join-Path $configFolder 'Report Detail Extract Script-PBIR.csx'
    $script2 = Join-Path $configFolder 'Report Detail Extract Script.csx'
    $missing = @()
    foreach ($f in @($blankModel, $script1, $script2)) { if (-not (Test-Path -LiteralPath $f -PathType Leaf)) { $missing += $f } }
    if ($missing.Count -gt 0) {
        $message = 'Required file(s) missing: ' + ($missing -join ', ')
        Set-IQItemDone -Stage $stage -ItemKey $key -Item 'Report Detail' -Status Failed -Message $message -Data @{ PbixCount = $pbixFiles.Count } | Out-Null
        $summary.Status = 'Failed'; $summary.Message = $message
        return $summary
    }
    try {
        if ((Select-String -LiteralPath $script2 -Pattern 'Expresssion' -SimpleMatch -Quiet)) {
            Write-IQLog -Level Warn -Stage $stage -Message "'Report Detail Extract Script.csx' still contains the misspelled JSON path 'Expresssion' (audit X3-B2); some column/measure lineage rows will be missing"
        }
    }
    catch { $null = $null }

    $timeout = 20
    try { $timeout = [int](Get-IQReportOption -Name 'ToolTimeoutMinutes' -Default 20) } catch { $timeout = 20 }
    if ($timeout -lt 1) { $timeout = 1 }
    $timeout = $timeout * 3

    # Working directory: BaseFolder when the run folder is the newest dated folder (monolith behaviour), else a scratch
    # folder whose only "Report Backups\<date>" entry is a link to the run folder.
    $dateName = [string]$script:IQ.RunId
    if ($dateName -notmatch '^\d{4}-\d{2}-\d{2}$') { $dateName = (Get-Date).ToString('yyyy-MM-dd') }
    $workingDirectory = [string]$script:IQ.BaseFolder
    $linkInfo = $null
    $newest = $null
    try { $newest = Get-IQDateFolder -Root ([string]$script:IQ.Paths.ReportBackups) } catch { $newest = $null }
    if (-not (Test-IQReportSamePath -A $newest -B $runFolder)) {
        Write-IQLog -Level Info -Stage $stage -Message ("Run folder '{0}' is not the newest dated folder under Report Backups; using a linked working folder" -f $runFolder)
        $linkInfo = New-IQReportDetailWorkingFolder -RunFolder $runFolder -DateName $dateName
        if ($null -eq $linkInfo) {
            $message = 'The run folder is not the newest dated folder under Report Backups and no link could be created; the csx scripts would process the wrong folder'
            Set-IQItemDone -Stage $stage -ItemKey $key -Item 'Report Detail' -Status Failed -Message $message -Data @{ PbixCount = $pbixFiles.Count; NewestDateFolder = [string]$newest } | Out-Null
            $summary.Status = 'Failed'; $summary.Message = $message
            return $summary
        }
        $workingDirectory = [string]$linkInfo.WorkingDirectory
    }

    # Fresh start for the extraction: stale TXT files from an earlier attempt and leftover unzip folders are removed.
    foreach ($name in @(Get-IQReportDetailTxtName)) { Remove-IQReportPath -Path (Join-Path $runFolder $name) }
    $removedBefore = Remove-IQReportDetailSubfolder -RunFolder $runFolder
    if ($removedBefore -gt 0) { Write-IQLog -Level Debug -Stage $stage -Message "Removed $removedBefore leftover folder(s) before extraction" }

    $previousEnv = @{ IMPACTIQ_BASE = $env:IMPACTIQ_BASE; IMPACTIQ_DATE_FOLDER = $env:IMPACTIQ_DATE_FOLDER; IMPACTIQ_REPORT_DATE = $env:IMPACTIQ_REPORT_DATE }
    $r1 = $null
    $r2 = $null
    try {
        $env:IMPACTIQ_BASE = [string]$script:IQ.BaseFolder
        $env:IMPACTIQ_DATE_FOLDER = $runFolder
        $env:IMPACTIQ_REPORT_DATE = $dateName

        Write-IQLog -Level Info -Stage $stage -Message 'Running Report Detail Extract Script-PBIR.csx (Tabular Editor)'
        $args1 = ('"{0}" -S "{1}"' -f $blankModel, $script1)
        $r1 = Invoke-IQTabularEditor -ArgumentList $args1 -WorkingDirectory $workingDirectory -TimeoutMinutes $timeout -LogName 'report-detail-pbir' -Stage $stage -Item 'PBIR script'
        $visualsTxt = Join-Path $runFolder 'Visuals.txt'
        if ($r1.Success -and -not (Test-Path -LiteralPath $visualsTxt)) {
            Write-IQLog -Level Warn -Stage $stage -Message "The PBIR script finished but wrote no Visuals.txt into $runFolder"
        }
        if ($r1.Success -or (Test-Path -LiteralPath $visualsTxt)) {
            Write-IQLog -Level Info -Stage $stage -Message 'Running Report Detail Extract Script.csx (classic Layout, Tabular Editor)'
            $args2 = ('"{0}" -S "{1}"' -f $blankModel, $script2)
            $r2 = Invoke-IQTabularEditor -ArgumentList $args2 -WorkingDirectory $workingDirectory -TimeoutMinutes $timeout -LogName 'report-detail-classic' -Stage $stage -Item 'Classic script'
        }
        else {
            Write-IQLog -Level Warn -Stage $stage -Message 'Skipping the classic Layout script because the PBIR script failed and produced no TXT files (it would append header-less rows)'
        }
    }
    finally {
        foreach ($k in @($previousEnv.Keys)) {
            $v = $previousEnv[$k]
            if ($null -eq $v) { Remove-Item -Path ('Env:' + $k) -ErrorAction SilentlyContinue }
            else { Set-Item -Path ('Env:' + $k) -Value $v }
        }
        $removedAfter = Remove-IQReportDetailSubfolder -RunFolder $runFolder
        if ($removedAfter -gt 0) { Write-IQLog -Level Info -Stage $stage -Message "Removed $removedAfter leftover VOL folder(s) from $runFolder" }
        Remove-IQReportDetailWorkingFolder -Info $linkInfo
        Remove-IQReportLocalhostFolder
    }

    $txtFiles = @()
    foreach ($name in @(Get-IQReportDetailTxtName)) {
        $p = Join-Path $runFolder $name
        if (Test-Path -LiteralPath $p -PathType Leaf) { $txtFiles += $p }
    }
    $summary.TxtFiles = $txtFiles
    $problems = @()
    if ($null -eq $r1 -or -not $r1.Success) { $problems += ('PBIR script failed: ' + (Get-IQReportProcessSummary -Result $r1)) }
    if ($null -ne $r1 -and $r1.Success -and $null -eq $r2) { $problems += 'classic script did not run' }
    if ($null -ne $r2 -and -not $r2.Success) { $problems += ('classic script failed: ' + (Get-IQReportProcessSummary -Result $r2)) }
    if ($txtFiles.Count -eq 0) { $problems += 'no TXT files were produced' }
    $data = @{
        PbixCount = $pbixFiles.Count; TxtFiles = @($txtFiles | ForEach-Object { Split-Path -Path $_ -Leaf }); WorkingDirectory = $workingDirectory
        Csx1ExitCode = $null; Csx1DurationSec = $null; Csx2ExitCode = $null; Csx2DurationSec = $null
    }
    if ($null -ne $r1) { $data.Csx1ExitCode = $r1.ExitCode; $data.Csx1DurationSec = $r1.DurationSec }
    if ($null -ne $r2) { $data.Csx2ExitCode = $r2.ExitCode; $data.Csx2DurationSec = $r2.DurationSec }
    if ($problems.Count -eq 0) {
        Set-IQItemDone -Stage $stage -ItemKey $key -Item 'Report Detail' -Outputs $txtFiles -Method 'TabularEditor' -Message ("{0} PBIX file(s) processed, {1} TXT file(s) written" -f $pbixFiles.Count, $txtFiles.Count) -Data $data | Out-Null
        $summary.Status = 'Succeeded'
        Write-IQLog -Level Success -Stage $stage -Message ("Report detail extraction finished: {0} TXT file(s) in {1}" -f $txtFiles.Count, $runFolder)
    }
    else {
        $message = $problems -join '; '
        Set-IQItemDone -Stage $stage -ItemKey $key -Item 'Report Detail' -Outputs $txtFiles -Status Failed -Method 'TabularEditor' -Message $message -Data $data | Out-Null
        $summary.Status = 'Failed'
        $summary.Message = $message
    }
    return $summary
}
