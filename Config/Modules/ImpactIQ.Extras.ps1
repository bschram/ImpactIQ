# ImpactIQ.Extras.ps1 - Extras stage: optional admin-API collectors (-IncludeAdminApis) and usage metrics (-IncludeUsageMetrics).
#
# Contract: brief sections 2.7 and 8.4. Nothing in this file existed in the monolith; every collector is opt-in, every
# collector is its own item checkpoint (stage 'Extras'), raw API responses are kept under State\runs\<RunId>\extracts\
# (admin\ and usage\) so a re-run never repeats a finished call, and the flattened rows are written with
# Save-IQInventory -Name 'extras-<SheetName>' so the Assemble stage can add one worksheet per file.
#
# Item checkpoints (Test-IQItemDone / Set-IQItemDone -Stage Extras):
#   admin-groups              GET admin/groups?$top=<n>&$skip=<k>&$expand=users,reports,datasets,dataflows,dashboards (paged with $skip)
#   admin-scan                Scanner API: GET admin/workspaces/modified -> POST admin/workspaces/getInfo (batches of 100,
#                             lineage/datasourceDetails/datasetSchema/datasetExpressions/getArtifactUsers) -> GET scanStatus/{id}
#                             polled until Succeeded -> GET scanResult/{id}; each batch result is cached on disk and reused
#   admin-activity-<yyyy-MM-dd>  GET admin/activityevents?startDateTime='<day>T00:00:00.000Z'&endDateTime='<day>T23:59:59.999Z'
#                             (continuationUri paging), one item per UTC day for the last -ActivityDays days (max 28)
#   usage-<workspaceId>       Get-IQUsageMetrics (ImpactIQ.Dax.ps1) per real workspace in scope
# The admin collectors run only when GET admin/capacities?$top=1 succeeds (Fabric administrator). A definite "no"
# (HTTP 401/403/404 -> $null) checkpoints admin-groups and admin-scan as Skipped with the reason and the run continues;
# a probe that fails for another reason (5xx after retries, network, expired token) is transient: the two items are
# checkpointed Failed (manifest.failures, stage CompletedWithErrors, exit 2) so the next start re-runs the stage and
# retries them. A Skipped checkpoint is not treated as final either: when a later run of the stage passes the probe,
# the collector runs for real.
#
# Raw JSON (never deleted by this module; API bodies are written verbatim, never re-serialised):
#   extracts\admin\groups-<page>.json            one file per admin/groups page
#   extracts\admin\scan-workspaces.json          the workspace ids that were scanned (with the batch layout)
#   extracts\admin\scan-<batch>-<hash8>.json     one scanResult per batch of <= 100 workspace ids (reused on re-run;
#                                                the ids are sorted before batching so a retry produces the same
#                                                batches, and a cached file is matched by its <hash8> even when its
#                                                batch number moved)
#   extracts\admin\activity-<yyyy-MM-dd>.json    { Date; EventCount; Pages; Partial; Events[] } per UTC day
#   extracts\usage\<safeWorkspaceId>.json        the Get-IQUsageMetrics result of one workspace
#
# Inventory files for Assemble (Save-IQInventory -Name 'extras-<SheetName>' -> inventory\extras-<sheetname>.json), each
# a self-describing object { SchemaVersion=1; Collector; SheetName; Columns[]; RowCount; Rows[]; CollectedUtc; Message }
# where every row carries every column (nested API values are flattened to compact JSON strings, lists of ids are
# joined with ";"). Sheet names and columns (unmapped scalar API fields are appended under their raw names, as the
# monolith's Rename-Properties did):
#   AdminWorkspaces      WorkspaceId, WorkspaceName, WorkspaceType, WorkspaceState, WorkspaceIsReadOnly,
#                        WorkspaceIsOnDedicatedCapacity, WorkspaceCapacityId, WorkspaceCapacityMigrationStatus,
#                        WorkspaceDescription, WorkspaceDefaultDatasetStorageFormat, WorkspaceHasWorkspaceLevelSettings,
#                        WorkspacePipelineId, WorkspaceUserCount, WorkspaceReportCount, WorkspaceDatasetCount,
#                        WorkspaceDataflowCount, WorkspaceDashboardCount, WorkspaceReportIds, WorkspaceDatasetIds,
#                        WorkspaceDataflowIds, WorkspaceDashboardIds
#   AdminWorkspaceUsers  WorkspaceId, WorkspaceName, UserEmailAddress, UserDisplayName, UserIdentifier, UserGraphId,
#                        UserPrincipalType, UserGroupUserAccessRight, UserType, UserProfileJson
#   ScanWorkspaces       (additive) WorkspaceId, WorkspaceName, WorkspaceType, WorkspaceState, WorkspaceIsOnDedicatedCapacity,
#                        WorkspaceCapacityId, WorkspaceDefaultDatasetStorageFormat, WorkspaceDescription,
#                        WorkspaceReportCount, WorkspaceDashboardCount, WorkspaceDatasetCount, WorkspaceDataflowCount,
#                        WorkspaceDatamartCount, WorkspaceUserCount
#   ScanDatasets         WorkspaceId, WorkspaceName, DatasetId, DatasetName, DatasetDescription, DatasetConfiguredBy,
#                        DatasetConfiguredById, DatasetCreatedDate, DatasetContentProviderType, DatasetTargetStorageMode,
#                        DatasetIsRefreshable, DatasetIsEffectiveIdentityRequired, DatasetIsEffectiveIdentityRolesRequired,
#                        DatasetIsOnPremGatewayRequired, DatasetEndorsement, DatasetCertifiedBy, DatasetSensitivityLabelId,
#                        DatasetDatasourceInstanceIds, DatasetUpstreamDataflowIds, DatasetUpstreamDatamartIds,
#                        DatasetTableCount, DatasetColumnCount, DatasetMeasureCount, DatasetRelationshipCount,
#                        DatasetExpressionCount, DatasetRoleNames, DatasetUserCount, DatasetSchemaMayNotBeUpToDate,
#                        DatasetSchemaRetrievalError, DatasetRefreshScheduleJson, DatasetDirectQueryRefreshScheduleJson
#   ScanTables           WorkspaceId, WorkspaceName, DatasetId, DatasetName, TableName, TableDescription, TableIsHidden,
#                        TableStorageMode, TableSourceExpression, TableColumnCount, TableMeasureCount
#   ScanColumns          WorkspaceId, WorkspaceName, DatasetId, DatasetName, TableName, ColumnName, ColumnDataType,
#                        ColumnType, ColumnIsHidden, ColumnExpression, ColumnDescription
#   ScanMeasures         WorkspaceId, WorkspaceName, DatasetId, DatasetName, TableName, MeasureName, MeasureExpression,
#                        MeasureIsHidden, MeasureDescription
#   ScanExpressions      (additive) WorkspaceId, WorkspaceName, DatasetId, DatasetName, ExpressionName, Expression,
#                        ExpressionDescription
#   ScanDatasources      DatasourceInstanceId, DatasourceType, DatasourceGatewayId, DatasourceServer, DatasourceDatabase,
#                        DatasourceUrl, DatasourcePath, DatasourceKind, DatasourceAccount, DatasourceDomain,
#                        DatasourceEmailAddress, DatasourceLoginServer, DatasourceClassInfo, DatasourceConnectionDetailsJson,
#                        DatasourceIsMisconfigured, DatasourceUsedByDatasetCount, DatasourceUsedByDataflowCount
#   ScanReports          WorkspaceId, WorkspaceName, ReportId, ReportName, ReportType, ReportDatasetId, ReportDatasetWorkspaceId,
#                        ReportCreatedDateTime, ReportModifiedDateTime, ReportCreatedBy, ReportCreatedById, ReportModifiedBy,
#                        ReportModifiedById, ReportAppId, ReportOriginalReportObjectId, ReportDescription, ReportEndorsement,
#                        ReportCertifiedBy, ReportSensitivityLabelId, ReportUserCount
#   ScanDashboards       WorkspaceId, WorkspaceName, DashboardId, DashboardName, DashboardIsReadOnly, DashboardAppId,
#                        DashboardTileCount, DashboardTileReportIds, DashboardTileDatasetIds, DashboardSensitivityLabelId,
#                        DashboardUserCount
#   ScanDataflows        WorkspaceId, WorkspaceName, DataflowId, DataflowName, DataflowDescription, DataflowConfiguredBy,
#                        DataflowModifiedBy, DataflowModifiedDateTime, DataflowGeneration, DataflowEndorsement,
#                        DataflowCertifiedBy, DataflowSensitivityLabelId, DataflowDatasourceInstanceIds,
#                        DataflowUpstreamDataflowIds, DataflowRefreshScheduleJson, DataflowUserCount
#   ScanUsers            WorkspaceId, WorkspaceName, ArtifactType (Workspace|Report|Dashboard|Dataset|Dataflow|Datamart),
#                        ArtifactId, ArtifactName, UserEmailAddress, UserDisplayName, UserIdentifier, UserGraphId,
#                        UserPrincipalType, UserType, UserAccessRight
#   ActivityEvents       Id, CreationTime, Operation, Activity, UserId, UserType, UserKey, Workload, ItemName, WorkSpaceName,
#                        WorkspaceId, CapacityId, CapacityName, DatasetName, DatasetId, ReportName, ReportId, ReportType,
#                        ArtifactKind, ArtifactId, ArtifactName, ObjectId, DistributionMethod, ConsumptionMethod, ClientIP,
#                        UserAgent, ActivityId, RequestId, IsSuccess, RefreshType, DataflowName, DataflowId, DashboardName,
#                        DashboardId, AppName, AppId, RecordType, OrganizationId, ActivityDate (UTC day of the request),
#                        + every other field the event carries (raw name; nested values as JSON)
#   UsageReportViews     WorkspaceId, WorkspaceName, UsageDatasetId, + the 'Report views' columns of the usage model
#                        (Date, ReportId, ReportName, UserId, UserKey, ConsumptionMethod, DistributionMethod, ReportType,
#                        AppName, CapacityId, CapacityName, DatasetName, UserAgent, CreationTime, ... as returned) and
#                        UserPrincipalName when the Users table exposes it
#   UsageReportPageViews WorkspaceId, WorkspaceName, UsageDatasetId, + the 'Report page views' columns (Date, ReportId,
#                        ReportName, SectionId, UserId, UserKey, Client, SessionSource, WorkspaceId(model), Timestamp, ...)
#
# Optional $script:IQ.Options keys (not entry-point parameters; defaults in Get-IQExtrasOption callers):
#   ActivityDays (entry point, default 30 -> clamped to ActivityMaxDays=28; the day window ends on the run's start
#   date - manifest startedUtc on a resume, Options.NowUtc / the clock otherwise - so a resumed run keeps its window),
#   ActivityMaxRows (250000 on PowerShell 7, 100000 on Windows PowerShell 5.1 whose JSON serialiser is much slower),
#   UsageDays (30), AdminGroupsPageSize (5000), AdminScanScopeOnly ($false: scan the whole tenant; $true: only
#   workspaces in scope), AdminScanMaxWorkspaces (0 = all), AdminScanBatchSize (100), AdminScanTimeoutMinutes (30),
#   ExtrasPollSeconds (10; the scanStatus poll honours a larger Retry-After header).
#
# Time budget (Options.TimeBudgetMinutes): Test-IQTimeBudget is consulted before every collector, before every Scanner
# batch and inside the scanStatus poll loop. A collector that stops on the budget is NOT checkpointed (the cached
# batches / day files stay on disk) and reports 'Paused', so Invoke-IQStage marks the stage Paused and the next start
# resumes it.
#
# Windows PowerShell 5.1 and PowerShell 7 compatible; nothing here is Windows-only. Dot-sourced from ImpactIQ.ps1, so
# $script:IQ is the shared context. Cross-module functions used (brief section 2): Write-IQLog, Get-IQSafeKey,
# Get-IQNowUtc, ConvertTo-IQJsonFile, ConvertFrom-IQJsonFile, Invoke-IQApi, Test-IQTimeBudget, Test-IQItemDone,
# Set-IQItemDone, Get-IQItemCheckpoint, Save-IQInventory, Get-IQSelectedWorkspaces, Get-IQUsageMetrics. The scanStatus
# poll reads the Retry-After header, which Invoke-IQApi does not expose, so it calls the Http module's request core
# (Get-IQApiUrl, Invoke-IQHttpRequest, Get-IQHttpHeaderValue, ConvertTo-IQRetryAfterDelay) directly. Private helpers
# are prefixed *-IQExtras* / *-IQAdmin* / *-IQScan* / *-IQUsage* and are not part of the contract.

# =====================================================================================================================
# Small private helpers
# =====================================================================================================================

function Get-IQExtrasMember {
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

function Get-IQExtrasMemberName {
    <#
    .SYNOPSIS
    The member names of a hashtable/dictionary or object, in their natural order (private). Wrap in @( ).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Object)
    $names = @()
    if ($null -eq $Object) { return $names }
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($k in $Object.Keys) { $names += [string]$k }
        return $names
    }
    foreach ($p in $Object.PSObject.Properties) { $names += $p.Name }
    return $names
}

function Get-IQExtrasOption {
    <#
    .SYNOPSIS
    Reads an entry from $script:IQ.Options with a default (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $false)][AllowNull()]$Default
    )
    if ($script:IQ -and $script:IQ.Options -and $script:IQ.Options.Contains($Name)) {
        $v = $script:IQ.Options[$Name]
        if ($null -ne $v -and -not ($v -is [string] -and [string]::IsNullOrWhiteSpace($v))) { return $v }
    }
    return $Default
}

function Get-IQExtrasFolder {
    <#
    .SYNOPSIS
    State\runs\<RunId>\extracts\<Name> (admin | usage), created when missing (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateSet('admin', 'usage')][string]$Name)
    if (-not $script:IQ) { throw 'ImpactIQ context is not initialised (Initialize-IQContext).' }
    $extracts = $null
    if ($script:IQ.ContainsKey('RunPaths') -and $null -ne $script:IQ.RunPaths -and $script:IQ.RunPaths.Extracts) { $extracts = [string]$script:IQ.RunPaths.Extracts }
    if ([string]::IsNullOrWhiteSpace($extracts)) {
        if ([string]::IsNullOrWhiteSpace([string]$script:IQ.RunPath)) { throw 'No active run (Initialize-IQRun has not been called).' }
        $extracts = Join-Path $script:IQ.RunPath 'extracts'
    }
    $folder = Join-Path $extracts $Name
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    return $folder
}

function Test-IQExtrasGuid {
    <#
    .SYNOPSIS
    $true when the value parses as a GUID (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $g = [guid]::Empty
    return [guid]::TryParse($Value, [ref]$g)
}

function Test-IQExtrasFileHasContent {
    <#
    .SYNOPSIS
    $true when the file exists and is larger than 0 bytes (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    try { return ((Get-Item -LiteralPath $Path).Length -gt 0) } catch { return $false }
}

function Start-IQExtrasSleep {
    <#
    .SYNOPSIS
    Sleeps for the given number of seconds (0 = no-op); separate function so tests can mock it (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][int]$Seconds)
    if ($Seconds -le 0) { return }
    Start-Sleep -Seconds $Seconds
}

function Write-IQExtrasRawFile {
    <#
    .SYNOPSIS
    Atomically writes an API body verbatim (UTF-8 without BOM, "<path>.tmp" then Move-Item -Force) so a raw cache never goes through ConvertTo-Json (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text,
        [Parameter(Mandatory = $true)][string]$Path
    )
    $dir = Split-Path -Path $Path -Parent
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = $Path + '.tmp'
    [System.IO.File]::WriteAllText($tmp, $Text, (New-Object System.Text.UTF8Encoding($false)))
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}

function Get-IQExtrasNowUtc {
    <#
    .SYNOPSIS
    The run's clock: Get-IQNowUtc (Options.NowUtc-aware) when the Common module exposes it, else [datetime]::UtcNow (private).
    #>
    [CmdletBinding()]
    param()
    if (Get-Command -Name 'Get-IQNowUtc' -ErrorAction SilentlyContinue) {
        try { return ([datetime](Get-IQNowUtc)).ToUniversalTime() } catch { return [datetime]::UtcNow }
    }
    return [datetime]::UtcNow
}

function ConvertTo-IQExtrasScalar {
    <#
    .SYNOPSIS
    Makes a value Excel-friendly: scalars pass through, nested objects/arrays become compact JSON, empty arrays become $null (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string] -or $Value -is [bool] -or $Value -is [datetime] -or $Value -is [decimal] -or $Value.GetType().IsPrimitive) { return $Value }
    if ($Value -is [guid] -or $Value -is [System.Uri] -or $Value -is [timespan]) { return [string]$Value }
    if ($Value -is [System.Collections.IEnumerable]) {
        $items = @($Value)
        if ($items.Count -eq 0) { return $null }
    }
    try { return (ConvertTo-Json -InputObject $Value -Depth 20 -Compress) } catch { return [string]$Value }
}

function ConvertTo-IQExtrasJoined {
    <#
    .SYNOPSIS
    Joins one property of every element of a collection with ";" (empty string when nothing) (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Items,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Property
    )
    $values = New-Object System.Collections.Generic.List[string]
    foreach ($i in @($Items)) {
        if ($null -eq $i) { continue }
        $v = $i
        if (-not [string]::IsNullOrEmpty($Property)) { $v = Get-IQExtrasMember -Object $i -Name $Property }
        if ($null -eq $v) { continue }
        $s = [string]$v
        if ($s.Length -gt 0 -and -not $values.Contains($s)) { $values.Add($s) }
    }
    return ($values -join ';')
}

function ConvertTo-IQExtrasRow {
    <#
    .SYNOPSIS
    Builds an ordered row: -Lead columns first, then -Map (raw property -> column) values, then every unmapped scalar API field under its raw name (private).
    .DESCRIPTION
    Mirrors the monolith's Rename-Properties behaviour (mapped names are always present, even when null; unmapped
    fields are kept). -Exclude lists raw properties that are consumed elsewhere (nested collections turned into
    counts / joined ids / their own sheets) so they are not duplicated as JSON blobs.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Object,
        [Parameter(Mandatory = $false)][AllowNull()][System.Collections.IDictionary]$Lead,
        [Parameter(Mandatory = $false)][AllowNull()][System.Collections.IDictionary]$Map,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$Exclude
    )
    $row = [ordered]@{}
    if ($null -ne $Lead) { foreach ($k in $Lead.Keys) { $row[[string]$k] = $Lead[$k] } }
    $mapped = @{}
    if ($null -ne $Map) {
        foreach ($raw in $Map.Keys) {
            $row[[string]$Map[$raw]] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $Object -Name ([string]$raw))
            $mapped[[string]$raw] = $true
        }
    }
    foreach ($e in @($Exclude)) { if ($null -ne $e) { $mapped[[string]$e] = $true } }
    # The unmapped fields are the bulk of every activity event (hundreds of thousands of rows on a busy tenant), so the
    # members are read directly and ConvertTo-IQExtrasScalar is only called for values that are not already scalar:
    # a PowerShell function call costs ~100-200 us on Windows PowerShell 5.1.
    if ($null -ne $Object) {
        if ($Object -is [System.Collections.IDictionary]) {
            foreach ($k in $Object.Keys) {
                $name = [string]$k
                if ($mapped.ContainsKey($name) -or $row.Contains($name)) { continue }
                $v = $Object[$k]
                if ($null -eq $v -or $v -is [string] -or $v -is [bool] -or $v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [datetime] -or $v -is [decimal]) { $row[$name] = $v }
                else { $row[$name] = ConvertTo-IQExtrasScalar -Value $v }
            }
        }
        else {
            foreach ($p in $Object.PSObject.Properties) {
                $name = $p.Name
                if ($mapped.ContainsKey($name) -or $row.Contains($name)) { continue }
                $v = $p.Value
                if ($null -eq $v -or $v -is [string] -or $v -is [bool] -or $v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [datetime] -or $v -is [decimal]) { $row[$name] = $v }
                else { $row[$name] = ConvertTo-IQExtrasScalar -Value $v }
            }
        }
    }
    return $row
}

function New-IQExtrasSheet {
    <#
    .SYNOPSIS
    Normalises rows into a self-describing sheet object { SchemaVersion; Collector; SheetName; Columns; RowCount; Rows; CollectedUtc; Message } where every row has every column (private).
    .DESCRIPTION
    Columns = -PreferredColumns (those that occur in at least one row, or all of them when -KeepPreferred) followed by
    every other property in order of first appearance. Nested values are flattened with ConvertTo-IQExtrasScalar
    (values that are already scalar - every cell of a ConvertTo-IQExtrasRow row - are copied without a function call,
    which keeps a 250 000-row ActivityEvents sheet feasible on Windows PowerShell 5.1).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SheetName,
        [Parameter(Mandatory = $true)][string]$Collector,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()]$Rows,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$PreferredColumns,
        [Parameter(Mandatory = $false)][switch]$KeepPreferred,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Message
    )
    $list = @($Rows | Where-Object { $null -ne $_ })
    $seen = @{}
    $columns = New-Object System.Collections.Generic.List[string]
    if ($KeepPreferred) {
        foreach ($c in @($PreferredColumns)) { if ($c -and -not $seen.ContainsKey($c)) { $seen[$c] = $true; $columns.Add($c) } }
    }
    # One pass over the rows collects the member names in order of first appearance (no per-row function calls).
    $present = @{}
    $order = New-Object System.Collections.Generic.List[string]
    foreach ($r in $list) {
        if ($r -is [System.Collections.IDictionary]) {
            foreach ($k in $r.Keys) { $n = [string]$k; if (-not $present.ContainsKey($n)) { $present[$n] = $true; $order.Add($n) } }
        }
        else {
            foreach ($p in $r.PSObject.Properties) { $n = $p.Name; if (-not $present.ContainsKey($n)) { $present[$n] = $true; $order.Add($n) } }
        }
    }
    foreach ($c in @($PreferredColumns)) { if ($c -and $present.ContainsKey($c) -and -not $seen.ContainsKey($c)) { $seen[$c] = $true; $columns.Add($c) } }
    foreach ($n in $order) { if (-not $seen.ContainsKey($n)) { $seen[$n] = $true; $columns.Add($n) } }
    $uniform = New-Object System.Collections.Generic.List[object]
    foreach ($r in $list) {
        $o = [ordered]@{}
        $isDict = ($r -is [System.Collections.IDictionary])
        foreach ($c in $columns) {
            $v = $null
            if ($isDict) { if ($r.Contains($c)) { $v = $r[$c] } }
            else { $prop = $r.PSObject.Properties[$c]; if ($null -ne $prop) { $v = $prop.Value } }
            if ($null -ne $v -and -not ($v -is [string] -or $v -is [bool] -or $v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [datetime] -or $v -is [decimal])) { $v = ConvertTo-IQExtrasScalar -Value $v }
            $o[$c] = $v
        }
        $uniform.Add([PSCustomObject]$o)
    }
    return [ordered]@{
        SchemaVersion = 1
        Collector     = $Collector
        SheetName     = $SheetName
        Columns       = $columns.ToArray()
        RowCount      = $uniform.Count
        Rows          = $uniform.ToArray()
        CollectedUtc  = [datetime]::UtcNow.ToString('o')
        Message       = [string]$Message
    }
}

function Save-IQExtrasSheet {
    <#
    .SYNOPSIS
    Writes a sheet object through Save-IQInventory -Name 'extras-<SheetName>' and returns the file path (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SheetName,
        [Parameter(Mandatory = $true)][string]$Collector,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()]$Rows,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$PreferredColumns,
        [Parameter(Mandatory = $false)][switch]$KeepPreferred,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Message
    )
    $sheet = New-IQExtrasSheet -SheetName $SheetName -Collector $Collector -Rows $Rows -PreferredColumns $PreferredColumns -KeepPreferred:$KeepPreferred -Message $Message
    $path = Save-IQInventory -Name ('extras-' + $SheetName) -Object $sheet
    Write-IQLog -Level Debug -Stage 'Extras' -Item $Collector -Message ("Sheet {0}: {1} row(s), {2} column(s) -> {3}" -f $SheetName, $sheet.RowCount, @($sheet.Columns).Count, $path)
    return [string]$path
}

function Get-IQExtrasSheetPath {
    <#
    .SYNOPSIS
    The inventory file path that Save-IQExtrasSheet writes for a sheet name, without writing it (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$SheetName)
    if (-not $script:IQ -or [string]::IsNullOrEmpty($script:IQ.RunPath)) { throw 'No active run. Call Initialize-IQRun first.' }
    return (Join-Path (Join-Path $script:IQ.RunPath 'inventory') ((Get-IQSafeKey -Value ('extras-' + $SheetName)) + '.json'))
}

function Get-IQExtrasUserRow {
    <#
    .SYNOPSIS
    Flattens one API user record into the UserEmailAddress/UserDisplayName/UserIdentifier/UserGraphId/UserPrincipalType/UserType columns plus the access right (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$User,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Lead,
        [Parameter(Mandatory = $false)][string]$AccessRightColumn = 'UserAccessRight'
    )
    $row = [ordered]@{}
    foreach ($k in $Lead.Keys) { $row[[string]$k] = $Lead[$k] }
    $row['UserEmailAddress'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $User -Name 'emailAddress')
    $row['UserDisplayName'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $User -Name 'displayName')
    $row['UserIdentifier'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $User -Name 'identifier')
    $row['UserGraphId'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $User -Name 'graphId')
    $row['UserPrincipalType'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $User -Name 'principalType')
    $row['UserType'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $User -Name 'userType')
    $right = $null
    foreach ($n in @('groupUserAccessRight', 'reportUserAccessRight', 'dashboardUserAccessRight', 'datasetUserAccessRight', 'dataflowUserAccessRight', 'datamartUserAccessRight', 'accessRight')) {
        $v = Get-IQExtrasMember -Object $User -Name $n
        if ($null -ne $v) { $right = [string]$v; break }
    }
    $row[$AccessRightColumn] = $right
    $userProfile = Get-IQExtrasMember -Object $User -Name 'profile'
    if ($null -ne $userProfile) { $row['UserProfileJson'] = ConvertTo-IQExtrasScalar -Value $userProfile }
    return $row
}

# =====================================================================================================================
# Admin probe
# =====================================================================================================================

function Test-IQAdminAccess {
    <#
    .SYNOPSIS
    $true when the signed-in user can call the Power BI admin APIs (GET admin/capacities?$top=1 succeeds).
    .DESCRIPTION
    A $null response (HTTP 401/403/404) means "not a Fabric administrator" for this run (IsAdmin = $false,
    Transient = $false). Any other failure (5xx after the retries, network, token refresh) is transient: IsAdmin =
    $false and Transient = $true, so the caller records the admin collectors as Failed (retried on resume) instead of
    Skipped. Returns @{ IsAdmin; Transient; Message }.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Stage = 'Extras')
    $result = @{ IsAdmin = $false; Transient = $false; Message = '' }
    try {
        $response = Invoke-IQApi -Method GET -Path 'admin/capacities' -Query @{ '$top' = 1 } -AllowNotFound -NoPaging -Stage $Stage
        if ($null -eq $response) {
            $result.Message = 'admin/capacities returned no response (HTTP 401/403/404): the signed-in user is not a Fabric administrator, so -IncludeAdminApis collectors are skipped.'
            return $result
        }
        $result.IsAdmin = $true
        $result.Message = 'Admin API probe succeeded (Fabric administrator).'
        return $result
    }
    catch {
        $result.Transient = $true
        $result.Message = 'Admin API probe failed (' + $_.Exception.Message + '): the -IncludeAdminApis collectors are recorded as Failed and retried on the next start.'
        return $result
    }
}

# =====================================================================================================================
# admin/groups (expand users, reports, datasets, dataflows, dashboards)
# =====================================================================================================================

function Get-IQAdminGroupInventory {
    <#
    .SYNOPSIS
    Collector admin-groups: pages GET admin/groups with $top/$skip and $expand, saves each page under extracts\admin and writes the AdminWorkspaces / AdminWorkspaceUsers sheets.
    .DESCRIPTION
    Returns @{ Success; Message; Outputs; WorkspaceCount; UserCount; Pages }. Never throws for API failures. A page
    after the first that returns nothing (HTTP 400/403/404) fails the collector: a truncated list must not be
    checkpointed as the complete tenant.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Stage = 'Extras')
    $collector = 'admin-groups'
    $result = @{ Success = $false; Message = ''; Outputs = @(); WorkspaceCount = 0; UserCount = 0; Pages = 0 }
    $folder = Get-IQExtrasFolder -Name 'admin'
    $pageSize = [int](Get-IQExtrasOption -Name 'AdminGroupsPageSize' -Default 5000)
    if ($pageSize -lt 1) { $pageSize = 5000 }
    if ($pageSize -gt 5000) { $pageSize = 5000 }
    $expand = 'users,reports,datasets,dataflows,dashboards'

    $groups = New-Object System.Collections.Generic.List[object]
    $skip = 0
    $page = 0
    $rawFiles = @()
    try {
        while ($true) {
            $page++
            # -Raw: the page body (up to 5 000 expanded workspaces) is written verbatim and parsed once.
            $raw = Invoke-IQApi -Method GET -Path 'admin/groups' -Query @{ '$top' = $pageSize; '$skip' = $skip; '$expand' = $expand } -NoPaging -Raw -Stage $Stage
            if ($null -eq $raw -or [string]::IsNullOrWhiteSpace([string]$raw)) {
                if ($page -eq 1) {
                    $result.Message = 'GET admin/groups returned no response (HTTP 400/403/404 - see the Warn line above)'
                    return $result
                }
                $result.Message = ('GET admin/groups page {0} ($skip={1}) returned no response (HTTP 400/403/404 - see the Warn line above); {2} workspace(s) read before it - the collector is retried on the next start' -f $page, $skip, $groups.Count)
                return $result
            }
            $response = ConvertFrom-Json -InputObject ([string]$raw)
            $rows = @(Get-IQExtrasMember -Object $response -Name 'value')
            $rows = @($rows | Where-Object { $null -ne $_ })
            $rawPath = Join-Path $folder ('groups-' + $page.ToString('000') + '.json')
            Write-IQExtrasRawFile -Text ([string]$raw) -Path $rawPath
            $rawFiles += $rawPath
            foreach ($g in $rows) { $groups.Add($g) }
            Write-IQLog -Level Debug -Stage $Stage -Item $collector -Message ("admin/groups page {0}: {1} workspace(s) (skip {2})" -f $page, $rows.Count, $skip)
            if ($rows.Count -lt $pageSize) { break }
            $skip += $pageSize
            if ($page -ge 1000) { Write-IQLog -Level Warn -Stage $Stage -Item $collector -Message 'admin/groups paging guard hit (1000 pages)'; break }
        }
    }
    catch {
        $result.Message = 'GET admin/groups failed: ' + $_.Exception.Message
        return $result
    }
    $result.Pages = $page

    $wsMap = [ordered]@{
        id                          = 'WorkspaceId'
        name                        = 'WorkspaceName'
        type                        = 'WorkspaceType'
        state                       = 'WorkspaceState'
        isReadOnly                  = 'WorkspaceIsReadOnly'
        isOnDedicatedCapacity       = 'WorkspaceIsOnDedicatedCapacity'
        capacityId                  = 'WorkspaceCapacityId'
        capacityMigrationStatus     = 'WorkspaceCapacityMigrationStatus'
        description                 = 'WorkspaceDescription'
        defaultDatasetStorageFormat = 'WorkspaceDefaultDatasetStorageFormat'
        hasWorkspaceLevelSettings   = 'WorkspaceHasWorkspaceLevelSettings'
        pipelineId                  = 'WorkspacePipelineId'
    }
    $wsRows = New-Object System.Collections.Generic.List[object]
    $userRows = New-Object System.Collections.Generic.List[object]
    foreach ($g in $groups) {
        $users = @(Get-IQExtrasMember -Object $g -Name 'users')
        $reports = @(Get-IQExtrasMember -Object $g -Name 'reports')
        $datasets = @(Get-IQExtrasMember -Object $g -Name 'datasets')
        $dataflows = @(Get-IQExtrasMember -Object $g -Name 'dataflows')
        $dashboards = @(Get-IQExtrasMember -Object $g -Name 'dashboards')
        $row = ConvertTo-IQExtrasRow -Object $g -Map $wsMap -Exclude @('users', 'reports', 'datasets', 'dataflows', 'dashboards', 'workbooks', 'datamarts')
        $row['WorkspaceUserCount'] = @($users | Where-Object { $null -ne $_ }).Count
        $row['WorkspaceReportCount'] = @($reports | Where-Object { $null -ne $_ }).Count
        $row['WorkspaceDatasetCount'] = @($datasets | Where-Object { $null -ne $_ }).Count
        $row['WorkspaceDataflowCount'] = @($dataflows | Where-Object { $null -ne $_ }).Count
        $row['WorkspaceDashboardCount'] = @($dashboards | Where-Object { $null -ne $_ }).Count
        $row['WorkspaceReportIds'] = ConvertTo-IQExtrasJoined -Items $reports -Property 'id'
        $row['WorkspaceDatasetIds'] = ConvertTo-IQExtrasJoined -Items $datasets -Property 'id'
        $row['WorkspaceDataflowIds'] = ConvertTo-IQExtrasJoined -Items $dataflows -Property 'objectId'
        $row['WorkspaceDashboardIds'] = ConvertTo-IQExtrasJoined -Items $dashboards -Property 'id'
        $wsRows.Add($row)
        $lead = [ordered]@{ WorkspaceId = [string](Get-IQExtrasMember -Object $g -Name 'id'); WorkspaceName = [string](Get-IQExtrasMember -Object $g -Name 'name') }
        foreach ($u in $users) {
            if ($null -eq $u) { continue }
            $userRows.Add((Get-IQExtrasUserRow -User $u -Lead $lead -AccessRightColumn 'UserGroupUserAccessRight'))
        }
    }
    $result.WorkspaceCount = $wsRows.Count
    $result.UserCount = $userRows.Count
    $wsColumns = @('WorkspaceId', 'WorkspaceName', 'WorkspaceType', 'WorkspaceState', 'WorkspaceIsReadOnly', 'WorkspaceIsOnDedicatedCapacity', 'WorkspaceCapacityId', 'WorkspaceCapacityMigrationStatus', 'WorkspaceDescription', 'WorkspaceDefaultDatasetStorageFormat', 'WorkspaceHasWorkspaceLevelSettings', 'WorkspacePipelineId', 'WorkspaceUserCount', 'WorkspaceReportCount', 'WorkspaceDatasetCount', 'WorkspaceDataflowCount', 'WorkspaceDashboardCount', 'WorkspaceReportIds', 'WorkspaceDatasetIds', 'WorkspaceDataflowIds', 'WorkspaceDashboardIds')
    $userColumns = @('WorkspaceId', 'WorkspaceName', 'UserEmailAddress', 'UserDisplayName', 'UserIdentifier', 'UserGraphId', 'UserPrincipalType', 'UserGroupUserAccessRight', 'UserType', 'UserProfileJson')
    $outputs = @()
    $outputs += Save-IQExtrasSheet -SheetName 'AdminWorkspaces' -Collector $collector -Rows $wsRows.ToArray() -PreferredColumns $wsColumns -KeepPreferred
    $outputs += Save-IQExtrasSheet -SheetName 'AdminWorkspaceUsers' -Collector $collector -Rows $userRows.ToArray() -PreferredColumns $userColumns -KeepPreferred
    $result.Outputs = @($outputs) + @($rawFiles)
    $result.Success = $true
    $result.Message = ('{0} workspace(s), {1} workspace user row(s) from {2} page(s)' -f $wsRows.Count, $userRows.Count, $page)
    return $result
}

# =====================================================================================================================
# Scanner API (admin/workspaces/modified -> getInfo -> scanStatus -> scanResult)
# =====================================================================================================================

function Get-IQAdminScanWorkspaceId {
    <#
    .SYNOPSIS
    The workspace ids to scan: GET admin/workspaces/modified (personal and inactive workspaces excluded), optionally restricted to the run scope / a maximum count (private).
    .DESCRIPTION
    The ids are returned sorted (ordinal, case-insensitive) so a retry slices the same batches and finds the cached
    batch files; the API's own order is unspecified. An empty list ("[]" - a tenant or scope without active
    non-personal workspaces) is returned as @(), which is not the same as a 403/404 ($null -> throw).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Stage = 'Extras')
    $collector = 'admin-scan'
    $raw = Invoke-IQApi -Method GET -Path 'admin/workspaces/modified' -Query @{ excludePersonalWorkspaces = 'True'; excludeInActiveWorkspaces = 'True' } -NoPaging -Raw -Stage $Stage
    if ($null -eq $raw) { throw 'GET admin/workspaces/modified returned no response (HTTP 400/403/404 - see the Warn line above)' }
    $text = ([string]$raw).Trim()
    if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1).Trim() }
    if ($text -eq '' -or $text -eq '[]') {
        Write-IQLog -Level Info -Stage $Stage -Item $collector -Message 'admin/workspaces/modified returned no workspaces (nothing to scan)'
        return @()
    }
    $response = ConvertFrom-Json -InputObject $text
    $entries = @()
    if ($response -is [System.Management.Automation.PSCustomObject] -and $null -ne $response.PSObject.Properties['value']) { $entries = @($response.value) }
    else { $entries = @($response) }
    $ids = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($e in $entries) {
        if ($null -eq $e) { continue }
        $id = $null
        if ($e -is [string]) { $id = $e } else { $id = [string](Get-IQExtrasMember -Object $e -Name 'id') }
        if (-not (Test-IQExtrasGuid -Value $id)) { continue }
        $key = $id.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $ids.Add($id)
    }
    $ids.Sort([System.StringComparer]::OrdinalIgnoreCase)
    $total = $ids.Count
    if ([bool](Get-IQExtrasOption -Name 'AdminScanScopeOnly' -Default $false)) {
        $inScope = @{}
        foreach ($w in @(Get-IQSelectedWorkspaces)) {
            $wid = [string](Get-IQExtrasMember -Object $w -Name 'WorkspaceId')
            if (Test-IQExtrasGuid -Value $wid) { $inScope[$wid.ToLowerInvariant()] = $true }
        }
        $filtered = New-Object System.Collections.Generic.List[string]
        foreach ($id in $ids) { if ($inScope.ContainsKey($id.ToLowerInvariant())) { $filtered.Add($id) } }
        Write-IQLog -Level Info -Stage $Stage -Item $collector -Message ("AdminScanScopeOnly: scanning {0} of {1} modified workspace(s) that are in the run scope" -f $filtered.Count, $total)
        $ids = $filtered
    }
    $max = [int](Get-IQExtrasOption -Name 'AdminScanMaxWorkspaces' -Default 0)
    if ($max -gt 0 -and $ids.Count -gt $max) {
        Write-IQLog -Level Warn -Stage $Stage -Item $collector -Message ("AdminScanMaxWorkspaces={0}: scanning only the first {0} of {1} workspace(s)" -f $max, $ids.Count)
        $ids = New-Object System.Collections.Generic.List[string](, [string[]]@($ids | Select-Object -First $max))
    }
    return $ids.ToArray()
}

function Get-IQAdminScanBatchHash {
    <#
    .SYNOPSIS
    First 8 hex chars of the SHA-1 of the sorted, lower-cased ids of a batch (part of the cached batch file name) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()][string[]]$Ids)
    $text = (@($Ids | ForEach-Object { ([string]$_).ToLowerInvariant() } | Sort-Object) -join ',')
    $sha = [System.Security.Cryptography.SHA1]::Create()
    try {
        $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($text))
    }
    finally { $sha.Dispose() }
    return (([System.BitConverter]::ToString($bytes) -replace '-', '').Substring(0, 8).ToLowerInvariant())
}

function Get-IQAdminScanStatus {
    <#
    .SYNOPSIS
    One GET admin/workspaces/scanStatus/{id} through the Http request core; returns @{ Status; Error; RetryAfter } (RetryAfter = the header in seconds, or $null) (private).
    .DESCRIPTION
    Invoke-IQApi hides the response headers, and the Scanner API tells the caller how long to wait through
    Retry-After on a 200 "Running"/"NotStarted" answer; polling faster than that only burns the getInfo/scanStatus
    quota. Throws when the request fails (after the Http module's own retries) or returns no body.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ScanId,
        [Parameter(Mandatory = $false)][string]$Stage = 'Extras'
    )
    $url = Get-IQApiUrl -Path ('admin/workspaces/scanStatus/' + $ScanId) -Api 'PowerBI'
    $response = Invoke-IQHttpRequest -Method GET -Url $url -Api 'PowerBI' -Stage $Stage
    if ($null -eq $response) { throw ('GET admin/workspaces/scanStatus/' + $ScanId + ' returned no response') }
    $content = $null
    if ($response -is [System.Collections.IDictionary]) { $content = $response['Content'] } else { $content = Get-IQExtrasMember -Object $response -Name 'Content' }
    if ($null -eq $content -or [string]::IsNullOrWhiteSpace([string]$content)) { throw ('GET admin/workspaces/scanStatus/' + $ScanId + ' returned an empty body') }
    $parsed = ConvertFrom-Json -InputObject ([string]$content)
    $headers = $null
    if ($response -is [System.Collections.IDictionary]) { $headers = $response['Headers'] } else { $headers = Get-IQExtrasMember -Object $response -Name 'Headers' }
    $retryAfter = $null
    if ($null -ne $headers) {
        try { $retryAfter = ConvertTo-IQRetryAfterDelay -Value (Get-IQHttpHeaderValue -Headers $headers -Name 'Retry-After') } catch { $retryAfter = $null }
    }
    return @{
        Status     = [string](Get-IQExtrasMember -Object $parsed -Name 'status')
        Error      = (Get-IQExtrasMember -Object $parsed -Name 'error')
        RetryAfter = $retryAfter
    }
}

function Invoke-IQAdminScanBatch {
    <#
    .SYNOPSIS
    Runs one Scanner API batch (getInfo -> scanStatus polling -> scanResult), writes the scanResult body verbatim to -OutPath and returns it parsed, or throws with the reason (private).
    .DESCRIPTION
    The poll waits max(Retry-After, ExtrasPollSeconds) between scanStatus calls, checks Test-IQTimeBudget on every
    turn (throws [System.OperationCanceledException] when the budget is used up) and gives up after
    AdminScanTimeoutMinutes with a [System.TimeoutException], so the caller can tell the three outcomes apart.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Ids,
        [Parameter(Mandatory = $false)][string]$Stage = 'Extras',
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Item,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$OutPath
    )
    $query = @{ lineage = 'True'; datasourceDetails = 'True'; datasetSchema = 'True'; datasetExpressions = 'True'; getArtifactUsers = 'True' }
    $body = @{ workspaces = @($Ids) }
    $accepted = Invoke-IQApi -Method POST -Path 'admin/workspaces/getInfo' -Query $query -Body $body -NoPaging -Stage $Stage
    if ($null -eq $accepted) { throw 'POST admin/workspaces/getInfo returned no response (HTTP 400/403/404 - see the Warn line above)' }
    $scanId = [string](Get-IQExtrasMember -Object $accepted -Name 'id')
    if ([string]::IsNullOrWhiteSpace($scanId)) { throw 'POST admin/workspaces/getInfo returned no scan id' }
    $status = [string](Get-IQExtrasMember -Object $accepted -Name 'status')
    $scanError = Get-IQExtrasMember -Object $accepted -Name 'error'
    $pollSeconds = [int](Get-IQExtrasOption -Name 'ExtrasPollSeconds' -Default 10)
    if ($pollSeconds -lt 0) { $pollSeconds = 0 }
    $timeoutMinutes = [int](Get-IQExtrasOption -Name 'AdminScanTimeoutMinutes' -Default 30)
    if ($timeoutMinutes -lt 1) { $timeoutMinutes = 1 }
    $deadline = [datetime]::UtcNow.AddMinutes($timeoutMinutes)
    $polls = 0
    $wait = $pollSeconds
    while ($status -ine 'Succeeded') {
        if ($status -ieq 'Failed') { throw ('scan ' + $scanId + ' ended with status Failed: ' + (ConvertTo-IQExtrasScalar -Value $scanError)) }
        if ([datetime]::UtcNow -ge $deadline) { throw (New-Object System.TimeoutException(('scan ' + $scanId + ' did not finish within ' + $timeoutMinutes + ' minute(s) (last status ' + $status + ')'))) }
        if (Test-IQTimeBudget -Stage $Stage -Item $Item) { throw (New-Object System.OperationCanceledException(('time budget reached while polling scan ' + $scanId + ' (last status ' + $status + '); the batch is scanned again on the next start'))) }
        Start-IQExtrasSleep -Seconds $wait
        $polls++
        $poll = Get-IQAdminScanStatus -ScanId $scanId -Stage $Stage
        $status = [string]$poll.Status
        $scanError = $poll.Error
        $wait = $pollSeconds
        if ($null -ne $poll.RetryAfter -and [int]$poll.RetryAfter -gt $wait) { $wait = [math]::Min(300, [int]$poll.RetryAfter) }
        if ($polls -gt 100000) { throw 'scanStatus polling guard hit' }
    }
    $rawResult = Invoke-IQApi -Method GET -Path ('admin/workspaces/scanResult/' + $scanId) -NoPaging -Raw -TimeoutSec 600 -Stage $Stage
    if ($null -eq $rawResult -or [string]::IsNullOrWhiteSpace([string]$rawResult)) { throw ('GET admin/workspaces/scanResult/' + $scanId + ' returned no response') }
    $scanResult = ConvertFrom-Json -InputObject ([string]$rawResult)
    if ($null -eq $scanResult) { throw ('GET admin/workspaces/scanResult/' + $scanId + ' returned an empty body') }
    if (-not [string]::IsNullOrWhiteSpace($OutPath)) { Write-IQExtrasRawFile -Text ([string]$rawResult) -Path $OutPath }
    Write-IQLog -Level Debug -Stage $Stage -Item $Item -Message ("scan {0}: {1} workspace id(s), {2} poll(s)" -f $scanId, $Ids.Count, $polls)
    return $scanResult
}

function ConvertTo-IQScanRowSet {
    <#
    .SYNOPSIS
    Flattens scanResult objects into the Scan* sheet row lists (private).
    .DESCRIPTION
    Returns a hashtable of List[object] keyed by sheet name: ScanWorkspaces, ScanDatasets, ScanTables, ScanColumns,
    ScanMeasures, ScanExpressions, ScanDatasources, ScanReports, ScanDashboards, ScanDataflows, ScanUsers.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()]$ScanResults)
    $sets = @{}
    foreach ($name in @('ScanWorkspaces', 'ScanDatasets', 'ScanTables', 'ScanColumns', 'ScanMeasures', 'ScanExpressions', 'ScanDatasources', 'ScanReports', 'ScanDashboards', 'ScanDataflows', 'ScanUsers')) {
        $sets[$name] = New-Object System.Collections.Generic.List[object]
    }
    $datasourceUse = @{}      # datasourceInstanceId -> @{ Datasets; Dataflows }
    $datasources = [ordered]@{}
    $misconfigured = @{}

    foreach ($scan in @($ScanResults)) {
        if ($null -eq $scan) { continue }
        foreach ($ds in @(Get-IQExtrasMember -Object $scan -Name 'datasourceInstances')) {
            if ($null -eq $ds) { continue }
            $id = [string](Get-IQExtrasMember -Object $ds -Name 'datasourceId')
            if ([string]::IsNullOrEmpty($id)) { continue }
            if (-not $datasources.Contains($id)) { $datasources[$id] = $ds }
        }
        foreach ($ds in @(Get-IQExtrasMember -Object $scan -Name 'misconfiguredDatasourceInstances')) {
            if ($null -eq $ds) { continue }
            $id = [string](Get-IQExtrasMember -Object $ds -Name 'datasourceId')
            if ([string]::IsNullOrEmpty($id)) { continue }
            $misconfigured[$id] = $true
            if (-not $datasources.Contains($id)) { $datasources[$id] = $ds }
        }

        foreach ($ws in @(Get-IQExtrasMember -Object $scan -Name 'workspaces')) {
            if ($null -eq $ws) { continue }
            $wsId = [string](Get-IQExtrasMember -Object $ws -Name 'id')
            $wsName = [string](Get-IQExtrasMember -Object $ws -Name 'name')
            $lead = [ordered]@{ WorkspaceId = $wsId; WorkspaceName = $wsName }
            $reports = @(Get-IQExtrasMember -Object $ws -Name 'reports')
            $dashboards = @(Get-IQExtrasMember -Object $ws -Name 'dashboards')
            $datasets = @(Get-IQExtrasMember -Object $ws -Name 'datasets')
            $dataflows = @(Get-IQExtrasMember -Object $ws -Name 'dataflows')
            $datamarts = @(Get-IQExtrasMember -Object $ws -Name 'datamarts')
            $wsUsers = @(Get-IQExtrasMember -Object $ws -Name 'users')

            $wsRow = ConvertTo-IQExtrasRow -Object $ws -Map ([ordered]@{
                    id                          = 'WorkspaceId'
                    name                        = 'WorkspaceName'
                    type                        = 'WorkspaceType'
                    state                       = 'WorkspaceState'
                    isOnDedicatedCapacity       = 'WorkspaceIsOnDedicatedCapacity'
                    capacityId                  = 'WorkspaceCapacityId'
                    defaultDatasetStorageFormat = 'WorkspaceDefaultDatasetStorageFormat'
                    description                 = 'WorkspaceDescription'
                }) -Exclude @('reports', 'dashboards', 'datasets', 'dataflows', 'datamarts', 'users', 'workbooks')
            $wsRow['WorkspaceReportCount'] = @($reports | Where-Object { $null -ne $_ }).Count
            $wsRow['WorkspaceDashboardCount'] = @($dashboards | Where-Object { $null -ne $_ }).Count
            $wsRow['WorkspaceDatasetCount'] = @($datasets | Where-Object { $null -ne $_ }).Count
            $wsRow['WorkspaceDataflowCount'] = @($dataflows | Where-Object { $null -ne $_ }).Count
            $wsRow['WorkspaceDatamartCount'] = @($datamarts | Where-Object { $null -ne $_ }).Count
            $wsRow['WorkspaceUserCount'] = @($wsUsers | Where-Object { $null -ne $_ }).Count
            $sets['ScanWorkspaces'].Add($wsRow)
            foreach ($u in $wsUsers) {
                if ($null -eq $u) { continue }
                $ulead = [ordered]@{ WorkspaceId = $wsId; WorkspaceName = $wsName; ArtifactType = 'Workspace'; ArtifactId = $wsId; ArtifactName = $wsName }
                $sets['ScanUsers'].Add((Get-IQExtrasUserRow -User $u -Lead $ulead))
            }

            foreach ($d in $datasets) {
                if ($null -eq $d) { continue }
                $dsId = [string](Get-IQExtrasMember -Object $d -Name 'id')
                $dsName = [string](Get-IQExtrasMember -Object $d -Name 'name')
                $tables = @(Get-IQExtrasMember -Object $d -Name 'tables')
                $dsUsers = @(Get-IQExtrasMember -Object $d -Name 'users')
                $expressions = @(Get-IQExtrasMember -Object $d -Name 'expressions')
                $roles = @(Get-IQExtrasMember -Object $d -Name 'roles')
                $relationships = @(Get-IQExtrasMember -Object $d -Name 'relationships')
                $usages = @(Get-IQExtrasMember -Object $d -Name 'datasourceUsages')
                $upstream = @(Get-IQExtrasMember -Object $d -Name 'upstreamDataflows')
                $upstreamMarts = @(Get-IQExtrasMember -Object $d -Name 'upstreamDatamarts')
                $endorsement = Get-IQExtrasMember -Object $d -Name 'endorsementDetails'
                $label = Get-IQExtrasMember -Object $d -Name 'sensitivityLabel'
                $row = ConvertTo-IQExtrasRow -Object $d -Lead $lead -Map ([ordered]@{
                        id                               = 'DatasetId'
                        name                             = 'DatasetName'
                        description                      = 'DatasetDescription'
                        configuredBy                     = 'DatasetConfiguredBy'
                        configuredById                   = 'DatasetConfiguredById'
                        createdDate                      = 'DatasetCreatedDate'
                        contentProviderType              = 'DatasetContentProviderType'
                        targetStorageMode                = 'DatasetTargetStorageMode'
                        isRefreshable                    = 'DatasetIsRefreshable'
                        isEffectiveIdentityRequired      = 'DatasetIsEffectiveIdentityRequired'
                        isEffectiveIdentityRolesRequired = 'DatasetIsEffectiveIdentityRolesRequired'
                        isOnPremGatewayRequired          = 'DatasetIsOnPremGatewayRequired'
                    }) -Exclude @('tables', 'users', 'expressions', 'roles', 'relationships', 'datasourceUsages', 'misconfiguredDatasourceUsages', 'upstreamDataflows', 'upstreamDatamarts', 'endorsementDetails', 'sensitivityLabel', 'schemaMayNotBeUpToDate', 'schemaRetrievalError', 'refreshSchedule', 'directQueryRefreshSchedule')
                $row['DatasetEndorsement'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $endorsement -Name 'endorsement')
                $row['DatasetCertifiedBy'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $endorsement -Name 'certifiedBy')
                $row['DatasetSensitivityLabelId'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $label -Name 'labelId')
                $row['DatasetDatasourceInstanceIds'] = ConvertTo-IQExtrasJoined -Items $usages -Property 'datasourceInstanceId'
                $row['DatasetUpstreamDataflowIds'] = ConvertTo-IQExtrasJoined -Items $upstream -Property 'targetDataflowId'
                $row['DatasetUpstreamDatamartIds'] = ConvertTo-IQExtrasJoined -Items $upstreamMarts -Property 'targetDatamartId'
                $columnCount = 0
                $measureCount = 0
                foreach ($t in $tables) {
                    if ($null -eq $t) { continue }
                    $columnCount += @(Get-IQExtrasMember -Object $t -Name 'columns' | Where-Object { $null -ne $_ }).Count
                    $measureCount += @(Get-IQExtrasMember -Object $t -Name 'measures' | Where-Object { $null -ne $_ }).Count
                }
                $row['DatasetTableCount'] = @($tables | Where-Object { $null -ne $_ }).Count
                $row['DatasetColumnCount'] = $columnCount
                $row['DatasetMeasureCount'] = $measureCount
                $row['DatasetRelationshipCount'] = @($relationships | Where-Object { $null -ne $_ }).Count
                $row['DatasetExpressionCount'] = @($expressions | Where-Object { $null -ne $_ }).Count
                $row['DatasetRoleNames'] = ConvertTo-IQExtrasJoined -Items $roles -Property 'name'
                $row['DatasetUserCount'] = @($dsUsers | Where-Object { $null -ne $_ }).Count
                $row['DatasetSchemaMayNotBeUpToDate'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $d -Name 'schemaMayNotBeUpToDate')
                $row['DatasetSchemaRetrievalError'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $d -Name 'schemaRetrievalError')
                $row['DatasetRefreshScheduleJson'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $d -Name 'refreshSchedule')
                $row['DatasetDirectQueryRefreshScheduleJson'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $d -Name 'directQueryRefreshSchedule')
                $sets['ScanDatasets'].Add($row)
                foreach ($u in $usages) {
                    $uid = [string](Get-IQExtrasMember -Object $u -Name 'datasourceInstanceId')
                    if ([string]::IsNullOrEmpty($uid)) { continue }
                    if (-not $datasourceUse.ContainsKey($uid)) { $datasourceUse[$uid] = @{ Datasets = 0; Dataflows = 0 } }
                    $datasourceUse[$uid].Datasets = [int]$datasourceUse[$uid].Datasets + 1
                }
                $dlead = [ordered]@{ WorkspaceId = $wsId; WorkspaceName = $wsName; DatasetId = $dsId; DatasetName = $dsName }
                foreach ($t in $tables) {
                    if ($null -eq $t) { continue }
                    $tName = [string](Get-IQExtrasMember -Object $t -Name 'name')
                    $columns = @(Get-IQExtrasMember -Object $t -Name 'columns')
                    $measures = @(Get-IQExtrasMember -Object $t -Name 'measures')
                    $sources = @(Get-IQExtrasMember -Object $t -Name 'source')
                    $tRow = ConvertTo-IQExtrasRow -Object $t -Lead $dlead -Map ([ordered]@{
                            name        = 'TableName'
                            description = 'TableDescription'
                            isHidden    = 'TableIsHidden'
                            storageMode = 'TableStorageMode'
                        }) -Exclude @('columns', 'measures', 'source')
                    $tRow['TableSourceExpression'] = ((@($sources | ForEach-Object { if ($null -ne $_) { [string](Get-IQExtrasMember -Object $_ -Name 'expression') } }) | Where-Object { -not [string]::IsNullOrEmpty($_) }) -join "`n")
                    $tRow['TableColumnCount'] = @($columns | Where-Object { $null -ne $_ }).Count
                    $tRow['TableMeasureCount'] = @($measures | Where-Object { $null -ne $_ }).Count
                    $sets['ScanTables'].Add($tRow)
                    $tlead = [ordered]@{ WorkspaceId = $wsId; WorkspaceName = $wsName; DatasetId = $dsId; DatasetName = $dsName; TableName = $tName }
                    foreach ($c in $columns) {
                        if ($null -eq $c) { continue }
                        $sets['ScanColumns'].Add((ConvertTo-IQExtrasRow -Object $c -Lead $tlead -Map ([ordered]@{
                                        name        = 'ColumnName'
                                        dataType    = 'ColumnDataType'
                                        columnType  = 'ColumnType'
                                        isHidden    = 'ColumnIsHidden'
                                        expression  = 'ColumnExpression'
                                        description = 'ColumnDescription'
                                    })))
                    }
                    foreach ($m in $measures) {
                        if ($null -eq $m) { continue }
                        $sets['ScanMeasures'].Add((ConvertTo-IQExtrasRow -Object $m -Lead $tlead -Map ([ordered]@{
                                        name        = 'MeasureName'
                                        expression  = 'MeasureExpression'
                                        isHidden    = 'MeasureIsHidden'
                                        description = 'MeasureDescription'
                                    })))
                    }
                }
                foreach ($e in $expressions) {
                    if ($null -eq $e) { continue }
                    $sets['ScanExpressions'].Add((ConvertTo-IQExtrasRow -Object $e -Lead $dlead -Map ([ordered]@{
                                    name        = 'ExpressionName'
                                    expression  = 'Expression'
                                    description = 'ExpressionDescription'
                                })))
                }
                foreach ($u in $dsUsers) {
                    if ($null -eq $u) { continue }
                    $ulead = [ordered]@{ WorkspaceId = $wsId; WorkspaceName = $wsName; ArtifactType = 'Dataset'; ArtifactId = $dsId; ArtifactName = $dsName }
                    $sets['ScanUsers'].Add((Get-IQExtrasUserRow -User $u -Lead $ulead))
                }
            }

            foreach ($r in $reports) {
                if ($null -eq $r) { continue }
                $rId = [string](Get-IQExtrasMember -Object $r -Name 'id')
                $rName = [string](Get-IQExtrasMember -Object $r -Name 'name')
                $rUsers = @(Get-IQExtrasMember -Object $r -Name 'users')
                $endorsement = Get-IQExtrasMember -Object $r -Name 'endorsementDetails'
                $label = Get-IQExtrasMember -Object $r -Name 'sensitivityLabel'
                $row = ConvertTo-IQExtrasRow -Object $r -Lead $lead -Map ([ordered]@{
                        id                     = 'ReportId'
                        name                   = 'ReportName'
                        reportType             = 'ReportType'
                        datasetId              = 'ReportDatasetId'
                        datasetWorkspaceId     = 'ReportDatasetWorkspaceId'
                        createdDateTime        = 'ReportCreatedDateTime'
                        modifiedDateTime       = 'ReportModifiedDateTime'
                        createdBy              = 'ReportCreatedBy'
                        createdById            = 'ReportCreatedById'
                        modifiedBy             = 'ReportModifiedBy'
                        modifiedById           = 'ReportModifiedById'
                        appId                  = 'ReportAppId'
                        originalReportObjectId = 'ReportOriginalReportObjectId'
                        description            = 'ReportDescription'
                    }) -Exclude @('users', 'endorsementDetails', 'sensitivityLabel', 'subscriptions')
                $row['ReportEndorsement'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $endorsement -Name 'endorsement')
                $row['ReportCertifiedBy'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $endorsement -Name 'certifiedBy')
                $row['ReportSensitivityLabelId'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $label -Name 'labelId')
                $row['ReportUserCount'] = @($rUsers | Where-Object { $null -ne $_ }).Count
                $sets['ScanReports'].Add($row)
                foreach ($u in $rUsers) {
                    if ($null -eq $u) { continue }
                    $ulead = [ordered]@{ WorkspaceId = $wsId; WorkspaceName = $wsName; ArtifactType = 'Report'; ArtifactId = $rId; ArtifactName = $rName }
                    $sets['ScanUsers'].Add((Get-IQExtrasUserRow -User $u -Lead $ulead))
                }
            }

            foreach ($b in $dashboards) {
                if ($null -eq $b) { continue }
                $bId = [string](Get-IQExtrasMember -Object $b -Name 'id')
                $bName = [string](Get-IQExtrasMember -Object $b -Name 'displayName')
                $bUsers = @(Get-IQExtrasMember -Object $b -Name 'users')
                $tiles = @(Get-IQExtrasMember -Object $b -Name 'tiles')
                $label = Get-IQExtrasMember -Object $b -Name 'sensitivityLabel'
                $row = ConvertTo-IQExtrasRow -Object $b -Lead $lead -Map ([ordered]@{
                        id          = 'DashboardId'
                        displayName = 'DashboardName'
                        isReadOnly  = 'DashboardIsReadOnly'
                        appId       = 'DashboardAppId'
                    }) -Exclude @('users', 'tiles', 'sensitivityLabel', 'subscriptions')
                $row['DashboardTileCount'] = @($tiles | Where-Object { $null -ne $_ }).Count
                $row['DashboardTileReportIds'] = ConvertTo-IQExtrasJoined -Items $tiles -Property 'reportId'
                $row['DashboardTileDatasetIds'] = ConvertTo-IQExtrasJoined -Items $tiles -Property 'datasetId'
                $row['DashboardSensitivityLabelId'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $label -Name 'labelId')
                $row['DashboardUserCount'] = @($bUsers | Where-Object { $null -ne $_ }).Count
                $sets['ScanDashboards'].Add($row)
                foreach ($u in $bUsers) {
                    if ($null -eq $u) { continue }
                    $ulead = [ordered]@{ WorkspaceId = $wsId; WorkspaceName = $wsName; ArtifactType = 'Dashboard'; ArtifactId = $bId; ArtifactName = $bName }
                    $sets['ScanUsers'].Add((Get-IQExtrasUserRow -User $u -Lead $ulead))
                }
            }

            foreach ($f in $dataflows) {
                if ($null -eq $f) { continue }
                $fId = [string](Get-IQExtrasMember -Object $f -Name 'objectId')
                $fName = [string](Get-IQExtrasMember -Object $f -Name 'name')
                $fUsers = @(Get-IQExtrasMember -Object $f -Name 'users')
                $usages = @(Get-IQExtrasMember -Object $f -Name 'datasourceUsages')
                $upstream = @(Get-IQExtrasMember -Object $f -Name 'upstreamDataflows')
                $endorsement = Get-IQExtrasMember -Object $f -Name 'endorsementDetails'
                $label = Get-IQExtrasMember -Object $f -Name 'sensitivityLabel'
                $row = ConvertTo-IQExtrasRow -Object $f -Lead $lead -Map ([ordered]@{
                        objectId         = 'DataflowId'
                        name             = 'DataflowName'
                        description      = 'DataflowDescription'
                        configuredBy     = 'DataflowConfiguredBy'
                        modifiedBy       = 'DataflowModifiedBy'
                        modifiedDateTime = 'DataflowModifiedDateTime'
                        generation       = 'DataflowGeneration'
                    }) -Exclude @('users', 'datasourceUsages', 'misconfiguredDatasourceUsages', 'upstreamDataflows', 'endorsementDetails', 'sensitivityLabel', 'refreshSchedule')
                $row['DataflowEndorsement'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $endorsement -Name 'endorsement')
                $row['DataflowCertifiedBy'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $endorsement -Name 'certifiedBy')
                $row['DataflowSensitivityLabelId'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $label -Name 'labelId')
                $row['DataflowDatasourceInstanceIds'] = ConvertTo-IQExtrasJoined -Items $usages -Property 'datasourceInstanceId'
                $row['DataflowUpstreamDataflowIds'] = ConvertTo-IQExtrasJoined -Items $upstream -Property 'targetDataflowId'
                $row['DataflowRefreshScheduleJson'] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $f -Name 'refreshSchedule')
                $row['DataflowUserCount'] = @($fUsers | Where-Object { $null -ne $_ }).Count
                $sets['ScanDataflows'].Add($row)
                foreach ($u in $usages) {
                    $uid = [string](Get-IQExtrasMember -Object $u -Name 'datasourceInstanceId')
                    if ([string]::IsNullOrEmpty($uid)) { continue }
                    if (-not $datasourceUse.ContainsKey($uid)) { $datasourceUse[$uid] = @{ Datasets = 0; Dataflows = 0 } }
                    $datasourceUse[$uid].Dataflows = [int]$datasourceUse[$uid].Dataflows + 1
                }
                foreach ($u in $fUsers) {
                    if ($null -eq $u) { continue }
                    $ulead = [ordered]@{ WorkspaceId = $wsId; WorkspaceName = $wsName; ArtifactType = 'Dataflow'; ArtifactId = $fId; ArtifactName = $fName }
                    $sets['ScanUsers'].Add((Get-IQExtrasUserRow -User $u -Lead $ulead))
                }
            }

            foreach ($m in $datamarts) {
                if ($null -eq $m) { continue }
                $mId = [string](Get-IQExtrasMember -Object $m -Name 'id')
                $mName = [string](Get-IQExtrasMember -Object $m -Name 'name')
                foreach ($u in @(Get-IQExtrasMember -Object $m -Name 'users')) {
                    if ($null -eq $u) { continue }
                    $ulead = [ordered]@{ WorkspaceId = $wsId; WorkspaceName = $wsName; ArtifactType = 'Datamart'; ArtifactId = $mId; ArtifactName = $mName }
                    $sets['ScanUsers'].Add((Get-IQExtrasUserRow -User $u -Lead $ulead))
                }
            }
        }
    }

    foreach ($id in $datasources.Keys) {
        $ds = $datasources[$id]
        $details = Get-IQExtrasMember -Object $ds -Name 'connectionDetails'
        $row = ConvertTo-IQExtrasRow -Object $ds -Map ([ordered]@{
                datasourceId   = 'DatasourceInstanceId'
                datasourceType = 'DatasourceType'
                gatewayId      = 'DatasourceGatewayId'
            }) -Exclude @('connectionDetails')
        foreach ($pair in @(@('server', 'DatasourceServer'), @('database', 'DatasourceDatabase'), @('url', 'DatasourceUrl'), @('path', 'DatasourcePath'), @('kind', 'DatasourceKind'), @('account', 'DatasourceAccount'), @('domain', 'DatasourceDomain'), @('emailAddress', 'DatasourceEmailAddress'), @('loginServer', 'DatasourceLoginServer'), @('classInfo', 'DatasourceClassInfo'))) {
            $row[$pair[1]] = ConvertTo-IQExtrasScalar -Value (Get-IQExtrasMember -Object $details -Name $pair[0])
        }
        $row['DatasourceConnectionDetailsJson'] = ConvertTo-IQExtrasScalar -Value $details
        $row['DatasourceIsMisconfigured'] = [bool]$misconfigured.ContainsKey($id)
        $use = $null
        if ($datasourceUse.ContainsKey($id)) { $use = $datasourceUse[$id] }
        $row['DatasourceUsedByDatasetCount'] = $(if ($null -ne $use) { [int]$use.Datasets } else { 0 })
        $row['DatasourceUsedByDataflowCount'] = $(if ($null -ne $use) { [int]$use.Dataflows } else { 0 })
        $sets['ScanDatasources'].Add($row)
    }
    return $sets
}

function Get-IQScanPreferredColumn {
    <#
    .SYNOPSIS
    The documented column order of one Scan* sheet (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$SheetName)
    switch ($SheetName) {
        'ScanWorkspaces' { return @('WorkspaceId', 'WorkspaceName', 'WorkspaceType', 'WorkspaceState', 'WorkspaceIsOnDedicatedCapacity', 'WorkspaceCapacityId', 'WorkspaceDefaultDatasetStorageFormat', 'WorkspaceDescription', 'WorkspaceReportCount', 'WorkspaceDashboardCount', 'WorkspaceDatasetCount', 'WorkspaceDataflowCount', 'WorkspaceDatamartCount', 'WorkspaceUserCount') }
        'ScanDatasets' { return @('WorkspaceId', 'WorkspaceName', 'DatasetId', 'DatasetName', 'DatasetDescription', 'DatasetConfiguredBy', 'DatasetConfiguredById', 'DatasetCreatedDate', 'DatasetContentProviderType', 'DatasetTargetStorageMode', 'DatasetIsRefreshable', 'DatasetIsEffectiveIdentityRequired', 'DatasetIsEffectiveIdentityRolesRequired', 'DatasetIsOnPremGatewayRequired', 'DatasetEndorsement', 'DatasetCertifiedBy', 'DatasetSensitivityLabelId', 'DatasetDatasourceInstanceIds', 'DatasetUpstreamDataflowIds', 'DatasetUpstreamDatamartIds', 'DatasetTableCount', 'DatasetColumnCount', 'DatasetMeasureCount', 'DatasetRelationshipCount', 'DatasetExpressionCount', 'DatasetRoleNames', 'DatasetUserCount', 'DatasetSchemaMayNotBeUpToDate', 'DatasetSchemaRetrievalError', 'DatasetRefreshScheduleJson', 'DatasetDirectQueryRefreshScheduleJson') }
        'ScanTables' { return @('WorkspaceId', 'WorkspaceName', 'DatasetId', 'DatasetName', 'TableName', 'TableDescription', 'TableIsHidden', 'TableStorageMode', 'TableSourceExpression', 'TableColumnCount', 'TableMeasureCount') }
        'ScanColumns' { return @('WorkspaceId', 'WorkspaceName', 'DatasetId', 'DatasetName', 'TableName', 'ColumnName', 'ColumnDataType', 'ColumnType', 'ColumnIsHidden', 'ColumnExpression', 'ColumnDescription') }
        'ScanMeasures' { return @('WorkspaceId', 'WorkspaceName', 'DatasetId', 'DatasetName', 'TableName', 'MeasureName', 'MeasureExpression', 'MeasureIsHidden', 'MeasureDescription') }
        'ScanExpressions' { return @('WorkspaceId', 'WorkspaceName', 'DatasetId', 'DatasetName', 'ExpressionName', 'Expression', 'ExpressionDescription') }
        'ScanDatasources' { return @('DatasourceInstanceId', 'DatasourceType', 'DatasourceGatewayId', 'DatasourceServer', 'DatasourceDatabase', 'DatasourceUrl', 'DatasourcePath', 'DatasourceKind', 'DatasourceAccount', 'DatasourceDomain', 'DatasourceEmailAddress', 'DatasourceLoginServer', 'DatasourceClassInfo', 'DatasourceConnectionDetailsJson', 'DatasourceIsMisconfigured', 'DatasourceUsedByDatasetCount', 'DatasourceUsedByDataflowCount') }
        'ScanReports' { return @('WorkspaceId', 'WorkspaceName', 'ReportId', 'ReportName', 'ReportType', 'ReportDatasetId', 'ReportDatasetWorkspaceId', 'ReportCreatedDateTime', 'ReportModifiedDateTime', 'ReportCreatedBy', 'ReportCreatedById', 'ReportModifiedBy', 'ReportModifiedById', 'ReportAppId', 'ReportOriginalReportObjectId', 'ReportDescription', 'ReportEndorsement', 'ReportCertifiedBy', 'ReportSensitivityLabelId', 'ReportUserCount') }
        'ScanDashboards' { return @('WorkspaceId', 'WorkspaceName', 'DashboardId', 'DashboardName', 'DashboardIsReadOnly', 'DashboardAppId', 'DashboardTileCount', 'DashboardTileReportIds', 'DashboardTileDatasetIds', 'DashboardSensitivityLabelId', 'DashboardUserCount') }
        'ScanDataflows' { return @('WorkspaceId', 'WorkspaceName', 'DataflowId', 'DataflowName', 'DataflowDescription', 'DataflowConfiguredBy', 'DataflowModifiedBy', 'DataflowModifiedDateTime', 'DataflowGeneration', 'DataflowEndorsement', 'DataflowCertifiedBy', 'DataflowSensitivityLabelId', 'DataflowDatasourceInstanceIds', 'DataflowUpstreamDataflowIds', 'DataflowRefreshScheduleJson', 'DataflowUserCount') }
        'ScanUsers' { return @('WorkspaceId', 'WorkspaceName', 'ArtifactType', 'ArtifactId', 'ArtifactName', 'UserEmailAddress', 'UserDisplayName', 'UserIdentifier', 'UserGraphId', 'UserPrincipalType', 'UserType', 'UserAccessRight') }
        default { return @() }
    }
}

function Get-IQAdminScanInventory {
    <#
    .SYNOPSIS
    Collector admin-scan: Scanner API over the tenant's (or the scope's) workspaces in batches of 100, results cached per batch under extracts\admin, flattened into the Scan* sheets.
    .DESCRIPTION
    A batch whose scan-<n>-<hash>.json already exists (non-empty) is reused - the ids are sorted before slicing so a
    retry produces the same batches, and a cached scan-*-<hash>.json is matched by its hash even when the batch
    number moved because a workspace appeared or disappeared - so a re-run only scans the batches that are missing.
    Returns @{ Success; Message; Outputs; WorkspaceCount; BatchCount; BatchesReused; BatchesFailed; RowCounts; Paused }.
    Success is $false when any batch failed (the item is re-tried on the next run, reusing the cached batches) or
    when the workspace list could not be read; a batch that timed out stops the remaining submissions for this run
    (a throttled tenant must not cascade into a series of 30-minute failures). Paused = $true (Success = $false, no
    sheets) when Test-IQTimeBudget stopped the loop: the caller must not checkpoint the item.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Stage = 'Extras')
    $collector = 'admin-scan'
    $result = @{ Success = $false; Message = ''; Outputs = @(); WorkspaceCount = 0; BatchCount = 0; BatchesReused = 0; BatchesFailed = 0; RowCounts = @{}; Paused = $false }
    $folder = Get-IQExtrasFolder -Name 'admin'
    $ids = @()
    try { $ids = @(Get-IQAdminScanWorkspaceId -Stage $Stage) }
    catch {
        $result.Message = 'Scanner API workspace list failed: ' + $_.Exception.Message
        return $result
    }
    # Deterministic batches (the id list is already sorted; kept here so the cache stays valid whatever the source order).
    $sortedIds = [string[]]@($ids | ForEach-Object { [string]$_ })
    [Array]::Sort($sortedIds, [System.StringComparer]::OrdinalIgnoreCase)
    $ids = @($sortedIds)
    $result.WorkspaceCount = $ids.Count
    $batchSize = [int](Get-IQExtrasOption -Name 'AdminScanBatchSize' -Default 100)
    if ($batchSize -lt 1 -or $batchSize -gt 100) { $batchSize = 100 }
    # Cached batch files of this run, by hash (scan-<n>-<hash8>.json): the index in the name is informational.
    $cachedByHash = @{}
    foreach ($f in @(Get-ChildItem -LiteralPath $folder -Filter 'scan-*.json' -File -ErrorAction SilentlyContinue)) {
        if ($f.Name -match '^scan-\d+-([0-9a-f]{8})\.json$' -and $f.Length -gt 0 -and -not $cachedByHash.ContainsKey($Matches[1])) { $cachedByHash[$Matches[1]] = $f.FullName }
    }
    $batches = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $ids.Count; $i += $batchSize) {
        $end = [math]::Min($ids.Count, $i + $batchSize) - 1
        $slice = @($ids[$i..$end])
        $index = $batches.Count + 1
        $hash = Get-IQAdminScanBatchHash -Ids $slice
        $file = Join-Path $folder ('scan-' + $index.ToString('000') + '-' + $hash + '.json')
        if (-not (Test-IQExtrasFileHasContent -Path $file) -and $cachedByHash.ContainsKey($hash)) { $file = $cachedByHash[$hash] }
        $batches.Add([ordered]@{ Index = $index; Ids = $slice; Hash = $hash; File = $file })
    }
    $layoutPath = Join-Path $folder 'scan-workspaces.json'
    ConvertTo-IQJsonFile -Object ([ordered]@{ CollectedUtc = [datetime]::UtcNow.ToString('o'); WorkspaceCount = $ids.Count; BatchSize = $batchSize; Batches = @($batches | ForEach-Object { [ordered]@{ Index = $_.Index; Hash = $_.Hash; File = $_.File; WorkspaceIds = $_.Ids } }) }) -Path $layoutPath
    $result.BatchCount = $batches.Count
    Write-IQLog -Level Info -Stage $Stage -Item $collector -Message ("Scanner API: {0} workspace(s) in {1} batch(es) of up to {2}" -f $ids.Count, $batches.Count, $batchSize)

    $scanResults = New-Object System.Collections.Generic.List[object]
    $failedBatches = @()
    $batchFiles = @()
    $stopReason = ''
    $pending = 0
    foreach ($b in $batches) {
        $item = ('admin-scan batch ' + $b.Index + '/' + $batches.Count)
        if (Test-IQExtrasFileHasContent -Path $b.File) {
            try {
                $cached = ConvertFrom-IQJsonFile -Path $b.File
                if ($null -ne $cached) {
                    $scanResults.Add($cached)
                    $batchFiles += $b.File
                    $result.BatchesReused++
                    Write-IQLog -Level Debug -Stage $Stage -Item $item -Message ('Reusing cached scan result ' + $b.File)
                    continue
                }
            }
            catch { Write-IQLog -Level Debug -Stage $Stage -Item $item -Message ('Cached scan result unreadable, scanning again: ' + $_.Exception.Message) }
        }
        if ($stopReason -ne '') { $pending++; continue }
        # Budget check before every submission: the finished batches are on disk, so stopping here loses nothing.
        if (Test-IQTimeBudget -Stage $Stage -Item $item) {
            $result.Paused = $true
            $result.Message = ('Time budget reached before batch {0} of {1}; {2} batch(es) cached - the scan continues on the next start' -f $b.Index, $batches.Count, $scanResults.Count)
            Write-IQLog -Level Warn -Stage $Stage -Item $collector -Message $result.Message
            return $result
        }
        try {
            $scan = Invoke-IQAdminScanBatch -Ids $b.Ids -Stage $Stage -Item $item -OutPath $b.File
            $scanResults.Add($scan)
            $batchFiles += $b.File
            Write-IQLog -Level Info -Stage $Stage -Item $item -Message ("Scanned {0} workspace(s)" -f $b.Ids.Count)
        }
        catch {
            if ($_.Exception -is [System.OperationCanceledException]) {
                $result.Paused = $true
                $result.Message = ('Time budget reached during batch {0} of {1}: {2}' -f $b.Index, $batches.Count, $_.Exception.Message)
                Write-IQLog -Level Warn -Stage $Stage -Item $collector -Message $result.Message
                return $result
            }
            $failedBatches += $b.Index
            $result.BatchesFailed++
            Write-IQLog -Level Warn -Stage $Stage -Item $item -Message ('Scan batch failed: ' + $_.Exception.Message)
            if ($_.Exception -is [System.TimeoutException]) {
                $stopReason = ('batch {0} timed out ({1}); the remaining batches are not submitted in this run and are scanned on the next start' -f $b.Index, $_.Exception.Message)
                Write-IQLog -Level Warn -Stage $Stage -Item $collector -Message $stopReason
            }
        }
    }

    $sets = ConvertTo-IQScanRowSet -ScanResults $scanResults.ToArray()
    $outputs = @()
    foreach ($name in @('ScanWorkspaces', 'ScanDatasets', 'ScanTables', 'ScanColumns', 'ScanMeasures', 'ScanExpressions', 'ScanDatasources', 'ScanReports', 'ScanDashboards', 'ScanDataflows', 'ScanUsers')) {
        $rows = @($sets[$name].ToArray())
        $result.RowCounts[$name] = $rows.Count
        $outputs += Save-IQExtrasSheet -SheetName $name -Collector $collector -Rows $rows -PreferredColumns (Get-IQScanPreferredColumn -SheetName $name) -KeepPreferred
    }
    $result.Outputs = @($outputs) + @($batchFiles) + @($layoutPath)
    $summary = ('{0} workspace(s) in {1} batch(es) ({2} reused): {3} datasets, {4} tables, {5} columns, {6} measures, {7} datasources, {8} reports, {9} dashboards, {10} dataflows, {11} user rows' -f $ids.Count, $batches.Count, $result.BatchesReused, $result.RowCounts['ScanDatasets'], $result.RowCounts['ScanTables'], $result.RowCounts['ScanColumns'], $result.RowCounts['ScanMeasures'], $result.RowCounts['ScanDatasources'], $result.RowCounts['ScanReports'], $result.RowCounts['ScanDashboards'], $result.RowCounts['ScanDataflows'], $result.RowCounts['ScanUsers'])
    if ($failedBatches.Count -gt 0) {
        $result.Message = ('{0} of {1} scan batch(es) failed (batch {2}); partial sheets written from the other batches. ' -f $failedBatches.Count, $batches.Count, ($failedBatches -join ', ')) + $summary
        if ($pending -gt 0) { $result.Message = ('{0} batch(es) not submitted after a timeout; ' -f $pending) + $result.Message }
        return $result
    }
    $result.Success = $true
    $result.Message = $summary
    return $result
}

# =====================================================================================================================
# admin/activityevents (one UTC day per item)
# =====================================================================================================================

function Get-IQAdminActivityDay {
    <#
    .SYNOPSIS
    Reads every activity event of one UTC day (GET admin/activityevents with continuationUri paging) and saves extracts\admin\activity-<day>.json (private).
    .DESCRIPTION
    Returns @{ Success; Message; Path; EventCount; Pages; Partial }. Never throws for API failures. A continuation
    page that returns nothing (HTTP 400/403/404, e.g. an expired continuation token) fails the day - a truncated day
    must not be checkpointed Succeeded, it is re-read on the next start.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][datetime]$Day,
        [Parameter(Mandatory = $false)][string]$Stage = 'Extras'
    )
    $dayText = $Day.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    $item = 'admin-activity-' + $dayText
    $folder = Get-IQExtrasFolder -Name 'admin'
    $path = Join-Path $folder ('activity-' + $dayText + '.json')
    $partial = ($dayText -eq (Get-IQExtrasNowUtc).ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture))
    $result = @{ Success = $false; Message = ''; Path = $path; EventCount = 0; Pages = 0; Partial = $partial }
    $start = "'" + $dayText + "T00:00:00.000Z'"
    $end = "'" + $dayText + "T23:59:59.999Z'"
    $events = New-Object System.Collections.Generic.List[object]
    $pages = 0
    try {
        $response = Invoke-IQApi -Method GET -Path 'admin/activityevents' -Query @{ startDateTime = $start; endDateTime = $end } -NoPaging -Stage $Stage
        if ($null -eq $response) {
            $result.Message = 'GET admin/activityevents returned no response (HTTP 400/403/404 - see the Warn line above)'
            return $result
        }
        while ($null -ne $response) {
            $pages++
            foreach ($e in @(Get-IQExtrasMember -Object $response -Name 'activityEventEntities')) { if ($null -ne $e) { $events.Add($e) } }
            $next = [string](Get-IQExtrasMember -Object $response -Name 'continuationUri')
            $last = Get-IQExtrasMember -Object $response -Name 'lastResultSet'
            if ([string]::IsNullOrWhiteSpace($next) -or ($null -ne $last -and [bool]$last)) { break }
            if ($pages -ge 10000) { Write-IQLog -Level Warn -Stage $Stage -Item $item -Message 'activityevents paging guard hit (10000 pages)'; break }
            $response = Invoke-IQApi -Method GET -Path $next -NoPaging -Stage $Stage
            if ($null -eq $response) {
                $result.Message = ('activityevents continuation page {0} returned no response (HTTP 400/403/404 - see the Warn line above) after {1} event(s); the day is re-read on the next start' -f ($pages + 1), $events.Count)
                return $result
            }
        }
    }
    catch {
        $result.Message = 'GET admin/activityevents failed: ' + $_.Exception.Message
        return $result
    }
    $result.Pages = $pages
    $result.EventCount = $events.Count
    try {
        ConvertTo-IQJsonFile -Object ([ordered]@{ Date = $dayText; EventCount = $events.Count; Pages = $pages; Partial = $partial; CollectedUtc = [datetime]::UtcNow.ToString('o'); Events = $events.ToArray() }) -Path $path
    }
    catch {
        $result.Message = 'Could not write ' + $path + ': ' + $_.Exception.Message
        return $result
    }
    $result.Success = $true
    $result.Message = ('{0} event(s) in {1} page(s)' -f $events.Count, $pages)
    if ($partial) { $result.Message += ' (today, partial day)' }
    return $result
}

function Get-IQActivityPreferredColumn {
    <#
    .SYNOPSIS
    The leading column order of the ActivityEvents sheet (private).
    #>
    [CmdletBinding()]
    param()
    return @('Id', 'CreationTime', 'Operation', 'Activity', 'UserId', 'UserType', 'UserKey', 'Workload', 'ItemName', 'WorkSpaceName', 'WorkspaceId', 'CapacityId', 'CapacityName', 'DatasetName', 'DatasetId', 'ReportName', 'ReportId', 'ReportType', 'ArtifactKind', 'ArtifactId', 'ArtifactName', 'ObjectId', 'DistributionMethod', 'ConsumptionMethod', 'ClientIP', 'UserAgent', 'ActivityId', 'RequestId', 'IsSuccess', 'RefreshType', 'DataflowName', 'DataflowId', 'DashboardName', 'DashboardId', 'AppName', 'AppId', 'RecordType', 'OrganizationId', 'ActivityDate')
}

function Save-IQActivitySheet {
    <#
    .SYNOPSIS
    Rebuilds the ActivityEvents sheet from every extracts\admin\activity-*.json of the run (capped at Options.ActivityMaxRows) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Stage = 'Extras')
    $collector = 'admin-activity'
    $folder = Get-IQExtrasFolder -Name 'admin'
    # Windows PowerShell 5.1 serialises a sheet of this size an order of magnitude slower than PowerShell 7 (the
    # per-day JSON files keep every event either way), so the default cap is lower there.
    $defaultMax = 250000
    if ($PSVersionTable.PSVersion.Major -lt 6) { $defaultMax = 100000 }
    $maxRows = [int](Get-IQExtrasOption -Name 'ActivityMaxRows' -Default $defaultMax)
    if ($maxRows -lt 1) { $maxRows = $defaultMax }
    $rows = New-Object System.Collections.Generic.List[object]
    $days = 0
    $truncated = $false
    foreach ($f in @(Get-ChildItem -LiteralPath $folder -Filter 'activity-*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        try {
            $doc = ConvertFrom-IQJsonFile -Path $f.FullName
            if ($null -eq $doc) { continue }
            $days++
            $date = [string](Get-IQExtrasMember -Object $doc -Name 'Date')
            foreach ($e in @(Get-IQExtrasMember -Object $doc -Name 'Events')) {
                if ($null -eq $e) { continue }
                if ($rows.Count -ge $maxRows) { $truncated = $true; break }
                $row = ConvertTo-IQExtrasRow -Object $e
                $row['ActivityDate'] = $date
                $rows.Add($row)
            }
        }
        catch { Write-IQLog -Level Warn -Stage $Stage -Item $collector -Message ("Could not read '{0}': {1}" -f $f.FullName, $_.Exception.Message) }
        if ($truncated) { break }
    }
    $message = ('{0} event(s) from {1} day file(s)' -f $rows.Count, $days)
    if ($truncated) {
        $message += (' - truncated at ActivityMaxRows={0}; the per-day JSON files under extracts\admin hold everything' -f $maxRows)
        Write-IQLog -Level Warn -Stage $Stage -Item $collector -Message $message
    }
    return (Save-IQExtrasSheet -SheetName 'ActivityEvents' -Collector $collector -Rows $rows.ToArray() -PreferredColumns (Get-IQActivityPreferredColumn) -Message $message)
}

function Get-IQAdminActivityDayList {
    <#
    .SYNOPSIS
    The UTC days to collect: the last -ActivityDays days ending on the run's start date, clamped to Options.ActivityMaxDays (28, the API window) (private).
    .DESCRIPTION
    The window ends on the day the run started (manifest startedUtc when this is a resumed run, else the run clock -
    Options.NowUtc in tests), so a run resumed days later collects the same days it checkpointed. The entry point's
    default -ActivityDays 30 is above the API window by design (collect everything available): that clamp is logged
    at Info; an explicit larger value is logged at Warn.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Stage = 'Extras')
    $days = [int](Get-IQExtrasOption -Name 'ActivityDays' -Default 30)
    $maxDays = [int](Get-IQExtrasOption -Name 'ActivityMaxDays' -Default 28)
    if ($maxDays -lt 1) { $maxDays = 28 }
    if ($days -lt 1) { $days = 1 }
    if ($days -gt $maxDays) {
        $level = 'Warn'
        if ($days -eq 30) { $level = 'Info' }
        Write-IQLog -Level $level -Stage $Stage -Item 'admin-activity' -Message ("-ActivityDays {0} exceeds the activity log window; collecting the last {1} day(s)" -f $days, $maxDays)
        $days = $maxDays
    }
    $today = (Get-IQExtrasNowUtc).Date
    if ($script:IQ -and $script:IQ.ContainsKey('IsResume') -and [bool]$script:IQ['IsResume'] -and $null -ne $script:IQ.Manifest) {
        $startedText = [string](Get-IQExtrasMember -Object $script:IQ.Manifest -Name 'startedUtc')
        if (-not [string]::IsNullOrWhiteSpace($startedText)) {
            try {
                $started = [datetime]::Parse($startedText, [System.Globalization.CultureInfo]::InvariantCulture, ([System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal))
                if ($started.Date -lt $today) {
                    Write-IQLog -Level Debug -Stage $Stage -Item 'admin-activity' -Message ('Resumed run: the activity window ends on the run start date ' + $started.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture))
                    $today = $started.Date
                }
            }
            catch { Write-IQLog -Level Debug -Stage $Stage -Item 'admin-activity' -Message ('manifest startedUtc not parseable (' + $startedText + '); using the clock') }
        }
    }
    $list = @()
    for ($i = $days - 1; $i -ge 0; $i--) { $list += $today.AddDays(-$i) }
    return $list
}

# =====================================================================================================================
# Usage metrics (per workspace, via Get-IQUsageMetrics in ImpactIQ.Dax.ps1)
# =====================================================================================================================

function Get-IQUsageWorkspaceList {
    <#
    .SYNOPSIS
    Real (GUID, non-synthetic) workspaces of the run scope for the usage collector (private).
    #>
    [CmdletBinding()]
    param()
    $out = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($w in @(Get-IQSelectedWorkspaces)) {
        if ($null -eq $w) { continue }
        $id = [string](Get-IQExtrasMember -Object $w -Name 'WorkspaceId')
        if (-not (Test-IQExtrasGuid -Value $id)) { continue }
        if ([bool](Get-IQExtrasMember -Object $w -Name 'WorkspaceIsSynthetic')) { continue }
        $key = $id.ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $out.Add([PSCustomObject]@{ WorkspaceId = $id; WorkspaceName = [string](Get-IQExtrasMember -Object $w -Name 'WorkspaceName') })
    }
    return $out.ToArray()
}

function Save-IQUsageSheet {
    <#
    .SYNOPSIS
    Rebuilds the UsageReportViews / UsageReportPageViews sheets from every extracts\usage\*.json of the run (private). Returns the two file paths.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Stage = 'Extras')
    $collector = 'usage'
    $folder = Get-IQExtrasFolder -Name 'usage'
    $views = New-Object System.Collections.Generic.List[object]
    $pageViews = New-Object System.Collections.Generic.List[object]
    $workspaces = 0
    $found = 0
    foreach ($f in @(Get-ChildItem -LiteralPath $folder -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        try {
            $doc = ConvertFrom-IQJsonFile -Path $f.FullName
            if ($null -eq $doc) { continue }
            $workspaces++
            if ([bool](Get-IQExtrasMember -Object $doc -Name 'Found')) { $found++ }
            foreach ($r in @(Get-IQExtrasMember -Object $doc -Name 'ReportViews')) { if ($null -ne $r) { $views.Add($r) } }
            foreach ($r in @(Get-IQExtrasMember -Object $doc -Name 'ReportPageViews')) { if ($null -ne $r) { $pageViews.Add($r) } }
        }
        catch { Write-IQLog -Level Warn -Stage $Stage -Item $collector -Message ("Could not read '{0}': {1}" -f $f.FullName, $_.Exception.Message) }
    }
    $lead = @('WorkspaceId', 'WorkspaceName', 'UsageDatasetId')
    $message = ('{0} workspace(s) checked, {1} with a usage model' -f $workspaces, $found)
    $paths = @()
    $paths += Save-IQExtrasSheet -SheetName 'UsageReportViews' -Collector $collector -Rows $views.ToArray() -PreferredColumns ($lead + @('Date', 'ReportId', 'ReportName', 'UserId', 'UserKey', 'UserPrincipalName', 'ConsumptionMethod', 'DistributionMethod', 'ReportType', 'AppName', 'CapacityId', 'CapacityName', 'DatasetName', 'UserAgent', 'CreationTime')) -Message $message
    $paths += Save-IQExtrasSheet -SheetName 'UsageReportPageViews' -Collector $collector -Rows $pageViews.ToArray() -PreferredColumns ($lead + @('Date', 'ReportId', 'ReportName', 'SectionId', 'UserId', 'UserKey', 'UserPrincipalName', 'Client', 'SessionSource', 'Timestamp')) -Message $message
    return $paths
}

function Invoke-IQUsageCollector {
    <#
    .SYNOPSIS
    Collector usage-<workspaceId>: Get-IQUsageMetrics per real workspace in scope, raw result under extracts\usage, one checkpoint per workspace, then the two Usage* sheets (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Stage = 'Extras')
    $summary = @{ Total = 0; Done = 0; Skipped = 0; Failed = 0; AlreadyDone = 0; Outputs = @(); BudgetStop = $false }
    $folder = Get-IQExtrasFolder -Name 'usage'
    $days = [int](Get-IQExtrasOption -Name 'UsageDays' -Default 30)
    if ($days -lt 1) { $days = 30 }
    $workspaces = @(Get-IQUsageWorkspaceList)
    $summary.Total = $workspaces.Count
    Write-IQLog -Level Info -Stage $Stage -Item 'usage' -Message ("Usage metrics: {0} workspace(s) in scope, last {1} day(s)" -f $workspaces.Count, $days)
    $index = 0
    foreach ($w in $workspaces) {
        $index++
        $key = 'usage-' + $w.WorkspaceId
        $item = $w.WorkspaceName
        if ([string]::IsNullOrWhiteSpace($item)) { $item = $w.WorkspaceId }
        if (Test-IQItemDone -Stage $Stage -ItemKey $key) {
            $summary.AlreadyDone++
            Write-IQLog -Level Debug -Stage $Stage -Item $item -Message 'Usage metrics already collected (checkpoint); skipping'
            continue
        }
        if (Test-IQTimeBudget -Stage $Stage -Item $item) {
            Write-IQLog -Level Warn -Stage $Stage -Item 'usage' -Message ("Time budget reached: {0} of {1} workspace(s) without usage metrics yet - collected on the next start." -f ($workspaces.Count - $index + 1), $workspaces.Count)
            $summary.BudgetStop = $true
            break
        }
        Write-IQLog -Level Info -Stage $Stage -Item $item -Message ("[{0}/{1}] Reading usage metrics" -f $index, $workspaces.Count)
        $path = Join-Path $folder ((Get-IQSafeKey -Value $w.WorkspaceId) + '.json')
        $usage = $null
        try { $usage = Get-IQUsageMetrics -WorkspaceId $w.WorkspaceId -WorkspaceName $w.WorkspaceName -Days $days }
        catch {
            Set-IQItemDone -Stage $Stage -ItemKey $key -Item $item -Status Failed -Method 'executeQueries' -Message ('Usage metrics failed: ' + $_.Exception.Message) | Out-Null
            $summary.Failed++
            continue
        }
        $message = [string](Get-IQExtrasMember -Object $usage -Name 'Message')
        $found = [bool](Get-IQExtrasMember -Object $usage -Name 'Found')
        $data = @{
            WorkspaceId     = $w.WorkspaceId
            WorkspaceName   = $w.WorkspaceName
            UsageDatasetId  = [string](Get-IQExtrasMember -Object $usage -Name 'DatasetId')
            UsageDatasetName = [string](Get-IQExtrasMember -Object $usage -Name 'DatasetName')
            ReportViews     = @(Get-IQExtrasMember -Object $usage -Name 'ReportViews').Count
            ReportPageViews = @(Get-IQExtrasMember -Object $usage -Name 'ReportPageViews').Count
            Days            = $days
            ExtractFile     = $path
        }
        if (-not $found) {
            if ($message -like 'Usage metrics failed*') {
                Set-IQItemDone -Stage $Stage -ItemKey $key -Item $item -Status Failed -Method 'executeQueries' -Message $message -Data $data | Out-Null
                $summary.Failed++
            }
            else {
                Set-IQItemDone -Stage $Stage -ItemKey $key -Item $item -Status Skipped -Method 'executeQueries' -Message $message -Data $data | Out-Null
                $summary.Skipped++
            }
            continue
        }
        try {
            $doc = [ordered]@{
                SchemaVersion   = 1
                WorkspaceId     = $w.WorkspaceId
                WorkspaceName   = $w.WorkspaceName
                Found           = $found
                DatasetId       = $data.UsageDatasetId
                DatasetName     = $data.UsageDatasetName
                Days            = $days
                Message         = $message
                CollectedUtc    = [datetime]::UtcNow.ToString('o')
                ReportViews     = @(Get-IQExtrasMember -Object $usage -Name 'ReportViews')
                ReportPageViews = @(Get-IQExtrasMember -Object $usage -Name 'ReportPageViews')
                Reports         = @(Get-IQExtrasMember -Object $usage -Name 'Reports')
                Users           = @(Get-IQExtrasMember -Object $usage -Name 'Users')
            }
            ConvertTo-IQJsonFile -Object $doc -Path $path
        }
        catch {
            Set-IQItemDone -Stage $Stage -ItemKey $key -Item $item -Status Failed -Method 'executeQueries' -Message ('Usage metrics read but the extract could not be written: ' + $_.Exception.Message) -Data $data | Out-Null
            $summary.Failed++
            continue
        }
        if ($message -like 'Usage metrics failed*') {
            Set-IQItemDone -Stage $Stage -ItemKey $key -Item $item -Status Failed -Method 'executeQueries' -Message $message -Outputs @($path) -Data $data | Out-Null
            $summary.Failed++
            continue
        }
        Set-IQItemDone -Stage $Stage -ItemKey $key -Item $item -Outputs @($path) -Method 'executeQueries' -Message $message -Data $data | Out-Null
        $summary.Done++
    }
    $summary.Outputs = @(Save-IQUsageSheet -Stage $Stage)
    Write-IQLog -Level Info -Stage $Stage -Item 'usage' -Message ("Usage metrics finished: {0} collected, {1} without a usage model, {2} failed, {3} already done of {4}" -f $summary.Done, $summary.Skipped, $summary.Failed, $summary.AlreadyDone, $summary.Total)
    return $summary
}

# =====================================================================================================================
# Stage body
# =====================================================================================================================

function Invoke-IQExtrasCollector {
    <#
    .SYNOPSIS
    Runs one single-checkpoint collector (admin-groups / admin-scan): skip when done, run, checkpoint Succeeded/Failed; a body that reports Paused (time budget) is not checkpointed (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ItemKey,
        [Parameter(Mandatory = $true)][string]$Item,
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][scriptblock]$Body,
        [Parameter(Mandatory = $false)][string]$Stage = 'Extras'
    )
    if (Test-IQItemDone -Stage $Stage -ItemKey $ItemKey) {
        $previous = Get-IQItemCheckpoint -Stage $Stage -ItemKey $ItemKey
        $previousStatus = [string](Get-IQExtrasMember -Object $previous -Name 'status')
        if ($previousStatus -ne 'Skipped') {
            Write-IQLog -Level Info -Stage $Stage -Item $Item -Message 'Already collected (checkpoint); skipping'
            return 'AlreadyDone'
        }
        # A Skipped checkpoint means the admin probe failed on an earlier run; now that it succeeded, collect for real.
        Write-IQLog -Level Info -Stage $Stage -Item $Item -Message 'Previously skipped (admin probe failed); collecting now'
    }
    if (Test-IQTimeBudget -Stage $Stage -Item $Item) {
        Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message 'Time budget reached - this collector runs on the next start.'
        return 'Paused'
    }
    Write-IQLog -Level Info -Stage $Stage -Item $Item -Message 'Collecting'
    $result = $null
    try { $result = & $Body }
    catch {
        $result = @{ Success = $false; Message = ('Unexpected error: ' + $_.Exception.Message); Outputs = @() }
        Write-IQLog -Level Debug -Stage $Stage -Item $Item -Message $_.Exception.ToString()
    }
    if ($null -eq $result -or -not ($result -is [System.Collections.IDictionary])) {
        $result = @{ Success = $false; Message = 'Collector returned no result'; Outputs = @() }
    }
    if ($result.Contains('Paused') -and [bool]$result['Paused']) {
        # The budget stopped the collector mid-way: its cached files stay, no checkpoint is written, and Invoke-IQStage
        # marks the stage Paused (Test-IQTimeBudget set BudgetExceeded) so the next start resumes it.
        Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message ('Time budget reached - ' + [string]$result.Message)
        return 'Paused'
    }
    $data = @{}
    foreach ($k in @($result.Keys)) {
        if ($k -in @('Success', 'Message', 'Outputs', 'Paused')) { continue }
        $data[$k] = $result[$k]
    }
    if (-not $result.Success) {
        Set-IQItemDone -Stage $Stage -ItemKey $ItemKey -Item $Item -Status Failed -Method $Method -Message ([string]$result.Message) -Data $data | Out-Null
        return 'Failed'
    }
    Set-IQItemDone -Stage $Stage -ItemKey $ItemKey -Item $Item -Outputs @($result.Outputs) -Method $Method -Message ([string]$result.Message) -Data $data | Out-Null
    Write-IQLog -Level Success -Stage $Stage -Item $Item -Message ([string]$result.Message)
    return 'Succeeded'
}

function Invoke-IQExtrasStage {
    <#
    .SYNOPSIS
    Extras stage body: optional admin-API collectors (-IncludeAdminApis) and per-workspace usage metrics (-IncludeUsageMetrics), each an item checkpoint.
    .DESCRIPTION
    -IncludeAdminApis: probe admin/capacities; when the user is a Fabric administrator run admin-groups, admin-scan and
    one admin-activity-<day> item per UTC day of the last -ActivityDays days (max 28); when the probe says "not an
    administrator" (401/403/404) record the two admin items as Skipped with the reason; when the probe failed for a
    transient reason (5xx, network, token) record them as Failed so the stage ends CompletedWithErrors and the next
    start retries them. -IncludeUsageMetrics: usage-<workspaceId> per real workspace in scope. Sheet
    files are rebuilt from the raw extracts on every run so Assemble always sees the complete data. Returns
    @{ AdminRequested; AdminAvailable; UsageRequested; Items = @{ <itemKey> = <status> }; Usage = <usage summary>;
    Outputs }.
    #>
    [CmdletBinding()]
    param()
    $stage = 'Extras'
    $summary = @{ AdminRequested = $false; AdminAvailable = $false; UsageRequested = $false; Items = [ordered]@{}; Usage = $null; Outputs = @() }
    $summary.AdminRequested = [bool](Get-IQExtrasOption -Name 'IncludeAdminApis' -Default $false)
    $summary.UsageRequested = [bool](Get-IQExtrasOption -Name 'IncludeUsageMetrics' -Default $false)
    if (-not $summary.AdminRequested -and -not $summary.UsageRequested) {
        Write-IQLog -Level Info -Stage $stage -Message 'No optional collectors requested (-IncludeAdminApis / -IncludeUsageMetrics); nothing to do.'
        return $summary
    }

    if ($summary.AdminRequested) {
        $probe = Test-IQAdminAccess -Stage $stage
        $summary.AdminAvailable = [bool]$probe.IsAdmin
        if (-not $probe.IsAdmin) {
            Write-IQLog -Level Warn -Stage $stage -Message $probe.Message
            $probeStatus = 'Skipped'
            if ([bool](Get-IQExtrasMember -Object $probe -Name 'Transient')) { $probeStatus = 'Failed' }
            foreach ($pair in @(@('admin-groups', 'Admin workspaces (admin/groups)'), @('admin-scan', 'Scanner API (admin/workspaces)'))) {
                if (Test-IQItemDone -Stage $stage -ItemKey $pair[0]) {
                    $previousStatus = [string](Get-IQExtrasMember -Object (Get-IQItemCheckpoint -Stage $stage -ItemKey $pair[0]) -Name 'status')
                    if ($previousStatus -eq 'Succeeded') { $summary.Items[$pair[0]] = 'AlreadyDone'; continue }
                    if ($probeStatus -eq 'Skipped') { $summary.Items[$pair[0]] = 'Skipped'; continue }
                }
                Set-IQItemDone -Stage $stage -ItemKey $pair[0] -Item $pair[1] -Status $probeStatus -Method 'AdminApi' -Message $probe.Message | Out-Null
                $summary.Items[$pair[0]] = $probeStatus
            }
        }
        else {
            Write-IQLog -Level Info -Stage $stage -Message $probe.Message
            $summary.Items['admin-groups'] = Invoke-IQExtrasCollector -ItemKey 'admin-groups' -Item 'Admin workspaces (admin/groups)' -Method 'AdminApi' -Stage $stage -Body { Get-IQAdminGroupInventory -Stage 'Extras' }
            $summary.Items['admin-scan'] = Invoke-IQExtrasCollector -ItemKey 'admin-scan' -Item 'Scanner API (admin/workspaces)' -Method 'ScannerApi' -Stage $stage -Body { Get-IQAdminScanInventory -Stage 'Extras' }

            $days = @(Get-IQAdminActivityDayList -Stage $stage)
            $activityDone = 0
            $activityFailed = 0
            $activitySkipped = 0
            foreach ($day in $days) {
                $dayText = $day.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
                $key = 'admin-activity-' + $dayText
                $item = 'Activity events ' + $dayText
                if (Test-IQItemDone -Stage $stage -ItemKey $key) {
                    $partial = $false
                    try { $partial = [bool](Get-IQExtrasMember -Object (Get-IQExtrasMember -Object (Get-IQItemCheckpoint -Stage $stage -ItemKey $key) -Name 'data') -Name 'Partial') } catch { $partial = $false }
                    if (-not $partial) {
                        $activitySkipped++
                        $summary.Items[$key] = 'AlreadyDone'
                        continue
                    }
                    Write-IQLog -Level Debug -Stage $stage -Item $item -Message 'Re-reading the partial (current) day'
                }
                if (Test-IQTimeBudget -Stage $stage -Item $item) {
                    Write-IQLog -Level Warn -Stage $stage -Item 'admin-activity' -Message 'Time budget reached - the remaining activity days are collected on the next start.'
                    $summary.Items[$key] = 'Paused'
                    break
                }
                $dayResult = Get-IQAdminActivityDay -Day $day -Stage $stage
                $data = @{ Date = $dayText; EventCount = [int]$dayResult.EventCount; Pages = [int]$dayResult.Pages; Partial = [bool]$dayResult.Partial; ExtractFile = [string]$dayResult.Path }
                if ($dayResult.Success) {
                    Set-IQItemDone -Stage $stage -ItemKey $key -Item $item -Outputs @($dayResult.Path) -Method 'AdminApi' -Message ([string]$dayResult.Message) -Data $data | Out-Null
                    $activityDone++
                    $summary.Items[$key] = 'Succeeded'
                }
                else {
                    Set-IQItemDone -Stage $stage -ItemKey $key -Item $item -Status Failed -Method 'AdminApi' -Message ([string]$dayResult.Message) -Data $data | Out-Null
                    $activityFailed++
                    $summary.Items[$key] = 'Failed'
                }
            }
            if ($script:IQ.ContainsKey('BudgetExceeded') -and [bool]$script:IQ['BudgetExceeded']) {
                # Rebuilding a sheet of up to ActivityMaxRows rows does not fit in the 2-minute shutdown grace; the day
                # files are complete and the sheet is rebuilt when the run resumes.
                Write-IQLog -Level Warn -Stage $stage -Item 'admin-activity' -Message 'Time budget reached - the ActivityEvents sheet is rebuilt from the day files on the next start.'
            }
            else {
                $summary.Outputs += Save-IQActivitySheet -Stage $stage
            }
            $level = 'Info'
            if ($activityFailed -gt 0) { $level = 'Warn' }
            Write-IQLog -Level $level -Stage $stage -Item 'admin-activity' -Message ("Activity events: {0} day(s) collected, {1} failed, {2} already done of {3}" -f $activityDone, $activityFailed, $activitySkipped, $days.Count)
        }
    }

    if ($summary.UsageRequested -and (Test-IQTimeBudget -Stage $stage -Item 'usage')) {
        Write-IQLog -Level Warn -Stage $stage -Item 'usage' -Message 'Time budget reached - usage metrics are collected on the next start.'
        $summary.Items['usage'] = 'Paused'
    }
    elseif ($summary.UsageRequested) {
        $summary.Usage = Invoke-IQUsageCollector -Stage $stage
        $summary.Outputs += @($summary.Usage.Outputs)
    }
    return $summary
}
