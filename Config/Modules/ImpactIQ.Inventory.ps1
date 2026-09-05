#Requires -Version 5.1
<#
    ImpactIQ v3 - Inventory stage (ImpactIQ.Inventory.ps1)

    Ports the REST metadata extraction of the original monolith ("Final PS Script.txt"):
      - rename maps (lines 1383-1521) and Rename-Properties (637-652) - copied verbatim
      - scope selection (1530-1860) as Resolve-IQScope: headless parameters first, interactive dialogs
        (Select-IQScopeInteractive from ImpactIQ.Interactive.ps1) only when nothing was given and the run is interactive
      - Fabric connections/gateways, apps + app reports (1874-2065) into global.json
      - the per-workspace loops (1940-2356) merged into ONE pass per workspace (Get-IQWorkspaceInventory), written to
        State\runs\<RunId>\inventory\ws-<id>.json and checkpointed with Set-IQItemDone -Stage Inventory
      - My Workspace (2356-2513) with the same shape (WorkspaceId 'My Workspace', WorkspaceIsSynthetic = $true)
      - NEW non-admin collectors (brief 6.1/6.2): Dashboards, DashboardTiles, WorkspaceUsers, DatasetUsers,
        DatasetParameters, DatasetDirectQueryRefreshSchedule, Capacities - each wrapped so a failure only yields an
        empty list plus one Debug log line.

    Every HTTP call goes through Invoke-IQApi (ImpactIQ.Http.ps1). Nothing here uses Read-Host, WinForms or globals;
    all shared state lives in $script:IQ (created by Initialize-IQContext in ImpactIQ.Common.ps1).
    Windows PowerShell 5.1 compatible.
#>

# =====================================================================================================================
# Rename maps and Rename-Properties (monolith lines 637-652 and 1383-1521, verbatim)
# =====================================================================================================================

function Get-IQRenameMap {
    <#
    .SYNOPSIS
        Returns one of the monolith's property rename maps (verbatim copies of "Final PS Script.txt" lines 1383-1521).
    .PARAMETER Name
        Workspace, Dataset, DatasetDatasource, DataflowDatasource, DatasetRefresh, DatasetRefreshSchedule, Dataflow,
        FabricDataflow, DataflowRefresh, DataflowLineage, Report, Page, App, AppReport or FabricItems.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Workspace', 'Dataset', 'DatasetDatasource', 'DataflowDatasource', 'DatasetRefresh', 'DatasetRefreshSchedule',
            'Dataflow', 'FabricDataflow', 'DataflowRefresh', 'DataflowLineage', 'Report', 'Page', 'App', 'AppReport', 'FabricItems')]
        [string]$Name
    )

    $workspaceRenameMap = @{
        "id" = "WorkspaceId";
        "name" = "WorkspaceName";
        "isReadOnly" = "WorkspaceIsReadOnly";
        "isOnDedicatedCapacity" = "WorkspaceIsOnDedicatedCapacity";
        "capacityId" = "WorkspaceCapacityId";
        "defaultDatasetStorageFormat" = "WorkspaceDefaultDatasetStorageFormat";
        "type" = "WorkspaceType"
    }

    $datasetRenameMap = @{
        "id" = "DatasetId";
        "name" = "DatasetName";
        "description" = "DatasetDescription";
        "webUrl" = "DatasetWebUrl";
        "addRowsAPIEnabled" = "DatasetAddRowsAPIEnabled";
        "configuredBy" = "DatasetConfiguredBy";
        "isRefreshable" = "DatasetIsRefreshable";
        "isEffectiveIdentityRequired" = "DatasetIsEffectiveIdentityRequired";
        "isEffectiveIdentityRolesRequired" = "DatasetIsEffectiveIdentityRolesRequired";
        "isOnPremGatewayRequired" = "DatasetIsOnPremGatewayRequired";
        "targetStorageMode" = "DatasetTargetStorageMode";
        "queryScaleOutSettings" = "DatasetQueryScaleOutSettings";
        "createdDate" = "DatasetCreatedDate"
    }

    $datasetDatasourceRenameMap = @{
        "datasourceType" = "DatasetDatasourceType";
        "datasourceId" = "DatasetDatasourceId";
        "gatewayId" = "DatasetDatasourceGatewayId";
        "connectionDetails" = "DatasetDatasourceConnectionDetails"
    }

    $dataflowDatasourceRenameMap = @{
        "datasourceType" = "DataflowDatasourceType";
        "datasourceId" = "DataflowDatasourceId";
        "gatewayId" = "DataflowDatasourceGatewayId";
        "connectionDetails" = "DataflowDatasourceConnectionDetails"
    }

    $datasetRefreshRenameMap = @{
        "requestId" = "DatasetRefreshRequestId";
        "id" = "DatasetRefreshId";
        "startTime" = "DatasetRefreshStartTime";
        "endTime" = "DatasetRefreshEndTime";
        "status" = "DatasetRefreshStatus";
        "refreshType" = "DatasetRefreshType"
    }

    $datasetRefreshScheduleRenameMap = @{
        "enabled" = "DatasetRefreshScheduleEnabled";
        "localTimeZoneId" = "DatasetRefreshScheduleLocalTimeZoneId";
        "notifyOption" = "DatasetRefreshScheduleNotifyOption"
    }

    $dataflowRenameMap = @{
        "configuredBy"      = "DataflowConfiguredBy";
        "description"       = "DataflowDescription";
        "modelUrl"         = "DataflowJsonURL";
        "modifiedBy"       = "DataflowModifiedBy";
        "modifiedDateTime" = "DataflowModifiedDateTime";
        "name"             = "DataflowName";
        "objectId"         = "DataflowId";
        "generation" = "DataflowGeneration"
    }

    # Define renaming map for Fabric Dataflows (Gen 2 CICD)
    $fabricDataflowRenameMap = @{
        "id" = "DataflowId";
        "displayName" = "DataflowName";
        "description" = "DataflowDescription"
    }

    $dataflowRefreshRenameMap = @{
        "requestId" = "DataflowRefreshRequestId";
        "id" = "DataflowRefreshId";
        "startTime" = "DataflowRefreshStartTime";
        "endTime" = "DataflowRefreshEndTime";
        "status" = "DataflowRefreshStatus" ;
        "refreshType" = "DataflowRefreshType" ;
        "errorInfo" = "DataflowErrorInfo"
    }

    $dataflowLineageRenameMap = @{
        "datasetObjectId"   = "DatasetId";
        "dataflowObjectId"  = "DataflowId";
        "workspaceObjectId" = "WorkspaceId"
    }

    $reportRenameMap = @{
        "id" = "ReportId";
        "name" = "ReportName";
        "description" = "ReportDescription";
        "webUrl" = "ReportWebUrl";
        "embedUrl" = "ReportEmbedUrl";
        "isFromPbix" = "ReportIsFromPbix";
        "isOwnedByMe" = "ReportIsOwnedByMe";
        "datasetId" = "DatasetId";
        "datasetWorkspaceId" = "DatasetWorkspaceId";
        "reportType" = "ReportType"
    }

    $pageRenameMap = @{
        "name" = "PageName";
        "displayName" = "PageDisplayName";
        "order" = "PageOrder"
    }

    $appRenameMap = @{
        "id" = "AppId";
        "name" = "AppName";
        "lastUpdate" = "AppLastUpdate";
        "description" = "AppDescription";
        "publishedBy" = "AppPublishedBy";
        "workspaceId" = "AppWorkspaceId";
        "users" = "AppUsers"
    }

    $appReportRenameMap = @{
        "id" = "AppReportId";
        "reportType" = "AppReportType";
        "name" = "ReportName";
        "webUrl" = "AppReportWebUrl";
        "embedUrl" = "AppReportEmbedUrl";
        "isOwnedByMe" = "AppReportIsOwnedByMe";
        "datasetId" = "AppReportDatasetId";
        "originalReportObjectId" = "ReportId";
        "users" = "AppUsers";
        "subscriptions" = "AppReportSubscriptions";
        "sections" = "AppReportSections"
    }

    # Define renaming map for Fabric Items
    $fabricItemsRenameMap = @{
        "id" = "FabricItemID";
        "type" = "FabricItemType";
        "displayName" = "FabricItemName";
        "description" = "FabricItemDescription"
    }

    switch ($Name) {
        'Workspace' { return $workspaceRenameMap }
        'Dataset' { return $datasetRenameMap }
        'DatasetDatasource' { return $datasetDatasourceRenameMap }
        'DataflowDatasource' { return $dataflowDatasourceRenameMap }
        'DatasetRefresh' { return $datasetRefreshRenameMap }
        'DatasetRefreshSchedule' { return $datasetRefreshScheduleRenameMap }
        'Dataflow' { return $dataflowRenameMap }
        'FabricDataflow' { return $fabricDataflowRenameMap }
        'DataflowRefresh' { return $dataflowRefreshRenameMap }
        'DataflowLineage' { return $dataflowLineageRenameMap }
        'Report' { return $reportRenameMap }
        'Page' { return $pageRenameMap }
        'App' { return $appRenameMap }
        'AppReport' { return $appReportRenameMap }
        'FabricItems' { return $fabricItemsRenameMap }
    }
}

function Rename-Properties {
    <#
    .SYNOPSIS
        Copies an API object into a new PSObject using a rename map; unmapped properties are kept verbatim (monolith 637-652).
    #>
    [CmdletBinding()]
    param ($object, $renameMap)
    $newObject = New-Object PSObject
    foreach ($originalName in $renameMap.Keys) {
        $newPropertyName = $renameMap[$originalName]
        $propertyValue = if ($object.PSObject.Properties[$originalName]) { $object.$originalName } else { $null }
        if ($newObject.PSObject.Properties[$newPropertyName]) { $newPropertyName += "_duplicate" }
        $newObject | Add-Member -MemberType NoteProperty -Name $newPropertyName -Value $propertyValue
    }
    foreach ($property in $object.PSObject.Properties) {
        if (-not $renameMap.ContainsKey($property.Name)) {
            $newObject | Add-Member -MemberType NoteProperty -Name $property.Name -Value $property.Value
        }
    }
    return $newObject
}

# =====================================================================================================================
# Private helpers
# =====================================================================================================================

function Add-IQNote {
    <#
    .SYNOPSIS
        Adds or overwrites a NoteProperty on a row (Add-Member -Force shorthand used by every collector).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $false)][AllowNull()]$Value
    )
    $Row | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
}

function ConvertTo-IQFlatRow {
    <#
    .SYNOPSIS
        Serialises nested arrays/objects of a row to compact JSON strings so JSON files and Excel cells hold real text (audit C5-12).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Row)
    if ($null -eq $Row) { return $null }
    foreach ($p in @($Row.PSObject.Properties)) {
        $v = $p.Value
        if ($null -eq $v -or $v -is [string] -or $v -is [System.ValueType]) { continue }
        if ($v -is [System.Collections.IDictionary] -or $v -is [System.Management.Automation.PSCustomObject] -or $v -is [System.Collections.IEnumerable]) {
            $isEmptyCollection = ($v -is [System.Collections.ICollection] -and $v.Count -eq 0)
            if ($isEmptyCollection) { $p.Value = $null }
            else {
                try { $p.Value = ConvertTo-Json -InputObject $v -Depth 20 -Compress }
                catch { $p.Value = [string]$v }
            }
        }
    }
    return $Row
}

function Get-IQInventoryOption {
    <#
    .SYNOPSIS
        Reads an entry-point option from $script:IQ.Options ($null when absent).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Name)
    if (-not $script:IQ) { return $null }
    return (Get-IQMemberValue -Object $script:IQ.Options -Name $Name)
}

function ConvertTo-IQIdList {
    <#
    .SYNOPSIS
        Normalises an id/name parameter (array, comma/semicolon separated string, or $null) to a trimmed string array.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Value)
    $out = @()
    foreach ($v in @($Value)) {
        if ($null -eq $v) { continue }
        foreach ($part in ([string]$v -split '[,;]')) {
            $t = $part.Trim()
            if ($t.Length -gt 0) { $out += $t }
        }
    }
    return $out
}

function Get-IQInventoryCache {
    <#
    .SYNOPSIS
        Per-run cache of raw /reports and /datasets responses keyed by workspace id (audit C4-10: scope scans are reused by the collectors).
    #>
    [CmdletBinding()]
    param()
    if (-not $script:IQ.ContainsKey('InventoryCache') -or $null -eq $script:IQ.InventoryCache) {
        $script:IQ.InventoryCache = @{ Reports = @{}; Datasets = @{}; MyWorkspaceReports = $null; MyWorkspaceDatasets = $null }
    }
    return $script:IQ.InventoryCache
}

function Invoke-IQInventoryGet {
    <#
    .SYNOPSIS
        GET through Invoke-IQApi that never throws: returns the parsed response or $null (Warn, or one Debug line with -Optional).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][ValidateSet('PowerBI', 'Fabric')][string]$Api = 'PowerBI',
        [Parameter(Mandatory = $false)][hashtable]$Query,
        [Parameter(Mandatory = $false)][switch]$Optional,
        [Parameter(Mandatory = $false)][string]$Description,
        [Parameter(Mandatory = $false)][string]$Item,
        [Parameter(Mandatory = $false)][AllowNull()][System.Collections.Generic.List[object]]$ErrorList
    )
    $desc = $Description
    if ([string]::IsNullOrEmpty($desc)) { $desc = "GET $Path" }
    $params = @{ Method = 'GET'; Path = $Path; Api = $Api; Stage = 'Inventory' }
    if ($Query) { $params.Query = $Query }
    if ($Optional) { $params.AllowNotFound = $true }
    try {
        return (Invoke-IQApi @params)
    }
    catch {
        $level = 'Warn'
        if ($Optional) { $level = 'Debug' }
        Write-IQLog -Level $level -Stage Inventory -Item $Item -Message ("{0} failed: {1}" -f $desc, $_.Exception.Message)
        if ($null -ne $ErrorList -and -not $Optional) {
            $ErrorList.Add([pscustomobject]@{ Collector = $desc; Path = $Path; Message = $_.Exception.Message })
        }
        return $null
    }
}

function Get-IQInventoryList {
    <#
    .SYNOPSIS
        Like Invoke-IQInventoryGet but returns the response's 'value' array (empty array on any failure).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][ValidateSet('PowerBI', 'Fabric')][string]$Api = 'PowerBI',
        [Parameter(Mandatory = $false)][hashtable]$Query,
        [Parameter(Mandatory = $false)][switch]$Optional,
        [Parameter(Mandatory = $false)][string]$Description,
        [Parameter(Mandatory = $false)][string]$Item,
        [Parameter(Mandatory = $false)][AllowNull()][System.Collections.Generic.List[object]]$ErrorList
    )
    $response = Invoke-IQInventoryGet @PSBoundParameters
    if ($null -eq $response) { return @() }
    if ($response -is [System.Management.Automation.PSCustomObject] -and $response.PSObject.Properties['value']) {
        if ($null -eq $response.value) { return @() }
        return @($response.value | Where-Object { $null -ne $_ })
    }
    if ($response -is [array]) { return @($response) }
    return @()
}

function New-IQPseudoWorkspace {
    <#
    .SYNOPSIS
        Builds the synthetic "My Workspace" / "Shared Reports (No Workspace Access)" workspace row (monolith 2368-2375, 2452-2460) flagged WorkspaceIsSynthetic (audit C6-02).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Name)
    return [PSCustomObject]@{
        WorkspaceId                    = $Name
        WorkspaceName                  = $Name
        WorkspaceType                  = "Workspace"
        WorkspaceIsReadOnly            = $false
        WorkspaceIsOnDedicatedCapacity = $false
        WorkspaceIsSynthetic           = $true
        WorkspaceApiScope              = 'myorg'
    }
}

function ConvertTo-IQPickerWorkspace {
    <#
    .SYNOPSIS
        Copy of a renamed workspace row that also carries the raw id/name/type fields the interactive pickers expect.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Workspace)
    return [pscustomobject]@{
        id                             = $Workspace.WorkspaceId
        name                           = $Workspace.WorkspaceName
        type                           = $Workspace.WorkspaceType
        isReadOnly                     = $Workspace.WorkspaceIsReadOnly
        isOnDedicatedCapacity          = $Workspace.WorkspaceIsOnDedicatedCapacity
        capacityId                     = $Workspace.WorkspaceCapacityId
        WorkspaceId                    = $Workspace.WorkspaceId
        WorkspaceName                  = $Workspace.WorkspaceName
        WorkspaceType                  = $Workspace.WorkspaceType
        WorkspaceIsOnDedicatedCapacity = $Workspace.WorkspaceIsOnDedicatedCapacity
    }
}

# =====================================================================================================================
# Workspace listing (monolith 1524-1527, paged per audit C4-11)
# =====================================================================================================================

function Get-IQWorkspaceList {
    <#
    .SYNOPSIS
        Lists every accessible workspace via GET groups ($top=5000, $skip paging) renamed with the workspace map; throws when the listing itself fails.
    .DESCRIPTION
        Rows carry WorkspaceId, WorkspaceName, WorkspaceType, WorkspaceIsReadOnly, WorkspaceIsOnDedicatedCapacity,
        WorkspaceCapacityId, WorkspaceDefaultDatasetStorageFormat (map), every other API field verbatim, plus
        WorkspaceState / WorkspaceHasWorkspaceLevelSettings when the service returns them and WorkspaceIsSynthetic = $false.
    #>
    [CmdletBinding()]
    param()
    $top = 5000
    $skip = 0
    $raw = New-Object System.Collections.Generic.List[object]
    do {
        $page = Invoke-IQApi -Method GET -Path 'groups' -Query @{ '$top' = $top; '$skip' = $skip } -Stage Inventory
        $rows = @()
        if ($null -ne $page -and $page.PSObject.Properties['value'] -and $null -ne $page.value) { $rows = @($page.value) }
        if ($rows.Count -gt 0) { $raw.AddRange([object[]]$rows) }
        $skip += $top
    } while ($rows.Count -eq $top -and $skip -lt 200000)

    if ($raw.Count -eq 0) {
        Write-IQLog -Level Warn -Stage Inventory -Message 'No workspaces were returned by GET groups (check sign-in and permissions).'
    }
    $map = Get-IQRenameMap -Name Workspace
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($w in $raw) {
        $row = Rename-Properties -object $w -renameMap $map
        if ($w.PSObject.Properties['state']) { Add-IQNote -Row $row -Name 'WorkspaceState' -Value $w.state }
        if ($w.PSObject.Properties['hasWorkspaceLevelSettings']) { Add-IQNote -Row $row -Name 'WorkspaceHasWorkspaceLevelSettings' -Value $w.hasWorkspaceLevelSettings }
        Add-IQNote -Row $row -Name 'WorkspaceIsSynthetic' -Value $false
        $out.Add((ConvertTo-IQFlatRow -Row $row))
    }
    Write-IQLog -Level Info -Stage Inventory -Message ("Found {0} accessible workspace(s)." -f $out.Count)
    return $out.ToArray()
}

# =====================================================================================================================
# Scope resolution (monolith 1530-1860)
# =====================================================================================================================

function Select-IQScopeWorkspace {
    <#
    .SYNOPSIS
        Filters renamed workspace rows by exact ids and/or wildcard names (-like); -All returns every row. Reports what was not found.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Workspaces,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$Ids,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$Names,
        [Parameter(Mandatory = $false)][switch]$All
    )
    $Ids = @(ConvertTo-IQIdList -Value $Ids)
    $Names = @(ConvertTo-IQIdList -Value $Names)
    $selected = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    $missingIds = @()
    $missingNames = @()
    if ($All) {
        foreach ($w in $Workspaces) { $selected.Add($w) }
    }
    else {
        foreach ($id in $Ids) {
            $hit = $false
            foreach ($w in $Workspaces) {
                if ([string]$w.WorkspaceId -eq $id) {
                    $hit = $true
                    $k = ([string]$w.WorkspaceId).ToLowerInvariant()
                    if (-not $seen.ContainsKey($k)) { $seen[$k] = $true; $selected.Add($w) }
                }
            }
            if (-not $hit) { $missingIds += $id }
        }
        foreach ($pattern in $Names) {
            $hit = $false
            foreach ($w in $Workspaces) {
                if ([string]$w.WorkspaceName -like $pattern) {
                    $hit = $true
                    $k = ([string]$w.WorkspaceId).ToLowerInvariant()
                    if (-not $seen.ContainsKey($k)) { $seen[$k] = $true; $selected.Add($w) }
                }
            }
            if (-not $hit) { $missingNames += $pattern }
        }
    }
    return @{ Selected = $selected.ToArray(); MissingIds = $missingIds; MissingNames = $missingNames }
}

function Get-IQScopeReportScan {
    <#
    .SYNOPSIS
        Lists the reports of the given workspaces (and optionally My Workspace) as picker-shaped rows (monolith 1590-1608 / 1766-1784); responses are cached for the collectors.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Workspaces,
        [Parameter(Mandatory = $false)][switch]$IncludeMyWorkspace
    )
    $cache = Get-IQInventoryCache
    $out = New-Object System.Collections.Generic.List[object]
    $total = @($Workspaces).Count
    $n = 0
    foreach ($ws in $Workspaces) {
        $n++
        $wsId = [string]$ws.WorkspaceId
        Write-IQLog -Level Debug -Stage Inventory -Message ("Scanning reports {0}/{1}: {2}" -f $n, $total, $ws.WorkspaceName)
        $reports = $null
        if ($cache.Reports.ContainsKey($wsId)) { $reports = $cache.Reports[$wsId] }
        else {
            $reports = @(Get-IQInventoryList -Path "groups/$wsId/reports" -Description "reports of workspace '$($ws.WorkspaceName)'" -Item $ws.WorkspaceName)
            $cache.Reports[$wsId] = $reports
        }
        foreach ($rpt in $reports) {
            $out.Add([pscustomobject]@{
                    ReportId           = $rpt.id
                    ReportName         = $rpt.name
                    WorkspaceId        = $wsId
                    WorkspaceName      = $ws.WorkspaceName
                    DatasetId          = $rpt.datasetId
                    DatasetWorkspaceId = $rpt.datasetWorkspaceId
                    IsMyWorkspace      = $false
                })
        }
    }
    if ($IncludeMyWorkspace) {
        # Audit C4-07: reports that live in My Workspace can be targeted too (skip app copies and shared reports).
        $mine = $cache.MyWorkspaceReports
        if ($null -eq $mine) {
            $mine = @(Get-IQInventoryList -Path 'reports' -Description 'My Workspace reports' -Item 'My Workspace')
            $cache.MyWorkspaceReports = $mine
        }
        foreach ($rpt in $mine) {
            if ($rpt.PSObject.Properties['appId'] -and $rpt.appId) { continue }
            if ($rpt.PSObject.Properties['isOwnedByMe'] -and $rpt.isOwnedByMe -eq $false) { continue }
            $out.Add([pscustomobject]@{
                    ReportId           = $rpt.id
                    ReportName         = $rpt.name
                    WorkspaceId        = 'My Workspace'
                    WorkspaceName      = 'My Workspace'
                    DatasetId          = $rpt.datasetId
                    DatasetWorkspaceId = $rpt.datasetWorkspaceId
                    IsMyWorkspace      = $true
                })
        }
    }
    return $out.ToArray()
}

function Get-IQScopeDatasetScan {
    <#
    .SYNOPSIS
        Lists the datasets of the given workspaces (and optionally My Workspace) as picker-shaped rows (monolith 1711-1731); responses are cached for the collectors.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Workspaces,
        [Parameter(Mandatory = $false)][switch]$IncludeMyWorkspace
    )
    $cache = Get-IQInventoryCache
    $out = New-Object System.Collections.Generic.List[object]
    $total = @($Workspaces).Count
    $n = 0
    foreach ($ws in $Workspaces) {
        $n++
        $wsId = [string]$ws.WorkspaceId
        Write-IQLog -Level Debug -Stage Inventory -Message ("Scanning datasets {0}/{1}: {2}" -f $n, $total, $ws.WorkspaceName)
        $datasets = $null
        if ($cache.Datasets.ContainsKey($wsId)) { $datasets = $cache.Datasets[$wsId] }
        else {
            $datasets = @(Get-IQInventoryList -Path "groups/$wsId/datasets" -Description "datasets of workspace '$($ws.WorkspaceName)'" -Item $ws.WorkspaceName)
            $cache.Datasets[$wsId] = $datasets
        }
        foreach ($ds in $datasets) {
            $out.Add([pscustomobject]@{ DatasetId = $ds.id; DatasetName = $ds.name; WorkspaceId = $wsId; WorkspaceName = $ws.WorkspaceName; IsMyWorkspace = $false })
        }
    }
    if ($IncludeMyWorkspace) {
        $mine = $cache.MyWorkspaceDatasets
        if ($null -eq $mine) {
            $mine = @(Get-IQInventoryList -Path 'datasets' -Description 'My Workspace datasets' -Item 'My Workspace')
            $cache.MyWorkspaceDatasets = $mine
        }
        foreach ($ds in $mine) {
            $out.Add([pscustomobject]@{ DatasetId = $ds.id; DatasetName = $ds.name; WorkspaceId = 'My Workspace'; WorkspaceName = 'My Workspace'; IsMyWorkspace = $true })
        }
    }
    return $out.ToArray()
}

function New-IQScopeObject {
    <#
    .SYNOPSIS
        Builds the scope hashtable ($script:IQ.Scope) incl. lower-cased lookup sets for the Reports/Models filters.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RunMode,
        [Parameter(Mandatory = $false)][AllowEmptyCollection()][array]$Workspaces = @(),
        [Parameter(Mandatory = $false)][bool]$IncludeMyWorkspace = $false,
        [Parameter(Mandatory = $false)][AllowEmptyCollection()][string[]]$ReportIds = @(),
        [Parameter(Mandatory = $false)][AllowEmptyCollection()][string[]]$DatasetIds = @(),
        [Parameter(Mandatory = $false)][AllowEmptyCollection()][string[]]$InaccessibleWorkspaceIds = @(),
        [Parameter(Mandatory = $false)][string]$Source = 'Parameters'
    )
    $reportSet = @{}
    foreach ($r in @($ReportIds)) { if ($r) { $reportSet[([string]$r).ToLowerInvariant()] = $true } }
    $datasetSet = @{}
    foreach ($d in @($DatasetIds)) { if ($d) { $datasetSet[([string]$d).ToLowerInvariant()] = $true } }
    $wsIds = @()
    $wsNames = @()
    foreach ($w in @($Workspaces)) { $wsIds += [string]$w.WorkspaceId; $wsNames += [string]$w.WorkspaceName }
    return @{
        RunMode                  = $RunMode
        Workspaces               = @($Workspaces)
        WorkspaceIds             = @($wsIds)
        WorkspaceNames           = @($wsNames)
        IncludeMyWorkspace       = [bool]$IncludeMyWorkspace
        ReportIds                = @($ReportIds)
        DatasetIds               = @($DatasetIds)
        ReportIdSet              = $reportSet
        DatasetIdSet             = $datasetSet
        InaccessibleWorkspaceIds = @($InaccessibleWorkspaceIds)
        Source                   = $Source
        ResolvedUtc              = [DateTime]::UtcNow.ToString('o')
    }
}

function Get-IQScopeFromManifest {
    <#
    .SYNOPSIS
        Rebuilds a scope hashtable from manifest.scope (ids only; Workspaces are filled by the caller) or $null when nothing is persisted.
    #>
    [CmdletBinding()]
    param()
    if (-not $script:IQ -or $null -eq $script:IQ.Manifest) { return $null }
    $s = Get-IQMemberValue -Object $script:IQ.Manifest -Name 'scope'
    if ($null -eq $s) { return $null }
    $runMode = [string](Get-IQMemberValue -Object $s -Name 'runMode')
    if ([string]::IsNullOrEmpty($runMode)) { $runMode = 'Workspaces' }
    $wsIds = @(ConvertTo-IQIdList -Value (Get-IQMemberValue -Object $s -Name 'workspaceIds'))
    $reportIds = @(ConvertTo-IQIdList -Value (Get-IQMemberValue -Object $s -Name 'reportIds'))
    $datasetIds = @(ConvertTo-IQIdList -Value (Get-IQMemberValue -Object $s -Name 'datasetIds'))
    $includeMy = [bool](Get-IQMemberValue -Object $s -Name 'includeMyWorkspace')
    $inaccessible = @(ConvertTo-IQIdList -Value (Get-IQMemberValue -Object $s -Name 'inaccessibleWorkspaceIds'))
    $scope = New-IQScopeObject -RunMode $runMode -IncludeMyWorkspace $includeMy -ReportIds $reportIds -DatasetIds $datasetIds -InaccessibleWorkspaceIds $inaccessible -Source 'Manifest'
    $scope.WorkspaceIds = $wsIds
    $scope.WorkspaceNames = @(ConvertTo-IQIdList -Value (Get-IQMemberValue -Object $s -Name 'workspaceNames'))
    return $scope
}

function Save-IQScope {
    <#
    .SYNOPSIS
        Stores the scope in $script:IQ.Scope and manifest.scope (runMode, workspaceIds, reportIds, datasetIds, includeMyWorkspace + additive fields), then saves the manifest.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Scope)
    $script:IQ.Scope = $Scope
    if ($null -ne $script:IQ.Manifest) {
        $script:IQ.Manifest['scope'] = [ordered]@{
            runMode                  = $Scope.RunMode
            workspaceIds             = @($Scope.WorkspaceIds)
            reportIds                = @($Scope.ReportIds)
            datasetIds               = @($Scope.DatasetIds)
            includeMyWorkspace       = [bool]$Scope.IncludeMyWorkspace
            workspaceNames           = @($Scope.WorkspaceNames)
            inaccessibleWorkspaceIds = @($Scope.InaccessibleWorkspaceIds)
            source                   = $Scope.Source
            resolvedUtc              = $Scope.ResolvedUtc
        }
        Save-IQManifest
    }
}

function Write-IQScopeSummary {
    <#
    .SYNOPSIS
        Logs the resolved scope (mode, workspaces, reports, models, My Workspace) so the log shows what actually ran (audit C4-08).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Scope)
    $my = 'no'
    if ($Scope.IncludeMyWorkspace) { $my = 'yes' }
    Write-IQLog -Level Info -Stage Inventory -Message ("Scope ({0}): mode={1}, {2} workspace(s), {3} report id(s), {4} model id(s), My Workspace={5}" -f `
            $Scope.Source, $Scope.RunMode, @($Scope.WorkspaceIds).Count, @($Scope.ReportIds).Count, @($Scope.DatasetIds).Count, $my)
    foreach ($w in @($Scope.Workspaces)) {
        Write-IQLog -Level Debug -Stage Inventory -Message ("  workspace: {0} ({1})" -f $w.WorkspaceName, $w.WorkspaceId)
    }
    foreach ($id in @($Scope.InaccessibleWorkspaceIds)) {
        Write-IQLog -Level Warn -Stage Inventory -Message "Workspace $id is required by the selection (remote model or connected report) but is not accessible - its content will be skipped."
    }
}

function Resolve-IQScope {
    <#
    .SYNOPSIS
        Resolves the run scope (brief 5.3): headless parameters, else the interactive dialog flow, else "no scope"; result in $script:IQ.Scope and manifest.scope.
    .DESCRIPTION
        Parameters default to $script:IQ.Options (RunMode, WorkspaceId, WorkspaceName, AllWorkspaces, IncludeMyWorkspace,
        ReportId, DatasetId). Workspaces mode selects by id/name/-AllWorkspaces. Reports mode scans the candidate workspaces
        for the given report ids and adds each report's datasetWorkspaceId workspace (monolith 1631-1658). Models mode
        locates the datasets, then scans every workspace's reports for datasetId in the selection (monolith 1753-1804).
        On a resumed run the persisted manifest.scope is reused (CLI scope parameters are ignored with a Warn when they
        differ). Returns $null when the interactive user cancels.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][ValidateSet('', 'Workspaces', 'Reports', 'Models')][string]$RunMode,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$WorkspaceId,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$WorkspaceName,
        [Parameter(Mandatory = $false)][switch]$AllWorkspaces,
        [Parameter(Mandatory = $false)][switch]$IncludeMyWorkspace,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$ReportId,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$DatasetId,
        [Parameter(Mandatory = $false)][AllowNull()][array]$Workspaces,
        [Parameter(Mandatory = $false)][switch]$NoInteractive,
        [Parameter(Mandatory = $false)][switch]$Force
    )
    if (-not $script:IQ) { throw 'ImpactIQ context is not initialised (call Initialize-IQContext first).' }

    # ---- defaults from the entry-point options ----------------------------------------------------------------
    if (-not $PSBoundParameters.ContainsKey('RunMode') -or [string]::IsNullOrEmpty($RunMode)) {
        $RunMode = [string](Get-IQInventoryOption -Name 'RunMode')
        if ([string]::IsNullOrEmpty($RunMode)) { $RunMode = 'Workspaces' }
    }
    if (-not $PSBoundParameters.ContainsKey('WorkspaceId')) { $WorkspaceId = @(ConvertTo-IQIdList -Value (Get-IQInventoryOption -Name 'WorkspaceId')) }
    if (-not $PSBoundParameters.ContainsKey('WorkspaceName')) { $WorkspaceName = @(ConvertTo-IQIdList -Value (Get-IQInventoryOption -Name 'WorkspaceName')) }
    if (-not $PSBoundParameters.ContainsKey('ReportId')) { $ReportId = @(ConvertTo-IQIdList -Value (Get-IQInventoryOption -Name 'ReportId')) }
    if (-not $PSBoundParameters.ContainsKey('DatasetId')) { $DatasetId = @(ConvertTo-IQIdList -Value (Get-IQInventoryOption -Name 'DatasetId')) }
    if (-not $PSBoundParameters.ContainsKey('AllWorkspaces')) { $AllWorkspaces = [bool](Get-IQInventoryOption -Name 'AllWorkspaces') }
    if (-not $PSBoundParameters.ContainsKey('IncludeMyWorkspace')) { $IncludeMyWorkspace = [bool](Get-IQInventoryOption -Name 'IncludeMyWorkspace') }
    $WorkspaceId = @(ConvertTo-IQIdList -Value $WorkspaceId)
    $WorkspaceName = @(ConvertTo-IQIdList -Value $WorkspaceName)
    $ReportId = @(ConvertTo-IQIdList -Value $ReportId)
    $DatasetId = @(ConvertTo-IQIdList -Value $DatasetId)
    $haveScopeParameters = ($WorkspaceId.Count -gt 0 -or $WorkspaceName.Count -gt 0 -or [bool]$AllWorkspaces -or $ReportId.Count -gt 0 -or $DatasetId.Count -gt 0 -or [bool]$IncludeMyWorkspace)

    # ---- resume: reuse the persisted scope ------------------------------------------------------------------------
    if (-not $Force -and $script:IQ.IsResume) {
        $persisted = Get-IQScopeFromManifest
        if ($null -ne $persisted -and (@($persisted.WorkspaceIds).Count -gt 0 -or $persisted.IncludeMyWorkspace)) {
            if ($haveScopeParameters) {
                $sameMode = ($persisted.RunMode -eq $RunMode)
                $sameIds = $true
                if ($WorkspaceId.Count -gt 0) { foreach ($id in $WorkspaceId) { if (@($persisted.WorkspaceIds) -notcontains $id) { $sameIds = $false } } }
                if ($ReportId.Count -gt 0) { foreach ($id in $ReportId) { if (@($persisted.ReportIds) -notcontains $id) { $sameIds = $false } } }
                if ($DatasetId.Count -gt 0) { foreach ($id in $DatasetId) { if (@($persisted.DatasetIds) -notcontains $id) { $sameIds = $false } } }
                if (-not $sameMode -or -not $sameIds) {
                    Write-IQLog -Level Warn -Stage Inventory -Message ("Resuming run '{0}': the persisted scope is reused and the scope parameters given now are ignored (use -Force for a fresh run)." -f $script:IQ.RunId)
                }
            }
            $rows = @()
            try { $rows = @(Get-IQInventory -Name 'workspaces') } catch { $rows = @() }
            $real = @($rows | Where-Object { $null -ne $_ -and -not [bool](Get-IQMemberValue -Object $_ -Name 'WorkspaceIsSynthetic') })
            if ($real.Count -eq 0 -and @($persisted.WorkspaceIds).Count -gt 0) {
                if ($null -eq $Workspaces) { $Workspaces = @(Get-IQWorkspaceList) }
                $real = @((Select-IQScopeWorkspace -Workspaces $Workspaces -Ids $persisted.WorkspaceIds).Selected)
            }
            $persisted.Workspaces = @($real)
            if ($real.Count -gt 0) {
                $persisted.WorkspaceIds = @($real | ForEach-Object { [string]$_.WorkspaceId })
                $persisted.WorkspaceNames = @($real | ForEach-Object { [string]$_.WorkspaceName })
            }
            $script:IQ.Scope = $persisted
            Write-IQScopeSummary -Scope $persisted
            return $persisted
        }
    }

    # ---- list workspaces (throws when the listing fails: no scope can be resolved without it, audit C4-03) ----
    if ($null -eq $Workspaces) { $Workspaces = @(Get-IQWorkspaceList) }
    $Workspaces = @($Workspaces)

    $source = 'Parameters'
    if (-not $haveScopeParameters) {
        $canPrompt = ([bool]$script:IQ.Interactive -and -not $NoInteractive -and (Get-Command -Name Select-IQScopeInteractive -ErrorAction SilentlyContinue))
        if (-not $canPrompt) {
            if ($RunMode -eq 'Reports') { throw "No scope: Reports mode requires -ReportId (one or more report ids); nothing is scanned by accident in a headless run." }
            if ($RunMode -eq 'Models') { throw "No scope: Models mode requires -DatasetId (one or more semantic model ids); nothing is scanned by accident in a headless run." }
            throw "No scope: specify -WorkspaceId / -WorkspaceName / -AllWorkspaces (Workspaces mode), -ReportId (Reports mode) or -DatasetId (Models mode), or -IncludeMyWorkspace. Nothing is scanned by accident in a headless run."
        }
        Write-IQLog -Level Info -Stage Inventory -Message 'No scope parameters given - opening the interactive selection dialogs.'
        $pickerRows = @($Workspaces | ForEach-Object { ConvertTo-IQPickerWorkspace -Workspace $_ })
        $selection = Select-IQScopeInteractive -Workspaces $pickerRows
        if ($null -eq $selection) {
            Write-IQLog -Level Warn -Stage Inventory -Message 'Run cancelled by user (scope selection).'
            $script:IQ.ScopeCancelled = $true
            return $null
        }
        $source = 'Interactive'
        $selMode = [string](Get-IQMemberValue -Object $selection -Name 'RunMode')
        if ($selMode) { $RunMode = $selMode }
        $WorkspaceId = @(ConvertTo-IQIdList -Value (Get-IQMemberValue -Object $selection -Name 'WorkspaceIds'))
        $WorkspaceName = @()
        $ReportId = @(ConvertTo-IQIdList -Value (Get-IQMemberValue -Object $selection -Name 'ReportIds'))
        $DatasetId = @(ConvertTo-IQIdList -Value (Get-IQMemberValue -Object $selection -Name 'DatasetIds'))
        $IncludeMyWorkspace = [bool](Get-IQMemberValue -Object $selection -Name 'IncludeMyWorkspace')
        if ([bool](Get-IQMemberValue -Object $selection -Name 'TimedOut')) {
            Write-IQLog -Level Warn -Stage Inventory -Message 'Interactive selection timed out - using the defaulted selection (legacy behaviour: all workspaces + My Workspace).'
        }
        $AllWorkspaces = ($RunMode -eq 'Workspaces' -and $WorkspaceId.Count -eq 0 -and -not $IncludeMyWorkspace)
    }

    $scope = $null
    switch ($RunMode) {
        'Reports' {
            if ($ReportId.Count -eq 0) { throw "Reports mode requires -ReportId (one or more report ids)." }
            $candidates = $Workspaces
            if ($WorkspaceId.Count -gt 0 -or $WorkspaceName.Count -gt 0) {
                $sel = Select-IQScopeWorkspace -Workspaces $Workspaces -Ids $WorkspaceId -Names $WorkspaceName
                foreach ($m in @($sel.MissingIds) + @($sel.MissingNames)) { Write-IQLog -Level Warn -Stage Inventory -Message "Workspace '$m' was not found among the accessible workspaces." }
                $candidates = @($sel.Selected)
                if ($candidates.Count -eq 0) {
                    Write-IQLog -Level Warn -Stage Inventory -Message 'None of the requested workspaces were found - scanning ALL accessible workspaces for the requested reports.'
                    $candidates = $Workspaces
                }
            }
            Write-IQLog -Level Info -Stage Inventory -Message ("Reports mode: scanning {0} workspace(s) for {1} report id(s)..." -f $candidates.Count, $ReportId.Count)
            $scan = @(Get-IQScopeReportScan -Workspaces $candidates -IncludeMyWorkspace:$IncludeMyWorkspace)
            $wanted = @{}
            foreach ($r in $ReportId) { $wanted[$r.ToLowerInvariant()] = $true }
            $selectedReports = @($scan | Where-Object { $_.ReportId -and $wanted.ContainsKey(([string]$_.ReportId).ToLowerInvariant()) })
            $foundIds = @{}
            foreach ($r in $selectedReports) { $foundIds[([string]$r.ReportId).ToLowerInvariant()] = $true }
            foreach ($r in $ReportId) { if (-not $foundIds.ContainsKey($r.ToLowerInvariant())) { Write-IQLog -Level Warn -Stage Inventory -Message "Report id '$r' was not found in the scanned workspaces." } }
            if ($selectedReports.Count -eq 0) { throw "None of the requested report ids were found in the scanned workspace(s)." }

            $needed = @()
            $datasetIds = @()
            $fromMy = $false
            foreach ($rpt in $selectedReports) {
                if ($rpt.IsMyWorkspace) { $fromMy = $true } elseif ($rpt.WorkspaceId) { $needed += [string]$rpt.WorkspaceId }
                if ($rpt.DatasetId) { $datasetIds += [string]$rpt.DatasetId }
                if ($rpt.DatasetWorkspaceId -and $rpt.DatasetWorkspaceId -ne $rpt.WorkspaceId) {
                    $needed += [string]$rpt.DatasetWorkspaceId
                    Write-IQLog -Level Info -Stage Inventory -Message "Report '$($rpt.ReportName)' uses a remote model in workspace $($rpt.DatasetWorkspaceId) - including that workspace."
                }
            }
            $needed = @($needed | Select-Object -Unique)
            $datasetIds = @($datasetIds | Select-Object -Unique)
            $accessible = @($Workspaces | Where-Object { $needed -contains [string]$_.WorkspaceId })
            $accessibleIds = @($accessible | ForEach-Object { [string]$_.WorkspaceId })
            $inaccessible = @($needed | Where-Object { $accessibleIds -notcontains $_ })   # audit C4-05
            $reportIdsOut = @($selectedReports | ForEach-Object { [string]$_.ReportId } | Select-Object -Unique)
            $scope = New-IQScopeObject -RunMode 'Reports' -Workspaces $accessible -IncludeMyWorkspace ([bool]$IncludeMyWorkspace -or $fromMy) -ReportIds $reportIdsOut -DatasetIds $datasetIds -InaccessibleWorkspaceIds $inaccessible -Source $source
            if ($accessible.Count -eq 0 -and -not $scope.IncludeMyWorkspace) { throw "No accessible workspaces found for the selected reports." }
        }
        'Models' {
            if ($DatasetId.Count -eq 0) { throw "Models mode requires -DatasetId (one or more semantic model ids)." }
            $candidates = $Workspaces
            if ($WorkspaceId.Count -gt 0 -or $WorkspaceName.Count -gt 0) {
                $sel = Select-IQScopeWorkspace -Workspaces $Workspaces -Ids $WorkspaceId -Names $WorkspaceName
                foreach ($m in @($sel.MissingIds) + @($sel.MissingNames)) { Write-IQLog -Level Warn -Stage Inventory -Message "Workspace '$m' was not found among the accessible workspaces." }
                $candidates = @($sel.Selected)
                if ($candidates.Count -eq 0) {
                    Write-IQLog -Level Warn -Stage Inventory -Message 'None of the requested workspaces were found - scanning ALL accessible workspaces for the requested models.'
                    $candidates = $Workspaces
                }
            }
            Write-IQLog -Level Info -Stage Inventory -Message ("Models mode: locating {0} model id(s) in {1} workspace(s)..." -f $DatasetId.Count, $candidates.Count)
            $wanted = @{}
            foreach ($d in $DatasetId) { $wanted[$d.ToLowerInvariant()] = $true }
            $dsScan = @(Get-IQScopeDatasetScan -Workspaces $candidates -IncludeMyWorkspace:$IncludeMyWorkspace)
            $selectedModels = @($dsScan | Where-Object { $_.DatasetId -and $wanted.ContainsKey(([string]$_.DatasetId).ToLowerInvariant()) })
            $foundIds = @{}
            foreach ($m in $selectedModels) { $foundIds[([string]$m.DatasetId).ToLowerInvariant()] = $true }
            foreach ($d in $DatasetId) { if (-not $foundIds.ContainsKey($d.ToLowerInvariant())) { Write-IQLog -Level Warn -Stage Inventory -Message "Model id '$d' was not found in the scanned workspaces." } }
            if ($selectedModels.Count -eq 0) { throw "None of the requested model ids were found in the scanned workspace(s)." }

            $needed = @()
            $fromMy = $false
            foreach ($mdl in $selectedModels) { if ($mdl.IsMyWorkspace) { $fromMy = $true } else { $needed += [string]$mdl.WorkspaceId } }

            # Now fetch reports from ALL workspaces to find which ones use the selected models (monolith 1759-1784)
            Write-IQLog -Level Info -Stage Inventory -Message 'Scanning all workspaces for reports connected to the selected model(s)...'
            $rptScan = @(Get-IQScopeReportScan -Workspaces $Workspaces -IncludeMyWorkspace:$IncludeMyWorkspace)
            $reportIdsOut = @()
            foreach ($rpt in $rptScan) {
                if ($rpt.DatasetId -and $wanted.ContainsKey(([string]$rpt.DatasetId).ToLowerInvariant())) {
                    $reportIdsOut += [string]$rpt.ReportId
                    if ($rpt.IsMyWorkspace) { $fromMy = $true } else { $needed += [string]$rpt.WorkspaceId }
                }
            }
            $needed = @($needed | Select-Object -Unique)
            $reportIdsOut = @($reportIdsOut | Select-Object -Unique)
            Write-IQLog -Level Info -Stage Inventory -Message ("Found {0} report(s) connected to {1} selected model(s)." -f $reportIdsOut.Count, $selectedModels.Count)
            $accessible = @($Workspaces | Where-Object { $needed -contains [string]$_.WorkspaceId })
            $modelIdsOut = @($selectedModels | ForEach-Object { [string]$_.DatasetId } | Select-Object -Unique)
            $scope = New-IQScopeObject -RunMode 'Models' -Workspaces $accessible -IncludeMyWorkspace ([bool]$IncludeMyWorkspace -or $fromMy) -ReportIds $reportIdsOut -DatasetIds $modelIdsOut -Source $source
            if ($accessible.Count -eq 0 -and -not $scope.IncludeMyWorkspace) { throw "No accessible workspaces found for the selected models." }
        }
        default {
            $sel = Select-IQScopeWorkspace -Workspaces $Workspaces -Ids $WorkspaceId -Names $WorkspaceName -All:$AllWorkspaces
            foreach ($m in @($sel.MissingIds) + @($sel.MissingNames)) { Write-IQLog -Level Warn -Stage Inventory -Message "Workspace '$m' was not found among the accessible workspaces." }
            $selected = @($sel.Selected)
            if ($selected.Count -eq 0 -and -not $IncludeMyWorkspace) {
                throw "No scope: none of the requested workspaces were found (and -IncludeMyWorkspace was not given)."
            }
            $scope = New-IQScopeObject -RunMode 'Workspaces' -Workspaces $selected -IncludeMyWorkspace ([bool]$IncludeMyWorkspace) -Source $source
        }
    }

    Save-IQScope -Scope $scope
    Write-IQScopeSummary -Scope $scope
    return $scope
}

function Get-IQEffectiveScope {
    <#
    .SYNOPSIS
        $script:IQ.Scope, or the scope rebuilt from the manifest (later stages / -Stages Assemble runs); $null when none.
    #>
    [CmdletBinding()]
    param()
    if (-not $script:IQ) { return $null }
    if ($null -ne $script:IQ.Scope) { return $script:IQ.Scope }
    $s = Get-IQScopeFromManifest
    if ($null -ne $s) { $script:IQ.Scope = $s }
    return $s
}

function Test-IQScopeIncludesDataset {
    <#
    .SYNOPSIS
        $true when the dataset id is inside the scope (always in Workspaces mode; selected ids only in Reports/Models modes - monolith 2016-2035).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Scope,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$DatasetId
    )
    if ($null -eq $Scope -or $Scope.RunMode -eq 'Workspaces') { return $true }
    if ([string]::IsNullOrEmpty($DatasetId)) { return $false }
    return [bool]$Scope.DatasetIdSet.ContainsKey($DatasetId.ToLowerInvariant())
}

function Test-IQScopeIncludesReport {
    <#
    .SYNOPSIS
        $true when the report id is inside the scope (always in Workspaces mode; selected ids only in Reports/Models modes).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Scope,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$ReportId
    )
    if ($null -eq $Scope -or $Scope.RunMode -eq 'Workspaces') { return $true }
    if ([string]::IsNullOrEmpty($ReportId)) { return $false }
    return [bool]$Scope.ReportIdSet.ContainsKey($ReportId.ToLowerInvariant())
}

# =====================================================================================================================
# Shared collectors used by both real workspaces and My Workspace
# =====================================================================================================================

function Add-IQDatasetRefreshScheduleRow {
    <#
    .SYNOPSIS
        Import refresh schedule rows for one dataset (one row per day/time, monolith 2091-2125) with DatasetRefreshScheduleKind Import/Unavailable; DQ rows go to a separate list (audit C5-14).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$BasePath,          # 'groups/<ws>/datasets/<id>' or 'datasets/<id>'
        [Parameter(Mandatory = $true)]$DatasetRow,
        [Parameter(Mandatory = $true)][string]$WorkspaceId,
        [Parameter(Mandatory = $true)][string]$WorkspaceName,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$ScheduleList,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$DirectQueryList
    )
    $datasetId = [string]$DatasetRow.DatasetId
    $datasetName = [string]$DatasetRow.DatasetName
    $map = Get-IQRenameMap -Name DatasetRefreshSchedule
    $refreshScheduleResponse = Invoke-IQInventoryGet -Path "$BasePath/refreshSchedule" -Optional -Description "refresh schedule of dataset '$datasetName'" -Item $WorkspaceName
    if ($null -ne $refreshScheduleResponse) {
        # Get base properties that apply to all rows
        $renamedScheduleRecord = Rename-Properties -object $refreshScheduleResponse -renameMap $map
        # Create separate rows for each day-time combination
        $days = if ($refreshScheduleResponse.days) { $refreshScheduleResponse.days } else { @($null) }
        $times = if ($refreshScheduleResponse.times) { $refreshScheduleResponse.times } else { @($null) }
        foreach ($day in $days) {
            foreach ($time in $times) {
                $scheduleRow = $renamedScheduleRecord.PSObject.Copy()
                Add-IQNote -Row $scheduleRow -Name 'WorkspaceId' -Value $WorkspaceId
                Add-IQNote -Row $scheduleRow -Name 'WorkspaceName' -Value $WorkspaceName
                Add-IQNote -Row $scheduleRow -Name 'DatasetId' -Value $datasetId
                Add-IQNote -Row $scheduleRow -Name 'DatasetName' -Value $datasetName
                Add-IQNote -Row $scheduleRow -Name 'DatasetRefreshScheduleDay' -Value $day
                Add-IQNote -Row $scheduleRow -Name 'DatasetRefreshScheduleTime' -Value $time
                Add-IQNote -Row $scheduleRow -Name 'DatasetRefreshScheduleKind' -Value 'Import'
                $ScheduleList.Add((ConvertTo-IQFlatRow -Row $scheduleRow))
            }
        }
        return
    }

    # NEW: DirectQuery refresh schedule (only tried when the import schedule is unavailable)
    $dq = Invoke-IQInventoryGet -Path "$BasePath/directQueryRefreshSchedule" -Optional -Description "DirectQuery refresh schedule of dataset '$datasetName'" -Item $WorkspaceName
    if ($null -ne $dq) {
        $days = if ($dq.PSObject.Properties['days'] -and $dq.days) { @($dq.days) } else { @($null) }
        $times = if ($dq.PSObject.Properties['times'] -and $dq.times) { @($dq.times) } else { @($null) }
        $frequency = $null
        if ($dq.PSObject.Properties['frequency']) { $frequency = $dq.frequency }
        $tz = $null
        if ($dq.PSObject.Properties['localTimeZoneId']) { $tz = $dq.localTimeZoneId }
        foreach ($day in $days) {
            foreach ($time in $times) {
                $DirectQueryList.Add([PSCustomObject]@{
                        DQFrequency       = $frequency
                        DQLocalTimeZoneId = $tz
                        DQDay             = $day
                        DQTime            = $time
                        DatasetId         = $datasetId
                        DatasetName       = $datasetName
                        WorkspaceId       = $WorkspaceId
                        WorkspaceName     = $WorkspaceName
                    })
            }
        }
        return
    }

    # Neither schedule is available: emit an explicit row instead of the monolith's silent all-null row.
    $row = Rename-Properties -object (New-Object PSObject) -renameMap $map
    Add-IQNote -Row $row -Name 'WorkspaceId' -Value $WorkspaceId
    Add-IQNote -Row $row -Name 'WorkspaceName' -Value $WorkspaceName
    Add-IQNote -Row $row -Name 'DatasetId' -Value $datasetId
    Add-IQNote -Row $row -Name 'DatasetName' -Value $datasetName
    Add-IQNote -Row $row -Name 'DatasetRefreshScheduleDay' -Value $null
    Add-IQNote -Row $row -Name 'DatasetRefreshScheduleTime' -Value $null
    Add-IQNote -Row $row -Name 'DatasetRefreshScheduleKind' -Value 'Unavailable'
    $ScheduleList.Add($row)
}

function Add-IQDatasetRefreshHistoryRow {
    <#
    .SYNOPSIS
        Refresh history rows of one dataset (monolith 2067-2088); honours Options.RefreshHistoryTop ($top) when set.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$BasePath,
        [Parameter(Mandatory = $true)]$DatasetRow,
        [Parameter(Mandatory = $true)][string]$WorkspaceId,
        [Parameter(Mandatory = $true)][string]$WorkspaceName,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$HistoryList,
        [Parameter(Mandatory = $false)][AllowNull()][System.Collections.Generic.List[object]]$ErrorList
    )
    $map = Get-IQRenameMap -Name DatasetRefresh
    $query = $null
    $top = Get-IQInventoryOption -Name 'RefreshHistoryTop'
    if ($null -ne $top -and [int]$top -gt 0) { $query = @{ '$top' = [int]$top } }
    $params = @{ Path = "$BasePath/refreshes"; Description = "refresh history of dataset '$($DatasetRow.DatasetName)'"; Item = $WorkspaceName; ErrorList = $ErrorList }
    if ($query) { $params.Query = $query }
    if ($DatasetRow.DatasetIsRefreshable -eq $false) { $params.Optional = $true }   # audit C5-10: expected to fail for DQ/live models
    foreach ($refresh in @(Get-IQInventoryList @params)) {
        $renamedRefreshRecord = Rename-Properties -object $refresh -renameMap $map
        Add-IQNote -Row $renamedRefreshRecord -Name 'WorkspaceId' -Value $WorkspaceId
        Add-IQNote -Row $renamedRefreshRecord -Name 'WorkspaceName' -Value $WorkspaceName
        Add-IQNote -Row $renamedRefreshRecord -Name 'DatasetId' -Value $DatasetRow.DatasetId
        Add-IQNote -Row $renamedRefreshRecord -Name 'DatasetName' -Value $DatasetRow.DatasetName
        $HistoryList.Add((ConvertTo-IQFlatRow -Row $renamedRefreshRecord))
    }
}

function Add-IQDatasetExtraRow {
    <#
    .SYNOPSIS
        NEW per-dataset collectors: parameters and dataset users (each optional: a 403/404 yields nothing plus one Debug line).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$BasePath,
        [Parameter(Mandatory = $true)]$DatasetRow,
        [Parameter(Mandatory = $true)][string]$WorkspaceId,
        [Parameter(Mandatory = $true)][string]$WorkspaceName,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$ParameterList,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$UserList
    )
    $datasetId = [string]$DatasetRow.DatasetId
    $datasetName = [string]$DatasetRow.DatasetName
    foreach ($p in @(Get-IQInventoryList -Path "$BasePath/parameters" -Optional -Description "parameters of dataset '$datasetName'" -Item $WorkspaceName)) {
        $suggested = $null
        if ($p.PSObject.Properties['suggestedValues'] -and $null -ne $p.suggestedValues) {
            try { $suggested = ConvertTo-Json -InputObject $p.suggestedValues -Depth 20 -Compress } catch { $suggested = [string]$p.suggestedValues }
        }
        $ParameterList.Add([PSCustomObject]@{
                ParameterName            = $p.name
                ParameterType            = $p.type
                ParameterIsRequired      = $p.isRequired
                ParameterCurrentValue    = $p.currentValue
                ParameterSuggestedValues = $suggested
                DatasetId                = $datasetId
                DatasetName              = $datasetName
                WorkspaceId              = $WorkspaceId
                WorkspaceName            = $WorkspaceName
            })
    }
    foreach ($u in @(Get-IQInventoryList -Path "$BasePath/users" -Optional -Description "users of dataset '$datasetName'" -Item $WorkspaceName)) {
        $UserList.Add([PSCustomObject]@{
                UserIdentifier             = $u.identifier
                UserPrincipalType          = $u.principalType
                UserDatasetUserAccessRight = $u.datasetUserAccessRight
                UserDisplayName            = $u.displayName
                UserEmailAddress           = $u.emailAddress
                UserGraphId                = $u.graphId
                DatasetId                  = $datasetId
                DatasetName                = $datasetName
                WorkspaceId                = $WorkspaceId
                WorkspaceName              = $WorkspaceName
            })
    }
}

function Add-IQDashboardRow {
    <#
    .SYNOPSIS
        NEW: dashboards and their tiles for a workspace ('groups/<id>' base) or My Workspace ('' base); optional collectors.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$BasePath,   # 'groups/<ws>' or ''
        [Parameter(Mandatory = $true)][string]$WorkspaceId,
        [Parameter(Mandatory = $true)][string]$WorkspaceName,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$DashboardList,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$TileList
    )
    $prefix = ''
    if (-not [string]::IsNullOrEmpty($BasePath)) { $prefix = $BasePath.TrimEnd('/') + '/' }
    foreach ($d in @(Get-IQInventoryList -Path "${prefix}dashboards" -Optional -Description "dashboards of workspace '$WorkspaceName'" -Item $WorkspaceName)) {
        $DashboardList.Add([PSCustomObject]@{
                DashboardId         = $d.id
                DashboardName       = $d.displayName
                DashboardIsReadOnly = $d.isReadOnly
                DashboardWebUrl     = $d.webUrl
                DashboardEmbedUrl   = $d.embedUrl
                WorkspaceId         = $WorkspaceId
                WorkspaceName       = $WorkspaceName
            })
        foreach ($t in @(Get-IQInventoryList -Path "${prefix}dashboards/$($d.id)/tiles" -Optional -Description "tiles of dashboard '$($d.displayName)'" -Item $WorkspaceName)) {
            $TileList.Add([PSCustomObject]@{
                    TileId        = $t.id
                    TileTitle     = $t.title
                    TileSubTitle  = $t.subTitle
                    TileRowSpan   = $t.rowSpan
                    TileColSpan   = $t.colSpan
                    TileEmbedUrl  = $t.embedUrl
                    ReportId      = $t.reportId
                    DatasetId     = $t.datasetId
                    DashboardId   = $d.id
                    DashboardName = $d.displayName
                    WorkspaceId   = $WorkspaceId
                    WorkspaceName = $WorkspaceName
                })
        }
    }
}

function New-IQWorkspaceInventoryObject {
    <#
    .SYNOPSIS
        Empty per-workspace inventory container (the keys of ws-<id>.json).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$WorkspaceId,
        [Parameter(Mandatory = $true)][string]$WorkspaceName,
        [Parameter(Mandatory = $false)][AllowNull()]$Workspace,
        [Parameter(Mandatory = $false)][string]$RunMode = 'Workspaces'
    )
    $inv = [ordered]@{
        WorkspaceId                       = $WorkspaceId
        WorkspaceName                     = $WorkspaceName
        Workspace                         = $Workspace
        RunMode                           = $RunMode
        CollectedUtc                      = $null
        IsSynthetic                       = $false
        SharedReportsFound                = $false
        Datasets                          = @()
        DatasetSources                    = @()
        DatasetRefreshHistory             = @()
        DatasetRefreshSchedule            = @()
        Reports                           = @()
        ReportPages                       = @()
        Dataflows                         = @()
        DataflowSources                   = @()
        DataflowLineage                   = @()
        DataflowRefreshHistory            = @()
        FabricItems                       = @()
        ItemConnections                   = @()
        ReportsWithSensitivityLabel       = @()
        Dashboards                        = @()
        DashboardTiles                    = @()
        WorkspaceUsers                    = @()
        DatasetUsers                      = @()
        DatasetParameters                 = @()
        DatasetDirectQueryRefreshSchedule = @()
        Errors                            = @()
    }
    return $inv
}

# =====================================================================================================================
# Per-workspace collection (monolith 1940-2356 merged into one pass)
# =====================================================================================================================

function Get-IQWorkspaceInventory {
    <#
    .SYNOPSIS
        Collects everything for one workspace (datasets, sources, refresh history/schedule, reports, pages, dataflows, lineage, Fabric items, item connections + the new collectors) into one object (ws-<id>.json shape).
    .DESCRIPTION
        Reproduces the monolith's per-workspace loops (1940-2356) with every call guarded. In Reports/Models modes only
        the selected datasets/reports (and their sources/pages/refreshes) are kept, exactly like the monolith's
        post-filters at 2016-2035; dataflows, lineage and Fabric items stay workspace-wide as before. The name lookups are
        hashtables shared across workspaces (DatasetId -> name, DataflowId -> name).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Workspace,
        [Parameter(Mandatory = $false)][AllowNull()]$Scope,
        [Parameter(Mandatory = $false)][hashtable]$DatasetNameLookup = @{},
        [Parameter(Mandatory = $false)][hashtable]$DataflowNameLookup = @{}
    )
    $wsId = [string]$Workspace.WorkspaceId
    $wsName = [string]$Workspace.WorkspaceName
    $runMode = 'Workspaces'
    if ($null -ne $Scope -and $Scope.RunMode) { $runMode = [string]$Scope.RunMode }
    $cache = Get-IQInventoryCache
    $errors = New-Object System.Collections.Generic.List[object]

    $datasets = New-Object System.Collections.Generic.List[object]
    $datasetSources = New-Object System.Collections.Generic.List[object]
    $refreshHistory = New-Object System.Collections.Generic.List[object]
    $refreshSchedule = New-Object System.Collections.Generic.List[object]
    $dqSchedule = New-Object System.Collections.Generic.List[object]
    $datasetParameters = New-Object System.Collections.Generic.List[object]
    $datasetUsers = New-Object System.Collections.Generic.List[object]
    $reports = New-Object System.Collections.Generic.List[object]
    $reportPages = New-Object System.Collections.Generic.List[object]
    $dataflows = New-Object System.Collections.Generic.List[object]
    $dataflowSources = New-Object System.Collections.Generic.List[object]
    $dataflowLineage = New-Object System.Collections.Generic.List[object]
    $dataflowRefreshHistory = New-Object System.Collections.Generic.List[object]
    $fabricItemRows = New-Object System.Collections.Generic.List[object]
    $itemConnections = New-Object System.Collections.Generic.List[object]
    $labelledReportIds = New-Object System.Collections.Generic.List[object]
    $dashboards = New-Object System.Collections.Generic.List[object]
    $tiles = New-Object System.Collections.Generic.List[object]
    $workspaceUsers = New-Object System.Collections.Generic.List[object]

    Write-IQLog -Level Info -Stage Inventory -Item $wsName -Message 'Report & Model metadata extraction started.'

    # ---- Fabric items + item connections (monolith 2282-2356), first so the sensitivity-label lookup is available ----
    $labelById = @{}
    $fabricItems = @(Get-IQInventoryList -Path "workspaces/$wsId/items" -Api Fabric -Optional -Description "Fabric items of workspace '$wsName'" -Item $wsName)
    $fabricItemsMap = Get-IQRenameMap -Name FabricItems
    foreach ($fabricItem in $fabricItems) {
        if ($null -eq $fabricItem) { continue }
        $labelId = $null
        if ($fabricItem.PSObject.Properties['sensitivityLabel'] -and $fabricItem.sensitivityLabel) {
            $labelId = Get-IQMemberValue -Object $fabricItem.sensitivityLabel -Name 'labelId'
            if ($null -eq $labelId) { $labelId = Get-IQMemberValue -Object $fabricItem.sensitivityLabel -Name 'id' }
            if ($null -eq $labelId) { $labelId = [string]$fabricItem.sensitivityLabel }
            $labelById[[string]$fabricItem.id] = $labelId
            if ($fabricItem.type -eq 'Report') { $labelledReportIds.Add([string]$fabricItem.id) }
        }
        foreach ($connection in @(Get-IQInventoryList -Path "workspaces/$wsId/items/$($fabricItem.id)/connections" -Api Fabric -Optional -Description "connections of Fabric item '$($fabricItem.displayName)'" -Item $wsName)) {
            $itemConnections.Add([PSCustomObject]@{
                    WorkspaceId      = $wsId
                    WorkspaceName    = $wsName
                    FabricItemId     = $fabricItem.id
                    FabricItemName   = $fabricItem.displayName
                    FabricItemType   = $fabricItem.type
                    ConnectionId     = $connection.id
                    ConnectionName   = $connection.displayName
                    GatewayId        = $connection.gatewayId
                    ConnectivityType = $connection.connectivityType
                    ConnectionType   = $connection.connectionDetails.type
                    ConnectionPath   = $connection.connectionDetails.path
                })
        }
    }
    foreach ($item in @($fabricItems | Where-Object { $null -ne $_ -and $_.type -ne 'Report' -and $_.type -ne 'SemanticModel' })) {
        $renamedItem = Rename-Properties -object $item -renameMap $fabricItemsMap
        Add-IQNote -Row $renamedItem -Name 'WorkspaceId' -Value $wsId
        Add-IQNote -Row $renamedItem -Name 'WorkspaceName' -Value $wsName
        $folderId = $null
        if ($item.PSObject.Properties['folderId']) { $folderId = $item.folderId }
        Add-IQNote -Row $renamedItem -Name 'FabricItemSensitivityLabelId' -Value $labelById[[string]$item.id]
        Add-IQNote -Row $renamedItem -Name 'FabricItemFolderId' -Value $folderId
        $fabricItemRows.Add((ConvertTo-IQFlatRow -Row $renamedItem))
    }

    # ---- Datasets + datasources + refresh history/schedule + NEW parameters/users (monolith 1946-1974, 2067-2125) ----
    $datasetMap = Get-IQRenameMap -Name Dataset
    $datasetDatasourceMap = Get-IQRenameMap -Name DatasetDatasource
    $datasetsRaw = $null
    if ($cache.Datasets.ContainsKey($wsId)) { $datasetsRaw = @($cache.Datasets[$wsId]); $cache.Datasets.Remove($wsId) }
    else { $datasetsRaw = @(Get-IQInventoryList -Path "groups/$wsId/datasets" -Description "datasets of workspace '$wsName'" -Item $wsName -ErrorList $errors) }
    foreach ($dataset in $datasetsRaw) {
        if ($null -eq $dataset) { continue }
        # Store the DatasetId and DatasetName in the lookup table (for every dataset, like the monolith, before filtering)
        if ($dataset.id) { $DatasetNameLookup[[string]$dataset.id] = $dataset.name }
        if (-not (Test-IQScopeIncludesDataset -Scope $Scope -DatasetId ([string]$dataset.id))) { continue }

        $renamedDataset = Rename-Properties -object $dataset -renameMap $datasetMap
        Add-IQNote -Row $renamedDataset -Name 'WorkspaceId' -Value $wsId
        Add-IQNote -Row $renamedDataset -Name 'WorkspaceName' -Value $wsName
        Add-IQNote -Row $renamedDataset -Name 'DatasetCapacityId' -Value (Get-IQMemberValue -Object $Workspace -Name 'WorkspaceCapacityId')
        Add-IQNote -Row $renamedDataset -Name 'DatasetSensitivityLabelId' -Value $labelById[[string]$dataset.id]
        $datasetRow = ConvertTo-IQFlatRow -Row $renamedDataset
        $datasets.Add($datasetRow)

        $base = "groups/$wsId/datasets/$($dataset.id)"
        # Fetch dataset sources
        foreach ($datasource in @(Get-IQInventoryList -Path "$base/datasources" -Description "datasources of dataset '$($dataset.name)'" -Item $wsName -ErrorList $errors)) {
            $renamedDatasource = Rename-Properties -object $datasource -renameMap $datasetDatasourceMap
            Add-IQNote -Row $renamedDatasource -Name 'WorkspaceId' -Value $wsId
            Add-IQNote -Row $renamedDatasource -Name 'WorkspaceName' -Value $wsName
            Add-IQNote -Row $renamedDatasource -Name 'DatasetId' -Value $dataset.id
            Add-IQNote -Row $renamedDatasource -Name 'DatasetName' -Value $dataset.name
            if ($datasource.connectionDetails) {
                $renamedDatasource.DatasetDatasourceConnectionDetails = $datasource.connectionDetails | ConvertTo-Json -Compress
            }
            $datasetSources.Add((ConvertTo-IQFlatRow -Row $renamedDatasource))
        }
        Add-IQDatasetRefreshHistoryRow -BasePath $base -DatasetRow $datasetRow -WorkspaceId $wsId -WorkspaceName $wsName -HistoryList $refreshHistory -ErrorList $errors
        Add-IQDatasetRefreshScheduleRow -BasePath $base -DatasetRow $datasetRow -WorkspaceId $wsId -WorkspaceName $wsName -ScheduleList $refreshSchedule -DirectQueryList $dqSchedule
        Add-IQDatasetExtraRow -BasePath $base -DatasetRow $datasetRow -WorkspaceId $wsId -WorkspaceName $wsName -ParameterList $datasetParameters -UserList $datasetUsers
    }

    # ---- Reports + pages (monolith 1976-2013) ----
    $reportMap = Get-IQRenameMap -Name Report
    $pageMap = Get-IQRenameMap -Name Page
    $reportsRaw = $null
    if ($cache.Reports.ContainsKey($wsId)) { $reportsRaw = @($cache.Reports[$wsId]); $cache.Reports.Remove($wsId) }
    else { $reportsRaw = @(Get-IQInventoryList -Path "groups/$wsId/reports" -Description "reports of workspace '$wsName'" -Item $wsName -ErrorList $errors) }
    foreach ($report in $reportsRaw) {
        if ($null -eq $report) { continue }
        if (-not (Test-IQScopeIncludesReport -Scope $Scope -ReportId ([string]$report.id))) { continue }
        $renamedReport = Rename-Properties -object $report -renameMap $reportMap
        Add-IQNote -Row $renamedReport -Name 'WorkspaceId' -Value $wsId
        Add-IQNote -Row $renamedReport -Name 'WorkspaceName' -Value $wsName
        # Retrieve and add the correct DatasetName from the lookup table if DatasetId exists
        $datasetId = [string]$report.datasetId
        if ($datasetId -and $DatasetNameLookup.ContainsKey($datasetId)) {
            Add-IQNote -Row $renamedReport -Name 'DatasetName' -Value $DatasetNameLookup[$datasetId]
        }
        else {
            Add-IQNote -Row $renamedReport -Name 'DatasetName' -Value 'Unknown Dataset'
        }
        Add-IQNote -Row $renamedReport -Name 'ReportSensitivityLabelId' -Value $labelById[[string]$report.id]
        Add-IQNote -Row $renamedReport -Name 'ReportHasSensitivityLabel' -Value ($labelledReportIds -contains [string]$report.id)
        $reports.Add((ConvertTo-IQFlatRow -Row $renamedReport))

        # Fetch report pages (paginated reports have none - audit C5-10)
        if ($report.PSObject.Properties['reportType'] -and $report.reportType -eq 'PaginatedReport') { continue }
        foreach ($page in @(Get-IQInventoryList -Path "groups/$wsId/reports/$($report.id)/pages" -Optional -Description "pages of report '$($report.name)'" -Item $wsName)) {
            $renamedPage = Rename-Properties -object $page -renameMap $pageMap
            Add-IQNote -Row $renamedPage -Name 'WorkspaceId' -Value $wsId
            Add-IQNote -Row $renamedPage -Name 'WorkspaceName' -Value $wsName
            Add-IQNote -Row $renamedPage -Name 'ReportId' -Value $report.id
            Add-IQNote -Row $renamedPage -Name 'ReportName' -Value $report.name
            $reportPages.Add((ConvertTo-IQFlatRow -Row $renamedPage))
        }
    }

    # ---- NEW: dashboards/tiles and workspace users ----
    Add-IQDashboardRow -BasePath "groups/$wsId" -WorkspaceId $wsId -WorkspaceName $wsName -DashboardList $dashboards -TileList $tiles
    foreach ($u in @(Get-IQInventoryList -Path "groups/$wsId/users" -Optional -Description "users of workspace '$wsName'" -Item $wsName)) {
        $workspaceUsers.Add([PSCustomObject]@{
                UserEmailAddress         = $u.emailAddress
                UserDisplayName          = $u.displayName
                UserIdentifier           = $u.identifier
                UserPrincipalType        = $u.principalType
                UserGroupUserAccessRight = $u.groupUserAccessRight
                UserGraphId              = $u.graphId
                WorkspaceId              = $wsId
                WorkspaceName            = $wsName
            })
    }

    # ---- Dataflows Gen1/Gen2 + datasources (monolith 2128-2183) ----
    $dataflowMap = Get-IQRenameMap -Name Dataflow
    $dataflowDatasourceMap = Get-IQRenameMap -Name DataflowDatasource
    foreach ($dataflow in @(Get-IQInventoryList -Path "groups/$wsId/dataflows" -Description "dataflows of workspace '$wsName'" -Item $wsName -ErrorList $errors)) {
        if ($null -eq $dataflow) { continue }
        $renamedDataflow = Rename-Properties -object $dataflow -renameMap $dataflowMap
        Add-IQNote -Row $renamedDataflow -Name 'WorkspaceId' -Value $wsId
        Add-IQNote -Row $renamedDataflow -Name 'WorkspaceName' -Value $wsName
        # Store DataflowId and DataflowName in a lookup table
        if ($dataflow.objectId) { $DataflowNameLookup[[string]$dataflow.objectId] = $dataflow.name }
        $dataflows.Add((ConvertTo-IQFlatRow -Row $renamedDataflow))

        # Fetch Dataflow Datasources (guarded against a null objectId - audit C5-07)
        if (-not $dataflow.objectId) { continue }
        foreach ($datasource in @(Get-IQInventoryList -Path "groups/$wsId/dataflows/$($dataflow.objectId)/datasources" -Description "datasources of dataflow '$($dataflow.name)'" -Item $wsName -ErrorList $errors)) {
            $renamedDataflowDatasource = Rename-Properties -object $datasource -renameMap $dataflowDatasourceMap
            Add-IQNote -Row $renamedDataflowDatasource -Name 'WorkspaceId' -Value $wsId
            Add-IQNote -Row $renamedDataflowDatasource -Name 'WorkspaceName' -Value $wsName
            Add-IQNote -Row $renamedDataflowDatasource -Name 'DataflowId' -Value $dataflow.objectId
            if ($datasource.connectionDetails) {
                $renamedDataflowDatasource.DataflowDatasourceConnectionDetails = $datasource.connectionDetails | ConvertTo-Json -Compress
            }
            if ($DataflowNameLookup.ContainsKey([string]$dataflow.objectId)) {
                Add-IQNote -Row $renamedDataflowDatasource -Name 'DataflowName' -Value $DataflowNameLookup[[string]$dataflow.objectId]
            }
            else {
                Add-IQNote -Row $renamedDataflowDatasource -Name 'DataflowName' -Value 'Unknown Dataflow'
            }
            $dataflowSources.Add((ConvertTo-IQFlatRow -Row $renamedDataflowDatasource))
        }
    }

    # ---- Gen 2 CICD dataflows via the Fabric API (monolith 2185-2217, fixed per audit C5-06: environment prefix, paged, Fabric token, logged) ----
    $fabricDataflowMap = Get-IQRenameMap -Name FabricDataflow
    foreach ($fabricDataflow in @(Get-IQInventoryList -Path "workspaces/$wsId/dataflows" -Api Fabric -Optional -Description "Fabric dataflows of workspace '$wsName'" -Item $wsName)) {
        if ($null -eq $fabricDataflow) { continue }
        $renamedFabricDataflow = Rename-Properties -object $fabricDataflow -renameMap $fabricDataflowMap
        Add-IQNote -Row $renamedFabricDataflow -Name 'WorkspaceId' -Value $wsId
        Add-IQNote -Row $renamedFabricDataflow -Name 'WorkspaceName' -Value $wsName
        Add-IQNote -Row $renamedFabricDataflow -Name 'DataflowGeneration' -Value 'Gen 2 CICD'
        $sensitivityId = $null
        if ($fabricDataflow.PSObject.Properties['sensitivityLabel'] -and $fabricDataflow.sensitivityLabel) {
            $sensitivityId = Get-IQMemberValue -Object $fabricDataflow.sensitivityLabel -Name 'labelId'
            if ($null -eq $sensitivityId) { $sensitivityId = Get-IQMemberValue -Object $fabricDataflow.sensitivityLabel -Name 'id' }
        }
        $folderId = $null
        if ($fabricDataflow.PSObject.Properties['folderId']) { $folderId = $fabricDataflow.folderId }
        Add-IQNote -Row $renamedFabricDataflow -Name 'DataflowSensitivityLabelId' -Value $sensitivityId
        Add-IQNote -Row $renamedFabricDataflow -Name 'DataflowFolderId' -Value $folderId
        if ($fabricDataflow.id) { $DataflowNameLookup[[string]$fabricDataflow.id] = $fabricDataflow.displayName }
        $dataflows.Add((ConvertTo-IQFlatRow -Row $renamedFabricDataflow))
    }

    # ---- Dataflow lineage (monolith 2222-2251); the API's workspaceObjectId is kept as DataflowWorkspaceId (audit C5-05) ----
    $lineageMap = Get-IQRenameMap -Name DataflowLineage
    foreach ($link in @(Get-IQInventoryList -Path "groups/$wsId/datasets/upstreamDataflows" -Description "dataflow lineage of workspace '$wsName'" -Item $wsName -ErrorList $errors)) {
        if ($null -eq $link) { continue }
        $renamedLink = Rename-Properties -object $link -renameMap $lineageMap
        $apiWorkspaceId = $null
        if ($link.PSObject.Properties['workspaceObjectId']) { $apiWorkspaceId = $link.workspaceObjectId }
        Add-IQNote -Row $renamedLink -Name 'DataflowWorkspaceId' -Value $apiWorkspaceId
        Add-IQNote -Row $renamedLink -Name 'WorkspaceId' -Value $wsId
        Add-IQNote -Row $renamedLink -Name 'WorkspaceName' -Value $wsName
        $dataflowId = [string]$link.dataflowObjectId
        if ($dataflowId -and $DataflowNameLookup.ContainsKey($dataflowId)) {
            Add-IQNote -Row $renamedLink -Name 'DataflowName' -Value $DataflowNameLookup[$dataflowId]
        }
        else {
            Add-IQNote -Row $renamedLink -Name 'DataflowName' -Value 'Unknown Dataflow'
        }
        $datasetId = [string]$link.datasetObjectId
        if ($datasetId -and $DatasetNameLookup.ContainsKey($datasetId)) {
            Add-IQNote -Row $renamedLink -Name 'DatasetName' -Value $DatasetNameLookup[$datasetId]
        }
        else {
            Add-IQNote -Row $renamedLink -Name 'DatasetName' -Value 'Unknown Dataset'
        }
        $dataflowLineage.Add((ConvertTo-IQFlatRow -Row $renamedLink))
    }

    # ---- Dataflow refresh history (monolith 2253-2280); Fabric CI/CD dataflows are skipped (audit C5-07) ----
    $dataflowRefreshMap = Get-IQRenameMap -Name DataflowRefresh
    foreach ($dataflow in $dataflows) {
        if ($dataflow.DataflowGeneration -eq 'Gen 2 CICD') { continue }
        if (-not $dataflow.DataflowId) { continue }
        foreach ($refresh in @(Get-IQInventoryList -Path "groups/$wsId/dataflows/$($dataflow.DataflowId)/transactions" -Optional -Description "refresh history of dataflow '$($dataflow.DataflowName)'" -Item $wsName)) {
            $renamedRefreshRecord = Rename-Properties -object $refresh -renameMap $dataflowRefreshMap
            Add-IQNote -Row $renamedRefreshRecord -Name 'WorkspaceId' -Value $wsId
            Add-IQNote -Row $renamedRefreshRecord -Name 'WorkspaceName' -Value $wsName
            Add-IQNote -Row $renamedRefreshRecord -Name 'DataflowId' -Value $dataflow.DataflowId
            if ($DataflowNameLookup.ContainsKey([string]$dataflow.DataflowId)) {
                Add-IQNote -Row $renamedRefreshRecord -Name 'DataflowName' -Value $DataflowNameLookup[[string]$dataflow.DataflowId]
            }
            else {
                Add-IQNote -Row $renamedRefreshRecord -Name 'DataflowName' -Value 'Unknown Dataflow'
            }
            $dataflowRefreshHistory.Add((ConvertTo-IQFlatRow -Row $renamedRefreshRecord))
        }
    }

    # ---- assemble ----
    $inv = New-IQWorkspaceInventoryObject -WorkspaceId $wsId -WorkspaceName $wsName -Workspace $Workspace -RunMode $runMode
    $inv.CollectedUtc = [DateTime]::UtcNow.ToString('o')
    $inv.Datasets = $datasets.ToArray()
    $inv.DatasetSources = $datasetSources.ToArray()
    $inv.DatasetRefreshHistory = $refreshHistory.ToArray()
    $inv.DatasetRefreshSchedule = $refreshSchedule.ToArray()
    $inv.Reports = $reports.ToArray()
    $inv.ReportPages = $reportPages.ToArray()
    $inv.Dataflows = $dataflows.ToArray()
    $inv.DataflowSources = $dataflowSources.ToArray()
    $inv.DataflowLineage = $dataflowLineage.ToArray()
    $inv.DataflowRefreshHistory = $dataflowRefreshHistory.ToArray()
    $inv.FabricItems = $fabricItemRows.ToArray()
    $inv.ItemConnections = $itemConnections.ToArray()
    $inv.ReportsWithSensitivityLabel = $labelledReportIds.ToArray()
    $inv.Dashboards = $dashboards.ToArray()
    $inv.DashboardTiles = $tiles.ToArray()
    $inv.WorkspaceUsers = $workspaceUsers.ToArray()
    $inv.DatasetUsers = $datasetUsers.ToArray()
    $inv.DatasetParameters = $datasetParameters.ToArray()
    $inv.DatasetDirectQueryRefreshSchedule = $dqSchedule.ToArray()
    $inv.Errors = $errors.ToArray()

    Write-IQLog -Level Info -Stage Inventory -Item $wsName -Message ("Collected {0} dataset(s), {1} report(s), {2} page(s), {3} dataflow(s), {4} Fabric item(s), {5} dashboard(s), {6} error(s)." -f `
            $datasets.Count, $reports.Count, $reportPages.Count, $dataflows.Count, $fabricItemRows.Count, $dashboards.Count, $errors.Count)
    return $inv
}

# =====================================================================================================================
# My Workspace (monolith 2356-2513)
# =====================================================================================================================

function Get-IQMyWorkspaceInventory {
    <#
    .SYNOPSIS
        Collects "My Workspace" (datasets, reports incl. shared ones, pages, refresh history/schedule + new collectors) using the myorg endpoints (monolith 2356-2513, brief 6.3).
    .DESCRIPTION
        Reports already captured from a selected workspace or an app (KnownReportIds) and app copies (appId) are skipped;
        reports not owned by the user land under the synthetic workspace 'Shared Reports (No Workspace Access)'.
        The Add-Member bug of the monolith (2445-2446: page rows without ReportId/ReportName, audit C6-01) is fixed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Scope,
        [Parameter(Mandatory = $false)][hashtable]$KnownReportIds = @{},
        [Parameter(Mandatory = $false)][hashtable]$DatasetNameLookup = @{}
    )
    $myWorkspaceId = 'My Workspace'
    $myWorkspaceName = 'My Workspace'
    $sharedName = 'Shared Reports (No Workspace Access)'
    $runMode = 'Workspaces'
    if ($null -ne $Scope -and $Scope.RunMode) { $runMode = [string]$Scope.RunMode }
    $cache = Get-IQInventoryCache
    $errors = New-Object System.Collections.Generic.List[object]

    $datasets = New-Object System.Collections.Generic.List[object]
    $datasetSources = New-Object System.Collections.Generic.List[object]
    $refreshHistory = New-Object System.Collections.Generic.List[object]
    $refreshSchedule = New-Object System.Collections.Generic.List[object]
    $dqSchedule = New-Object System.Collections.Generic.List[object]
    $datasetParameters = New-Object System.Collections.Generic.List[object]
    $datasetUsers = New-Object System.Collections.Generic.List[object]
    $reports = New-Object System.Collections.Generic.List[object]
    $reportPages = New-Object System.Collections.Generic.List[object]
    $dashboards = New-Object System.Collections.Generic.List[object]
    $tiles = New-Object System.Collections.Generic.List[object]

    Write-IQLog -Level Info -Stage Inventory -Item $myWorkspaceName -Message 'My Workspace metadata extract started.'

    # Fetch datasets from "My Workspace"
    $datasetMap = Get-IQRenameMap -Name Dataset
    $datasetDatasourceMap = Get-IQRenameMap -Name DatasetDatasource
    $myDatasets = $cache.MyWorkspaceDatasets
    if ($null -eq $myDatasets) { $myDatasets = @(Get-IQInventoryList -Path 'datasets' -Description 'My Workspace datasets' -Item $myWorkspaceName -ErrorList $errors) }
    $cache.MyWorkspaceDatasets = $null
    foreach ($dataset in @($myDatasets)) {
        if ($null -eq $dataset) { continue }
        if ($dataset.id) { $DatasetNameLookup[[string]$dataset.id] = $dataset.name }
        if (-not (Test-IQScopeIncludesDataset -Scope $Scope -DatasetId ([string]$dataset.id))) { continue }
        $renamedDataset = Rename-Properties -object $dataset -renameMap $datasetMap
        Add-IQNote -Row $renamedDataset -Name 'WorkspaceId' -Value $myWorkspaceId
        Add-IQNote -Row $renamedDataset -Name 'WorkspaceName' -Value $myWorkspaceName
        Add-IQNote -Row $renamedDataset -Name 'DatasetCapacityId' -Value $null
        $datasetRow = ConvertTo-IQFlatRow -Row $renamedDataset
        $datasets.Add($datasetRow)

        $base = "datasets/$($dataset.id)"
        foreach ($datasource in @(Get-IQInventoryList -Path "$base/datasources" -Description "datasources of dataset '$($dataset.name)'" -Item $myWorkspaceName -ErrorList $errors)) {
            $renamedDatasource = Rename-Properties -object $datasource -renameMap $datasetDatasourceMap
            Add-IQNote -Row $renamedDatasource -Name 'WorkspaceId' -Value $myWorkspaceId
            Add-IQNote -Row $renamedDatasource -Name 'WorkspaceName' -Value $myWorkspaceName
            Add-IQNote -Row $renamedDatasource -Name 'DatasetId' -Value $dataset.id
            Add-IQNote -Row $renamedDatasource -Name 'DatasetName' -Value $dataset.name
            if ($datasource.connectionDetails) {
                $renamedDatasource.DatasetDatasourceConnectionDetails = $datasource.connectionDetails | ConvertTo-Json -Compress
            }
            $datasetSources.Add((ConvertTo-IQFlatRow -Row $renamedDatasource))
        }
        Add-IQDatasetRefreshHistoryRow -BasePath $base -DatasetRow $datasetRow -WorkspaceId $myWorkspaceId -WorkspaceName $myWorkspaceName -HistoryList $refreshHistory -ErrorList $errors
        Add-IQDatasetRefreshScheduleRow -BasePath $base -DatasetRow $datasetRow -WorkspaceId $myWorkspaceId -WorkspaceName $myWorkspaceName -ScheduleList $refreshSchedule -DirectQueryList $dqSchedule
        Add-IQDatasetExtraRow -BasePath $base -DatasetRow $datasetRow -WorkspaceId $myWorkspaceId -WorkspaceName $myWorkspaceName -ParameterList $datasetParameters -UserList $datasetUsers
    }

    # Fetch reports from "My Workspace"
    $reportMap = Get-IQRenameMap -Name Report
    $pageMap = Get-IQRenameMap -Name Page
    $myReports = $cache.MyWorkspaceReports
    if ($null -eq $myReports) { $myReports = @(Get-IQInventoryList -Path 'reports' -Description 'My Workspace reports' -Item $myWorkspaceName -ErrorList $errors) }
    $cache.MyWorkspaceReports = $null
    $sharedReportExists = $false
    foreach ($report in @($myReports)) {
        if ($null -eq $report) { continue }
        # Skip if already captured elsewhere (workspace/app)
        if ($report.id -and $KnownReportIds.ContainsKey(([string]$report.id).ToLowerInvariant())) { continue }
        # Skip reports with appId - these are shared via apps, not actually in My Workspace
        if ($report.PSObject.Properties['appId'] -and $report.appId) { continue }
        if (-not (Test-IQScopeIncludesReport -Scope $Scope -ReportId ([string]$report.id))) { continue }

        if ($report.PSObject.Properties['isOwnedByMe'] -and $report.isOwnedByMe -eq $false) {
            $workspaceIdValue = $sharedName
            $workspaceNameValue = $sharedName
            $sharedReportExists = $true
        }
        else {
            $workspaceIdValue = $myWorkspaceId
            $workspaceNameValue = $myWorkspaceName
        }

        $renamedReport = Rename-Properties -object $report -renameMap $reportMap
        Add-IQNote -Row $renamedReport -Name 'WorkspaceId' -Value $workspaceIdValue
        Add-IQNote -Row $renamedReport -Name 'WorkspaceName' -Value $workspaceNameValue
        $datasetId = [string]$report.datasetId
        if ($datasetId -and $DatasetNameLookup.ContainsKey($datasetId)) {
            Add-IQNote -Row $renamedReport -Name 'DatasetName' -Value $DatasetNameLookup[$datasetId]
        }
        else {
            Add-IQNote -Row $renamedReport -Name 'DatasetName' -Value 'Unknown Dataset'
        }
        Add-IQNote -Row $renamedReport -Name 'ReportSensitivityLabelId' -Value $null
        Add-IQNote -Row $renamedReport -Name 'ReportHasSensitivityLabel' -Value $false
        $reports.Add((ConvertTo-IQFlatRow -Row $renamedReport))

        if ($report.PSObject.Properties['reportType'] -and $report.reportType -eq 'PaginatedReport') { continue }
        foreach ($page in @(Get-IQInventoryList -Path "reports/$($report.id)/pages" -Optional -Description "pages of report '$($report.name)'" -Item $myWorkspaceName)) {
            $renamedPage = Rename-Properties -object $page -renameMap $pageMap
            Add-IQNote -Row $renamedPage -Name 'WorkspaceId' -Value $workspaceIdValue
            Add-IQNote -Row $renamedPage -Name 'WorkspaceName' -Value $workspaceNameValue
            Add-IQNote -Row $renamedPage -Name 'ReportId' -Value $report.id       # audit C6-01 fix
            Add-IQNote -Row $renamedPage -Name 'ReportName' -Value $report.name
            $reportPages.Add((ConvertTo-IQFlatRow -Row $renamedPage))
        }
    }

    # NEW: My Workspace dashboards + tiles
    Add-IQDashboardRow -BasePath '' -WorkspaceId $myWorkspaceId -WorkspaceName $myWorkspaceName -DashboardList $dashboards -TileList $tiles

    $inv = New-IQWorkspaceInventoryObject -WorkspaceId $myWorkspaceId -WorkspaceName $myWorkspaceName -Workspace (New-IQPseudoWorkspace -Name $myWorkspaceName) -RunMode $runMode
    $inv.CollectedUtc = [DateTime]::UtcNow.ToString('o')
    $inv.IsSynthetic = $true
    $inv.SharedReportsFound = $sharedReportExists
    $inv.Datasets = $datasets.ToArray()
    $inv.DatasetSources = $datasetSources.ToArray()
    $inv.DatasetRefreshHistory = $refreshHistory.ToArray()
    $inv.DatasetRefreshSchedule = $refreshSchedule.ToArray()
    $inv.Reports = $reports.ToArray()
    $inv.ReportPages = $reportPages.ToArray()
    $inv.Dashboards = $dashboards.ToArray()
    $inv.DashboardTiles = $tiles.ToArray()
    $inv.DatasetUsers = $datasetUsers.ToArray()
    $inv.DatasetParameters = $datasetParameters.ToArray()
    $inv.DatasetDirectQueryRefreshSchedule = $dqSchedule.ToArray()
    $inv.Errors = $errors.ToArray()

    Write-IQLog -Level Info -Stage Inventory -Item $myWorkspaceName -Message ("Collected {0} dataset(s), {1} report(s) ({2} shared), {3} page(s), {4} dashboard(s), {5} error(s)." -f `
            $datasets.Count, $reports.Count, @($reports | Where-Object { $_.WorkspaceId -eq $sharedName }).Count, $reportPages.Count, $dashboards.Count, $errors.Count)
    return $inv
}

# =====================================================================================================================
# Global collectors (monolith 1874-1936 connections/gateways, 2036-2065 apps; NEW capacities)
# =====================================================================================================================

function Get-IQGlobalInventory {
    <#
    .SYNOPSIS
        Collects the tenant-wide lists: Fabric Connections and Gateways (monolith 1874-1936), Apps + AppReports for the selected workspaces (2036-2065) and NEW Capacities.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowEmptyCollection()][array]$Workspaces = @())
    $connections = New-Object System.Collections.Generic.List[object]
    $gateways = New-Object System.Collections.Generic.List[object]
    $capacities = New-Object System.Collections.Generic.List[object]
    $apps = New-Object System.Collections.Generic.List[object]
    $appReports = New-Object System.Collections.Generic.List[object]
    $appReportIds = New-Object System.Collections.Generic.List[object]
    $originalReportObjectIds = New-Object System.Collections.Generic.List[object]
    $errors = New-Object System.Collections.Generic.List[object]

    Write-IQLog -Level Info -Stage Inventory -Message 'Fabric connection metadata extraction started.'
    foreach ($connection in @(Get-IQInventoryList -Path 'connections' -Api Fabric -Optional -Description 'Fabric List Connections')) {
        $detailsJson = $null
        if ($connection.PSObject.Properties['connectionDetails'] -and $null -ne $connection.connectionDetails) {
            try { $detailsJson = ConvertTo-Json -InputObject $connection.connectionDetails -Depth 20 -Compress } catch { $detailsJson = $null }
        }
        $connections.Add([PSCustomObject]@{
                ConnectionId                   = $connection.id
                ConnectionName                 = $connection.displayName
                GatewayId                      = $connection.gatewayId
                ConnectivityType               = $connection.connectivityType
                ConnectionType                 = $connection.connectionDetails.type
                ConnectionPath                 = $connection.connectionDetails.path
                PrivacyLevel                   = $connection.privacyLevel
                CredentialType                 = $connection.credentialDetails.credentialType
                SingleSignOnType               = $connection.credentialDetails.singleSignOnType
                ConnectionEncryption           = $connection.credentialDetails.connectionEncryption
                SkipTestConnection             = $connection.credentialDetails.skipTestConnection
                AllowConnectionUsageInGateway  = $connection.allowConnectionUsageInGateway
                AllowUsageInUserControlledCode = $connection.allowUsageInUserControlledCode
                CreatedDateTime                = $connection.connectionRecency.createdDateTime
                LastBoundDateTime              = $connection.connectionRecency.lastBoundDateTime
                LastCredentialUsedDateTime     = $connection.connectionRecency.lastCredentialUsedDateTime
                MyLastBoundDateTime            = $connection.connectionRecency.myLastBoundDateTime
                MyLastCredentialUsedDateTime   = $connection.connectionRecency.myLastCredentialUsedDateTime
                ConnectionDetailsJson          = $detailsJson
            })
    }
    Write-IQLog -Level Info -Stage Inventory -Message ("Fabric connection metadata extraction completed. Found {0} connection(s)." -f $connections.Count)

    Write-IQLog -Level Info -Stage Inventory -Message 'Fabric gateway metadata extraction started.'
    foreach ($gateway in @(Get-IQInventoryList -Path 'gateways' -Api Fabric -Optional -Description 'Fabric List Gateways')) {
        $gateways.Add([PSCustomObject]@{
                GatewayId                    = $gateway.id
                GatewayName                  = $gateway.displayName
                GatewayType                  = $gateway.type
                GatewayVersion               = $gateway.version
                CapacityId                   = $gateway.capacityId
                NumberOfMemberGateways       = $gateway.numberOfMemberGateways
                MinMemberGatewayCount        = $gateway.minMemberGatewayCount
                MaxMemberGatewayCount        = $gateway.maxMemberGatewayCount
                LoadBalancingSetting         = $gateway.loadBalancingSetting
                AllowCloudConnectionRefresh  = $gateway.allowCloudConnectionRefresh
                AllowCustomConnectors        = $gateway.allowCustomConnectors
                InactivityMinutesBeforeSleep = $gateway.inactivityMinutesBeforeSleep
                SubscriptionId               = $gateway.virtualNetworkAzureResource.subscriptionId
                ResourceGroupName            = $gateway.virtualNetworkAzureResource.resourceGroupName
                VirtualNetworkName           = $gateway.virtualNetworkAzureResource.virtualNetworkName
                SubnetName                   = $gateway.virtualNetworkAzureResource.subnetName
            })
    }
    Write-IQLog -Level Info -Stage Inventory -Message ("Fabric gateway metadata extraction completed. Found {0} gateway(s)." -f $gateways.Count)

    # NEW: capacities (Power BI GET capacities - the ones the user can see)
    foreach ($cap in @(Get-IQInventoryList -Path 'capacities' -Optional -Description 'capacities')) {
        $admins = $null
        if ($cap.PSObject.Properties['admins'] -and $null -ne $cap.admins) { $admins = (@($cap.admins) | ForEach-Object { [string]$_ }) -join ';' }
        $capacities.Add([PSCustomObject]@{
                CapacityId               = $cap.id
                CapacityDisplayName      = $cap.displayName
                CapacitySku              = $cap.sku
                CapacityState            = $cap.state
                CapacityRegion           = $cap.region
                CapacityAdmins           = $admins
                CapacityUsersAccessRight = $cap.capacityUserAccessRight
            })
    }
    Write-IQLog -Level Info -Stage Inventory -Message ("Found {0} capacity(ies)." -f $capacities.Count)

    # Fetch Apps and App Reports that are in filtered workspaces (monolith 2036-2065)
    $appMap = Get-IQRenameMap -Name App
    $appReportMap = Get-IQRenameMap -Name AppReport
    $wsIds = @($Workspaces | ForEach-Object { [string]$_.WorkspaceId })
    foreach ($app in @(Get-IQInventoryList -Path 'apps' -Description 'apps' -ErrorList $errors)) {
        if ($null -eq $app) { continue }
        if ($wsIds -contains [string]$app.workspaceId) {
            $renamedApp = Rename-Properties -object $app -renameMap $appMap
            $apps.Add((ConvertTo-IQFlatRow -Row $renamedApp))
            # Fetch reports within each app
            foreach ($report in @(Get-IQInventoryList -Path "apps/$($app.id)/reports" -Optional -Description "reports of app '$($app.name)'")) {
                $renamedAppReport = Rename-Properties -object $report -renameMap $appReportMap
                Add-IQNote -Row $renamedAppReport -Name 'AppId' -Value $app.id
                Add-IQNote -Row $renamedAppReport -Name 'AppName' -Value $app.name
                $appReports.Add((ConvertTo-IQFlatRow -Row $renamedAppReport))
                if ($report.id) { $appReportIds.Add([string]$report.id) }
                if ($report.PSObject.Properties['originalReportObjectId'] -and $report.originalReportObjectId) { $originalReportObjectIds.Add([string]$report.originalReportObjectId) }
            }
        }
    }
    Write-IQLog -Level Info -Stage Inventory -Message ("Found {0} app(s) with {1} app report(s) in the selected workspaces." -f $apps.Count, $appReports.Count)

    return [ordered]@{
        CollectedUtc            = [DateTime]::UtcNow.ToString('o')
        Connections             = $connections.ToArray()
        Gateways                = $gateways.ToArray()
        Capacities              = $capacities.ToArray()
        Apps                    = $apps.ToArray()
        AppReports              = $appReports.ToArray()
        AppReportIds            = $appReportIds.ToArray()
        OriginalReportObjectIds = $originalReportObjectIds.ToArray()
        Errors                  = $errors.ToArray()
    }
}

# =====================================================================================================================
# Post-processing and the stage body
# =====================================================================================================================

function Update-IQInventoryUnknownName {
    <#
    .SYNOPSIS
        Back-fills 'Unknown Dataset' / 'Unknown Dataflow' names in every ws-*.json using the complete name lookups (the monolith depended on workspace order).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$DatasetNameLookup,
        [Parameter(Mandatory = $true)][hashtable]$DataflowNameLookup
    )
    $fixed = 0
    foreach ($inv in @(Get-IQAllWorkspaceInventories)) {
        $changed = $false
        foreach ($row in @(Get-IQMemberValue -Object $inv -Name 'Reports')) {
            if ($null -eq $row) { continue }
            $id = [string](Get-IQMemberValue -Object $row -Name 'DatasetId')
            if ((Get-IQMemberValue -Object $row -Name 'DatasetName') -eq 'Unknown Dataset' -and $id -and $DatasetNameLookup.ContainsKey($id)) {
                Add-IQNote -Row $row -Name 'DatasetName' -Value $DatasetNameLookup[$id]; $changed = $true; $fixed++
            }
        }
        foreach ($row in @(Get-IQMemberValue -Object $inv -Name 'DataflowLineage')) {
            if ($null -eq $row) { continue }
            $dsId = [string](Get-IQMemberValue -Object $row -Name 'DatasetId')
            $dfId = [string](Get-IQMemberValue -Object $row -Name 'DataflowId')
            if ((Get-IQMemberValue -Object $row -Name 'DatasetName') -eq 'Unknown Dataset' -and $dsId -and $DatasetNameLookup.ContainsKey($dsId)) {
                Add-IQNote -Row $row -Name 'DatasetName' -Value $DatasetNameLookup[$dsId]; $changed = $true; $fixed++
            }
            if ((Get-IQMemberValue -Object $row -Name 'DataflowName') -eq 'Unknown Dataflow' -and $dfId -and $DataflowNameLookup.ContainsKey($dfId)) {
                Add-IQNote -Row $row -Name 'DataflowName' -Value $DataflowNameLookup[$dfId]; $changed = $true; $fixed++
            }
        }
        foreach ($collection in @('DataflowSources', 'DataflowRefreshHistory')) {
            foreach ($row in @(Get-IQMemberValue -Object $inv -Name $collection)) {
                if ($null -eq $row) { continue }
                $dfId = [string](Get-IQMemberValue -Object $row -Name 'DataflowId')
                if ((Get-IQMemberValue -Object $row -Name 'DataflowName') -eq 'Unknown Dataflow' -and $dfId -and $DataflowNameLookup.ContainsKey($dfId)) {
                    Add-IQNote -Row $row -Name 'DataflowName' -Value $DataflowNameLookup[$dfId]; $changed = $true; $fixed++
                }
            }
        }
        if ($changed) {
            $name = 'ws-' + [string](Get-IQMemberValue -Object $inv -Name 'WorkspaceId')
            Save-IQInventory -Name $name -Object $inv | Out-Null
        }
    }
    if ($fixed -gt 0) { Write-IQLog -Level Debug -Stage Inventory -Message "Back-filled $fixed unknown dataset/dataflow name(s) across the inventory files." }
    return $fixed
}

function Get-IQInventoryNameLookup {
    <#
    .SYNOPSIS
        Seeds the DatasetId->name and DataflowId->name lookups from the inventory files already on disk (resume).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$DatasetNameLookup,
        [Parameter(Mandatory = $true)][hashtable]$DataflowNameLookup
    )
    $known = @{}
    foreach ($inv in @(Get-IQAllWorkspaceInventories)) {
        foreach ($row in @(Get-IQMemberValue -Object $inv -Name 'Datasets')) {
            if ($null -eq $row) { continue }
            $id = [string](Get-IQMemberValue -Object $row -Name 'DatasetId')
            if ($id) { $DatasetNameLookup[$id] = Get-IQMemberValue -Object $row -Name 'DatasetName' }
        }
        foreach ($row in @(Get-IQMemberValue -Object $inv -Name 'Dataflows')) {
            if ($null -eq $row) { continue }
            $id = [string](Get-IQMemberValue -Object $row -Name 'DataflowId')
            if ($id) { $DataflowNameLookup[$id] = Get-IQMemberValue -Object $row -Name 'DataflowName' }
        }
        if (-not [bool](Get-IQMemberValue -Object $inv -Name 'IsSynthetic')) {
            foreach ($row in @(Get-IQMemberValue -Object $inv -Name 'Reports')) {
                if ($null -eq $row) { continue }
                $id = [string](Get-IQMemberValue -Object $row -Name 'ReportId')
                if ($id) { $known[$id.ToLowerInvariant()] = $true }
            }
        }
    }
    return $known
}

function Invoke-IQInventoryStage {
    <#
    .SYNOPSIS
        Inventory stage body: resolves the scope, writes workspaces.json and global.json, then one checkpointed ws-<id>.json per workspace (+ My Workspace).
    .DESCRIPTION
        Resume: workspaces whose Inventory checkpoint exists (done\Inventory\<wsId>.json with the ws file still present)
        are skipped; the global block has its own checkpoint ('global'). Returns a summary hashtable
        (WorkspaceCount, Collected, Skipped, Failed). Throws when the scope cannot be resolved or the user cancels
        ($script:IQ.ScopeCancelled is set to $true in that case so the entry point can exit gracefully).
    #>
    [CmdletBinding()]
    param()
    $stage = 'Inventory'
    $scope = Resolve-IQScope
    if ($null -eq $scope) {
        $script:IQ.ScopeCancelled = $true
        throw 'Scope selection cancelled by user.'
    }
    $realWorkspaces = @($scope.Workspaces)

    # ---- workspaces.json (initial; re-written after capacities and My Workspace are known) ----
    $wsRows = New-Object System.Collections.Generic.List[object]
    foreach ($w in $realWorkspaces) {
        if (-not $w.PSObject.Properties['WorkspaceIsSynthetic']) { Add-IQNote -Row $w -Name 'WorkspaceIsSynthetic' -Value $false }
        $wsRows.Add($w)
    }
    if ($scope.IncludeMyWorkspace) { $wsRows.Add((New-IQPseudoWorkspace -Name 'My Workspace')) }
    $workspacesPath = Save-IQInventory -Name 'workspaces' -Object $wsRows.ToArray()

    # ---- global block (checkpointed as one item) ----
    $globalInv = $null
    if (Test-IQItemDone -Stage $stage -ItemKey 'global') {
        Write-IQLog -Level Info -Stage $stage -Message 'Global inventory (connections, gateways, capacities, apps) already collected - skipping.'
        $globalInv = Get-IQInventory -Name 'global'
    }
    if ($null -eq $globalInv) {
        try {
            $globalInv = Get-IQGlobalInventory -Workspaces $realWorkspaces
            $globalPath = Save-IQInventory -Name 'global' -Object $globalInv
            $status = 'Succeeded'
            $message = $null
            if (@($globalInv.Errors).Count -gt 0) {
                $status = 'Failed'
                $message = (@($globalInv.Errors) | ForEach-Object { $_.Collector + ': ' + $_.Message }) -join ' | '
            }
            Set-IQItemDone -Stage $stage -ItemKey 'global' -Item 'Global inventory' -Outputs @($globalPath) -Status $status -Message $message -Data @{
                Connections = @($globalInv.Connections).Count; Gateways = @($globalInv.Gateways).Count; Capacities = @($globalInv.Capacities).Count
                Apps = @($globalInv.Apps).Count; AppReports = @($globalInv.AppReports).Count
            } | Out-Null
        }
        catch {
            Set-IQItemDone -Stage $stage -ItemKey 'global' -Item 'Global inventory' -Status Failed -Message $_.Exception.Message | Out-Null
            Write-IQLog -Level Error -Stage $stage -Message 'Global inventory failed' -Exception $_.Exception
            $globalInv = [ordered]@{ Connections = @(); Gateways = @(); Capacities = @(); Apps = @(); AppReports = @(); AppReportIds = @(); OriginalReportObjectIds = @(); Errors = @() }
        }
    }

    # CapacityName joined from Capacities (brief 6.2)
    $capacityNames = @{}
    foreach ($c in @(Get-IQMemberValue -Object $globalInv -Name 'Capacities')) {
        if ($null -eq $c) { continue }
        $cid = [string](Get-IQMemberValue -Object $c -Name 'CapacityId')
        if ($cid) { $capacityNames[$cid.ToLowerInvariant()] = Get-IQMemberValue -Object $c -Name 'CapacityDisplayName' }
    }
    foreach ($w in $wsRows) {
        $capId = [string](Get-IQMemberValue -Object $w -Name 'WorkspaceCapacityId')
        $capName = $null
        if ($capId -and $capacityNames.ContainsKey($capId.ToLowerInvariant())) { $capName = $capacityNames[$capId.ToLowerInvariant()] }
        Add-IQNote -Row $w -Name 'CapacityName' -Value $capName
    }
    $workspacesPath = Save-IQInventory -Name 'workspaces' -Object $wsRows.ToArray()

    # ---- per-workspace collection ----
    $datasetNameLookup = @{}
    $dataflowNameLookup = @{}
    $knownReportIds = Get-IQInventoryNameLookup -DatasetNameLookup $datasetNameLookup -DataflowNameLookup $dataflowNameLookup
    foreach ($id in @(Get-IQMemberValue -Object $globalInv -Name 'AppReportIds')) { if ($id) { $knownReportIds[([string]$id).ToLowerInvariant()] = $true } }
    foreach ($id in @(Get-IQMemberValue -Object $globalInv -Name 'OriginalReportObjectIds')) { if ($id) { $knownReportIds[([string]$id).ToLowerInvariant()] = $true } }

    $total = $realWorkspaces.Count
    $index = 0
    $collected = 0
    $skipped = 0
    $failed = 0
    $budgetStop = $false
    foreach ($ws in $realWorkspaces) {
        $index++
        $wsId = [string]$ws.WorkspaceId
        $wsName = [string]$ws.WorkspaceName
        if (Test-IQItemDone -Stage $stage -ItemKey $wsId) {
            Write-IQLog -Level Info -Stage $stage -Item $wsName -Message ("Workspace {0}/{1} already inventoried - skipping (resume)." -f $index, $total)
            $skipped++
            continue
        }
        if (Test-IQTimeBudget -Stage $stage -Item $wsName) {
            Write-IQLog -Level Warn -Stage $stage -Message ("Time budget reached: {0} of {1} workspace(s) not inventoried yet - they are collected on the next start." -f ($total - $index + 1), $total)
            $budgetStop = $true
            break
        }
        Write-IQLog -Level Info -Stage $stage -Item $wsName -Message ("Workspace {0}/{1} ({2})" -f $index, $total, $wsId)
        try {
            $inv = Get-IQWorkspaceInventory -Workspace $ws -Scope $scope -DatasetNameLookup $datasetNameLookup -DataflowNameLookup $dataflowNameLookup
            $path = Save-IQInventory -Name ('ws-' + $wsId) -Object $inv
            foreach ($r in @($inv.Reports)) { if ($r.ReportId) { $knownReportIds[([string]$r.ReportId).ToLowerInvariant()] = $true } }
            $counts = @{
                Datasets = @($inv.Datasets).Count; Reports = @($inv.Reports).Count; ReportPages = @($inv.ReportPages).Count
                Dataflows = @($inv.Dataflows).Count; FabricItems = @($inv.FabricItems).Count; Dashboards = @($inv.Dashboards).Count
                WorkspaceUsers = @($inv.WorkspaceUsers).Count; Errors = @($inv.Errors).Count
            }
            if (@($inv.Errors).Count -gt 0) {
                $message = (@($inv.Errors) | ForEach-Object { $_.Collector + ': ' + $_.Message }) -join ' | '
                Set-IQItemDone -Stage $stage -ItemKey $wsId -Item $wsName -Outputs @($path) -Status Failed -Message ("Partial inventory - " + $message) -Data $counts | Out-Null
                $failed++
            }
            else {
                Set-IQItemDone -Stage $stage -ItemKey $wsId -Item $wsName -Outputs @($path) -Status Succeeded -Data $counts | Out-Null
                $collected++
            }
        }
        catch {
            Set-IQItemDone -Stage $stage -ItemKey $wsId -Item $wsName -Status Failed -Message $_.Exception.Message | Out-Null
            Write-IQLog -Level Error -Stage $stage -Item $wsName -Message 'Workspace inventory failed' -Exception $_.Exception
            $failed++
        }
    }

    # ---- My Workspace ----
    $sharedFound = $false
    if ($scope.IncludeMyWorkspace -and $budgetStop) {
        Write-IQLog -Level Warn -Stage $stage -Item 'My Workspace' -Message 'Time budget reached - My Workspace is inventoried on the next start.'
    }
    elseif ($scope.IncludeMyWorkspace) {
        $myKey = 'My Workspace'
        if (Test-IQItemDone -Stage $stage -ItemKey $myKey) {
            Write-IQLog -Level Info -Stage $stage -Item $myKey -Message 'My Workspace already inventoried - skipping (resume).'
            $skipped++
            $existing = Get-IQInventory -Name ('ws-' + $myKey)
            if ($null -ne $existing) { $sharedFound = [bool](Get-IQMemberValue -Object $existing -Name 'SharedReportsFound') }
        }
        else {
            try {
                $inv = Get-IQMyWorkspaceInventory -Scope $scope -KnownReportIds $knownReportIds -DatasetNameLookup $datasetNameLookup
                $path = Save-IQInventory -Name ('ws-' + $myKey) -Object $inv
                $sharedFound = [bool]$inv.SharedReportsFound
                $counts = @{ Datasets = @($inv.Datasets).Count; Reports = @($inv.Reports).Count; ReportPages = @($inv.ReportPages).Count; Dashboards = @($inv.Dashboards).Count; Errors = @($inv.Errors).Count }
                if (@($inv.Errors).Count -gt 0) {
                    $message = (@($inv.Errors) | ForEach-Object { $_.Collector + ': ' + $_.Message }) -join ' | '
                    Set-IQItemDone -Stage $stage -ItemKey $myKey -Item $myKey -Outputs @($path) -Status Failed -Message ("Partial inventory - " + $message) -Data $counts | Out-Null
                    $failed++
                }
                else {
                    Set-IQItemDone -Stage $stage -ItemKey $myKey -Item $myKey -Outputs @($path) -Status Succeeded -Data $counts | Out-Null
                    $collected++
                }
            }
            catch {
                Set-IQItemDone -Stage $stage -ItemKey $myKey -Item $myKey -Status Failed -Message $_.Exception.Message | Out-Null
                Write-IQLog -Level Error -Stage $stage -Item $myKey -Message 'My Workspace inventory failed' -Exception $_.Exception
                $failed++
            }
        }
    }
    else {
        Write-IQLog -Level Info -Stage $stage -Message "Skipping 'My Workspace' (not selected)."
    }

    # ---- post-processing: unknown names, final workspaces.json (with the Shared Reports pseudo row when needed) ----
    Update-IQInventoryUnknownName -DatasetNameLookup $datasetNameLookup -DataflowNameLookup $dataflowNameLookup | Out-Null
    if ($sharedFound) {
        $row = New-IQPseudoWorkspace -Name 'Shared Reports (No Workspace Access)'
        Add-IQNote -Row $row -Name 'CapacityName' -Value $null
        $wsRows.Add($row)
    }
    $workspacesPath = Save-IQInventory -Name 'workspaces' -Object $wsRows.ToArray()

    $summary = @{ WorkspaceCount = $wsRows.Count; Collected = $collected; Skipped = $skipped; Failed = $failed; WorkspacesFile = $workspacesPath; BudgetStop = $budgetStop }
    Write-IQLog -Level Success -Stage $stage -Message ("Inventory complete: {0} workspace(s) in scope, {1} collected, {2} skipped (resume), {3} failed." -f $wsRows.Count, $collected, $skipped, $failed)
    return $summary
}

# =====================================================================================================================
# Accessors for later stages
# =====================================================================================================================

function Get-IQSelectedWorkspaces {
    <#
    .SYNOPSIS
        The workspaces of the run (renamed rows from inventory\workspaces.json, incl. the pseudo rows unless -ExcludeSynthetic); falls back to the resolved scope.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][switch]$ExcludeSynthetic)
    $rows = @()
    try { $rows = @(Get-IQInventory -Name 'workspaces') } catch { $rows = @() }
    $rows = @($rows | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) {
        $scope = Get-IQEffectiveScope
        if ($null -ne $scope) {
            $rows = @($scope.Workspaces)
            if ($scope.IncludeMyWorkspace) { $rows += New-IQPseudoWorkspace -Name 'My Workspace' }
        }
    }
    if ($ExcludeSynthetic) { $rows = @($rows | Where-Object { -not [bool](Get-IQMemberValue -Object $_ -Name 'WorkspaceIsSynthetic') }) }
    return $rows
}

function Get-IQSelectedItemList {
    <#
    .SYNOPSIS
        Flattens one collection ('Datasets' or 'Reports') from every ws-*.json, applies the scope filter and enriches rows with the workspace's capacity/synthetic flags (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Datasets', 'Reports')][string]$Collection,
        [Parameter(Mandatory = $false)][switch]$IncludeSharedReports
    )
    $scope = Get-IQEffectiveScope
    $wsById = @{}
    foreach ($w in @(Get-IQSelectedWorkspaces)) {
        $id = [string](Get-IQMemberValue -Object $w -Name 'WorkspaceId')
        if ($id) { $wsById[$id.ToLowerInvariant()] = $w }
    }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($inv in @(Get-IQAllWorkspaceInventories)) {
        foreach ($row in @(Get-IQMemberValue -Object $inv -Name $Collection)) {
            if ($null -eq $row) { continue }
            $wsId = [string](Get-IQMemberValue -Object $row -Name 'WorkspaceId')
            if (-not $IncludeSharedReports -and $wsId -eq 'Shared Reports (No Workspace Access)') { continue }
            if ($Collection -eq 'Datasets') {
                if (-not (Test-IQScopeIncludesDataset -Scope $scope -DatasetId ([string](Get-IQMemberValue -Object $row -Name 'DatasetId')))) { continue }
            }
            else {
                if (-not (Test-IQScopeIncludesReport -Scope $scope -ReportId ([string](Get-IQMemberValue -Object $row -Name 'ReportId')))) { continue }
            }
            $ws = $null
            if ($wsId -and $wsById.ContainsKey($wsId.ToLowerInvariant())) { $ws = $wsById[$wsId.ToLowerInvariant()] }
            $dedicated = $false
            $synthetic = $false
            if ($null -ne $ws) {
                $dedicated = [bool](Get-IQMemberValue -Object $ws -Name 'WorkspaceIsOnDedicatedCapacity')
                $synthetic = [bool](Get-IQMemberValue -Object $ws -Name 'WorkspaceIsSynthetic')
            }
            elseif ($wsId -eq 'My Workspace' -or $wsId -eq 'Shared Reports (No Workspace Access)') { $synthetic = $true }
            if (-not $row.PSObject.Properties['WorkspaceIsOnDedicatedCapacity']) { Add-IQNote -Row $row -Name 'WorkspaceIsOnDedicatedCapacity' -Value $dedicated }
            if (-not $row.PSObject.Properties['WorkspaceIsSynthetic']) { Add-IQNote -Row $row -Name 'WorkspaceIsSynthetic' -Value $synthetic }
            $out.Add($row)
        }
    }
    return $out.ToArray()
}

function Get-IQSelectedDatasets {
    <#
    .SYNOPSIS
        Every dataset row in scope (monolith $datasetsInfo shape: DatasetId, DatasetName, WorkspaceId, WorkspaceName, ... plus WorkspaceIsOnDedicatedCapacity/WorkspaceIsSynthetic) from the inventory files.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][switch]$IncludeSharedReports)
    return @(Get-IQSelectedItemList -Collection Datasets -IncludeSharedReports:$IncludeSharedReports)
}

function Get-IQSelectedReports {
    <#
    .SYNOPSIS
        Every report row in scope (monolith $reportsInfo shape: ReportId, ReportName, DatasetId, DatasetWorkspaceId, WorkspaceId, WorkspaceName, ReportHasSensitivityLabel, ...); 'Shared Reports (No Workspace Access)' rows only with -IncludeSharedReports.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][switch]$IncludeSharedReports)
    return @(Get-IQSelectedItemList -Collection Reports -IncludeSharedReports:$IncludeSharedReports)
}

function Get-IQReportsWithSensitivityLabel {
    <#
    .SYNOPSIS
        Hashtable (report id -> $true) of reports that carry a sensitivity label per the Fabric items API (monolith $globalInv:ReportsWithSensitivityLabel), read from the inventory files.
    #>
    [CmdletBinding()]
    param()
    $set = @{}
    foreach ($inv in @(Get-IQAllWorkspaceInventories)) {
        foreach ($id in @(Get-IQMemberValue -Object $inv -Name 'ReportsWithSensitivityLabel')) {
            if ($null -ne $id -and [string]$id) { $set[[string]$id] = $true }
        }
    }
    return $set
}
