#Requires -Version 5.1
<#
    ImpactIQ v3 - Assemble stage (ImpactIQ.Assemble.ps1)

    Rebuilds the four workbooks from the run state and the backup folders, every run (brief sections 2.7, 6.4, 9, 14):

      Power BI Environment Detail.xlsx   <- State\runs\<RunId>\inventory\workspaces.json, global.json, ws-*.json,
                                            the manifest (RunSummary, Failures), inventory\extras-*.json (Extras stage)
      Report Detail.xlsx                 <- every *.txt in Report Backups\<RunId>\ (monolith 3063-3110 parsing:
                                            tab-split with the exact header count, UTF-8, sheet name = file name)
      Model Detail.xlsx                  <- *.csv in Model Backups\<RunId>\ ("Semantic Models" = every CSV that is not
                                            *_MD.csv, "Measure Dependencies" = *_MD.csv; monolith 3209-3290)
      Dataflow Detail.xlsx               <- State\runs\<RunId>\extracts\dataflows\*.json (Sheet1: the six monolith
                                            columns + the five DataTable artefact columns the PBIT expects), written to
                                            Dataflow Backups\<RunId>\ and copied to the BaseFolder (monolith 3701-3728)

    Rules applied (audit findings C6-13/14/15/16, C8-04/05/06/09/11/15, C9-05/11):
      - every sheet listed in Config\SheetContract.json exists even when empty; missing contract columns are added as
        empty cells (Ensure-IQSheetColumns); an empty collection becomes a header row plus ONE row of empty strings
        (the monolith's Dataflow "dummy row" behaviour, generalised - the PBIT filters nulls/empties). NOTE: EPPlus
        does not store empty-string cells, so on disk such a sheet is header-only (row 2 is an empty <row> element);
        Import-Excel returns zero rows for it - read the header with Import-Excel -NoHeader or Open-ExcelPackage;
      - real DateTime values (PowerShell 7 parses ISO date-times in JSON into DateTime, 5.1 keeps ISO text) are
        written as Excel dates (built-in date-time format, so Import-Excel/EPPlus read them back as [datetime]);
        -NoNumberConversion * is always passed;
      - the column set of a sheet is the UNION of every row's properties (first-appearance order), so heterogeneous
        rows no longer lose properties;
      - values go through a typed System.Data.DataTable and Export-Excel -InputObject (EPPlus LoadFromDataTable):
        strings are written verbatim (IDs with leading zeros stay text, "=..." is never turned into a formula, nothing
        is number-converted), real numbers/booleans/dates keep their type, nested values become compact JSON, and
        cells longer than Excel's 32,767-character limit are truncated (counted and logged); a sheet never exceeds
        Excel's 1,048,576-row limit (rows beyond 1,048,575 data rows are dropped, counted and logged);
      - each workbook is written to a temporary file in the target folder and copied over the previous workbook only
        after it was closed successfully, so a failed build never destroys the last good output (stale *.tmp-*.xlsx
        files left by a killed process are removed first; .NET file APIs are used so '[' / ']' in the BaseFolder
        never glob);
      - one sheet that cannot be written (EPPlus error) does not lose the workbook: the sheet is replaced by its
        header-only table (every contract sheet still exists), the failure is logged and the workbook item is
        recorded as Failed; a source file that cannot be read (CSV, txt, dataflow extract, extras JSON) is recorded
        as a failed Assemble item ("source-<file>") so manifest.failures, the Failures sheet and the exit code show
        that the workbook is incomplete (brief section 0.4);
      - Message / Error text copied from the manifest and the inventory Errors[] goes through ConvertTo-IQRedactedText
        (the workbooks are shared; connection-string passwords and tokens must not land in them);
      - -AutoNameRange is kept where the monolith used it (Report/Model Detail); -AutoSize is used on Windows only
        (ImportExcel cannot auto-fit without libgdiplus on Linux) and is capped by ImportExcel's MaxAutoSizeRows;
      - worksheet names are made Excel-safe (31 characters, no []:*?/\); "DatasetDirectQueryRefreshSchedule" (33
        characters) is written as "DatasetDQRefreshSchedule".

    Windows PowerShell 5.1 and PowerShell 7 compatible. Dot-sourced from ImpactIQ.ps1, so $script:IQ is the shared
    context. Cross-module functions used (brief section 2): Write-IQLog, Get-IQSafeKey, ConvertFrom-IQJsonFile,
    Get-IQDateFolder, Get-IQInventory, Get-IQAllWorkspaceInventories, Set-IQItemDone, Save-IQManifest,
    ConvertTo-IQRedactedText.
    Private helpers are prefixed *-IQAsm* / *-IQSheet* and are not part of the cross-module contract.
#>

# Sheet order of "Power BI Environment Detail.xlsx": the monolith's 17 sheets (lines 2521-2537) first, in the same
# order, then the additive sheets of brief section 6.4. Extras sheets (section 8.4) follow in file order.
$script:IQEnvironmentSheetOrder = @(
    'Workspaces', 'FabricItems', 'Connections', 'Gateways', 'ItemConnections', 'Datasets', 'DatasetSourcesInfo',
    'DatasetRefreshHistory', 'DatasetRefreshSchedule', 'Dataflows', 'DataflowLineage', 'DataflowSourcesInfo',
    'DataflowRefreshHistory', 'Reports', 'ReportPages', 'Apps', 'AppReports',
    'Dashboards', 'DashboardTiles', 'Capacities', 'WorkspaceUsers', 'DatasetUsers', 'DatasetParameters',
    'DatasetDirectQueryRefreshSchedule', 'RunSummary', 'Failures', 'InventoryErrors'
)

# Where each environment sheet comes from: 'workspaces' = inventory\workspaces.json, 'global' = a key of global.json,
# 'ws' = a key of every ws-*.json (concatenated), 'manifest' = built from the manifest.
$script:IQEnvironmentSheetSource = @{
    'Workspaces'                        = @{ Source = 'workspaces'; Key = $null }
    'FabricItems'                       = @{ Source = 'ws'; Key = 'FabricItems' }
    'Connections'                       = @{ Source = 'global'; Key = 'Connections' }
    'Gateways'                          = @{ Source = 'global'; Key = 'Gateways' }
    'ItemConnections'                   = @{ Source = 'ws'; Key = 'ItemConnections' }
    'Datasets'                          = @{ Source = 'ws'; Key = 'Datasets' }
    'DatasetSourcesInfo'                = @{ Source = 'ws'; Key = 'DatasetSources' }
    'DatasetRefreshHistory'             = @{ Source = 'ws'; Key = 'DatasetRefreshHistory' }
    'DatasetRefreshSchedule'            = @{ Source = 'ws'; Key = 'DatasetRefreshSchedule' }
    'Dataflows'                         = @{ Source = 'ws'; Key = 'Dataflows' }
    'DataflowLineage'                   = @{ Source = 'ws'; Key = 'DataflowLineage' }
    'DataflowSourcesInfo'               = @{ Source = 'ws'; Key = 'DataflowSources' }
    'DataflowRefreshHistory'            = @{ Source = 'ws'; Key = 'DataflowRefreshHistory' }
    'Reports'                           = @{ Source = 'ws'; Key = 'Reports' }
    'ReportPages'                       = @{ Source = 'ws'; Key = 'ReportPages' }
    'Apps'                              = @{ Source = 'global'; Key = 'Apps' }
    'AppReports'                        = @{ Source = 'global'; Key = 'AppReports' }
    'Dashboards'                        = @{ Source = 'ws'; Key = 'Dashboards' }
    'DashboardTiles'                    = @{ Source = 'ws'; Key = 'DashboardTiles' }
    'Capacities'                        = @{ Source = 'global'; Key = 'Capacities' }
    'WorkspaceUsers'                    = @{ Source = 'ws'; Key = 'WorkspaceUsers' }
    'DatasetUsers'                      = @{ Source = 'ws'; Key = 'DatasetUsers' }
    'DatasetParameters'                 = @{ Source = 'ws'; Key = 'DatasetParameters' }
    'DatasetDirectQueryRefreshSchedule' = @{ Source = 'ws'; Key = 'DatasetDirectQueryRefreshSchedule' }
    'RunSummary'                        = @{ Source = 'manifest'; Key = 'RunSummary' }
    'Failures'                          = @{ Source = 'manifest'; Key = 'Failures' }
    'InventoryErrors'                   = @{ Source = 'manifest'; Key = 'InventoryErrors' }
}

# Default columns of the sheets that are not in SheetContract.json (brief sections 6.1, 6.2, 6.4), so that an empty
# collection still produces a meaningful header row.
$script:IQAssembleDefaultColumns = @{
    'Dashboards'                        = @('DashboardId', 'DashboardName', 'DashboardIsReadOnly', 'DashboardWebUrl', 'DashboardEmbedUrl', 'WorkspaceId', 'WorkspaceName')
    'DashboardTiles'                    = @('TileId', 'TileTitle', 'TileSubTitle', 'TileRowSpan', 'TileColSpan', 'TileEmbedUrl', 'ReportId', 'DatasetId', 'DashboardId', 'DashboardName', 'WorkspaceId', 'WorkspaceName')
    'Capacities'                        = @('CapacityId', 'CapacityDisplayName', 'CapacitySku', 'CapacityState', 'CapacityRegion', 'CapacityAdmins', 'CapacityUsersAccessRight')
    'WorkspaceUsers'                    = @('UserEmailAddress', 'UserDisplayName', 'UserIdentifier', 'UserPrincipalType', 'UserGroupUserAccessRight', 'UserGraphId', 'WorkspaceId', 'WorkspaceName')
    'DatasetUsers'                      = @('UserIdentifier', 'UserPrincipalType', 'UserDatasetUserAccessRight', 'UserDisplayName', 'UserEmailAddress', 'DatasetId', 'DatasetName', 'WorkspaceId', 'WorkspaceName')
    'DatasetParameters'                 = @('ParameterName', 'ParameterType', 'ParameterIsRequired', 'ParameterCurrentValue', 'DatasetId', 'DatasetName', 'WorkspaceId', 'WorkspaceName')
    'DatasetDirectQueryRefreshSchedule' = @('DQFrequency', 'DQLocalTimeZoneId', 'DQDay', 'DQTime', 'DatasetId', 'DatasetName', 'WorkspaceId', 'WorkspaceName')
    'RunSummary'                        = @('RunId', 'Stage', 'Status', 'StartedUtc', 'EndedUtc', 'DurationSeconds', 'ItemsDone', 'ItemsFailed', 'Error', 'Environment', 'Auth', 'RunMode', 'Machine', 'User', 'PSVersion', 'IsAzureDevOps', 'ResumeCount')
    'Failures'                          = @('Stage', 'ItemKey', 'Item', 'Message', 'TimeUtc', 'RunId')
    'InventoryErrors'                   = @('WorkspaceId', 'WorkspaceName', 'Collector', 'Path', 'Message')
    'Gateways'                          = @('GatewayId', 'GatewayName', 'GatewayType', 'GatewayDisplayName')
    'Connections'                       = @('ConnectionId', 'ConnectionName', 'ConnectionType', 'ConnectionDisplayName')
    'ItemConnections'                   = @('FabricItemID', 'FabricItemName', 'FabricItemType', 'ConnectionId', 'WorkspaceId', 'WorkspaceName')
}

# The monolith exported a System.Data.DataTable through the pipeline, which added these DataRow properties as columns;
# the PBIT's ExpectedColumns list for Sheet1 therefore includes them (brief section 14). Written as empty strings.
$script:IQDataflowSheetColumns = @('Dataflow ID', 'Dataflow Name', 'Query Name', 'Query', 'Report Date', 'Workspace Name - Dataflow Name')
$script:IQDataflowArtefactColumns = @('RowError', 'RowState', 'Table', 'ItemArray', 'HasErrors')

# Excel-safe worksheet names for sheet names longer than 31 characters (Excel hard limit).
$script:IQSheetNameAlias = @{
    'DatasetDirectQueryRefreshSchedule' = 'DatasetDQRefreshSchedule'
}

# Canonical Extras sheet names (brief section 8.4); inventory file names are lower-cased by Get-IQSafeKey.
$script:IQExtrasSheetNames = @(
    'AdminWorkspaces', 'AdminWorkspaceUsers', 'ScanDatasets', 'ScanTables', 'ScanColumns', 'ScanMeasures',
    'ScanDatasources', 'ScanReports', 'ScanDashboards', 'ScanDataflows', 'ScanUsers', 'ActivityEvents',
    'UsageReportViews', 'UsageReportPageViews'
)

$script:IQExcelCellLimit = 32767
# Excel's hard limit is 1,048,576 rows per worksheet; row 1 is the header, so at most 1,048,575 data rows are written.
$script:IQExcelMaxDataRows = 1048575

# Source files that could not be read while the current workbook was built (@{ Path; Message }); reset per workbook by
# Invoke-IQAssembleStage and turned into failed "source-<file>" Assemble items so the run does not end 'Completed'
# while a sheet silently lacks rows (brief section 0.4).
$script:IQAssembleSourceErrors = New-Object System.Collections.Generic.List[object]

# =====================================================================================================================
# Small private helpers (rows may be PSCustomObjects from JSON, hashtables/ordered dictionaries, or DataRows)
# =====================================================================================================================

function Add-IQAsmSourceError {
    <#
    .SYNOPSIS
        Logs a Warn for a source file that could not be read and remembers it for Invoke-IQAssembleStage (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Message,
        [Parameter(Mandatory = $false)][AllowNull()][System.Exception]$Exception,
        [Parameter(Mandatory = $false)][AllowNull()][string]$Item
    )
    if ([string]::IsNullOrWhiteSpace($Item)) { $Item = [System.IO.Path]::GetFileName($Path) }
    Write-IQLog -Level Warn -Stage 'Assemble' -Item $Item -Message $Message -Exception $Exception
    if ($null -eq $script:IQAssembleSourceErrors) { $script:IQAssembleSourceErrors = New-Object System.Collections.Generic.List[object] }
    $script:IQAssembleSourceErrors.Add(@{ Path = $Path; Message = $Message })
}

function Remove-IQAsmStaleTempFile {
    <#
    .SYNOPSIS
        Deletes "<BaseName>.tmp-*.xlsx" files a killed earlier run left next to a workbook (private; never throws).
    .DESCRIPTION
        Uses the .NET directory API (no PowerShell wildcard globbing of the folder path) so a BaseFolder containing
        '[' or ']' works; returns the number of files removed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Folder,
        [Parameter(Mandatory = $true)][string]$BaseName
    )
    $removed = 0
    if ([string]::IsNullOrWhiteSpace($Folder) -or -not [System.IO.Directory]::Exists($Folder)) { return $removed }
    $stale = @()
    try { $stale = @([System.IO.Directory]::GetFiles($Folder, $BaseName + '.tmp-*.xlsx')) } catch { $stale = @() }
    foreach ($f in $stale) {
        try { [System.IO.File]::Delete($f); $removed++ }
        catch { Write-IQLog -Level Debug -Stage 'Assemble' -Message ("Stale temporary workbook '{0}' could not be removed: {1}" -f $f, $_.Exception.Message) }
    }
    if ($removed -gt 0) {
        Write-IQLog -Level Debug -Stage 'Assemble' -Item $BaseName -Message ("Removed {0} stale temporary workbook file(s) left by an earlier run." -f $removed)
    }
    return $removed
}

function Get-IQAsmMember {
    <#
    .SYNOPSIS
        Reads a named member from a dictionary or an object property; $null when absent (private).
    .DESCRIPTION
        Standard PowerShell output semantics apply: an array member is emitted element by element, so wrap the call
        in @( ) when a list is expected. Use Test-IQAsmArrayMember when the question is "is this member an array".
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

function Test-IQAsmHasMember {
    <#
    .SYNOPSIS
        $true when the dictionary key / object property exists (even with a $null or empty value) (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )
    if ($null -eq $Object) { return $false }
    if ($Object -is [System.Collections.IDictionary]) { return [bool]$Object.Contains($Name) }
    if ($Object -is [string] -or $Object -is [System.ValueType]) { return $false }
    return ($null -ne $Object.PSObject.Properties[$Name])
}

function Test-IQAsmArrayMember {
    <#
    .SYNOPSIS
        $true when the named member holds an array/list (also when it is empty); avoids the pipeline unrolling of Get-IQAsmMember (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )
    if (-not (Test-IQAsmHasMember -Object $Object -Name $Name)) { return $false }
    $value = $null
    if ($Object -is [System.Collections.IDictionary]) { $value = $Object[$Name] }
    else { $value = $Object.PSObject.Properties[$Name].Value }
    if ($null -eq $value) { return $false }
    if ($value -is [string]) { return $false }
    return (($value -is [System.Array]) -or ($value -is [System.Collections.IList]))
}

function Get-IQAsmMemberName {
    <#
    .SYNOPSIS
        The member names of a dictionary or object, in declaration order; wrap in @( ) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Object)
    $names = @()
    if ($null -eq $Object) { return $names }
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($k in @($Object.Keys)) { $names += [string]$k }
        return $names
    }
    if ($Object -is [string] -or $Object -is [System.ValueType]) { return $names }
    foreach ($p in $Object.PSObject.Properties) { $names += $p.Name }
    return $names
}

function ConvertTo-IQAsmRowMap {
    <#
    .SYNOPSIS
        Normalises one row (object / dictionary) into an ordered dictionary of name -> value (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Row)
    $map = [ordered]@{}
    if ($null -eq $Row) { return $map }
    if ($Row -is [System.Collections.IDictionary]) {
        foreach ($k in @($Row.Keys)) { $map[[string]$k] = $Row[$k] }
        return $map
    }
    if ($Row -is [string] -or $Row -is [System.ValueType]) {
        $map['Value'] = $Row
        return $map
    }
    foreach ($p in $Row.PSObject.Properties) {
        if ($p.MemberType -in @('NoteProperty', 'Property', 'AliasProperty', 'ScriptProperty')) { $map[$p.Name] = $p.Value }
    }
    return $map
}

function ConvertTo-IQAsmCellText {
    <#
    .SYNOPSIS
        Converts any value to the text written into a string-typed sheet column (invariant culture, nested -> compact JSON) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Value)
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return '' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [bool]) { if ($Value) { return 'True' } else { return 'False' } }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [System.DateTimeOffset]) { return $Value.ToString('yyyy-MM-dd HH:mm:ss zzz', [System.Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [System.IFormattable]) { return $Value.ToString($null, [System.Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [System.Collections.IDictionary] -or $Value -is [System.Collections.IEnumerable] -or $Value -is [System.Management.Automation.PSCustomObject]) {
        try { return (ConvertTo-Json -InputObject $Value -Compress -Depth 20) } catch { return [string]$Value }
    }
    return [string]$Value
}

function Get-IQAsmValueKind {
    <#
    .SYNOPSIS
        Classifies a value for column typing: Null, Bool, DateTime, Integer, Real, String or Other (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Value)
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return 'Null' }
    if ($Value -is [string]) { return 'String' }
    if ($Value -is [bool]) { return 'Bool' }
    if ($Value -is [datetime]) { return 'DateTime' }
    if ($Value -is [byte] -or $Value -is [sbyte] -or $Value -is [int16] -or $Value -is [uint16] -or $Value -is [int32] -or $Value -is [uint32] -or $Value -is [int64]) { return 'Integer' }
    if ($Value -is [single] -or $Value -is [double] -or $Value -is [decimal]) { return 'Real' }
    return 'Other'
}

function Get-IQSafeSheetName {
    <#
    .SYNOPSIS
        Excel-safe worksheet name: alias map, forbidden characters replaced by "_", trimmed to 31 characters (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Name)
    $safe = $Name
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'Sheet' }
    if ($script:IQSheetNameAlias.ContainsKey($safe)) { $safe = $script:IQSheetNameAlias[$safe] }
    $safe = $safe -replace '[\[\]:\*\?/\\]', '_'
    $safe = $safe.Trim("'")
    if ($safe.Length -gt 31) { $safe = $safe.Substring(0, 31) }
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'Sheet' }
    return $safe
}

function Get-IQAsmRunPath {
    <#
    .SYNOPSIS
        Full path of a sub-folder of the current run state (inventory, extracts, ...), or $null without an active run (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$SubFolder)
    if (-not $script:IQ) { return $null }
    if ($script:IQ.ContainsKey('RunPaths') -and $null -ne $script:IQ.RunPaths) {
        $known = $script:IQ.RunPaths[$SubFolder]
        if (-not [string]::IsNullOrWhiteSpace([string]$known)) { return [string]$known }
    }
    if ([string]::IsNullOrWhiteSpace([string]$script:IQ.RunPath)) { return $null }
    return (Join-Path $script:IQ.RunPath $SubFolder.ToLowerInvariant())
}

function Get-IQAssembleFolder {
    <#
    .SYNOPSIS
        The backup folder Assemble reads for Model / Report / Dataflow: <root>\<RunId> when it exists; without an active run the newest yyyy-MM-dd folder (Warn); else $null.
    .DESCRIPTION
        With an active run (Initialize-IQRun set $script:IQ.RunPaths) a missing run folder yields $null and a Warn -
        the builders then emit header-only sheets. Falling back to the newest dated folder of ANOTHER run would present
        that run's CSV/txt rows as this run's output (and write this run's Dataflow workbook into that folder), so the
        monolith's "latest dated folder" rule only applies when no run is active (e.g. the helpers are called directly).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateSet('Model', 'Report', 'Dataflow')][string]$Kind)
    $root = $null
    if ($script:IQ -and $script:IQ.Paths) { $root = [string]$script:IQ.Paths[$Kind + 'Backups'] }
    $hasRun = [bool]($script:IQ -and $script:IQ.ContainsKey('RunPaths') -and $null -ne $script:IQ.RunPaths)
    $runFolder = $null
    if ($hasRun) { $runFolder = [string]$script:IQ.RunPaths[$Kind + 'Backups'] }
    if ([string]::IsNullOrWhiteSpace($runFolder) -and -not [string]::IsNullOrWhiteSpace($root) -and $script:IQ -and -not [string]::IsNullOrWhiteSpace([string]$script:IQ.RunId)) {
        $runFolder = Join-Path $root $script:IQ.RunId
    }
    if (-not [string]::IsNullOrWhiteSpace($runFolder) -and (Test-Path -LiteralPath $runFolder)) { return $runFolder }
    if ($hasRun) {
        Write-IQLog -Level Warn -Stage 'Assemble' -Message ("{0} Backups folder '{1}' of run '{2}' not found - the {0} sheets are built header-only (another run's folder is never used)." -f $Kind, $runFolder, [string]$script:IQ.RunId)
        return $null
    }
    if (-not [string]::IsNullOrWhiteSpace($root)) {
        $latest = Get-IQDateFolder -Root $root
        if ($latest) {
            Write-IQLog -Level Warn -Stage 'Assemble' -Message ("No active run; {0} Backups: using the newest dated folder '{1}'." -f $Kind, $latest)
            return $latest
        }
    }
    return $null
}

# =====================================================================================================================
# Sheet contract
# =====================================================================================================================

function Get-IQSheetContract {
    <#
    .SYNOPSIS
        Loads Config\SheetContract.json (workbook -> sheet -> expectedColumns); $null with a Warn when missing or unreadable.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        $configFolder = $null
        if ($script:IQ) { $configFolder = [string]$script:IQ.ConfigFolder }
        if ([string]::IsNullOrWhiteSpace($configFolder) -and $script:IQ -and $script:IQ.BaseFolder) { $configFolder = Join-Path $script:IQ.BaseFolder 'Config' }
        if ([string]::IsNullOrWhiteSpace($configFolder)) { $configFolder = Join-Path $PSScriptRoot '..' }
        $Path = Join-Path $configFolder 'SheetContract.json'
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-IQLog -Level Warn -Stage 'Assemble' -Message ("SheetContract.json not found at '{0}'; contract sheets/columns cannot be enforced." -f $Path)
        return $null
    }
    try {
        $contract = ConvertFrom-IQJsonFile -Path $Path
        if ($null -eq (Get-IQAsmMember -Object $contract -Name 'workbooks')) {
            Write-IQLog -Level Warn -Stage 'Assemble' -Message ("SheetContract.json at '{0}' has no 'workbooks' node." -f $Path)
            return $null
        }
        return $contract
    }
    catch {
        Write-IQLog -Level Warn -Stage 'Assemble' -Message ("SheetContract.json could not be read: {0}" -f $_.Exception.Message) -Exception $_.Exception
        return $null
    }
}

function Get-IQContractSheetName {
    <#
    .SYNOPSIS
        The sheet names the contract lists for a workbook (e.g. "Report Detail.xlsx"), in contract order; wrap in @( ).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Workbook,
        [Parameter(Mandatory = $false)][AllowNull()]$Contract
    )
    if ($null -eq $Contract) { return @() }
    $workbooks = Get-IQAsmMember -Object $Contract -Name 'workbooks'
    $wb = Get-IQAsmMember -Object $workbooks -Name $Workbook
    return @(Get-IQAsmMemberName -Object $wb)
}

function Get-IQContractColumn {
    <#
    .SYNOPSIS
        The expected columns of one sheet of one workbook from the contract (empty when not listed); wrap in @( ).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Workbook,
        [Parameter(Mandatory = $true)][string]$SheetName,
        [Parameter(Mandatory = $false)][AllowNull()]$Contract
    )
    if ($null -eq $Contract) { return @() }
    $workbooks = Get-IQAsmMember -Object $Contract -Name 'workbooks'
    $wb = Get-IQAsmMember -Object $workbooks -Name $Workbook
    $sheet = Get-IQAsmMember -Object $wb -Name $SheetName
    $cols = @()
    foreach ($c in @(Get-IQAsmMember -Object $sheet -Name 'expectedColumns')) {
        if (-not [string]::IsNullOrWhiteSpace([string]$c)) { $cols += [string]$c }
    }
    return $cols
}

function Ensure-IQSheetColumns {
    <#
    .SYNOPSIS
        Adds every column of Config\SheetContract.json that a sheet lacks (as empty text cells).
    .DESCRIPTION
        -Table (the DataTable built by ConvertTo-IQSheetTable): columns are appended in contract order, existing
        columns keep their order; returns the names that were added.
        -Rows (plain row objects / dictionaries): returns new [PSCustomObject] rows that carry every contract column
        (missing ones as ''); wrap in @( ). An empty -Rows input yields no rows (the caller decides about placeholders).
        With -Contract omitted the contract is loaded from Config\SheetContract.json. The name "Ensure-IQSheetColumns"
        is fixed by the brief (section 6.4/9).
    #>
    [CmdletBinding(DefaultParameterSetName = 'Table')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = 'Function name mandated by the implementation brief.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification = 'Function name mandated by the implementation brief.')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'Table')][System.Data.DataTable]$Table,
        [Parameter(Mandatory = $true, ParameterSetName = 'Rows')][AllowNull()][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory = $true)][string]$Workbook,
        [Parameter(Mandatory = $true)][string]$SheetName,
        [Parameter(Mandatory = $false)][AllowNull()]$Contract
    )
    if ($null -eq $Contract) { $Contract = Get-IQSheetContract }
    $contractColumns = @(Get-IQContractColumn -Workbook $Workbook -SheetName $SheetName -Contract $Contract)

    if ($PSCmdlet.ParameterSetName -eq 'Rows') {
        $out = New-Object System.Collections.Generic.List[object]
        foreach ($row in @($Rows)) {
            if ($null -eq $row) { continue }
            $map = ConvertTo-IQAsmRowMap -Row $row
            foreach ($col in $contractColumns) { if (-not $map.Contains($col)) { $map[$col] = '' } }
            $out.Add([PSCustomObject]$map)
        }
        return $out.ToArray()
    }

    $added = @()
    foreach ($col in $contractColumns) {
        if ($Table.Columns.Contains($col)) { continue }
        $newCol = New-Object System.Data.DataColumn($col, [string])
        $newCol.DefaultValue = ''
        [void]$Table.Columns.Add($newCol)
        foreach ($r in $Table.Rows) { $r[$col] = '' }
        $added += $col
    }
    if ($added.Count -gt 0) {
        Write-IQLog -Level Debug -Stage 'Assemble' -Item $SheetName -Message ("Added {0} missing contract column(s): {1}" -f $added.Count, ($added -join ', '))
    }
    return $added
}

# =====================================================================================================================
# Rows -> typed DataTable
# =====================================================================================================================

function ConvertTo-IQSheetTable {
    <#
    .SYNOPSIS
        Builds a typed System.Data.DataTable from row objects: union of all properties (first-appearance order), leading/trailing fixed columns, verbatim strings.
    .DESCRIPTION
        Column type = bool / datetime / long / double when every non-null value of the column has that kind, else
        string (values converted with ConvertTo-IQAsmCellText; nested objects become compact JSON). Strings longer than
        32,767 characters are truncated (Excel limit) and counted in the table's ExtendedProperties['Truncated'];
        rows beyond 1,048,575 (Excel's row limit minus the header) are dropped, counted in
        ExtendedProperties['TruncatedRows'] and logged at Warn. -LeadingColumns are placed first (created empty when
        absent), -TrailingColumns are appended when absent.
        Performance: Assemble runs after the time budget is spent, so the per-cell work is kept to plain operators
        (member enumeration once per row into name/value arrays, string and $null fast paths, column ordinals); the
        helper functions are only called for the uncommon non-string values.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$LeadingColumns,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$TrailingColumns,
        [Parameter(Mandatory = $false)][string]$SheetName = 'Sheet'
    )
    $rowNames = New-Object System.Collections.Generic.List[object]    # per row: string[] member names
    $rowValues = New-Object System.Collections.Generic.List[object]   # per row: object[] member values
    $columns = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($c in @($LeadingColumns)) {
        if ([string]::IsNullOrWhiteSpace([string]$c) -or $seen.ContainsKey([string]$c)) { continue }
        $seen[[string]$c] = $true; $columns.Add([string]$c)
    }
    $memberTypes = @('NoteProperty', 'Property', 'AliasProperty', 'ScriptProperty')
    $maxRows = [int]$script:IQExcelMaxDataRows
    $truncatedRows = 0
    $lastSignature = $null
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        if ($rowNames.Count -ge $maxRows) { $truncatedRows++; continue }
        $names = $null
        $values = $null
        if ($row -is [System.Collections.IDictionary]) {
            $names = New-Object System.Collections.Generic.List[string]
            $values = New-Object System.Collections.Generic.List[object]
            foreach ($k in @($row.Keys)) { $names.Add([string]$k); $values.Add($row[$k]) }
            $names = $names.ToArray(); $values = $values.ToArray()
        }
        elseif ($row -is [string] -or $row -is [System.ValueType]) {
            $names = [string[]]@('Value'); $values = [object[]]@($row)
        }
        elseif ($row -is [System.Management.Automation.PSCustomObject]) {
            # JSON / [PSCustomObject] rows (the bulk of every sheet): the names come from one engine-side member
            # enumeration and the values are read by index - several times faster than a scripted foreach over
            # PSObject.Properties with .Name/.Value per member. Values are NOT taken by member enumeration because
            # that would flatten array-valued cells and shift the row.
            $props = @($row.PSObject.Properties)
            $names = [string[]]@($props.Name)
            $values = New-Object object[] $props.Count
            for ($i = 0; $i -lt $props.Count; $i++) { $values[$i] = $props[$i].Value }
        }
        else {
            # Other objects (DataRow, class instances): keep the property kinds the sheet should carry.
            $nl = New-Object System.Collections.Generic.List[string]
            $vl = New-Object System.Collections.Generic.List[object]
            foreach ($p in $row.PSObject.Properties) {
                if ($memberTypes -contains $p.MemberType) { $nl.Add($p.Name); $vl.Add($p.Value) }
            }
            $names = $nl.ToArray(); $values = $vl.ToArray()
        }
        # Column discovery only when the row shape differs from the previous row (rows of one collector share it).
        $signature = [string]::Join([string][char]31, $names)
        if ($signature -ne $lastSignature) {
            foreach ($k in $names) {
                if (-not $seen.ContainsKey($k)) { $seen[$k] = $true; $columns.Add($k) }
            }
            $lastSignature = $signature
        }
        $rowNames.Add($names)
        $rowValues.Add($values)
    }
    foreach ($c in @($TrailingColumns)) {
        if ([string]::IsNullOrWhiteSpace([string]$c) -or $seen.ContainsKey([string]$c)) { continue }
        $seen[[string]$c] = $true; $columns.Add([string]$c)
    }

    # Column typing: one pass over the values (strings and nulls decided inline, helper only for the rest).
    $kinds = @{}
    foreach ($c in $columns) { $kinds[$c] = $null }
    for ($ri = 0; $ri -lt $rowNames.Count; $ri++) {
        $names = $rowNames[$ri]; $values = $rowValues[$ri]
        for ($i = 0; $i -lt $names.Length; $i++) {
            $v = $values[$i]
            if ($null -eq $v) { continue }
            $c = $names[$i]
            $current = $kinds[$c]
            if ($current -eq 'String') { continue }
            if ($v -is [string]) { $kinds[$c] = 'String'; continue }
            $kind = Get-IQAsmValueKind -Value $v
            if ($kind -eq 'Null') { continue }
            if ($kind -eq 'Other') { $kind = 'String' }
            if ($null -eq $current) { $kinds[$c] = $kind }
            elseif ($current -eq $kind) { continue }
            elseif (($current -eq 'Integer' -and $kind -eq 'Real') -or ($current -eq 'Real' -and $kind -eq 'Integer')) { $kinds[$c] = 'Real' }
            else { $kinds[$c] = 'String' }
        }
    }

    $table = New-Object System.Data.DataTable
    $table.TableName = (Get-IQSafeSheetName -Name $SheetName)
    $ordinal = @{}
    $isText = @{}
    foreach ($c in $columns) {
        $type = [string]
        switch ([string]$kinds[$c]) {
            'Bool' { $type = [bool] }
            'DateTime' { $type = [datetime] }
            'Integer' { $type = [long] }
            'Real' { $type = [double] }
            default { $type = [string] }
        }
        $col = New-Object System.Data.DataColumn($c, $type)
        # A text column defaults to '' (a missing/null value is an empty cell, as before); typed columns default to DBNull.
        if ($type -eq [string]) { $col.DefaultValue = ''; $isText[$c] = $true } else { $isText[$c] = $false }
        [void]$table.Columns.Add($col)
        $ordinal[$c] = $col.Ordinal
    }

    $truncated = 0
    $cellLimit = [int]$script:IQExcelCellLimit
    for ($ri = 0; $ri -lt $rowNames.Count; $ri++) {
        $names = $rowNames[$ri]; $values = $rowValues[$ri]
        $dr = $table.NewRow()
        for ($i = 0; $i -lt $names.Length; $i++) {
            $v = $values[$i]
            if ($null -eq $v -or $v -is [System.DBNull]) { continue }   # defaults: '' for text, DBNull for typed
            $c = $names[$i]
            if ($isText[$c]) {
                if ($v -is [string]) { $text = $v } else { $text = ConvertTo-IQAsmCellText -Value $v }
                if ($text.Length -gt $cellLimit) { $text = $text.Substring(0, $cellLimit); $truncated++ }
                $dr[$ordinal[$c]] = $text
                continue
            }
            switch ([string]$kinds[$c]) {
                'Bool' { $dr[$ordinal[$c]] = [bool]$v }
                'DateTime' { $dr[$ordinal[$c]] = [datetime]$v }
                'Integer' { $dr[$ordinal[$c]] = [long]$v }
                'Real' { $dr[$ordinal[$c]] = [double]$v }
            }
        }
        $table.Rows.Add($dr)
    }
    $table.ExtendedProperties['Truncated'] = $truncated
    $table.ExtendedProperties['TruncatedRows'] = $truncatedRows
    $table.ExtendedProperties['SourceRows'] = $rowNames.Count + $truncatedRows
    if ($truncated -gt 0) {
        Write-IQLog -Level Warn -Stage 'Assemble' -Item $SheetName -Message ("{0} cell(s) longer than {1} characters were truncated (Excel limit)." -f $truncated, $script:IQExcelCellLimit)
    }
    if ($truncatedRows -gt 0) {
        Write-IQLog -Level Warn -Stage 'Assemble' -Item $SheetName -Message ("{0} row(s) beyond Excel's limit of {1:n0} data rows per sheet were not written." -f $truncatedRows, $maxRows)
    }
    return , $table
}

function Add-IQSheetPlaceholderRow {
    <#
    .SYNOPSIS
        Adds the monolith's "dummy row" (every column an empty string) when the table has no rows; returns $true when added.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][System.Data.DataTable]$Table)
    if ($Table.Rows.Count -gt 0) { return $false }
    if ($Table.Columns.Count -eq 0) { return $false }
    $dr = $Table.NewRow()
    foreach ($col in $Table.Columns) {
        if ($col.DataType -eq [string]) { $dr[$col.ColumnName] = '' } else { $dr[$col.ColumnName] = [System.DBNull]::Value }
    }
    $Table.Rows.Add($dr)
    return $true
}

function New-IQSheetTable {
    <#
    .SYNOPSIS
        Rows -> DataTable with contract columns ensured and the placeholder row added when empty (the per-sheet pipeline).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Workbook,
        [Parameter(Mandatory = $true)][string]$SheetName,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory = $false)][AllowNull()]$Contract,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$LeadingColumns,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$DefaultColumns
    )
    $trailing = @()
    if ($null -ne $DefaultColumns) { $trailing = @($DefaultColumns) }
    $table = ConvertTo-IQSheetTable -Rows $Rows -LeadingColumns $LeadingColumns -TrailingColumns $trailing -SheetName $SheetName
    [void](Ensure-IQSheetColumns -Table $table -Workbook $Workbook -SheetName $SheetName -Contract $Contract)
    if (Add-IQSheetPlaceholderRow -Table $table) {
        Write-IQLog -Level Debug -Stage 'Assemble' -Item $SheetName -Message 'No rows - header plus one empty row written.'
    }
    return , $table
}

# =====================================================================================================================
# Workbook writer (temp file, then move)
# =====================================================================================================================

function Write-IQWorkbook {
    <#
    .SYNOPSIS
        Writes an ordered set of sheet tables to a workbook via a temporary file that replaces the target only after a successful save.
    .DESCRIPTION
        -Sheets is an ordered dictionary sheetName -> System.Data.DataTable (Export-Excel -InputObject fast path).
        Returns @{ Path; Sheets = [string[]] final worksheet names; Rows = total rows; FailedSheets = [string[]] }.
        A sheet whose Export-Excel call fails (EPPlus error) is logged at Error and written again as its header-only
        table (columns only), so every contract sheet still exists and the other sheets are not lost; its name is
        returned in FailedSheets (the caller records the workbook item as Failed). Throws when the workbook itself
        cannot be written or moved into place (e.g. the target is open in Excel); the previous workbook is left
        untouched in that case. Stale "<name>.tmp-*.xlsx" files of a killed earlier run are removed first.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][System.Collections.Specialized.OrderedDictionary]$Sheets,
        [Parameter(Mandatory = $false)][switch]$AutoSize,
        [Parameter(Mandatory = $false)][switch]$AutoNameRange
    )
    if ($Sheets.Count -eq 0) { throw "No sheets to write to '$Path'." }
    $folder = Split-Path -Path $Path -Parent
    if ($folder -and -not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    [void](Remove-IQAsmStaleTempFile -Folder $folder -BaseName $baseName)
    $tmp = Join-Path $folder ($baseName + '.tmp-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.xlsx')

    $useAutoSize = $false
    if ($AutoSize -and $script:IQ -and $script:IQ.IsWindows) { $useAutoSize = $true }

    $pkg = $null
    $finalNames = @()
    $failedSheets = @()
    $usedNames = @{}
    $totalRows = 0
    try {
        foreach ($name in @($Sheets.Keys)) {
            $table = $Sheets[$name]
            if ($null -eq $table -or -not ($table -is [System.Data.DataTable])) {
                Write-IQLog -Level Warn -Stage 'Assemble' -Item ([string]$name) -Message 'Sheet skipped: not a DataTable.'
                continue
            }
            if ($table.Columns.Count -eq 0) {
                Write-IQLog -Level Info -Stage 'Assemble' -Item ([string]$name) -Message 'Sheet not written: no rows and no column list (nothing to put in a header).'
                continue
            }
            $sheetName = Get-IQSafeSheetName -Name ([string]$name)
            $n = 1
            while ($usedNames.ContainsKey($sheetName.ToLowerInvariant())) {
                $n++
                $suffix = ' (' + $n + ')'
                $stem = Get-IQSafeSheetName -Name ([string]$name)
                if ($stem.Length + $suffix.Length -gt 31) { $stem = $stem.Substring(0, 31 - $suffix.Length) }
                $sheetName = $stem + $suffix
            }
            $usedNames[$sheetName.ToLowerInvariant()] = $true
            if ($sheetName -ne [string]$name) {
                Write-IQLog -Level Debug -Stage 'Assemble' -Item ([string]$name) -Message ("Worksheet name written as '{0}' (Excel naming rules)." -f $sheetName)
            }

            # -NoNumberConversion '*' (audit C8-15 / X1-22): the DataTable path already writes strings verbatim, the
            # switch keeps that guarantee if the input path ever changes (piped objects).
            # DateTime-typed columns are written as real Excel dates with ImportExcel's built-in date-time number format
            # (NumFmtId 22); a custom format string must NOT be used here - EPPlus/Import-Excel only read a numeric cell
            # back as [datetime] when its number format is a built-in date format.
            $params = @{ WorksheetName = $sheetName; PassThru = $true; NoNumberConversion = @('*') }
            if ($useAutoSize) { $params['AutoSize'] = $true }
            if ($AutoNameRange) { $params['AutoNameRange'] = $true }
            $written = $false
            try {
                if ($null -eq $pkg) { $params['Path'] = $tmp } else { $params['ExcelPackage'] = $pkg }
                $pkg = Export-Excel -InputObject $table @params
                $written = $true
            }
            catch {
                # One failing sheet must not cost the workbook (every contract sheet missing = PBIT fatal, brief
                # section 14): drop the half-built worksheet and write the header-only table instead.
                $failedSheets += [string]$name
                Write-IQLog -Level Error -Stage 'Assemble' -Item $sheetName -Message ("Sheet could not be written ({0} row(s), {1} column(s)); written header-only instead: {2}" -f $table.Rows.Count, $table.Columns.Count, $_.Exception.Message) -Exception $_.Exception
                if ($null -ne $pkg) {
                    try { if ($null -ne $pkg.Workbook.Worksheets[$sheetName]) { $pkg.Workbook.Worksheets.Delete($sheetName) } }
                    catch { Write-IQLog -Level Debug -Stage 'Assemble' -Item $sheetName -Message ("Partial worksheet could not be removed: {0}" -f $_.Exception.Message) }
                }
                $headerOnly = $table.Clone()   # same columns, no rows
                [void](Add-IQSheetPlaceholderRow -Table $headerOnly)
                $params.Remove('Path'); $params.Remove('ExcelPackage')
                if ($null -eq $pkg) { $params['Path'] = $tmp } else { $params['ExcelPackage'] = $pkg }
                $pkg = Export-Excel -InputObject $headerOnly @params
                $table = $headerOnly
            }
            $finalNames += $sheetName
            $totalRows += $table.Rows.Count
            if ($written) { Write-IQLog -Level Debug -Stage 'Assemble' -Item $sheetName -Message ("{0} row(s), {1} column(s)" -f $table.Rows.Count, $table.Columns.Count) }
        }
        if ($null -eq $pkg) { throw "No sheet could be written to '$Path'." }
        Close-ExcelPackage -ExcelPackage $pkg
        $pkg = $null
        if (-not (Test-Path -LiteralPath $tmp) -or (Get-Item -LiteralPath $tmp).Length -le 0) { throw "Temporary workbook '$tmp' was not created." }
        # .NET copy + delete instead of Move-Item: -Destination has no -LiteralPath variant and Windows PowerShell 5.1
        # globs '[' / ']' in it (a BaseFolder like 'D:\PBI Governance [PROD]' broke the final move). Copy-then-delete
        # keeps the previous workbook until the new one is fully in place and still fails cleanly when the target is
        # locked (open in Excel).
        [System.IO.File]::Copy($tmp, $Path, $true)
        [System.IO.File]::Delete($tmp)
    }
    catch {
        if ($null -ne $pkg) { try { Close-ExcelPackage -ExcelPackage $pkg -NoSave } catch { Write-IQLog -Level Debug -Stage 'Assemble' -Message ("Temp package close failed: {0}" -f $_.Exception.Message) } }
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        throw
    }
    return @{ Path = $Path; Sheets = $finalNames; Rows = $totalRows; FailedSheets = $failedSheets }
}

# =====================================================================================================================
# Report Detail: *.txt -> sheets (monolith 3063-3110)
# =====================================================================================================================

function ConvertFrom-IQReportDetailText {
    <#
    .SYNOPSIS
        Parses one tab-separated csx output file (UTF-8): header line, then rows split into exactly the header count of columns; wrap in @( ).
    .DESCRIPTION
        Monolith 3086-3103 with the audit fixes: a one-line (header-only) file yields zero rows instead of a garbage
        sheet (C8-04), column order is preserved with [ordered] (C8-06), missing trailing values become empty strings,
        completely empty lines are ignored. Returns @{ Headers = [string[]]; Rows = [object[]] }.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)
    $result = @{ Headers = @(); Rows = @() }
    if (-not (Test-Path -LiteralPath $Path)) { return $result }
    $lines = @([System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8))
    if ($lines.Count -eq 0) { return $result }
    $headerLine = $lines[0]
    if ($headerLine.Length -gt 0 -and [int][char]$headerLine[0] -eq 0xFEFF) { $headerLine = $headerLine.Substring(1) }
    $headers = @($headerLine -split "`t")
    $result.Headers = $headers
    if ($headers.Count -eq 0 -or ($headers.Count -eq 1 -and [string]::IsNullOrWhiteSpace($headers[0]))) { return $result }
    $rows = New-Object System.Collections.Generic.List[object]
    for ($li = 1; $li -lt $lines.Count; $li++) {
        $line = $lines[$li]
        if ($null -eq $line -or $line.Length -eq 0) { continue }
        $values = @($line -split "`t", $headers.Count)   # exact number of columns as headers (monolith 3096)
        $obj = [ordered]@{}
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $name = $headers[$i]
            if ([string]::IsNullOrEmpty($name)) { $name = 'Column' + ($i + 1) }
            if ($obj.Contains($name)) { $name = $name + '_' + ($i + 1) }
            if ($i -lt $values.Count) { $obj[$name] = $values[$i] } else { $obj[$name] = '' }
        }
        $rows.Add([PSCustomObject]$obj)
    }
    $result.Rows = $rows.ToArray()
    return $result
}

function Get-IQReportDetailSheetMap {
    <#
    .SYNOPSIS
        Ordered dictionary sheetName -> @{ Rows; Columns } for every *.txt in the report backup folder (file order, as the monolith), then missing contract sheets.
    .DESCRIPTION
        Columns are the header line of the file, so a header-only txt (e.g. ReportExports.txt when nothing was
        exported, ExtractErrors.txt, ReportObjects_UnusedObjects.txt, or a contract sheet whose csx writes extra
        columns) still yields its header row in the workbook. A file that cannot be read is recorded with
        Add-IQAsmSourceError (failed Assemble item) and its sheet is written header-only.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Folder,
        [Parameter(Mandatory = $false)][AllowNull()]$Contract
    )
    $map = [ordered]@{}
    if (-not [string]::IsNullOrWhiteSpace($Folder) -and (Test-Path -LiteralPath $Folder)) {
        foreach ($txt in @(Get-ChildItem -LiteralPath $Folder -Filter '*.txt' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
            $sheet = [System.IO.Path]::GetFileNameWithoutExtension($txt.Name)
            try {
                $parsed = ConvertFrom-IQReportDetailText -Path $txt.FullName
                $headers = @($parsed.Headers | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
                $map[$sheet] = @{ Rows = @($parsed.Rows); Columns = @($headers) }
                Write-IQLog -Level Debug -Stage 'Assemble' -Item $sheet -Message ("{0}: {1} row(s), {2} column(s)" -f $txt.Name, @($parsed.Rows).Count, @($parsed.Headers).Count)
            }
            catch {
                Add-IQAsmSourceError -Path $txt.FullName -Item $sheet -Message ("Could not parse '{0}': {1}" -f $txt.FullName, $_.Exception.Message) -Exception $_.Exception
                if (-not $map.Contains($sheet)) { $map[$sheet] = @{ Rows = @(); Columns = @() } }
            }
        }
    }
    foreach ($sheet in @(Get-IQContractSheetName -Workbook 'Report Detail.xlsx' -Contract $Contract)) {
        if (-not $map.Contains($sheet)) { $map[$sheet] = @{ Rows = @(); Columns = @() } }
    }
    return $map
}

# =====================================================================================================================
# Model Detail: *.csv -> "Semantic Models" / "Measure Dependencies" (monolith 3209-3290)
# =====================================================================================================================

function Get-IQModelDetailRow {
    <#
    .SYNOPSIS
        Concatenates the csx/DAX CSV files of the model backup folder (Import-Csv, UTF-8): all *.csv except *_MD.csv, or only *_MD.csv with -Dependencies; wrap in @( ).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Folder,
        [Parameter(Mandatory = $false)][switch]$Dependencies
    )
    $rows = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Folder) -or -not (Test-Path -LiteralPath $Folder)) { return $rows.ToArray() }
    $files = 0
    foreach ($csv in @(Get-ChildItem -LiteralPath $Folder -Filter '*.csv' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        $isMd = ($csv.Name -like '*_MD.csv')
        if ($Dependencies -and -not $isMd) { continue }
        if (-not $Dependencies -and $isMd) { continue }
        try {
            $data = @(Import-Csv -LiteralPath $csv.FullName -Encoding UTF8)
            foreach ($r in $data) { if ($null -ne $r) { $rows.Add($r) } }
            $files++
        }
        catch {
            Add-IQAsmSourceError -Path $csv.FullName -Item $csv.Name -Message ("Could not read CSV '{0}': {1}" -f $csv.FullName, $_.Exception.Message) -Exception $_.Exception
        }
    }
    $kind = 'Semantic Models'
    if ($Dependencies) { $kind = 'Measure Dependencies' }
    Write-IQLog -Level Debug -Stage 'Assemble' -Item $kind -Message ("{0} CSV file(s), {1} row(s) from '{2}'" -f $files, $rows.Count, $Folder)
    return $rows.ToArray()
}

# =====================================================================================================================
# Dataflow Detail: extracts\dataflows\*.json -> Sheet1 (monolith 3498-3510, 3701-3713)
# =====================================================================================================================

function Get-IQDataflowDetailRow {
    <#
    .SYNOPSIS
        The Sheet1 rows of "Dataflow Detail.xlsx" from every extracts\dataflows\*.json of the run (Queries arrays, file order); wrap in @( ).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Folder)
    if ([string]::IsNullOrWhiteSpace($Folder)) {
        $extracts = Get-IQAsmRunPath -SubFolder 'Extracts'
        if (-not [string]::IsNullOrWhiteSpace($extracts)) { $Folder = Join-Path $extracts 'dataflows' }
    }
    $rows = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Folder) -or -not (Test-Path -LiteralPath $Folder)) { return $rows.ToArray() }
    $files = 0
    foreach ($f in @(Get-ChildItem -LiteralPath $Folder -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        try {
            $extract = ConvertFrom-IQJsonFile -Path $f.FullName
            if ($null -eq $extract) { continue }
            $files++
            $queries = Get-IQAsmMember -Object $extract -Name 'Queries'
            if ($null -eq $queries -and $extract -is [System.Array]) { $queries = $extract }   # tolerate a bare array of rows
            foreach ($q in @($queries)) { if ($null -ne $q) { $rows.Add($q) } }
        }
        catch {
            Add-IQAsmSourceError -Path $f.FullName -Item $f.Name -Message ("Could not read dataflow extract '{0}': {1}" -f $f.FullName, $_.Exception.Message) -Exception $_.Exception
        }
    }
    Write-IQLog -Level Debug -Stage 'Assemble' -Item 'Sheet1' -Message ("{0} dataflow extract(s), {1} query row(s)" -f $files, $rows.Count)
    return $rows.ToArray()
}

# =====================================================================================================================
# Environment Detail: inventory + manifest + extras -> sheet rows
# =====================================================================================================================

function Get-IQAssembleRunSummaryRow {
    <#
    .SYNOPSIS
        RunSummary rows: one "(Run)" row followed by one row per stage of the manifest (pipeline order); wrap in @( ).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Manifest)
    if ($null -eq $Manifest -and $script:IQ) { $Manifest = $script:IQ.Manifest }
    if ($null -eq $Manifest) { return @() }
    if (Get-Command -Name 'Get-IQRunSummary' -ErrorAction SilentlyContinue) {
        try { return @(Get-IQRunSummary -Manifest $Manifest) }
        catch { Write-IQLog -Level Debug -Stage 'Assemble' -Message ("Get-IQRunSummary failed, building RunSummary locally: {0}" -f $_.Exception.Message) }
    }
    $order = @('Inventory', 'ModelBackup', 'ReportBackup', 'ReportDetail', 'ModelDetail', 'Dataflows', 'Extras', 'Assemble')
    $runId = [string](Get-IQAsmMember -Object $Manifest -Name 'runId')
    $hostInfo = Get-IQAsmMember -Object $Manifest -Name 'host'
    $scope = Get-IQAsmMember -Object $Manifest -Name 'scope'
    $stages = Get-IQAsmMember -Object $Manifest -Name 'stages'
    $failures = @(Get-IQAsmMember -Object $Manifest -Name 'failures')
    $environment = [string](Get-IQAsmMember -Object $Manifest -Name 'environment')
    $auth = [string](Get-IQAsmMember -Object $Manifest -Name 'auth')
    $resumeCount = Get-IQAsmMember -Object $Manifest -Name 'resumeCount'
    if ($null -eq $resumeCount) { $resumeCount = 0 }
    $names = @(Get-IQAsmMemberName -Object $stages)
    $ordered = @($order | Where-Object { $_ -in $names }) + @($names | Where-Object { $_ -notin $order })

    $common = [ordered]@{
        Environment   = $environment
        Auth          = $auth
        RunMode       = [string](Get-IQAsmMember -Object $scope -Name 'runMode')
        Machine       = [string](Get-IQAsmMember -Object $hostInfo -Name 'machine')
        User          = [string](Get-IQAsmMember -Object $hostInfo -Name 'user')
        PSVersion     = [string](Get-IQAsmMember -Object $hostInfo -Name 'psVersion')
        IsAzureDevOps = [bool](Get-IQAsmMember -Object $hostInfo -Name 'isAzureDevOps')
        ResumeCount   = [int]$resumeCount
    }
    $rows = New-Object System.Collections.Generic.List[object]
    $totalDone = 0; $totalFailed = 0
    $stageRows = @()
    foreach ($name in $ordered) {
        $st = Get-IQAsmMember -Object $stages -Name $name
        $done = Get-IQAsmMember -Object $st -Name 'itemsDone'; if ($null -eq $done) { $done = 0 }
        $failed = Get-IQAsmMember -Object $st -Name 'itemsFailed'; if ($null -eq $failed) { $failed = 0 }
        $totalDone += [int]$done; $totalFailed += [int]$failed
        $dur = $null
        $s = $null; $e = $null
        try { $sv = Get-IQAsmMember -Object $st -Name 'startedUtc'; if ($sv) { $s = [datetime]::Parse([string]$sv, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal) } } catch { $s = $null }
        try { $ev = Get-IQAsmMember -Object $st -Name 'endedUtc'; if ($ev) { $e = [datetime]::Parse([string]$ev, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AdjustToUniversal) } } catch { $e = $null }
        if ($null -ne $s -and $null -ne $e) { $dur = [math]::Round(($e - $s).TotalSeconds, 1) }
        $row = [ordered]@{
            RunId           = $runId
            Stage           = $name
            Status          = [string](Get-IQAsmMember -Object $st -Name 'status')
            StartedUtc      = [string](Get-IQAsmMember -Object $st -Name 'startedUtc')
            EndedUtc        = [string](Get-IQAsmMember -Object $st -Name 'endedUtc')
            DurationSeconds = $dur
            ItemsDone       = [int]$done
            ItemsFailed     = [int]$failed
            Error           = ConvertTo-IQRedactedText -Text ([string](Get-IQAsmMember -Object $st -Name 'error'))
        }
        foreach ($k in $common.Keys) { $row[$k] = $common[$k] }
        $stageRows += [PSCustomObject]$row
    }
    $runError = ''
    if ($failures.Count -gt 0) { $runError = ('{0} item failure(s)' -f $failures.Count) }
    $runRow = [ordered]@{
        RunId           = $runId
        Stage           = '(Run)'
        Status          = [string](Get-IQAsmMember -Object $Manifest -Name 'status')
        StartedUtc      = [string](Get-IQAsmMember -Object $Manifest -Name 'startedUtc')
        EndedUtc        = [string](Get-IQAsmMember -Object $Manifest -Name 'endedUtc')
        DurationSeconds = $null
        ItemsDone       = $totalDone
        ItemsFailed     = $totalFailed
        Error           = $runError
    }
    foreach ($k in $common.Keys) { $runRow[$k] = $common[$k] }
    $rows.Add([PSCustomObject]$runRow)
    foreach ($r in $stageRows) { $rows.Add($r) }
    return $rows.ToArray()
}

function Get-IQAssembleFailureRow {
    <#
    .SYNOPSIS
        Failures rows from manifest.failures (Stage, ItemKey, Item, Message, TimeUtc, RunId); wrap in @( ).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Manifest)
    if ($null -eq $Manifest -and $script:IQ) { $Manifest = $script:IQ.Manifest }
    $rows = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Manifest) { return $rows.ToArray() }
    $runId = [string](Get-IQAsmMember -Object $Manifest -Name 'runId')
    foreach ($f in @(Get-IQAsmMember -Object $Manifest -Name 'failures')) {
        if ($null -eq $f) { continue }
        $rows.Add([PSCustomObject]@{
                Stage   = [string](Get-IQAsmMember -Object $f -Name 'stage')
                ItemKey = [string](Get-IQAsmMember -Object $f -Name 'itemKey')
                Item    = [string](Get-IQAsmMember -Object $f -Name 'item')
                Message = ConvertTo-IQRedactedText -Text ([string](Get-IQAsmMember -Object $f -Name 'message'))
                TimeUtc = [string](Get-IQAsmMember -Object $f -Name 'timeUtc')
                RunId   = $runId
            })
    }
    return $rows.ToArray()
}

function Get-IQAssembleInventoryErrorRow {
    <#
    .SYNOPSIS
        InventoryErrors rows: the Errors[] recorded in global.json and every ws-*.json (collector failures that only yielded empty lists); wrap in @( ).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Global,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()][object[]]$WorkspaceInventories
    )
    $rows = New-Object System.Collections.Generic.List[object]
    $add = {
        param($err, $wsId, $wsName)
        if ($null -eq $err) { return }
        $rows.Add([PSCustomObject]@{
                WorkspaceId   = $wsId
                WorkspaceName = $wsName
                Collector     = [string](Get-IQAsmMember -Object $err -Name 'Collector')
                Path          = [string](Get-IQAsmMember -Object $err -Name 'Path')
                Message       = ConvertTo-IQRedactedText -Text ([string](Get-IQAsmMember -Object $err -Name 'Message'))
            })
    }
    foreach ($e in @(Get-IQAsmMember -Object $Global -Name 'Errors')) { & $add $e '' '(global)' }
    foreach ($inv in @($WorkspaceInventories)) {
        if ($null -eq $inv) { continue }
        $wsId = [string](Get-IQAsmMember -Object $inv -Name 'WorkspaceId')
        $wsName = [string](Get-IQAsmMember -Object $inv -Name 'WorkspaceName')
        foreach ($e in @(Get-IQAsmMember -Object $inv -Name 'Errors')) { & $add $e $wsId $wsName }
    }
    return $rows.ToArray()
}

function Get-IQExtraSheetMap {
    <#
    .SYNOPSIS
        Sheets from every inventory\extras-*.json written by the Extras stage (sheet name = collector name): @{ Sheets = <ordered name -> rows>; Columns = @{ name -> string[] } }.
    .DESCRIPTION
        Accepted file shapes:
          - the Extras module's sheet object { SheetName; Collector; Columns[]; Rows[] } (Rows may be empty - the
            Columns list then still yields a header row);
          - a bare JSON array of rows (sheet named from the file name, canonical casing restored for the known
            collectors);
          - an object whose properties are arrays (one sheet per property, property name = sheet name);
          - anything else becomes a one-row sheet.
        Rows of two files that resolve to the same sheet name are concatenated.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$InventoryFolder)
    $map = [ordered]@{}
    $columns = @{}
    $result = @{ Sheets = $map; Columns = $columns }
    if ([string]::IsNullOrWhiteSpace($InventoryFolder)) { $InventoryFolder = Get-IQAsmRunPath -SubFolder 'Inventory' }
    if ([string]::IsNullOrWhiteSpace($InventoryFolder) -or -not (Test-Path -LiteralPath $InventoryFolder)) { return $result }
    $canonical = @{}
    foreach ($n in $script:IQExtrasSheetNames) { $canonical[$n.ToLowerInvariant()] = $n }

    $addRows = {
        param([string]$sheetName, [object[]]$rows, [string[]]$cols)
        if ($map.Contains($sheetName)) { $map[$sheetName] = @($map[$sheetName]) + @($rows) }
        else { $map[$sheetName] = @($rows) }
        if ($null -ne $cols -and $cols.Count -gt 0 -and -not $columns.ContainsKey($sheetName)) { $columns[$sheetName] = @($cols) }
    }

    foreach ($f in @(Get-ChildItem -LiteralPath $InventoryFolder -Filter 'extras-*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $rawName = $stem.Substring('extras-'.Length)
        $defaultName = $rawName
        if ($canonical.ContainsKey($rawName.ToLowerInvariant())) { $defaultName = $canonical[$rawName.ToLowerInvariant()] }
        if ([string]::IsNullOrWhiteSpace($defaultName)) { $defaultName = $stem }
        try {
            # A top-level JSON array is unrolled by ConvertFrom-IQJsonFile (1 element -> the element, 0 -> $null), so the
            # shape is decided from the first non-blank character of the file, not from the parsed value. The file is
            # sniffed with a reader (not ReadAllText): an ActivityEvents file can be hundreds of MB and would otherwise
            # be held twice in memory on Windows PowerShell 5.1.
            $firstChar = ''
            $sr = [System.IO.File]::OpenText($f.FullName)
            try {
                do { $ch = $sr.Read() } while ($ch -ge 0 -and ([char]$ch -eq [char]0xFEFF -or [char]::IsWhiteSpace([char]$ch)))
                if ($ch -ge 0) { $firstChar = [string][char]$ch }
            }
            finally { $sr.Dispose() }
            if ($firstChar -eq '[') {
                & $addRows $defaultName @(ConvertFrom-IQJsonFile -Path $f.FullName | Where-Object { $null -ne $_ }) @()
                continue
            }
            $obj = ConvertFrom-IQJsonFile -Path $f.FullName
            if ($null -eq $obj) { & $addRows $defaultName @() @(); continue }
            if ($obj -is [string] -or $obj -is [System.ValueType]) { & $addRows $defaultName @([PSCustomObject]@{ Value = $obj }) @(); continue }

            # Shape 1: { SheetName; Columns; Rows } (Rows/Items may be empty)
            $rowsName = $null
            foreach ($candidate in @('Rows', 'rows', 'Items', 'items')) {
                if (Test-IQAsmHasMember -Object $obj -Name $candidate) { $rowsName = $candidate; break }
            }
            if ($null -ne $rowsName) {
                $sheetName = $defaultName
                foreach ($candidate in @('SheetName', 'sheetName', 'Sheet', 'sheet', 'Collector', 'collector', 'Name', 'name')) {
                    $v = Get-IQAsmMember -Object $obj -Name $candidate
                    if ($v -is [string] -and -not [string]::IsNullOrWhiteSpace($v)) { $sheetName = $v; break }
                }
                $cols = @()
                foreach ($candidate in @('Columns', 'columns')) {
                    if (Test-IQAsmHasMember -Object $obj -Name $candidate) {
                        foreach ($c in @(Get-IQAsmMember -Object $obj -Name $candidate)) { if (-not [string]::IsNullOrWhiteSpace([string]$c)) { $cols += [string]$c } }
                        break
                    }
                }
                & $addRows $sheetName @(Get-IQAsmMember -Object $obj -Name $rowsName | Where-Object { $null -ne $_ }) $cols
                continue
            }

            # Shape 2: object of arrays (one sheet per array property)
            $arrayProps = @()
            foreach ($name in @(Get-IQAsmMemberName -Object $obj)) {
                if (Test-IQAsmArrayMember -Object $obj -Name $name) { $arrayProps += $name }
            }
            if ($arrayProps.Count -gt 0) {
                foreach ($name in $arrayProps) { & $addRows $name @(Get-IQAsmMember -Object $obj -Name $name | Where-Object { $null -ne $_ }) @() }
                continue
            }

            # Shape 3: a single row object
            & $addRows $defaultName @($obj) @()
        }
        catch {
            Add-IQAsmSourceError -Path $f.FullName -Item $f.Name -Message ("Could not read extras file '{0}': {1}" -f $f.FullName, $_.Exception.Message) -Exception $_.Exception
        }
    }
    if ($map.Count -gt 0) {
        Write-IQLog -Level Info -Stage 'Assemble' -Message ("Extras sheets found: {0}" -f (@($map.Keys) -join ', '))
    }
    return $result
}

function Get-IQEnvironmentSheetMap {
    <#
    .SYNOPSIS
        Ordered dictionary sheetName -> rows for "Power BI Environment Detail.xlsx" (monolith order, then the new sheets, then extras).
    .DESCRIPTION
        Sources: inventory\workspaces.json, global.json (Connections, Gateways, Capacities, Apps, AppReports), every
        ws-*.json (concatenated per collection), the manifest (RunSummary, Failures), Errors[] (InventoryErrors) and
        extras-*.json. Returns @{ Sheets = <ordered>; Extras = [string[]] extras sheet names }.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Manifest)
    if ($null -eq $Manifest -and $script:IQ) { $Manifest = $script:IQ.Manifest }
    $sheets = [ordered]@{}

    $workspaces = @()
    $global = $null
    $wsInventories = @()
    $inventoryFolder = Get-IQAsmRunPath -SubFolder 'Inventory'
    if (-not [string]::IsNullOrWhiteSpace($inventoryFolder) -and (Test-Path -LiteralPath $inventoryFolder)) {
        try { $workspaces = @(Get-IQInventory -Name 'workspaces') } catch { Write-IQLog -Level Warn -Stage 'Assemble' -Message ("workspaces.json could not be read: {0}" -f $_.Exception.Message) }
        try { $global = Get-IQInventory -Name 'global' } catch { Write-IQLog -Level Warn -Stage 'Assemble' -Message ("global.json could not be read: {0}" -f $_.Exception.Message) }
        try { $wsInventories = @(Get-IQAllWorkspaceInventories) } catch { Write-IQLog -Level Warn -Stage 'Assemble' -Message ("ws-*.json could not be read: {0}" -f $_.Exception.Message) }
    }
    else {
        Write-IQLog -Level Warn -Stage 'Assemble' -Message 'No inventory folder for this run - the Environment workbook will contain header-only sheets.'
    }
    $workspaces = @($workspaces | Where-Object { $null -ne $_ })
    Write-IQLog -Level Info -Stage 'Assemble' -Message ("Inventory: {0} workspace row(s), {1} workspace file(s), global.json {2}" -f $workspaces.Count, $wsInventories.Count, $(if ($null -ne $global) { 'present' } else { 'missing' }))

    foreach ($sheet in $script:IQEnvironmentSheetOrder) {
        $src = $script:IQEnvironmentSheetSource[$sheet]
        $rows = @()
        switch ($src.Source) {
            'workspaces' { $rows = $workspaces }
            'global' { $rows = @(Get-IQAsmMember -Object $global -Name $src.Key | Where-Object { $null -ne $_ }) }
            'ws' {
                $list = New-Object System.Collections.Generic.List[object]
                foreach ($inv in $wsInventories) {
                    foreach ($r in @(Get-IQAsmMember -Object $inv -Name $src.Key)) { if ($null -ne $r) { $list.Add($r) } }
                }
                $rows = $list.ToArray()
            }
            'manifest' {
                switch ($src.Key) {
                    'RunSummary' { $rows = @(Get-IQAssembleRunSummaryRow -Manifest $Manifest) }
                    'Failures' { $rows = @(Get-IQAssembleFailureRow -Manifest $Manifest) }
                    'InventoryErrors' { $rows = @(Get-IQAssembleInventoryErrorRow -Global $global -WorkspaceInventories $wsInventories) }
                }
            }
        }
        $sheets[$sheet] = @($rows)
    }

    $extras = Get-IQExtraSheetMap -InventoryFolder $inventoryFolder
    $extraNames = @()
    $extraColumns = @{}
    foreach ($name in @($extras.Sheets.Keys)) {
        $target = [string]$name
        if ($sheets.Contains($target)) { $target = 'Extras' + $target }   # never overwrite a core sheet
        $sheets[$target] = @($extras.Sheets[$name])
        if ($extras.Columns.ContainsKey([string]$name)) { $extraColumns[$target] = @($extras.Columns[[string]$name]) }
        $extraNames += $target
    }
    return @{ Sheets = $sheets; Extras = $extraNames; Columns = $extraColumns }
}

# =====================================================================================================================
# Workbook builders
# =====================================================================================================================

function Build-IQEnvironmentWorkbook {
    <#
    .SYNOPSIS
        Builds "Power BI Environment Detail.xlsx" in the BaseFolder from inventory, manifest and extras; returns the Write-IQWorkbook result.
    #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = 'Build-* mirrors the stage vocabulary of the brief; private to this module.')]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][AllowNull()]$Contract,
        [Parameter(Mandatory = $false)][AllowNull()]$Manifest
    )
    $workbook = 'Power BI Environment Detail.xlsx'
    $source = Get-IQEnvironmentSheetMap -Manifest $Manifest
    $tables = [ordered]@{}
    foreach ($name in @($source.Sheets.Keys)) {
        $defaults = $null
        if ($script:IQAssembleDefaultColumns.ContainsKey([string]$name)) { $defaults = $script:IQAssembleDefaultColumns[[string]$name] }
        elseif ($source.Columns.ContainsKey([string]$name)) { $defaults = @($source.Columns[[string]$name]) }   # extras sheet: header from the collector's Columns[]
        $tables[$name] = New-IQSheetTable -Workbook $workbook -SheetName ([string]$name) -Rows $source.Sheets[$name] -Contract $Contract -DefaultColumns $defaults
    }
    # Every contract sheet must exist even if the source map did not produce it (defensive).
    foreach ($sheet in @(Get-IQContractSheetName -Workbook $workbook -Contract $Contract)) {
        if (-not $tables.Contains($sheet)) { $tables[$sheet] = New-IQSheetTable -Workbook $workbook -SheetName $sheet -Rows @() -Contract $Contract }
    }
    return (Write-IQWorkbook -Path $Path -Sheets $tables -AutoSize)
}

function Build-IQReportWorkbook {
    <#
    .SYNOPSIS
        Builds "Report Detail.xlsx" from the *.txt files of the report backup folder; returns the Write-IQWorkbook result.
    #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = 'Build-* mirrors the stage vocabulary of the brief; private to this module.')]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Folder,
        [Parameter(Mandatory = $false)][AllowNull()]$Contract
    )
    $workbook = 'Report Detail.xlsx'
    if ([string]::IsNullOrWhiteSpace($Folder)) { $Folder = Get-IQAssembleFolder -Kind 'Report' }
    if ([string]::IsNullOrWhiteSpace($Folder)) { Write-IQLog -Level Warn -Stage 'Assemble' -Message 'No Report Backups folder found - Report Detail.xlsx will contain header-only sheets.' }
    else { Write-IQLog -Level Info -Stage 'Assemble' -Message ("Report Detail source folder: {0}" -f $Folder) }
    $source = Get-IQReportDetailSheetMap -Folder $Folder -Contract $Contract
    $tables = [ordered]@{}
    foreach ($name in @($source.Keys)) {
        $entry = $source[$name]
        # The txt header row is the sheet's default column list, so header-only files keep their columns.
        $tables[$name] = New-IQSheetTable -Workbook $workbook -SheetName ([string]$name) -Rows @($entry.Rows) -Contract $Contract -DefaultColumns @($entry.Columns)
    }
    return (Write-IQWorkbook -Path $Path -Sheets $tables -AutoNameRange)
}

function Build-IQModelWorkbook {
    <#
    .SYNOPSIS
        Builds "Model Detail.xlsx" ("Semantic Models", "Measure Dependencies") from the CSVs of the model backup folder; returns the Write-IQWorkbook result.
    #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = 'Build-* mirrors the stage vocabulary of the brief; private to this module.')]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Folder,
        [Parameter(Mandatory = $false)][AllowNull()]$Contract
    )
    $workbook = 'Model Detail.xlsx'
    if ([string]::IsNullOrWhiteSpace($Folder)) { $Folder = Get-IQAssembleFolder -Kind 'Model' }
    if ([string]::IsNullOrWhiteSpace($Folder)) { Write-IQLog -Level Warn -Stage 'Assemble' -Message 'No Model Backups folder found - Model Detail.xlsx will contain header-only sheets.' }
    else { Write-IQLog -Level Info -Stage 'Assemble' -Message ("Model Detail source folder: {0}" -f $Folder) }
    $tables = [ordered]@{}
    $tables['Semantic Models'] = New-IQSheetTable -Workbook $workbook -SheetName 'Semantic Models' -Rows @(Get-IQModelDetailRow -Folder $Folder) -Contract $Contract
    $tables['Measure Dependencies'] = New-IQSheetTable -Workbook $workbook -SheetName 'Measure Dependencies' -Rows @(Get-IQModelDetailRow -Folder $Folder -Dependencies) -Contract $Contract
    foreach ($sheet in @(Get-IQContractSheetName -Workbook $workbook -Contract $Contract)) {
        if (-not $tables.Contains($sheet)) { $tables[$sheet] = New-IQSheetTable -Workbook $workbook -SheetName $sheet -Rows @() -Contract $Contract }
    }
    return (Write-IQWorkbook -Path $Path -Sheets $tables -AutoNameRange)
}

function Build-IQDataflowWorkbook {
    <#
    .SYNOPSIS
        Builds "Dataflow Detail.xlsx" (Sheet1) in the dataflow backup folder and copies it to the BaseFolder; returns @{ Path; CopyPath; Sheets; Rows }.
    #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification = 'Build-* mirrors the stage vocabulary of the brief; private to this module.')]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$CopyPath,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$ExtractFolder,
        [Parameter(Mandatory = $false)][AllowNull()]$Contract
    )
    $workbook = 'Dataflow Detail.xlsx'
    $rows = @(Get-IQDataflowDetailRow -Folder $ExtractFolder)
    # Rows: the six monolith columns first, then the five DataRow artefact columns (empty strings), then additive ones.
    $shaped = New-Object System.Collections.Generic.List[object]
    foreach ($r in $rows) {
        $map = ConvertTo-IQAsmRowMap -Row $r
        $o = [ordered]@{}
        foreach ($c in $script:IQDataflowSheetColumns) { if ($map.Contains($c)) { $o[$c] = $map[$c] } else { $o[$c] = '' } }
        foreach ($c in $script:IQDataflowArtefactColumns) { $o[$c] = '' }
        foreach ($k in @($map.Keys)) { if (-not $o.Contains($k)) { $o[$k] = $map[$k] } }
        $shaped.Add([PSCustomObject]$o)
    }
    $leading = @($script:IQDataflowSheetColumns) + @($script:IQDataflowArtefactColumns)
    $tables = [ordered]@{}
    $tables['Sheet1'] = New-IQSheetTable -Workbook $workbook -SheetName 'Sheet1' -Rows $shaped.ToArray() -Contract $Contract -LeadingColumns $leading
    foreach ($sheet in @(Get-IQContractSheetName -Workbook $workbook -Contract $Contract)) {
        if (-not $tables.Contains($sheet)) { $tables[$sheet] = New-IQSheetTable -Workbook $workbook -SheetName $sheet -Rows @() -Contract $Contract }
    }
    $result = Write-IQWorkbook -Path $Path -Sheets $tables -AutoSize
    $result['CopyPath'] = $null
    if (-not [string]::IsNullOrWhiteSpace($CopyPath) -and $CopyPath -ne $Path) {
        $copyFolder = Split-Path -Path $CopyPath -Parent
        if ($copyFolder -and -not (Test-Path -LiteralPath $copyFolder)) { New-Item -ItemType Directory -Path $copyFolder -Force | Out-Null }
        # Stale copy temp files of earlier module versions / killed runs, then a .NET copy: the source is a complete,
        # closed workbook and File.Copy never globs '[' / ']' in the BaseFolder (Copy-Item/Move-Item -Destination do).
        [void](Remove-IQAsmStaleTempFile -Folder $copyFolder -BaseName ([System.IO.Path]::GetFileNameWithoutExtension($CopyPath)))
        [System.IO.File]::Copy($Path, $CopyPath, $true)
        $result['CopyPath'] = $CopyPath
    }
    return $result
}

# =====================================================================================================================
# Stage body
# =====================================================================================================================

function Set-IQManifestOutput {
    <#
    .SYNOPSIS
        Stores the workbook paths in manifest.outputs (environmentWorkbook, reportWorkbook, modelWorkbook, dataflowWorkbook) and saves the manifest.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Outputs)
    if (-not $script:IQ -or $null -eq $script:IQ.Manifest) { return }
    $m = $script:IQ.Manifest
    $existing = Get-IQAsmMember -Object $m -Name 'outputs'
    $merged = [ordered]@{}
    foreach ($k in @(Get-IQAsmMemberName -Object $existing)) { $merged[$k] = Get-IQAsmMember -Object $existing -Name $k }
    foreach ($k in @('environmentWorkbook', 'reportWorkbook', 'modelWorkbook', 'dataflowWorkbook', 'dataflowWorkbookInRunFolder')) {
        if ($Outputs.ContainsKey($k)) { $merged[$k] = $Outputs[$k] }
    }
    if ($m -is [System.Collections.IDictionary]) { $m['outputs'] = $merged }
    else { $m | Add-Member -NotePropertyName 'outputs' -NotePropertyValue ([PSCustomObject]$merged) -Force }
    try { Save-IQManifest } catch { Write-IQLog -Level Warn -Stage 'Assemble' -Message ("Manifest could not be saved: {0}" -f $_.Exception.Message) }
}

function Invoke-IQAssembleStage {
    <#
    .SYNOPSIS
        Assemble stage body: rebuilds the four workbooks from state files and backup folders and records manifest.outputs.
    .DESCRIPTION
        Each workbook is built independently (a failure of one is recorded with Set-IQItemDone -Status Failed and the
        others are still built), written to a temp file and moved into place, and checkpointed as an Assemble item
        (itemKeys environment-workbook, report-workbook, model-workbook, dataflow-workbook). A workbook whose sheets
        were partly written header-only (Write-IQWorkbook FailedSheets) is recorded as Failed although the file exists;
        every source file that could not be read is recorded as a failed "source-<file>" item, so the stage ends
        CompletedWithErrors and the exit code reflects the incomplete output (brief section 0.4). Returns
        @{ EnvironmentWorkbook; ReportWorkbook; ModelWorkbook; DataflowWorkbook; Built; Failed; SourceErrors; Outputs }.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$OutputFolder,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$ContractPath
    )
    if (-not $script:IQ) { throw 'ImpactIQ context not initialised. Call Initialize-IQContext first.' }
    $stage = 'Assemble'
    $started = [System.Diagnostics.Stopwatch]::StartNew()

    try { Import-Module ImportExcel -ErrorAction Stop }
    catch { throw ("The ImportExcel module is required to build the workbooks (Install-Module ImportExcel -Scope CurrentUser). {0}" -f $_.Exception.Message) }
    if (-not (Get-Command -Name 'Export-Excel' -ErrorAction SilentlyContinue)) { throw 'Export-Excel (ImportExcel module) is not available.' }

    if ([string]::IsNullOrWhiteSpace($OutputFolder)) { $OutputFolder = [string]$script:IQ.BaseFolder }
    if ([string]::IsNullOrWhiteSpace($OutputFolder)) { throw 'No output folder (BaseFolder) is set.' }
    if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }

    $contract = Get-IQSheetContract -Path $ContractPath
    $manifest = $script:IQ.Manifest
    $hasManifest = ($null -ne $manifest)
    Write-IQLog -Level Info -Stage $stage -Message ("Assembling workbooks for run '{0}' into '{1}'" -f [string]$script:IQ.RunId, $OutputFolder)

    $outputs = @{}
    $result = @{ EnvironmentWorkbook = $null; ReportWorkbook = $null; ModelWorkbook = $null; DataflowWorkbook = $null; Built = 0; Failed = 0; SourceErrors = 0; Outputs = $outputs }

    # The Dataflow workbook always goes into THIS run's Dataflow Backups\<RunId> folder (created when it was deleted
    # after Initialize-IQRun) - never into another run's dated folder. Without an active run (helpers called directly)
    # the monolith's newest-dated-folder rule applies, else the output folder.
    $dataflowFolder = $null
    if ($script:IQ.ContainsKey('RunPaths') -and $null -ne $script:IQ.RunPaths -and -not [string]::IsNullOrWhiteSpace([string]$script:IQ.RunPaths['DataflowBackups'])) {
        $dataflowFolder = [string]$script:IQ.RunPaths['DataflowBackups']
        if (-not (Test-Path -LiteralPath $dataflowFolder)) { New-Item -ItemType Directory -Path $dataflowFolder -Force | Out-Null }
    }
    else {
        $dataflowFolder = Get-IQAssembleFolder -Kind 'Dataflow'
        if ([string]::IsNullOrWhiteSpace($dataflowFolder)) { $dataflowFolder = $OutputFolder }
    }

    $jobs = @(
        @{ Key = 'environment-workbook'; Name = 'Power BI Environment Detail.xlsx'; ResultKey = 'EnvironmentWorkbook'; OutputKey = 'environmentWorkbook'
            Build = { param($path) Build-IQEnvironmentWorkbook -Path $path -Contract $contract -Manifest $manifest }
            Path = (Join-Path $OutputFolder 'Power BI Environment Detail.xlsx') },
        @{ Key = 'report-workbook'; Name = 'Report Detail.xlsx'; ResultKey = 'ReportWorkbook'; OutputKey = 'reportWorkbook'
            Build = { param($path) Build-IQReportWorkbook -Path $path -Contract $contract }
            Path = (Join-Path $OutputFolder 'Report Detail.xlsx') },
        @{ Key = 'model-workbook'; Name = 'Model Detail.xlsx'; ResultKey = 'ModelWorkbook'; OutputKey = 'modelWorkbook'
            Build = { param($path) Build-IQModelWorkbook -Path $path -Contract $contract }
            Path = (Join-Path $OutputFolder 'Model Detail.xlsx') },
        @{ Key = 'dataflow-workbook'; Name = 'Dataflow Detail.xlsx'; ResultKey = 'DataflowWorkbook'; OutputKey = 'dataflowWorkbook'
            Build = { param($path) Build-IQDataflowWorkbook -Path (Join-Path $dataflowFolder 'Dataflow Detail.xlsx') -CopyPath $path -Contract $contract }
            Path = (Join-Path $OutputFolder 'Dataflow Detail.xlsx') }
    )

    foreach ($job in $jobs) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $script:IQAssembleSourceErrors = New-Object System.Collections.Generic.List[object]
        try {
            $r = & $job.Build $job.Path
            $finalPath = [string]$job.Path
            $result[$job.ResultKey] = $finalPath
            $outputs[$job.OutputKey] = $finalPath
            $outs = @($finalPath)
            if ($job.Key -eq 'dataflow-workbook' -and $r -and $r['Path'] -and ([string]$r['Path']) -ne $finalPath) {
                $outputs['dataflowWorkbookInRunFolder'] = [string]$r['Path']
                $outs += [string]$r['Path']
            }
            $sheetCount = 0; $rowCount = 0; $failedSheets = @()
            if ($r) { $sheetCount = @($r['Sheets']).Count; $rowCount = [int]$r['Rows']; $failedSheets = @($r['FailedSheets'] | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) }
            $sourceErrors = @()
            if ($null -ne $script:IQAssembleSourceErrors) { $sourceErrors = $script:IQAssembleSourceErrors.ToArray() }
            $result.Built++
            $result.SourceErrors += $sourceErrors.Count
            if ($failedSheets.Count -gt 0) {
                # The file exists (every contract sheet is present, the failed ones header-only) but it is incomplete.
                $msg = ("{0} sheet(s) could not be written and are header-only: {1}" -f $failedSheets.Count, ($failedSheets -join ', '))
                Write-IQLog -Level Error -Stage $stage -Item $job.Name -Message ("{0} sheet(s), {1} row(s) written to '{2}' in {3:n1} s - INCOMPLETE: {4}" -f $sheetCount, $rowCount, $finalPath, $sw.Elapsed.TotalSeconds, $msg)
                if ($hasManifest) {
                    try { Set-IQItemDone -Stage $stage -ItemKey $job.Key -Item $job.Name -Outputs $outs -Status Failed -Method 'ImportExcel' -Message $msg -Data @{ Sheets = $sheetCount; Rows = $rowCount; FailedSheets = $failedSheets } | Out-Null }
                    catch { Write-IQLog -Level Warn -Stage $stage -Message ("Checkpoint could not be written: {0}" -f $_.Exception.Message) }
                }
            }
            else {
                $level = 'Success'
                $note = ''
                if ($sourceErrors.Count -gt 0) { $level = 'Warn'; $note = (' - {0} source file(s) could not be read (see Failures)' -f $sourceErrors.Count) }
                Write-IQLog -Level $level -Stage $stage -Item $job.Name -Message ("{0} sheet(s), {1} row(s) written to '{2}' in {3:n1} s{4}" -f $sheetCount, $rowCount, $finalPath, $sw.Elapsed.TotalSeconds, $note)
                if ($hasManifest) {
                    Set-IQItemDone -Stage $stage -ItemKey $job.Key -Item $job.Name -Outputs $outs -Method 'ImportExcel' -Data @{ Sheets = $sheetCount; Rows = $rowCount; SourceErrors = $sourceErrors.Count } | Out-Null
                }
            }
            # Every unreadable source file is a recorded per-item failure (brief section 0.4): the sheet lacks its rows.
            if ($hasManifest) {
                foreach ($se in $sourceErrors) {
                    $sePath = [string]$se['Path']
                    try { Set-IQItemDone -Stage $stage -ItemKey ('source-' + (Get-IQSafeKey -Value $sePath)) -Item $sePath -Status Failed -Method 'ImportExcel' -Message ([string]$se['Message']) | Out-Null }
                    catch { Write-IQLog -Level Warn -Stage $stage -Message ("Checkpoint could not be written: {0}" -f $_.Exception.Message) }
                }
            }
        }
        catch {
            $result.Failed++
            Write-IQLog -Level Error -Stage $stage -Item $job.Name -Message ("Workbook build failed: {0}" -f $_.Exception.Message) -Exception $_.Exception
            if ($hasManifest) {
                try { Set-IQItemDone -Stage $stage -ItemKey $job.Key -Item $job.Name -Status Failed -Method 'ImportExcel' -Message $_.Exception.Message | Out-Null }
                catch { Write-IQLog -Level Warn -Stage $stage -Message ("Checkpoint could not be written: {0}" -f $_.Exception.Message) }
            }
        }
    }
    $script:IQAssembleSourceErrors = New-Object System.Collections.Generic.List[object]

    if ($hasManifest) { Set-IQManifestOutput -Outputs $outputs }
    $level = 'Success'
    if ($result.Failed -gt 0 -or $result.SourceErrors -gt 0) { $level = 'Warn' }
    Write-IQLog -Level $level -Stage $stage -Message ("Assemble finished: {0} workbook(s) built, {1} failed, {2} unreadable source file(s), {3:n1} s" -f $result.Built, $result.Failed, $result.SourceErrors, $started.Elapsed.TotalSeconds)
    return $result
}
