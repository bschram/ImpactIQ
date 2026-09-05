# ImpactIQ.Dax.ps1 - executeQueries client, INFO.VIEW.* / INFO.* model-detail fallback, DAX reference extractor and usage metrics.
#
# Contract: brief sections 2.7, 7.3, 8.4 and 13 (+ task models-revise).
#   Invoke-IQDaxQuery -WorkspaceId -DatasetId -Dax [-Impersonate]  -> array of flattened row objects (throws on error)
#   Get-IQModelDetailViaDax -Dataset -OutputFolder                   -> writes "<CleanWs> ~ <CleanModel>.csv" and "..._MD.csv"
#   Get-IQModelBackupFileName -Dataset                               -> "<CleanWs> ~ <CleanModel>" (the base name used everywhere)
#   Get-IQDaxReferences -Expression -KnownTables -KnownMeasures      -> approximate direct DAX references (tables/columns/measures)
#   ConvertTo-IQMeasureDependencyRows / Write-IQCsvFile / New-IQModelDetailRow / Get-IQModelDetailHeader /
#   Get-IQMeasureDependencyHeader                                    -> CSV building blocks shared with ImpactIQ.Bim.ps1
#   Get-IQUsageMetrics -WorkspaceId -WorkspaceName [-Days 30]        -> usage-metrics rows for the Extras stage
#
# Windows PowerShell 5.1 and PowerShell 7 compatible. Loaded by dot-sourcing from ImpactIQ.ps1, so $script:IQ is the
# shared context created by Initialize-IQContext (ImpactIQ.Common.ps1). Nothing here is Windows-only.
#
# Cross-module functions used (brief section 2): Write-IQLog, Get-IQCleanName, Get-IQSafeKey, Invoke-IQApi,
# ConvertTo-IQJsonFile, ConvertFrom-IQJsonFile. Private helpers are prefixed *-IQDax* / *-IQCsv* and are not part of
# the contract.
#
# executeQueries facts that shape Get-IQModelDetailViaDax (audit research-dax.md): the JSON endpoint officially supports
# only DAX (raw INFO.* / DMV queries are rejected with HTTP 400, engine error 3239575574, on most tenants since 2025),
# while INFO.VIEW.TABLES/COLUMNS/MEASURES/RELATIONSHIPS are widely reported to work with Build permission (Pro included).
# [Expression] columns are blank unless the caller has write permission on the model; INFO.CALCDEPENDENCY,
# INFO.TABLEPERMISSIONS and INFO.ANNOTATIONS need write permission even over XMLA.
#
# TOM / TMSCHEMA enum codes reproduced as strings (raw INFO.* returns the integer codes; INFO.VIEW.* returns the names):
#   RelationshipEndCardinality  0=None 1=One 2=Many
#   CrossFilteringBehavior      1=OneDirection 2=BothDirections 3=Automatic
#   SecurityFilteringBehavior   1=OneDirection 2=BothDirections 3=None
#   ModeType (partition Mode)   0=Import 1=DirectQuery 2=Default 3=Push 4=Dual 5=DirectLake
#   PartitionSourceType (Type)  1=Query 2=Calculated 3=None 4=M 5=Entity 6=PolicyRange 7=CalculationGroup 8=Inferred
#   ColumnType (Type)           1=Data 2=Calculated 3=RowNumber 4=CalculatedTableColumn
#   ModelPermission (roles)     1=None 2=Read 3=ReadRefresh 4=Refresh 5=Administrator
#   MetadataPermission          1=Default 2=None 3=Read
# ConvertTo-IQDaxEnumName accepts both the code and the name.

function Get-IQModelBackupFileName {
    <#
    .SYNOPSIS
    Returns the "<CleanWs> ~ <CleanModel>" base name used for every model backup / detail file of a dataset.
    .DESCRIPTION
    Applies Get-IQCleanName (the exact monolith sanitiser) to WorkspaceName and DatasetName and joins them with " ~ ".
    Pass either a dataset object (properties WorkspaceName / DatasetName, monolith names) or the two names explicitly.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)][AllowNull()]$Dataset,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$WorkspaceName,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$DatasetName
    )
    if ($null -ne $Dataset) {
        if (-not $PSBoundParameters.ContainsKey('WorkspaceName')) { $WorkspaceName = [string](Get-IQDaxMember -Object $Dataset -Name 'WorkspaceName') }
        if (-not $PSBoundParameters.ContainsKey('DatasetName')) { $DatasetName = [string](Get-IQDaxMember -Object $Dataset -Name 'DatasetName') }
    }
    return ((Get-IQCleanName -Name $WorkspaceName) + ' ~ ' + (Get-IQCleanName -Name $DatasetName))
}

function Get-IQDaxMember {
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

function ConvertTo-IQDaxText {
    <#
    .SYNOPSIS
    Converts a DAX result value to the string written into the CSV: $null -> '', booleans -> True/False (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false, Position = 0)][AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [bool]) { if ($Value) { return 'True' } else { return 'False' } }
    if ($Value -is [System.DBNull]) { return '' }
    if ($Value -is [double] -or $Value -is [decimal] -or $Value -is [single]) {
        return ([string]$Value)
    }
    return [string]$Value
}

function ConvertTo-IQDaxEnumName {
    <#
    .SYNOPSIS
    Maps an INFO.* integer enum code (or an already-named value) to its TOM enum name (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Value,
        [Parameter(Mandatory = $true)][hashtable]$Map,
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$Default = ''
    )
    if ($null -eq $Value) { return $Default }
    $text = ([string]$Value).Trim()
    if ($text -eq '') { return $Default }
    $code = 0
    if ([int]::TryParse($text, [ref]$code)) {
        if ($Map.ContainsKey($code)) { return [string]$Map[$code] }
        return $text
    }
    foreach ($name in $Map.Values) {
        if ([string]$name -ieq $text) { return [string]$name }
    }
    return $text
}

function Test-IQDaxGuid {
    <#
    .SYNOPSIS
    True when the value looks like a GUID (workspace ids of real workspaces; pseudo workspaces are names) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value.Trim() -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')
}

function Get-IQDaxDatasetPath {
    <#
    .SYNOPSIS
    Relative API path for a dataset: groups/{ws}/datasets/{id} for a real workspace, datasets/{id} for My Workspace (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$WorkspaceId,
        [Parameter(Mandatory = $true)][string]$DatasetId
    )
    if (Test-IQDaxGuid -Value $WorkspaceId) { return ('groups/' + $WorkspaceId.Trim() + '/datasets/' + $DatasetId.Trim()) }
    return ('datasets/' + $DatasetId.Trim())
}

function Get-IQDaxErrorText {
    <#
    .SYNOPSIS
    Extracts the human-readable DAX/engine error from an executeQueries error object or a raw 400 body (private).
    .DESCRIPTION
    Handles the service shapes { error: { code, message, "pbi.error": { code, details: [ { code, detail: { value } } ] } } }
    and the per-result { error: {...} } object. Returns '' when nothing useful is found.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$ErrorObject)
    if ($null -eq $ErrorObject) { return '' }
    $obj = $ErrorObject
    if ($obj -is [string]) {
        $text = $obj.Trim()
        if ($text -eq '') { return '' }
        try { $obj = $text | ConvertFrom-Json -ErrorAction Stop } catch { return $text }
    }
    $parts = @()
    try {
        $err = Get-IQDaxMember -Object $obj -Name 'error'
        if ($null -eq $err) { $err = $obj }
        $code = Get-IQDaxMember -Object $err -Name 'code'
        $message = Get-IQDaxMember -Object $err -Name 'message'
        if ($code) { $parts += [string]$code }
        if ($message) { $parts += [string]$message }
        $pbi = Get-IQDaxMember -Object $err -Name 'pbi.error'
        if ($null -ne $pbi) {
            $details = Get-IQDaxMember -Object $pbi -Name 'details'
            foreach ($d in @($details)) {
                if ($null -eq $d) { continue }
                $detail = Get-IQDaxMember -Object $d -Name 'detail'
                $value = $null
                if ($null -ne $detail) { $value = Get-IQDaxMember -Object $detail -Name 'value' }
                if ($null -eq $value) { $value = Get-IQDaxMember -Object $d -Name 'message' }
                if ($null -ne $value -and ([string]$value).Trim() -ne '') {
                    $dcode = [string](Get-IQDaxMember -Object $d -Name 'code')
                    if ($dcode -and $dcode -ne 'DetailsMessage') { $parts += ($dcode + ': ' + [string]$value) }
                    else { $parts += [string]$value }
                }
            }
        }
        $inner = Get-IQDaxMember -Object $err -Name 'details'
        if ($null -ne $inner -and $parts.Count -le 1) {
            foreach ($d in @($inner)) {
                $m = Get-IQDaxMember -Object $d -Name 'message'
                if ($m) { $parts += [string]$m }
            }
        }
    }
    catch { $parts += [string]$obj }
    $result = (($parts | Where-Object { $_ -and $_.Trim() }) -join ' | ')
    if ($result.Length -gt 1000) { $result = $result.Substring(0, 1000) + '...' }
    return $result
}

function ConvertFrom-IQDaxRows {
    <#
    .SYNOPSIS
    Strips the [Column] / Table[Column] / 'Table'[Column] wrapper from executeQueries row property names.
    .DESCRIPTION
    Column names are shortened to the bare column name when that is unambiguous within the result set; when two
    source columns would collapse to the same short name their original names are kept. Returns an array of
    PSCustomObjects with properties in the original order.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false, Position = 0)][AllowNull()][AllowEmptyCollection()]$Rows)
    $list = @($Rows | Where-Object { $null -ne $_ })
    if ($list.Count -eq 0) { return @() }

    # Build the name map from the union of property names (includeNulls=true makes every row carry every column,
    # but be defensive about rows that omit nulls).
    $names = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($row in $list) {
        foreach ($p in $row.PSObject.Properties) {
            if (-not $seen.ContainsKey($p.Name)) { $seen[$p.Name] = $true; $names.Add($p.Name) }
        }
    }
    $shortOf = @{}
    $counts = @{}
    foreach ($n in $names) {
        $short = $n
        if ($n -match '^\[(.*)\]$') { $short = $Matches[1] }
        elseif ($n -match "^(?:'(.*)'|([^'\[\]]+))\[(.*)\]$") { $short = $Matches[3] }
        $shortOf[$n] = $short
        if ($counts.ContainsKey($short)) { $counts[$short] = [int]$counts[$short] + 1 } else { $counts[$short] = 1 }
    }
    foreach ($n in @($shortOf.Keys)) {
        if ([int]$counts[$shortOf[$n]] -gt 1) { $shortOf[$n] = $n }
    }

    $out = New-Object System.Collections.Generic.List[object]
    foreach ($row in $list) {
        $o = [ordered]@{}
        foreach ($n in $names) {
            $prop = $row.PSObject.Properties[$n]
            $value = $null
            if ($null -ne $prop) { $value = $prop.Value }
            $o[$shortOf[$n]] = $value
        }
        $out.Add([PSCustomObject]$o)
    }
    return $out.ToArray()
}

function Invoke-IQDaxQuery {
    <#
    .SYNOPSIS
    Runs one DAX query against a dataset through the executeQueries REST API and returns the flattened rows.
    .DESCRIPTION
    POST groups/{ws}/datasets/{id}/executeQueries (datasets/{id}/executeQueries for My Workspace) with
    { queries:[{query}], serializerSettings:{ includeNulls:true } } (+ impersonatedUserName when -Impersonate is given).
    Rows come back with the [Table].[Column] wrapper stripped (ConvertFrom-IQDaxRows). Throws a System.Exception with
    the engine error text when the service reports an error (HTTP 400 body or results[].error); a 403/404/400
    handled by Invoke-IQApi (which returns $null) is surfaced as an exception too, using $script:IQ.LastHttpError
    when the Http module records it. Logs a Warn when the 100 000-row cap of the API is hit.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$WorkspaceId,
        [Parameter(Mandatory = $true)][string]$DatasetId,
        [Parameter(Mandatory = $true)][string]$Dax,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Impersonate,
        [Parameter(Mandatory = $false)][int]$TimeoutSec = 600,
        [Parameter(Mandatory = $false)][string]$Stage = 'Dax',
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Item
    )
    $path = (Get-IQDaxDatasetPath -WorkspaceId $WorkspaceId -DatasetId $DatasetId) + '/executeQueries'
    $body = [ordered]@{
        queries            = @([ordered]@{ query = $Dax })
        serializerSettings = [ordered]@{ includeNulls = $true }
    }
    if (-not [string]::IsNullOrWhiteSpace($Impersonate)) { $body['impersonatedUserName'] = $Impersonate.Trim() }

    $queryLabel = ($Dax -replace '\s+', ' ').Trim()
    if ($queryLabel.Length -gt 120) { $queryLabel = $queryLabel.Substring(0, 120) + '...' }
    Write-IQLog -Level Debug -Stage $Stage -Item $Item -Message ("DAX: {0}" -f $queryLabel)

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    if ($script:IQ -and $script:IQ.ContainsKey('LastHttpError')) { $script:IQ.LastHttpError = $null }
    $response = Invoke-IQApi -Method POST -Path $path -Body $body -NoPaging -TimeoutSec $TimeoutSec -Stage $Stage
    $stopwatch.Stop()

    if ($null -eq $response) {
        # Invoke-IQApi returns $null for 400/403/404 (logged with the body at Warn). Surface a useful message.
        $status = $null
        $detail = ''
        try {
            if ($script:IQ -and $script:IQ.ContainsKey('LastHttpError') -and $null -ne $script:IQ.LastHttpError) {
                $status = Get-IQDaxMember -Object $script:IQ.LastHttpError -Name 'StatusCode'
                $detail = Get-IQDaxErrorText -ErrorObject (Get-IQDaxMember -Object $script:IQ.LastHttpError -Name 'Body')
            }
        }
        catch { $detail = '' }
        $message = 'executeQueries returned no response for ' + $path
        if ($null -ne $status) { $message = ('executeQueries returned HTTP {0} for {1}' -f $status, $path) }
        else { $message += ' (HTTP 400/403/404 - see the Warn line above for the response body)' }
        if ($detail) { $message += ': ' + $detail }
        else { $message += '. Check Build permission on the dataset and the "Semantic Model Execute Queries REST API" tenant setting.' }
        throw (New-Object System.Exception($message))
    }

    $results = Get-IQDaxMember -Object $response -Name 'results'
    if ($null -eq $results) {
        $topError = Get-IQDaxMember -Object $response -Name 'error'
        if ($null -ne $topError) { throw (New-Object System.Exception('executeQueries error: ' + (Get-IQDaxErrorText -ErrorObject $response))) }
        throw (New-Object System.Exception('executeQueries returned an unexpected response (no results) for ' + $path))
    }
    $first = @($results)[0]
    $resultError = Get-IQDaxMember -Object $first -Name 'error'
    if ($null -ne $resultError) {
        throw (New-Object System.Exception('DAX query error: ' + (Get-IQDaxErrorText -ErrorObject $first)))
    }
    $tables = @(Get-IQDaxMember -Object $first -Name 'tables')
    $rows = @()
    if ($tables.Count -gt 0 -and $null -ne $tables[0]) {
        $rows = @(Get-IQDaxMember -Object $tables[0] -Name 'rows')
    }
    $rows = @($rows | Where-Object { $null -ne $_ })
    if ($rows.Count -ge 100000) {
        Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message ("DAX result hit the executeQueries 100 000-row cap ({0}); the result may be truncated." -f $queryLabel)
    }
    Write-IQLog -Level Debug -Stage $Stage -Item $Item -Message ("DAX rows: {0} ({1} ms)" -f $rows.Count, $stopwatch.ElapsedMilliseconds)
    return @(ConvertFrom-IQDaxRows -Rows $rows)
}

function Get-IQModelDetailHeader {
    <#
    .SYNOPSIS
    The 20 column names of "<CleanWs> ~ <CleanModel>.csv" (Semantic Models sheet), in order (brief section 13).
    #>
    [CmdletBinding()]
    param()
    return @('Type', 'Table', 'Name', 'FormatString', 'DisplayFolder', 'Description', 'IsHidden', 'TableStorageMode', 'Expression',
        'ModelAsOfDate', 'ModelName', 'ModelID', 'RelationshipFromTable', 'RelationshipFromColumn', 'RelationshipToTable',
        'RelationshipToColumn', 'RelationshipStatus', 'RelationshipFromCardinality', 'RelationshipToCardinality',
        'RelationshipCrossFilteringBehavior')
}

function Get-IQMeasureDependencyHeader {
    <#
    .SYNOPSIS
    The 7 column names of "<CleanWs> ~ <CleanModel>_MD.csv" (Measure Dependencies sheet), in order (brief section 13).
    #>
    [CmdletBinding()]
    param()
    return @('ObjectName', 'ObjectType', 'DependsOn', 'DependsOnType', 'ModelAsOfDate', 'ModelName', 'ModelID')
}

function Format-IQCsvField {
    <#
    .SYNOPSIS
    Quotes one CSV field exactly like the csx FormatField lambda: always quoted, embedded quotes doubled (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false, Position = 0)][AllowNull()]$Value)
    $text = ConvertTo-IQDaxText -Value $Value
    if ($text -eq '') { return '""' }
    return ('"' + $text.Replace('"', '""') + '"')
}

function Write-IQCsvFile {
    <#
    .SYNOPSIS
    Writes a fully quoted CSV (CRLF, UTF-8 with BOM) atomically: <path>.tmp then Move-Item -Force (private).
    .DESCRIPTION
    -Rows is an array of hashtables/ordered dictionaries keyed by header name (missing keys -> ""), or of object[] in
    header order. Every field is quoted and embedded double quotes are doubled, matching the Tabular Editor scripts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string[]]$Header,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()]$Rows
    )
    $sb = New-Object System.Text.StringBuilder
    $headerLine = (($Header | ForEach-Object { '"' + $_.Replace('"', '""') + '"' }) -join ',')
    [void]$sb.Append($headerLine).Append("`r`n")
    $count = 0
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $fields = New-Object System.Collections.Generic.List[string]
        if ($row -is [System.Collections.IDictionary]) {
            foreach ($h in $Header) {
                $v = $null
                if ($row.Contains($h)) { $v = $row[$h] }
                $fields.Add((Format-IQCsvField -Value $v))
            }
        }
        elseif ($row -is [System.Collections.IList]) {
            for ($i = 0; $i -lt $Header.Count; $i++) {
                $v = $null
                if ($i -lt $row.Count) { $v = $row[$i] }
                $fields.Add((Format-IQCsvField -Value $v))
            }
        }
        else {
            foreach ($h in $Header) {
                $fields.Add((Format-IQCsvField -Value (Get-IQDaxMember -Object $row -Name $h)))
            }
        }
        [void]$sb.Append(($fields -join ',')).Append("`r`n")
        $count++
    }
    $dir = Split-Path -Path $Path -Parent
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = $Path + '.tmp'
    $encoding = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($tmp, $sb.ToString(), $encoding)
    Move-Item -LiteralPath $tmp -Destination $Path -Force
    return $count
}

function New-IQModelDetailRow {
    <#
    .SYNOPSIS
    Builds one 20-column Semantic Models row (ordered dictionary) from named fields; unspecified fields are "" (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Type,
        [Parameter(Mandatory = $true)][hashtable]$Common,
        [Parameter(Mandatory = $false)][AllowNull()][hashtable]$Fields
    )
    $row = [ordered]@{}
    foreach ($h in (Get-IQModelDetailHeader)) { $row[$h] = '' }
    $row['Type'] = $Type
    $row['ModelAsOfDate'] = [string]$Common['ModelAsOfDate']
    $row['ModelName'] = [string]$Common['ModelName']
    $row['ModelID'] = [string]$Common['ModelID']
    if ($null -ne $Fields) {
        foreach ($k in $Fields.Keys) { $row[$k] = ConvertTo-IQDaxText -Value $Fields[$k] }
    }
    return $row
}

function ConvertTo-IQDaxTableRef {
    <#
    .SYNOPSIS
    DAX full name of a table, 'Table' with embedded single quotes doubled (Tabular Editor DaxObjectFullName rule) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Table)
    if ($null -eq $Table) { $Table = '' }
    return ("'" + $Table.Replace("'", "''") + "'")
}

function ConvertTo-IQDaxBracketRef {
    <#
    .SYNOPSIS
    [Name] with embedded ']' doubled (Tabular Editor DaxObjectFullName rule) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Name)
    if ($null -eq $Name) { $Name = '' }
    return ('[' + $Name.Replace(']', ']]') + ']')
}

function Get-IQDaxExtractFolder {
    <#
    .SYNOPSIS
    Returns (and creates) <RunPath>\extracts\dax\<safeKey> where raw INFO.* results are cached (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Key)
    $root = $null
    if ($script:IQ) {
        if ($script:IQ.ContainsKey('RunPaths') -and $script:IQ.RunPaths -and $script:IQ.RunPaths.Extracts) { $root = [string]$script:IQ.RunPaths.Extracts }
        elseif ($script:IQ.ContainsKey('RunPath') -and -not [string]::IsNullOrWhiteSpace([string]$script:IQ.RunPath)) { $root = Join-Path ([string]$script:IQ.RunPath) 'extracts' }
        elseif ($script:IQ.ContainsKey('StatePath') -and -not [string]::IsNullOrWhiteSpace([string]$script:IQ.StatePath)) { $root = Join-Path ([string]$script:IQ.StatePath) 'extracts' }
    }
    if (-not $root) { $root = Join-Path ([System.IO.Path]::GetTempPath()) 'ImpactIQ\extracts' }
    $folder = Join-Path (Join-Path $root 'dax') (Get-IQSafeKey -Value $Key)
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    return $folder
}

function Get-IQDaxInfoRowSet {
    <#
    .SYNOPSIS
    Runs one INFO.* query for a dataset, caching the flattened rows as JSON under extracts\dax\<key>\<name>.json (private).
    .DESCRIPTION
    When the JSON already exists (and -Refresh is not set) it is loaded instead of re-querying, so a re-run can rebuild
    the CSVs without touching the service. Throws on query failure (the caller decides whether that is fatal).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ExtractFolder,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$WorkspaceId,
        [Parameter(Mandatory = $true)][string]$DatasetId,
        [Parameter(Mandatory = $true)][string]$Dax,
        [Parameter(Mandatory = $false)][switch]$Refresh,
        [Parameter(Mandatory = $false)][string]$Stage = 'ModelDetail',
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Item
    )
    $file = Join-Path $ExtractFolder ($Name + '.json')
    if (-not $Refresh -and (Test-Path -LiteralPath $file)) {
        try {
            $cached = @(ConvertFrom-IQJsonFile -Path $file)
            Write-IQLog -Level Debug -Stage $Stage -Item $Item -Message ("Reusing cached DAX extract {0} ({1} rows)" -f $Name, $cached.Count)
            return $cached
        }
        catch {
            Write-IQLog -Level Debug -Stage $Stage -Item $Item -Message ("Cached extract {0} unreadable, re-querying: {1}" -f $Name, $_.Exception.Message)
        }
    }
    $rows = @(Invoke-IQDaxQuery -WorkspaceId $WorkspaceId -DatasetId $DatasetId -Dax $Dax -Stage $Stage -Item $Item)
    try { ConvertTo-IQJsonFile -Object $rows -Path $file }
    catch { Write-IQLog -Level Warn -Stage $Stage -Item $Item -Message ("Could not cache DAX extract {0}: {1}" -f $Name, $_.Exception.Message) }
    return $rows
}

function Get-IQDaxIndex {
    <#
    .SYNOPSIS
    Builds a hashtable keyed by the string form of an ID column from INFO.* rows (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()]$Rows,
        [Parameter(Mandatory = $false)][string]$IdName = 'ID'
    )
    $index = @{}
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $id = Get-IQDaxMember -Object $r -Name $IdName
        if ($null -eq $id) { continue }
        $index[[string]$id] = $r
    }
    return $index
}

function Get-IQDaxIndexedName {
    <#
    .SYNOPSIS
    Looks up the display name of an object by ID in an index built by Get-IQDaxIndex; '' when unknown (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Index,
        [Parameter(Mandatory = $false)][AllowNull()]$Id,
        [Parameter(Mandatory = $false)][string[]]$NameProperty = @('Name')
    )
    if ($null -eq $Id) { return '' }
    $key = [string]$Id
    if (-not $Index.ContainsKey($key)) { return '' }
    $obj = $Index[$key]
    foreach ($p in $NameProperty) {
        $v = Get-IQDaxMember -Object $obj -Name $p
        if ($null -ne $v -and ([string]$v) -ne '') { return [string]$v }
    }
    return ''
}

function Get-IQDaxColumnName {
    <#
    .SYNOPSIS
    Display name of an INFO.COLUMNS row: ExplicitName, else InferredName (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Column)
    if ($null -eq $Column) { return '' }
    $n = Get-IQDaxMember -Object $Column -Name 'ExplicitName'
    if ($null -ne $n -and ([string]$n) -ne '') { return [string]$n }
    $n = Get-IQDaxMember -Object $Column -Name 'InferredName'
    if ($null -ne $n) { return [string]$n }
    $n = Get-IQDaxMember -Object $Column -Name 'Name'
    if ($null -ne $n) { return [string]$n }
    return ''
}

function ConvertTo-IQDaxCodeOnly {
    <#
    .SYNOPSIS
    Blanks string literals and comments in a DAX expression so identifier regexes only see code (private).
    .DESCRIPTION
    "..." literals ("" escapes) become a single space; // and -- line comments and /* */ block comments are removed;
    'Table' and [Name] identifiers are copied verbatim ('' and ]] escapes honoured) so comment markers inside them
    are not mistaken for comments.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false, Position = 0)][AllowNull()][AllowEmptyString()][string]$Expression)
    if ([string]::IsNullOrEmpty($Expression)) { return '' }
    $sb = New-Object System.Text.StringBuilder
    $text = $Expression
    $n = $text.Length
    $i = 0
    while ($i -lt $n) {
        $c = $text[$i]
        $next = [char]0
        if ($i + 1 -lt $n) { $next = $text[$i + 1] }
        if ($c -eq '"') {
            $i++
            while ($i -lt $n) {
                if ($text[$i] -eq '"') {
                    if ($i + 1 -lt $n -and $text[$i + 1] -eq '"') { $i += 2; continue }
                    $i++
                    break
                }
                $i++
            }
            [void]$sb.Append(' ')
            continue
        }
        if ($c -eq "'" -or $c -eq '[') {
            $close = "'"
            if ($c -eq '[') { $close = ']' }
            $start = $i
            $i++
            while ($i -lt $n) {
                if ($text[$i] -eq $close) {
                    if ($i + 1 -lt $n -and $text[$i + 1] -eq $close) { $i += 2; continue }
                    $i++
                    break
                }
                $i++
            }
            [void]$sb.Append($text.Substring($start, $i - $start))
            continue
        }
        if (($c -eq '/' -and $next -eq '/') -or ($c -eq '-' -and $next -eq '-')) {
            while ($i -lt $n -and $text[$i] -ne "`n" -and $text[$i] -ne "`r") { $i++ }
            [void]$sb.Append(' ')
            continue
        }
        if ($c -eq '/' -and $next -eq '*') {
            $end = $text.IndexOf('*/', $i + 2)
            if ($end -lt 0) { $i = $n } else { $i = $end + 2 }
            [void]$sb.Append(' ')
            continue
        }
        [void]$sb.Append($c)
        $i++
    }
    return $sb.ToString()
}

function Get-IQDaxReferences {
    <#
    .SYNOPSIS
    Approximate DAX reference extractor: the tables, columns and measures a DAX expression refers to directly (regex based).
    .DESCRIPTION
    Stand-in for Tabular Editor's DependsOn when TE2 cannot run (Bim parser path, DAX fallback without INFO.CALCDEPENDENCY).
    String literals and comments are ignored (ConvertTo-IQDaxCodeOnly); then 'Table'[Name], Table[Name], [Name] and
    standalone 'Table' / Table references are matched. Resolution (identifiers are case-insensitive, like DAX):
      - 'Table'[Name] / Table[Name]: a column of that table, unless Name is a known measure (then Measure [Name]).
      - [Name]: a known measure first; else a column of -CurrentTable; else a column that exists in exactly one known
        table; else ignored.
      - standalone 'Table' or an unquoted identifier that is a known table and is not followed by "(" or "[" and is not
        a VAR declared in the expression: a Table dependency (DependsOnType CalculationGroupTable when the table is a
        calculation group).
    Approximate by design: variables shadowing table names that are not declared with VAR in the same expression,
    references built dynamically (TREATAS, string concatenation, user hierarchies), function-name collisions and
    indirect (transitive) dependencies are not resolved. Only DIRECT references are returned, matching the csx.
    Output objects: @{ DependsOn (TE2 DaxObjectFullName: 'Table'[Column], [Measure], 'Table'); DependsOnType
    (Column|Measure|Table|CalculationGroupTable); Table; Name }, de-duplicated, in order of first appearance.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)][AllowNull()][AllowEmptyString()][string]$Expression,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()][string[]]$KnownTables,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()][string[]]$KnownMeasures,
        [Parameter(Mandatory = $false)][AllowNull()][hashtable]$KnownColumns,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$CurrentTable,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()][string[]]$CalculationGroupTables
    )
    $out = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Expression)) { return $out.ToArray() }

    # Case-insensitive lookups (original spelling preserved for output).
    $tableCase = @{}
    foreach ($t in @($KnownTables)) { if (-not [string]::IsNullOrEmpty($t)) { $tableCase[$t.ToLowerInvariant()] = $t } }
    $measureCase = @{}
    foreach ($m in @($KnownMeasures)) { if (-not [string]::IsNullOrEmpty($m)) { $measureCase[$m.ToLowerInvariant()] = $m } }
    $calcGroupSet = @{}
    foreach ($cg in @($CalculationGroupTables)) { if (-not [string]::IsNullOrEmpty($cg)) { $calcGroupSet[$cg.ToLowerInvariant()] = $true } }
    $columnsByTable = @{}
    $ownersByColumn = @{}
    $haveColumns = $false
    if ($null -ne $KnownColumns) {
        foreach ($tk in $KnownColumns.Keys) {
            $tName = [string]$tk
            if ($tName -eq '') { continue }
            $haveColumns = $true
            $lt = $tName.ToLowerInvariant()
            if (-not $tableCase.ContainsKey($lt)) { $tableCase[$lt] = $tName }
            $map = @{}
            foreach ($cn in @($KnownColumns[$tk])) {
                if ([string]::IsNullOrEmpty($cn)) { continue }
                $lc = ([string]$cn).ToLowerInvariant()
                $map[$lc] = [string]$cn
                if (-not $ownersByColumn.ContainsKey($lc)) { $ownersByColumn[$lc] = New-Object System.Collections.Generic.List[string] }
                $ownersByColumn[$lc].Add($tName)
            }
            $columnsByTable[$lt] = $map
        }
    }
    $currentLower = ''
    if (-not [string]::IsNullOrEmpty($CurrentTable)) { $currentLower = $CurrentTable.ToLowerInvariant() }

    $code = ConvertTo-IQDaxCodeOnly -Expression $Expression
    # VAR names shadow tables inside the expression (approximation: any VAR declared anywhere in the expression).
    $varNames = @{}
    foreach ($vm in [regex]::Matches($code, '(?i)\bVAR\s+([A-Za-z_]\w*)')) { $varNames[$vm.Groups[1].Value.ToLowerInvariant()] = $true }

    $seen = @{}
    $add = {
        param($type, $table, $name)
        $full = ''
        $depType = $type
        switch ($type) {
            'Measure' { $full = ConvertTo-IQDaxBracketRef -Name $name }
            'Column' { $full = (ConvertTo-IQDaxTableRef -Table $table) + (ConvertTo-IQDaxBracketRef -Name $name) }
            'Table' {
                $full = ConvertTo-IQDaxTableRef -Table $table
                if ($calcGroupSet.ContainsKey($table.ToLowerInvariant())) { $depType = 'CalculationGroupTable' }
            }
        }
        $key = $depType + '|' + $full
        if ($seen.ContainsKey($key)) { return }
        $seen[$key] = $true
        $out.Add([PSCustomObject]@{ DependsOn = $full; DependsOnType = $depType; Table = $table; Name = $name })
    }
    $resolveTableName = {
        param($raw)
        $lt = $raw.ToLowerInvariant()
        if ($tableCase.ContainsKey($lt)) { return $tableCase[$lt] }
        return $raw
    }
    $resolveColumnName = {
        param($table, $raw)
        $lt = $table.ToLowerInvariant()
        $lc = $raw.ToLowerInvariant()
        if ($columnsByTable.ContainsKey($lt) -and $columnsByTable[$lt].ContainsKey($lc)) { return $columnsByTable[$lt][$lc] }
        return $raw
    }
    $tableHasColumn = {
        param($table, $raw)
        $lt = $table.ToLowerInvariant()
        return ($columnsByTable.ContainsKey($lt) -and $columnsByTable[$lt].ContainsKey($raw.ToLowerInvariant()))
    }

    $pattern = "(?<qt>'(?:[^']|'')+')(?:\s*\[(?<qc>(?:[^\]]|\]\])+)\])?|(?<![\w\]\)'`"\.])(?<ut>[A-Za-z_]\w*)(?:\s*\[(?<uc>(?:[^\]]|\]\])+)\])?|\[(?<bc>(?:[^\]]|\]\])+)\]"
    foreach ($m in [regex]::Matches($code, $pattern)) {
        $tableRaw = $null
        $bracketRaw = $null
        $quotedTable = $false
        if ($m.Groups['qt'].Success) {
            $q = $m.Groups['qt'].Value
            $tableRaw = $q.Substring(1, $q.Length - 2).Replace("''", "'")
            $quotedTable = $true
            if ($m.Groups['qc'].Success) { $bracketRaw = $m.Groups['qc'].Value.Replace(']]', ']') }
        }
        elseif ($m.Groups['ut'].Success) {
            $tableRaw = $m.Groups['ut'].Value
            if ($m.Groups['uc'].Success) { $bracketRaw = $m.Groups['uc'].Value.Replace(']]', ']') }
        }
        elseif ($m.Groups['bc'].Success) {
            $bracketRaw = $m.Groups['bc'].Value.Replace(']]', ']')
        }

        if ($null -ne $tableRaw -and -not $quotedTable) {
            # Unquoted identifier: only a table when known (and not a VAR); otherwise it is a function/keyword/variable.
            $lt = $tableRaw.ToLowerInvariant()
            if ($varNames.ContainsKey($lt) -or -not $tableCase.ContainsKey($lt)) {
                if ($null -ne $bracketRaw) { $tableRaw = $null } else { continue }
            }
        }

        if ($null -ne $tableRaw -and $null -ne $bracketRaw) {
            $table = & $resolveTableName $tableRaw
            $lb = $bracketRaw.ToLowerInvariant()
            if ((& $tableHasColumn $table $bracketRaw)) { & $add 'Column' $table (& $resolveColumnName $table $bracketRaw) }
            elseif ($measureCase.ContainsKey($lb)) { & $add 'Measure' $table $measureCase[$lb] }
            else { & $add 'Column' $table $bracketRaw }
            continue
        }
        if ($null -ne $tableRaw) {
            # Standalone table reference: not when followed by "(" (function) or "[" (handled above).
            $rest = $code.Substring($m.Index + $m.Length).TrimStart()
            if ($rest.Length -gt 0 -and ($rest[0] -eq '(' -or $rest[0] -eq '[')) { continue }
            if (-not $quotedTable -and $rest.Length -gt 0 -and $rest[0] -eq '.') { continue }
            $table = & $resolveTableName $tableRaw
            if ($quotedTable -or $tableCase.ContainsKey($tableRaw.ToLowerInvariant())) { & $add 'Table' $table $table }
            continue
        }
        if ($null -ne $bracketRaw) {
            $lb = $bracketRaw.ToLowerInvariant()
            if ($measureCase.ContainsKey($lb)) { & $add 'Measure' '' $measureCase[$lb]; continue }
            if ($currentLower -ne '' -and (& $tableHasColumn $CurrentTable $bracketRaw)) { & $add 'Column' (& $resolveTableName $CurrentTable) (& $resolveColumnName $CurrentTable $bracketRaw); continue }
            if ($ownersByColumn.ContainsKey($lb) -and $ownersByColumn[$lb].Count -eq 1) {
                $owner = $ownersByColumn[$lb][0]
                & $add 'Column' $owner (& $resolveColumnName $owner $bracketRaw)
                continue
            }
            if (-not $haveColumns -and $currentLower -ne '' -and $measureCase.Count -gt 0) {
                # No column catalogue: a bracket that is not a known measure is assumed to be a column of the current table.
                & $add 'Column' (& $resolveTableName $CurrentTable) $bracketRaw
            }
        }
    }
    return $out.ToArray()
}

function ConvertTo-IQMeasureDependencyRows {
    <#
    .SYNOPSIS
    Builds "_MD.csv" rows (Get-IQMeasureDependencyHeader order) for measures / calculated columns / calculation items via Get-IQDaxReferences.
    .DESCRIPTION
    -Objects is an array of @{ ObjectName; ObjectType (Measure|CalculatedColumn|CalculationItem); Table; Expression }
    in the csx emission order (measures, calculated columns, calculation items). Returns ordered dictionaries with
    ObjectName, ObjectType, DependsOn, DependsOnType, ModelAsOfDate, ModelName, ModelID (from -Common).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()]$Objects,
        [Parameter(Mandatory = $true)][hashtable]$Common,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()][string[]]$KnownTables,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()][string[]]$KnownMeasures,
        [Parameter(Mandatory = $false)][AllowNull()][hashtable]$KnownColumns,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()][string[]]$CalculationGroupTables
    )
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($o in @($Objects)) {
        if ($null -eq $o) { continue }
        $expr = [string](Get-IQDaxMember -Object $o -Name 'Expression')
        if ([string]::IsNullOrWhiteSpace($expr)) { continue }
        $refs = @(Get-IQDaxReferences -Expression $expr -KnownTables $KnownTables -KnownMeasures $KnownMeasures -KnownColumns $KnownColumns -CurrentTable ([string](Get-IQDaxMember -Object $o -Name 'Table')) -CalculationGroupTables $CalculationGroupTables)
        foreach ($r in $refs) {
            $rows.Add([ordered]@{
                    ObjectName = [string](Get-IQDaxMember -Object $o -Name 'ObjectName'); ObjectType = [string](Get-IQDaxMember -Object $o -Name 'ObjectType')
                    DependsOn = $r.DependsOn; DependsOnType = $r.DependsOnType
                    ModelAsOfDate = $Common.ModelAsOfDate; ModelName = $Common.ModelName; ModelID = $Common.ModelID
                })
        }
    }
    return $rows.ToArray()
}

function Test-IQDaxInfoUnsupportedMessage {
    <#
    .SYNOPSIS
    True when an executeQueries error text looks like "INFO functions are not supported here" (HTTP 400 / engine 3239575574) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false, Position = 0)][AllowNull()][AllowEmptyString()][string]$Message)
    if ([string]::IsNullOrWhiteSpace($Message)) { return $false }
    return ($Message -match '(?i)\bINFO\b|not supported|unsupported|3239575574|Failed to execute the DAX query|HTTP 400|no response')
}

function Get-IQDaxShapeBool {
    <#
    .SYNOPSIS
    Normalises a bool-ish INFO value ("true"/1/$true) to $true/$false, $null when unknown (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false, Position = 0)][AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [bool]) { return $Value }
    $t = ([string]$Value).Trim()
    if ($t -ieq 'true' -or $t -eq '1') { return $true }
    if ($t -ieq 'false' -or $t -eq '0') { return $false }
    if ($t -eq '') { return $null }
    try { return [System.Convert]::ToBoolean($Value) } catch { return $null }
}

function ConvertTo-IQDaxModelShape {
    <#
    .SYNOPSIS
    Normalises INFO.VIEW.* rows (or raw INFO.* rows joined on IDs) into one name-keyed model shape (private).
    .DESCRIPTION
    Returns @{ Tables; Columns; Measures; Relationships; TableNameById; ColumnNameById }. Tables: @{ ID; Name; Description;
    IsHidden; StorageMode; IsCalculationGroup; Expression }. Columns: @{ ID; Table; Name; Type (Data|Calculated|
    CalculatedTableColumn|RowNumber); FormatString; DisplayFolder; Description; IsHidden; Expression }. Measures: @{ ID; Table;
    Name; Expression; FormatString; DisplayFolder; Description; IsHidden }. Relationships: @{ ID; Name; FromTable; FromColumn;
    ToTable; ToColumn; FromCardinality; ToCardinality; CrossFilteringBehavior; IsActive } with TOM enum names.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Data,
        [Parameter(Mandatory = $true)][bool]$ViewMode
    )
    $cardinalityMap = @{ 0 = 'None'; 1 = 'One'; 2 = 'Many' }
    $crossFilterMap = @{ 1 = 'OneDirection'; 2 = 'BothDirections'; 3 = 'Automatic' }
    $columnTypeMap = @{ 1 = 'Data'; 2 = 'Calculated'; 3 = 'RowNumber'; 4 = 'CalculatedTableColumn' }
    $modeMap = @{ 0 = 'Import'; 1 = 'DirectQuery'; 2 = 'Default'; 3 = 'Push'; 4 = 'Dual'; 5 = 'DirectLake' }
    $shape = @{ Tables = @(); Columns = @(); Measures = @(); Relationships = @(); TableNameById = @{}; ColumnNameById = @{} }
    $tables = New-Object System.Collections.Generic.List[object]
    $columns = New-Object System.Collections.Generic.List[object]
    $measures = New-Object System.Collections.Generic.List[object]
    $relationships = New-Object System.Collections.Generic.List[object]

    if ($ViewMode) {
        foreach ($t in @($Data['tables'])) {
            if ($null -eq $t) { continue }
            $id = [string](Get-IQDaxMember -Object $t -Name 'ID')
            $name = [string](Get-IQDaxMember -Object $t -Name 'Name')
            if ($id -ne '') { $shape.TableNameById[$id] = $name }
            $prec = Get-IQDaxMember -Object $t -Name 'CalculationGroupPrecedence'
            $tables.Add(@{
                    ID = $id; Name = $name; Description = (Get-IQDaxMember -Object $t -Name 'Description'); IsHidden = (Get-IQDaxShapeBool (Get-IQDaxMember -Object $t -Name 'IsHidden'))
                    StorageMode = (ConvertTo-IQDaxEnumName -Value (Get-IQDaxMember -Object $t -Name 'StorageMode') -Map $modeMap -Default '')
                    IsCalculationGroup = ($null -ne $prec -and ([string]$prec).Trim() -ne ''); Expression = (Get-IQDaxMember -Object $t -Name 'Expression')
                })
        }
        foreach ($c in @($Data['columns'])) {
            if ($null -eq $c) { continue }
            $id = [string](Get-IQDaxMember -Object $c -Name 'ID')
            $name = [string](Get-IQDaxMember -Object $c -Name 'Name')
            if ($id -ne '') { $shape.ColumnNameById[$id] = $name }
            $type = ConvertTo-IQDaxEnumName -Value (Get-IQDaxMember -Object $c -Name 'Type') -Map $columnTypeMap -Default 'Data'
            if ([string](Get-IQDaxMember -Object $c -Name 'DataCategory') -eq 'RowNumber') { $type = 'RowNumber' }
            $columns.Add(@{
                    ID = $id; Table = [string](Get-IQDaxMember -Object $c -Name 'Table'); Name = $name; Type = $type
                    FormatString = (Get-IQDaxMember -Object $c -Name 'FormatString'); DisplayFolder = (Get-IQDaxMember -Object $c -Name 'DisplayFolder')
                    Description = (Get-IQDaxMember -Object $c -Name 'Description'); IsHidden = (Get-IQDaxShapeBool (Get-IQDaxMember -Object $c -Name 'IsHidden'))
                    Expression = (Get-IQDaxMember -Object $c -Name 'Expression')
                })
        }
        foreach ($m in @($Data['measures'])) {
            if ($null -eq $m) { continue }
            $measures.Add(@{
                    ID = [string](Get-IQDaxMember -Object $m -Name 'ID'); Table = [string](Get-IQDaxMember -Object $m -Name 'Table'); Name = [string](Get-IQDaxMember -Object $m -Name 'Name')
                    Expression = (Get-IQDaxMember -Object $m -Name 'Expression'); FormatString = (Get-IQDaxMember -Object $m -Name 'FormatString')
                    DisplayFolder = (Get-IQDaxMember -Object $m -Name 'DisplayFolder'); Description = (Get-IQDaxMember -Object $m -Name 'Description')
                    IsHidden = (Get-IQDaxShapeBool (Get-IQDaxMember -Object $m -Name 'IsHidden'))
                })
        }
        foreach ($r in @($Data['relationships'])) {
            if ($null -eq $r) { continue }
            $relationships.Add(@{
                    ID = [string](Get-IQDaxMember -Object $r -Name 'ID'); Name = [string](Get-IQDaxMember -Object $r -Name 'Name')
                    FromTable = [string](Get-IQDaxMember -Object $r -Name 'FromTable'); FromColumn = [string](Get-IQDaxMember -Object $r -Name 'FromColumn')
                    ToTable = [string](Get-IQDaxMember -Object $r -Name 'ToTable'); ToColumn = [string](Get-IQDaxMember -Object $r -Name 'ToColumn')
                    FromCardinality = (ConvertTo-IQDaxEnumName -Value (Get-IQDaxMember -Object $r -Name 'FromCardinality') -Map $cardinalityMap)
                    ToCardinality = (ConvertTo-IQDaxEnumName -Value (Get-IQDaxMember -Object $r -Name 'ToCardinality') -Map $cardinalityMap)
                    CrossFilteringBehavior = (ConvertTo-IQDaxEnumName -Value (Get-IQDaxMember -Object $r -Name 'CrossFilteringBehavior') -Map $crossFilterMap)
                    IsActive = (Get-IQDaxShapeBool (Get-IQDaxMember -Object $r -Name 'IsActive'))
                })
        }
    }
    else {
        $tableIndex = Get-IQDaxIndex -Rows $Data['tables']
        $columnIndex = Get-IQDaxIndex -Rows $Data['columns']
        foreach ($t in @($Data['tables'])) {
            if ($null -eq $t) { continue }
            $id = [string](Get-IQDaxMember -Object $t -Name 'ID')
            $name = [string](Get-IQDaxMember -Object $t -Name 'Name')
            if ($id -ne '') { $shape.TableNameById[$id] = $name }
            $cgId = Get-IQDaxMember -Object $t -Name 'CalculationGroupID'
            $tables.Add(@{
                    ID = $id; Name = $name; Description = (Get-IQDaxMember -Object $t -Name 'Description'); IsHidden = (Get-IQDaxShapeBool (Get-IQDaxMember -Object $t -Name 'IsHidden'))
                    StorageMode = ''; IsCalculationGroup = ($null -ne $cgId -and ([string]$cgId).Trim() -ne ''); Expression = $null
                })
        }
        foreach ($c in @($Data['columns'])) {
            if ($null -eq $c) { continue }
            $id = [string](Get-IQDaxMember -Object $c -Name 'ID')
            $name = Get-IQDaxColumnName -Column $c
            if ($id -ne '') { $shape.ColumnNameById[$id] = $name }
            $columns.Add(@{
                    ID = $id; Table = (Get-IQDaxIndexedName -Index $tableIndex -Id (Get-IQDaxMember -Object $c -Name 'TableID')); Name = $name
                    Type = (ConvertTo-IQDaxEnumName -Value (Get-IQDaxMember -Object $c -Name 'Type') -Map $columnTypeMap -Default 'Data')
                    FormatString = (Get-IQDaxMember -Object $c -Name 'FormatString'); DisplayFolder = (Get-IQDaxMember -Object $c -Name 'DisplayFolder')
                    Description = (Get-IQDaxMember -Object $c -Name 'Description'); IsHidden = (Get-IQDaxShapeBool (Get-IQDaxMember -Object $c -Name 'IsHidden'))
                    Expression = (Get-IQDaxMember -Object $c -Name 'Expression')
                })
        }
        foreach ($m in @($Data['measures'])) {
            if ($null -eq $m) { continue }
            $measures.Add(@{
                    ID = [string](Get-IQDaxMember -Object $m -Name 'ID'); Table = (Get-IQDaxIndexedName -Index $tableIndex -Id (Get-IQDaxMember -Object $m -Name 'TableID')); Name = [string](Get-IQDaxMember -Object $m -Name 'Name')
                    Expression = (Get-IQDaxMember -Object $m -Name 'Expression'); FormatString = (Get-IQDaxMember -Object $m -Name 'FormatString')
                    DisplayFolder = (Get-IQDaxMember -Object $m -Name 'DisplayFolder'); Description = (Get-IQDaxMember -Object $m -Name 'Description')
                    IsHidden = (Get-IQDaxShapeBool (Get-IQDaxMember -Object $m -Name 'IsHidden'))
                })
        }
        foreach ($r in @($Data['relationships'])) {
            if ($null -eq $r) { continue }
            $fromColumn = ''
            $toColumn = ''
            $fcid = [string](Get-IQDaxMember -Object $r -Name 'FromColumnID')
            $tcid = [string](Get-IQDaxMember -Object $r -Name 'ToColumnID')
            if ($columnIndex.ContainsKey($fcid)) { $fromColumn = Get-IQDaxColumnName -Column $columnIndex[$fcid] }
            if ($columnIndex.ContainsKey($tcid)) { $toColumn = Get-IQDaxColumnName -Column $columnIndex[$tcid] }
            $relationships.Add(@{
                    ID = [string](Get-IQDaxMember -Object $r -Name 'ID'); Name = [string](Get-IQDaxMember -Object $r -Name 'Name')
                    FromTable = (Get-IQDaxIndexedName -Index $tableIndex -Id (Get-IQDaxMember -Object $r -Name 'FromTableID')); FromColumn = $fromColumn
                    ToTable = (Get-IQDaxIndexedName -Index $tableIndex -Id (Get-IQDaxMember -Object $r -Name 'ToTableID')); ToColumn = $toColumn
                    FromCardinality = (ConvertTo-IQDaxEnumName -Value (Get-IQDaxMember -Object $r -Name 'FromCardinality') -Map $cardinalityMap)
                    ToCardinality = (ConvertTo-IQDaxEnumName -Value (Get-IQDaxMember -Object $r -Name 'ToCardinality') -Map $cardinalityMap)
                    CrossFilteringBehavior = (ConvertTo-IQDaxEnumName -Value (Get-IQDaxMember -Object $r -Name 'CrossFilteringBehavior') -Map $crossFilterMap)
                    IsActive = (Get-IQDaxShapeBool (Get-IQDaxMember -Object $r -Name 'IsActive'))
                })
        }
    }
    $shape.Tables = $tables.ToArray()
    $shape.Columns = $columns.ToArray()
    $shape.Measures = $measures.ToArray()
    $shape.Relationships = $relationships.ToArray()
    return $shape
}

function ConvertTo-IQDaxCalcDependencyRows {
    <#
    .SYNOPSIS
    Maps INFO.CALCDEPENDENCY() rows to "_MD.csv" rows (brief section 13 mapping; direct dependencies only) (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()]$Rows,
        [Parameter(Mandatory = $true)][hashtable]$Common,
        [Parameter(Mandatory = $false)][AllowNull()][hashtable]$CalcGroupTableNames
    )
    if ($null -eq $CalcGroupTableNames) { $CalcGroupTableNames = @{} }
    $objectTypeMap = @{ 'MEASURE' = 'Measure'; 'CALC_COLUMN' = 'CalculatedColumn'; 'CALCULATION_ITEM' = 'CalculationItem'; 'CALC_ITEM' = 'CalculationItem' }
    $mdRows = New-Object System.Collections.Generic.List[object]
    $mdSeen = @{}
    foreach ($d in @($Rows)) {
        if ($null -eq $d) { continue }
        $objType = ([string](Get-IQDaxMember -Object $d -Name 'OBJECT_TYPE')).Trim().ToUpperInvariant()
        if (-not $objectTypeMap.ContainsKey($objType)) { continue }
        $refType = ([string](Get-IQDaxMember -Object $d -Name 'REFERENCED_OBJECT_TYPE')).Trim().ToUpperInvariant()
        $refTable = [string](Get-IQDaxMember -Object $d -Name 'REFERENCED_TABLE')
        $refObject = [string](Get-IQDaxMember -Object $d -Name 'REFERENCED_OBJECT')
        $dependsOn = $null
        $dependsOnType = $null
        switch ($refType) {
            'MEASURE' { $dependsOn = ConvertTo-IQDaxBracketRef -Name $refObject; $dependsOnType = 'Measure' }
            'COLUMN' { $dependsOn = (ConvertTo-IQDaxTableRef -Table $refTable) + (ConvertTo-IQDaxBracketRef -Name $refObject); $dependsOnType = 'Column' }
            'CALC_COLUMN' { $dependsOn = (ConvertTo-IQDaxTableRef -Table $refTable) + (ConvertTo-IQDaxBracketRef -Name $refObject); $dependsOnType = 'Column' }
            'TABLE' {
                $dependsOn = ConvertTo-IQDaxTableRef -Table $refTable
                if ($CalcGroupTableNames.ContainsKey($refTable)) { $dependsOnType = 'CalculationGroupTable' } else { $dependsOnType = 'Table' }
            }
            'CALC_TABLE' {
                $dependsOn = ConvertTo-IQDaxTableRef -Table $refTable
                if ($CalcGroupTableNames.ContainsKey($refTable)) { $dependsOnType = 'CalculationGroupTable' } else { $dependsOnType = 'Table' }
            }
            'CALCULATION_ITEM' { $dependsOn = (ConvertTo-IQDaxTableRef -Table $refTable) + (ConvertTo-IQDaxBracketRef -Name $refObject); $dependsOnType = 'CalculationItem' }
            'CALC_ITEM' { $dependsOn = (ConvertTo-IQDaxTableRef -Table $refTable) + (ConvertTo-IQDaxBracketRef -Name $refObject); $dependsOnType = 'CalculationItem' }
            default { $dependsOn = $null }
        }
        if ($null -eq $dependsOn) { continue }
        $objectName = [string](Get-IQDaxMember -Object $d -Name 'OBJECT')
        $objectTable = [string](Get-IQDaxMember -Object $d -Name 'TABLE')
        $dedupeKey = ($objType + '|' + $objectTable + '|' + $objectName + '|' + $dependsOn + '|' + $dependsOnType)
        if ($mdSeen.ContainsKey($dedupeKey)) { continue }
        $mdSeen[$dedupeKey] = $true
        $mdRows.Add([ordered]@{
                ObjectName = $objectName; ObjectType = $objectTypeMap[$objType]; DependsOn = $dependsOn; DependsOnType = $dependsOnType
                ModelAsOfDate = $Common.ModelAsOfDate; ModelName = $Common.ModelName; ModelID = $Common.ModelID
            })
    }
    return $mdRows.ToArray()
}

function Get-IQModelDetailViaDax {
    <#
    .SYNOPSIS
    Produces "<CleanWs> ~ <CleanModel>.csv" and "..._MD.csv" for a dataset from INFO.VIEW.* / INFO.* DAX queries (executeQueries).
    .DESCRIPTION
    Reproduces the rows of "Model Detail Extract Script.csx" and "Measure Dependency Extract Script.csx" (brief section 13).
    Query plan (one executeQueries call per query, each guarded, raw results cached under extracts\dax\<datasetId>\*.json
    and reused on re-runs):
      1. INFO.VIEW.TABLES() (falls back to raw INFO.TABLES() on older engines); a failure of both = not available (Failed).
      2. INFO.VIEW.COLUMNS / MEASURES / RELATIONSHIPS (or the raw INFO.* equivalents joined on IDs).
      3. Best effort raw INFO.PARTITIONS / INFO.MODEL / INFO.ROLES / INFO.TABLEPERMISSIONS / INFO.CALCULATIONGROUPS /
         INFO.CALCULATIONITEMS / INFO.HIERARCHIES / INFO.LEVELS / INFO.CALCDEPENDENCY. The JSON executeQueries endpoint
         officially does not support raw INFO functions: the first HTTP 400 / "not supported" answer skips the remaining
         raw queries with ONE Warn, the rows of those kinds are simply missing and the part names are returned in
         Unavailable (and appended to Message, hence to the checkpoint).
    Dependency rows: INFO.CALCDEPENDENCY() when it returned rows, else Get-IQDaxReferences over the measure / calculated
    column / calculation item expressions that were returned. INFO.VIEW.MEASURES()[Expression] is blank unless the caller
    has write permission on the model (Contributor+); then no dependency rows can be produced and Message says so.
    Never throws; returns @{ Success; Csv; MdCsv; Outputs; Message; InfoUnsupported; RowCount; DependencyRowCount;
    Method='Dax'; BaseName; Unavailable; DependencySource; ExpressionsMasked }.
    ModelName = "<CleanWs> ~ <CleanModel>"; ModelID = DatasetId for dedicated-capacity workspaces, else the same
    "<CleanWs> ~ <CleanModel>" string (what the PBIT joins on for Pro models); ModelAsOfDate = RunId when it is a date.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Dataset,
        [Parameter(Mandatory = $true)][string]$OutputFolder,
        [Parameter(Mandatory = $false)][AllowNull()][object]$IsDedicated,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$ModelAsOfDate,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$BaseName,
        [Parameter(Mandatory = $false)][switch]$Refresh,
        [Parameter(Mandatory = $false)][string]$Stage = 'ModelDetail'
    )
    $datasetId = [string](Get-IQDaxMember -Object $Dataset -Name 'DatasetId')
    $workspaceId = [string](Get-IQDaxMember -Object $Dataset -Name 'WorkspaceId')
    if ([string]::IsNullOrWhiteSpace($BaseName)) { $BaseName = Get-IQModelBackupFileName -Dataset $Dataset }
    $item = $BaseName
    $result = @{ Success = $false; Csv = $null; MdCsv = $null; Outputs = @(); Message = ''; InfoUnsupported = $false; RowCount = 0; DependencyRowCount = 0; Method = 'Dax'; BaseName = $BaseName; Unavailable = @(); DependencySource = 'none'; ExpressionsMasked = $false }
    if ([string]::IsNullOrWhiteSpace($datasetId)) { $result.Message = 'Dataset has no DatasetId'; return $result }

    # Dedicated capacity decides the ModelID convention (brief section 13 / audit x1 section 4.1).
    $dedicated = $false
    if ($PSBoundParameters.ContainsKey('IsDedicated') -and $null -ne $IsDedicated) { $dedicated = [bool]$IsDedicated }
    else {
        $flag = Get-IQDaxMember -Object $Dataset -Name 'WorkspaceIsOnDedicatedCapacity'
        if ($null -ne $flag) { try { $dedicated = [System.Convert]::ToBoolean($flag) } catch { $dedicated = ([string]$flag -ieq 'true') } }
    }
    if ([string]::IsNullOrWhiteSpace($ModelAsOfDate)) { $ModelAsOfDate = Get-IQDaxModelAsOfDate }
    $modelId = $BaseName
    if ($dedicated) { $modelId = $datasetId }
    $common = @{ ModelAsOfDate = $ModelAsOfDate; ModelName = $BaseName; ModelID = $modelId }

    if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
    $csvPath = Join-Path $OutputFolder ($BaseName + '.csv')
    $mdPath = Join-Path $OutputFolder ($BaseName + '_MD.csv')
    $extractFolder = Get-IQDaxExtractFolder -Key $datasetId
    $queryParams = @{ ExtractFolder = $extractFolder; WorkspaceId = $workspaceId; DatasetId = $datasetId; Refresh = $Refresh; Stage = $Stage; Item = $item }
    $data = @{}
    $warnings = @()
    $unavailable = New-Object System.Collections.Generic.List[string]

    # 1. Tables: INFO.VIEW.TABLES() first (works with Build permission on Pro and capacity), raw INFO.TABLES() as the fallback.
    $viewMode = $true
    try {
        $data['tables'] = @(Get-IQDaxInfoRowSet @queryParams -Name 'view-tables' -Dax 'EVALUATE INFO.VIEW.TABLES()')
    }
    catch {
        $viewError = $_.Exception.Message
        Write-IQLog -Level Debug -Stage $Stage -Item $item -Message ("INFO.VIEW.TABLES() failed ({0}); trying INFO.TABLES()" -f $viewError)
        $viewMode = $false
        try {
            $data['tables'] = @(Get-IQDaxInfoRowSet @queryParams -Name 'tables' -Dax 'EVALUATE INFO.TABLES()')
        }
        catch {
            $rawError = $_.Exception.Message
            $result.Message = 'INFO.VIEW.TABLES() failed: ' + $viewError + ' | INFO.TABLES() failed: ' + $rawError
            if ((Test-IQDaxInfoUnsupportedMessage -Message $viewError) -or (Test-IQDaxInfoUnsupportedMessage -Message $rawError)) { $result.InfoUnsupported = $true }
            Write-IQLog -Level Warn -Stage $Stage -Item $item -Message ("DAX model detail unavailable: {0}" -f $result.Message)
            return $result
        }
    }

    # 2. Columns, measures, relationships (each guarded: a failure yields no rows of that kind + a Warn).
    $coreQueries = [ordered]@{}
    if ($viewMode) {
        $coreQueries['columns'] = @{ Name = 'view-columns'; Dax = 'EVALUATE INFO.VIEW.COLUMNS()' }
        $coreQueries['measures'] = @{ Name = 'view-measures'; Dax = 'EVALUATE INFO.VIEW.MEASURES()' }
        $coreQueries['relationships'] = @{ Name = 'view-relationships'; Dax = 'EVALUATE INFO.VIEW.RELATIONSHIPS()' }
    }
    else {
        $coreQueries['columns'] = @{ Name = 'columns'; Dax = 'EVALUATE INFO.COLUMNS()' }
        $coreQueries['measures'] = @{ Name = 'measures'; Dax = 'EVALUATE INFO.MEASURES()' }
        $coreQueries['relationships'] = @{ Name = 'relationships'; Dax = 'EVALUATE INFO.RELATIONSHIPS()' }
    }
    foreach ($key in $coreQueries.Keys) {
        $q = $coreQueries[$key]
        try { $data[$key] = @(Get-IQDaxInfoRowSet @queryParams -Name $q.Name -Dax $q.Dax) }
        catch {
            $data[$key] = @()
            $warnings += ('{0}: {1}' -f $key, $_.Exception.Message)
            Write-IQLog -Level Warn -Stage $Stage -Item $item -Message ("{0} failed; rows of that kind will be missing: {1}" -f $q.Dax, $_.Exception.Message)
        }
    }

    # 3. Best-effort raw INFO.* (officially unsupported on the JSON endpoint): stop at the first "not supported" answer.
    $rawQueries = [ordered]@{
        partitions       = 'EVALUATE INFO.PARTITIONS()'
        model            = 'EVALUATE INFO.MODEL()'
        roles            = 'EVALUATE INFO.ROLES()'
        tablepermissions = 'EVALUATE INFO.TABLEPERMISSIONS()'
        calcgroups       = 'EVALUATE INFO.CALCULATIONGROUPS()'
        calcitems        = 'EVALUATE INFO.CALCULATIONITEMS()'
        hierarchies      = 'EVALUATE INFO.HIERARCHIES()'
        levels           = 'EVALUATE INFO.LEVELS()'
        calcdependency   = 'EVALUATE INFO.CALCDEPENDENCY()'
    }
    $rawBlocked = $false
    $blockReason = ''
    foreach ($name in $rawQueries.Keys) {
        if ($rawBlocked) { $data[$name] = @(); if ($name -ne 'model') { $unavailable.Add($name) }; continue }
        try { $data[$name] = @(Get-IQDaxInfoRowSet @queryParams -Name $name -Dax $rawQueries[$name]) }
        catch {
            $data[$name] = @()
            if ($name -ne 'model') { $unavailable.Add($name) }
            if (Test-IQDaxInfoUnsupportedMessage -Message $_.Exception.Message) {
                $rawBlocked = $true
                $blockReason = $_.Exception.Message
            }
            else {
                $warnings += ('{0}: {1}' -f $name, $_.Exception.Message)
                Write-IQLog -Level Warn -Stage $Stage -Item $item -Message ("{0} failed; rows of that kind will be missing: {1}" -f $rawQueries[$name], $_.Exception.Message)
            }
        }
    }
    if ($rawBlocked) {
        Write-IQLog -Level Warn -Stage $Stage -Item $item -Message ("Raw INFO.* functions are not available through executeQueries for this model; skipped: {0}. Partitions, roles/RLS filters, calculation items, hierarchies and INFO.CALCDEPENDENCY rows will be missing (engine: {1})" -f ($unavailable -join ', '), $blockReason)
    }
    $result.Unavailable = @($unavailable.ToArray())

    # Shapes and indexes.
    $shape = ConvertTo-IQDaxModelShape -Data $data -ViewMode $viewMode
    $modeMap = @{ 0 = 'Import'; 1 = 'DirectQuery'; 2 = 'Default'; 3 = 'Push'; 4 = 'Dual'; 5 = 'DirectLake' }
    $defaultMode = 'Default'
    $modelRow = @($data['model'])
    if ($modelRow.Count -gt 0 -and $null -ne $modelRow[0]) {
        $dm = ConvertTo-IQDaxEnumName -Value (Get-IQDaxMember -Object $modelRow[0] -Name 'DefaultMode') -Map $modeMap -Default 'Default'
        if ($dm -and $dm -ne 'Default') { $defaultMode = $dm }
    }
    $resolveMode = {
        param($value)
        $m = ConvertTo-IQDaxEnumName -Value $value -Map $modeMap -Default ''
        if ($m -eq 'Default') { return $defaultMode }
        return $m
    }
    $tableNameById = $shape.TableNameById
    $tableIndex = Get-IQDaxIndex -Rows $data['tables']
    $hierarchyIndex = Get-IQDaxIndex -Rows $data['hierarchies']
    $roleIndex = Get-IQDaxIndex -Rows $data['roles']
    $calcGroupById = @{}
    foreach ($cg in @($data['calcgroups'])) {
        $tid = [string](Get-IQDaxMember -Object $cg -Name 'TableID')
        if ($tableNameById.ContainsKey($tid)) { $calcGroupById[[string](Get-IQDaxMember -Object $cg -Name 'ID')] = $tableNameById[$tid] }
    }
    $calcGroupTableNames = @{}
    foreach ($t in $shape.Tables) { if ($t.IsCalculationGroup) { $calcGroupTableNames[$t.Name] = $true } }
    foreach ($n in $calcGroupById.Values) { $calcGroupTableNames[[string]$n] = $true }
    # First (lowest ID) partition per table gives the table storage mode when INFO.VIEW.TABLES() did not (TE takes Partitions[0]).
    $firstPartitionMode = @{}
    foreach ($p in (@($data['partitions']) | Sort-Object { [double](Get-IQDaxMember -Object $_ -Name 'ID') })) {
        $tid = [string](Get-IQDaxMember -Object $p -Name 'TableID')
        if ($tid -eq '' -or -not $tableNameById.ContainsKey($tid)) { continue }
        $tn = $tableNameById[$tid]
        if (-not $firstPartitionMode.ContainsKey($tn)) { $firstPartitionMode[$tn] = (& $resolveMode (Get-IQDaxMember -Object $p -Name 'Mode')) }
    }

    $rows = New-Object System.Collections.Generic.List[object]

    # Tables (all, including calculation-group tables).
    foreach ($t in $shape.Tables) {
        $mode = [string]$t.StorageMode
        if ($mode -eq '' -and $firstPartitionMode.ContainsKey($t.Name)) { $mode = $firstPartitionMode[$t.Name] }
        $rows.Add((New-IQModelDetailRow -Type 'Table' -Common $common -Fields @{ Table = $t.Name; Name = $t.Name; IsHidden = $t.IsHidden; TableStorageMode = $mode; Description = $t.Description }))
    }

    # Calculation groups (Table = Name = the calc-group table name) and their items.
    $groupTables = New-Object System.Collections.Generic.List[string]
    foreach ($t in $shape.Tables) { if ($t.IsCalculationGroup) { $groupTables.Add($t.Name) } }
    foreach ($n in $calcGroupById.Values) { if (-not $groupTables.Contains([string]$n)) { $groupTables.Add([string]$n) } }
    $calcItemObjects = New-Object System.Collections.Generic.List[object]
    foreach ($groupName in $groupTables) {
        $tableRow = @($shape.Tables | Where-Object { $_.Name -eq $groupName })
        $desc = $null
        $hidden = $null
        if ($tableRow.Count -gt 0) { $desc = $tableRow[0].Description; $hidden = $tableRow[0].IsHidden }
        $cgRow = @($data['calcgroups'] | Where-Object { $null -ne $_ -and $calcGroupById.ContainsKey([string](Get-IQDaxMember -Object $_ -Name 'ID')) -and $calcGroupById[[string](Get-IQDaxMember -Object $_ -Name 'ID')] -eq $groupName })
        if (($null -eq $desc -or [string]$desc -eq '') -and $cgRow.Count -gt 0) { $desc = Get-IQDaxMember -Object $cgRow[0] -Name 'Description' }
        $rows.Add((New-IQModelDetailRow -Type 'CalculationGroup' -Common $common -Fields @{ Table = $groupName; Name = $groupName; Description = $desc; IsHidden = $hidden }))
        if ($cgRow.Count -gt 0) {
            $cgId = [string](Get-IQDaxMember -Object $cgRow[0] -Name 'ID')
            foreach ($ci in (@($data['calcitems']) | Sort-Object { [double](Get-IQDaxMember -Object $_ -Name 'Ordinal') })) {
                if ([string](Get-IQDaxMember -Object $ci -Name 'CalculationGroupID') -ne $cgId) { continue }
                $ciName = Get-IQDaxMember -Object $ci -Name 'Name'
                $ciExpr = Get-IQDaxMember -Object $ci -Name 'Expression'
                $rows.Add((New-IQModelDetailRow -Type 'CalculationItem' -Common $common -Fields @{ Table = $groupName; Name = $ciName; Description = (Get-IQDaxMember -Object $ci -Name 'Description'); Expression = $ciExpr }))
                $calcItemObjects.Add(@{ ObjectName = [string]$ciName; ObjectType = 'CalculationItem'; Table = $groupName; Expression = $ciExpr })
            }
        }
    }

    # Columns: every non-RowNumber column as "Column"; calculated columns again as "CalculatedColumn" (mirrors the csx).
    $knownColumns = @{}
    foreach ($c in $shape.Columns) {
        if ($c.Type -eq 'RowNumber') { continue }
        if (-not $knownColumns.ContainsKey($c.Table)) { $knownColumns[$c.Table] = @() }
        $knownColumns[$c.Table] += $c.Name
        $rows.Add((New-IQModelDetailRow -Type 'Column' -Common $common -Fields @{ Table = $c.Table; Name = $c.Name; FormatString = $c.FormatString; DisplayFolder = $c.DisplayFolder; Description = $c.Description; IsHidden = $c.IsHidden }))
    }
    $calcColumnObjects = New-Object System.Collections.Generic.List[object]
    foreach ($c in $shape.Columns) {
        if ($c.Type -ne 'Calculated') { continue }
        $rows.Add((New-IQModelDetailRow -Type 'CalculatedColumn' -Common $common -Fields @{ Table = $c.Table; Name = $c.Name; FormatString = $c.FormatString; DisplayFolder = $c.DisplayFolder; Description = $c.Description; IsHidden = $c.IsHidden; Expression = $c.Expression }))
        $calcColumnObjects.Add(@{ ObjectName = $c.Name; ObjectType = 'CalculatedColumn'; Table = $c.Table; Expression = $c.Expression })
    }

    # Measures.
    $measureObjects = New-Object System.Collections.Generic.List[object]
    $measureNames = @()
    $blankExpressions = 0
    foreach ($m in $shape.Measures) {
        $rows.Add((New-IQModelDetailRow -Type 'Measure' -Common $common -Fields @{ Table = $m.Table; Name = $m.Name; FormatString = $m.FormatString; DisplayFolder = $m.DisplayFolder; Description = $m.Description; IsHidden = $m.IsHidden; Expression = $m.Expression }))
        $measureObjects.Add(@{ ObjectName = $m.Name; ObjectType = 'Measure'; Table = $m.Table; Expression = $m.Expression })
        $measureNames += $m.Name
        if ([string]::IsNullOrWhiteSpace([string]$m.Expression)) { $blankExpressions++ }
    }
    if ($shape.Measures.Count -gt 0 -and $blankExpressions -eq $shape.Measures.Count) {
        $result.ExpressionsMasked = $true
        Write-IQLog -Level Warn -Stage $Stage -Item $item -Message ("All {0} measure expressions came back blank: INFO.VIEW.MEASURES()[Expression] is only populated for callers with write permission on the semantic model (workspace Contributor or above)" -f $shape.Measures.Count)
    }

    # Hierarchies and levels (raw INFO.HIERARCHIES / INFO.LEVELS, when available).
    foreach ($h in @($data['hierarchies'])) {
        $tn = ''
        $tid = [string](Get-IQDaxMember -Object $h -Name 'TableID')
        if ($tableNameById.ContainsKey($tid)) { $tn = $tableNameById[$tid] }
        $rows.Add((New-IQModelDetailRow -Type 'Hierarchy' -Common $common -Fields @{ Table = $tn; Name = (Get-IQDaxMember -Object $h -Name 'Name'); DisplayFolder = (Get-IQDaxMember -Object $h -Name 'DisplayFolder'); Description = (Get-IQDaxMember -Object $h -Name 'Description'); IsHidden = (Get-IQDaxMember -Object $h -Name 'IsHidden') }))
    }
    foreach ($l in @($data['levels'])) {
        $hid = [string](Get-IQDaxMember -Object $l -Name 'HierarchyID')
        $tableName = ''
        if ($hierarchyIndex.ContainsKey($hid)) {
            $tid = [string](Get-IQDaxMember -Object $hierarchyIndex[$hid] -Name 'TableID')
            if ($tableNameById.ContainsKey($tid)) { $tableName = $tableNameById[$tid] }
        }
        $rows.Add((New-IQModelDetailRow -Type 'Level' -Common $common -Fields @{ Table = $tableName; Name = (Get-IQDaxMember -Object $l -Name 'Name'); Description = (Get-IQDaxMember -Object $l -Name 'Description') }))
    }

    # Partitions (Expression = M / DAX / query text from QueryDefinition).
    foreach ($p in @($data['partitions'])) {
        $expr = Get-IQDaxMember -Object $p -Name 'QueryDefinition'
        if ($null -eq $expr -or [string]$expr -eq '') { $expr = Get-IQDaxMember -Object $p -Name 'Expression' }
        $tn = ''
        $tid = [string](Get-IQDaxMember -Object $p -Name 'TableID')
        if ($tableNameById.ContainsKey($tid)) { $tn = $tableNameById[$tid] }
        $rows.Add((New-IQModelDetailRow -Type 'Partition' -Common $common -Fields @{ Table = $tn; Name = (Get-IQDaxMember -Object $p -Name 'Name'); Description = (Get-IQDaxMember -Object $p -Name 'Description'); TableStorageMode = (& $resolveMode (Get-IQDaxMember -Object $p -Name 'Mode')); Expression = $expr }))
    }

    # RLS filters: one row per table permission, Name = role name.
    foreach ($tp in @($data['tablepermissions'])) {
        $roleName = Get-IQDaxIndexedName -Index $roleIndex -Id (Get-IQDaxMember -Object $tp -Name 'RoleID')
        $tn = ''
        $tid = [string](Get-IQDaxMember -Object $tp -Name 'TableID')
        if ($tableNameById.ContainsKey($tid)) { $tn = $tableNameById[$tid] }
        $rows.Add((New-IQModelDetailRow -Type 'RLSFilter' -Common $common -Fields @{ Table = $tn; Name = $roleName; Expression = (Get-IQDaxMember -Object $tp -Name 'FilterExpression') }))
    }

    # Relationships.
    foreach ($r in $shape.Relationships) {
        $status = ''
        if ($null -ne $r.IsActive) { if ($r.IsActive) { $status = 'True' } else { $status = 'False' } }
        $rows.Add((New-IQModelDetailRow -Type 'Relationship' -Common $common -Fields @{
                    Table = $r.FromTable; Name = $r.FromColumn; Expression = $r.Name
                    RelationshipFromTable = $r.FromTable; RelationshipFromColumn = $r.FromColumn; RelationshipToTable = $r.ToTable; RelationshipToColumn = $r.ToColumn
                    RelationshipStatus = $status; RelationshipFromCardinality = $r.FromCardinality; RelationshipToCardinality = $r.ToCardinality
                    RelationshipCrossFilteringBehavior = $r.CrossFilteringBehavior
                }))
    }

    # Measure dependencies: INFO.CALCDEPENDENCY() when it returned rows, else the regex reference extractor.
    $mdRows = @()
    if (@($data['calcdependency']).Count -gt 0) {
        $mdRows = @(ConvertTo-IQDaxCalcDependencyRows -Rows $data['calcdependency'] -Common $common -CalcGroupTableNames $calcGroupTableNames)
        $result.DependencySource = 'INFO.CALCDEPENDENCY'
    }
    else {
        $objects = @($measureObjects.ToArray()) + @($calcColumnObjects.ToArray()) + @($calcItemObjects.ToArray())
        $tableNames = @($shape.Tables | ForEach-Object { $_.Name })
        $mdRows = @(ConvertTo-IQMeasureDependencyRows -Objects $objects -Common $common -KnownTables $tableNames -KnownMeasures $measureNames -KnownColumns $knownColumns -CalculationGroupTables @($calcGroupTableNames.Keys))
        if ($objects.Count -gt 0 -and -not $result.ExpressionsMasked) { $result.DependencySource = 'expression parsing (approximate)' }
    }

    try {
        $result.RowCount = Write-IQCsvFile -Path $csvPath -Header (Get-IQModelDetailHeader) -Rows $rows.ToArray()
        $result.DependencyRowCount = Write-IQCsvFile -Path $mdPath -Header (Get-IQMeasureDependencyHeader) -Rows $mdRows
    }
    catch {
        $result.Message = 'Could not write CSV files: ' + $_.Exception.Message
        Write-IQLog -Level Error -Stage $Stage -Item $item -Message $result.Message -Exception $_.Exception
        return $result
    }
    $result.Success = $true
    $result.Csv = $csvPath
    $result.MdCsv = $mdPath
    $result.Outputs = @($csvPath, $mdPath)
    $source = 'INFO.VIEW.*'
    if (-not $viewMode) { $source = 'INFO.*' }
    $result.Message = ('{0} object rows, {1} dependency rows via DAX {2} (dependencies: {3})' -f $result.RowCount, $result.DependencyRowCount, $source, $result.DependencySource)
    if ($result.ExpressionsMasked) { $result.Message += '; measure expressions masked (write permission on the model required)' }
    if ($unavailable.Count -gt 0) { $result.Message += ' (unavailable via REST: ' + ($unavailable -join ', ') + ')' }
    if ($warnings.Count -gt 0) { $result.Message += ' (partial: ' + ($warnings -join '; ') + ')' }
    Write-IQLog -Level Success -Stage $Stage -Item $item -Message $result.Message
    return $result
}

function Get-IQDaxModelAsOfDate {
    <#
    .SYNOPSIS
    ModelAsOfDate for CSV rows: the RunId when it is a yyyy-MM-dd folder name, else the run start date, else today (private).
    #>
    [CmdletBinding()]
    param()
    $runId = $null
    if ($script:IQ -and $script:IQ.ContainsKey('RunId')) { $runId = [string]$script:IQ.RunId }
    if ($runId -and $runId -match '^\d{4}-\d{2}-\d{2}$') { return $runId }
    try {
        if ($script:IQ -and $script:IQ.Manifest) {
            $started = Get-IQDaxMember -Object $script:IQ.Manifest -Name 'startedUtc'
            if ($started) {
                $dt = [datetime]::Parse([string]$started, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::RoundtripKind)
                return $dt.ToLocalTime().ToString('yyyy-MM-dd')
            }
        }
    }
    catch { $null = $null }
    return (Get-Date -Format 'yyyy-MM-dd')
}

function Get-IQUsageMetrics {
    <#
    .SYNOPSIS
    Reads the usage-metrics semantic model of a workspace via DAX (Report views / Report page views for the last N days).
    .DESCRIPTION
    Finds a dataset named "Report Usage Metrics Model" (new usage metrics) or "Usage Metrics Report" (classic) in the
    workspace, discovers its tables with INFO.TABLES()/INFO.COLUMNS(), then reads 'Report views', 'Report page views',
    'Reports' and 'Users' (each guarded) filtered on the table's date column over the last -Days days. Report and
    user names are joined in when the lookup tables expose ReportId/UserId keys. Never throws; returns
    @{ Found; DatasetId; DatasetName; ReportViews; ReportPageViews; Reports; Users; Message } with WorkspaceId /
    WorkspaceName stamped on every row.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$WorkspaceId,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$WorkspaceName,
        [Parameter(Mandatory = $false)][int]$Days = 30,
        [Parameter(Mandatory = $false)][string]$Stage = 'Extras'
    )
    $result = @{ Found = $false; DatasetId = $null; DatasetName = $null; ReportViews = @(); ReportPageViews = @(); Reports = @(); Users = @(); Message = '' }
    $item = $WorkspaceName
    if ([string]::IsNullOrWhiteSpace($item)) { $item = $WorkspaceId }
    if ($Days -lt 1) { $Days = 1 }
    try {
        $listPath = 'datasets'
        if (Test-IQDaxGuid -Value $WorkspaceId) { $listPath = 'groups/' + $WorkspaceId + '/datasets' }
        $response = Invoke-IQApi -Method GET -Path $listPath -AllowNotFound -Stage $Stage
        $datasets = @()
        if ($null -ne $response) { $datasets = @(Get-IQDaxMember -Object $response -Name 'value') }
        $usage = @($datasets | Where-Object { $null -ne $_ -and ([string]$_.name -eq 'Report Usage Metrics Model' -or [string]$_.name -eq 'Usage Metrics Report') })
        if ($usage.Count -eq 0) {
            $result.Message = 'No usage metrics semantic model in this workspace (open the workspace usage metrics report once to create it).'
            Write-IQLog -Level Debug -Stage $Stage -Item $item -Message $result.Message
            return $result
        }
        $chosen = @($usage | Where-Object { [string]$_.name -eq 'Report Usage Metrics Model' })
        if ($chosen.Count -eq 0) { $chosen = $usage }
        $usageDataset = $chosen[0]
        $result.Found = $true
        $result.DatasetId = [string]$usageDataset.id
        $result.DatasetName = [string]$usageDataset.name

        $tables = @(Invoke-IQDaxQuery -WorkspaceId $WorkspaceId -DatasetId $result.DatasetId -Dax 'EVALUATE INFO.TABLES()' -Stage $Stage -Item $item)
        $columns = @()
        try { $columns = @(Invoke-IQDaxQuery -WorkspaceId $WorkspaceId -DatasetId $result.DatasetId -Dax 'EVALUATE INFO.COLUMNS()' -Stage $Stage -Item $item) }
        catch { Write-IQLog -Level Debug -Stage $Stage -Item $item -Message ("INFO.COLUMNS() failed on the usage model: " + $_.Exception.Message) }
        $tableNames = @{}
        foreach ($t in $tables) { $n = [string](Get-IQDaxMember -Object $t -Name 'Name'); if ($n) { $tableNames[$n] = [string](Get-IQDaxMember -Object $t -Name 'ID') } }

        $findDateColumn = {
            param($tableName)
            $tid = $tableNames[$tableName]
            $candidates = @($columns | Where-Object { [string](Get-IQDaxMember -Object $_ -Name 'TableID') -eq $tid })
            foreach ($c in $candidates) { if ((Get-IQDaxColumnName -Column $c) -ieq 'Date') { return 'Date' } }
            foreach ($c in $candidates) {
                $dt = Get-IQDaxMember -Object $c -Name 'ExplicitDataType'
                if ($null -eq $dt) { $dt = Get-IQDaxMember -Object $c -Name 'InferredDataType' }
                if ([string]$dt -eq '9') { return (Get-IQDaxColumnName -Column $c) }
            }
            return $null
        }
        $readTable = {
            param($tableName, [bool]$filterByDate)
            if (-not $tableNames.ContainsKey($tableName)) { return @() }
            $tref = ConvertTo-IQDaxTableRef -Table $tableName
            $dax = ('EVALUATE TOPN(100000, {0})' -f $tref)
            if ($filterByDate) {
                $dateCol = & $findDateColumn $tableName
                if ($dateCol) { $dax = ('EVALUATE TOPN(100000, FILTER({0}, {0}{1} >= TODAY() - {2}))' -f $tref, (ConvertTo-IQDaxBracketRef -Name $dateCol), $Days) }
            }
            try { return @(Invoke-IQDaxQuery -WorkspaceId $WorkspaceId -DatasetId $result.DatasetId -Dax $dax -Stage $Stage -Item $item) }
            catch {
                Write-IQLog -Level Warn -Stage $Stage -Item $item -Message ("Usage table '{0}' could not be read: {1}" -f $tableName, $_.Exception.Message)
                return @()
            }
        }
        $reports = @(& $readTable 'Reports' $false)
        $users = @(& $readTable 'Users' $false)
        $reportViews = @(& $readTable 'Report views' $true)
        $pageViews = @(& $readTable 'Report page views' $true)

        $reportNameById = @{}
        foreach ($r in $reports) {
            $id = Get-IQDaxMember -Object $r -Name 'ReportId'
            if ($null -eq $id) { $id = Get-IQDaxMember -Object $r -Name 'ReportGuid' }
            $name = Get-IQDaxMember -Object $r -Name 'ReportName'
            if ($null -eq $name) { $name = Get-IQDaxMember -Object $r -Name 'Name' }
            if ($null -ne $id -and $null -ne $name) { $reportNameById[[string]$id] = [string]$name }
        }
        $userNameById = @{}
        foreach ($u in $users) {
            $id = Get-IQDaxMember -Object $u -Name 'UserId'
            $name = $null
            foreach ($p in @('UserPrincipalName', 'UserName', 'User', 'DisplayName', 'Name')) { $v = Get-IQDaxMember -Object $u -Name $p; if ($null -ne $v -and [string]$v -ne '') { $name = [string]$v; break } }
            if ($null -ne $id -and $null -ne $name) { $userNameById[[string]$id] = $name }
        }
        $stamp = {
            param($rowsIn)
            $out = @()
            foreach ($r in @($rowsIn)) {
                if ($null -eq $r) { continue }
                $o = [ordered]@{ WorkspaceId = $WorkspaceId; WorkspaceName = $WorkspaceName; UsageDatasetId = $result.DatasetId }
                foreach ($p in $r.PSObject.Properties) { $o[$p.Name] = $p.Value }
                $rid = Get-IQDaxMember -Object $r -Name 'ReportId'
                if ($null -ne $rid -and -not $o.Contains('ReportName') -and $reportNameById.ContainsKey([string]$rid)) { $o['ReportName'] = $reportNameById[[string]$rid] }
                $uid = Get-IQDaxMember -Object $r -Name 'UserId'
                if ($null -ne $uid -and -not $o.Contains('UserPrincipalName') -and $userNameById.ContainsKey([string]$uid)) { $o['UserPrincipalName'] = $userNameById[[string]$uid] }
                $out += [PSCustomObject]$o
            }
            return $out
        }
        $result.ReportViews = @(& $stamp $reportViews)
        $result.ReportPageViews = @(& $stamp $pageViews)
        $result.Reports = @(& $stamp $reports)
        $result.Users = @(& $stamp $users)
        $result.Message = ('{0} report views, {1} page views (last {2} days) from "{3}"' -f $result.ReportViews.Count, $result.ReportPageViews.Count, $Days, $result.DatasetName)
        Write-IQLog -Level Info -Stage $Stage -Item $item -Message $result.Message
    }
    catch {
        $result.Message = 'Usage metrics failed: ' + $_.Exception.Message
        Write-IQLog -Level Warn -Stage $Stage -Item $item -Message $result.Message
    }
    return $result
}
