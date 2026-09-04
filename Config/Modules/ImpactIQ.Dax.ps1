# ImpactIQ.Dax.ps1 - executeQueries client, INFO.* model-detail fallback and usage-metrics queries.
#
# Contract: brief sections 2.7, 7.3, 8.4 and 13.
#   Invoke-IQDaxQuery -WorkspaceId -DatasetId -Dax [-Impersonate]  -> array of flattened row objects (throws on error)
#   Get-IQModelDetailViaDax -Dataset -OutputFolder                   -> writes "<CleanWs> ~ <CleanModel>.csv" and "..._MD.csv"
#   Get-IQModelBackupFileName -Dataset                               -> "<CleanWs> ~ <CleanModel>" (the base name used everywhere)
#   Get-IQUsageMetrics -WorkspaceId -WorkspaceName [-Days 30]        -> usage-metrics rows for the Extras stage
#
# Windows PowerShell 5.1 and PowerShell 7 compatible. Loaded by dot-sourcing from ImpactIQ.ps1, so $script:IQ is the
# shared context created by Initialize-IQContext (ImpactIQ.Common.ps1). Nothing here is Windows-only.
#
# Cross-module functions used (brief section 2): Write-IQLog, Get-IQCleanName, Get-IQSafeKey, Invoke-IQApi,
# ConvertTo-IQJsonFile, ConvertFrom-IQJsonFile. Private helpers are prefixed *-IQDax* / *-IQCsv* and are not part of
# the contract.
#
# TOM / TMSCHEMA enum codes reproduced as strings (Tabular Object Model enumerations; INFO.* returns the integer codes):
#   RelationshipEndCardinality  0=None 1=One 2=Many
#   CrossFilteringBehavior      1=OneDirection 2=BothDirections 3=Automatic
#   SecurityFilteringBehavior   1=OneDirection 2=BothDirections 3=None
#   ModeType (partition Mode)   0=Default 1=Import 2=DirectQuery 3=Push 4=Dual 5=DirectLake
#   PartitionSourceType (Type)  1=Query 2=Calculated 3=None 4=M 5=Entity 6=PolicyRange 7=CalculationGroup 8=Inferred
#   ColumnType (Type)           1=RowNumber 2=Data 3=Calculated 4=CalculatedTableColumn
#   ModelPermission (roles)     1=None 2=Read 3=ReadRefresh 4=Refresh 5=Administrator
#   MetadataPermission          1=Default 2=None 3=Read
# Newer engines may already return the names instead of the codes; ConvertTo-IQDaxEnumName accepts both.

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

function Get-IQModelDetailViaDax {
    <#
    .SYNOPSIS
    Produces "<CleanWs> ~ <CleanModel>.csv" and "..._MD.csv" for a dataset from INFO.* DAX queries (executeQueries).
    .DESCRIPTION
    Reproduces the rows of "Model Detail Extract Script.csx" and "Measure Dependency Extract Script.csx" (brief section 13)
    from INFO.TABLES/COLUMNS/MEASURES/RELATIONSHIPS/PARTITIONS/ROLES/TABLEPERMISSIONS/CALCULATIONGROUPS/
    CALCULATIONITEMS/HIERARCHIES/LEVELS/CALCDEPENDENCY (+ INFO.MODEL for the default storage mode). Raw results are
    cached under extracts\dax\<datasetId>\*.json and reused on re-runs. Never throws; returns
    @{ Success; Csv; MdCsv; Outputs; Message; InfoUnsupported; RowCount; DependencyRowCount; Method='Dax'; BaseName }.
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
    $result = @{ Success = $false; Csv = $null; MdCsv = $null; Outputs = @(); Message = ''; InfoUnsupported = $false; RowCount = 0; DependencyRowCount = 0; Method = 'Dax'; BaseName = $BaseName }
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

    # 1. INFO.TABLES first: an error here means INFO functions / executeQueries are unavailable for this model.
    $tables = @()
    try {
        $tables = @(Get-IQDaxInfoRowSet @queryParams -Name 'tables' -Dax 'EVALUATE INFO.TABLES()')
    }
    catch {
        $msg = $_.Exception.Message
        $result.Message = 'INFO.TABLES() failed: ' + $msg
        if ($msg -match '(?i)\bINFO\b|not supported|Unsupported') { $result.InfoUnsupported = $true }
        Write-IQLog -Level Warn -Stage $Stage -Item $item -Message ("DAX model detail unavailable: {0}" -f $result.Message)
        return $result
    }

    # 2. The remaining queries. Each is guarded: a missing INFO function (older engine) yields an empty set and a Warn.
    $queries = [ordered]@{
        model            = 'EVALUATE INFO.MODEL()'
        columns          = 'EVALUATE INFO.COLUMNS()'
        measures         = 'EVALUATE INFO.MEASURES()'
        relationships    = 'EVALUATE INFO.RELATIONSHIPS()'
        partitions       = 'EVALUATE INFO.PARTITIONS()'
        roles            = 'EVALUATE INFO.ROLES()'
        tablepermissions = 'EVALUATE INFO.TABLEPERMISSIONS()'
        calcgroups       = 'EVALUATE INFO.CALCULATIONGROUPS()'
        calcitems        = 'EVALUATE INFO.CALCULATIONITEMS()'
        hierarchies      = 'EVALUATE INFO.HIERARCHIES()'
        levels           = 'EVALUATE INFO.LEVELS()'
        calcdependency   = 'EVALUATE INFO.CALCDEPENDENCY()'
    }
    $data = @{ tables = $tables }
    $warnings = @()
    foreach ($name in $queries.Keys) {
        try {
            $data[$name] = @(Get-IQDaxInfoRowSet @queryParams -Name $name -Dax $queries[$name])
        }
        catch {
            $data[$name] = @()
            $warnings += ('{0}: {1}' -f $name, $_.Exception.Message)
            Write-IQLog -Level Warn -Stage $Stage -Item $item -Message ("{0} failed; rows of that kind will be missing: {1}" -f $queries[$name], $_.Exception.Message)
        }
    }

    # Enum maps.
    $modeMap = @{ 0 = 'Default'; 1 = 'Import'; 2 = 'DirectQuery'; 3 = 'Push'; 4 = 'Dual'; 5 = 'DirectLake' }
    $cardinalityMap = @{ 0 = 'None'; 1 = 'One'; 2 = 'Many' }
    $crossFilterMap = @{ 1 = 'OneDirection'; 2 = 'BothDirections'; 3 = 'Automatic' }

    # Default storage mode of the model (used when a partition inherits Mode=Default; audit X1-19).
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

    # Indexes.
    $tableIndex = Get-IQDaxIndex -Rows $data['tables']
    $columnIndex = Get-IQDaxIndex -Rows $data['columns']
    $hierarchyIndex = Get-IQDaxIndex -Rows $data['hierarchies']
    $roleIndex = Get-IQDaxIndex -Rows $data['roles']
    $calcGroupTableIds = @{}
    foreach ($cg in @($data['calcgroups'])) {
        $tid = Get-IQDaxMember -Object $cg -Name 'TableID'
        if ($null -ne $tid) { $calcGroupTableIds[[string]$tid] = $true }
    }
    $calcGroupTableNames = @{}
    foreach ($tid in $calcGroupTableIds.Keys) {
        $n = Get-IQDaxIndexedName -Index $tableIndex -Id $tid
        if ($n) { $calcGroupTableNames[$n] = $true }
    }
    # First (lowest ID) partition per table gives the table storage mode (TE takes Partitions[0]).
    $firstPartitionMode = @{}
    foreach ($p in (@($data['partitions']) | Sort-Object { [double](Get-IQDaxMember -Object $_ -Name 'ID') })) {
        $tid = [string](Get-IQDaxMember -Object $p -Name 'TableID')
        if ($tid -eq '') { continue }
        if (-not $firstPartitionMode.ContainsKey($tid)) { $firstPartitionMode[$tid] = (& $resolveMode (Get-IQDaxMember -Object $p -Name 'Mode')) }
    }

    $rows = New-Object System.Collections.Generic.List[object]

    # Tables (all, including calculation-group tables).
    foreach ($t in @($data['tables'])) {
        $tid = [string](Get-IQDaxMember -Object $t -Name 'ID')
        $name = [string](Get-IQDaxMember -Object $t -Name 'Name')
        $mode = ''
        if ($firstPartitionMode.ContainsKey($tid)) { $mode = $firstPartitionMode[$tid] }
        $rows.Add((New-IQModelDetailRow -Type 'Table' -Common $common -Fields @{
                    Table = $name; Name = $name; IsHidden = (Get-IQDaxMember -Object $t -Name 'IsHidden'); TableStorageMode = $mode
                    Description = (Get-IQDaxMember -Object $t -Name 'Description')
                }))
    }

    # Calculation groups (Table = Name = the calc-group table name) and their items.
    foreach ($cg in @($data['calcgroups'])) {
        $tid = [string](Get-IQDaxMember -Object $cg -Name 'TableID')
        $tableRow = $null
        if ($tableIndex.ContainsKey($tid)) { $tableRow = $tableIndex[$tid] }
        $groupName = Get-IQDaxIndexedName -Index $tableIndex -Id $tid
        $desc = Get-IQDaxMember -Object $cg -Name 'Description'
        if (($null -eq $desc -or [string]$desc -eq '') -and $null -ne $tableRow) { $desc = Get-IQDaxMember -Object $tableRow -Name 'Description' }
        $hidden = $null
        if ($null -ne $tableRow) { $hidden = Get-IQDaxMember -Object $tableRow -Name 'IsHidden' }
        $rows.Add((New-IQModelDetailRow -Type 'CalculationGroup' -Common $common -Fields @{ Table = $groupName; Name = $groupName; Description = $desc; IsHidden = $hidden }))
        $cgId = [string](Get-IQDaxMember -Object $cg -Name 'ID')
        foreach ($ci in @($data['calcitems'])) {
            if ([string](Get-IQDaxMember -Object $ci -Name 'CalculationGroupID') -ne $cgId) { continue }
            $rows.Add((New-IQModelDetailRow -Type 'CalculationItem' -Common $common -Fields @{
                        Table = $groupName; Name = (Get-IQDaxMember -Object $ci -Name 'Name')
                        Description = (Get-IQDaxMember -Object $ci -Name 'Description'); Expression = (Get-IQDaxMember -Object $ci -Name 'Expression')
                    }))
        }
    }

    # Columns: every non-RowNumber column as "Column"; calculated columns (Type=3) again as "CalculatedColumn" (mirrors the csx).
    $columnRows = @($data['columns'])
    foreach ($c in $columnRows) {
        $ctype = 0
        try { $ctype = [int](Get-IQDaxMember -Object $c -Name 'Type') } catch { $ctype = 0 }
        if ($ctype -eq 1) { continue }
        $rows.Add((New-IQModelDetailRow -Type 'Column' -Common $common -Fields @{
                    Table = (Get-IQDaxIndexedName -Index $tableIndex -Id (Get-IQDaxMember -Object $c -Name 'TableID')); Name = (Get-IQDaxColumnName -Column $c)
                    FormatString = (Get-IQDaxMember -Object $c -Name 'FormatString'); DisplayFolder = (Get-IQDaxMember -Object $c -Name 'DisplayFolder')
                    Description = (Get-IQDaxMember -Object $c -Name 'Description'); IsHidden = (Get-IQDaxMember -Object $c -Name 'IsHidden')
                }))
    }
    foreach ($c in $columnRows) {
        $ctype = 0
        try { $ctype = [int](Get-IQDaxMember -Object $c -Name 'Type') } catch { $ctype = 0 }
        if ($ctype -ne 3) { continue }
        $rows.Add((New-IQModelDetailRow -Type 'CalculatedColumn' -Common $common -Fields @{
                    Table = (Get-IQDaxIndexedName -Index $tableIndex -Id (Get-IQDaxMember -Object $c -Name 'TableID')); Name = (Get-IQDaxColumnName -Column $c)
                    FormatString = (Get-IQDaxMember -Object $c -Name 'FormatString'); DisplayFolder = (Get-IQDaxMember -Object $c -Name 'DisplayFolder')
                    Description = (Get-IQDaxMember -Object $c -Name 'Description'); IsHidden = (Get-IQDaxMember -Object $c -Name 'IsHidden')
                    Expression = (Get-IQDaxMember -Object $c -Name 'Expression')
                }))
    }

    # Measures.
    foreach ($m in @($data['measures'])) {
        $rows.Add((New-IQModelDetailRow -Type 'Measure' -Common $common -Fields @{
                    Table = (Get-IQDaxIndexedName -Index $tableIndex -Id (Get-IQDaxMember -Object $m -Name 'TableID')); Name = (Get-IQDaxMember -Object $m -Name 'Name')
                    FormatString = (Get-IQDaxMember -Object $m -Name 'FormatString'); DisplayFolder = (Get-IQDaxMember -Object $m -Name 'DisplayFolder')
                    Description = (Get-IQDaxMember -Object $m -Name 'Description'); IsHidden = (Get-IQDaxMember -Object $m -Name 'IsHidden')
                    Expression = (Get-IQDaxMember -Object $m -Name 'Expression')
                }))
    }

    # Hierarchies and levels.
    foreach ($h in @($data['hierarchies'])) {
        $rows.Add((New-IQModelDetailRow -Type 'Hierarchy' -Common $common -Fields @{
                    Table = (Get-IQDaxIndexedName -Index $tableIndex -Id (Get-IQDaxMember -Object $h -Name 'TableID')); Name = (Get-IQDaxMember -Object $h -Name 'Name')
                    DisplayFolder = (Get-IQDaxMember -Object $h -Name 'DisplayFolder'); Description = (Get-IQDaxMember -Object $h -Name 'Description')
                    IsHidden = (Get-IQDaxMember -Object $h -Name 'IsHidden')
                }))
    }
    foreach ($l in @($data['levels'])) {
        $hid = [string](Get-IQDaxMember -Object $l -Name 'HierarchyID')
        $tableName = ''
        if ($hierarchyIndex.ContainsKey($hid)) { $tableName = Get-IQDaxIndexedName -Index $tableIndex -Id (Get-IQDaxMember -Object $hierarchyIndex[$hid] -Name 'TableID') }
        $rows.Add((New-IQModelDetailRow -Type 'Level' -Common $common -Fields @{
                    Table = $tableName; Name = (Get-IQDaxMember -Object $l -Name 'Name'); Description = (Get-IQDaxMember -Object $l -Name 'Description')
                }))
    }

    # Partitions (Expression = M / DAX / query text from QueryDefinition).
    foreach ($p in @($data['partitions'])) {
        $expr = Get-IQDaxMember -Object $p -Name 'QueryDefinition'
        if ($null -eq $expr -or [string]$expr -eq '') { $expr = Get-IQDaxMember -Object $p -Name 'Expression' }
        $rows.Add((New-IQModelDetailRow -Type 'Partition' -Common $common -Fields @{
                    Table = (Get-IQDaxIndexedName -Index $tableIndex -Id (Get-IQDaxMember -Object $p -Name 'TableID')); Name = (Get-IQDaxMember -Object $p -Name 'Name')
                    Description = (Get-IQDaxMember -Object $p -Name 'Description'); TableStorageMode = (& $resolveMode (Get-IQDaxMember -Object $p -Name 'Mode'))
                    Expression = $expr
                }))
    }

    # RLS filters: one row per table permission, Name = role name.
    foreach ($tp in @($data['tablepermissions'])) {
        $roleName = Get-IQDaxIndexedName -Index $roleIndex -Id (Get-IQDaxMember -Object $tp -Name 'RoleID')
        $rows.Add((New-IQModelDetailRow -Type 'RLSFilter' -Common $common -Fields @{
                    Table = (Get-IQDaxIndexedName -Index $tableIndex -Id (Get-IQDaxMember -Object $tp -Name 'TableID')); Name = $roleName
                    Expression = (Get-IQDaxMember -Object $tp -Name 'FilterExpression')
                }))
    }

    # Relationships.
    foreach ($r in @($data['relationships'])) {
        $fromTable = Get-IQDaxIndexedName -Index $tableIndex -Id (Get-IQDaxMember -Object $r -Name 'FromTableID')
        $toTable = Get-IQDaxIndexedName -Index $tableIndex -Id (Get-IQDaxMember -Object $r -Name 'ToTableID')
        $fromColumn = ''
        $toColumn = ''
        $fcid = [string](Get-IQDaxMember -Object $r -Name 'FromColumnID')
        $tcid = [string](Get-IQDaxMember -Object $r -Name 'ToColumnID')
        if ($columnIndex.ContainsKey($fcid)) { $fromColumn = Get-IQDaxColumnName -Column $columnIndex[$fcid] }
        if ($columnIndex.ContainsKey($tcid)) { $toColumn = Get-IQDaxColumnName -Column $columnIndex[$tcid] }
        $isActive = Get-IQDaxMember -Object $r -Name 'IsActive'
        $status = ''
        if ($null -ne $isActive) { try { if ([System.Convert]::ToBoolean($isActive)) { $status = 'True' } else { $status = 'False' } } catch { $status = [string]$isActive } }
        $rows.Add((New-IQModelDetailRow -Type 'Relationship' -Common $common -Fields @{
                    Table = $fromTable; Name = $fromColumn; Expression = (Get-IQDaxMember -Object $r -Name 'Name')
                    RelationshipFromTable = $fromTable; RelationshipFromColumn = $fromColumn; RelationshipToTable = $toTable; RelationshipToColumn = $toColumn
                    RelationshipStatus = $status
                    RelationshipFromCardinality = (ConvertTo-IQDaxEnumName -Value (Get-IQDaxMember -Object $r -Name 'FromCardinality') -Map $cardinalityMap)
                    RelationshipToCardinality = (ConvertTo-IQDaxEnumName -Value (Get-IQDaxMember -Object $r -Name 'ToCardinality') -Map $cardinalityMap)
                    RelationshipCrossFilteringBehavior = (ConvertTo-IQDaxEnumName -Value (Get-IQDaxMember -Object $r -Name 'CrossFilteringBehavior') -Map $crossFilterMap)
                }))
    }

    # Measure dependencies from INFO.CALCDEPENDENCY() (direct dependencies; brief section 13 mapping).
    $objectTypeMap = @{ 'MEASURE' = 'Measure'; 'CALC_COLUMN' = 'CalculatedColumn'; 'CALCULATION_ITEM' = 'CalculationItem'; 'CALC_ITEM' = 'CalculationItem' }
    $mdRows = New-Object System.Collections.Generic.List[object]
    $mdSeen = @{}
    foreach ($d in @($data['calcdependency'])) {
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
                if ($calcGroupTableNames.ContainsKey($refTable)) { $dependsOnType = 'CalculationGroupTable' } else { $dependsOnType = 'Table' }
            }
            'CALC_TABLE' {
                $dependsOn = ConvertTo-IQDaxTableRef -Table $refTable
                if ($calcGroupTableNames.ContainsKey($refTable)) { $dependsOnType = 'CalculationGroupTable' } else { $dependsOnType = 'Table' }
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
                ModelAsOfDate = $common.ModelAsOfDate; ModelName = $common.ModelName; ModelID = $common.ModelID
            })
    }

    try {
        $result.RowCount = Write-IQCsvFile -Path $csvPath -Header (Get-IQModelDetailHeader) -Rows $rows.ToArray()
        $result.DependencyRowCount = Write-IQCsvFile -Path $mdPath -Header (Get-IQMeasureDependencyHeader) -Rows $mdRows.ToArray()
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
    $result.Message = ('{0} object rows, {1} dependency rows via DAX INFO.*' -f $result.RowCount, $result.DependencyRowCount)
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
