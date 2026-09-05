# ImpactIQ.Dataflows.ps1 - Dataflows stage: Gen1 JSON backup, Gen2 (Fabric CI/CD) definition backup, query parsing.
#
# Contract: brief sections 2.7 and 8.3. Port of "Final PS Script.txt" lines 3298-3735 (Export-FabricDataflow,
# Parse-FabricDataflowContent, the Gen1 loop) with the audit C9 findings applied:
#   C9-01 every call is guarded, nothing stale is written, the pseudo "My Workspace" row is skipped
#   C9-02 Gen2 getDefinition goes through Invoke-IQFabricLro (202 / Location polling / result)
#   C9-03 / C9-11 / C9-16 one line-anchored tokenizer (Split-IQMSharedQuery) replaces the two lazy regexes:
#         nested-bracket attribute records, dotted / non-ASCII identifiers, doubled quotes in #"..." names, a flexible
#         "section X;" split, linear time, no $matches shadowing
#   C9-04 the Gen1 "unescape" only runs when the document is demonstrably still escaped
#   C9-06 per-dataflow checkpoints (Test-IQItemDone / Set-IQItemDone), no folder wipe, parsed queries persisted to
#         State\runs\<RunId>\extracts\dataflows\<safeKey>.json immediately (Assemble rebuilds the workbook from those)
#   C9-07 dataflow list comes from the inventory (ws-*.json) instead of re-listing; live listing only as a fallback
#   C9-08 every failure is logged and recorded in the manifest; summary line at the end
#   C9-09 ReportDate = RunId when it is a yyyy-MM-dd date (invariant), no "latest folder" heuristics
#   C9-10 exact bytes for .pq (WriteAllBytes), UTF-8 without BOM for the Gen1 .txt, name-collision suffix "~<id8>",
#         trailing "." / " " trimmed, > 240 char path warning
#   C9-12 / C9-13 queriesMetadata (Load Enabled, Query Group), Gen1 entities and every Gen2 definition part are kept
#   C9-14 cheap derived columns per query (Source Functions, Referenced Queries, Line Count, Uses Native Query)
#   C9-15 definition part paths are validated against the target folder before writing
#
# Files written (backup folder = <BaseFolder>\Dataflow Backups\<RunId>\; names use Get-IQCleanName, unchanged):
#   <CleanWs> ~ <CleanDf>.txt           Gen1: the raw "Get Dataflow" model.json body (UTF-8, no BOM)   [monolith name]
#   <CleanWs> ~ <CleanDf>.pq            Gen2: the mashup.pq definition part, byte-exact                  [monolith name]
#   <CleanWs> ~ <CleanDf>.definition\   Gen2: every definition part (queryMetadata.json, *.pq, *.mdf, .platform ...)
#   State\runs\<RunId>\extracts\dataflows\<safeDataflowId>.json   parsed queries + entities (schema in the header of
#                                        New-IQDataflowExtract). "Queries" rows carry the exact Sheet1 columns of
#                                        "Dataflow Detail.xlsx" ("Dataflow ID", "Dataflow Name", "Query Name", "Query",
#                                        "Report Date", "Workspace Name - Dataflow Name") followed by additive columns.
#
# Windows PowerShell 5.1 and PowerShell 7 compatible; nothing here is Windows-only. Dot-sourced from ImpactIQ.ps1, so
# $script:IQ is the shared context. Cross-module functions used (brief section 2): Write-IQLog, Get-IQCleanName,
# Get-IQSafeKey, ConvertTo-IQJsonFile, ConvertFrom-IQJsonFile, Invoke-IQApi, Invoke-IQFabricLro, Test-IQItemDone,
# Set-IQItemDone, Get-IQAllWorkspaceInventories, Get-IQSelectedWorkspaces. Private helpers are prefixed *-IQDataflow* /
# *-IQM* and are not part of the contract.

# =====================================================================================================================
# Small private helpers
# =====================================================================================================================

function Get-IQDataflowMember {
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

function Get-IQDataflowRunFolder {
    <#
    .SYNOPSIS
    The dataflow backup folder of the current run (<BaseFolder>\Dataflow Backups\<RunId>), created when missing (private).
    #>
    [CmdletBinding()]
    param()
    if (-not $script:IQ) { throw 'ImpactIQ context is not initialised (Initialize-IQContext).' }
    $folder = $null
    if ($script:IQ.ContainsKey('RunPaths') -and $null -ne $script:IQ.RunPaths -and $script:IQ.RunPaths.DataflowBackups) {
        $folder = [string]$script:IQ.RunPaths.DataflowBackups
    }
    if ([string]::IsNullOrWhiteSpace($folder)) {
        if ([string]::IsNullOrWhiteSpace([string]$script:IQ.RunId)) { throw 'No active run (Initialize-IQRun has not been called).' }
        $folder = Join-Path $script:IQ.Paths.DataflowBackups $script:IQ.RunId
    }
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    return $folder
}

function Get-IQDataflowExtractFolder {
    <#
    .SYNOPSIS
    State\runs\<RunId>\extracts\dataflows (created when missing) - one JSON per dataflow with the parsed queries (private).
    #>
    [CmdletBinding()]
    param()
    if (-not $script:IQ) { throw 'ImpactIQ context is not initialised (Initialize-IQContext).' }
    $extracts = $null
    if ($script:IQ.ContainsKey('RunPaths') -and $null -ne $script:IQ.RunPaths -and $script:IQ.RunPaths.Extracts) { $extracts = [string]$script:IQ.RunPaths.Extracts }
    if ([string]::IsNullOrWhiteSpace($extracts)) {
        if ([string]::IsNullOrWhiteSpace([string]$script:IQ.RunPath)) { throw 'No active run (Initialize-IQRun has not been called).' }
        $extracts = Join-Path $script:IQ.RunPath 'extracts'
    }
    $folder = Join-Path $extracts 'dataflows'
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    return $folder
}

function Get-IQDataflowReportDate {
    <#
    .SYNOPSIS
    The "Report Date" value (yyyy-MM-dd): the RunId when it is a date, else today (audit C9-09, invariant culture) (private).
    #>
    [CmdletBinding()]
    param()
    $runId = ''
    if ($script:IQ) { $runId = [string]$script:IQ.RunId }
    if ($runId -match '^\d{4}-\d{2}-\d{2}$') {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact($runId, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            return $parsed.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
        }
    }
    return (Get-Date).ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-IQDataflowFileStem {
    <#
    .SYNOPSIS
    "<CleanWs> ~ <CleanDf>" via Get-IQCleanName (monolith naming), with trailing "." / " " trimmed (audit C9-10) (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$WorkspaceName,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$DataflowName
    )
    $stem = (Get-IQCleanName -Name $WorkspaceName) + ' ~ ' + (Get-IQCleanName -Name $DataflowName)
    $stem = $stem.TrimEnd('.', ' ')
    if ([string]::IsNullOrWhiteSpace($stem)) { $stem = 'Dataflow' }
    return $stem
}

function Write-IQDataflowTextFile {
    <#
    .SYNOPSIS
    Writes text as UTF-8 without BOM through a temp file + Move-Item -Force so a half-written backup never looks complete (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Text
    )
    $dir = Split-Path -Path $Path -Parent
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = $Path + '.tmp'
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($tmp, [string]$Text, $enc)
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}

function Write-IQDataflowBinaryFile {
    <#
    .SYNOPSIS
    Writes bytes through a temp file + Move-Item -Force (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][byte[]]$Bytes
    )
    $dir = Split-Path -Path $Path -Parent
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = $Path + '.tmp'
    [System.IO.File]::WriteAllBytes($tmp, $Bytes)
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}

function Test-IQDataflowFileHasContent {
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

# =====================================================================================================================
# M (Power Query) section tokenizer - shared by Gen1 (pbi:mashup.document) and Gen2 (mashup.pq)
# =====================================================================================================================

function Test-IQMBalancedRecord {
    <#
    .SYNOPSIS
    $true when the text is exactly one M record literal "[ ... ]" (brackets balance to zero only at the last char; strings are skipped) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $false }
    $t = $Text.Trim()
    if ($t.Length -lt 2 -or $t[0] -ne '[' -or $t[$t.Length - 1] -ne ']') { return $false }
    $depth = 0
    $inString = $false
    $last = $t.Length - 1
    for ($i = 0; $i -le $last; $i++) {
        $c = $t[$i]
        if ($inString) {
            if ($c -eq '"') {
                if ($i -lt $last -and $t[$i + 1] -eq '"') { $i++ }   # doubled quote inside a string
                else { $inString = $false }
            }
            continue
        }
        if ($c -eq '"') { $inString = $true; continue }
        if ($c -eq '[') { $depth++; continue }
        if ($c -eq ']') {
            $depth--
            if ($depth -lt 0) { return $false }
            if ($depth -eq 0 -and $i -ne $last) { return $false }   # closed before the end: "[a=1] x"
        }
    }
    return ($depth -eq 0 -and -not $inString)
}

function Split-IQMTrailingAttribute {
    <#
    .SYNOPSIS
    Peels a trailing own-line attribute record "[...]" (belongs to the NEXT shared member) off a query body (private).
    .DESCRIPTION
    Returns @{ Body; Attribute }. The record must start at the beginning of a line, be a single balanced record that
    runs to the end of the text, and the text before it must be empty or end with ";" (the end of the previous member),
    so records that are part of the query expression itself are never peeled (audit C9-03).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Text)
    $result = @{ Body = [string]$Text; Attribute = '' }
    if ([string]::IsNullOrWhiteSpace($Text)) { return $result }
    $trimmedEnd = $Text.TrimEnd()
    if (-not $trimmedEnd.EndsWith(']')) { return $result }

    # Candidate start offsets: every line whose first non-blank character is "[" (checked from the last one backwards).
    $candidates = New-Object System.Collections.Generic.List[int]
    $pos = 0
    while ($pos -le $Text.Length) {
        $lineEnd = $Text.IndexOf("`n", $pos)
        if ($lineEnd -lt 0) { $lineEnd = $Text.Length }
        $line = $Text.Substring($pos, $lineEnd - $pos)
        $trimmedLine = $line.TrimStart()
        if ($trimmedLine.StartsWith('[')) { $candidates.Add($pos + ($line.Length - $trimmedLine.Length)) }
        if ($lineEnd -ge $Text.Length) { break }
        $pos = $lineEnd + 1
    }
    for ($i = $candidates.Count - 1; $i -ge 0; $i--) {
        $offset = $candidates[$i]
        $before = $Text.Substring(0, $offset).TrimEnd()
        if ($before.Length -gt 0 -and -not $before.EndsWith(';')) { continue }
        $candidate = $Text.Substring($offset).Trim()
        if (Test-IQMBalancedRecord -Text $candidate) {
            $result.Body = $Text.Substring(0, $offset)
            $result.Attribute = $candidate
            return $result
        }
    }
    return $result
}

function Split-IQMSharedQuery {
    <#
    .SYNOPSIS
    Tokenises the body of an M section into its "shared" members (audit C9-03 tokenizer; replaces the monolith regex).
    .DESCRIPTION
    Headers are matched line-anchored: optional same-line attribute record, "shared", then #"quoted name" (doubled
    quotes un-doubled) or a regular identifier (Unicode letters, digits, "_" and "." allowed). The text between two
    headers is the member expression; a trailing own-line attribute record (e.g. Gen2 "[DataDestinations = {...}]",
    "[StagingDefinition = ...]") is peeled off and attached to the following member. Returns objects
    @{ Name; Expression (trailing ";" removed, trimmed); Attributes (record text or ''); Index }.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Section)
    $out = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($Section)) { return $out.ToArray() }

    $headerRegex = New-Object System.Text.RegularExpressions.Regex('(?m)^[ \t]*(?:(\[[^\r\n]*\])[ \t]*)?shared[ \t]+(?:#"((?:[^"]|"")*)"|([\p{L}_][\p{L}\p{N}_.]*))[ \t]*=')
    $headers = $headerRegex.Matches($Section)
    if ($headers.Count -eq 0) { return $out.ToArray() }

    # Text before the first header may be an attribute record for the first member.
    $pendingAttribute = ''
    $leading = $Section.Substring(0, $headers[0].Index)
    if (-not [string]::IsNullOrWhiteSpace($leading)) {
        $peeled = Split-IQMTrailingAttribute -Text $leading
        if ($peeled.Attribute) { $pendingAttribute = $peeled.Attribute }
    }

    for ($i = 0; $i -lt $headers.Count; $i++) {
        $m = $headers[$i]
        $start = $m.Index + $m.Length
        $end = $Section.Length
        if ($i + 1 -lt $headers.Count) { $end = $headers[$i + 1].Index }
        $body = $Section.Substring($start, $end - $start)

        $peeled = Split-IQMTrailingAttribute -Text $body
        $body = $peeled.Body

        $name = ''
        if ($m.Groups[2].Success) { $name = $m.Groups[2].Value -replace '""', '"' }
        elseif ($m.Groups[3].Success) { $name = $m.Groups[3].Value }
        $attribute = $pendingAttribute
        if ($m.Groups[1].Success -and -not [string]::IsNullOrWhiteSpace($m.Groups[1].Value)) { $attribute = $m.Groups[1].Value.Trim() }
        $pendingAttribute = [string]$peeled.Attribute

        $expression = $body.Trim()
        $expression = $expression -replace ';\s*$', ''   # monolith: remove the trailing semicolon
        $expression = $expression.Trim()
        $out.Add([PSCustomObject]@{ Name = $name; Expression = $expression; Attributes = $attribute; Index = $i })
    }
    return $out.ToArray()
}

function Get-IQMSectionBody {
    <#
    .SYNOPSIS
    Returns the text after the "section <Name>;" header of an M document (fallback: whole document when no header) (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Content,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Item
    )
    if ([string]::IsNullOrWhiteSpace($Content)) { return '' }
    # Monolith split on the literal 'section Section1;' - accept any section name and spacing (audit C9-11).
    # Everything after the FIRST section header is the member list (an M document has exactly one section).
    $header = [regex]::Match($Content, '(?m)^[ \t]*section[ \t]+[^;\r\n]+;[ \t]*\r?$')
    if ($header.Success) {
        return $Content.Substring($header.Index + $header.Length)
    }
    if ($Content -match '(?m)^[ \t]*(?:\[[^\r\n]*\][ \t]*)?shared[ \t]+') {
        Write-IQLog -Level Debug -Stage 'Dataflows' -Item $Item -Message 'No "section <name>;" header found; parsing the whole document'
        return $Content
    }
    return ''
}

function Get-IQMDerivedColumn {
    <#
    .SYNOPSIS
    Cheap per-query governance columns from the expression text (audit C9-14) (private).
    .DESCRIPTION
    Returns @{ SourceFunctions; ReferencedQueries; LineCount; UsesNativeQuery } - source connector functions
    (Sql.Database, Web.Contents, ...), the other shared members referenced by the expression, line count and a
    heuristic native-query flag (Value.NativeQuery or an options record with Query = "...").
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Expression,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$OwnName,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()][string[]]$AllNames
    )
    $result = @{ SourceFunctions = ''; ReferencedQueries = ''; LineCount = 0; UsesNativeQuery = $false }
    if ([string]::IsNullOrEmpty($Expression)) { return $result }

    $sourcePattern = '\b([A-Z][A-Za-z0-9]*\.(?:Database|Databases|Contents|Files|Tables|Dataflows|Workspaces|Query|Feed|Cube|Cubes|Document|Workbook|DataSource|Folders|Blobs|Containers|Lakehouse|Warehouse|Data|Catalog|Databricks|Dataset))\s*\('
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($m in [regex]::Matches($Expression, $sourcePattern)) {
        $fn = $m.Groups[1].Value
        if ($fn -like 'Table.*' -or $fn -like 'List.*' -or $fn -like 'Record.*' -or $fn -like 'Text.*' -or $fn -like 'Value.*') { continue }
        if (-not $found.Contains($fn)) { $found.Add($fn) }
    }
    $result.SourceFunctions = ($found -join '; ')

    $refs = New-Object System.Collections.Generic.List[string]
    foreach ($n in @($AllNames)) {
        if ([string]::IsNullOrEmpty($n) -or $n -eq $OwnName) { continue }
        $escaped = [regex]::Escape($n)
        $pattern = $null
        if ($n -match '^[\p{L}_][\p{L}\p{N}_.]*$') { $pattern = '(?:#"' + $escaped + '"|(?<![\p{L}\p{N}_."#])' + $escaped + '(?![\p{L}\p{N}_."]))' }
        else { $pattern = '#"' + [regex]::Escape(($n -replace '"', '""')) + '"' }
        if ([regex]::IsMatch($Expression, $pattern)) { if (-not $refs.Contains($n)) { $refs.Add($n) } }
    }
    $result.ReferencedQueries = ($refs -join '; ')

    $result.LineCount = ([regex]::Matches($Expression, "`n")).Count + 1
    $result.UsesNativeQuery = [bool]([regex]::IsMatch($Expression, '(?s)Value\.NativeQuery\s*\(|\[[^\]]*\bQuery\s*=\s*"'))
    return $result
}

function Get-IQDataflowQueryMetaMap {
    <#
    .SYNOPSIS
    Normalises "queriesMetadata" (object keyed by query name, or an array) into a hashtable name -> @{ LoadEnabled; IsHidden; QueryGroupId; QueryGroup } (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$QueriesMetadata,
        [Parameter(Mandatory = $false)][AllowNull()]$QueryGroups
    )
    $groupNames = @{}
    foreach ($g in @($QueryGroups)) {
        if ($null -eq $g) { continue }
        $gid = [string](Get-IQDataflowMember -Object $g -Name 'id')
        $gname = Get-IQDataflowMember -Object $g -Name 'name'
        if ($null -eq $gname) { $gname = Get-IQDataflowMember -Object $g -Name 'displayName' }
        if ($gid -and $null -ne $gname) { $groupNames[$gid] = [string]$gname }
    }
    $map = @{}
    if ($null -eq $QueriesMetadata) { return $map }
    $entries = @()
    if ($QueriesMetadata -is [System.Collections.IDictionary]) {
        foreach ($k in $QueriesMetadata.Keys) { $entries += , @{ Key = [string]$k; Value = $QueriesMetadata[$k] } }
    }
    elseif ($QueriesMetadata -is [System.Management.Automation.PSCustomObject]) {
        foreach ($p in $QueriesMetadata.PSObject.Properties) { $entries += , @{ Key = [string]$p.Name; Value = $p.Value } }
    }
    elseif ($QueriesMetadata -is [System.Collections.IEnumerable] -and -not ($QueriesMetadata -is [string])) {
        foreach ($v in $QueriesMetadata) {
            if ($null -eq $v) { continue }
            $qn = Get-IQDataflowMember -Object $v -Name 'queryName'
            if ($null -eq $qn) { $qn = Get-IQDataflowMember -Object $v -Name 'name' }
            $entries += , @{ Key = [string]$qn; Value = $v }
        }
    }
    foreach ($e in $entries) {
        $v = $e.Value
        $name = [string](Get-IQDataflowMember -Object $v -Name 'queryName')
        if ([string]::IsNullOrEmpty($name)) { $name = [string]$e.Key }
        if ([string]::IsNullOrEmpty($name)) { continue }
        $gid = [string](Get-IQDataflowMember -Object $v -Name 'queryGroupId')
        $gname = $null
        if ($gid -and $groupNames.ContainsKey($gid)) { $gname = $groupNames[$gid] }
        elseif ($gid) { $gname = $gid }
        $map[$name] = @{
            LoadEnabled  = (Get-IQDataflowMember -Object $v -Name 'loadEnabled')
            IsHidden     = (Get-IQDataflowMember -Object $v -Name 'isHidden')
            QueryGroupId = $gid
            QueryGroup   = $gname
            QueryId      = [string](Get-IQDataflowMember -Object $v -Name 'queryId')
        }
    }
    return $map
}

function ConvertTo-IQDataflowQueryRow {
    <#
    .SYNOPSIS
    Builds the "Dataflow Detail.xlsx" Sheet1 rows (exact monolith columns first, additive columns after) from tokenised members (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()][array]$Members,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$DataflowId,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$DataflowName,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$WorkspaceName,
        [Parameter(Mandatory = $true)][string]$ReportDate,
        [Parameter(Mandatory = $false)][AllowNull()][hashtable]$QueryMetadata
    )
    $rows = New-Object System.Collections.Generic.List[object]
    $stem = (Get-IQCleanName -Name $WorkspaceName) + ' ~ ' + (Get-IQCleanName -Name $DataflowName)   # monolith "Workspace Name - Dataflow Name"
    $names = @()
    foreach ($m in @($Members)) { if ($null -ne $m -and -not [string]::IsNullOrWhiteSpace([string]$m.Name)) { $names += [string]$m.Name } }
    foreach ($m in @($Members)) {
        if ($null -eq $m) { continue }
        $queryName = [string]$m.Name
        $queryExpression = [string]$m.Expression
        # Monolith: skip empty names / expressions.
        if ([string]::IsNullOrWhiteSpace($queryName) -or [string]::IsNullOrWhiteSpace($queryExpression)) { continue }
        $derived = Get-IQMDerivedColumn -Expression $queryExpression -OwnName $queryName -AllNames $names
        $loadEnabled = $null
        $isHidden = $null
        $group = $null
        if ($null -ne $QueryMetadata -and $QueryMetadata.ContainsKey($queryName)) {
            $meta = $QueryMetadata[$queryName]
            $loadEnabled = $meta.LoadEnabled
            $isHidden = $meta.IsHidden
            $group = $meta.QueryGroup
        }
        $rows.Add([PSCustomObject]([ordered]@{
                    'Dataflow ID'                    = $DataflowId
                    'Dataflow Name'                  = $DataflowName
                    'Query Name'                     = $queryName
                    'Query'                          = $queryExpression
                    'Report Date'                    = $ReportDate
                    'Workspace Name - Dataflow Name' = $stem
                    'Load Enabled'                   = $loadEnabled
                    'Is Hidden'                      = $isHidden
                    'Query Group'                    = $group
                    'Attributes'                     = [string]$m.Attributes
                    'Source Functions'               = $derived.SourceFunctions
                    'Referenced Queries'             = $derived.ReferencedQueries
                    'Line Count'                     = $derived.LineCount
                    'Uses Native Query'              = $derived.UsesNativeQuery
                    'Query Length'                   = $queryExpression.Length
                }))
    }
    return $rows.ToArray()
}

function ConvertFrom-IQDataflowDocument {
    <#
    .SYNOPSIS
    Parses an M document (Gen1 pbi:mashup.document or Gen2 mashup.pq) into "Dataflow Detail.xlsx" Sheet1 rows.
    .DESCRIPTION
    Public replacement for the monolith's Parse-FabricDataflowContent and the inline Gen1 regex block: splits on the
    "section <name>;" header (any name) and tokenises the shared members with Split-IQMSharedQuery. Rows carry the
    exact monolith columns ("Dataflow ID", "Dataflow Name", "Query Name", "Query", "Report Date",
    "Workspace Name - Dataflow Name") plus the additive columns documented in ConvertTo-IQDataflowQueryRow.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Content,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$DataflowId,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$DataflowName,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$WorkspaceName,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$ReportDate,
        [Parameter(Mandatory = $false)][AllowNull()][hashtable]$QueryMetadata,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Item
    )
    if ([string]::IsNullOrWhiteSpace($ReportDate)) { $ReportDate = Get-IQDataflowReportDate }
    $section = Get-IQMSectionBody -Content $Content -Item $Item
    $members = @(Split-IQMSharedQuery -Section $section)
    return @(ConvertTo-IQDataflowQueryRow -Members $members -DataflowId $DataflowId -DataflowName $DataflowName -WorkspaceName $WorkspaceName -ReportDate $ReportDate -QueryMetadata $QueryMetadata)
}

# =====================================================================================================================
# Gen1 (Power BI "Get Dataflow" model.json)
# =====================================================================================================================

function Get-IQDataflowGen1Document {
    <#
    .SYNOPSIS
    Extracts pbi:mashup.document from a parsed Gen1 model.json, un-escaping only when the text is still JSON-escaped (audit C9-04) (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$ModelJson)
    if ($null -eq $ModelJson) { return $null }
    $mashup = Get-IQDataflowMember -Object $ModelJson -Name 'pbi:mashup'
    if ($null -eq $mashup) { return $null }
    $document = Get-IQDataflowMember -Object $mashup -Name 'document'
    if ($null -eq $document) { return $null }
    $text = [string]$document
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    # ConvertFrom-Json already un-escaped the string. Only a double-escaped document (no real newline but literal \n
    # sequences) needs the monolith's replacements; running them unconditionally corrupts UNC paths like "\\srv\new".
    if ($text -notmatch "`n" -and $text -match '\\r\\n|\\n') {
        $text = $text -replace '\\r\\n', "`n" -replace '\\n', "`n" -replace '\\"', '"'
    }
    return $text
}

function ConvertTo-IQDataflowEntityRow {
    <#
    .SYNOPSIS
    Entity / attribute rows from a Gen1 model.json ("entities[]" with attributes and pbi:refreshPolicy) (audit C9-12) (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$ModelJson,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$DataflowId,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$DataflowName,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$WorkspaceName
    )
    $rows = New-Object System.Collections.Generic.List[object]
    if ($null -eq $ModelJson) { return $rows.ToArray() }
    $stem = (Get-IQCleanName -Name $WorkspaceName) + ' ~ ' + (Get-IQCleanName -Name $DataflowName)
    foreach ($entity in @(Get-IQDataflowMember -Object $ModelJson -Name 'entities')) {
        if ($null -eq $entity) { continue }
        $entityName = [string](Get-IQDataflowMember -Object $entity -Name 'name')
        $entityType = [string](Get-IQDataflowMember -Object $entity -Name '$type')
        $description = Get-IQDataflowMember -Object $entity -Name 'description'
        $policy = Get-IQDataflowMember -Object $entity -Name 'pbi:refreshPolicy'
        $policyJson = ''
        if ($null -ne $policy) { try { $policyJson = ConvertTo-Json -InputObject $policy -Depth 20 -Compress } catch { $policyJson = [string]$policy } }
        $incremental = ($null -ne $policy)
        $attributes = @(Get-IQDataflowMember -Object $entity -Name 'attributes')
        $partitionCount = @(Get-IQDataflowMember -Object $entity -Name 'partitions').Count
        if ($attributes.Count -eq 0) {
            $rows.Add([PSCustomObject]([ordered]@{
                        'Dataflow ID'                    = $DataflowId
                        'Dataflow Name'                  = $DataflowName
                        'Entity'                         = $entityName
                        'Entity Type'                    = $entityType
                        'Column'                         = ''
                        'Data Type'                      = ''
                        'Incremental Refresh'            = $incremental
                        'Refresh Policy JSON'            = $policyJson
                        'Description'                    = $description
                        'Partition Count'                = $partitionCount
                        'Workspace Name - Dataflow Name' = $stem
                    }))
            continue
        }
        foreach ($attr in $attributes) {
            if ($null -eq $attr) { continue }
            $rows.Add([PSCustomObject]([ordered]@{
                        'Dataflow ID'                    = $DataflowId
                        'Dataflow Name'                  = $DataflowName
                        'Entity'                         = $entityName
                        'Entity Type'                    = $entityType
                        'Column'                         = [string](Get-IQDataflowMember -Object $attr -Name 'name')
                        'Data Type'                      = [string](Get-IQDataflowMember -Object $attr -Name 'dataType')
                        'Incremental Refresh'            = $incremental
                        'Refresh Policy JSON'            = $policyJson
                        'Description'                    = $description
                        'Partition Count'                = $partitionCount
                        'Workspace Name - Dataflow Name' = $stem
                    }))
        }
    }
    return $rows.ToArray()
}

function Export-IQGen1Dataflow {
    <#
    .SYNOPSIS
    Backs up one Gen1 dataflow (GET groups/{ws}/dataflows/{id}, raw body -> "<CleanWs> ~ <CleanDf>.txt") and parses its queries (private).
    .DESCRIPTION
    Returns @{ Success; Message; BackupPath; Queries; Entities; Metadata; HasMashup }. Never throws: transport errors,
    empty bodies and unparsable JSON are returned as Success = $false with a message (audit C9-01).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Work,
        [Parameter(Mandatory = $true)][string]$ReportDate
    )
    $stage = 'Dataflows'
    $result = @{ Success = $false; Message = ''; BackupPath = $Work.BackupPath; Queries = @(); Entities = @(); Metadata = @{}; HasMashup = $false; Method = 'PowerBI-GetDataflow' }
    $path = 'groups/' + $Work.WorkspaceId + '/dataflows/' + $Work.DataflowId
    $body = $null
    try {
        $body = Invoke-IQApi -Method GET -Path $path -Raw -NoPaging -Stage $stage
    }
    catch {
        $result.Message = 'Dataflow export failed: ' + $_.Exception.Message
        return $result
    }
    if ($null -eq $body -or [string]::IsNullOrWhiteSpace([string]$body)) {
        $result.Message = 'Dataflow export returned no content (HTTP 400/403/404 - see the Warn line above)'
        return $result
    }
    $body = [string]$body
    try { Write-IQDataflowTextFile -Path $Work.BackupPath -Text $body }
    catch {
        $result.Message = 'Could not write the backup file ' + $Work.BackupPath + ': ' + $_.Exception.Message
        return $result
    }
    $modelJson = $null
    try { $modelJson = ConvertFrom-Json -InputObject $body }
    catch {
        $result.Message = 'Backup written but the response is not valid JSON: ' + $_.Exception.Message
        $result.Success = $true   # the raw backup is still valuable
        return $result
    }
    $document = Get-IQDataflowGen1Document -ModelJson $modelJson
    $mashup = Get-IQDataflowMember -Object $modelJson -Name 'pbi:mashup'
    $metadataMap = @{}
    if ($null -ne $mashup) {
        $metadataMap = Get-IQDataflowQueryMetaMap -QueriesMetadata (Get-IQDataflowMember -Object $mashup -Name 'queriesMetadata') -QueryGroups (Get-IQDataflowMember -Object $modelJson -Name 'pbi:QueryGroups')
        $result.Metadata = @{
            FetchedTime        = [string](Get-IQDataflowMember -Object $mashup -Name 'fetchedTime')
            AllowNativeQueries = (Get-IQDataflowMember -Object $mashup -Name 'allowNativeQueries')
            FastCombine        = (Get-IQDataflowMember -Object $mashup -Name 'fastCombine')
            Culture            = [string](Get-IQDataflowMember -Object $modelJson -Name 'culture')
            ModifiedTime       = [string](Get-IQDataflowMember -Object $modelJson -Name 'modifiedTime')
            Version            = [string](Get-IQDataflowMember -Object $modelJson -Name 'version')
        }
    }
    $result.Entities = @(ConvertTo-IQDataflowEntityRow -ModelJson $modelJson -DataflowId $Work.DataflowId -DataflowName $Work.DataflowName -WorkspaceName $Work.WorkspaceName)
    if ([string]::IsNullOrWhiteSpace($document)) {
        $result.Success = $true
        $result.Message = 'Backup written; the export has no pbi:mashup document (classic Gen2 or empty dataflow) - no queries parsed'
        return $result
    }
    $result.HasMashup = $true
    $result.Queries = @(ConvertFrom-IQDataflowDocument -Content $document -DataflowId $Work.DataflowId -DataflowName $Work.DataflowName -WorkspaceName $Work.WorkspaceName -ReportDate $ReportDate -QueryMetadata $metadataMap -Item $Work.Item)
    $result.Success = $true
    $result.Message = ('{0} quer{1} parsed' -f $result.Queries.Count, $(if ($result.Queries.Count -eq 1) { 'y' } else { 'ies' }))
    return $result
}

# =====================================================================================================================
# Gen2 CI/CD (Fabric getDefinition)
# =====================================================================================================================

function Test-IQDataflowPartPathSafe {
    <#
    .SYNOPSIS
    $true when the definition part path stays inside the target folder (no "..", no rooted paths) (audit C9-15) (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Folder,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$PartPath
    )
    if ([string]::IsNullOrWhiteSpace($PartPath)) { return $false }
    if ($PartPath -match '^[A-Za-z]:' -or $PartPath.StartsWith('\') -or $PartPath.StartsWith('/')) { return $false }
    $relative = $PartPath -replace '/', [System.IO.Path]::DirectorySeparatorChar -replace '\\', [System.IO.Path]::DirectorySeparatorChar
    try {
        $root = [System.IO.Path]::GetFullPath($Folder).TrimEnd([System.IO.Path]::DirectorySeparatorChar) + [System.IO.Path]::DirectorySeparatorChar
        $full = [System.IO.Path]::GetFullPath((Join-Path $Folder $relative))
        return $full.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)
    }
    catch { return $false }
}

function Export-IQFabricDataflow {
    <#
    .SYNOPSIS
    Backs up one Gen2 CI/CD dataflow via Fabric getDefinition (LRO-aware) and parses its mashup.pq (port of Export-FabricDataflow).
    .DESCRIPTION
    POST workspaces/{ws}/dataflows/{id}/getDefinition through Invoke-IQFabricLro (audit C9-02). Every definition part
    is decoded into "<CleanWs> ~ <CleanDf>.definition\<part path>" (C9-13, paths validated per C9-15); the mashup .pq
    part is also written byte-exact as "<CleanWs> ~ <CleanDf>.pq" (the monolith's backup file, C9-10); queryMetadata.json
    supplies Load Enabled / Is Hidden / Query Group. Returns @{ Success; Message; BackupPath; DefinitionFolder;
    Queries; Entities; Metadata; PartsCount }. Never throws.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Work,
        [Parameter(Mandatory = $true)][string]$ReportDate
    )
    $stage = 'Dataflows'
    $result = @{ Success = $false; Message = ''; BackupPath = $Work.BackupPath; DefinitionFolder = $Work.DefinitionFolder; Queries = @(); Entities = @(); Metadata = @{}; PartsCount = 0; Method = 'Fabric-getDefinition' }
    $path = 'workspaces/' + $Work.WorkspaceId + '/dataflows/' + $Work.DataflowId + '/getDefinition'
    $response = $null
    try {
        $response = Invoke-IQFabricLro -Method POST -Path $path -TimeoutMinutes 10 -Stage $stage
    }
    catch {
        $result.Message = 'Fabric getDefinition failed: ' + $_.Exception.Message
        return $result
    }
    if ($null -eq $response) {
        $result.Message = 'Fabric getDefinition returned nothing (no Fabric token for this environment, no read/write permission on the dataflow, an encrypted sensitivity label, or the operation failed - see the Warn/Debug lines above)'
        return $result
    }
    $definition = Get-IQDataflowMember -Object $response -Name 'definition'
    $parts = @()
    if ($null -ne $definition) { $parts = @(Get-IQDataflowMember -Object $definition -Name 'parts') }
    $parts = @($parts | Where-Object { $null -ne $_ })
    if ($parts.Count -eq 0) {
        $result.Message = 'Fabric getDefinition returned a definition without parts'
        return $result
    }

    $folder = $Work.DefinitionFolder
    try {
        if (Test-Path -LiteralPath $folder) { Remove-Item -LiteralPath $folder -Recurse -Force -ErrorAction Stop }
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    catch {
        $result.Message = 'Could not prepare the definition folder ' + $folder + ': ' + $_.Exception.Message
        return $result
    }

    $pqPart = $null
    $pqBytes = $null
    $metadataBytes = $null
    $written = 0
    foreach ($part in $parts) {
        $partPath = [string](Get-IQDataflowMember -Object $part -Name 'path')
        $payloadType = [string](Get-IQDataflowMember -Object $part -Name 'payloadType')
        $payload = Get-IQDataflowMember -Object $part -Name 'payload'
        if (-not (Test-IQDataflowPartPathSafe -Folder $folder -PartPath $partPath)) {
            Write-IQLog -Level Warn -Stage $stage -Item $Work.Item -Message ("Skipping definition part with an unsafe path '{0}'" -f $partPath)
            continue
        }
        $bytes = $null
        try {
            if ($payloadType -eq 'InlineBase64') { $bytes = [System.Convert]::FromBase64String([string]$payload) }
            elseif ($null -ne $payload) { $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$payload) }
            else { $bytes = New-Object byte[] 0 }
        }
        catch {
            Write-IQLog -Level Warn -Stage $stage -Item $Work.Item -Message ("Definition part '{0}' could not be decoded ({1}): {2}" -f $partPath, $payloadType, $_.Exception.Message)
            continue
        }
        $relative = $partPath -replace '/', [System.IO.Path]::DirectorySeparatorChar
        $target = Join-Path $folder $relative
        try { Write-IQDataflowBinaryFile -Path $target -Bytes $bytes; $written++ }
        catch {
            Write-IQLog -Level Warn -Stage $stage -Item $Work.Item -Message ("Definition part '{0}' could not be written: {1}" -f $partPath, $_.Exception.Message)
            continue
        }
        $leaf = Split-Path -Path $partPath -Leaf
        if ($leaf -match '(?i)\.pq$') {
            # Prefer mashup.pq; otherwise the first .pq part (monolith: Select-Object -First 1).
            if ($null -eq $pqPart -or $leaf -ieq 'mashup.pq') { $pqPart = $partPath; $pqBytes = $bytes }
        }
        elseif ($leaf -ieq 'queryMetadata.json') { $metadataBytes = $bytes }
    }
    $result.PartsCount = $written
    if ($null -eq $pqBytes) {
        $result.Message = ('Definition saved ({0} part(s)) but it contains no .pq part' -f $written)
        $result.Success = ($written -gt 0)
        return $result
    }
    try { Write-IQDataflowBinaryFile -Path $Work.BackupPath -Bytes $pqBytes }
    catch {
        $result.Message = 'Could not write ' + $Work.BackupPath + ': ' + $_.Exception.Message
        return $result
    }

    $metadataMap = @{}
    if ($null -ne $metadataBytes) {
        try {
            $metaText = [System.Text.Encoding]::UTF8.GetString($metadataBytes)
            if ($metaText.Length -gt 0 -and $metaText[0] -eq [char]0xFEFF) { $metaText = $metaText.Substring(1) }
            $metaJson = ConvertFrom-Json -InputObject $metaText
            $metadataMap = Get-IQDataflowQueryMetaMap -QueriesMetadata (Get-IQDataflowMember -Object $metaJson -Name 'queriesMetadata') -QueryGroups (Get-IQDataflowMember -Object $metaJson -Name 'queryGroups')
            $result.Metadata = @{
                FormatVersion = [string](Get-IQDataflowMember -Object $metaJson -Name 'formatVersion')
                Name          = [string](Get-IQDataflowMember -Object $metaJson -Name 'name')
                Connections   = @(Get-IQDataflowMember -Object $metaJson -Name 'connections').Count
            }
        }
        catch { Write-IQLog -Level Debug -Stage $stage -Item $Work.Item -Message ('queryMetadata.json could not be parsed: ' + $_.Exception.Message) }
    }
    $pqText = [System.Text.Encoding]::UTF8.GetString($pqBytes)
    if ($pqText.Length -gt 0 -and $pqText[0] -eq [char]0xFEFF) { $pqText = $pqText.Substring(1) }
    $result.Queries = @(ConvertFrom-IQDataflowDocument -Content $pqText -DataflowId $Work.DataflowId -DataflowName $Work.DataflowName -WorkspaceName $Work.WorkspaceName -ReportDate $ReportDate -QueryMetadata $metadataMap -Item $Work.Item)
    $result.Success = $true
    $result.Message = ('{0} part(s) saved, {1} quer{2} parsed from {3}' -f $written, $result.Queries.Count, $(if ($result.Queries.Count -eq 1) { 'y' } else { 'ies' }), $pqPart)
    return $result
}

# =====================================================================================================================
# Work list
# =====================================================================================================================

function Get-IQDataflowLiveList {
    <#
    .SYNOPSIS
    Fallback dataflow list straight from the APIs for one workspace (used only when no inventory files exist) (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$WorkspaceId,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$WorkspaceName
    )
    $stage = 'Dataflows'
    $rows = New-Object System.Collections.Generic.List[object]
    try {
        $response = Invoke-IQApi -Method GET -Path ('groups/' + $WorkspaceId + '/dataflows') -Stage $stage
        foreach ($d in @(Get-IQDataflowMember -Object $response -Name 'value')) {
            if ($null -eq $d) { continue }
            $rows.Add([PSCustomObject]@{ DataflowId = [string]$d.objectId; DataflowName = [string]$d.name; DataflowGeneration = (Get-IQDataflowMember -Object $d -Name 'generation'); WorkspaceId = $WorkspaceId; WorkspaceName = $WorkspaceName })
        }
    }
    catch { Write-IQLog -Level Warn -Stage $stage -Item $WorkspaceName -Message ('Dataflow list failed: ' + $_.Exception.Message) }
    try {
        $fabric = Invoke-IQApi -Method GET -Path ('workspaces/' + $WorkspaceId + '/dataflows') -Api Fabric -AllowNotFound -Stage $stage
        foreach ($d in @(Get-IQDataflowMember -Object $fabric -Name 'value')) {
            if ($null -eq $d) { continue }
            $rows.Add([PSCustomObject]@{ DataflowId = [string]$d.id; DataflowName = [string]$d.displayName; DataflowGeneration = 'Gen 2 CICD'; WorkspaceId = $WorkspaceId; WorkspaceName = $WorkspaceName })
        }
    }
    catch { Write-IQLog -Level Debug -Stage $stage -Item $WorkspaceName -Message ('Fabric dataflow list failed: ' + $_.Exception.Message) }
    return $rows.ToArray()
}

function Get-IQDataflowWorkList {
    <#
    .SYNOPSIS
    One work item per dataflow in the run (from ws-*.json "Dataflows"; live listing when no inventory exists), with backup paths resolved (private).
    .DESCRIPTION
    Pseudo workspaces ("My Workspace", "Shared Reports (No Workspace Access)") are skipped (audit C9-01/C9-08).
    Duplicated ids keep the "Gen 2 CICD" row. File names: "<CleanWs> ~ <CleanDf>" (+ " ~<first 8 of id>" when another
    dataflow already owns that name, C9-10). Returns objects @{ Key; Item; WorkspaceId; WorkspaceName; DataflowId;
    DataflowName; Generation; IsGen2Cicd; Stem; BackupPath; DefinitionFolder; ExtractPath }.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RunFolder,
        [Parameter(Mandatory = $true)][string]$ExtractFolder
    )
    $stage = 'Dataflows'
    $syntheticIds = @('My Workspace', 'Shared Reports (No Workspace Access)')
    $candidates = New-Object System.Collections.Generic.List[object]
    $inventories = @()
    try { $inventories = @(Get-IQAllWorkspaceInventories) } catch { $inventories = @() }
    $inventories = @($inventories | Where-Object { $null -ne $_ })
    if ($inventories.Count -gt 0) {
        foreach ($inv in $inventories) {
            $wsId = [string](Get-IQDataflowMember -Object $inv -Name 'WorkspaceId')
            $wsName = [string](Get-IQDataflowMember -Object $inv -Name 'WorkspaceName')
            if ([bool](Get-IQDataflowMember -Object $inv -Name 'IsSynthetic') -or $syntheticIds -contains $wsId) { continue }
            foreach ($row in @(Get-IQDataflowMember -Object $inv -Name 'Dataflows')) {
                if ($null -eq $row) { continue }
                $rowWs = [string](Get-IQDataflowMember -Object $row -Name 'WorkspaceId')
                if ([string]::IsNullOrWhiteSpace($rowWs)) { $rowWs = $wsId }
                $rowWsName = [string](Get-IQDataflowMember -Object $row -Name 'WorkspaceName')
                if ([string]::IsNullOrWhiteSpace($rowWsName)) { $rowWsName = $wsName }
                $candidates.Add([PSCustomObject]@{
                        DataflowId         = [string](Get-IQDataflowMember -Object $row -Name 'DataflowId')
                        DataflowName       = [string](Get-IQDataflowMember -Object $row -Name 'DataflowName')
                        DataflowGeneration = (Get-IQDataflowMember -Object $row -Name 'DataflowGeneration')
                        WorkspaceId        = $rowWs
                        WorkspaceName      = $rowWsName
                    })
            }
        }
    }
    else {
        Write-IQLog -Level Warn -Stage $stage -Message 'No workspace inventory files found (Inventory stage not run?); listing dataflows directly from the APIs'
        foreach ($ws in @(Get-IQSelectedWorkspaces)) {
            if ($null -eq $ws) { continue }
            $wsId = [string](Get-IQDataflowMember -Object $ws -Name 'WorkspaceId')
            $wsName = [string](Get-IQDataflowMember -Object $ws -Name 'WorkspaceName')
            if ([string]::IsNullOrWhiteSpace($wsId) -or $syntheticIds -contains $wsId -or [bool](Get-IQDataflowMember -Object $ws -Name 'WorkspaceIsSynthetic')) { continue }
            foreach ($row in @(Get-IQDataflowLiveList -WorkspaceId $wsId -WorkspaceName $wsName)) { $candidates.Add($row) }
        }
    }

    # De-duplicate by id (prefer the Gen 2 CICD row), keep a stable order.
    $byId = [ordered]@{}
    foreach ($c in $candidates) {
        $id = [string]$c.DataflowId
        if ([string]::IsNullOrWhiteSpace($id)) { continue }
        $key = $id.ToLowerInvariant()
        $isCicd = ([string]$c.DataflowGeneration -eq 'Gen 2 CICD')
        if ($byId.Contains($key)) {
            $existingIsCicd = ([string]$byId[$key].DataflowGeneration -eq 'Gen 2 CICD')
            if ($isCicd -and -not $existingIsCicd) { $byId[$key] = $c }
            continue
        }
        $byId[$key] = $c
    }

    $usedStems = @{}
    $work = New-Object System.Collections.Generic.List[object]
    foreach ($key in $byId.Keys) {
        $c = $byId[$key]
        $id = [string]$c.DataflowId
        $stem = Get-IQDataflowFileStem -WorkspaceName $c.WorkspaceName -DataflowName $c.DataflowName
        $stemKey = $stem.ToLowerInvariant()
        if ($usedStems.ContainsKey($stemKey) -and $usedStems[$stemKey] -ne $key) {
            $suffix = $id
            if ($suffix.Length -gt 8) { $suffix = $suffix.Substring(0, 8) }
            $stem = $stem + ' ~' + $suffix
            $stemKey = $stem.ToLowerInvariant()
            Write-IQLog -Level Debug -Stage $stage -Message ("Backup name collision for dataflow {0}; using '{1}'" -f $id, $stem)
        }
        $usedStems[$stemKey] = $key
        $isCicd = ([string]$c.DataflowGeneration -eq 'Gen 2 CICD')
        $extension = '.txt'
        if ($isCicd) { $extension = '.pq' }
        $backupPath = Join-Path $RunFolder ($stem + $extension)
        if ($backupPath.Length -gt 240) { Write-IQLog -Level Warn -Stage $stage -Message ("Backup path is {0} characters long and may exceed MAX_PATH: {1}" -f $backupPath.Length, $backupPath) }
        $generationText = [string]$c.DataflowGeneration
        if ([string]::IsNullOrWhiteSpace($generationText)) { $generationText = 'Gen1' }
        $work.Add([PSCustomObject]@{
                Key              = $id
                Item             = ($c.WorkspaceName + ' ~ ' + $c.DataflowName)
                WorkspaceId      = [string]$c.WorkspaceId
                WorkspaceName    = [string]$c.WorkspaceName
                DataflowId       = $id
                DataflowName     = [string]$c.DataflowName
                Generation       = $generationText
                IsGen2Cicd       = $isCicd
                Stem             = $stem
                BackupPath       = $backupPath
                DefinitionFolder = (Join-Path $RunFolder ($stem + '.definition'))
                ExtractPath      = (Join-Path $ExtractFolder ((Get-IQSafeKey -Value $id) + '.json'))
            })
    }
    return $work.ToArray()
}

# =====================================================================================================================
# Extract file (one per dataflow) and the stage body
# =====================================================================================================================

function New-IQDataflowExtract {
    <#
    .SYNOPSIS
    Builds the per-dataflow extract object written to extracts\dataflows\<safeKey>.json (private).
    .DESCRIPTION
    Schema: { SchemaVersion=1; DataflowId; DataflowName; WorkspaceId; WorkspaceName; Generation; Method; Status;
    Message; BackupFile; DefinitionFolder; PartsCount; ReportDate; QueryCount; EntityCount; HasMashup; Metadata{};
    CollectedUtc; Queries[] (Sheet1 rows: "Dataflow ID", "Dataflow Name", "Query Name", "Query", "Report Date",
    "Workspace Name - Dataflow Name", then additive "Load Enabled", "Is Hidden", "Query Group", "Attributes",
    "Source Functions", "Referenced Queries", "Line Count", "Uses Native Query", "Query Length"); Entities[] }.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Work,
        [Parameter(Mandatory = $true)][hashtable]$Result,
        [Parameter(Mandatory = $true)][string]$ReportDate
    )
    $definitionFolder = $null
    if ($Result.ContainsKey('DefinitionFolder') -and $Result.DefinitionFolder -and (Test-Path -LiteralPath $Result.DefinitionFolder)) { $definitionFolder = [string]$Result.DefinitionFolder }
    $partsCount = 0
    if ($Result.ContainsKey('PartsCount')) { $partsCount = [int]$Result.PartsCount }
    $hasMashup = $false
    if ($Result.ContainsKey('HasMashup')) { $hasMashup = [bool]$Result.HasMashup } elseif (@($Result.Queries).Count -gt 0) { $hasMashup = $true }
    $status = 'Failed'
    if ($Result.Success) { $status = 'Succeeded' }
    return [ordered]@{
        SchemaVersion    = 1
        DataflowId       = $Work.DataflowId
        DataflowName     = $Work.DataflowName
        WorkspaceId      = $Work.WorkspaceId
        WorkspaceName    = $Work.WorkspaceName
        Generation       = $Work.Generation
        Method           = [string]$Result.Method
        Status           = $status
        Message          = [string]$Result.Message
        BackupFile       = [string]$Work.BackupPath
        DefinitionFolder = $definitionFolder
        PartsCount       = $partsCount
        ReportDate       = $ReportDate
        QueryCount       = @($Result.Queries).Count
        EntityCount      = @($Result.Entities).Count
        HasMashup        = $hasMashup
        Metadata         = $Result.Metadata
        CollectedUtc     = [datetime]::UtcNow.ToString('o')
        Queries          = @($Result.Queries)
        Entities         = @($Result.Entities)
    }
}

function Get-IQDataflowExtractRow {
    <#
    .SYNOPSIS
    Reads every extracts\dataflows\*.json of the current run and returns the combined "Dataflow Detail.xlsx" Sheet1 rows (or the entity rows with -Entities).
    .DESCRIPTION
    Convenience for the Assemble stage / tests: rows come back in file-name order; the header-only dummy row of the
    monolith is NOT added here (Assemble adds it when the sheet would otherwise be empty). Wrap the call in @( ).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][switch]$Entities)
    $folder = Get-IQDataflowExtractFolder
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($f in @(Get-ChildItem -LiteralPath $folder -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        try {
            $extract = ConvertFrom-IQJsonFile -Path $f.FullName
            if ($null -eq $extract) { continue }
            $name = 'Queries'
            if ($Entities) { $name = 'Entities' }
            foreach ($r in @(Get-IQDataflowMember -Object $extract -Name $name)) { if ($null -ne $r) { $rows.Add($r) } }
        }
        catch { Write-IQLog -Level Warn -Stage 'Dataflows' -Message ("Could not read dataflow extract '{0}': {1}" -f $f.FullName, $_.Exception.Message) }
    }
    return $rows.ToArray()
}

function Invoke-IQDataflowsStage {
    <#
    .SYNOPSIS
    Dataflows stage body: backs up every dataflow in the run (Gen1 JSON / Gen2 CI/CD definition), parses the queries and checkpoints each dataflow (itemKey = DataflowId).
    .DESCRIPTION
    Port of monolith 3298-3735 with the audit C9 fixes (see the file header). Per dataflow: skip when
    Test-IQItemDone says so (resume); otherwise export (Export-IQGen1Dataflow / Export-IQFabricDataflow), write
    extracts\dataflows\<safeKey>.json, then Set-IQItemDone with the backup file and the extract as outputs.
    Failures are recorded (Set-IQItemDone -Status Failed) and the loop continues. Returns
    @{ Total; Done; Failed; AlreadyDone; Gen1; Gen2; Queries; RunFolder; ExtractFolder }.
    #>
    [CmdletBinding()]
    param()
    $stage = 'Dataflows'
    $summary = @{ Total = 0; Done = 0; Failed = 0; AlreadyDone = 0; Gen1 = 0; Gen2 = 0; Queries = 0; RunFolder = $null; ExtractFolder = $null; BudgetStop = $false }
    $runFolder = Get-IQDataflowRunFolder
    $extractFolder = Get-IQDataflowExtractFolder
    $summary.RunFolder = $runFolder
    $summary.ExtractFolder = $extractFolder
    $reportDate = Get-IQDataflowReportDate

    $work = @(Get-IQDataflowWorkList -RunFolder $runFolder -ExtractFolder $extractFolder)
    $summary.Total = $work.Count
    Write-IQLog -Level Info -Stage $stage -Message ("Dataflow backup: {0} dataflow(s) in scope; folder {1}" -f $work.Count, $runFolder)
    if ($work.Count -eq 0) { return $summary }

    $index = 0
    foreach ($w in $work) {
        $index++
        if (Test-IQItemDone -Stage $stage -ItemKey $w.Key) {
            $summary.AlreadyDone++
            Write-IQLog -Level Debug -Stage $stage -Item $w.Item -Message 'Already done (checkpoint); skipping'
            continue
        }
        if (Test-IQTimeBudget -Stage $stage -Item $w.Item) {
            Write-IQLog -Level Warn -Stage $stage -Message ("Time budget reached: {0} of {1} dataflow(s) not backed up yet - they are exported on the next start." -f ($work.Count - $index + 1), $work.Count)
            $summary.BudgetStop = $true
            break
        }
        Write-IQLog -Level Info -Stage $stage -Item $w.Item -Message ("[{0}/{1}] Exporting {2} dataflow" -f $index, $work.Count, $w.Generation)
        $result = $null
        try {
            if ($w.IsGen2Cicd) { $result = Export-IQFabricDataflow -Work $w -ReportDate $reportDate }
            else { $result = Export-IQGen1Dataflow -Work $w -ReportDate $reportDate }
        }
        catch {
            $result = @{ Success = $false; Message = ('Unexpected error: ' + $_.Exception.Message); Queries = @(); Entities = @(); Metadata = @{}; Method = $(if ($w.IsGen2Cicd) { 'Fabric-getDefinition' } else { 'PowerBI-GetDataflow' }) }
            Write-IQLog -Level Debug -Stage $stage -Item $w.Item -Message $_.Exception.ToString()
        }

        $data = @{
            WorkspaceId   = $w.WorkspaceId
            WorkspaceName = $w.WorkspaceName
            DataflowId    = $w.DataflowId
            DataflowName  = $w.DataflowName
            Generation    = $w.Generation
            Method        = [string]$result.Method
            BackupFile    = [string]$w.BackupPath
            QueryCount    = @($result.Queries).Count
            EntityCount   = @($result.Entities).Count
            PartsCount    = $(if ($result.ContainsKey('PartsCount')) { [int]$result.PartsCount } else { 0 })
            ExtractFile   = [string]$w.ExtractPath
        }

        if (-not $result.Success) {
            # Never leave a stale/empty backup behind that could be mistaken for a good one (audit C9-01).
            if ((Test-Path -LiteralPath $w.BackupPath) -and -not (Test-IQDataflowFileHasContent -Path $w.BackupPath)) { Remove-Item -LiteralPath $w.BackupPath -Force -ErrorAction SilentlyContinue }
            Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $w.Item -Status Failed -Method ([string]$result.Method) -Message ([string]$result.Message) -Data $data | Out-Null
            $summary.Failed++
            continue
        }

        try {
            $extract = New-IQDataflowExtract -Work $w -Result $result -ReportDate $reportDate
            ConvertTo-IQJsonFile -Object $extract -Path $w.ExtractPath
        }
        catch {
            Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $w.Item -Status Failed -Method ([string]$result.Method) -Message ('Backup succeeded but the extract could not be written: ' + $_.Exception.Message) -Data $data | Out-Null
            $summary.Failed++
            continue
        }
        $outputs = @($w.ExtractPath)
        if (Test-IQDataflowFileHasContent -Path $w.BackupPath) { $outputs = @($w.BackupPath) + $outputs }
        Set-IQItemDone -Stage $stage -ItemKey $w.Key -Item $w.Item -Outputs $outputs -Method ([string]$result.Method) -Message ([string]$result.Message) -Data $data | Out-Null
        $summary.Done++
        $summary.Queries += @($result.Queries).Count
        if ($w.IsGen2Cicd) { $summary.Gen2++ } else { $summary.Gen1++ }
        Write-IQLog -Level Success -Stage $stage -Item $w.Item -Message ([string]$result.Message)
    }

    $level = 'Info'
    if ($summary.Failed -gt 0) { $level = 'Warn' }
    Write-IQLog -Level $level -Stage $stage -Message ("Dataflow backup finished: {0} Gen1, {1} Gen2 backed up ({2} queries), {3} failed, {4} already done of {5}" -f $summary.Gen1, $summary.Gen2, $summary.Queries, $summary.Failed, $summary.AlreadyDone, $summary.Total)
    return $summary
}
