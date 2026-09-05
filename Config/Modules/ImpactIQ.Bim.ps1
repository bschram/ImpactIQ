# ImpactIQ.Bim.ps1 - TMSL (.bim) parser and Model Detail CSV export without Tabular Editor.
#
# Contract: brief sections 7.2 and 13 (+ task models-revise).
#   ConvertFrom-IQBimModel -Path                                              -> object graph (hashtable) of a TMSL model
#   Export-IQModelDetailFromBim -BimPath -OutputFolder -ModelName -ModelId -AsOfDate
#                                                                              -> writes "<ModelName>.csv" and "<ModelName>_MD.csv"
#                                                                                 with exactly the columns/row kinds of the csx scripts
#
# A .bim file IS TMSL JSON: { name, id, compatibilityLevel, model: { tables[]: { columns[], measures[], partitions[],
# hierarchies[], calculationGroup: { calculationItems[] } }, relationships[], roles[]: { tablePermissions[] } } }. The same
# JSON comes from Tabular Editor (-B export, possibly with expressions split into string arrays), pbi-tools (Pro PBIX) and the
# Fabric semanticModels/{id}/getDefinition?format=TMSL "model.bim" part. Dependency rows come from Get-IQDaxReferences
# (ImpactIQ.Dax.ps1), an approximate regex extractor - see its help for what it does and does not resolve.
#
# Windows PowerShell 5.1 and PowerShell 7 compatible; nothing here is Windows-only. Requires ImpactIQ.Dax.ps1 (CSV building
# blocks: Write-IQCsvFile, New-IQModelDetailRow, Get-IQModelDetailHeader, Get-IQMeasureDependencyHeader,
# ConvertTo-IQMeasureDependencyRows) which ImpactIQ.ps1 loads first; when this file is dot-sourced on its own the Dax module
# next to it is loaded automatically.
#
# Cross-module functions used (brief section 2): Write-IQLog. Private helpers are prefixed *-IQBim* and are not part of the contract.

if (-not (Get-Command -Name 'Write-IQCsvFile' -ErrorAction SilentlyContinue)) {
    $iqBimDaxModule = Join-Path $PSScriptRoot 'ImpactIQ.Dax.ps1'
    if (Test-Path -LiteralPath $iqBimDaxModule) { . $iqBimDaxModule }
}

function Get-IQBimMember {
    <#
    .SYNOPSIS
    Reads a named member from a dictionary (JavaScriptSerializer output) or an object property (ConvertFrom-Json output); $null when absent (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        foreach ($k in $Object.Keys) { if ([string]$k -ieq $Name) { return $Object[$k] } }
        return $null
    }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -ne $prop) { return $prop.Value }
    foreach ($p in $Object.PSObject.Properties) { if ($p.Name -ieq $Name) { return $p.Value } }
    return $null
}

function ConvertFrom-IQBimJson {
    <#
    .SYNOPSIS
    Parses TMSL JSON text; on Windows PowerShell 5.1 large documents (> 2 MB, beyond ConvertFrom-Json's MaxJsonLength) use JavaScriptSerializer (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { throw 'The .bim file is empty.' }
    $body = $Text.TrimStart([char]0xFEFF, ' ', "`t", "`r", "`n")
    if ($PSVersionTable.PSVersion.Major -ge 6) { return ($body | ConvertFrom-Json -ErrorAction Stop) }
    if ($body.Length -lt 2000000) {
        try { return ($body | ConvertFrom-Json -ErrorAction Stop) }
        catch { if ($_.Exception.Message -notmatch 'maxJsonLength|MaxJsonLength') { throw } }
    }
    Add-Type -AssemblyName System.Web.Extensions -ErrorAction Stop
    $serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $serializer.MaxJsonLength = [int]::MaxValue
    $serializer.RecursionLimit = 200
    return $serializer.DeserializeObject($body)
}

function ConvertTo-IQBimText {
    <#
    .SYNOPSIS
    TMSL string-or-string[] value to one string: arrays (Tabular Editor "SplitMultilineStrings") are joined with CRLF like TOM does; $null -> '' (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false, Position = 0)][AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.Collections.IEnumerable]) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($v in $Value) { if ($null -eq $v) { $parts.Add('') } else { $parts.Add([string]$v) } }
        return ($parts.ToArray() -join "`r`n")
    }
    return [string]$Value
}

function ConvertTo-IQBimBool {
    <#
    .SYNOPSIS
    TMSL boolean (absent = default) to [bool] (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)][AllowNull()]$Value,
        [Parameter(Mandatory = $false)][bool]$Default = $false
    )
    if ($null -eq $Value) { return $Default }
    if ($Value -is [bool]) { return $Value }
    $t = ([string]$Value).Trim()
    if ($t -ieq 'true' -or $t -eq '1') { return $true }
    if ($t -ieq 'false' -or $t -eq '0' -or $t -eq '') { return $false }
    try { return [System.Convert]::ToBoolean($Value) } catch { return $Default }
}

function ConvertTo-IQBimEnumName {
    <#
    .SYNOPSIS
    TMSL camelCase enum value (many, oneDirection, directQuery ...) to the TOM PascalCase name the csx scripts write (Many, OneDirection, DirectQuery ...) (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, Position = 0)][AllowNull()]$Value,
        [Parameter(Mandatory = $false)][AllowNull()][hashtable]$Map,
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$Default = ''
    )
    if ($null -eq $Value) { return $Default }
    $text = ([string]$Value).Trim()
    if ($text -eq '') { return $Default }
    if ($null -ne $Map) {
        $lower = $text.ToLowerInvariant()
        if ($Map.ContainsKey($lower)) { return [string]$Map[$lower] }
    }
    return ($text.Substring(0, 1).ToUpperInvariant() + $text.Substring(1))
}

function ConvertFrom-IQBimModel {
    <#
    .SYNOPSIS
    Parses a TMSL .bim file into a plain object graph (hashtables/arrays) with TOM-style names and enum values.
    .DESCRIPTION
    Returns @{ Path; Name; Id; CompatibilityLevel; DefaultMode; Tables; Relationships; Roles; CalculationGroupTables }.
      Table:        @{ Name; Description; IsHidden; StorageMode; IsCalculationGroup; Columns; Measures; Partitions; Hierarchies; CalculationItems }
      Column:       @{ Name; Type (Data|Calculated|CalculatedTableColumn|RowNumber); DataType; SourceColumn; FormatString; DisplayFolder; Description; IsHidden; Expression }
      Measure:      @{ Name; Expression; FormatString; DisplayFolder; Description; IsHidden }
      Partition:    @{ Name; Description; Mode (Import|DirectQuery|Dual|Push|DirectLake|...); SourceType; Expression }
      Hierarchy:    @{ Name; DisplayFolder; Description; IsHidden; Levels = @(@{ Name; Column; Ordinal; Description }) }
      CalculationItem: @{ Name; Expression; Description; Ordinal }
      Relationship: @{ Name; FromTable; FromColumn; ToTable; ToColumn; FromCardinality (Many|One|None); ToCardinality; CrossFilteringBehavior (OneDirection|BothDirections|Automatic); IsActive }
      Role:         @{ Name; ModelPermission; TablePermissions = @(@{ Table; FilterExpression }) }
    TMSL defaults applied: fromCardinality many, toCardinality one, crossFilteringBehavior oneDirection, isActive true,
    partition mode default/absent = model defaultMode (default import). StorageMode of a table = mode of its first
    partition (what the csx writes). Expressions given as string arrays are joined with CRLF. Throws on unreadable JSON.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true, Position = 0)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "The .bim file does not exist: $Path" }
    $text = [System.IO.File]::ReadAllText($Path)
    $root = ConvertFrom-IQBimJson -Text $text
    $model = Get-IQBimMember -Object $root -Name 'model'
    if ($null -eq $model) { throw "Not a TMSL model file (no 'model' object): $Path" }

    $modeMap = @{ 'import' = 'Import'; 'directquery' = 'DirectQuery'; 'dual' = 'Dual'; 'push' = 'Push'; 'directlake' = 'DirectLake'; 'default' = 'Default' }
    $cardinalityMap = @{ 'many' = 'Many'; 'one' = 'One'; 'none' = 'None' }
    $crossFilterMap = @{ 'onedirection' = 'OneDirection'; 'bothdirections' = 'BothDirections'; 'automatic' = 'Automatic' }
    $columnTypeMap = @{ 'data' = 'Data'; 'calculated' = 'Calculated'; 'calculatedtablecolumn' = 'CalculatedTableColumn'; 'rownumber' = 'RowNumber' }

    $defaultMode = ConvertTo-IQBimEnumName -Value (Get-IQBimMember -Object $model -Name 'defaultMode') -Map $modeMap -Default 'Import'
    if ($defaultMode -eq 'Default' -or $defaultMode -eq '') { $defaultMode = 'Import' }
    $resolveMode = {
        param($value)
        $m = ConvertTo-IQBimEnumName -Value $value -Map $modeMap -Default $defaultMode
        if ($m -eq 'Default') { return $defaultMode }
        return $m
    }

    $result = @{
        Path = $Path; Name = [string](Get-IQBimMember -Object $root -Name 'name'); Id = [string](Get-IQBimMember -Object $root -Name 'id')
        CompatibilityLevel = Get-IQBimMember -Object $root -Name 'compatibilityLevel'; DefaultMode = $defaultMode
        Tables = @(); Relationships = @(); Roles = @(); CalculationGroupTables = @()
    }
    $tables = New-Object System.Collections.Generic.List[object]
    $calcGroupNames = New-Object System.Collections.Generic.List[string]
    foreach ($t in @(Get-IQBimMember -Object $model -Name 'tables')) {
        if ($null -eq $t) { continue }
        $tableName = [string](Get-IQBimMember -Object $t -Name 'name')
        $columns = New-Object System.Collections.Generic.List[object]
        foreach ($c in @(Get-IQBimMember -Object $t -Name 'columns')) {
            if ($null -eq $c) { continue }
            $columns.Add(@{
                    Name = [string](Get-IQBimMember -Object $c -Name 'name')
                    Type = (ConvertTo-IQBimEnumName -Value (Get-IQBimMember -Object $c -Name 'type') -Map $columnTypeMap -Default 'Data')
                    DataType = [string](Get-IQBimMember -Object $c -Name 'dataType'); SourceColumn = [string](Get-IQBimMember -Object $c -Name 'sourceColumn')
                    FormatString = (ConvertTo-IQBimText (Get-IQBimMember -Object $c -Name 'formatString')); DisplayFolder = [string](Get-IQBimMember -Object $c -Name 'displayFolder')
                    Description = (ConvertTo-IQBimText (Get-IQBimMember -Object $c -Name 'description')); IsHidden = (ConvertTo-IQBimBool (Get-IQBimMember -Object $c -Name 'isHidden'))
                    Expression = (ConvertTo-IQBimText (Get-IQBimMember -Object $c -Name 'expression'))
                })
        }
        $measures = New-Object System.Collections.Generic.List[object]
        foreach ($m in @(Get-IQBimMember -Object $t -Name 'measures')) {
            if ($null -eq $m) { continue }
            $measures.Add(@{
                    Name = [string](Get-IQBimMember -Object $m -Name 'name'); Expression = (ConvertTo-IQBimText (Get-IQBimMember -Object $m -Name 'expression'))
                    FormatString = (ConvertTo-IQBimText (Get-IQBimMember -Object $m -Name 'formatString')); DisplayFolder = [string](Get-IQBimMember -Object $m -Name 'displayFolder')
                    Description = (ConvertTo-IQBimText (Get-IQBimMember -Object $m -Name 'description')); IsHidden = (ConvertTo-IQBimBool (Get-IQBimMember -Object $m -Name 'isHidden'))
                })
        }
        $partitions = New-Object System.Collections.Generic.List[object]
        foreach ($p in @(Get-IQBimMember -Object $t -Name 'partitions')) {
            if ($null -eq $p) { continue }
            $source = Get-IQBimMember -Object $p -Name 'source'
            $sourceType = ''
            $expr = ''
            if ($null -ne $source) {
                $sourceType = ConvertTo-IQBimEnumName -Value (Get-IQBimMember -Object $source -Name 'type') -Default ''
                $expr = ConvertTo-IQBimText (Get-IQBimMember -Object $source -Name 'expression')
                if ($expr -eq '') { $expr = ConvertTo-IQBimText (Get-IQBimMember -Object $source -Name 'query') }
                if ($expr -eq '') { $expr = ConvertTo-IQBimText (Get-IQBimMember -Object $source -Name 'entityName') }
            }
            $partitions.Add(@{
                    Name = [string](Get-IQBimMember -Object $p -Name 'name'); Description = (ConvertTo-IQBimText (Get-IQBimMember -Object $p -Name 'description'))
                    Mode = (& $resolveMode (Get-IQBimMember -Object $p -Name 'mode')); SourceType = $sourceType; Expression = $expr
                })
        }
        $hierarchies = New-Object System.Collections.Generic.List[object]
        foreach ($h in @(Get-IQBimMember -Object $t -Name 'hierarchies')) {
            if ($null -eq $h) { continue }
            $levels = New-Object System.Collections.Generic.List[object]
            foreach ($l in @(Get-IQBimMember -Object $h -Name 'levels')) {
                if ($null -eq $l) { continue }
                $levels.Add(@{
                        Name = [string](Get-IQBimMember -Object $l -Name 'name'); Column = [string](Get-IQBimMember -Object $l -Name 'column')
                        Ordinal = (Get-IQBimMember -Object $l -Name 'ordinal'); Description = (ConvertTo-IQBimText (Get-IQBimMember -Object $l -Name 'description'))
                    })
            }
            $hierarchies.Add(@{
                    Name = [string](Get-IQBimMember -Object $h -Name 'name'); DisplayFolder = [string](Get-IQBimMember -Object $h -Name 'displayFolder')
                    Description = (ConvertTo-IQBimText (Get-IQBimMember -Object $h -Name 'description')); IsHidden = (ConvertTo-IQBimBool (Get-IQBimMember -Object $h -Name 'isHidden'))
                    Levels = $levels.ToArray()
                })
        }
        $calcGroup = Get-IQBimMember -Object $t -Name 'calculationGroup'
        $items = New-Object System.Collections.Generic.List[object]
        $isCalcGroup = ($null -ne $calcGroup)
        if ($isCalcGroup) {
            $calcGroupNames.Add($tableName)
            $ordered = @(Get-IQBimMember -Object $calcGroup -Name 'calculationItems') | Where-Object { $null -ne $_ } | Sort-Object { $o = Get-IQBimMember -Object $_ -Name 'ordinal'; if ($null -eq $o) { 0 } else { [double]$o } }
            foreach ($ci in @($ordered)) {
                $items.Add(@{
                        Name = [string](Get-IQBimMember -Object $ci -Name 'name'); Expression = (ConvertTo-IQBimText (Get-IQBimMember -Object $ci -Name 'expression'))
                        Description = (ConvertTo-IQBimText (Get-IQBimMember -Object $ci -Name 'description')); Ordinal = (Get-IQBimMember -Object $ci -Name 'ordinal')
                    })
            }
        }
        $storageMode = ''
        if ($partitions.Count -gt 0) { $storageMode = [string]$partitions[0].Mode }
        $tables.Add(@{
                Name = $tableName; Description = (ConvertTo-IQBimText (Get-IQBimMember -Object $t -Name 'description'))
                IsHidden = (ConvertTo-IQBimBool (Get-IQBimMember -Object $t -Name 'isHidden')); StorageMode = $storageMode
                IsCalculationGroup = $isCalcGroup; CalculationGroupDescription = (ConvertTo-IQBimText (Get-IQBimMember -Object $calcGroup -Name 'description'))
                Columns = $columns.ToArray(); Measures = $measures.ToArray(); Partitions = $partitions.ToArray(); Hierarchies = $hierarchies.ToArray(); CalculationItems = $items.ToArray()
            })
    }
    $relationships = New-Object System.Collections.Generic.List[object]
    foreach ($r in @(Get-IQBimMember -Object $model -Name 'relationships')) {
        if ($null -eq $r) { continue }
        $relationships.Add(@{
                Name = [string](Get-IQBimMember -Object $r -Name 'name')
                FromTable = [string](Get-IQBimMember -Object $r -Name 'fromTable'); FromColumn = [string](Get-IQBimMember -Object $r -Name 'fromColumn')
                ToTable = [string](Get-IQBimMember -Object $r -Name 'toTable'); ToColumn = [string](Get-IQBimMember -Object $r -Name 'toColumn')
                FromCardinality = (ConvertTo-IQBimEnumName -Value (Get-IQBimMember -Object $r -Name 'fromCardinality') -Map $cardinalityMap -Default 'Many')
                ToCardinality = (ConvertTo-IQBimEnumName -Value (Get-IQBimMember -Object $r -Name 'toCardinality') -Map $cardinalityMap -Default 'One')
                CrossFilteringBehavior = (ConvertTo-IQBimEnumName -Value (Get-IQBimMember -Object $r -Name 'crossFilteringBehavior') -Map $crossFilterMap -Default 'OneDirection')
                IsActive = (ConvertTo-IQBimBool (Get-IQBimMember -Object $r -Name 'isActive') -Default $true)
            })
    }
    $roles = New-Object System.Collections.Generic.List[object]
    foreach ($role in @(Get-IQBimMember -Object $model -Name 'roles')) {
        if ($null -eq $role) { continue }
        $permissions = New-Object System.Collections.Generic.List[object]
        foreach ($tp in @(Get-IQBimMember -Object $role -Name 'tablePermissions')) {
            if ($null -eq $tp) { continue }
            $permissions.Add(@{ Table = [string](Get-IQBimMember -Object $tp -Name 'name'); FilterExpression = (ConvertTo-IQBimText (Get-IQBimMember -Object $tp -Name 'filterExpression')) })
        }
        $roles.Add(@{ Name = [string](Get-IQBimMember -Object $role -Name 'name'); ModelPermission = (ConvertTo-IQBimEnumName -Value (Get-IQBimMember -Object $role -Name 'modelPermission') -Default ''); TablePermissions = $permissions.ToArray() })
    }
    $result.Tables = $tables.ToArray()
    $result.Relationships = $relationships.ToArray()
    $result.Roles = $roles.ToArray()
    $result.CalculationGroupTables = $calcGroupNames.ToArray()
    return $result
}

function Get-IQBimModelDetailRows {
    <#
    .SYNOPSIS
    Builds the "<model>.csv" rows (brief section 13, csx emission order) from a parsed model (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Model,
        [Parameter(Mandatory = $true)][hashtable]$Common
    )
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($t in $Model.Tables) {
        $rows.Add((New-IQModelDetailRow -Type 'Table' -Common $Common -Fields @{ Table = $t.Name; Name = $t.Name; IsHidden = $t.IsHidden; TableStorageMode = $t.StorageMode; Description = $t.Description }))
    }
    foreach ($t in $Model.Tables) {
        if (-not $t.IsCalculationGroup) { continue }
        $desc = $t.Description
        if ([string]::IsNullOrEmpty($desc)) { $desc = $t.CalculationGroupDescription }
        $rows.Add((New-IQModelDetailRow -Type 'CalculationGroup' -Common $Common -Fields @{ Table = $t.Name; Name = $t.Name; Description = $desc; IsHidden = $t.IsHidden }))
        foreach ($ci in $t.CalculationItems) {
            $rows.Add((New-IQModelDetailRow -Type 'CalculationItem' -Common $Common -Fields @{ Table = $t.Name; Name = $ci.Name; Description = $ci.Description; Expression = $ci.Expression }))
        }
    }
    foreach ($t in $Model.Tables) {
        foreach ($c in $t.Columns) {
            if ($c.Type -eq 'RowNumber') { continue }
            $rows.Add((New-IQModelDetailRow -Type 'Column' -Common $Common -Fields @{ Table = $t.Name; Name = $c.Name; FormatString = $c.FormatString; DisplayFolder = $c.DisplayFolder; Description = $c.Description; IsHidden = $c.IsHidden }))
        }
    }
    foreach ($t in $Model.Tables) {
        foreach ($c in $t.Columns) {
            if ($c.Type -ne 'Calculated') { continue }
            $rows.Add((New-IQModelDetailRow -Type 'CalculatedColumn' -Common $Common -Fields @{ Table = $t.Name; Name = $c.Name; FormatString = $c.FormatString; DisplayFolder = $c.DisplayFolder; Description = $c.Description; IsHidden = $c.IsHidden; Expression = $c.Expression }))
        }
    }
    foreach ($t in $Model.Tables) {
        foreach ($m in $t.Measures) {
            $rows.Add((New-IQModelDetailRow -Type 'Measure' -Common $Common -Fields @{ Table = $t.Name; Name = $m.Name; FormatString = $m.FormatString; DisplayFolder = $m.DisplayFolder; Description = $m.Description; IsHidden = $m.IsHidden; Expression = $m.Expression }))
        }
    }
    foreach ($t in $Model.Tables) {
        foreach ($h in $t.Hierarchies) {
            $rows.Add((New-IQModelDetailRow -Type 'Hierarchy' -Common $Common -Fields @{ Table = $t.Name; Name = $h.Name; DisplayFolder = $h.DisplayFolder; Description = $h.Description; IsHidden = $h.IsHidden }))
        }
    }
    foreach ($t in $Model.Tables) {
        foreach ($h in $t.Hierarchies) {
            foreach ($l in $h.Levels) {
                $rows.Add((New-IQModelDetailRow -Type 'Level' -Common $Common -Fields @{ Table = $t.Name; Name = $l.Name; Description = $l.Description }))
            }
        }
    }
    foreach ($t in $Model.Tables) {
        foreach ($p in $t.Partitions) {
            $rows.Add((New-IQModelDetailRow -Type 'Partition' -Common $Common -Fields @{ Table = $t.Name; Name = $p.Name; Description = $p.Description; TableStorageMode = $p.Mode; Expression = $p.Expression }))
        }
    }
    foreach ($role in $Model.Roles) {
        foreach ($tp in $role.TablePermissions) {
            $rows.Add((New-IQModelDetailRow -Type 'RLSFilter' -Common $Common -Fields @{ Table = $tp.Table; Name = $role.Name; Expression = $tp.FilterExpression }))
        }
    }
    foreach ($r in $Model.Relationships) {
        $status = 'False'
        if ($r.IsActive) { $status = 'True' }
        $rows.Add((New-IQModelDetailRow -Type 'Relationship' -Common $Common -Fields @{
                    Table = $r.FromTable; Name = $r.FromColumn; Expression = $r.Name
                    RelationshipFromTable = $r.FromTable; RelationshipFromColumn = $r.FromColumn; RelationshipToTable = $r.ToTable; RelationshipToColumn = $r.ToColumn
                    RelationshipStatus = $status; RelationshipFromCardinality = $r.FromCardinality; RelationshipToCardinality = $r.ToCardinality
                    RelationshipCrossFilteringBehavior = $r.CrossFilteringBehavior
                }))
    }
    return $rows.ToArray()
}

function Get-IQBimDependencyRows {
    <#
    .SYNOPSIS
    Builds the "<model>_MD.csv" rows (measures, then calculated columns, then calculation items) from a parsed model via Get-IQDaxReferences (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][hashtable]$Model,
        [Parameter(Mandatory = $true)][hashtable]$Common
    )
    $tableNames = @()
    $measureNames = @()
    $knownColumns = @{}
    foreach ($t in $Model.Tables) {
        $tableNames += $t.Name
        $knownColumns[$t.Name] = @($t.Columns | Where-Object { $_.Type -ne 'RowNumber' } | ForEach-Object { $_.Name })
        foreach ($m in $t.Measures) { $measureNames += $m.Name }
    }
    $objects = New-Object System.Collections.Generic.List[object]
    foreach ($t in $Model.Tables) { foreach ($m in $t.Measures) { $objects.Add(@{ ObjectName = $m.Name; ObjectType = 'Measure'; Table = $t.Name; Expression = $m.Expression }) } }
    foreach ($t in $Model.Tables) { foreach ($c in $t.Columns) { if ($c.Type -eq 'Calculated') { $objects.Add(@{ ObjectName = $c.Name; ObjectType = 'CalculatedColumn'; Table = $t.Name; Expression = $c.Expression }) } } }
    foreach ($t in $Model.Tables) { if ($t.IsCalculationGroup) { foreach ($ci in $t.CalculationItems) { $objects.Add(@{ ObjectName = $ci.Name; ObjectType = 'CalculationItem'; Table = $t.Name; Expression = $ci.Expression }) } } }
    return @(ConvertTo-IQMeasureDependencyRows -Objects $objects.ToArray() -Common $Common -KnownTables $tableNames -KnownMeasures $measureNames -KnownColumns $knownColumns -CalculationGroupTables @($Model.CalculationGroupTables))
}

function Export-IQModelDetailFromBim {
    <#
    .SYNOPSIS
    Writes "<ModelName>.csv" (Semantic Models rows) and "<ModelName>_MD.csv" (Measure Dependencies rows) from a TMSL .bim file.
    .DESCRIPTION
    Same files, header, quoting (every field quoted, " doubled, CRLF, UTF-8 with BOM) and row kinds as the Tabular Editor
    csx scripts (brief section 13): Table, CalculationGroup, CalculationItem, Column, CalculatedColumn (again, with
    Expression), Measure, Hierarchy, Level, Partition, RLSFilter, Relationship. ModelName should be the
    "<CleanWs> ~ <CleanModel>" base name (the PBIT splits it on "~"); ModelId the dataset GUID for dedicated-capacity
    models and the same base name for Pro models (the PBIT join rule, audit x1 section 4.1); AsOfDate the run date
    (yyyy-MM-dd). Dependency rows are approximate (Get-IQDaxReferences). Never throws; returns
    @{ Success; Csv; MdCsv; Outputs; Message; RowCount; DependencyRowCount; Method='Bim'; ModelName; BimName; BimId }.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$BimPath,
        [Parameter(Mandatory = $true)][string]$OutputFolder,
        [Parameter(Mandatory = $true)][string]$ModelName,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$ModelId,
        [Parameter(Mandatory = $true)][string]$AsOfDate,
        [Parameter(Mandatory = $false)][string]$Stage = 'ModelDetail'
    )
    $result = @{ Success = $false; Csv = $null; MdCsv = $null; Outputs = @(); Message = ''; RowCount = 0; DependencyRowCount = 0; Method = 'Bim'; ModelName = $ModelName; BimName = ''; BimId = '' }
    $model = $null
    try { $model = ConvertFrom-IQBimModel -Path $BimPath }
    catch {
        $result.Message = 'Could not parse the .bim file: ' + $_.Exception.Message
        Write-IQLog -Level Warn -Stage $Stage -Item $ModelName -Message $result.Message
        return $result
    }
    $result.BimName = [string]$model.Name
    $result.BimId = [string]$model.Id
    $common = @{ ModelAsOfDate = $AsOfDate; ModelName = $ModelName; ModelID = $ModelId }
    try {
        if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
        $csvPath = Join-Path $OutputFolder ($ModelName + '.csv')
        $mdPath = Join-Path $OutputFolder ($ModelName + '_MD.csv')
        $rows = @(Get-IQBimModelDetailRows -Model $model -Common $common)
        $mdRows = @(Get-IQBimDependencyRows -Model $model -Common $common)
        $result.RowCount = Write-IQCsvFile -Path $csvPath -Header (Get-IQModelDetailHeader) -Rows $rows
        $result.DependencyRowCount = Write-IQCsvFile -Path $mdPath -Header (Get-IQMeasureDependencyHeader) -Rows $mdRows
    }
    catch {
        $result.Message = 'Could not build the model detail CSVs from the .bim: ' + $_.Exception.Message
        Write-IQLog -Level Error -Stage $Stage -Item $ModelName -Message $result.Message -Exception $_.Exception
        return $result
    }
    $result.Success = $true
    $result.Csv = $csvPath
    $result.MdCsv = $mdPath
    $result.Outputs = @($csvPath, $mdPath)
    $result.Message = ('{0} object rows, {1} dependency rows parsed from {2} (dependencies: expression parsing, approximate)' -f $result.RowCount, $result.DependencyRowCount, (Split-Path -Path $BimPath -Leaf))
    Write-IQLog -Level Success -Stage $Stage -Item $ModelName -Message $result.Message
    return $result
}
