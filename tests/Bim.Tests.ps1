# Pester 5 tests for ImpactIQ.Bim.ps1 (TMSL parser + Model Detail CSV export) and the DAX reference extractor it relies on.
# Runs on Linux pwsh: only Common, Dax and Bim are loaded; no context, no network.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester Should operators are registered dynamically; the compatibility profiles do not know Pester 5')]
param()

BeforeAll {
    $repo = Split-Path -Parent $PSScriptRoot
    $script:IQ = $null
    . (Join-Path $repo 'Config/Modules/ImpactIQ.Common.ps1')
    . (Join-Path $repo 'Config/Modules/ImpactIQ.Dax.ps1')
    . (Join-Path $repo 'Config/Modules/ImpactIQ.Bim.ps1')
    $script:FixturePath = Join-Path $PSScriptRoot 'fixtures/bim/sample-model.bim'
    $script:OutFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('iq-bim-tests-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Path $script:OutFolder -Force | Out-Null
    $script:ModelName = 'Finance (Prod) ~ Ledger'
    $script:ModelId = '11111111-1111-1111-1111-111111111111'
    $script:Result = Export-IQModelDetailFromBim -BimPath $script:FixturePath -OutputFolder $script:OutFolder -ModelName $script:ModelName -ModelId $script:ModelId -AsOfDate '2026-09-05'
    $script:CsvLines = @([System.IO.File]::ReadAllText($script:Result.Csv) -split "`r`n")
    $script:MdLines = @([System.IO.File]::ReadAllText($script:Result.MdCsv) -split "`r`n")
    $script:Rows = @(Import-Csv -Path $script:Result.Csv -Encoding UTF8)
    $script:MdRows = @(Import-Csv -Path $script:Result.MdCsv -Encoding UTF8)
}

AfterAll {
    if ($script:OutFolder -and (Test-Path -LiteralPath $script:OutFolder)) { Remove-Item -LiteralPath $script:OutFolder -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'ConvertFrom-IQBimModel' {
    BeforeAll { $script:Model = ConvertFrom-IQBimModel -Path $script:FixturePath }

    It 'reads the database name, id and tables' {
        $script:Model.Name | Should -Be 'Finance (Prod) ~ Ledger'
        $script:Model.Id | Should -Be '11111111-1111-1111-1111-111111111111'
        @($script:Model.Tables).Count | Should -Be 3
        @($script:Model.Tables | ForEach-Object { $_.Name }) | Should -Be @('Sales', 'Date', 'Time Intelligence')
    }
    It 'identifies the calculation group table and its items in ordinal order' {
        $script:Model.CalculationGroupTables | Should -Be @('Time Intelligence')
        $cg = $script:Model.Tables | Where-Object { $_.Name -eq 'Time Intelligence' }
        $cg.IsCalculationGroup | Should -BeTrue
        @($cg.CalculationItems | ForEach-Object { $_.Name }) | Should -Be @('Current', 'YTD')
    }
    It 'joins string-array expressions with CRLF and keeps plain strings' {
        $sales = $script:Model.Tables | Where-Object { $_.Name -eq 'Sales' }
        ($sales.Measures | Where-Object { $_.Name -eq 'Total Cost' }).Expression | Should -Be "`r`nSUM ( Sales[Cost] )"
        ($sales.Measures | Where-Object { $_.Name -eq 'Total Sales' }).Expression | Should -Be 'SUM(Sales[Amount])'
        ($sales.Columns | Where-Object { $_.Name -eq 'Amount x2' }).Type | Should -Be 'Calculated'
        $sales.Partitions[0].Expression | Should -Match '^let\r\n    Source = Sql\.Database\("srv", "db"\),'
    }
    It 'applies TMSL defaults: partition mode Import, cardinality Many/One, OneDirection, active' {
        ($script:Model.Tables | Where-Object { $_.Name -eq 'Date' }).StorageMode | Should -Be 'Import'
        $r1 = $script:Model.Relationships | Where-Object { $_.FromColumn -eq 'DateKey' }
        $r1.FromCardinality | Should -Be 'Many'
        $r1.ToCardinality | Should -Be 'One'
        $r1.CrossFilteringBehavior | Should -Be 'OneDirection'
        $r1.IsActive | Should -BeTrue
        $r2 = $script:Model.Relationships | Where-Object { $_.FromColumn -eq 'ShipDateKey' }
        $r2.CrossFilteringBehavior | Should -Be 'BothDirections'
        $r2.IsActive | Should -BeFalse
    }
    It 'reads roles and table permissions' {
        @($script:Model.Roles).Count | Should -Be 1
        $script:Model.Roles[0].Name | Should -Be 'Region Manager'
        $script:Model.Roles[0].ModelPermission | Should -Be 'Read'
        $script:Model.Roles[0].TablePermissions[0].Table | Should -Be 'Sales'
        $script:Model.Roles[0].TablePermissions[0].FilterExpression | Should -Be '[Region] = USERNAME()'
    }
    It 'throws on a file that is not TMSL' {
        $bad = Join-Path $script:OutFolder 'bad.bim'
        '{"name":"x"}' | Set-Content -Path $bad -Encoding UTF8
        { ConvertFrom-IQBimModel -Path $bad } | Should -Throw
    }
}

Describe 'Export-IQModelDetailFromBim' {
    It 'succeeds and reports the row counts' {
        $script:Result.Success | Should -BeTrue
        $script:Result.Method | Should -Be 'Bim'
        $script:Result.RowCount | Should -Be 31
        $script:Result.DependencyRowCount | Should -Be 6
        @($script:Result.Outputs).Count | Should -Be 2
        (Split-Path -Leaf $script:Result.Csv) | Should -Be 'Finance (Prod) ~ Ledger.csv'
        (Split-Path -Leaf $script:Result.MdCsv) | Should -Be 'Finance (Prod) ~ Ledger_MD.csv'
    }
    It 'writes the exact 20-column Semantic Models header (brief section 13)' {
        $script:CsvLines[0] | Should -Be '"Type","Table","Name","FormatString","DisplayFolder","Description","IsHidden","TableStorageMode","Expression","ModelAsOfDate","ModelName","ModelID","RelationshipFromTable","RelationshipFromColumn","RelationshipToTable","RelationshipToColumn","RelationshipStatus","RelationshipFromCardinality","RelationshipToCardinality","RelationshipCrossFilteringBehavior"'
        (Get-IQModelDetailHeader) -join ',' | Should -Be 'Type,Table,Name,FormatString,DisplayFolder,Description,IsHidden,TableStorageMode,Expression,ModelAsOfDate,ModelName,ModelID,RelationshipFromTable,RelationshipFromColumn,RelationshipToTable,RelationshipToColumn,RelationshipStatus,RelationshipFromCardinality,RelationshipToCardinality,RelationshipCrossFilteringBehavior'
    }
    It 'writes the exact 7-column Measure Dependencies header' {
        $script:MdLines[0] | Should -Be '"ObjectName","ObjectType","DependsOn","DependsOnType","ModelAsOfDate","ModelName","ModelID"'
    }
    It 'writes UTF-8 with BOM and CRLF' {
        $bytes = [System.IO.File]::ReadAllBytes($script:Result.Csv)
        $bytes[0] | Should -Be 0xEF
        $bytes[1] | Should -Be 0xBB
        $bytes[2] | Should -Be 0xBF
        ([System.IO.File]::ReadAllText($script:Result.Csv)) | Should -Match "`r`n"
    }
    It 'emits every row kind with the expected counts' {
        $counts = @{}
        foreach ($g in ($script:Rows | Group-Object Type)) { $counts[$g.Name] = $g.Count }
        $counts['Table'] | Should -Be 3
        $counts['CalculationGroup'] | Should -Be 1
        $counts['CalculationItem'] | Should -Be 2
        $counts['Column'] | Should -Be 12
        $counts['CalculatedColumn'] | Should -Be 1
        $counts['Measure'] | Should -Be 3
        $counts['Hierarchy'] | Should -Be 1
        $counts['Level'] | Should -Be 2
        $counts['Partition'] | Should -Be 3
        $counts['RLSFilter'] | Should -Be 1
        $counts['Relationship'] | Should -Be 2
        $script:Rows.Count | Should -Be 31
    }
    It 'emits the row kinds in the csx order' {
        $order = @($script:Rows | ForEach-Object { $_.Type } | Select-Object -Unique)
        $order | Should -Be @('Table', 'CalculationGroup', 'CalculationItem', 'Column', 'CalculatedColumn', 'Measure', 'Hierarchy', 'Level', 'Partition', 'RLSFilter', 'Relationship')
    }
    It 'stamps ModelAsOfDate, ModelName and ModelID on every row' {
        @($script:Rows | Where-Object { $_.ModelAsOfDate -ne '2026-09-05' -or $_.ModelName -ne $script:ModelName -or $_.ModelID -ne $script:ModelId }).Count | Should -Be 0
        @($script:MdRows | Where-Object { $_.ModelAsOfDate -ne '2026-09-05' -or $_.ModelName -ne $script:ModelName -or $_.ModelID -ne $script:ModelId }).Count | Should -Be 0
    }
    It 'fills the Table rows (storage mode from the first partition, hidden flag as True/False)' {
        $sales = $script:Rows | Where-Object { $_.Type -eq 'Table' -and $_.Name -eq 'Sales' }
        $sales.Table | Should -Be 'Sales'
        $sales.TableStorageMode | Should -Be 'Import'
        $sales.IsHidden | Should -Be 'False'
        $sales.Expression | Should -Be ''
        ($script:Rows | Where-Object { $_.Type -eq 'Table' -and $_.Name -eq 'Date' }).TableStorageMode | Should -Be 'Import'
    }
    It 'fills Column and CalculatedColumn rows like the csx (calculated column listed twice, expression only on the second)' {
        $amount = $script:Rows | Where-Object { $_.Type -eq 'Column' -and $_.Name -eq 'Amount' }
        $amount.Table | Should -Be 'Sales'
        $amount.FormatString | Should -Be '#,0.00'
        $amount.DisplayFolder | Should -Be 'Money'
        $amount.Description | Should -Be 'Sales "amount" in USD'
        $amount.IsHidden | Should -Be 'False'
        $amount.Expression | Should -Be ''
        $ccAsColumn = $script:Rows | Where-Object { $_.Type -eq 'Column' -and $_.Name -eq 'Amount x2' }
        $ccAsColumn.Expression | Should -Be ''
        $cc = $script:Rows | Where-Object { $_.Type -eq 'CalculatedColumn' }
        $cc.Name | Should -Be 'Amount x2'
        $cc.Expression | Should -Be "`r`nSales[Amount] * 2 // doubled for the demo"
        $cc.FormatString | Should -Be '0'
        ($script:Rows | Where-Object { $_.Type -eq 'Column' -and $_.Name -eq 'DateKey' }).IsHidden | Should -Be 'True'
    }
    It 'fills Measure rows' {
        $m = $script:Rows | Where-Object { $_.Type -eq 'Measure' -and $_.Name -eq 'Total Sales' }
        $m.Table | Should -Be 'Sales'
        $m.Expression | Should -Be 'SUM(Sales[Amount])'
        $m.FormatString | Should -Be '#,0.00'
        $m.DisplayFolder | Should -Be 'KPIs'
        $m.Description | Should -Be 'Sum of amount'
        $m.IsHidden | Should -Be 'False'
        ($script:Rows | Where-Object { $_.Type -eq 'Measure' -and $_.Name -eq 'Margin' }).Expression | Should -Match 'VAR sales = \[Total Sales\]'
    }
    It 'fills CalculationGroup / CalculationItem rows' {
        $cg = $script:Rows | Where-Object { $_.Type -eq 'CalculationGroup' }
        $cg.Table | Should -Be 'Time Intelligence'
        $cg.Name | Should -Be 'Time Intelligence'
        $cg.Description | Should -Be 'Time calculation group'
        $cg.IsHidden | Should -Be 'False'
        $ytd = $script:Rows | Where-Object { $_.Type -eq 'CalculationItem' -and $_.Name -eq 'YTD' }
        $ytd.Table | Should -Be 'Time Intelligence'
        $ytd.Expression | Should -Be "CALCULATE ( SELECTEDMEASURE (), DATESYTD ( 'Date'[Date] ) )"
        ($script:Rows | Where-Object { $_.Type -eq 'CalculationItem' -and $_.Name -eq 'Current' }).Description | Should -Be 'As is'
    }
    It 'fills Hierarchy and Level rows (Level.Table = the hierarchy table)' {
        $h = $script:Rows | Where-Object { $_.Type -eq 'Hierarchy' }
        $h.Table | Should -Be 'Date'
        $h.Name | Should -Be 'Calendar'
        $h.DisplayFolder | Should -Be 'Hierarchies'
        $h.Description | Should -Be 'Year > Month'
        $levels = @($script:Rows | Where-Object { $_.Type -eq 'Level' })
        @($levels | ForEach-Object { $_.Name }) | Should -Be @('Year', 'Month')
        $levels[1].Table | Should -Be 'Date'
        $levels[1].Description | Should -Be 'Month name'
    }
    It 'fills Partition rows with the M expression verbatim and the mode' {
        $p = $script:Rows | Where-Object { $_.Type -eq 'Partition' -and $_.Table -eq 'Sales' }
        $p.Name | Should -Be 'Sales'
        $p.TableStorageMode | Should -Be 'Import'
        $p.Expression | Should -Match 'Sql\.Database\("srv", "db"\)'
        $p.Expression | Should -Match 'Source\{\[Schema="dbo",Item="Sales"\]\}\[Data\]'
        ($script:Rows | Where-Object { $_.Type -eq 'Partition' -and $_.Table -eq 'Time Intelligence' }).Expression | Should -Be ''
    }
    It 'fills the RLSFilter row (Table = permission table, Name = role)' {
        $rls = $script:Rows | Where-Object { $_.Type -eq 'RLSFilter' }
        $rls.Table | Should -Be 'Sales'
        $rls.Name | Should -Be 'Region Manager'
        $rls.Expression | Should -Be '[Region] = USERNAME()'
    }
    It 'fills Relationship rows with TOM enum names and defaults' {
        $r1 = $script:Rows | Where-Object { $_.Type -eq 'Relationship' -and $_.Name -eq 'DateKey' }
        $r1.Table | Should -Be 'Sales'
        $r1.Expression | Should -Be '8a6d4c8e-2222-4b2c-9d3e-000000000001'
        $r1.RelationshipFromTable | Should -Be 'Sales'
        $r1.RelationshipFromColumn | Should -Be 'DateKey'
        $r1.RelationshipToTable | Should -Be 'Date'
        $r1.RelationshipToColumn | Should -Be 'Date'
        $r1.RelationshipStatus | Should -Be 'True'
        $r1.RelationshipFromCardinality | Should -Be 'Many'
        $r1.RelationshipToCardinality | Should -Be 'One'
        $r1.RelationshipCrossFilteringBehavior | Should -Be 'OneDirection'
        $r2 = $script:Rows | Where-Object { $_.Type -eq 'Relationship' -and $_.Name -eq 'ShipDateKey' }
        $r2.RelationshipStatus | Should -Be 'False'
        $r2.RelationshipCrossFilteringBehavior | Should -Be 'BothDirections'
    }
    It 'writes the dependency rows for measures, the calculated column and the calculation items' {
        $script:MdRows.Count | Should -Be 6
        $ts = @($script:MdRows | Where-Object { $_.ObjectName -eq 'Total Sales' })
        $ts.Count | Should -Be 1
        $ts[0].ObjectType | Should -Be 'Measure'
        $ts[0].DependsOn | Should -Be "'Sales'[Amount]"
        $ts[0].DependsOnType | Should -Be 'Column'
        ($script:MdRows | Where-Object { $_.ObjectName -eq 'Total Cost' }).DependsOn | Should -Be "'Sales'[Cost]"
        $margin = @($script:MdRows | Where-Object { $_.ObjectName -eq 'Margin' })
        @($margin | ForEach-Object { $_.DependsOn }) | Should -Be @('[Total Sales]', '[Total Cost]')
        @($margin | ForEach-Object { $_.DependsOnType } | Select-Object -Unique) | Should -Be @('Measure')
        $cc = $script:MdRows | Where-Object { $_.ObjectType -eq 'CalculatedColumn' }
        $cc.ObjectName | Should -Be 'Amount x2'
        $cc.DependsOn | Should -Be "'Sales'[Amount]"
        $ytd = $script:MdRows | Where-Object { $_.ObjectType -eq 'CalculationItem' }
        $ytd.ObjectName | Should -Be 'YTD'
        $ytd.DependsOn | Should -Be "'Date'[Date]"
        $ytd.DependsOnType | Should -Be 'Column'
    }
    It 'ignores references inside comments and string literals' {
        @($script:MdRows | Where-Object { $_.DependsOn -like '*Fake*' -or $_.DependsOn -like '*Not A Measure*' }).Count | Should -Be 0
    }
    It 'emits the dependency rows in csx order (measures, calculated columns, calculation items)' {
        @($script:MdRows | ForEach-Object { $_.ObjectType } | Select-Object -Unique) | Should -Be @('Measure', 'CalculatedColumn', 'CalculationItem')
    }
    It 'returns a failed result instead of throwing for a missing file' {
        $r = Export-IQModelDetailFromBim -BimPath (Join-Path $script:OutFolder 'missing.bim') -OutputFolder $script:OutFolder -ModelName 'X ~ Y' -ModelId 'X ~ Y' -AsOfDate '2026-09-05'
        $r.Success | Should -BeFalse
        $r.Message | Should -Match 'Could not parse'
    }
}

Describe 'Get-IQDaxReferences' {
    BeforeAll {
        $script:Tables = @('Sales', 'Date', 'Time Intelligence')
        $script:Measures = @('Total Sales', 'Total Cost')
        $script:Columns = @{ Sales = @('Amount', 'Cost', 'Region'); Date = @('Date', 'Year'); 'Time Intelligence' = @('Name') }
    }
    It 'resolves quoted and unquoted table-column references and measures' {
        $refs = @(Get-IQDaxReferences -Expression "CALCULATE([Total Sales], FILTER('Date', 'Date'[Year] = 2024), Sales[Region] = ""x"")" -KnownTables $script:Tables -KnownMeasures $script:Measures -KnownColumns $script:Columns)
        @($refs | ForEach-Object { $_.DependsOn }) | Should -Be @('[Total Sales]', "'Date'", "'Date'[Year]", "'Sales'[Region]")
        @($refs | ForEach-Object { $_.DependsOnType }) | Should -Be @('Measure', 'Table', 'Column', 'Column')
    }
    It 'treats a table-qualified measure as a measure and resolves bare columns against the current table' {
        $refs = @(Get-IQDaxReferences -Expression 'Sales[Total Cost] + [Amount]' -KnownTables $script:Tables -KnownMeasures $script:Measures -KnownColumns $script:Columns -CurrentTable 'Sales')
        @($refs | ForEach-Object { $_.DependsOn }) | Should -Be @('[Total Cost]', "'Sales'[Amount]")
    }
    It 'ignores comments, string literals, function names and VAR names that look like tables' {
        $expr = "VAR Sales = 1 // Sales[Amount] in a comment`r`n/* 'Date'[Date] */ RETURN Sales + COUNTROWS ( Date ) + [Cost] + LEN(""[Region]"")"
        $refs = @(Get-IQDaxReferences -Expression $expr -KnownTables $script:Tables -KnownMeasures $script:Measures -KnownColumns $script:Columns -CurrentTable 'Sales')
        @($refs | ForEach-Object { $_.DependsOn }) | Should -Be @("'Date'", "'Sales'[Cost]")
    }
    It 'escapes quotes and brackets like Tabular Editor and types calculation group tables' {
        $refs = @(Get-IQDaxReferences -Expression "'Date''s'[Col]]x] + COUNTROWS('Time Intelligence')" -KnownTables @("Date's", 'Time Intelligence') -KnownMeasures @() -KnownColumns @{ "Date's" = @('Col]x') } -CalculationGroupTables @('Time Intelligence'))
        $refs[0].DependsOn | Should -Be "'Date''s'[Col]]x]"
        $refs[0].DependsOnType | Should -Be 'Column'
        $refs[1].DependsOn | Should -Be "'Time Intelligence'"
        $refs[1].DependsOnType | Should -Be 'CalculationGroupTable'
    }
    It 'de-duplicates and returns nothing for blank expressions' {
        @(Get-IQDaxReferences -Expression '[Total Sales] + [Total Sales]' -KnownTables $script:Tables -KnownMeasures $script:Measures).Count | Should -Be 1
        @(Get-IQDaxReferences -Expression '' -KnownTables $script:Tables -KnownMeasures $script:Measures).Count | Should -Be 0
        @(Get-IQDaxReferences -Expression 'SELECTEDMEASURE()' -KnownTables $script:Tables -KnownMeasures $script:Measures).Count | Should -Be 0
    }
}
