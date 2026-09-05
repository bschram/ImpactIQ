# Dax.Tests.ps1 - ImpactIQ.Dax.ps1 (brief section 7.3, 13): executeQueries client and the INFO.* model-detail fallback.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Ids = Get-IQTestFixtureJson -Relative 'ids.json'
    $script:Base = Initialize-IQTestContext -Prefix 'dax'
    Mock Get-IQToken { 'test-token' }
    Mock Invoke-IQApi { Invoke-IQTestApiFixture -Method $Method -Path $Path -Body $Body -Query $Query -Api $Api -Raw:$Raw }
    $script:ModelDetailHeader = 'Type,Table,Name,FormatString,DisplayFolder,Description,IsHidden,TableStorageMode,Expression,ModelAsOfDate,ModelName,ModelID,RelationshipFromTable,RelationshipFromColumn,RelationshipToTable,RelationshipToColumn,RelationshipStatus,RelationshipFromCardinality,RelationshipToCardinality,RelationshipCrossFilteringBehavior'
    $script:MdHeader = 'ObjectName,ObjectType,DependsOn,DependsOnType,ModelAsOfDate,ModelName,ModelID'
    $script:Dataset = [pscustomobject]@{ DatasetId = $script:Ids.d1; DatasetName = 'Finance Model'; WorkspaceId = $script:Ids.ws1; WorkspaceName = 'Finance [Prod]'; WorkspaceIsOnDedicatedCapacity = $true }
    function Get-CsvHeaderLine { param([string]$Path) return ((Get-Content -LiteralPath $Path -Encoding UTF8 -TotalCount 1) -replace '^﻿', '') }
    function Get-UnquotedHeader { param([string]$Path) return ((Get-CsvHeaderLine -Path $Path) -replace '"', '') }
}
AfterAll {
    $script:IQTestDaxOverrides.Clear()
    Remove-IQTestFolder -Path $script:Base
}

Describe 'ConvertFrom-IQDaxRows' {
    It 'strips [Column], Table[Column] and ''Table''[Column] wrappers when unambiguous' {
        $rows = @([pscustomobject]@{ '[ID]' = 1; 'Sales[Amount]' = 2; "'Date Table'[Year]" = 2026 })
        $out = @(ConvertFrom-IQDaxRows -Rows $rows)
        $out.Count | Should -Be 1
        @($out[0].PSObject.Properties.Name) | Should -Be @('ID', 'Amount', 'Year')
        $out[0].Year | Should -Be 2026
    }
    It 'keeps the original names when two columns would collapse to the same short name' {
        $rows = @([pscustomobject]@{ 'Sales[Name]' = 'a'; 'Product[Name]' = 'b' })
        $out = @(ConvertFrom-IQDaxRows -Rows $rows)
        @($out[0].PSObject.Properties.Name) | Should -Be @('Sales[Name]', 'Product[Name]')
    }
    It 'returns an empty array for no rows' {
        @(ConvertFrom-IQDaxRows -Rows @()).Count | Should -Be 0
        @(ConvertFrom-IQDaxRows -Rows $null).Count | Should -Be 0
    }
}

Describe 'Invoke-IQDaxQuery' {
    BeforeEach { $script:IQTestApiCalls.Clear() }
    It 'POSTs groups/{ws}/datasets/{id}/executeQueries with the query and includeNulls and returns flattened rows' {
        $rows = @(Invoke-IQDaxQuery -WorkspaceId $script:Ids.ws1 -DatasetId $script:Ids.d1 -Dax 'EVALUATE INFO.VIEW.TABLES()')
        $rows.Count | Should -Be 3
        $rows[0].Name | Should -Be 'Sales'
        $rows[0].PSObject.Properties['[Name]'] | Should -BeNullOrEmpty
        $script:IQTestApiCalls[0] | Should -Be ('PowerBI POST groups/' + $script:Ids.ws1 + '/datasets/' + $script:Ids.d1 + '/executeQueries')
    }
    It 'uses the group-less path for My Workspace datasets' {
        Invoke-IQDaxQuery -WorkspaceId 'My Workspace' -DatasetId $script:Ids.d1 -Dax 'EVALUATE INFO.VIEW.TABLES()' | Out-Null
        $script:IQTestApiCalls[0] | Should -Be ('PowerBI POST datasets/' + $script:Ids.d1 + '/executeQueries')
    }
    It 'sends the body shape { queries[{query}], serializerSettings{includeNulls:true} } (+ impersonatedUserName)' {
        Mock Invoke-IQApi { $script:LastBody = $Body; return (Get-IQTestFixtureJson -Relative 'dax/info-view-tables.json') }
        Invoke-IQDaxQuery -WorkspaceId $script:Ids.ws1 -DatasetId $script:Ids.d1 -Dax 'EVALUATE INFO.VIEW.TABLES()' -Impersonate 'user@contoso.gov' | Out-Null
        $json = ConvertTo-Json -InputObject $script:LastBody -Depth 20 -Compress
        $json | Should -Match '"queries":\[\{"query":"EVALUATE INFO\.VIEW\.TABLES\(\)"\}\]'
        $json | Should -Match '"includeNulls":true'
        $json | Should -Match '"impersonatedUserName":"user@contoso\.gov"'
    }
    It 'throws with the engine message when results[].error is returned' {
        Mock Invoke-IQApi { return (Get-IQTestFixtureJson -Relative 'dax/error-result.json') }
        { Invoke-IQDaxQuery -WorkspaceId $script:Ids.ws1 -DatasetId $script:Ids.d1 -Dax 'EVALUATE INFO.TABLES()' } | Should -Throw -ExpectedMessage '*INFO*'
    }
    It 'throws a permission / tenant-setting hint when Invoke-IQApi returns $null (HTTP 400/403/404)' {
        Mock Invoke-IQApi { return $null }
        { Invoke-IQDaxQuery -WorkspaceId $script:Ids.ws1 -DatasetId $script:Ids.d1 -Dax 'EVALUATE INFO.TABLES()' } | Should -Throw
    }
}

Describe 'Get-IQModelDetailViaDax (INFO fixtures -> csx-shaped CSVs)' {
    BeforeAll {
        $script:OutFolder = Join-Path $script:Base 'dax-out'
        $script:IQTestApiCalls.Clear()
        $script:Result = Get-IQModelDetailViaDax -Dataset $script:Dataset -OutputFolder $script:OutFolder
        $script:Csv = Join-Path $script:OutFolder 'Finance (Prod) ~ Finance Model.csv'
        $script:Md = Join-Path $script:OutFolder 'Finance (Prod) ~ Finance Model_MD.csv'
        if (Test-Path -LiteralPath $script:Csv) { $script:Rows = @(Import-Csv -LiteralPath $script:Csv -Encoding UTF8) } else { $script:Rows = @() }
        if (Test-Path -LiteralPath $script:Md) { $script:MdRows = @(Import-Csv -LiteralPath $script:Md -Encoding UTF8) } else { $script:MdRows = @() }
    }
    It 'succeeds and writes "<CleanWs> ~ <CleanModel>.csv" and "..._MD.csv"' {
        $script:Result.Success | Should -BeTrue -Because $script:Result.Message
        $script:Result.Method | Should -Be 'Dax'
        $script:Csv | Should -Exist
        $script:Md | Should -Exist
        @($script:Result.Outputs) | Should -Contain $script:Csv
        @($script:Result.Outputs) | Should -Contain $script:Md
    }
    It 'the Semantic Models header equals the csx column list exactly (x1 section 2.3)' {
        Get-UnquotedHeader -Path $script:Csv | Should -BeExactly $script:ModelDetailHeader
    }
    It 'the Measure Dependencies header equals the csx column list exactly (x1 section 3.3)' {
        Get-UnquotedHeader -Path $script:Md | Should -BeExactly $script:MdHeader
    }
    It 'quotes every field and doubles embedded quotes' {
        $line = @(Get-Content -LiteralPath $script:Csv -Encoding UTF8 | Where-Object { $_ -like '"Column","Sales","Amount"*' })[0]
        $line | Should -Not -BeNullOrEmpty
        $line | Should -Match 'Sales ""amount"" in USD'
        ($line -split ',' | Select-Object -First 3) | ForEach-Object { $_ | Should -Match '^".*"$' }
    }
    It 'emits every row kind of the csx script' {
        $types = @($script:Rows | ForEach-Object { $_.Type } | Sort-Object -Unique)
        foreach ($t in @('Table', 'CalculationGroup', 'CalculationItem', 'Column', 'CalculatedColumn', 'Measure', 'Hierarchy', 'Level', 'Partition', 'RLSFilter', 'Relationship')) {
            $types | Should -Contain $t
        }
    }
    It 'Table rows carry the storage mode; the RowNumber column is excluded; calculated columns appear as Column AND CalculatedColumn' {
        $sales = @($script:Rows | Where-Object { $_.Type -eq 'Table' -and $_.Name -eq 'Sales' })[0]
        $sales.Table | Should -Be 'Sales'
        $sales.TableStorageMode | Should -Be 'Import'
        $sales.IsHidden | Should -Be 'False'
        @($script:Rows | Where-Object { $_.Type -eq 'Column' -and $_.Name -like 'RowNumber*' }).Count | Should -Be 0
        @($script:Rows | Where-Object { $_.Type -eq 'Column' -and $_.Name -eq 'Margin' }).Count | Should -Be 1
        $cc = @($script:Rows | Where-Object { $_.Type -eq 'CalculatedColumn' -and $_.Name -eq 'Margin' })
        $cc.Count | Should -Be 1
        $cc[0].Expression | Should -Be 'Sales[Amount] - Sales[Cost]'
        $amount = @($script:Rows | Where-Object { $_.Type -eq 'Column' -and $_.Name -eq 'Amount' })[0]
        $amount.FormatString | Should -Be '#,0.00'
        $amount.DisplayFolder | Should -Be 'Money'
        $amount.Expression | Should -Be ''
    }
    It 'Measure rows have Expression/FormatString/DisplayFolder/Description/IsHidden' {
        $m = @($script:Rows | Where-Object { $_.Type -eq 'Measure' -and $_.Name -eq 'Total Sales' })[0]
        $m.Table | Should -Be 'Sales'
        $m.Expression | Should -Be 'SUM(Sales[Amount])'
        $m.FormatString | Should -Be '#,0'
        $m.DisplayFolder | Should -Be 'KPIs'
        $m.Description | Should -Be 'Sum of amount'
        $m.IsHidden | Should -Be 'False'
        (@($script:Rows | Where-Object { $_.Type -eq 'Measure' -and $_.Name -eq 'Total Cost' })[0]).IsHidden | Should -Be 'True'
    }
    It 'Relationship rows follow the csx layout (Table=From table, Name=From column, Expression=relationship name)' {
        $r = @($script:Rows | Where-Object { $_.Type -eq 'Relationship' })
        $r.Count | Should -Be 1
        $r[0].Table | Should -Be 'Sales'
        $r[0].Name | Should -Be 'DateKey'
        $r[0].Expression | Should -Be 'e1a2b3c4-0000-4000-8000-000000000200'
        $r[0].RelationshipFromTable | Should -Be 'Sales'
        $r[0].RelationshipFromColumn | Should -Be 'DateKey'
        $r[0].RelationshipToTable | Should -Be 'Date'
        $r[0].RelationshipToColumn | Should -Be 'Date'
        $r[0].RelationshipStatus | Should -Be 'True'
        $r[0].RelationshipFromCardinality | Should -Be 'Many'
        $r[0].RelationshipToCardinality | Should -Be 'One'
        $r[0].RelationshipCrossFilteringBehavior | Should -Be 'OneDirection'
    }
    It 'Partition, RLSFilter, CalculationGroup/Item, Hierarchy and Level rows are joined on IDs' {
        $p = @($script:Rows | Where-Object { $_.Type -eq 'Partition' -and $_.Table -eq 'Sales' })[0]
        $p.Expression | Should -Match 'Sql\.Database'
        $p.TableStorageMode | Should -Be 'Import'
        $rls = @($script:Rows | Where-Object { $_.Type -eq 'RLSFilter' })[0]
        $rls.Table | Should -Be 'Sales'
        $rls.Name | Should -Be 'Region Manager'
        $rls.Expression | Should -Be '[Region] = "West"'
        $cg = @($script:Rows | Where-Object { $_.Type -eq 'CalculationGroup' })[0]
        $cg.Table | Should -Be 'Time Intelligence'
        $cg.Name | Should -Be 'Time Intelligence'
        $ci = @($script:Rows | Where-Object { $_.Type -eq 'CalculationItem' -and $_.Name -eq 'YTD' })[0]
        $ci.Table | Should -Be 'Time Intelligence'
        $ci.Expression | Should -Match 'DATESYTD'
        $h = @($script:Rows | Where-Object { $_.Type -eq 'Hierarchy' })[0]
        $h.Table | Should -Be 'Date'
        $h.Name | Should -Be 'Calendar'
        $h.DisplayFolder | Should -Be 'Hierarchies'
        $lv = @($script:Rows | Where-Object { $_.Type -eq 'Level' })
        $lv.Count | Should -Be 2
        $lv[0].Table | Should -Be 'Date'
    }
    It 'stamps ModelAsOfDate = RunId, ModelName = "<CleanWs> ~ <CleanModel>" and ModelID = dataset id for dedicated capacity' {
        foreach ($row in $script:Rows) {
            $row.ModelAsOfDate | Should -Be $script:IQ.RunId
            $row.ModelName | Should -Be 'Finance (Prod) ~ Finance Model'
            $row.ModelID | Should -Be $script:Ids.d1
        }
    }
    It 'maps INFO.CALCDEPENDENCY rows to the _MD.csv shape (direct dependencies only, other object types skipped)' {
        $script:Result.DependencySource | Should -Be 'INFO.CALCDEPENDENCY'
        $script:MdRows.Count | Should -Be 8
        $ts = @($script:MdRows | Where-Object { $_.ObjectName -eq 'Total Sales' })
        $ts.Count | Should -Be 1
        $ts[0].ObjectType | Should -Be 'Measure'
        $ts[0].DependsOn | Should -Be "'Sales'[Amount]"
        $ts[0].DependsOnType | Should -Be 'Column'
        $mp = @($script:MdRows | Where-Object { $_.ObjectName -eq 'Margin %' })
        $mp.Count | Should -Be 2
        @($mp | ForEach-Object { $_.DependsOn }) | Should -Contain '[Total Sales]'
        @($mp | ForEach-Object { $_.DependsOnType } | Sort-Object -Unique) | Should -Be @('Measure')
        $cc = @($script:MdRows | Where-Object { $_.ObjectName -eq 'Margin' })
        $cc.Count | Should -Be 2
        $cc[0].ObjectType | Should -Be 'CalculatedColumn'
        $ytd = @($script:MdRows | Where-Object { $_.ObjectName -eq 'YTD' })
        $ytd.Count | Should -Be 2
        $ytd[0].ObjectType | Should -Be 'CalculationItem'
        @($ytd | ForEach-Object { $_.DependsOn }) | Should -Contain "'Date'"
        @($script:MdRows | Where-Object { $_.ObjectName -eq 'Region Manager' }).Count | Should -Be 0 -Because 'ROWS_ALLOWED is not a csx object type'
        $script:MdRows[0].ModelName | Should -Be 'Finance (Prod) ~ Finance Model'
        $script:MdRows[0].ModelID | Should -Be $script:Ids.d1
    }
    It 'caches every raw result under extracts\dax\<key>\ and a re-run rebuilds the CSVs without querying' {
        $extract = Join-Path (Join-Path $script:IQ.RunPaths.Extracts 'dax') (Get-IQSafeKey -Value $script:Ids.d1)
        $extract | Should -Exist
        @(Get-ChildItem -LiteralPath $extract -Filter '*.json').Count | Should -BeGreaterOrEqual 4
        Remove-Item -LiteralPath $script:Csv -Force
        $script:IQTestApiCalls.Clear()
        $again = Get-IQModelDetailViaDax -Dataset $script:Dataset -OutputFolder $script:OutFolder
        $again.Success | Should -BeTrue
        $script:Csv | Should -Exist
        $script:IQTestApiCalls.Count | Should -Be 0
    }
    It 'uses ModelID = "<CleanWs> ~ <CleanModel>" for Pro (shared capacity) models' {
        $pro = [pscustomobject]@{ DatasetId = $script:Ids.d1; DatasetName = 'Finance Model'; WorkspaceId = $script:Ids.ws1; WorkspaceName = 'Finance [Prod]'; WorkspaceIsOnDedicatedCapacity = $false }
        $folder = Join-Path $script:Base 'dax-pro'
        $r = Get-IQModelDetailViaDax -Dataset $pro -OutputFolder $folder
        $r.Success | Should -BeTrue
        (@(Import-Csv -LiteralPath $r.Csv -Encoding UTF8)[0]).ModelID | Should -Be 'Finance (Prod) ~ Finance Model'
    }
}

Describe 'Get-IQModelDetailViaDax when raw INFO.* is rejected (HTTP 400, executeQueries limitation)' {
    BeforeAll {
        foreach ($k in @('info-partitions', 'info-model', 'info-roles', 'info-tablepermissions', 'info-calculationgroups', 'info-calculationitems', 'info-hierarchies', 'info-levels', 'info-calcdependency')) {
            $script:IQTestDaxOverrides[$k] = $null
        }
        $ds = [pscustomobject]@{ DatasetId = $script:Ids.d3; DatasetName = 'Sales Model'; WorkspaceId = $script:Ids.ws2; WorkspaceName = 'Sales'; WorkspaceIsOnDedicatedCapacity = $false }
        $script:IQTestApiCalls.Clear()
        $script:ProResult = Get-IQModelDetailViaDax -Dataset $ds -OutputFolder (Join-Path $script:Base 'dax-raw-blocked')
    }
    AfterAll { $script:IQTestDaxOverrides.Clear() }
    It 'still succeeds from INFO.VIEW.* and reports the unavailable parts' {
        $script:ProResult.Success | Should -BeTrue -Because $script:ProResult.Message
        @($script:ProResult.Unavailable).Count | Should -BeGreaterOrEqual 1
        $script:ProResult.Csv | Should -Exist
        $script:ProResult.MdCsv | Should -Exist
        Get-UnquotedHeader -Path $script:ProResult.Csv | Should -BeExactly $script:ModelDetailHeader
    }
    It 'stops sending raw INFO.* queries after the first "not supported" answer (one Warn)' {
        $raw = @($script:IQTestApiCalls | Where-Object { $_ -like '*executeQueries' })
        $raw.Count | Should -BeLessThan 13
        (Get-Content -LiteralPath $script:IQ.LogFile -Raw) | Should -Match '(?i)not available through executeQueries|not supported'
    }
    It 'derives measure dependencies by expression parsing when INFO.CALCDEPENDENCY is unavailable' {
        $script:ProResult.DependencySource | Should -Not -Be 'INFO.CALCDEPENDENCY'
        $md = @(Import-Csv -LiteralPath $script:ProResult.MdCsv -Encoding UTF8)
        $md.Count | Should -BeGreaterOrEqual 1
        @($md | Where-Object { $_.ObjectName -eq 'Margin %' -and $_.DependsOn -eq '[Total Sales]' -and $_.DependsOnType -eq 'Measure' }).Count | Should -Be 1
    }
    It 'fails cleanly (Success=$false, InfoUnsupported) when even INFO.VIEW.TABLES() and INFO.TABLES() are rejected' {
        $script:IQTestDaxOverrides['info-view-tables'] = $null
        $script:IQTestDaxOverrides['info-tables'] = $null
        $ds = [pscustomobject]@{ DatasetId = $script:Ids.d4; DatasetName = 'HR Model'; WorkspaceId = $script:Ids.ws3; WorkspaceName = 'HR Analytics'; WorkspaceIsOnDedicatedCapacity = $false }
        $r = Get-IQModelDetailViaDax -Dataset $ds -OutputFolder (Join-Path $script:Base 'dax-none')
        $r.Success | Should -BeFalse
        $r.Message | Should -Not -BeNullOrEmpty
        $r.InfoUnsupported | Should -BeTrue
    }
}

Describe 'Export-IQModelDetailFromBim (TE2-free .bim path)' -Skip:(-not (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'Config/Modules/ImpactIQ.Bim.ps1'))) {
    BeforeAll {
        $script:BimResult = Export-IQModelDetailFromBim -BimPath (Get-IQTestFixturePath -Relative 'bim/sample-model.bim') -OutputFolder (Join-Path $script:Base 'bim-out') -ModelName 'Finance (Prod) ~ Ledger' -ModelId '11111111-1111-1111-1111-111111111111' -AsOfDate '2026-09-04'
    }
    It 'writes both CSVs with the exact csx headers' {
        $script:BimResult.Success | Should -BeTrue -Because $script:BimResult.Message
        Get-UnquotedHeader -Path $script:BimResult.Csv | Should -BeExactly $script:ModelDetailHeader
        Get-UnquotedHeader -Path $script:BimResult.MdCsv | Should -BeExactly $script:MdHeader
        @(Import-Csv -LiteralPath $script:BimResult.Csv -Encoding UTF8).Count | Should -BeGreaterThan 3
    }
}
