# Assemble.Tests.ps1 - ImpactIQ.Assemble.ps1 (brief section 9, 14): the four workbooks against Config/SheetContract.json.
# Discovery-time values (Pester evaluates -Skip / -TestCases before BeforeAll runs).
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
$script:HasImportExcel = [bool](Get-Module -ListAvailable -Name ImportExcel | Select-Object -First 1)
$script:ContractCases = @()
$script:ContractFile = Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'Config') 'SheetContract.json'
if (Test-Path -LiteralPath $script:ContractFile) {
    $iqContractDoc = ConvertFrom-Json -InputObject ([System.IO.File]::ReadAllText($script:ContractFile))
    foreach ($iqWb in $iqContractDoc.workbooks.PSObject.Properties) {
        foreach ($iqSheet in $iqWb.Value.PSObject.Properties) {
            $script:ContractCases += @{ Workbook = $iqWb.Name; Sheet = $iqSheet.Name; Columns = @($iqSheet.Value.expectedColumns) }
        }
    }
}

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    # top-level (discovery) variables are not visible in the run phase: recompute what the run needs
    $script:HasImportExcel = [bool](Get-Module -ListAvailable -Name ImportExcel | Select-Object -First 1)
    if ($script:HasImportExcel) { Import-Module ImportExcel -ErrorAction Stop -WarningAction SilentlyContinue }
    $script:Ids = Get-IQTestFixtureJson -Relative 'ids.json'
    Mock Get-IQToken { 'test-token' }
    Mock Invoke-IQApi { Invoke-IQTestApiFixture -Method $Method -Path $Path -Body $Body -Query $Query -Api $Api -Raw:$Raw }

    # A realistic run state: inventory from the API fixtures, one manifest failure, csx-shaped extract files, a dataflow extract.
    $script:Base = Initialize-IQTestContext -Options @{ RunMode = 'Workspaces'; WorkspaceId = @($script:Ids.ws1, $script:Ids.ws2) } -Prefix 'asm'
    Invoke-IQStage -Name Inventory -Body { Invoke-IQInventoryStage | Out-Null } | Out-Null
    Set-IQItemDone -Stage 'ModelBackup' -ItemKey $script:Ids.d3 -Item 'Sales ~ Sales Model' -Status Failed -Message 'XMLA endpoint disabled on this capacity' | Out-Null
    Copy-IQTestFixtureFolder -Relative 'extracts/report-detail' -Destination $script:IQ.RunPaths.ReportBackups
    Remove-Item -LiteralPath (Join-Path $script:IQ.RunPaths.ReportBackups '_columns-used.json') -Force -ErrorAction SilentlyContinue
    # A header-only NON-contract txt (ReportExports.txt as written when nothing was exported) must still become a sheet.
    [System.IO.File]::WriteAllText((Join-Path $script:IQ.RunPaths.ReportBackups 'ReportExports.txt'), "ReportName`tReportID`tExportStatus`tExportMessage`r`n", (New-Object System.Text.UTF8Encoding($false)))
    # A temp workbook left behind by a killed earlier run must be cleaned up by the next Assemble.
    [System.IO.File]::WriteAllText((Join-Path $script:Base 'Power BI Environment Detail.tmp-deadbeef.xlsx'), 'stale')
    Copy-IQTestFixtureFolder -Relative 'extracts/model-detail' -Destination $script:IQ.RunPaths.ModelBackups
    $dfFolder = Join-Path $script:IQ.RunPaths.Extracts 'dataflows'
    New-Item -ItemType Directory -Path $dfFolder -Force | Out-Null
    $gen1 = Get-IQTestFixtureJson -Relative 'dataflows/gen1-model.json'
    $queries = @(ConvertFrom-IQDataflowDocument -Content ([string]$gen1.'pbi:mashup'.document) -DataflowId $script:Ids.df1 -DataflowName 'Sales [Prod]/Data' -WorkspaceName 'Finance [Prod]' -ReportDate $script:IQ.RunId)
    ConvertTo-IQJsonFile -Object ([ordered]@{ SchemaVersion = 1; DataflowId = $script:Ids.df1; DataflowName = 'Sales [Prod]/Data'; WorkspaceId = $script:Ids.ws1; WorkspaceName = 'Finance [Prod]'; Generation = 'Gen1'; Status = 'Succeeded'; ReportDate = $script:IQ.RunId; QueryCount = $queries.Count; Queries = $queries; Entities = @() }) -Path (Join-Path $dfFolder ((Get-IQSafeKey -Value $script:Ids.df1) + '.json'))
    $script:Contract = ConvertFrom-IQJsonFile -Path (Join-Path (Join-Path $script:Base 'Config') 'SheetContract.json')
    $script:Result = $null
    if ($script:HasImportExcel) { $script:Status = Invoke-IQStage -Name Assemble -Body { $script:Result = Invoke-IQAssembleStage } }
    $script:EnvPath = Join-Path $script:Base 'Power BI Environment Detail.xlsx'
    $script:RepPath = Join-Path $script:Base 'Report Detail.xlsx'
    $script:ModPath = Join-Path $script:Base 'Model Detail.xlsx'
    $script:DfPath = Join-Path $script:Base 'Dataflow Detail.xlsx'
}
AfterAll { Remove-IQTestFolder -Path $script:Base }

Describe 'Invoke-IQAssembleStage' -Skip:(-not $script:HasImportExcel) {
    It 'builds the four workbooks in the BaseFolder (and the dataflow workbook in its run folder) without leaving temp files' {
        $script:Status | Should -Be 'Completed'
        foreach ($p in @($script:EnvPath, $script:RepPath, $script:ModPath, $script:DfPath)) { $p | Should -Exist }
        (Join-Path $script:IQ.RunPaths.DataflowBackups 'Dataflow Detail.xlsx') | Should -Exist
        @(Get-ChildItem -LiteralPath $script:Base -Recurse -Filter '*.tmp*').Count | Should -Be 0
    }
    It 'records manifest.outputs and one checkpoint per workbook' {
        $m = ConvertFrom-IQJsonFile -Path $script:IQ.RunPaths.Manifest
        $m.outputs.environmentWorkbook | Should -Be $script:EnvPath
        $m.outputs.reportWorkbook | Should -Be $script:RepPath
        $m.outputs.modelWorkbook | Should -Be $script:ModPath
        $m.outputs.dataflowWorkbook | Should -Be $script:DfPath
        [int]$m.stages.Assemble.itemsDone | Should -Be 4
    }
    It 'every contract sheet exists with every expected column: <Workbook> / <Sheet>' -TestCases $script:ContractCases {
        param($Workbook, $Sheet, $Columns)
        $path = Join-Path $script:Base $Workbook
        @(Get-ExcelSheetInfo -Path $path | Where-Object { $_.Name -eq $Sheet }).Count | Should -Be 1
        $header = @(Get-IQTestSheetHeader -Path $path -Sheet $Sheet)
        $header.Count | Should -BeGreaterOrEqual 1
        $missing = @($Columns | Where-Object { $_ -notin $header })
        $missing | Should -BeNullOrEmpty -Because ("columns missing from {0}/{1}: {2}" -f $Workbook, $Sheet, ($missing -join ','))
    }
}

Describe 'Power BI Environment Detail.xlsx content' -Skip:(-not $script:HasImportExcel) {
    BeforeAll { $script:EnvSheets = @(Get-ExcelSheetInfo -Path $script:EnvPath | Sort-Object Index | ForEach-Object { $_.Name }) }
    It 'keeps the monolith sheet order for the first 17 sheets' {
        $expected = @('Workspaces', 'FabricItems', 'Connections', 'Gateways', 'ItemConnections', 'Datasets', 'DatasetSourcesInfo', 'DatasetRefreshHistory', 'DatasetRefreshSchedule', 'Dataflows', 'DataflowLineage', 'DataflowSourcesInfo', 'DataflowRefreshHistory', 'Reports', 'ReportPages', 'Apps', 'AppReports')
        (($script:EnvSheets | Select-Object -First 17) -join '|') | Should -Be ($expected -join '|')
    }
    It 'adds the new sheets (Dashboards, DashboardTiles, Capacities, users, parameters, DQ schedule, RunSummary, Failures)' {
        foreach ($s in @('Dashboards', 'DashboardTiles', 'Capacities', 'WorkspaceUsers', 'DatasetUsers', 'DatasetParameters', 'RunSummary', 'Failures')) { $script:EnvSheets | Should -Contain $s }
        @($script:EnvSheets | Where-Object { $_ -like 'DatasetD*RefreshSchedule' }).Count | Should -Be 1 -Because 'DatasetDirectQueryRefreshSchedule (31-char Excel limit -> DatasetDQRefreshSchedule)'
    }
    It 'Workspaces / Datasets / Reports rows come from the inventory with the monolith columns' {
        $ws = @(Import-Excel -Path $script:EnvPath -WorksheetName 'Workspaces')
        @($ws | Where-Object { $_.WorkspaceId -eq $script:Ids.ws1 }).Count | Should -Be 1
        $ds = @(Import-Excel -Path $script:EnvPath -WorksheetName 'Datasets')
        $ds.Count | Should -Be 3
        (@($ds | Where-Object { $_.DatasetId -eq $script:Ids.d1 })[0]).DatasetName | Should -Be 'Finance Model'
        $rp = @(Import-Excel -Path $script:EnvPath -WorksheetName 'Reports')
        (@($rp | Where-Object { $_.ReportId -eq $script:Ids.r2 })[0]).DatasetWorkspaceId | Should -Be $script:Ids.ws1
        $pg = @(Import-Excel -Path $script:EnvPath -WorksheetName 'ReportPages')
        $pg.Count | Should -BeGreaterOrEqual 5
    }
    It 'Failures and RunSummary sheets reflect the manifest' {
        $f = @(Import-Excel -Path $script:EnvPath -WorksheetName 'Failures')
        $f.Count | Should -Be 1
        $f[0].Stage | Should -Be 'ModelBackup'
        $f[0].Message | Should -Match 'XMLA endpoint disabled'
        $rs = @(Import-Excel -Path $script:EnvPath -WorksheetName 'RunSummary')
        $rs[0].Stage | Should -Be '(Run)'
        @($rs | Where-Object { $_.Stage -eq 'Inventory' }).Count | Should -Be 1
    }
    It 'Capacities and DashboardTiles rows are typed sensibly' {
        $cap = @(Import-Excel -Path $script:EnvPath -WorksheetName 'Capacities')
        $cap.Count | Should -Be 1
        $cap[0].CapacityAdmins | Should -Be 'admin@contoso.gov;capadmin@contoso.gov'
        $tiles = @(Import-Excel -Path $script:EnvPath -WorksheetName 'DashboardTiles')
        $tiles[0].TileRowSpan | Should -Be 2
    }
}

Describe 'Report Detail.xlsx content (csx *.txt parsing, monolith 3063-3110)' -Skip:(-not $script:HasImportExcel) {
    It 'parses tab-separated UTF-8 text with the exact header count' {
        $pages = @(Import-Excel -Path $script:RepPath -WorksheetName 'Pages')
        $pages.Count | Should -Be 2
        $pages[1].Name | Should -Be 'Détail ünïcode'
        $pages[0].ReportID | Should -Be $script:Ids.r1
    }
    It 'keeps a DAX expression starting with "=" verbatim and the 12-column ReportLevelMeasures layout' {
        $rlm = @(Import-Excel -Path $script:RepPath -WorksheetName 'ReportLevelMeasures')
        $rlm.Count | Should -Be 1
        $rlm[0].Expression | Should -Be "=SUM('Sales'[Amount])"
        $rlm[0].FormatString | Should -Be '#,0'
        $rlm[0].ReportDate | Should -Match '^2026-09-04'
    }
    It 'a header-only txt file gives a header-only sheet with the contract columns (C8-04)' {
        @(Import-Excel -Path $script:RepPath -WorksheetName 'Bookmarks' -WarningAction SilentlyContinue).Count | Should -Be 0
        (Get-IQTestSheetHeader -Path $script:RepPath -Sheet 'Bookmarks') | Should -Contain 'SuppressData'
    }
}

Describe 'Model Detail.xlsx content (*.csv / *_MD.csv split, monolith 3209-3290)' -Skip:(-not $script:HasImportExcel) {
    It 'Semantic Models rows come from every *.csv except *_MD.csv' {
        $sm = @(Import-Excel -Path $script:ModPath -WorksheetName 'Semantic Models')
        $sm.Count | Should -Be 5
        (@($sm | Where-Object { $_.Type -eq 'Measure' })[0]).Description | Should -Be 'Über total'
        $sm[0].ModelName | Should -Be 'Finance (Prod) ~ Finance Model'
        $sm[0].ModelID | Should -Be $script:Ids.d1
    }
    It 'Measure Dependencies rows come from *_MD.csv' {
        $md = @(Import-Excel -Path $script:ModPath -WorksheetName 'Measure Dependencies')
        $md.Count | Should -Be 4
        $md[0].DependsOn | Should -Be "'Sales'[Amount]"
        @($md | ForEach-Object { $_.ObjectType } | Sort-Object -Unique) | Should -Be @('CalculatedColumn', 'CalculationItem', 'Measure')
    }
}

Describe 'Dataflow Detail.xlsx content (brief section 14)' -Skip:(-not $script:HasImportExcel) {
    It 'Sheet1 has the six monolith columns then the five DataTable artefact columns, with rows from the extracts' {
        $rows = @(Import-Excel -Path $script:DfPath -WorksheetName 'Sheet1')
        $rows.Count | Should -Be 6
        (@($rows[0].PSObject.Properties.Name) | Select-Object -First 11) -join ',' | Should -Be 'Dataflow ID,Dataflow Name,Query Name,Query,Report Date,Workspace Name - Dataflow Name,RowError,RowState,Table,ItemArray,HasErrors'
        $rows[0].'Workspace Name - Dataflow Name' | Should -Be 'Finance (Prod) ~ Sales (Prod) Data'
        $rows[0].'Report Date' | Should -Match '^2026-09-04'
    }
}

Describe 'Assemble on an empty run state' -Skip:(-not $script:HasImportExcel) {
    BeforeAll {
        $script:Base2 = Initialize-IQTestContext -RunId 'empty-run' -Prefix 'asm-empty'
        $script:Result2 = Invoke-IQAssembleStage
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base2 }
    It 'still creates every contract sheet header-only (a missing sheet is fatal for the PBIT)' {
        $script:Result2.Built | Should -Be 4
        $script:Result2.Failed | Should -Be 0
        foreach ($wb in $script:Contract.workbooks.PSObject.Properties) {
            $path = Join-Path $script:Base2 $wb.Name
            $path | Should -Exist
            $sheets = @(Get-ExcelSheetInfo -Path $path | ForEach-Object { $_.Name })
            foreach ($sheet in $wb.Value.PSObject.Properties) {
                $sheets | Should -Contain $sheet.Name
                $hdr = @(Get-IQTestSheetHeader -Path $path -Sheet $sheet.Name)
                @($sheet.Value.expectedColumns | Where-Object { $_ -notin $hdr }) | Should -BeNullOrEmpty
            }
        }
        (Get-IQTestSheetHeader -Path (Join-Path $script:Base2 'Dataflow Detail.xlsx') -Sheet 'Sheet1') -join ',' | Should -Be 'Dataflow ID,Dataflow Name,Query Name,Query,Report Date,Workspace Name - Dataflow Name,RowError,RowState,Table,ItemArray,HasErrors'
    }
    It 'rebuilding over existing workbooks works (temp file then move)' {
        $r = Invoke-IQAssembleStage
        $r.Built | Should -Be 4
        @(Get-ChildItem -LiteralPath $script:Base2 -Recurse -Filter '*.tmp*').Count | Should -Be 0
    }
}

Describe 'Report Detail.xlsx: header-only non-contract txt files keep their header (ASM-04)' -Skip:(-not $script:HasImportExcel) {
    It 'ReportExports.txt with only a header row becomes a ReportExports sheet with those columns' {
        @(Get-ExcelSheetInfo -Path $script:RepPath | Where-Object { $_.Name -eq 'ReportExports' }).Count | Should -Be 1
        (Get-IQTestSheetHeader -Path $script:RepPath -Sheet 'ReportExports') -join ',' | Should -Be 'ReportName,ReportID,ExportStatus,ExportMessage'
    }
    It 'the stale temp workbook of an earlier run was removed by the rebuild (ASM-06)' {
        (Join-Path $script:Base 'Power BI Environment Detail.tmp-deadbeef.xlsx') | Should -Not -Exist
    }
}

Describe 'ConvertTo-IQSheetTable typing, fast paths and Excel limits (ASM-01, ASM-03)' {
    It 'types columns from the values, keeps strings verbatim and fills missing/null cells (union of properties)' {
        $rows = @(
            [pscustomobject]@{ Id = '007'; Count = 1; Flag = $true; When = [datetime]'2026-01-02T03:04:05'; Nested = @{ a = 1 }; Note = $null }
            @{ Id = '=SUM(1)'; Count = 2.5; Flag = $false; Extra = 'x' }
        )
        $t = ConvertTo-IQSheetTable -Rows $rows -SheetName 'T'
        @($t.Columns | ForEach-Object { $_.ColumnName }) -join ',' | Should -Be 'Id,Count,Flag,When,Nested,Note,Extra'
        $t.Columns['Id'].DataType | Should -Be ([string])
        $t.Columns['Count'].DataType | Should -Be ([double])     # int + real -> Real
        $t.Columns['Flag'].DataType | Should -Be ([bool])
        $t.Columns['When'].DataType | Should -Be ([datetime])
        $t.Rows[0]['Id'] | Should -Be '007'
        $t.Rows[1]['Id'] | Should -Be '=SUM(1)'
        $t.Rows[0]['Nested'] | Should -Be '{"a":1}'
        $t.Rows[0]['Note'] | Should -Be ''                       # explicit $null in a text column -> empty string
        $t.Rows[0]['Extra'] | Should -Be ''                      # property absent from the row -> empty string
        $t.Rows[1]['When'] | Should -Be ([System.DBNull]::Value)  # absent from a typed column -> DBNull
        $t.Rows[1]['Count'] | Should -Be 2.5
    }
    It 'a column mixing strings and numbers becomes text with invariant-culture numbers' {
        $t = ConvertTo-IQSheetTable -Rows @([pscustomobject]@{ V = 1.5 }, [pscustomobject]@{ V = 'a' }) -SheetName 'T'
        $t.Columns['V'].DataType | Should -Be ([string])
        $t.Rows[0]['V'] | Should -Be '1.5'
    }
    It 'truncates cells beyond 32,767 characters and counts them' {
        $t = ConvertTo-IQSheetTable -Rows @([pscustomobject]@{ V = ('x' * 40000) }) -SheetName 'T'
        ([string]$t.Rows[0]['V']).Length | Should -Be 32767
        [int]$t.ExtendedProperties['Truncated'] | Should -Be 1
    }
    It 'never exceeds the Excel row limit: rows beyond the cap are dropped, counted and the rest is kept' {
        $saved = $script:IQExcelMaxDataRows
        try {
            $script:IQExcelMaxDataRows = 3
            $rows = @(1..5 | ForEach-Object { [pscustomobject]@{ N = $_ } })
            $t = ConvertTo-IQSheetTable -Rows $rows -SheetName 'T'
            $t.Rows.Count | Should -Be 3
            [int]$t.ExtendedProperties['TruncatedRows'] | Should -Be 2
            [int]$t.ExtendedProperties['SourceRows'] | Should -Be 5
        }
        finally { $script:IQExcelMaxDataRows = $saved }
        $script:IQExcelMaxDataRows | Should -Be 1048575
    }
}

Describe 'Assemble redacts secrets in the Message / Error columns (ASM-09)' {
    BeforeAll {
        if (Get-Command -Name 'Get-IQRunSummary' -ErrorAction SilentlyContinue) { Mock Get-IQRunSummary { throw 'forced local RunSummary path' } }
    }
    It 'InventoryErrors.Message masks connection-string passwords' {
        $g = [pscustomobject]@{ Errors = @([pscustomobject]@{ Collector = 'datasources'; Path = 'datasets/x/datasources'; Message = 'Server=srv;Password=hunter2;Timeout=30' }) }
        $rows = @(Get-IQAssembleInventoryErrorRow -Global $g -WorkspaceInventories @())
        $rows.Count | Should -Be 1
        $rows[0].Message | Should -Match 'Password=\*\*\*'
        $rows[0].Message | Should -Not -Match 'hunter2'
    }
    It 'Failures.Message masks bearer tokens' {
        $m = @{ runId = 'r'; failures = @(@{ stage = 'S'; itemKey = 'k'; item = 'i'; message = 'HTTP 401 with Bearer eyJabc.def.ghi'; timeUtc = 't' }) }
        $rows = @(Get-IQAssembleFailureRow -Manifest $m)
        $rows[0].Message | Should -Match 'Bearer \*\*\*'
        $rows[0].Message | Should -Not -Match 'eyJabc'
    }
    It 'RunSummary.Error masks client secrets' {
        $m = @{ runId = 'r'; status = 'Failed'; stages = @{ Inventory = @{ status = 'Failed'; itemsDone = 0; itemsFailed = 0; error = 'client_secret=abc123 rejected' } }; failures = @() }
        $rows = @(Get-IQAssembleRunSummaryRow -Manifest $m)
        $inv = @($rows | Where-Object { $_.Stage -eq 'Inventory' })[0]
        $inv.Error | Should -Match 'client_secret=\*\*\*'
        $inv.Error | Should -Not -Match 'abc123'
    }
}

Describe 'Write-IQWorkbook: per-sheet isolation, literal paths, stale temp files (ASM-01, ASM-06, ASM-07)' -Skip:(-not $script:HasImportExcel) {
    BeforeAll {
        # Two folders: one whose name contains [ ] (the module must write/copy there with literal paths) and a plain one
        # for reading content back (ImportExcel's own Import-Excel/Get-ExcelSheetInfo -Path glob '[' / ']').
        $script:WbFolder = Join-Path ([System.IO.Path]::GetTempPath()) ('ImpactIQ-tests/asm-wb-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + ' [PROD]')
        $script:WbPlain = Join-Path ([System.IO.Path]::GetTempPath()) ('ImpactIQ-tests/asm-wb-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '-plain')
        New-Item -ItemType Directory -Path $script:WbFolder -Force | Out-Null
        New-Item -ItemType Directory -Path $script:WbPlain -Force | Out-Null
        # The sheet 'Bad' fails in EPPlus when it carries its data rows; the header-only retry (one placeholder row) succeeds.
        Mock Export-Excel -ParameterFilter { $WorksheetName -eq 'Bad' -and $null -ne $InputObject -and $InputObject.Rows.Count -ge 2 } { throw 'EPPlus: simulated LoadFromDataTable failure' }
    }
    AfterAll { Remove-IQTestFolder -Path $script:WbFolder; Remove-IQTestFolder -Path $script:WbPlain }
    It 'writes into a folder whose name contains [ ] and removes a stale temp file first' {
        $stale = Join-Path $script:WbFolder 'Book.tmp-0badf00d.xlsx'
        [System.IO.File]::WriteAllText($stale, 'stale')
        $sheets = [ordered]@{}
        $sheets['Good'] = ConvertTo-IQSheetTable -Rows @([pscustomobject]@{ A = 1 }) -SheetName 'Good'
        $target = Join-Path $script:WbFolder 'Book.xlsx'
        $r = Write-IQWorkbook -Path $target -Sheets $sheets
        [System.IO.File]::Exists($target) | Should -BeTrue
        [System.IO.File]::Exists($stale) | Should -BeFalse
        @([System.IO.Directory]::GetFiles($script:WbFolder, '*.tmp-*')).Count | Should -Be 0
        @($r.FailedSheets).Count | Should -Be 0
        $readable = Join-Path $script:WbPlain 'Book.xlsx'
        [System.IO.File]::Copy($target, $readable, $true)
        @(Import-Excel -Path $readable -WorksheetName 'Good').Count | Should -Be 1
    }
    It 'one failing sheet is written header-only, the other sheets survive and the failure is reported' {
        $sheets = [ordered]@{}
        $sheets['Good'] = ConvertTo-IQSheetTable -Rows @([pscustomobject]@{ A = 1 }, [pscustomobject]@{ A = 2 }) -SheetName 'Good'
        $sheets['Bad'] = ConvertTo-IQSheetTable -Rows @([pscustomobject]@{ B = 'x' }, [pscustomobject]@{ B = 'y' }) -SheetName 'Bad'
        $sheets['Also'] = ConvertTo-IQSheetTable -Rows @([pscustomobject]@{ C = $true }) -SheetName 'Also'
        $target = Join-Path $script:WbPlain 'Partial.xlsx'
        $r = Write-IQWorkbook -Path $target -Sheets $sheets
        @($r.FailedSheets) | Should -Be @('Bad')
        ($r.Sheets -join ',') | Should -Be 'Good,Bad,Also'
        @(Get-ExcelSheetInfo -Path $target | ForEach-Object { $_.Name }) -join ',' | Should -Be 'Good,Bad,Also'
        @(Import-Excel -Path $target -WorksheetName 'Good').Count | Should -Be 2
        @(Import-Excel -Path $target -WorksheetName 'Bad' -WarningAction SilentlyContinue).Count | Should -Be 0
        (Get-IQTestSheetHeader -Path $target -Sheet 'Bad') | Should -Be @('B')
        @(Import-Excel -Path $target -WorksheetName 'Also').Count | Should -Be 1
    }
    It 'Build-IQDataflowWorkbook copies the workbook into a [ ] folder with the .NET API' {
        $runCopy = Join-Path $script:WbFolder 'run'
        $r = Build-IQDataflowWorkbook -Path (Join-Path $runCopy 'Dataflow Detail.xlsx') -CopyPath (Join-Path $script:WbFolder 'Dataflow Detail.xlsx') -ExtractFolder (Join-Path $script:WbFolder 'no-such-extracts') -Contract $script:Contract
        [System.IO.File]::Exists((Join-Path $runCopy 'Dataflow Detail.xlsx')) | Should -BeTrue
        [System.IO.File]::Exists((Join-Path $script:WbFolder 'Dataflow Detail.xlsx')) | Should -BeTrue
        $r.CopyPath | Should -Be (Join-Path $script:WbFolder 'Dataflow Detail.xlsx')
    }
}

Describe 'Assemble records unreadable source files and never reads another run''s folder (ASM-02, ASM-08)' -Skip:(-not $script:HasImportExcel) {
    BeforeAll {
        $script:Base3 = Initialize-IQTestContext -RunId 'src-err' -Prefix 'asm-src'
        $inv = $script:IQ.RunPaths.Inventory
        New-Item -ItemType Directory -Path $inv -Force | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $inv 'extras-scandatasets.json'), '{ "SheetName": "ScanDatasets", this is not json')
        $df = Join-Path $script:IQ.RunPaths.Extracts 'dataflows'
        New-Item -ItemType Directory -Path $df -Force | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $df 'broken.json'), '{ "Queries": [ {')
        # Another run's dated folders exist; this run's Model / Dataflow folders were deleted after Initialize-IQRun.
        $script:OtherModel = Join-Path $script:IQ.Paths.ModelBackups '2020-01-01'
        New-Item -ItemType Directory -Path $script:OtherModel -Force | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $script:OtherModel 'Old ~ Model.csv'), "ModelName,Type`r`nOld Model,Table`r`n")
        $script:OtherDataflow = Join-Path $script:IQ.Paths.DataflowBackups '2020-01-01'
        New-Item -ItemType Directory -Path $script:OtherDataflow -Force | Out-Null
        Remove-Item -LiteralPath $script:IQ.RunPaths.ModelBackups -Recurse -Force
        Remove-Item -LiteralPath $script:IQ.RunPaths.DataflowBackups -Recurse -Force
        $script:Result3 = $null
        $script:Status3 = Invoke-IQStage -Name Assemble -Body { $script:Result3 = Invoke-IQAssembleStage }
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base3 }
    It 'builds every workbook but ends CompletedWithErrors with one failed "source-<file>" item per unreadable file' {
        $script:Status3 | Should -Be 'CompletedWithErrors'
        $script:Result3.Built | Should -Be 4
        $script:Result3.Failed | Should -Be 0
        $script:Result3.SourceErrors | Should -Be 2
        $m = ConvertFrom-IQJsonFile -Path $script:IQ.RunPaths.Manifest
        [int]$m.stages.Assemble.itemsDone | Should -Be 4
        [int]$m.stages.Assemble.itemsFailed | Should -Be 2
        $f = @($m.failures | Where-Object { $_.stage -eq 'Assemble' })
        $f.Count | Should -Be 2
        @($f | Where-Object { $_.itemKey -like 'source-*' }).Count | Should -Be 2
        @($f | Where-Object { $_.item -like '*extras-scandatasets.json' }).Count | Should -Be 1
        @($f | Where-Object { $_.item -like '*broken.json' }).Count | Should -Be 1
        (Join-Path $script:Base3 'Power BI Environment Detail.xlsx') | Should -Exist
    }
    It 'does not fall back to another run''s Model Backups folder while a run is active' {
        Get-IQAssembleFolder -Kind 'Model' | Should -BeNullOrEmpty
        @(Import-Excel -Path (Join-Path $script:Base3 'Model Detail.xlsx') -WorksheetName 'Semantic Models' -WarningAction SilentlyContinue).Count | Should -Be 0
    }
    It 'writes the Dataflow workbook into this run''s (re-created) Dataflow Backups folder, not the other run''s' {
        (Join-Path $script:IQ.RunPaths.DataflowBackups 'Dataflow Detail.xlsx') | Should -Exist
        (Join-Path $script:OtherDataflow 'Dataflow Detail.xlsx') | Should -Not -Exist
        $m = ConvertFrom-IQJsonFile -Path $script:IQ.RunPaths.Manifest
        $m.outputs.dataflowWorkbookInRunFolder | Should -Be (Join-Path $script:IQ.RunPaths.DataflowBackups 'Dataflow Detail.xlsx')
    }
}

Describe 'Get-IQAssembleFolder without an active run keeps the newest-dated-folder rule (ASM-08)' {
    BeforeAll {
        $script:Base4 = Initialize-IQTestContext -NoRun -Prefix 'asm-norun'
        $script:Dated = Join-Path $script:IQ.Paths.ReportBackups '2021-05-06'
        New-Item -ItemType Directory -Path $script:Dated -Force | Out-Null
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base4 }
    It 'returns the newest yyyy-MM-dd folder when no run was initialised' {
        Get-IQAssembleFolder -Kind 'Report' | Should -Be $script:Dated
    }
}
