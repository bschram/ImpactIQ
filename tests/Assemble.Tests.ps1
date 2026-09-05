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
