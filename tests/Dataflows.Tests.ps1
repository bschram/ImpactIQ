# Dataflows.Tests.ps1 - ImpactIQ.Dataflows.ps1 (brief section 8.3, audit C9): Gen1/Gen2 parsing and the stage.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Ids = Get-IQTestFixtureJson -Relative 'ids.json'
    $script:Gen1Model = Get-IQTestFixtureJson -Relative 'dataflows/gen1-model.json'
    $script:Gen1Document = [string]$script:Gen1Model.'pbi:mashup'.document
    $script:Gen2Pq = Get-IQTestFixtureText -Relative 'dataflows/gen2-mashup.pq'
    Mock Get-IQToken { 'test-token' }
    Mock Invoke-IQApi { Invoke-IQTestApiFixture -Method $Method -Path $Path -Body $Body -Query $Query -Api $Api -Raw:$Raw }
    Mock Invoke-IQFabricLro {
        $script:IQTestApiCalls.Add(('Fabric LRO {0} {1}' -f $Method, $Path))
        if ($Path -like ('*dataflows/' + $script:Ids.fdf1 + '/getDefinition')) { return (Get-IQTestGen2Definition) }
        return $null
    }
}
AfterAll { Remove-IQTestFolder -Path $script:Base }

Describe 'ConvertFrom-IQDataflowDocument (Gen1 mashup document)' {
    BeforeAll {
        $script:Rows = @(ConvertFrom-IQDataflowDocument -Content $script:Gen1Document -DataflowId $script:Ids.df1 -DataflowName 'Sales [Prod]/Data' -WorkspaceName 'Finance [Prod]' -ReportDate '2026-09-04')
        $script:Names = @($script:Rows | ForEach-Object { $_.'Query Name' })
    }
    It 'finds every shared member (6) and no false split on "shared" inside a string' {
        $script:Rows.Count | Should -Be 6
        $script:Names | Should -Contain 'Sales'
        $script:Names | Should -Not -Contain 'Foo'
    }
    It 'handles #"names with spaces", dotted and non-ASCII identifiers and doubled quotes' {
        $script:Names | Should -Contain 'Customers 2024'
        $script:Names | Should -Contain 'Region.Lookup'
        $script:Names | Should -Contain 'Ärzte'
        $script:Names | Should -Contain 'Say "Hi"'
    }
    It 'keeps nested let bodies intact and strips the trailing semicolon' {
        $sales = @($script:Rows | Where-Object { $_.'Query Name' -eq 'Sales' })[0]
        $sales.Query | Should -Match '^let'
        $sales.Query | Should -Match 'Sql\.Database\("sql-finance\.contoso\.gov", "Ledger"\)'
        $sales.Query | Should -Match '\[Data\]\s*\r?\nin\s*\r?\n\s*T$'
        $sales.Query | Should -Not -Match ';\s*$'
        $cust = @($script:Rows | Where-Object { $_.'Query Name' -eq 'Customers 2024' })[0]
        $cust.Query | Should -Match 'see shared Foo = bar'
    }
    It 'peels an attribute record in brackets off the previous body and keeps backslashes untouched' {
        $files = @($script:Rows | Where-Object { $_.'Query Name' -eq 'Files' })[0]
        $files.Query | Should -Match 'Folder\.Files\("\\\\fileserver\\new reports\\"\)'
        $files.Query | Should -Match 'C:\\rn\\data\\'
        $arzte = @($script:Rows | Where-Object { $_.'Query Name' -eq 'Ärzte' })[0]
        $arzte.Query | Should -Not -Match 'Description = "nested'
        if ($files.PSObject.Properties['Attributes']) { $files.Attributes | Should -Match 'nested \[bracket\] record' }
    }
    It 'emits the six monolith Sheet1 columns first, in order, with the cleaned "Workspace Name - Dataflow Name"' {
        $cols = @($script:Rows[0].PSObject.Properties.Name)
        ($cols | Select-Object -First 6) -join ',' | Should -Be 'Dataflow ID,Dataflow Name,Query Name,Query,Report Date,Workspace Name - Dataflow Name'
        $script:Rows[0].'Dataflow ID' | Should -Be $script:Ids.df1
        $script:Rows[0].'Dataflow Name' | Should -Be 'Sales [Prod]/Data'
        $script:Rows[0].'Report Date' | Should -Be '2026-09-04'
        $script:Rows[0].'Workspace Name - Dataflow Name' | Should -Be 'Finance (Prod) ~ Sales (Prod) Data'
    }
}

Describe 'ConvertFrom-IQDataflowDocument (Gen2 .pq with attribute records)' {
    BeforeAll {
        $script:G2 = @(ConvertFrom-IQDataflowDocument -Content $script:Gen2Pq -DataflowId $script:Ids.fdf1 -DataflowName 'Sales Gen2' -WorkspaceName 'Sales' -ReportDate '2026-09-04')
    }
    It 'finds the 3 shared queries and attaches the [DataDestinations = ...] record to its own query' {
        $script:G2.Count | Should -Be 3
        @($script:G2 | ForEach-Object { $_.'Query Name' }) | Should -Be @('Sales', 'Customers 2024', 'Sales_DataDestination')
        $sales = $script:G2[0]
        $sales.Query | Should -Not -Match 'DataDestinations'
        $sales.Query | Should -Not -Match ';\s*$'
        if ($sales.PSObject.Properties['Attributes']) { $sales.Attributes | Should -Match '^\[DataDestinations' }
        $script:G2[1].Query | Should -Match '^let\s*\r?\n\s*Source = Sales\s*\r?\nin\s*\r?\n\s*Source$'
    }
    It 'handles a single-line let ... in query with a trailing semicolon' {
        $script:G2[2].Query | Should -Be 'let Pattern = Lakehouse.Contents([HierarchicalNavigation = null]) in Pattern'
    }
}

Describe 'Invoke-IQDataflowsStage' {
    BeforeAll {
        $script:Base = Initialize-IQTestContext -Options @{ RunMode = 'Workspaces'; WorkspaceId = @($script:Ids.ws1, $script:Ids.ws2) } -Prefix 'df'
        Invoke-IQStage -Name Inventory -Body { Invoke-IQInventoryStage | Out-Null } | Should -Be 'Completed'
        # add a Gen1 dataflow whose export fails (the API answers $null) to the ws1 inventory
        $script:BrokenId = '30000000-0000-4000-8000-0000000000ff'
        $w1 = Get-IQInventory -Name ('ws-' + $script:Ids.ws1)
        $broken = [pscustomobject]@{ DataflowId = $script:BrokenId; DataflowName = 'Broken Flow'; DataflowGeneration = 1; WorkspaceId = $script:Ids.ws1; WorkspaceName = 'Finance [Prod]' }
        $w1.Dataflows = @(@($w1.Dataflows) + @($broken))
        Save-IQInventory -Name ('ws-' + $script:Ids.ws1) -Object $w1 | Out-Null
        $script:IQTestApiOverrides[(ConvertTo-IQTestFixtureName -Api PowerBI -Path ('groups/' + $script:Ids.ws1 + '/dataflows/' + $script:BrokenId))] = $null
        $script:IQTestApiCalls.Clear()
        $script:Status = Invoke-IQStage -Name Dataflows -Body { Invoke-IQDataflowsStage | Out-Null }
        $script:RunFolder = $script:IQ.RunPaths.DataflowBackups
        $script:ExtractFolder = Join-Path $script:IQ.RunPaths.Extracts 'dataflows'
    }
    AfterAll { $script:IQTestApiOverrides.Clear() }
    It 'completes with errors (one broken dataflow) and checkpoints each dataflow by DataflowId' {
        $script:Status | Should -Be 'CompletedWithErrors'
        (Get-IQItemCheckpoint -Stage Dataflows -ItemKey $script:Ids.df1).status | Should -Be 'Succeeded'
        (Get-IQItemCheckpoint -Stage Dataflows -ItemKey $script:Ids.fdf1).status | Should -Be 'Succeeded'
        (Get-IQItemCheckpoint -Stage Dataflows -ItemKey $script:BrokenId).status | Should -Be 'Failed'
        @($script:IQ.Manifest.failures | Where-Object { $_.itemKey -eq $script:BrokenId }).Count | Should -Be 1
    }
    It 'writes the Gen1 backup as "<CleanWs> ~ <CleanDf>.txt" with the raw JSON body (UTF-8, no BOM)' {
        $txt = Join-Path $script:RunFolder 'Finance (Prod) ~ Sales (Prod) Data.txt'
        $txt | Should -Exist
        $bytes = [System.IO.File]::ReadAllBytes($txt)
        (($bytes[0] -eq 0xEF) -and ($bytes[1] -eq 0xBB)) | Should -BeFalse
        $parsed = ConvertFrom-Json -InputObject ([System.IO.File]::ReadAllText($txt))
        $parsed.'pbi:mashup'.document | Should -Match 'shared Sales'
        @((Get-IQItemCheckpoint -Stage Dataflows -ItemKey $script:Ids.df1).outputs) | Should -Contain $txt
    }
    It 'writes the Gen2 backup as "<CleanWs> ~ <CleanDf>.pq" (byte exact) via Invoke-IQFabricLro getDefinition' {
        $pq = Join-Path $script:RunFolder 'Sales ~ Sales Gen2.pq'
        $pq | Should -Exist
        [System.IO.File]::ReadAllBytes($pq).Length | Should -Be ([System.IO.File]::ReadAllBytes((Get-IQTestFixturePath -Relative 'dataflows/gen2-mashup.pq')).Length)
        @($script:IQTestApiCalls | Where-Object { $_ -like ('Fabric LRO POST workspaces/' + $script:Ids.ws2 + '/dataflows/' + $script:Ids.fdf1 + '/getDefinition') }).Count | Should -Be 1
        (Join-Path $script:RunFolder 'evil.txt') | Should -Not -Exist -Because 'definition parts outside the target folder are skipped'
    }
    It 'does not leave a stale or empty backup for the failed dataflow' {
        (Join-Path $script:RunFolder 'Finance (Prod) ~ Broken Flow.txt') | Should -Not -Exist
    }
    It 'saves one extract JSON per dataflow with the parsed queries' {
        $extA = ConvertFrom-IQJsonFile -Path (Join-Path $script:ExtractFolder ((Get-IQSafeKey -Value $script:Ids.df1) + '.json'))
        $extA | Should -Not -BeNullOrEmpty
        $extA.DataflowId | Should -Be $script:Ids.df1
        [int]$extA.QueryCount | Should -Be 6
        @($extA.Queries)[0].'Query Name' | Should -Be 'Sales'
        @($extA.Queries)[0].'Report Date' | Should -Be $script:IQ.RunId
        $extC = ConvertFrom-IQJsonFile -Path (Join-Path $script:ExtractFolder ((Get-IQSafeKey -Value $script:Ids.fdf1) + '.json'))
        [int]$extC.QueryCount | Should -Be 3
        $extC.Generation | Should -Match 'Gen 2'
    }
    It 'Get-IQDataflowExtractRow combines the extracts into Sheet1-shaped rows' {
        $rows = @(Get-IQDataflowExtractRow)
        $rows.Count | Should -Be 9
        (@($rows[0].PSObject.Properties.Name) | Select-Object -First 6) -join ',' | Should -Be 'Dataflow ID,Dataflow Name,Query Name,Query,Report Date,Workspace Name - Dataflow Name'
    }
    It 'skips pseudo workspaces (My Workspace) and never calls the API for them' {
        @($script:IQTestApiCalls | Where-Object { $_ -like '*My Workspace*' }).Count | Should -Be 0
    }
    It 'resume: checkpointed dataflows are not re-fetched, a missing output is re-exported, a fixed failure is retried' {
        $script:IQTestApiOverrides.Clear()
        $script:IQTestApiOverrides[(ConvertTo-IQTestFixtureName -Api PowerBI -Path ('groups/' + $script:Ids.ws1 + '/dataflows/' + $script:BrokenId))] = { Get-IQTestFixtureText -Relative 'dataflows/gen1-model.json' }
        $pq = Join-Path $script:RunFolder 'Sales ~ Sales Gen2.pq'
        Remove-Item -LiteralPath $pq -Force
        $script:IQTestApiCalls.Clear()
        Invoke-IQStage -Name Dataflows -Body { Invoke-IQDataflowsStage | Out-Null } | Should -Be 'Completed'
        @($script:IQTestApiCalls | Where-Object { $_ -like ('*dataflows/' + $script:Ids.df1) }).Count | Should -Be 0
        @($script:IQTestApiCalls | Where-Object { $_ -like ('*' + $script:Ids.fdf1 + '/getDefinition') }).Count | Should -Be 1
        $pq | Should -Exist
        (Get-IQItemCheckpoint -Stage Dataflows -ItemKey $script:BrokenId).status | Should -Be 'Succeeded'
        @($script:IQ.Manifest.failures | Where-Object { $_.itemKey -eq $script:BrokenId }).Count | Should -Be 0
    }
}
