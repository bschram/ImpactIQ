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

Describe 'Gen1 model.json helpers (entity rows, query groups) and Referenced Queries' {
    It 'Incremental Refresh comes from the refresh policy TYPE (IncrementalRefreshPolicy), never from its mere presence' {
        $rows = @(ConvertTo-IQDataflowEntityRow -ModelJson $script:Gen1Model -DataflowId $script:Ids.df1 -DataflowName 'Sales [Prod]/Data' -WorkspaceName 'Finance [Prod]')
        $rows.Count | Should -Be 3
        $sales = @($rows | Where-Object { $_.Entity -eq 'Sales' })
        $sales.Count | Should -Be 2
        foreach ($r in $sales) {
            $r.'Incremental Refresh' | Should -BeTrue
            $r.'Refresh Policy Type' | Should -Be 'IncrementalRefreshPolicy'
        }
        $empty = @($rows | Where-Object { $_.Entity -eq 'Empty' })[0]
        $empty.'Incremental Refresh' | Should -BeFalse -Because 'FullRefreshPolicy is what the service emits for every non-incremental entity'
        $empty.'Refresh Policy Type' | Should -Be 'FullRefreshPolicy'
        $empty.Column | Should -Be ''
    }
    It 'a FullRefreshPolicy entity and an entity without a policy are both not incremental; incrementalPeriods alone still counts' {
        $model = [pscustomobject]@{
            entities = @(
                [pscustomobject]@{ '$type' = 'LocalEntity'; name = 'Plain'; attributes = @(); 'pbi:refreshPolicy' = [pscustomobject]@{ '$type' = 'FullRefreshPolicy'; location = 'Plain.csv' } },
                [pscustomobject]@{ '$type' = 'LocalEntity'; name = 'NoPolicy'; attributes = @() },
                [pscustomobject]@{ '$type' = 'LocalEntity'; name = 'Periods'; attributes = @(); 'pbi:refreshPolicy' = [pscustomobject]@{ incrementalPeriods = 12; incrementalGranularity = 'Day' } }
            )
        }
        $rows = @(ConvertTo-IQDataflowEntityRow -ModelJson $model -DataflowId 'x' -DataflowName 'x' -WorkspaceName 'x')
        @($rows | Where-Object { $_.Entity -eq 'Plain' })[0].'Incremental Refresh' | Should -BeFalse
        @($rows | Where-Object { $_.Entity -eq 'NoPolicy' })[0].'Incremental Refresh' | Should -BeFalse
        @($rows | Where-Object { $_.Entity -eq 'NoPolicy' })[0].'Refresh Policy Type' | Should -Be ''
        @($rows | Where-Object { $_.Entity -eq 'Periods' })[0].'Incremental Refresh' | Should -BeTrue
    }
    It 'query groups come from the "pbi:QueryGroups" annotation (JSON string); a top-level member wins; bad JSON yields none' {
        $fromAnnotation = @(Get-IQDataflowGen1QueryGroup -ModelJson $script:Gen1Model)
        $fromAnnotation.Count | Should -Be 1
        $fromAnnotation[0].name | Should -Be 'Facts'
        $topLevel = [pscustomobject]@{
            'pbi:QueryGroups' = @([pscustomobject]@{ id = 'g1'; name = 'Top' })
            annotations       = @([pscustomobject]@{ name = 'pbi:QueryGroups'; value = '[{"id":"g1","name":"Annotation"}]' })
        }
        @(Get-IQDataflowGen1QueryGroup -ModelJson $topLevel)[0].name | Should -Be 'Top'
        $bad = [pscustomobject]@{ annotations = @([pscustomobject]@{ name = 'pbi:QueryGroups'; value = 'not json [' }, [pscustomobject]@{ name = 'other'; value = '[]' }) }
        @(Get-IQDataflowGen1QueryGroup -ModelJson $bad).Count | Should -Be 0
        @(Get-IQDataflowGen1QueryGroup -ModelJson $null).Count | Should -Be 0
    }
    It '"Query Group" shows the group NAME resolved through the annotation, not the group id' {
        $map = Get-IQDataflowQueryMetaMap -QueriesMetadata $script:Gen1Model.'pbi:mashup'.queriesMetadata -QueryGroups (Get-IQDataflowGen1QueryGroup -ModelJson $script:Gen1Model)
        $map['Sales'].QueryGroup | Should -Be 'Facts'
        $rows = @(ConvertFrom-IQDataflowDocument -Content $script:Gen1Document -DataflowId $script:Ids.df1 -DataflowName 'Sales [Prod]/Data' -WorkspaceName 'Finance [Prod]' -ReportDate '2026-09-04' -QueryMetadata $map)
        $sales = @($rows | Where-Object { $_.'Query Name' -eq 'Sales' })[0]
        $sales.'Query Group' | Should -Be 'Facts'
        $sales.'Load Enabled' | Should -BeTrue
    }
    It 'Referenced Queries ignores identifiers the expression declares itself (Source / Data let-steps) and keeps real references' {
        $doc = "section Section1;`nshared Source = Sql.Database(`"srv`", `"db`");`nshared Sales = let`n    Source = Sql.Database(`"srv`", `"db`"),`n    Data = Source{[Schema = `"dbo`", Item = `"Sales`"]}[Data]`nin`n    Data;`nshared Data = let Pattern = Sales in Pattern;`nshared #`"Sales Copy`" = let Source = #`"Sales`" in Source;`nshared Single = let Data = 1 in Data;"
        $rows = @(ConvertFrom-IQDataflowDocument -Content $doc -DataflowId 'x' -DataflowName 'x' -WorkspaceName 'x' -ReportDate '2026-09-04')
        $byName = @{}
        foreach ($r in $rows) { $byName[$r.'Query Name'] = $r }
        $byName['Sales'].'Referenced Queries' | Should -Be '' -Because 'Source and Data are local steps of Sales, not the shared queries of the same name'
        $byName['Data'].'Referenced Queries' | Should -Be 'Sales'
        $byName['Sales Copy'].'Referenced Queries' | Should -Be 'Sales' -Because '#"Sales" is the same identifier as Sales; the local Source step is not a reference'
        $byName['Single'].'Referenced Queries' | Should -Be '' -Because 'a single-line "let Data = ..." declares Data'
        $byName['Source'].'Referenced Queries' | Should -Be ''
        # the fixture: "Customers 2024" has "Source = Sales" -> Sales is referenced; Sales itself references nothing
        $fixtureRows = @(ConvertFrom-IQDataflowDocument -Content $script:Gen1Document -DataflowId 'x' -DataflowName 'x' -WorkspaceName 'x' -ReportDate '2026-09-04')
        @($fixtureRows | Where-Object { $_.'Query Name' -eq 'Customers 2024' })[0].'Referenced Queries' | Should -Be 'Sales'
        @($fixtureRows | Where-Object { $_.'Query Name' -eq 'Sales' })[0].'Referenced Queries' | Should -Be ''
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
        $definition = Join-Path $script:RunFolder 'Sales ~ Sales Gen2.definition'
        (Join-Path $definition 'mashup.pq') | Should -Exist
        (Join-Path $definition 'queryMetadata.json') | Should -Exist
        @((Get-IQItemCheckpoint -Stage Dataflows -ItemKey $script:Ids.fdf1).outputs) | Should -Contain $definition -Because 'a deleted definition folder must invalidate the checkpoint'
        @(Get-ChildItem -LiteralPath $script:RunFolder -Directory | Where-Object { $_.Name -like '*.tmp-*' }).Count | Should -Be 0 -Because 'the temp folder is swapped into place'
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
        @($extA.Queries)[0].'Query Group' | Should -Be 'Facts' -Because 'the group name comes from the pbi:QueryGroups annotation'
        [int]$extA.EntityCount | Should -Be 3
        @($extA.Entities | Where-Object { $_.Entity -eq 'Empty' })[0].'Incremental Refresh' | Should -BeFalse
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
    It 'Report Date for a non-date RunId is the run start date from the manifest (one date per run), not today' {
        $savedRunId = $script:IQ.RunId
        $savedStart = $script:IQ.Manifest['startedUtc']
        try {
            $script:IQ.RunId = 'nightly-42'
            $script:IQ.Manifest['startedUtc'] = '2026-01-15T12:00:00.0000000Z'
            Get-IQDataflowReportDate | Should -Be '2026-01-15'
            $script:IQ.RunId = '2026-09-04'
            Get-IQDataflowReportDate | Should -Be '2026-09-04'
        }
        finally {
            $script:IQ.RunId = $savedRunId
            $script:IQ.Manifest['startedUtc'] = $savedStart
        }
    }
    It 'backup names: a Gen1 and a Gen2 CI/CD dataflow with the same name keep "<stem>.txt" / "<stem>.pq"; a second Gen1 gets the " ~<id8>" suffix' {
        Mock Get-IQAllWorkspaceInventories {
            @([pscustomobject]@{
                    WorkspaceId = 'ws-x'; WorkspaceName = 'WS'; IsSynthetic = $false
                    Dataflows   = @(
                        [pscustomobject]@{ DataflowId = '11111111-aaaa-4aaa-8aaa-000000000001'; DataflowName = 'Sales'; DataflowGeneration = 1; WorkspaceId = 'ws-x'; WorkspaceName = 'WS' },
                        [pscustomobject]@{ DataflowId = '22222222-bbbb-4bbb-8bbb-000000000002'; DataflowName = 'Sales'; DataflowGeneration = 'Gen 2 CICD'; WorkspaceId = 'ws-x'; WorkspaceName = 'WS' },
                        [pscustomobject]@{ DataflowId = '33333333-cccc-4ccc-8ccc-000000000003'; DataflowName = 'Sales'; DataflowGeneration = 1; WorkspaceId = 'ws-x'; WorkspaceName = 'WS' }
                    )
                })
        }
        $list = @(Get-IQDataflowWorkList -RunFolder $script:RunFolder -ExtractFolder $script:ExtractFolder)
        @($list | ForEach-Object { Split-Path -Path $_.BackupPath -Leaf }) | Should -Be @('WS ~ Sales.txt', 'WS ~ Sales.pq', 'WS ~ Sales ~33333333.txt')
        @($list | ForEach-Object { Split-Path -Path $_.DefinitionFolder -Leaf }) | Should -Be @('WS ~ Sales.definition', 'WS ~ Sales.definition', 'WS ~ Sales ~33333333.definition')
    }
    It 'Gen2: a re-export whose parts cannot be saved keeps the previous definition folder and .pq, leaves no temp folder, fails the checkpoint and removes the stale extract' {
        $definition = Join-Path $script:RunFolder 'Sales ~ Sales Gen2.definition'
        $pq = Join-Path $script:RunFolder 'Sales ~ Sales Gen2.pq'
        $extract = Join-Path $script:ExtractFolder ((Get-IQSafeKey -Value $script:Ids.fdf1) + '.json')
        (Join-Path $definition 'mashup.pq') | Should -Exist
        $extract | Should -Exist
        Remove-Item -LiteralPath $pq -Force   # invalidates the checkpoint so the dataflow is re-exported
        Mock Invoke-IQFabricLro {
            [pscustomobject]@{ definition = [pscustomobject]@{ parts = @([pscustomobject]@{ path = '../evil.txt'; payloadType = 'InlineBase64'; payload = 'eA==' }, [pscustomobject]@{ path = 'broken.pq'; payloadType = 'InlineBase64'; payload = '%%% not base64 %%%' }) } }
        }
        Invoke-IQStage -Name Dataflows -Body { Invoke-IQDataflowsStage | Out-Null } | Should -Be 'CompletedWithErrors'
        (Get-IQItemCheckpoint -Stage Dataflows -ItemKey $script:Ids.fdf1).status | Should -Be 'Failed'
        (Join-Path $definition 'mashup.pq') | Should -Exist -Because 'the previous definition is only replaced by a good export'
        (Join-Path $definition 'broken.pq') | Should -Not -Exist
        @(Get-ChildItem -LiteralPath $script:RunFolder -Directory | Where-Object { $_.Name -like '*.tmp-*' }).Count | Should -Be 0
        (Join-Path $script:RunFolder 'evil.txt') | Should -Not -Exist
        $extract | Should -Not -Exist -Because 'Assemble must not emit stale rows for a dataflow the manifest reports as failed'
        @($script:IQ.Manifest.failures | Where-Object { $_.itemKey -eq $script:Ids.fdf1 }).Count | Should -Be 1
    }
    It 'Gen1: a body that is not valid JSON is a failure (retried on resume) that keeps the raw .txt and removes the stale extract' {
        $txt = Join-Path $script:RunFolder 'Finance (Prod) ~ Sales (Prod) Data.txt'
        $extract = Join-Path $script:ExtractFolder ((Get-IQSafeKey -Value $script:Ids.df1) + '.json')
        $extract | Should -Exist
        Remove-Item -LiteralPath $txt -Force
        $script:IQTestApiOverrides[(ConvertTo-IQTestFixtureName -Api PowerBI -Path ('groups/' + $script:Ids.ws1 + '/dataflows/' + $script:Ids.df1))] = '{ "name": "Sales", "pbi:mashup": { "document": "section Section1;'
        try {
            Invoke-IQStage -Name Dataflows -Body { Invoke-IQDataflowsStage | Out-Null } | Should -Be 'CompletedWithErrors'
        }
        finally { $script:IQTestApiOverrides.Clear() }
        $cp = Get-IQItemCheckpoint -Stage Dataflows -ItemKey $script:Ids.df1
        $cp.status | Should -Be 'Failed'
        $cp.message | Should -Match 'not valid JSON'
        $txt | Should -Exist -Because 'the raw body is kept for inspection'
        $extract | Should -Not -Exist
        (Get-IQItemCheckpoint -Stage Dataflows -ItemKey $script:Ids.fdf1).status | Should -Be 'Succeeded' -Because 'the Gen2 dataflow that failed in the previous test is retried with the good definition'
        (Test-IQItemDone -Stage Dataflows -ItemKey $script:Ids.df1) | Should -BeFalse -Because 'a failed dataflow is exported again on resume'
        # and the fixed dataflow is picked up again on the next start
        Invoke-IQStage -Name Dataflows -Body { Invoke-IQDataflowsStage | Out-Null } | Should -Be 'Completed'
        (Get-IQItemCheckpoint -Stage Dataflows -ItemKey $script:Ids.df1).status | Should -Be 'Succeeded'
        $extract | Should -Exist
    }
}
