# Inventory.Tests.ps1 - ImpactIQ.Inventory.ps1 (brief section 2.6, 5.3, 6) with Invoke-IQApi answered from tests/fixtures/api.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Ids = Get-IQTestFixtureJson -Relative 'ids.json'
    Mock Get-IQToken { 'test-token' }
    Mock Invoke-IQApi { Invoke-IQTestApiFixture -Method $Method -Path $Path -Body $Body -Query $Query -Api $Api -Raw:$Raw }
    function Read-WorkspaceFile { param([string]$Id) return (Get-IQInventory -Name ('ws-' + $Id)) }
    function Start-InventoryRun {
        param([hashtable]$Options)
        $script:Base = Initialize-IQTestContext -Options $Options -Prefix 'inv'
        $script:Summary = $null
        $script:StageStatus = Invoke-IQStage -Name Inventory -Body { $script:Summary = Invoke-IQInventoryStage }
    }
}
AfterAll { Remove-IQTestFolder -Path $script:Base }

Describe 'Get-IQWorkspaceList' {
    BeforeAll { $script:Base = Initialize-IQTestContext -Prefix 'inv-list' }
    AfterAll { Remove-IQTestFolder -Path $script:Base }
    It 'lists groups with $top=5000 paging and renames with the monolith workspace map' {
        $rows = @(Get-IQWorkspaceList)
        $rows.Count | Should -Be 4
        $first = @($rows | Where-Object { $_.WorkspaceId -eq $script:Ids.ws1 })[0]
        $first.WorkspaceName | Should -Be 'Finance [Prod]'
        $first.WorkspaceType | Should -Be 'Workspace'
        $first.WorkspaceIsReadOnly | Should -BeFalse
        $first.WorkspaceIsOnDedicatedCapacity | Should -BeTrue
        $first.WorkspaceCapacityId | Should -Be $script:Ids.cap1
        $first.WorkspaceDefaultDatasetStorageFormat | Should -Be 'Large'
        $first.WorkspaceState | Should -Be 'Active'
        @($script:IQTestApiCalls | Where-Object { $_ -like 'PowerBI GET groups' }).Count | Should -BeGreaterOrEqual 1
    }
}

Describe 'Workspaces mode (ids + wildcard name + My Workspace)' {
    BeforeAll {
        Start-InventoryRun -Options @{ RunMode = 'Workspaces'; WorkspaceId = @($script:Ids.ws1, $script:Ids.ws2); WorkspaceName = @('HR*'); IncludeMyWorkspace = $true }
        $script:W1 = Read-WorkspaceFile $script:Ids.ws1
        $script:W2 = Read-WorkspaceFile $script:Ids.ws2
        $script:W3 = Read-WorkspaceFile $script:Ids.ws3
        $script:My = Read-WorkspaceFile 'My Workspace'
        $script:Global = Get-IQInventory -Name 'global'
        $script:WsRows = @(Get-IQInventory -Name 'workspaces')
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base }
    It 'completes and collects the selected workspaces plus My Workspace' {
        $script:StageStatus | Should -Be 'Completed'
        $script:Summary.Collected | Should -Be 4
        (Join-Path (Join-Path $script:IQ.RunPath 'inventory') 'workspaces.json') | Should -Exist
        (Join-Path (Join-Path $script:IQ.RunPath 'inventory') 'global.json') | Should -Exist
        (Join-Path (Join-Path $script:IQ.RunPath 'inventory') ('ws-' + $script:Ids.ws1 + '.json')) | Should -Exist
    }
    It 'workspaces.json holds the 3 selected rows plus the My Workspace pseudo row with CapacityName joined' {
        @($script:WsRows | Where-Object { $_.WorkspaceId -eq $script:Ids.ws4 }).Count | Should -Be 0
        $fin = @($script:WsRows | Where-Object { $_.WorkspaceId -eq $script:Ids.ws1 })[0]
        $fin.CapacityName | Should -Be 'Premium P1'
        $my = @($script:WsRows | Where-Object { $_.WorkspaceId -eq 'My Workspace' })
        $my.Count | Should -Be 1
        $my[0].WorkspaceName | Should -Be 'My Workspace'
    }
    It 'ws-<id>.json has every collection key (existing + new)' {
        foreach ($key in @('WorkspaceId', 'WorkspaceName', 'Datasets', 'DatasetSources', 'DatasetRefreshHistory', 'DatasetRefreshSchedule', 'Reports', 'ReportPages', 'Dataflows', 'DataflowSources', 'DataflowLineage', 'DataflowRefreshHistory', 'FabricItems', 'ItemConnections', 'ReportsWithSensitivityLabel', 'Dashboards', 'DashboardTiles', 'WorkspaceUsers', 'DatasetUsers', 'DatasetParameters', 'DatasetDirectQueryRefreshSchedule')) {
            $script:W1.PSObject.Properties[$key] | Should -Not -BeNullOrEmpty -Because "ws file key $key"
        }
    }
    It 'Datasets rows use the monolith property names plus DatasetCapacityId' {
        $ds = @($script:W1.Datasets)
        $ds.Count | Should -Be 2
        $d1 = @($ds | Where-Object { $_.DatasetId -eq $script:Ids.d1 })[0]
        $d1.DatasetName | Should -Be 'Finance Model'
        $d1.DatasetWebUrl | Should -Match '^https://'
        $d1.DatasetConfiguredBy | Should -Be 'analyst@contoso.gov'
        $d1.DatasetIsRefreshable | Should -BeTrue
        $d1.DatasetTargetStorageMode | Should -Be 'Abf'
        $d1.DatasetCreatedDate | Should -Not -BeNullOrEmpty
        $d1.WorkspaceId | Should -Be $script:Ids.ws1
        $d1.WorkspaceName | Should -Be 'Finance [Prod]'
        $d1.DatasetCapacityId | Should -Be $script:Ids.cap1
        $d1.PSObject.Properties['id'] | Should -BeNullOrEmpty -Because 'renamed keys are removed'
    }
    It 'Reports rows carry ReportId/ReportName/DatasetId/DatasetWorkspaceId/ReportType/ReportIsFromPbix' {
        $r6 = @($script:W1.Reports | Where-Object { $_.ReportId -eq $script:Ids.r6 })[0]
        $r6.ReportName | Should -Be 'Cross Report'
        $r6.DatasetId | Should -Be $script:Ids.d3
        $r6.DatasetWorkspaceId | Should -Be $script:Ids.ws2
        $r6.ReportType | Should -Be 'PowerBIReport'
        $r6.ReportIsFromPbix | Should -BeTrue
        $r6.ReportWebUrl | Should -Match '^https://'
        $r6.WorkspaceName | Should -Be 'Finance [Prod]'
    }
    It 'ReportPages rows are renamed and carry the report id; no pages call is made for paginated reports' {
        $pages = @($script:W1.ReportPages)
        $pages.Count | Should -Be 3
        $p = @($pages | Where-Object { $_.ReportId -eq $script:Ids.r1 })
        $p.Count | Should -Be 2
        $p[0].PageName | Should -Match '^ReportSection'
        $p[0].PageDisplayName | Should -Be 'Overview'
        $p[0].PageOrder | Should -Be 0
        @($script:IQTestApiCalls | Where-Object { $_ -like ('*reports/' + $script:Ids.r4 + '/pages') }).Count | Should -Be 0
    }
    It 'DatasetSources, refresh history (with nested JSON flattened) and schedules are collected' {
        @($script:W1.DatasetSources).Count | Should -Be 1
        $src = @($script:W1.DatasetSources)[0]
        $src.DatasetDatasourceType | Should -Be 'Sql'
        $src.DatasetId | Should -Be $script:Ids.d1
        $hist = @($script:W1.DatasetRefreshHistory)
        $hist.Count | Should -Be 2
        $hist[0].DatasetRefreshStatus | Should -BeIn @('Completed', 'Failed')
        $hist[0].DatasetRefreshType | Should -Be 'Scheduled'
        ($hist | Where-Object { $null -ne $_.refreshAttempts } | Select-Object -First 1).refreshAttempts | Should -BeOfType [string]
        $sched = @($script:W1.DatasetRefreshSchedule | Where-Object { $_.DatasetId -eq $script:Ids.d1 })
        $sched.Count | Should -Be 4 -Because '2 days x 2 times'
        $sched[0].DatasetRefreshScheduleEnabled | Should -BeTrue
        $sched[0].DatasetRefreshScheduleLocalTimeZoneId | Should -Be 'UTC'
        $sched[0].DatasetRefreshScheduleNotifyOption | Should -Be 'MailOnFailure'
    }
    It 'DirectQuery refresh schedule rows use the DQ* columns' {
        $dq = @($script:W1.DatasetDirectQueryRefreshSchedule)
        $dq.Count | Should -BeGreaterOrEqual 1
        $dq[0].DQFrequency | Should -Be 15
        $dq[0].DQLocalTimeZoneId | Should -Be 'UTC'
        $dq[0].DatasetId | Should -Be $script:Ids.d2
        $dq[0].DatasetName | Should -Be 'DQ Model'
    }
    It 'Dataflows, sources, lineage and refresh history keep the monolith names; CICD dataflows come from Fabric' {
        $df = @($script:W1.Dataflows)
        $df.Count | Should -Be 1
        $df[0].DataflowId | Should -Be $script:Ids.df1
        $df[0].DataflowName | Should -Be 'Sales [Prod]/Data'
        $df[0].DataflowJsonURL | Should -Match 'model\.json$'
        $df[0].DataflowConfiguredBy | Should -Be 'analyst@contoso.gov'
        @($script:W1.DataflowSources)[0].DataflowDatasourceType | Should -Be 'Sql'
        @($script:W1.DataflowRefreshHistory)[0].DataflowRefreshStatus | Should -Be 'Success'
        $lineage = @($script:W2.DataflowLineage)
        $lineage.Count | Should -Be 1
        $lineage[0].DatasetId | Should -Be $script:Ids.d3
        $lineage[0].DataflowId | Should -Be $script:Ids.df1
        $lineage[0].WorkspaceId | Should -Be $script:Ids.ws2
        $cicd = @($script:W2.Dataflows)
        $cicd.Count | Should -Be 1
        $cicd[0].DataflowId | Should -Be $script:Ids.fdf1
        $cicd[0].DataflowName | Should -Be 'Sales Gen2'
        $cicd[0].DataflowGeneration | Should -Be 'Gen 2 CICD'
    }
    It 'FabricItems keep the monolith map and exclude Report/SemanticModel items; ItemConnections collected' {
        $items = @($script:W1.FabricItems)
        $items.Count | Should -Be 1
        $items[0].FabricItemID | Should -Be $script:Ids.nb1
        $items[0].FabricItemType | Should -Be 'Notebook'
        $items[0].FabricItemName | Should -Be 'Finance Notebook'
        @($script:W1.ItemConnections).Count | Should -Be 1
        @($script:W1.ReportsWithSensitivityLabel) | Should -Contain $script:Ids.r1
    }
    It 'NEW collectors: Dashboards, DashboardTiles, WorkspaceUsers (403 -> empty), DatasetUsers, DatasetParameters' {
        $db = @($script:W1.Dashboards)
        $db.Count | Should -Be 1
        $db[0].DashboardId | Should -Be $script:Ids.db1
        $db[0].DashboardName | Should -Be 'Executive Dashboard'
        $db[0].DashboardIsReadOnly | Should -BeFalse
        $db[0].WorkspaceName | Should -Be 'Finance [Prod]'
        $tiles = @($script:W1.DashboardTiles)
        $tiles.Count | Should -Be 1
        $tiles[0].TileTitle | Should -Be 'Revenue YTD'
        $tiles[0].TileRowSpan | Should -Be 2
        $tiles[0].ReportId | Should -Be $script:Ids.r1
        $tiles[0].DashboardId | Should -Be $script:Ids.db1
        $users = @($script:W1.WorkspaceUsers)
        $users.Count | Should -Be 2
        ($users | Where-Object { $_.UserPrincipalType -eq 'User' }).UserEmailAddress | Should -Be 'admin@contoso.gov'
        ($users | Where-Object { $_.UserPrincipalType -eq 'User' }).UserGroupUserAccessRight | Should -Be 'Admin'
        @($script:W2.WorkspaceUsers).Count | Should -Be 0 -Because 'a 403 yields an empty list, not a failure'
        $dsu = @($script:W1.DatasetUsers)
        $dsu.Count | Should -Be 1
        $dsu[0].UserDatasetUserAccessRight | Should -Be 'ReadWriteReshareExplore'
        $dsu[0].DatasetName | Should -Be 'Finance Model'
        $prm = @($script:W1.DatasetParameters)
        $prm.Count | Should -Be 1
        $prm[0].ParameterName | Should -Be 'ServerName'
        $prm[0].ParameterType | Should -Be 'Text'
        $prm[0].ParameterIsRequired | Should -BeTrue
        $prm[0].ParameterCurrentValue | Should -Be 'sql-finance.contoso.gov'
    }
    It 'global.json has Apps (filtered to scope), AppReports, Connections, Gateways and Capacities' {
        @($script:Global.Apps).Count | Should -Be 1
        @($script:Global.Apps)[0].AppId | Should -Be $script:Ids.app1
        @($script:Global.Apps)[0].AppWorkspaceId | Should -Be $script:Ids.ws1
        $ar = @($script:Global.AppReports)
        $ar.Count | Should -Be 1
        $ar[0].ReportId | Should -Be $script:Ids.r1 -Because 'originalReportObjectId -> ReportId'
        $ar[0].AppReportId | Should -Be $script:Ids.ap1
        @($script:Global.Connections).Count | Should -Be 1
        @($script:Global.Gateways).Count | Should -Be 1
        $cap = @($script:Global.Capacities)
        $cap.Count | Should -Be 1
        $cap[0].CapacityId | Should -Be $script:Ids.cap1
        $cap[0].CapacityDisplayName | Should -Be 'Premium P1'
        $cap[0].CapacitySku | Should -Be 'P1'
        $cap[0].CapacityState | Should -Be 'Active'
        $cap[0].CapacityRegion | Should -Be 'US Gov Virginia'
        $cap[0].CapacityAdmins | Should -Be 'admin@contoso.gov;capadmin@contoso.gov'
        $cap[0].CapacityUsersAccessRight | Should -Be 'Admin'
    }
    It 'My Workspace uses the group-less endpoints; page rows carry ReportId/ReportName (C6-01); shared reports get the pseudo workspace' {
        @($script:My.Datasets).Count | Should -Be 1
        @($script:My.Datasets)[0].DatasetId | Should -Be $script:Ids.md1
        $reports = @($script:My.Reports)
        @($reports | Where-Object { $_.ReportId -eq $script:Ids.mr1 }).Count | Should -Be 1
        @($reports | Where-Object { $_.ReportId -eq $script:Ids.r1 }).Count | Should -Be 0 -Because 'already collected in its real workspace'
        @($reports | Where-Object { $_.ReportId -eq $script:Ids.ap1 }).Count | Should -Be 0 -Because 'app copies are skipped'
        $shared = @($reports | Where-Object { $_.ReportId -eq $script:Ids.sr1 })
        $shared.Count | Should -Be 1
        $shared[0].WorkspaceId | Should -Be 'Shared Reports (No Workspace Access)'
        $pages = @($script:My.ReportPages)
        $pages.Count | Should -Be 2
        $mine = @($pages | Where-Object { $_.ReportId -eq $script:Ids.mr1 })
        $mine.Count | Should -Be 1
        $mine[0].ReportName | Should -Be 'My Report'
        @($script:IQTestApiCalls | Where-Object { $_ -eq 'PowerBI GET datasets' }).Count | Should -BeGreaterOrEqual 1
        @($script:IQTestApiCalls | Where-Object { $_ -eq 'PowerBI GET reports' }).Count | Should -BeGreaterOrEqual 1
    }
    It 'checkpoints every workspace, My Workspace and the global block' {
        Test-IQItemDone -Stage Inventory -ItemKey $script:Ids.ws1 | Should -BeTrue
        Test-IQItemDone -Stage Inventory -ItemKey $script:Ids.ws3 | Should -BeTrue
        Test-IQItemDone -Stage Inventory -ItemKey 'My Workspace' | Should -BeTrue
        Test-IQItemDone -Stage Inventory -ItemKey 'global' | Should -BeTrue
    }
    It 'persists the scope in the manifest' {
        $script:IQ.Manifest.scope.runMode | Should -Be 'Workspaces'
        @($script:IQ.Manifest.scope.workspaceIds) | Should -Contain $script:Ids.ws1
        @($script:IQ.Manifest.scope.workspaceIds) | Should -Contain $script:Ids.ws3
        @($script:IQ.Manifest.scope.workspaceIds).Count | Should -Be 3
        [bool]$script:IQ.Manifest.scope.includeMyWorkspace | Should -BeTrue
    }
    It 'Get-IQSelectedWorkspaces / Datasets / Reports flatten the inventory with the monolith property names' {
        @(Get-IQSelectedWorkspaces -ExcludeSynthetic).Count | Should -Be 3
        $ds = @(Get-IQSelectedDatasets)
        $ds.Count | Should -Be 5 -Because '2 + 1 + 1 + My Workspace'
        $d1 = @($ds | Where-Object { $_.DatasetId -eq $script:Ids.d1 })[0]
        $d1.WorkspaceIsOnDedicatedCapacity | Should -BeTrue
        $d1.WorkspaceName | Should -Be 'Finance [Prod]'
        $rp = @(Get-IQSelectedReports)
        @($rp | Where-Object { $_.ReportId -eq $script:Ids.sr1 }).Count | Should -Be 0 -Because 'shared reports cannot be backed up'
        @($rp | Where-Object { $_.ReportId -eq $script:Ids.r2 }).Count | Should -Be 1
        (Get-IQReportsWithSensitivityLabel).ContainsKey($script:Ids.r1) | Should -BeTrue
    }
    It 'resume: a second run skips every checkpointed workspace without per-workspace calls' {
        $before = $script:IQTestApiCalls.Count
        $script:IQ.IsResume = $true
        $script:IQ.Scope = $null
        $script:IQ.Manifest.stages.Inventory.status = 'Interrupted'
        $script:Summary = $null
        Invoke-IQStage -Name Inventory -Body { $script:Summary = Invoke-IQInventoryStage } | Should -Be 'Completed'
        $script:Summary.Skipped | Should -Be 4
        $script:Summary.Collected | Should -Be 0
        $newCalls = @($script:IQTestApiCalls | Select-Object -Skip $before)
        @($newCalls | Where-Object { $_ -like ('*groups/' + $script:Ids.ws1 + '/datasets*') }).Count | Should -Be 0
    }
}

Describe 'Reports mode (remote model workspace inclusion, monolith 1640-1660)' {
    BeforeAll { Start-InventoryRun -Options @{ RunMode = 'Reports'; ReportId = @($script:Ids.r2) } }
    AfterAll { Remove-IQTestFolder -Path $script:Base }
    It 'selects the report workspace AND the workspace hosting the remote model' {
        $script:StageStatus | Should -Be 'Completed'
        $s = $script:IQ.Scope
        $s.RunMode | Should -Be 'Reports'
        @($s.WorkspaceIds) | Should -Contain $script:Ids.ws2
        @($s.WorkspaceIds) | Should -Contain $script:Ids.ws1
        @($s.WorkspaceIds).Count | Should -Be 2
        @($s.ReportIds) | Should -Be @($script:Ids.r2)
        @($s.DatasetIds) | Should -Be @($script:Ids.d1)
    }
    It 'filters the collected datasets/reports to the selection but keeps dataflows workspace-wide' {
        $w1 = Read-WorkspaceFile $script:Ids.ws1
        $w2 = Read-WorkspaceFile $script:Ids.ws2
        @($w1.Datasets).Count | Should -Be 1
        @($w1.Datasets)[0].DatasetId | Should -Be $script:Ids.d1
        @($w1.Reports).Count | Should -Be 0
        @($w2.Reports).Count | Should -Be 1
        @($w2.Reports)[0].ReportId | Should -Be $script:Ids.r2
        @($w2.Datasets).Count | Should -Be 0
        @($w1.Dataflows).Count | Should -Be 1
        @(Get-IQSelectedReports).Count | Should -Be 1
        @(Get-IQSelectedDatasets).Count | Should -Be 1
        @($script:IQTestApiCalls | Where-Object { $_ -like ('*datasets/' + $script:Ids.d2 + '/refreshes') }).Count | Should -Be 0 -Because 'unselected datasets get no per-item calls'
    }
    It 'requires -ReportId' {
        Initialize-IQTestContext -Options @{ RunMode = 'Reports' } -Prefix 'inv-noreport' | Out-Null
        { Resolve-IQScope | Out-Null } | Should -Throw -ExpectedMessage '*ReportId*'
        Remove-IQTestFolder -Path $script:IQ.BaseFolder
    }
}

Describe 'Models mode (monolith 1755-1790)' {
    BeforeAll { Start-InventoryRun -Options @{ RunMode = 'Models'; DatasetId = @($script:Ids.d1) } }
    AfterAll { Remove-IQTestFolder -Path $script:Base }
    It 'includes the model workspace and every workspace with a report on that model' {
        $script:StageStatus | Should -Be 'Completed'
        $s = $script:IQ.Scope
        foreach ($rid in @($script:Ids.r1, $script:Ids.r2, $script:Ids.r9)) { @($s.ReportIds) | Should -Contain $rid }
        @($s.ReportIds).Count | Should -Be 3
        foreach ($wid in @($script:Ids.ws1, $script:Ids.ws2, $script:Ids.ws4)) { @($s.WorkspaceIds) | Should -Contain $wid }
        @($s.WorkspaceIds).Count | Should -Be 3
        @(Get-IQSelectedDatasets).Count | Should -Be 1
        @(Get-IQSelectedReports).Count | Should -Be 3
    }
}

Describe 'Headless scope guards (brief section 5.3)' {
    It 'throws "No scope" when nothing selects a workspace and the run is not interactive' {
        Initialize-IQTestContext -Options @{ RunMode = 'Workspaces' } -Prefix 'inv-noscope' | Out-Null
        { Resolve-IQScope | Out-Null } | Should -Throw -ExpectedMessage 'No scope*'
        Remove-IQTestFolder -Path $script:IQ.BaseFolder
    }
    It '-AllWorkspaces selects every workspace' {
        Initialize-IQTestContext -Options @{ RunMode = 'Workspaces'; AllWorkspaces = $true } -Prefix 'inv-all' | Out-Null
        @((Resolve-IQScope).WorkspaceIds).Count | Should -Be 4
        Remove-IQTestFolder -Path $script:IQ.BaseFolder
    }
    It '-WorkspaceName wildcards and unknown ids are handled (unknown -> Warn, not fatal)' {
        Initialize-IQTestContext -Options @{ RunMode = 'Workspaces'; WorkspaceName = @('*Analytics'); WorkspaceId = @('00000000-0000-4000-8000-00000000dead') } -Prefix 'inv-wild' | Out-Null
        $s = Resolve-IQScope
        @($s.WorkspaceIds) | Should -Be @($script:Ids.ws3)
        Remove-IQTestFolder -Path $script:IQ.BaseFolder
    }
    It 'My Workspace only is a valid headless scope' {
        Start-InventoryRun -Options @{ RunMode = 'Workspaces'; IncludeMyWorkspace = $true }
        $script:StageStatus | Should -Be 'Completed'
        @(Get-IQSelectedDatasets).Count | Should -Be 1
        Remove-IQTestFolder -Path $script:Base
    }
}

Describe 'Failure handling' {
    BeforeAll {
        Mock Invoke-IQApi {
            if ($Path -like ('groups/' + $script:Ids.ws3 + '/reports')) { throw 'HTTP 500 for GET groups/.../reports after 5 attempt(s)' }
            Invoke-IQTestApiFixture -Method $Method -Path $Path -Body $Body -Query $Query -Api $Api -Raw:$Raw
        }
        Start-InventoryRun -Options @{ RunMode = 'Workspaces'; WorkspaceId = @($script:Ids.ws3) }
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base }
    It 'a core collector exception marks the workspace Failed, keeps the partial file and continues' {
        $script:StageStatus | Should -Be 'CompletedWithErrors'
        $script:Summary.Failed | Should -Be 1
        Test-IQItemDone -Stage Inventory -ItemKey $script:Ids.ws3 | Should -BeFalse
        $w3 = Read-WorkspaceFile $script:Ids.ws3
        @($w3.Datasets).Count | Should -Be 1
        @($w3.Errors).Count | Should -BeGreaterOrEqual 1
        @($script:IQ.Manifest.failures).Count | Should -Be 1
        $script:IQ.Manifest.failures[0].stage | Should -Be 'Inventory'
    }
}
