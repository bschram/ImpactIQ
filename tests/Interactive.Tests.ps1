# Interactive.Tests.ps1 - ImpactIQ.Interactive.ps1 selection wrappers (brief section 2.8) with every WinForms dialog mocked.
# The module is WinForms-only, so TestHelpers.ps1 does not load it; this file dot-sources it explicitly and replaces
# Assert-IQInteractiveHost and the Show-* dialogs with Pester mocks. Guards INT-01 (timeout data matches the log),
# INT-02 (picker listings feed the inventory cache), INT-03 (OK with nothing ticked) and INT-04 (TimedOut propagation).
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    . (Join-Path (Join-Path (Join-Path (Get-IQTestRepoRoot) 'Config') 'Modules') 'ImpactIQ.Interactive.ps1')
    $script:Ids = Get-IQTestFixtureJson -Relative 'ids.json'
    Mock Get-IQToken { 'test-token' }
    Mock Invoke-IQApi { Invoke-IQTestApiFixture -Method $Method -Path $Path -Body $Body -Query $Query -Api $Api -Raw:$Raw }
    Mock Assert-IQInteractiveHost { }

    function Start-InteractiveContext {
        param([hashtable]$Options, [string]$Prefix = 'int')
        $script:Base = Initialize-IQTestContext -Options $Options -Prefix $Prefix
        $script:IQ.Interactive = $true
        $script:Workspaces = @(Get-IQWorkspaceList)
        $script:PickerRows = @($script:Workspaces | ForEach-Object { [pscustomobject]@{ id = [string]$_.WorkspaceId; name = [string]$_.WorkspaceName } })
        $script:IQTestApiCalls.Clear()
    }
    function New-PickerResult {
        # Mirrors what the verbatim Show-WorkspacePicker returns (SelectedWorkspaceIds / IncludeMyWorkspace / TimedOut / NothingChecked).
        param([array]$Ids = @(), [bool]$IncludeMy = $false, [bool]$TimedOut = $false, [bool]$NothingChecked = $false)
        return [pscustomobject]@{ SelectedWorkspaceIds = @($Ids); IncludeMyWorkspace = $IncludeMy; TimedOut = $TimedOut; NothingChecked = $NothingChecked }
    }
    function Get-ReportListCallCount { return @($script:IQTestApiCalls | Where-Object { $_ -like 'PowerBI GET groups/*/reports' }).Count }
    function Get-DatasetListCallCount { return @($script:IQTestApiCalls | Where-Object { $_ -like 'PowerBI GET groups/*/datasets' }).Count }
    function Get-LogText { return [System.IO.File]::ReadAllText($script:IQ.LogFile) }
}
AfterAll { $script:IQTestApiOverrides.Clear(); Remove-IQTestFolder -Path $script:Base }

Describe 'Workspaces mode: picker timeout really selects every workspace plus My Workspace (INT-01)' {
    BeforeAll {
        Start-InteractiveContext -Options @{ RunMode = 'Workspaces' } -Prefix 'int-timeout'
        Mock Show-RunModeDialog { 'Workspaces' }
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base }
    It 'timeout with nothing ticked returns every accessible workspace id, IncludeMyWorkspace and TimedOut' {
        # The verbatim picker forces IncludeMyWorkspace on timeout and therefore never substitutes the ids itself.
        Mock Show-WorkspacePicker { New-PickerResult -Ids @() -IncludeMy $true -TimedOut $true -NothingChecked $true }
        $r = Select-IQScopeInteractive -Workspaces $script:PickerRows
        $r.RunMode | Should -Be 'Workspaces'
        @($r.WorkspaceIds).Count | Should -Be 4
        foreach ($k in @('ws1', 'ws2', 'ws3', 'ws4')) { @($r.WorkspaceIds) | Should -Contain $script:Ids.$k }
        $r.IncludeMyWorkspace | Should -BeTrue
        $r.TimedOut | Should -BeTrue
        (Get-LogText) | Should -Match 'timed out - running against every accessible workspace plus My Workspace'
    }
    It 'timeout with some workspaces ticked keeps only those (plus My Workspace)' {
        Mock Show-WorkspacePicker { New-PickerResult -Ids @($script:Ids.ws2) -IncludeMy $true -TimedOut $true -NothingChecked $false }
        $r = Select-IQScopeInteractive -Workspaces $script:PickerRows
        @($r.WorkspaceIds) | Should -Be @($script:Ids.ws2)
        $r.IncludeMyWorkspace | Should -BeTrue
        $r.TimedOut | Should -BeTrue
    }
    It 'Resolve-IQScope builds the whole-tenant scope the timeout message promises' {
        Mock Show-WorkspacePicker { New-PickerResult -Ids @() -IncludeMy $true -TimedOut $true -NothingChecked $true }
        $scope = Resolve-IQScope
        $scope.Source | Should -Be 'Interactive'
        @($scope.WorkspaceIds).Count | Should -Be 4
        $scope.IncludeMyWorkspace | Should -BeTrue
        (Get-LogText) | Should -Match 'Interactive selection timed out'
    }
}

Describe 'Workspaces mode: OK with nothing ticked never widens the scope silently (INT-03)' {
    BeforeAll {
        Start-InteractiveContext -Options @{ RunMode = 'Workspaces' } -Prefix 'int-empty'
        Mock Show-RunModeDialog { 'Workspaces' }
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base }
    It 'asks again instead of accepting the picker''s "every workspace id" substitution' {
        $script:PickerCalls = 0
        Mock Show-WorkspacePicker {
            $script:PickerCalls++
            if ($script:PickerCalls -eq 1) {
                # Exactly what the verbatim picker returns for OK + nothing ticked + "Include My Workspace" off.
                return (New-PickerResult -Ids @($Workspaces | ForEach-Object { $_.id }) -IncludeMy $false -TimedOut $false -NothingChecked $true)
            }
            return (New-PickerResult -Ids @($script:Ids.ws3) -IncludeMy $false -TimedOut $false -NothingChecked $false)
        }
        $r = Select-IQScopeInteractive -Workspaces $script:PickerRows
        $script:PickerCalls | Should -Be 2
        @($r.WorkspaceIds) | Should -Be @($script:Ids.ws3)
        $r.IncludeMyWorkspace | Should -BeFalse
        $r.TimedOut | Should -BeFalse
        (Get-LogText) | Should -Match 'No workspaces selected\. Please select at least one workspace'
    }
    It '"My Workspace only" (nothing ticked + Include My Workspace) stays a valid selection' {
        $script:PickerCalls = 0
        Mock Show-WorkspacePicker { $script:PickerCalls++; New-PickerResult -Ids @() -IncludeMy $true -TimedOut $false -NothingChecked $true }
        $r = Select-IQScopeInteractive -Workspaces $script:PickerRows
        $script:PickerCalls | Should -Be 1
        @($r.WorkspaceIds).Count | Should -Be 0
        $r.IncludeMyWorkspace | Should -BeTrue
        $r.TimedOut | Should -BeFalse
    }
    It 'a normal selection is returned as-is' {
        Mock Show-WorkspacePicker { New-PickerResult -Ids @($script:Ids.ws1, $script:Ids.ws2) -IncludeMy $false -TimedOut $false -NothingChecked $false }
        $r = Select-IQScopeInteractive -Workspaces $script:PickerRows
        @($r.WorkspaceIds) | Should -Be @($script:Ids.ws1, $script:Ids.ws2)
        $r.ReportIds.Count | Should -Be 0
        $r.DatasetIds.Count | Should -Be 0
    }
    It 'Cancel on the run-mode dialog returns $null' {
        Mock Show-RunModeDialog { 'Cancel' }
        Select-IQScopeInteractive -Workspaces $script:PickerRows | Should -BeNullOrEmpty
    }
}

Describe 'Reports mode "Specific": nothing ticked means "all" and the picker timeout is reported (INT-03 / INT-04)' {
    BeforeAll {
        Start-InteractiveContext -Options @{ } -Prefix 'int-reports'
        Mock Show-RunModeDialog { 'Reports' }
        Mock Show-ReportScopeDialog { 'Specific' }
        Mock Show-ReportPicker { @($Reports | Where-Object { $_.ReportId -eq $script:Ids.r1 }) }
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base }
    It 'nothing ticked scans every workspace but returns WorkspaceIds = @() (not hundreds of explicit ids)' {
        Mock Show-WorkspacePicker { New-PickerResult -Ids @($Workspaces | ForEach-Object { $_.id }) -IncludeMy $false -TimedOut $false -NothingChecked $true }
        $before = Get-ReportListCallCount
        $r = Select-IQScopeInteractive -Workspaces $script:PickerRows
        $r.RunMode | Should -Be 'Reports'
        @($r.WorkspaceIds).Count | Should -Be 0
        @($r.ReportIds) | Should -Be @($script:Ids.r1)
        @($r.DatasetIds) | Should -Be @($script:Ids.d1)
        $r.TimedOut | Should -BeFalse
        ((Get-ReportListCallCount) - $before) | Should -Be 4
        (Get-LogText) | Should -Match 'No workspaces selected - showing reports from ALL workspaces instead'
    }
    It 'a timed-out workspace picker is reported as TimedOut, not as a deliberate selection' {
        Mock Show-WorkspacePicker { New-PickerResult -Ids @() -IncludeMy $true -TimedOut $true -NothingChecked $true }
        $r = Select-IQScopeInteractive -Workspaces $script:PickerRows
        $r.TimedOut | Should -BeTrue
        @($r.WorkspaceIds).Count | Should -Be 0
        @($r.ReportIds) | Should -Be @($script:Ids.r1)
        (Get-LogText) | Should -Match 'Workspace picker timed out - showing reports from ALL workspaces'
    }
    It 'a real selection lists only the ticked workspace and carries its id' {
        Start-InteractiveContext -Options @{ } -Prefix 'int-reports-one'
        Mock Show-WorkspacePicker { New-PickerResult -Ids @($script:Ids.ws1) -IncludeMy $false -TimedOut $false -NothingChecked $false }
        $r = Select-IQScopeInteractive -Workspaces $script:PickerRows
        @($r.WorkspaceIds) | Should -Be @($script:Ids.ws1)
        $r.TimedOut | Should -BeFalse
        (Get-ReportListCallCount) | Should -Be 1
        @($script:IQTestApiCalls) | Should -Contain ('PowerBI GET groups/{0}/reports' -f $script:Ids.ws1)
    }
    It '"All" scope never reports a timeout' {
        Mock Show-ReportScopeDialog { 'All' }
        Mock Show-WorkspacePicker { throw 'Show-WorkspacePicker must not be shown for the All scope' }
        $r = Select-IQScopeInteractive -Workspaces $script:PickerRows
        $r.TimedOut | Should -BeFalse
        @($r.WorkspaceIds).Count | Should -Be 0
    }
}

Describe 'Models mode "Specific": picker timeout is reported (INT-04)' {
    BeforeAll {
        Start-InteractiveContext -Options @{ } -Prefix 'int-models'
        Mock Show-RunModeDialog { 'Models' }
        Mock Show-ModelScopeDialog { 'Specific' }
        Mock Show-ModelPicker { @($Models | Where-Object { $_.DatasetId -eq $script:Ids.d3 }) }
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base }
    It 'timed-out picker => TimedOut = $true with WorkspaceIds = @()' {
        Mock Show-WorkspacePicker { New-PickerResult -Ids @() -IncludeMy $true -TimedOut $true -NothingChecked $true }
        $r = Select-IQScopeInteractive -Workspaces $script:PickerRows
        $r.RunMode | Should -Be 'Models'
        $r.TimedOut | Should -BeTrue
        @($r.WorkspaceIds).Count | Should -Be 0
        @($r.DatasetIds) | Should -Be @($script:Ids.d3)
        $r.ReportIds.Count | Should -Be 0
        (Get-LogText) | Should -Match 'Workspace picker timed out - showing models from ALL workspaces'
    }
    It 'a real selection => TimedOut = $false and only that workspace is listed' {
        Start-InteractiveContext -Options @{ } -Prefix 'int-models-one'
        Mock Show-WorkspacePicker { New-PickerResult -Ids @($script:Ids.ws2) -IncludeMy $false -TimedOut $false -NothingChecked $false }
        $r = Select-IQScopeInteractive -Workspaces $script:PickerRows
        $r.TimedOut | Should -BeFalse
        @($r.WorkspaceIds) | Should -Be @($script:Ids.ws2)
        (Get-DatasetListCallCount) | Should -Be 1
    }
}

Describe 'Picker listings populate the per-run inventory cache (INT-02)' {
    BeforeAll { Start-InteractiveContext -Options @{ } -Prefix 'int-cache' }
    AfterAll { $script:IQTestApiOverrides.Clear(); Remove-IQTestFolder -Path $script:Base }
    It 'reports listed for the picker are stored under the workspace id and Get-IQScopeReportScan does not list again' {
        $rows = @(Get-IQInteractiveWorkspaceItemList -Workspaces $script:PickerRows -Kind Reports)
        $rows.Count | Should -Be 7
        (Get-ReportListCallCount) | Should -Be 4
        $cache = Get-IQInventoryCache
        foreach ($k in @('ws1', 'ws2', 'ws3', 'ws4')) { $cache.Reports.ContainsKey([string]$script:Ids.$k) | Should -BeTrue }
        @($cache.Reports[[string]$script:Ids.ws1]).Count | Should -Be 2
        $scan = @(Get-IQScopeReportScan -Workspaces $script:Workspaces)
        $scan.Count | Should -Be 7
        (Get-ReportListCallCount) | Should -Be 4
        # A second pass through the picker (Cancel -> back -> Specific again) is a cache hit as well.
        @(Get-IQInteractiveWorkspaceItemList -Workspaces $script:PickerRows -Kind Reports).Count | Should -Be 7
        (Get-ReportListCallCount) | Should -Be 4
    }
    It 'datasets listed for the picker are stored and Get-IQScopeDatasetScan does not list again' {
        $rows = @(Get-IQInteractiveWorkspaceItemList -Workspaces $script:PickerRows -Kind Models)
        $rows.Count | Should -Be 5
        (Get-DatasetListCallCount) | Should -Be 4
        (Get-IQInventoryCache).Datasets.ContainsKey([string]$script:Ids.ws2) | Should -BeTrue
        @(Get-IQScopeDatasetScan -Workspaces $script:Workspaces).Count | Should -Be 5
        (Get-DatasetListCallCount) | Should -Be 4
    }
    It 'a failed listing (403/404 => $null) is NOT cached, so the Inventory stage re-fetches with error bookkeeping (INV-01)' {
        Start-InteractiveContext -Options @{ } -Prefix 'int-cache-miss'
        $script:IQTestApiOverrides[('PowerBI__groups__{0}__reports.json' -f $script:Ids.ws4)] = { $null }
        try {
            $rows = @(Get-IQInteractiveWorkspaceItemList -Workspaces $script:PickerRows -Kind Reports)
            $rows.Count | Should -Be 6
            $cache = Get-IQInventoryCache
            $cache.Reports.ContainsKey([string]$script:Ids.ws1) | Should -BeTrue
            $cache.Reports.ContainsKey([string]$script:Ids.ws4) | Should -BeFalse
        }
        finally { $script:IQTestApiOverrides.Clear() }
    }
    It 'an empty workspace is cached as an empty list (not re-listed as a miss)' {
        Start-InteractiveContext -Options @{ } -Prefix 'int-cache-empty'
        $script:IQTestApiOverrides[('PowerBI__groups__{0}__reports.json' -f $script:Ids.ws3)] = { [pscustomobject]@{ value = @() } }
        try {
            @(Get-IQInteractiveWorkspaceItemList -Workspaces $script:PickerRows -Kind Reports).Count | Should -Be 6
            (Get-IQInventoryCache).Reports.ContainsKey([string]$script:Ids.ws3) | Should -BeTrue
            @(Get-IQInteractiveWorkspaceItemList -Workspaces $script:PickerRows -Kind Reports).Count | Should -Be 6
            (Get-ReportListCallCount) | Should -Be 4
        }
        finally { $script:IQTestApiOverrides.Clear() }
    }
    It 'end to end: Reports mode through Resolve-IQScope lists the tenant once, not twice' {
        Start-InteractiveContext -Options @{ } -Prefix 'int-cache-e2e'
        Mock Show-RunModeDialog { 'Reports' }
        Mock Show-ReportScopeDialog { 'All' }
        Mock Show-ReportPicker { @($Reports | Where-Object { $_.ReportId -eq $script:Ids.r5 }) }
        $scope = Resolve-IQScope
        $scope.RunMode | Should -Be 'Reports'
        $scope.Source | Should -Be 'Interactive'
        @($scope.ReportIds) | Should -Contain $script:Ids.r5
        (Get-ReportListCallCount) | Should -Be 4
    }
}

Describe 'Select-IQScopeInteractive input handling' {
    BeforeAll {
        Start-InteractiveContext -Options @{ } -Prefix 'int-input'
        Mock Show-RunModeDialog { 'Workspaces' }
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base }
    It 'accepts renamed inventory rows (WorkspaceId / WorkspaceName) as well as raw {id; name} rows' {
        Mock Show-WorkspacePicker { New-PickerResult -Ids @($Workspaces[0].id) -IncludeMy $false -TimedOut $false -NothingChecked $false }
        $r = Select-IQScopeInteractive -Workspaces $script:Workspaces
        @($r.WorkspaceIds) | Should -Be @([string]$script:Workspaces[0].WorkspaceId)
    }
    It 'an empty workspace list is an error, not a cancel' {
        { Select-IQScopeInteractive -Workspaces @() } | Should -Throw -ExpectedMessage 'No workspaces are available for selection*'
    }
}
