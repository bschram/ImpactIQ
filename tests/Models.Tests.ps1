# Models.Tests.ps1 - ImpactIQ.Models.ps1 (brief section 7.1, 7.2): ModelBackup / ModelDetail helpers and the stage
# bookkeeping that must hold on resumed runs (review findings M-01 .. M-09). Tabular Editor never runs here: the process
# batch is mocked so the tests are host-independent (Linux CI and Windows PowerShell 5.1).
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Ids = Get-IQTestFixtureJson -Relative 'ids.json'
    $script:Base = Initialize-IQTestContext -Prefix 'models' -RunId '2026-09-04'
    Mock Get-IQToken { 'test-token' }
    $script:FixtureBim = Get-IQTestFixturePath -Relative 'bim/sample-model.bim'
    $script:RunFolder = Get-IQModelRunFolder

    function New-TestWork {
        param([string]$Key, [string]$BaseName, [string]$WorkspaceId, [string]$WorkspaceName = 'WS', [string]$DatasetName = 'Model', [bool]$Dedicated = $true)
        $folder = Get-IQModelRunFolder
        return @{
            Key = $Key; Item = $BaseName; DatasetId = $Key; DatasetName = $DatasetName; WorkspaceId = $WorkspaceId; WorkspaceName = $WorkspaceName
            IsDedicated = $Dedicated; IsPseudoWorkspace = $false; NoAccess = $false; BaseName = $BaseName
            BimPath = (Join-Path $folder ($BaseName + '.bim')); CsvPath = (Join-Path $folder ($BaseName + '.csv')); MdPath = (Join-Path $folder ($BaseName + '_MD.csv'))
            Dataset = [pscustomobject]@{ DatasetId = $Key; DatasetName = $DatasetName; WorkspaceId = $WorkspaceId; WorkspaceName = $WorkspaceName; WorkspaceIsOnDedicatedCapacity = $Dedicated }
        }
    }
    function New-TestDefinition {
        param([string]$Name = 'Any')
        $bim = '{"name":"' + $Name + '","compatibilityLevel":1567,"model":{"tables":[]}}'
        return ([pscustomobject]@{ definition = [pscustomobject]@{ parts = @([pscustomobject]@{ path = 'model.bim'; payloadType = 'InlineBase64'; payload = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($bim)) }) } })
    }
    function New-TestProcessResult {
        param([int]$ExitCode = 0, [bool]$TimedOut = $false, [string]$StartError = $null)
        return @{ ExitCode = $ExitCode; TimedOut = $TimedOut; StdOut = ''; StdErr = ''; OutFile = $null; ErrFile = $null; DurationSec = 1; StartError = $StartError }
    }
    function Get-Checkpoint { param([string]$Stage, [string]$Key) return (Get-IQItemCheckpoint -Stage $Stage -ItemKey $Key) }
}
AfterAll { Remove-IQTestFolder -Path $script:Base }

Describe 'Test-IQModelBimComplete (M-02: truncated .bim files are not backups)' {
    It 'accepts the sample TMSL fixture' {
        Test-IQModelBimComplete -Path $script:FixtureBim | Should -BeTrue
    }
    It 'rejects a truncated copy and deletes it with -RemoveInvalid' {
        $text = [System.IO.File]::ReadAllText($script:FixtureBim)
        $partial = Join-Path $script:RunFolder 'truncated.bim'
        [System.IO.File]::WriteAllText($partial, $text.Substring(0, [int]($text.Length / 2)))
        Test-IQModelBimComplete -Path $partial | Should -BeFalse
        $partial | Should -Exist -Because 'without -RemoveInvalid the file is left alone'
        Test-IQModelBimComplete -Path $partial -RemoveInvalid | Should -BeFalse
        $partial | Should -Not -Exist
    }
    It 'rejects JSON without a model object and missing / empty files' {
        $noModel = Join-Path $script:RunFolder 'nomodel.bim'
        [System.IO.File]::WriteAllText($noModel, '{"name":"x"}')
        Test-IQModelBimComplete -Path $noModel | Should -BeFalse
        Test-IQModelBimComplete -Path (Join-Path $script:RunFolder 'missing.bim') | Should -BeFalse
        Test-IQModelBimComplete -Path '' | Should -BeFalse
        Remove-Item -LiteralPath $noModel -Force
    }
}

Describe 'Complete-IQModelBackupJob (M-01 / M-02)' {
    BeforeEach {
        $script:W = New-TestWork -Key 'job-1' -BaseName 'WS ~ Job' -WorkspaceId $script:Ids.ws1
        $script:Map = @{ 'job-1' = $script:W }
        if (Test-Path -LiteralPath $script:W.BimPath) { Remove-Item -LiteralPath $script:W.BimPath -Force }
    }
    It 'a timed-out (killed) run with a partial .bim is Failed and the partial file is deleted' {
        [System.IO.File]::WriteAllText($script:W.BimPath, '{"name":"WS ~ Job","model":{"tables":[{"name":"S')
        Complete-IQModelBackupJob -Entry @{ ItemKey = 'job-1'; Result = (New-TestProcessResult -ExitCode -1 -TimedOut $true) } -WorkMap $script:Map
        $script:W.BimPath | Should -Not -Exist
        (Get-Checkpoint -Stage ModelBackup -Key 'job-1').status | Should -Be 'Failed'
        (Get-Checkpoint -Stage ModelBackup -Key 'job-1').message | Should -Match 'timed out'
        $script:W.Checkpointed | Should -Be 'Failed'
    }
    It 'a non-empty but truncated .bim with exit code 0 is Failed, not Succeeded' {
        [System.IO.File]::WriteAllText($script:W.BimPath, '{"name":"WS ~ Job","model":{"tables":[{"name":"Sales","columns":[')
        Complete-IQModelBackupJob -Entry @{ ItemKey = 'job-1'; Result = (New-TestProcessResult -ExitCode 0) } -WorkMap $script:Map
        $script:W.BimPath | Should -Not -Exist
        (Get-Checkpoint -Stage ModelBackup -Key 'job-1').status | Should -Be 'Failed'
        $script:W.Checkpointed | Should -Be 'Failed'
    }
    It 'a complete .bim with a non-zero exit code is Succeeded with a note and records Checkpointed' {
        Copy-Item -LiteralPath $script:FixtureBim -Destination $script:W.BimPath -Force
        Complete-IQModelBackupJob -Entry @{ ItemKey = 'job-1'; Result = (New-TestProcessResult -ExitCode 1) } -WorkMap $script:Map
        $cp = Get-Checkpoint -Stage ModelBackup -Key 'job-1'
        $cp.status | Should -Be 'Succeeded'
        $cp.method | Should -Be 'XMLA'
        $cp.message | Should -Match 'exit code 1'
        $script:W.Checkpointed | Should -Be 'Succeeded'
    }
    It 'with -DeferFailure a failure is stored on the work entry and nothing is checkpointed' {
        Set-IQItemDone -Stage ModelBackup -ItemKey 'job-1' -Item 'WS ~ Job' -Status Failed -Message 'stale from an earlier attempt' | Out-Null
        Complete-IQModelBackupJob -Entry @{ ItemKey = 'job-1'; Result = (New-TestProcessResult -ExitCode 1) } -WorkMap $script:Map -DeferFailure
        $script:W.XmlaFailure | Should -Match 'no complete \.bim'
        $script:W.ContainsKey('Checkpointed') | Should -BeFalse
        (Get-Checkpoint -Stage ModelBackup -Key 'job-1').message | Should -Be 'stale from an earlier attempt' -Because 'the deferred failure does not touch the checkpoint'
    }
}

Describe 'Save-IQModelDefinitionFromFabric circuit breakers (M-04)' {
    BeforeEach { $script:FabricCalls = 0 }
    It 'per-item refusals (Invoke-IQFabricLro returns $null) never disable the dedicated-model fallback' {
        Mock Invoke-IQFabricLro { $script:FabricCalls++; return $null }
        $state = @{ Enabled = $true; Reason = ''; ConsecutiveFailures = 0; MaxConsecutiveFailures = 3; BestEffortEnabled = $true; BestEffortReason = ''; BestEffortConsecutiveFailures = 0; MaxBestEffortConsecutiveFailures = 3 }
        foreach ($i in 1..5) {
            $w = New-TestWork -Key ('ref-' + $i) -BaseName ('WS ~ Ref' + $i) -WorkspaceId $script:Ids.ws1
            (Save-IQModelDefinitionFromFabric -Work $w -FabricState $state).Success | Should -BeFalse
        }
        $script:FabricCalls | Should -Be 5
        $state.Enabled | Should -BeTrue
        $state.ConsecutiveFailures | Should -Be 0
    }
    It 'three failed Pro best-effort attempts stop only the best-effort attempts' {
        Mock Invoke-IQFabricLro { $script:FabricCalls++; return $null }
        $state = New-IQModelFabricState
        $state.Enabled | Should -BeTrue -Because 'Get-IQToken is mocked to return a Fabric token'
        foreach ($i in 1..4) {
            $w = New-TestWork -Key ('pro-' + $i) -BaseName ('WS ~ Pro' + $i) -WorkspaceId $script:Ids.ws2 -Dedicated $false
            Save-IQModelDefinitionFromFabric -Work $w -FabricState $state -BestEffort | Out-Null
        }
        $script:FabricCalls | Should -Be 3 -Because 'the fourth Pro attempt is skipped by the best-effort breaker'
        $state.BestEffortEnabled | Should -BeFalse
        $state.Enabled | Should -BeTrue
        $dedicated = New-TestWork -Key 'ded-1' -BaseName 'WS ~ Ded' -WorkspaceId $script:Ids.ws1
        Save-IQModelDefinitionFromFabric -Work $dedicated -FabricState $state | Out-Null
        $script:FabricCalls | Should -Be 4 -Because 'dedicated models are still attempted'
    }
    It 'three consecutive transport failures (the call throws) disable the fallback for everyone' {
        Mock Invoke-IQFabricLro { $script:FabricCalls++; throw 'name resolution failed' }
        $state = New-IQModelFabricState
        foreach ($i in 1..4) {
            $w = New-TestWork -Key ('tr-' + $i) -BaseName ('WS ~ Tr' + $i) -WorkspaceId $script:Ids.ws1
            Save-IQModelDefinitionFromFabric -Work $w -FabricState $state | Out-Null
        }
        $script:FabricCalls | Should -Be 3
        $state.Enabled | Should -BeFalse
        $state.Reason | Should -Match 'transport'
    }
    It 'a success resets both counters and writes the model.bim part' {
        Mock Invoke-IQFabricLro { $script:FabricCalls++; return (New-TestDefinition -Name 'WS ~ Ok') }
        $state = @{ Enabled = $true; Reason = ''; ConsecutiveFailures = 2; MaxConsecutiveFailures = 3; BestEffortEnabled = $true; BestEffortReason = ''; BestEffortConsecutiveFailures = 2; MaxBestEffortConsecutiveFailures = 3 }
        $w = New-TestWork -Key 'ok-1' -BaseName 'WS ~ Ok' -WorkspaceId $script:Ids.ws1
        $r = Save-IQModelDefinitionFromFabric -Work $w -FabricState $state -BestEffort
        $r.Success | Should -BeTrue
        $w.BimPath | Should -Exist
        $state.ConsecutiveFailures | Should -Be 0
        $state.BestEffortConsecutiveFailures | Should -Be 0
    }
}

Describe 'Invoke-IQModelBackupStage with a mocked Tabular Editor batch (M-01, M-05, M-06, M-09)' {
    BeforeAll {
        $script:StageIds = Get-IQTestFixtureJson -Relative 'ids.json'
        $script:Datasets = @(
            [pscustomobject]@{ DatasetId = $script:StageIds.d1; DatasetName = 'Retry Model'; WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Finance; "Prod"'; WorkspaceIsOnDedicatedCapacity = $true },
            [pscustomobject]@{ DatasetId = $script:StageIds.d2; DatasetName = 'Quote "Model"'; WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Finance; "Prod"'; WorkspaceIsOnDedicatedCapacity = $true }
        )
        $script:Workspaces = @([pscustomobject]@{ WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Finance; "Prod"'; WorkspaceIsOnDedicatedCapacity = $true })
        $script:IQ.Tools = @{ TabularEditorWorks = $true; TabularEditorPath = (Join-Path $script:Base 'TabularEditor.exe'); TabularEditorPreflight = $null }
        Mock Test-IQModelTabularEditorAvailable { $true }
        Mock Get-IQSelectedDatasets { $script:Datasets }
        Mock Get-IQSelectedWorkspaces { $script:Workspaces }
        $script:BatchJobs = @()
        $script:FabricCalls = 0
        # The batch never writes a .bim (XMLA "fails"); the callback runs exactly like the real pool does.
        Mock Invoke-IQProcessBatch {
            $script:BatchJobs = @($Jobs)
            $out = @()
            foreach ($j in $Jobs) {
                $entry = @{ ItemKey = $j.ItemKey; Item = $j.Item; Result = (New-TestProcessResult -ExitCode 1) }
                if ($OnJobComplete) { & $OnJobComplete $entry }
                $out += , $entry
            }
            return $out
        }
        Mock Invoke-IQFabricLro { $script:FabricCalls++; return (New-TestDefinition -Name 'Finance Prod ~ Retry Model') }
        # A stale Failed checkpoint from an "earlier attempt" of this run (first attempt had no Fabric token).
        Set-IQItemDone -Stage ModelBackup -ItemKey $script:StageIds.d1 -Item 'stale' -Status Failed -Method 'XMLA' -Message 'earlier attempt: XMLA export produced no .bim' | Out-Null
        $script:IQ.IsResume = $true
        $script:Summary = Invoke-IQModelBackupStage
        $script:IQ.IsResume = $false
    }
    It 'a workspace name with ; and " is exported via XMLA (URL-encoded Data Source), only the quoted dataset name is blocked (M-05)' {
        @($script:BatchJobs).Count | Should -Be 1
        $script:BatchJobs[0].ItemKey | Should -Be $script:StageIds.d1
        $script:BatchJobs[0].ArgumentList | Should -Match 'Finance%3B%20%22Prod%22'
        $cp2 = Get-Checkpoint -Stage ModelBackup -Key $script:StageIds.d2
        $cp2.status | Should -Be 'Succeeded' -Because 'the blocked dataset went through the Fabric fallback'
        $cp2.method | Should -Be 'FabricDefinition'
        $cp2.message | Should -Match 'double quote'
    }
    It 'the deferred XMLA failure runs the Fabric fallback although a stale checkpoint exists (M-01)' {
        $script:FabricCalls | Should -Be 2 -Because 'one call for the blocked dataset, one for the failed XMLA export'
        $cp = Get-Checkpoint -Stage ModelBackup -Key $script:StageIds.d1
        $cp.status | Should -Be 'Succeeded'
        $cp.method | Should -Be 'FabricDefinition'
        $cp.message | Should -Match 'no complete \.bim'
        $cp.message | Should -Not -Match 'earlier attempt'
    }
    It 'counts the outcomes from this batch, not from the checkpoint files' {
        $script:Summary.Total | Should -Be 2
        $script:Summary.Done | Should -Be 2
        $script:Summary.ViaFabric | Should -Be 2
        $script:Summary.ViaXmla | Should -Be 0
        $script:Summary.Failed | Should -Be 0
    }
    It 'an unreadable inventory fails the stage instead of completing it with 0 items (M-06)' {
        Mock Get-IQSelectedDatasets { throw 'inventory file corrupt' }
        { Invoke-IQModelBackupStage } | Should -Throw -ExpectedMessage '*inventory file corrupt*'
        Invoke-IQStage -Name 'ModelDetail' -Body { Invoke-IQModelDetailStage | Out-Null } | Should -Be 'Failed'
        $script:IQ.Manifest.stages.ModelDetail.error | Should -Match 'inventory file corrupt'
    }
}

Describe 'Get-IQModelDatabaseNameFromBim (M-08: head-only read, memoised)' {
    It 'reads the top-level name of the fixture without deserialising the model' {
        Get-IQModelDatabaseNameFromBim -BimPath $script:FixtureBim | Should -Be 'Finance (Prod) ~ Ledger'
    }
    It 'decodes JSON escapes in the name and ignores names inside the model object' {
        $p = Join-Path $script:RunFolder 'escaped.bim'
        [System.IO.File]::WriteAllText($p, '{ "compatibilityLevel": 1567, "name": "A \"quoted\" \\ name", "model": { "tables": [ { "name": "Sales" } ] } }')
        Get-IQModelDatabaseNameFromBim -BimPath $p | Should -Be 'A "quoted" \ name'
        $p2 = Join-Path $script:RunFolder 'noname.bim'
        [System.IO.File]::WriteAllText($p2, '{ "compatibilityLevel": 1567, "model": { "tables": [ { "name": "Sales" } ] } }')
        Get-IQModelDatabaseNameFromBim -BimPath $p2 | Should -BeNullOrEmpty
        Remove-Item -LiteralPath $p, $p2 -Force
    }
    It 'returns $null for a truncated file and a missing file' {
        $p = Join-Path $script:RunFolder 'trunc-name.bim'
        [System.IO.File]::WriteAllText($p, '{ "name": "Half')
        Get-IQModelDatabaseNameFromBim -BimPath $p | Should -BeNullOrEmpty
        Get-IQModelDatabaseNameFromBim -BimPath (Join-Path $script:RunFolder 'nope.bim') | Should -BeNullOrEmpty
        Remove-Item -LiteralPath $p -Force
    }
    It 'Get-IQModelDetailPlan does not hand an unreadable .bim to Tabular Editor' {
        $p = Join-Path $script:RunFolder 'unreadable.bim'
        [System.IO.File]::WriteAllText($p, 'not json at all')
        $plan = Get-IQModelDetailPlan -Method 'Auto' -TabularEditorAvailable $true -BimPath $p -BaseName 'X ~ Y'
        $plan.Steps | Should -Be @('Bim', 'Dax')
        $plan.Why | Should -Match 'could not be read'
        $good = Get-IQModelDetailPlan -Method 'Auto' -TabularEditorAvailable $true -BimPath $script:FixtureBim -BaseName 'Finance (Prod) ~ Ledger'
        $good.Steps | Should -Be @('TabularEditor', 'Bim', 'Dax')
        Remove-Item -LiteralPath $p -Force
    }
}

Describe 'Find-IQModelDetailCsv (M-03: only the run folder is searched)' {
    BeforeAll {
        $script:Other = Join-Path ([string]$script:IQ.Paths.ModelBackups) '2026-09-05'
        New-Item -ItemType Directory -Path $script:Other -Force | Out-Null
        $script:W3 = New-TestWork -Key 'csv-1' -BaseName 'WS ~ Csv' -WorkspaceId $script:Ids.ws1
        $script:StaleCsv = Join-Path $script:Other 'WS ~ Csv.csv'
        $script:StaleMd = Join-Path $script:Other 'WS ~ Csv_MD.csv'
        [System.IO.File]::WriteAllText($script:StaleCsv, 'Type,Table' + [Environment]::NewLine)
        [System.IO.File]::WriteAllText($script:StaleMd, 'ObjectName' + [Environment]::NewLine)
    }
    AfterAll { Remove-IQTestFolder -Path $script:Other }
    It 'does not take (or move) a CSV from another run''s newer date folder' {
        $found = Find-IQModelDetailCsv -Work $script:W3 -RunFolder $script:RunFolder -BimPath $null
        $found.Csv | Should -BeNullOrEmpty
        $found.Md | Should -BeNullOrEmpty
        $script:StaleCsv | Should -Exist
        $script:StaleMd | Should -Exist
    }
    It 'renames a CSV the csx wrote under the .bim database name inside the run folder' {
        $bim = Join-Path $script:RunFolder 'WS ~ Csv.bim'
        [System.IO.File]::WriteAllText($bim, '{"name":"Db Name","model":{}}')
        [System.IO.File]::WriteAllText((Join-Path $script:RunFolder 'Db Name.csv'), 'Type,Table' + [Environment]::NewLine)
        [System.IO.File]::WriteAllText((Join-Path $script:RunFolder 'Db Name_MD.csv'), 'ObjectName' + [Environment]::NewLine)
        $found = Find-IQModelDetailCsv -Work $script:W3 -RunFolder $script:RunFolder -BimPath $bim
        $found.Csv | Should -Be $script:W3.CsvPath
        $found.Md | Should -Be $script:W3.MdPath
        $script:W3.CsvPath | Should -Exist
        (Join-Path $script:RunFolder 'Db Name.csv') | Should -Not -Exist
        Remove-Item -LiteralPath $bim, $script:W3.CsvPath, $script:W3.MdPath -Force
    }
}

Describe 'Get-IQModelReportBackupBimIndex / Get-IQModelBimPath (M-07: one checkpoint scan per stage)' {
    BeforeAll {
        $script:ProBim = Join-Path $script:RunFolder 'Pro ~ Report Model.bim'
        Copy-Item -LiteralPath $script:FixtureBim -Destination $script:ProBim -Force
        Set-IQItemDone -Stage ReportBackup -ItemKey $script:Ids.r1 -Item 'Pro ~ Report' -Outputs @((Join-Path $script:RunFolder 'Pro ~ Report.pbix'), $script:ProBim) -Data @{ DatasetId = $script:Ids.d3.ToUpperInvariant(); BimPath = $script:ProBim } | Out-Null
        Set-IQItemDone -Stage ReportBackup -ItemKey $script:Ids.r2 -Item 'Pro ~ Other' -Outputs @() -Data @{ DatasetId = $script:Ids.d4 } | Out-Null
        $script:W4 = New-TestWork -Key $script:Ids.d3 -BaseName 'Pro ~ Model' -WorkspaceId $script:Ids.ws2 -Dedicated $false
    }
    It 'indexes only checkpoints that name a .bim, keyed by lower-case dataset id' {
        $index = Get-IQModelReportBackupBimIndex
        $index.Count | Should -Be 1
        $index.ContainsKey($script:Ids.d3.ToLowerInvariant()) | Should -BeTrue
        @($index[$script:Ids.d3.ToLowerInvariant()])[0] | Should -Be $script:ProBim
    }
    It 'resolves the Pro model .bim through the index (case-insensitive) and without it' {
        Get-IQModelBimPath -Work $script:W4 -ReportBackupIndex (Get-IQModelReportBackupBimIndex) | Should -Be $script:ProBim
        Get-IQModelBimPath -Work $script:W4 | Should -Be $script:ProBim
        Get-IQModelBimPath -Work $script:W4 -ReportBackupIndex @{} | Should -BeNullOrEmpty
    }
}
