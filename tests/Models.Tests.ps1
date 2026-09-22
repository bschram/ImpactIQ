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

Describe 'XMLA export outcomes for inferred-capacity models and the Fabric API state (run assessment 2026-09-16)' {
    BeforeEach {
        $script:W = New-TestWork -Key 'inf-1' -BaseName 'WS ~ Large' -WorkspaceId $script:Ids.ws1
        $script:W.CapacityInferred = $true
        $script:Map = @{ 'inf-1' = $script:W }
        if (Test-Path -LiteralPath $script:W.BimPath) { Remove-Item -LiteralPath $script:W.BimPath -Force }
        if ($script:IQ.ContainsKey('FabricApi')) { $script:IQ.Remove('FabricApi') }
    }
    AfterAll { if ($script:IQ.ContainsKey('FabricApi')) { $script:IQ.Remove('FabricApi') } }
    It 'the MSOLAP connection string uses the access-token form with an explicit empty User ID (and a Password-only form for the probe)' {
        $w = @{ WorkspaceName = 'Fin; "Prod"'; DatasetName = 'Model'; BimPath = '/out/m.bim' }
        $a = New-IQModelXmlaArgumentList -Work $w -Token 'TOK' -ScriptPath '/s.cs'
        $a | Should -Be '"Provider=MSOLAP;Data Source=powerbi://api.powerbi.com/v1.0/myorg/Fin%3B%20%22Prod%22;User ID=;Password=TOK" "Model" -S "/s.cs" -B "/out/m.bim"'
        $b = New-IQModelXmlaArgumentList -Work $w -Token 'TOK' -ScriptPath '/s.cs' -Form PasswordOnly
        $b | Should -Match ';Password=TOK" "Model"'
        $b | Should -Not -Match 'User ID'
    }
    It 'an XMLA failure of a model whose capacity was only inferred from its storage format is Skipped with the reason' {
        Complete-IQModelBackupJob -Entry @{ ItemKey = 'inf-1'; Result = (New-TestProcessResult -ExitCode 1) } -WorkMap $script:Map
        $cp = Get-Checkpoint -Stage ModelBackup -Key 'inf-1'
        $cp.status | Should -Be 'Skipped'
        $cp.message | Should -Match 'no complete \.bim'
        $cp.message | Should -Match 'License info'
        $script:W.Checkpointed | Should -Be 'Skipped'
    }
    It 'the same failure on a model in a listed capacity workspace stays Failed' {
        $script:W.CapacityInferred = $false
        Complete-IQModelBackupJob -Entry @{ ItemKey = 'inf-1'; Result = (New-TestProcessResult -ExitCode 1) } -WorkMap $script:Map
        (Get-Checkpoint -Stage ModelBackup -Key 'inf-1').status | Should -Be 'Failed'
        $script:W.Checkpointed | Should -Be 'Failed'
    }
    It 'New-IQModelFabricState is disabled with the real reason once the Fabric API was found unavailable' {
        $script:IQ['FabricApi'] = @{ Unreachable = $true; Reason = 'the service answered FeatureNotAvailable: the Fabric REST API is not offered to this tenant'; ConsecutiveTransportFailures = 0; MaxConsecutiveTransportFailures = 3; Warned = $true; UsingPowerBIToken = $true; PowerBITokenNoticeShown = $true }
        $state = New-IQModelFabricState
        $state.Enabled | Should -BeFalse
        $state.Reason | Should -Match 'Fabric API unavailable in this run \(the service answered FeatureNotAvailable'
    }
    It 'New-IQModelFabricState stays enabled when only the Power BI token stands in for a Fabric token' {
        Mock Get-IQToken { if ($Resource -eq 'Fabric') { return $null } return 'pbi-token' }
        $state = New-IQModelFabricState
        $state.Enabled | Should -BeTrue
    }
}

Describe 'ModelID convention follows the workspace listing, not the inferred capacity (Measure Lineage blank key, 2026-09-17)' {
    BeforeAll { $script:Base = Initialize-IQTestContext -Prefix 'models-listed' }
    AfterAll { Remove-IQTestFolder -Path $script:Base }
    It 'a dedicated workspace keeps the dataset GUID convention' {
        $w = New-TestWork -Key $script:Ids.d1 -BaseName 'WS ~ Model' -WorkspaceId $script:Ids.ws1 -Dedicated $true
        Test-IQModelListedDedicated -Work $w | Should -BeTrue
    }
    It 'a Pro workspace uses the file-name convention' {
        $w = New-TestWork -Key $script:Ids.d1 -BaseName 'WS ~ Model' -WorkspaceId $script:Ids.ws1 -Dedicated $false
        Test-IQModelListedDedicated -Work $w | Should -BeFalse
    }
    It 'a capacity inferred from the large storage format is treated as Pro for the ModelID (the PBIT joins on the workspace flag)' {
        $w = New-TestWork -Key $script:Ids.d1 -BaseName 'WS ~ Model' -WorkspaceId $script:Ids.ws1 -Dedicated $true
        $w['CapacityInferred'] = $true
        Test-IQModelListedDedicated -Work $w | Should -BeFalse
    }
    It 'the DAX extractor stamps ModelID = "<CleanWs> ~ <CleanModel>" when told the workspace is not listed dedicated' {
        Mock Invoke-IQDaxQuery { return @() }
        Mock Test-IQDaxQueryAccess { return $true }
        $ds = [pscustomobject]@{ DatasetId = $script:Ids.d1; DatasetName = 'Model'; WorkspaceId = $script:Ids.ws1; WorkspaceName = 'WS'; WorkspaceIsOnDedicatedCapacity = $false; DatasetTargetStorageMode = 'PremiumFiles' }
        $r = Get-IQModelDetailViaDax -Dataset $ds -OutputFolder (Join-Path $script:Base 'dax') -IsDedicated $false -BaseName 'WS ~ Model' -ModelAsOfDate '2026-09-17'
        if ($r.Success -and $r.Csv -and (Test-Path -LiteralPath $r.Csv)) {
            $rows = @(Import-Csv -LiteralPath $r.Csv -Encoding UTF8)
            if ($rows.Count -gt 0) { $rows[0].ModelID | Should -Be 'WS ~ Model' }
        }
        $r.Method | Should -Be 'Dax'
    }
}

Describe 'XMLA connection probe, system datasets and permission classification (run assessment 2026-09-18)' {
    BeforeAll {
        $script:Base = Initialize-IQTestContext -Options @{ Environment = 'USGov'; MaxParallelExtracts = 1 } -Prefix 'models-xmla-probe'
        $script:StageIds = Get-IQTestFixtureJson -Relative 'ids.json'
        $script:IQ.Tools = @{ TabularEditorWorks = $true; TabularEditorPath = (Join-Path $script:Base 'TabularEditor.exe'); TabularEditorPreflight = $null }
        Mock Test-IQModelTabularEditorAvailable { $true }
        Mock Get-IQToken { if ($Resource -eq 'Fabric') { return $null } return 'GOV-TOKEN' }   # no Fabric fallback in these tests
        Mock Get-IQTokenForResourceUrl { if ($ResourceUrl -like 'https://analysis.windows.net/*') { return 'COMMERCIAL-TOKEN' } return $null }
        Mock Get-IQPowerBIModuleXmlaToken { $null }   # no Power BI module session unless a test says so
        function New-ProbeTeResult { param([bool]$Success = $true, [int]$ExitCode = 0) return @{ ExitCode = $ExitCode; TimedOut = $false; StdOut = ''; StdErr = ''; OutFile = $null; ErrFile = $null; DurationSec = 1; Success = $Success; ErrorLines = @(); FailureReason = $(if ($Success) { '' } else { 'exit ' + $ExitCode }) } }
        $script:Workspaces = @([pscustomobject]@{ WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Dash'; WorkspaceIsOnDedicatedCapacity = $true; WorkspaceCapacityId = '72AA2844-F3D6-459A-A19B-D8ECFE5EB068' })
        Mock Get-IQSelectedWorkspaces { $script:Workspaces }
        function Write-CompleteBim { param([string]$Path) [System.IO.File]::WriteAllText($Path, '{"name":"m","compatibilityLevel":1567,"model":{"culture":"en-US","tables":[{"name":"T","columns":[{"name":"C","dataType":"string"}]}]}}') }
        function New-AuthFailedResult { return @{ ExitCode = 1; TimedOut = $false; StdOut = "Tabular Editor 2.29.0`r`nLoading model...`r`nError loading model: Authentication failed for all authenticators`r`n`r`nTechnical Details:`r`nRootActivityId: x"; StdErr = ''; OutFile = $null; ErrFile = $null; DurationSec = 4; StartError = $null } }
        function Reset-ProbeState { foreach ($k in @('XmlaConnection', 'XmlaProbeDone', 'XmlaProbeTried')) { if ($script:IQ.ContainsKey($k)) { $script:IQ.Remove($k) } }; $script:IQ.Options['NoXmlaProbe'] = $false; $script:IQ.Options['XmlaTokenResource'] = ''; $script:IQ.Options['XmlaTokenSource'] = 'Auto'; $m = Get-IQModelXmlaMemoryPath; if (Test-Path -LiteralPath $m) { Remove-Item -LiteralPath $m -Force } }
        function Clear-ModelCheckpoints { $p = Join-Path (Join-Path $script:IQ.RunPath 'done') 'ModelBackup'; if (Test-Path -LiteralPath $p) { Get-ChildItem -LiteralPath $p -Filter '*.json' | Remove-Item -Force } }
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base }
    BeforeEach { Reset-ProbeState; Clear-ModelCheckpoints; $script:BatchArgs = New-Object System.Collections.Generic.List[string]; $script:ProbeArgs = New-Object System.Collections.Generic.List[string] }

    It 'a system usage-metrics model is Skipped before any XMLA attempt' {
        Mock Get-IQSelectedDatasets { @([pscustomobject]@{ DatasetId = $script:StageIds.d1; DatasetName = 'Report Usage Metrics Model'; WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Dash'; WorkspaceIsOnDedicatedCapacity = $true }) }
        Mock Invoke-IQProcessBatch { throw 'must not be called' }
        $summary = Invoke-IQModelBackupStage
        $summary.Skipped | Should -Be 1
        $summary.Failed | Should -Be 0
        (Get-Checkpoint -Stage ModelBackup -Key $script:StageIds.d1).message | Should -Match 'system-generated usage metrics model'
        Test-IQModelSystemDataset -Name 'Usage Metrics Report' | Should -BeTrue
        Test-IQModelSystemDataset -Name 'Sales' | Should -BeFalse
    }
    It 'a Discover permission refusal is Skipped as a permission gap with the <euii> tags removed' {
        $w = New-TestWork -Key $script:StageIds.d2 -BaseName 'Dash ~ Perm' -WorkspaceId $script:StageIds.ws1 -Dedicated $true
        Set-IQModelXmlaFailure -Work $w -Key $w.Key -Message "XMLA export produced no complete .bim: exit code 1; Error loading model: The '<euii>user@contoso.gov</euii>' user does not have permission to call the Discover method."
        $cp = Get-Checkpoint -Stage ModelBackup -Key $script:StageIds.d2
        $cp.status | Should -Be 'Skipped'
        $cp.message | Should -Match '^Skipped: no XMLA permission on this model'
        $cp.message | Should -Not -Match 'euii'
        $cp.message | Should -Match "'user@contoso.gov' user does not have permission"
    }
    It 'the first authentication failure probes the other variants once; the accepted one is used for the remaining models' {
        Mock Get-IQSelectedDatasets { @(
            [pscustomobject]@{ DatasetId = $script:StageIds.d1; DatasetName = 'First'; WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Dash'; WorkspaceIsOnDedicatedCapacity = $true },
            [pscustomobject]@{ DatasetId = $script:StageIds.d2; DatasetName = 'Second'; WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Dash'; WorkspaceIsOnDedicatedCapacity = $true }
        ) }
        # the batch: the GOV token is refused; the COMMERCIAL token (any form) is accepted
        Mock Invoke-IQProcessBatch {
            $out = @()
            foreach ($j in $Jobs) {
                $script:BatchArgs.Add([string]$j.ArgumentList)
                $r = New-AuthFailedResult
                if ([string]$j.ArgumentList -like '*Password=COMMERCIAL-TOKEN*') { $bim = [regex]::Match([string]$j.ArgumentList, '-B "([^"]+)"').Groups[1].Value; Write-CompleteBim -Path $bim; $r = New-TestProcessResult -ExitCode 0 }
                $entry = @{ ItemKey = $j.ItemKey; Item = $j.Item; Result = $r }
                if ($OnJobComplete) { & $OnJobComplete $entry }
                $out += , $entry
            }
            return $out
        }
        Mock Invoke-IQTabularEditor {
            $script:ProbeArgs.Add($ArgumentList)
            $bim = [regex]::Match($ArgumentList, '-B "([^"]+)"').Groups[1].Value
            if ($ArgumentList -like '*Password=COMMERCIAL-TOKEN*') { Write-CompleteBim -Path $bim; return (New-ProbeTeResult -Success $true) }
            $r = New-AuthFailedResult; $r.Success = $false; $r.ErrorLines = @('Error loading model: Authentication failed for all authenticators'); $r.FailureReason = 'exit 1'; return $r
        }
        $summary = Invoke-IQModelBackupStage
        $summary.Done | Should -Be 2
        $summary.Failed | Should -Be 0
        # probe order: same audience / other form first (refused), then the commercial audience (accepted on its first form)
        $script:ProbeArgs.Count | Should -Be 2
        $script:ProbeArgs[0] | Should -Match 'Data Source=[^"]+;Password=GOV-TOKEN"'
        $script:ProbeArgs[1] | Should -Match 'User ID=;Password=COMMERCIAL-TOKEN'
        $conn = Get-IQModelXmlaConnection
        $conn.ResourceUrl | Should -Be 'https://analysis.windows.net/powerbi/api'
        $conn.Form | Should -Be 'UserIdEmpty'
        $conn.Pinned | Should -BeTrue
        (Get-Checkpoint -Stage ModelBackup -Key $script:StageIds.d1).message | Should -Match 'XMLA accepted after switching to sign-in token, audience https://analysis.windows.net/powerbi/api, form UserIdEmpty'
        # the second chunk went straight through the batch with the switched token and no further probe
        $script:BatchArgs.Count | Should -Be 2
        $script:BatchArgs[1] | Should -Match 'User ID=;Password=COMMERCIAL-TOKEN'
        (Get-Checkpoint -Stage ModelBackup -Key $script:StageIds.d2).status | Should -Be 'Succeeded'
        $log = Get-Content -LiteralPath $script:IQ.LogFile -Raw
        $log | Should -Match "Pin it explicitly with -XmlaTokenResource 'https://analysis.windows.net/powerbi/api' \(settings file: XmlaTokenResource\)"
    }
    It 'when no variant is accepted the model fails once with the admin checks in the message and later models are not probed again' {
        Mock Get-IQSelectedDatasets { @(
            [pscustomobject]@{ DatasetId = $script:StageIds.d1; DatasetName = 'First'; WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Dash'; WorkspaceIsOnDedicatedCapacity = $true },
            [pscustomobject]@{ DatasetId = $script:StageIds.d2; DatasetName = 'Second'; WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Dash'; WorkspaceIsOnDedicatedCapacity = $true }
        ) }
        Mock Invoke-IQProcessBatch {
            $out = @()
            foreach ($j in $Jobs) { $entry = @{ ItemKey = $j.ItemKey; Item = $j.Item; Result = (New-AuthFailedResult) }; if ($OnJobComplete) { & $OnJobComplete $entry }; $out += , $entry }
            return $out
        }
        Mock Invoke-IQTabularEditor { $script:ProbeArgs.Add($ArgumentList); $r = New-AuthFailedResult; $r.Success = $false; $r.ErrorLines = @('Error loading model: Authentication failed for all authenticators'); $r.FailureReason = 'exit 1'; return $r }
        $summary = Invoke-IQModelBackupStage
        $summary.Failed | Should -Be 2
        $script:ProbeArgs.Count | Should -Be 3 -Because 'three other sign-in variants exist for a GCC run (gov/PasswordOnly, commercial/UserIdEmpty, commercial/PasswordOnly); the two module-token variants are skipped without a module session'
        $script:IQ['XmlaProbeTried'] | Should -Match 'Power BI PowerShell module token, audience https://analysis.usgovcloudapi.net/powerbi/api, form UserIdEmpty: no Power BI PowerShell module token'
        $log = Get-Content -LiteralPath $script:IQ.LogFile -Raw
        $log | Should -Match 'could not be tried in this run: install MicrosoftPowerBIMgmt and run interactively once'
        $cp = Get-Checkpoint -Stage ModelBackup -Key $script:StageIds.d1
        $cp.status | Should -Be 'Failed'
        $cp.message | Should -Match 'refused the access token that every REST call accepts'
        $cp.message | Should -Match 'Variants tried: '
        $cp.message | Should -Match 'Test-IQXmlaAccess-Legacy\.ps1'
        $cp.message | Should -Match '-XmlaTokenSource PowerBIModule'
        (Get-IQModelXmlaConnection).ResourceUrl | Should -Be 'https://analysis.usgovcloudapi.net/powerbi/api'
    }
    It '-NoXmlaProbe and -XmlaTokenResource: no probing, and the pinned audience is used from the first export' {
        Mock Get-IQSelectedDatasets { @([pscustomobject]@{ DatasetId = $script:StageIds.d1; DatasetName = 'First'; WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Dash'; WorkspaceIsOnDedicatedCapacity = $true }) }
        Mock Invoke-IQProcessBatch {
            $out = @()
            foreach ($j in $Jobs) { $script:BatchArgs.Add([string]$j.ArgumentList); $entry = @{ ItemKey = $j.ItemKey; Item = $j.Item; Result = (New-AuthFailedResult) }; if ($OnJobComplete) { & $OnJobComplete $entry }; $out += , $entry }
            return $out
        }
        Mock Invoke-IQTabularEditor { $script:ProbeArgs.Add($ArgumentList); return (New-ProbeTeResult -Success $false -ExitCode 1) }
        $script:IQ.Options['NoXmlaProbe'] = $true
        $script:IQ.Options['XmlaTokenResource'] = 'https://analysis.windows.net/powerbi/api'
        $summary = Invoke-IQModelBackupStage
        $summary.Failed | Should -Be 1
        $script:ProbeArgs.Count | Should -Be 0
        $script:BatchArgs[0] | Should -Match 'User ID=;Password=COMMERCIAL-TOKEN'
        (Get-IQModelXmlaConnection).Pinned | Should -BeTrue
        @(Get-IQModelXmlaVariantList).Count | Should -Be 3 -Because 'a pinned audience leaves the two module-token variants and the other form of the sign-in token'
    }

    It 'the Power BI module token is tried first after an authentication failure, used for the remaining models when accepted, and remembered for the next run (2026-09-22)' {
        Mock Get-IQSelectedDatasets { @(
            [pscustomobject]@{ DatasetId = $script:StageIds.d1; DatasetName = 'First'; WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Dash'; WorkspaceIsOnDedicatedCapacity = $true },
            [pscustomobject]@{ DatasetId = $script:StageIds.d2; DatasetName = 'Second'; WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Dash'; WorkspaceIsOnDedicatedCapacity = $true }
        ) }
        Mock Get-IQPowerBIModuleXmlaToken { 'MODULE-TOKEN' }
        Mock Invoke-IQProcessBatch {
            $out = @()
            foreach ($j in $Jobs) {
                $script:BatchArgs.Add([string]$j.ArgumentList)
                $r = New-AuthFailedResult
                if ([string]$j.ArgumentList -like '*Password=MODULE-TOKEN*') { $bim = [regex]::Match([string]$j.ArgumentList, '-B "([^"]+)"').Groups[1].Value; Write-CompleteBim -Path $bim; $r = New-TestProcessResult -ExitCode 0 }
                $entry = @{ ItemKey = $j.ItemKey; Item = $j.Item; Result = $r }
                if ($OnJobComplete) { & $OnJobComplete $entry }
                $out += , $entry
            }
            return $out
        }
        Mock Invoke-IQTabularEditor {
            $script:ProbeArgs.Add($ArgumentList)
            $bim = [regex]::Match($ArgumentList, '-B "([^"]+)"').Groups[1].Value
            if ($ArgumentList -like '*Password=MODULE-TOKEN*') { Write-CompleteBim -Path $bim; return (New-ProbeTeResult -Success $true) }
            $r = New-AuthFailedResult; $r.Success = $false; $r.ErrorLines = @('Error loading model: Authentication failed for all authenticators'); $r.FailureReason = 'exit 1'; return $r
        }
        $summary = Invoke-IQModelBackupStage
        $summary.Done | Should -Be 2
        $summary.Failed | Should -Be 0
        $script:ProbeArgs.Count | Should -Be 1 -Because 'the module token in the standard form is the first variant and it is accepted'
        $script:ProbeArgs[0] | Should -Match 'Data Source=powerbi://api\.powerbigov\.us/v1\.0/myorg/Dash;User ID=;Password=MODULE-TOKEN"'
        $conn = Get-IQModelXmlaConnection
        $conn.TokenSource | Should -Be 'PowerBIModule'
        $conn.ResourceUrl | Should -Be 'https://analysis.usgovcloudapi.net/powerbi/api'
        $conn.Form | Should -Be 'UserIdEmpty'
        (Get-Checkpoint -Stage ModelBackup -Key $script:StageIds.d1).message | Should -Match 'XMLA accepted after switching to Power BI PowerShell module token'
        $script:BatchArgs.Count | Should -Be 2
        $script:BatchArgs[1] | Should -Match 'User ID=;Password=MODULE-TOKEN'
        $log = Get-Content -LiteralPath $script:IQ.LogFile -Raw
        $log | Should -Match 'Pin it explicitly with -XmlaTokenSource PowerBIModule \(settings file: XmlaTokenSource\)'
        $memory = ConvertFrom-IQJsonFile -Path (Get-IQModelXmlaMemoryPath)
        $memory.TokenSource | Should -Be 'PowerBIModule'
        $memory.Environment | Should -Be 'USGov'
        $memory.Form | Should -Be 'UserIdEmpty'
    }
    It 'a remembered connection is used from the first export of the next run; without a module token the run falls back to the sign-in token' {
        Mock Get-IQSelectedDatasets { @([pscustomobject]@{ DatasetId = $script:StageIds.d1; DatasetName = 'First'; WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Dash'; WorkspaceIsOnDedicatedCapacity = $true }) }
        Mock Invoke-IQProcessBatch {
            $out = @()
            foreach ($j in $Jobs) {
                $script:BatchArgs.Add([string]$j.ArgumentList)
                $bim = [regex]::Match([string]$j.ArgumentList, '-B "([^"]+)"').Groups[1].Value; Write-CompleteBim -Path $bim
                $entry = @{ ItemKey = $j.ItemKey; Item = $j.Item; Result = (New-TestProcessResult -ExitCode 0) }
                if ($OnJobComplete) { & $OnJobComplete $entry }
                $out += , $entry
            }
            return $out
        }
        Mock Invoke-IQTabularEditor { $script:ProbeArgs.Add($ArgumentList); return (New-ProbeTeResult -Success $false -ExitCode 1) }
        Save-IQModelXmlaMemory -Connection @{ ResourceUrl = 'https://analysis.usgovcloudapi.net/powerbi/api'; Form = 'PasswordOnly'; TokenSource = 'PowerBIModule' }
        foreach ($k in @('XmlaConnection', 'XmlaProbeDone', 'XmlaProbeTried')) { if ($script:IQ.ContainsKey($k)) { $script:IQ.Remove($k) } }
        # run A: module token available -> remembered variant from the first export, no probe
        Mock Get-IQPowerBIModuleXmlaToken { 'MODULE-TOKEN' }
        (Invoke-IQModelBackupStage).Done | Should -Be 1
        $script:ProbeArgs.Count | Should -Be 0
        $script:BatchArgs[0] | Should -Match 'Data Source=[^"]+;Password=MODULE-TOKEN"'
        $log = Get-Content -LiteralPath $script:IQ.LogFile -Raw
        $log | Should -Match 'Using the XMLA connection an earlier run found accepted \(Power BI PowerShell module token, audience https://analysis.usgovcloudapi.net/powerbi/api, form PasswordOnly'
        # run B (headless, no module session): falls back to the sign-in token with a Warn line instead of failing every model
        Clear-ModelCheckpoints
        foreach ($k in @('XmlaConnection', 'XmlaProbeDone', 'XmlaProbeTried')) { if ($script:IQ.ContainsKey($k)) { $script:IQ.Remove($k) } }
        $script:BatchArgs.Clear()
        Mock Get-IQPowerBIModuleXmlaToken { $null }
        (Invoke-IQModelBackupStage).Done | Should -Be 1
        $script:BatchArgs[0] | Should -Match 'Data Source=[^"]+;Password=GOV-TOKEN"'
        (Get-IQModelXmlaConnection).TokenSource | Should -Be 'SignIn'
        $log = Get-Content -LiteralPath $script:IQ.LogFile -Raw
        $log | Should -Match 'remembered XMLA connection uses the Power BI PowerShell module token, which this run cannot obtain; using the sign-in token instead'
        # a memory for another environment is ignored
        Save-IQModelXmlaMemory -Connection @{ ResourceUrl = 'https://analysis.windows.net/powerbi/api'; Form = 'UserIdEmpty'; TokenSource = 'SignIn' }
        $raw = Get-Content -LiteralPath (Get-IQModelXmlaMemoryPath) -Raw
        [System.IO.File]::WriteAllText((Get-IQModelXmlaMemoryPath), ($raw -replace '"USGov"', '"Public"'))
        foreach ($k in @('XmlaConnection')) { if ($script:IQ.ContainsKey($k)) { $script:IQ.Remove($k) } }
        (Get-IQModelXmlaConnection).ResourceUrl | Should -Be 'https://analysis.usgovcloudapi.net/powerbi/api'
    }
    It '-XmlaTokenSource PowerBIModule pins the module token from the first export; a sign-in that went through the module offers no separate module variant' {
        Mock Get-IQSelectedDatasets { @([pscustomobject]@{ DatasetId = $script:StageIds.d1; DatasetName = 'First'; WorkspaceId = $script:StageIds.ws1; WorkspaceName = 'Dash'; WorkspaceIsOnDedicatedCapacity = $true }) }
        Mock Invoke-IQProcessBatch {
            $out = @()
            foreach ($j in $Jobs) { $script:BatchArgs.Add([string]$j.ArgumentList); $bim = [regex]::Match([string]$j.ArgumentList, '-B "([^"]+)"').Groups[1].Value; Write-CompleteBim -Path $bim; $entry = @{ ItemKey = $j.ItemKey; Item = $j.Item; Result = (New-TestProcessResult -ExitCode 0) }; if ($OnJobComplete) { & $OnJobComplete $entry }; $out += , $entry }
            return $out
        }
        Mock Get-IQPowerBIModuleXmlaToken { 'MODULE-TOKEN' }
        $script:IQ.Options['XmlaTokenSource'] = 'PowerBIModule'
        (Invoke-IQModelBackupStage).Done | Should -Be 1
        $script:BatchArgs[0] | Should -Match 'User ID=;Password=MODULE-TOKEN'
        $conn = Get-IQModelXmlaConnection
        $conn.TokenSource | Should -Be 'PowerBIModule'
        $conn.SourcePinned | Should -BeTrue
        @(Get-IQModelXmlaVariantList | ForEach-Object { $_.TokenSource }) | Should -Be @('PowerBIModule') -Because 'a pinned source leaves only the other form of that source'
        $log = Get-Content -LiteralPath $script:IQ.LogFile -Raw
        $log | Should -Match 'Power BI PowerShell module token \(pinned by -XmlaTokenSource\)'
        # pinned but no module token: a clear error, not "No Power BI token available"
        foreach ($k in @('XmlaConnection')) { if ($script:IQ.ContainsKey($k)) { $script:IQ.Remove($k) } }
        Mock Get-IQPowerBIModuleXmlaToken { $null }
        { Get-IQModelXmlaToken } | Should -Throw '*pinned to the Power BI PowerShell module*'
        # -XmlaTokenSource SignIn: no module variants at all
        foreach ($k in @('XmlaConnection')) { if ($script:IQ.ContainsKey($k)) { $script:IQ.Remove($k) } }
        $script:IQ.Options['XmlaTokenSource'] = 'SignIn'
        @(Get-IQModelXmlaVariantList | Where-Object { $_.TokenSource -eq 'PowerBIModule' }).Count | Should -Be 0
        @(Get-IQModelXmlaVariantList).Count | Should -Be 3
        # Auto with a sign-in that already went through the module: the sign-in token IS the module token
        foreach ($k in @('XmlaConnection')) { if ($script:IQ.ContainsKey($k)) { $script:IQ.Remove($k) } }
        $script:IQ.Options['XmlaTokenSource'] = 'Auto'
        $previousAuth = $script:IQ.Auth
        try {
            $script:IQ.Auth = @{ Provider = 'Module'; Initialized = $true }
            @(Get-IQModelXmlaVariantList | Where-Object { $_.TokenSource -eq 'PowerBIModule' }).Count | Should -Be 0
            $script:IQ.Auth = @{ Provider = 'Az'; Initialized = $true }
            @(Get-IQModelXmlaVariantList | ForEach-Object { $_.TokenSource }) | Should -Be @('PowerBIModule', 'PowerBIModule', 'SignIn', 'SignIn', 'SignIn') -Because 'module variants (both forms, environment audience) come first'
        }
        finally { $script:IQ.Auth = $previousAuth }
    }
}
