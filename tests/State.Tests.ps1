# State.Tests.ps1 - ImpactIQ.State.ps1 (brief section 2.4, 5.2): manifest, resume decisions, checkpoints, stage runner.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Base = Initialize-IQTestContext -Options @{ Environment = 'USGov'; RunMode = 'Workspaces'; Credential = 'not-a-real-credential'; TokenCacheKey = 'cache-secret'; MaxRetries = 5; WorkspaceName = @('A*'); DeviceCodeWebhookUrl = 'https://contoso.webhook.office.com/webhookb2/SECRET-GUID/IncomingWebhook/abc' } -NoRun -Prefix 'state'
    function Set-Clock { param([string]$Iso) $script:IQ.Options['NowUtc'] = $Iso }
    function Get-ManifestFromDisk { return (ConvertFrom-IQJsonFile -Path (Join-Path $script:IQ.RunPath 'manifest.json')) }
}
AfterAll { Remove-IQTestFolder -Path $script:Base }

Describe 'Initialize-IQRun (fresh run) and the manifest' {
    BeforeAll {
        Set-Clock '2026-09-04T10:00:00Z'
        $script:M = Initialize-IQRun -RunId '' -ResumePolicy Auto
    }
    It 'defaults the RunId to today (yyyy-MM-dd) and is not a resume' {
        $script:IQ.RunId | Should -Be '2026-09-04'
        $script:IQ.IsResume | Should -BeFalse
        $script:IQ.RunPath | Should -Be (Join-Path (Join-Path (Join-Path $script:Base 'State') 'runs') '2026-09-04')
    }
    It 'writes manifest.json with the section 2.4 schema' {
        (Join-Path $script:IQ.RunPath 'manifest.json') | Should -Exist
        $m = Get-ManifestFromDisk
        $m.schemaVersion | Should -Be 1
        $m.runId | Should -Be '2026-09-04'
        $m.status | Should -Be 'Running'
        $m.environment | Should -Be 'USGov'
        foreach ($key in @('startedUtc', 'updatedUtc', 'endedUtc', 'host', 'auth', 'options', 'scope', 'stages', 'failures', 'outputs')) {
            $m.PSObject.Properties[$key] | Should -Not -BeNullOrEmpty -Because "manifest key $key"
        }
        $m.host.PSObject.Properties['machine'] | Should -Not -BeNullOrEmpty
        $m.host.PSObject.Properties['psVersion'] | Should -Not -BeNullOrEmpty
        foreach ($key in @('environmentWorkbook', 'reportWorkbook', 'modelWorkbook', 'dataflowWorkbook')) { $m.outputs.PSObject.Properties[$key] | Should -Not -BeNullOrEmpty }
    }
    It 'never stores secrets in the manifest options' {
        $json = Get-Content -LiteralPath (Join-Path $script:IQ.RunPath 'manifest.json') -Raw
        $json | Should -Not -Match 'cache-secret'
        $json | Should -Not -Match 'not-a-real-credential'
        $json | Should -Not -Match 'SECRET-GUID' -Because 'an incoming-webhook URL is the credential'
        $script:M.options.DeviceCodeWebhookUrl | Should -Be 'https://contoso.webhook.office.com/***'
        $script:M.options.MaxRetries | Should -Be 5
    }
    It 'creates the backup folders for the run id' {
        (Join-Path (Join-Path $script:Base 'Model Backups') '2026-09-04') | Should -Exist
        (Join-Path (Join-Path $script:Base 'Report Backups') '2026-09-04') | Should -Exist
        (Join-Path (Join-Path $script:Base 'Dataflow Backups') '2026-09-04') | Should -Exist
    }
    It 'rejects a RunId that is not a plain folder name' {
        { Initialize-IQRun -RunId '../escape' -ResumePolicy Never } | Should -Throw
    }
}

Describe 'Item checkpoints and the stage runner' {
    BeforeAll {
        Set-Clock '2026-09-04T10:05:00Z'
        Initialize-IQRun -RunId '2026-09-04' -ResumePolicy Never | Out-Null
        $script:WsFile = Save-IQInventory -Name 'ws-WS1' -Object @{ WorkspaceId = 'WS1'; Datasets = @(@{ DatasetId = 'd1' }, @{ DatasetId = 'd2' }) }
        $script:Status = Invoke-IQStage -Name 'Inventory' -Body {
            Set-IQItemDone -Stage Inventory -ItemKey 'WS1' -Item 'Workspace One' -Outputs @($script:WsFile) -Data @{ Datasets = 2 } | Out-Null
            Set-IQItemDone -Stage Inventory -ItemKey 'WS2' -Item 'Workspace Two' -Status Failed -Message 'HTTP 403 Password=abc;' | Out-Null
            Set-IQItemDone -Stage Inventory -ItemKey 'WS3' -Item 'Workspace Three' -Status Skipped -Message 'Pro workspace' | Out-Null
        }
    }
    It 'reports CompletedWithErrors when an item failed and counts done/failed' {
        $script:Status | Should -Be 'CompletedWithErrors'
        $script:IQ.Manifest.stages.Inventory.status | Should -Be 'CompletedWithErrors'
        [int]$script:IQ.Manifest.stages.Inventory.itemsFailed | Should -Be 1
        [int]$script:IQ.Manifest.stages.Inventory.itemsDone | Should -BeGreaterOrEqual 1
    }
    It 'records the failure in manifest.failures (redacted)' {
        $f = @($script:IQ.Manifest.failures)
        $f.Count | Should -Be 1
        $f[0].stage | Should -Be 'Inventory'
        $f[0].itemKey | Should -Be 'WS2'
        $f[0].item | Should -Be 'Workspace Two'
        $f[0].message | Should -Not -Match 'abc'
        $f[0].timeUtc | Should -Not -BeNullOrEmpty
    }
    It 'writes the checkpoint file done\<Stage>\<safeKey>.json' {
        (Join-Path (Join-Path (Join-Path $script:IQ.RunPath 'done') 'Inventory') 'ws1.json') | Should -Exist
        $cp = Get-IQItemCheckpoint -Stage Inventory -ItemKey 'WS1'
        $cp.status | Should -Be 'Succeeded'
        @($cp.outputs).Count | Should -Be 1
        $cp.data.Datasets | Should -Be 2
    }
    It 'Test-IQItemDone is true for Succeeded and Skipped, false for Failed and unknown' {
        Test-IQItemDone -Stage Inventory -ItemKey 'WS1' | Should -BeTrue
        Test-IQItemDone -Stage Inventory -ItemKey 'WS3' | Should -BeTrue
        Test-IQItemDone -Stage Inventory -ItemKey 'WS2' | Should -BeFalse
        Test-IQItemDone -Stage Inventory -ItemKey 'nope' | Should -BeFalse
        Get-IQItemCheckpoint -Stage Inventory -ItemKey 'nope' | Should -BeNullOrEmpty
    }
    It 'invalidates a checkpoint whose output file is missing' {
        Remove-Item -LiteralPath $script:WsFile -Force
        Test-IQItemDone -Stage Inventory -ItemKey 'WS1' | Should -BeFalse
        Save-IQInventory -Name 'ws-WS1' -Object @{ WorkspaceId = 'WS1'; Datasets = @() } | Out-Null
        Test-IQItemDone -Stage Inventory -ItemKey 'WS1' | Should -BeTrue
    }
    It 'removes the failure entry when the item later succeeds' {
        Set-IQItemDone -Stage Inventory -ItemKey 'WS2' -Item 'Workspace Two' -Status Succeeded | Out-Null
        @($script:IQ.Manifest.failures).Count | Should -Be 0
        [int]$script:IQ.Manifest.stages.Inventory.itemsFailed | Should -Be 0
    }
    It 'Save-IQInventory / Get-IQInventory / Get-IQAllWorkspaceInventories use the inventory folder' {
        Save-IQInventory -Name 'global' -Object @{ Apps = @(1, 2, 3) } | Out-Null
        Save-IQInventory -Name 'ws-WS2' -Object @{ WorkspaceId = 'WS2' } | Out-Null
        (Get-IQInventory -Name 'global').Apps.Count | Should -Be 3
        (Join-Path (Join-Path $script:IQ.RunPath 'inventory') 'global.json') | Should -Exist
        @(Get-IQAllWorkspaceInventories).Count | Should -Be 2
        Get-IQInventory -Name 'does-not-exist' | Should -BeNullOrEmpty
    }
    It 'a stage whose body throws is Failed but does not rethrow unless -Fatal' {
        Invoke-IQStage -Name 'Dataflows' -Body { throw 'non fatal' } | Should -Be 'Failed'
        $script:IQ.Manifest.stages.Dataflows.error | Should -Match 'non fatal'
        { Invoke-IQStage -Name 'ModelBackup' -Body { throw 'fatal!' } -Fatal } | Should -Throw -ExpectedMessage '*fatal!*'
        $script:IQ.Manifest.stages.ModelBackup.status | Should -Be 'Failed'
    }
    It 'a clean stage is Completed with timestamps and $IQ.CurrentStage set during the body' {
        $script:SeenStage = $null
        Invoke-IQStage -Name 'Assemble' -Body { $script:SeenStage = $script:IQ.CurrentStage } | Should -Be 'Completed'
        $script:SeenStage | Should -Be 'Assemble'
        $script:IQ.Manifest.stages.Assemble.startedUtc | Should -Not -BeNullOrEmpty
        $script:IQ.Manifest.stages.Assemble.endedUtc | Should -Not -BeNullOrEmpty
    }
    It 'Complete-IQRun derives CompletedWithErrors and the exit code is 2' {
        Complete-IQRun | Should -Be 'CompletedWithErrors'
        $m = Get-ManifestFromDisk
        $m.status | Should -Be 'CompletedWithErrors'
        $m.endedUtc | Should -Not -BeNullOrEmpty
        Get-IQExitCode -Manifest $script:IQ.Manifest | Should -Be 2
        Get-IQExitCode -Manifest $m | Should -Be 2
    }
    It 'Get-IQRunSummary yields one (Run) row and stage rows in pipeline order' {
        $rows = @(Get-IQRunSummary)
        $rows[0].Stage | Should -Be '(Run)'
        $names = @($rows | Select-Object -Skip 1 | ForEach-Object { $_.Stage })
        [array]::IndexOf($names, 'Inventory') | Should -BeLessThan ([array]::IndexOf($names, 'ModelBackup'))
        [array]::IndexOf($names, 'Dataflows') | Should -BeLessThan ([array]::IndexOf($names, 'Assemble'))
    }
}

Describe 'Resume decisions (brief section 5.2)' {
    It 'Auto: a manifest for the RunId that is not Completed is resumed; Completed stages are skipped, Assemble re-runs' {
        Set-Clock '2026-09-04T11:00:00Z'
        $script:IQ.Manifest.status = 'CompletedWithErrors'; Save-IQManifest
        $m = Initialize-IQRun -RunId '2026-09-04' -ResumePolicy Auto
        $script:IQ.IsResume | Should -BeTrue
        $m.status | Should -Be 'Running'
        $script:IQ.Manifest.stages.Inventory.status = 'Completed'; Save-IQManifest
        Invoke-IQStage -Name 'Inventory' -Body { throw 'must be skipped' } | Should -Be 'Completed'
        Invoke-IQStage -Name 'Assemble' -Body { throw 'assemble always runs' } | Should -Be 'Failed'
        Invoke-IQStage -Name 'Dataflows' -Body { 'ran' } | Should -Be 'Completed' -Because 'a Failed stage is re-run on resume'
    }
    It 'Auto + RefreshInventory re-runs a Completed Inventory stage and clears its checkpoints, ws-*.json / global.json and Inventory failures' {
        # a workspace that disappeared since the original run, and failure entries from Inventory and another stage
        $staleWs = Save-IQInventory -Name 'ws-GONE' -Object @{ WorkspaceId = 'GONE' }
        Set-IQItemDone -Stage Inventory -ItemKey 'GONE' -Item 'Removed workspace' -Status Failed -Message 'HTTP 404' | Out-Null
        Set-IQItemDone -Stage Dataflows -ItemKey 'dfx' -Item 'dfx' -Status Failed -Message 'kept' | Out-Null
        $script:IQ.Options['RefreshInventory'] = $true
        try {
            Invoke-IQStage -Name 'Inventory' -Body { 'ran again' } | Should -Be 'Completed'
            (Join-Path (Join-Path $script:IQ.RunPath 'done') 'Inventory') | Should -Not -Exist
            $staleWs | Should -Not -Exist
            (Join-Path (Join-Path $script:IQ.RunPath 'inventory') 'global.json') | Should -Not -Exist
            @(Get-IQAllWorkspaceInventories).Count | Should -Be 0
            @($script:IQ.Manifest.failures | Where-Object { $_.stage -eq 'Inventory' }).Count | Should -Be 0
            @($script:IQ.Manifest.failures | Where-Object { $_.stage -eq 'Dataflows' }).Count | Should -Be 1 -Because 'only Inventory failures are refreshed'
        }
        finally { $script:IQ.Options['RefreshInventory'] = $false }
        Set-IQItemDone -Stage Dataflows -ItemKey 'dfx' -Item 'dfx' -Status Succeeded | Out-Null
        Complete-IQRun -Status Completed | Out-Null
    }
    It 'Auto without -RunId resumes the newest Running/Paused/Failed run within ResumeMaxAgeDays' {
        $script:IQ.Manifest.status = 'Failed'; Save-IQManifest
        Set-Clock '2026-09-05T08:00:00Z'
        Initialize-IQRun -ResumePolicy Auto -ResumeMaxAgeDays 3 | Out-Null
        $script:IQ.IsResume | Should -BeTrue
        $script:IQ.RunId | Should -Be '2026-09-04'
        Complete-IQRun -Status Completed | Out-Null
    }
    It 'Auto with an explicit -RunId ignores other runs' {
        $script:IQ.Manifest.status = 'Failed'; Save-IQManifest
        Set-Clock '2026-09-05T09:00:00Z'
        Initialize-IQRun -RunId '2026-09-05' -ResumePolicy Auto | Out-Null
        $script:IQ.IsResume | Should -BeFalse
        $script:IQ.RunId | Should -Be '2026-09-05'
        Complete-IQRun -Status Completed | Out-Null
    }
    It 'Auto does not resume a failed run older than ResumeMaxAgeDays' {
        Initialize-IQRun -RunId '2026-09-04' -ResumePolicy Always | Out-Null
        $script:IQ.Manifest.status = 'Failed'; Save-IQManifest
        Set-Clock '2026-09-09T08:00:00Z'
        Initialize-IQRun -ResumePolicy Auto -ResumeMaxAgeDays 3 | Out-Null
        $script:IQ.IsResume | Should -BeFalse
        $script:IQ.RunId | Should -Be '2026-09-09'
        Complete-IQRun -Status Completed | Out-Null
    }
    It 'Auto with a Completed manifest for the RunId starts fresh, archives the manifest and clears the backup folders' {
        $marker = Join-Path $script:IQ.RunPaths.ModelBackups 'old.bim'
        New-Item -ItemType File -Path $marker -Force | Out-Null
        $m = Initialize-IQRun -RunId '2026-09-09' -ResumePolicy Auto
        $script:IQ.IsResume | Should -BeFalse
        $m.status | Should -Be 'Running'
        @(Get-ChildItem -LiteralPath $script:IQ.RunPath -Filter 'manifest.*.json').Count | Should -BeGreaterOrEqual 1
        $marker | Should -Not -Exist
    }
    It 'Always resumes any existing manifest (even Completed) and starts a new manifest without clearing when none exists' {
        Complete-IQRun -Status Completed | Out-Null
        Initialize-IQRun -RunId '2026-09-09' -ResumePolicy Always | Out-Null
        $script:IQ.IsResume | Should -BeTrue
        # "-Stages Assemble" implies Always: backups already exported today (State folder lost) must survive
        $keep = Join-Path (Join-Path $script:IQ.Paths.ModelBackups 'custom-run') 'keep.bim'
        New-Item -ItemType File -Path $keep -Force | Out-Null
        Initialize-IQRun -RunId 'custom-run' -ResumePolicy Always | Out-Null
        $script:IQ.IsResume | Should -BeFalse
        $script:IQ.RunId | Should -Be 'custom-run'
        $keep | Should -Exist
    }
    It 'Never and -Force start fresh and delete the run state' {
        Set-IQItemDone -Stage Extras -ItemKey 'k' -Item 'k' | Out-Null
        Initialize-IQRun -RunId 'custom-run' -ResumePolicy Never | Out-Null
        $script:IQ.IsResume | Should -BeFalse
        (Join-Path (Join-Path $script:IQ.RunPath 'done') 'Extras') | Should -Not -Exist
        Set-IQItemDone -Stage Extras -ItemKey 'k' -Item 'k' | Out-Null
        Initialize-IQRun -RunId 'custom-run' -ResumePolicy Auto -Force | Out-Null
        $script:IQ.IsResume | Should -BeFalse
        (Join-Path (Join-Path $script:IQ.RunPath 'done') 'Extras') | Should -Not -Exist
    }
    It 'a resumed run never clears the backup folders' {
        $marker = Join-Path $script:IQ.RunPaths.ReportBackups 'keep.pbix'
        New-Item -ItemType File -Path $marker -Force | Out-Null
        $script:IQ.Manifest.status = 'Failed'; Save-IQManifest
        Initialize-IQRun -RunId 'custom-run' -ResumePolicy Auto | Out-Null
        $script:IQ.IsResume | Should -BeTrue
        $marker | Should -Exist
    }
}

Describe 'Time budget: Paused stages, Paused run, exit code 3, resume' {
    BeforeAll {
        Set-Clock '2026-09-10T10:00:00Z'
        $script:IQ.Options['TimeBudgetMinutes'] = 0
        $script:IQ.BudgetExceeded = $false
        # a dated RunId: only auto-generated RunIds are adopted by a start without -RunId (last test of this block)
        Initialize-IQRun -RunId '2026-09-10' -ResumePolicy Never | Out-Null
        $script:IQ.Options['TimeBudgetMinutes'] = 30
        $script:IQ.StartedUtc = [datetime]::UtcNow
    }
    AfterAll {
        $script:IQ.Options['TimeBudgetMinutes'] = 0
        $script:IQ.BudgetExceeded = $false
        $script:IQ.StartedUtc = [datetime]::UtcNow
    }
    It 'a stage whose body stops on the budget is Paused (not Completed) and the rest of the items stay unchecked' {
        $script:BudgetSeen = @()
        $status = Invoke-IQStage -Name 'Dataflows' -Body {
            foreach ($k in @('df1', 'df2', 'df3')) {
                if (Test-IQTimeBudget -Stage 'Dataflows' -Item $k) { break }
                Set-IQItemDone -Stage Dataflows -ItemKey $k -Item $k | Out-Null
                $script:BudgetSeen += $k
                # the budget "runs out" after the first item
                $script:IQ.StartedUtc = [datetime]::UtcNow.AddMinutes(-29)
            }
        }
        $status | Should -Be 'Paused'
        $script:BudgetSeen | Should -Be @('df1')
        $script:IQ.Manifest.stages.Dataflows.status | Should -Be 'Paused'
        [int]$script:IQ.Manifest.stages.Dataflows.itemsDone | Should -Be 1
        $script:IQ.BudgetExceeded | Should -BeTrue
    }
    It 'later stages are marked Paused without running, but Assemble still runs' {
        $script:AssembleRan = $false
        Invoke-IQStage -Name 'Extras' -Body { throw 'must not run' } | Should -Be 'Paused'
        $script:IQ.Manifest.stages.Extras.status | Should -Be 'Paused'
        Invoke-IQStage -Name 'Assemble' -Body { $script:AssembleRan = $true } | Should -Be 'Completed'
        $script:AssembleRan | Should -BeTrue
        $script:IQ.Manifest.stages.Assemble.status | Should -Be 'Completed'
    }
    It 'Complete-IQRun marks the run Paused and the exit code is 3' {
        Complete-IQRun | Should -Be 'Paused'
        $m = Get-ManifestFromDisk
        $m.status | Should -Be 'Paused'
        Get-IQExitCode -Manifest $script:IQ.Manifest | Should -Be 3
        Get-IQExitCode -Manifest $m | Should -Be 3
    }
    It 'Auto resume treats a Paused run like Running (same RunId) and re-runs the Paused stages, skipping done items' {
        $script:IQ.BudgetExceeded = $false
        $script:IQ.StartedUtc = [datetime]::UtcNow
        Initialize-IQRun -RunId '2026-09-10' -ResumePolicy Auto | Out-Null
        $script:IQ.IsResume | Should -BeTrue
        $script:IQ.Manifest.status | Should -Be 'Running'
        Test-IQItemDone -Stage Dataflows -ItemKey 'df1' | Should -BeTrue
        $script:Ran = @()
        Invoke-IQStage -Name 'Dataflows' -Body {
            foreach ($k in @('df1', 'df2', 'df3')) {
                if (Test-IQItemDone -Stage Dataflows -ItemKey $k) { continue }
                Set-IQItemDone -Stage Dataflows -ItemKey $k -Item $k | Out-Null
                $script:Ran += $k
            }
        } | Should -Be 'Completed'
        $script:Ran | Should -Be @('df2', 'df3')
        Invoke-IQStage -Name 'Extras' -Body { 'ok' } | Should -Be 'Completed'
        Complete-IQRun | Should -Be 'Completed'
        Get-IQExitCode -Manifest $script:IQ.Manifest | Should -Be 0
    }
    It 'a budget that trips only inside Assemble (every stage Completed) does not pause the run' {
        $script:IQ.BudgetExceeded = $true
        try { Get-IQRunStatus | Should -Be 'Completed' } finally { $script:IQ.BudgetExceeded = $false }
    }
    It 'Auto without -RunId also picks up a Paused run within ResumeMaxAgeDays' {
        $script:IQ.Manifest.status = 'Paused'; Save-IQManifest
        Set-Clock '2026-09-11T10:00:00Z'
        Initialize-IQRun -RunId '' -ResumePolicy Auto | Out-Null
        $script:IQ.RunId | Should -Be '2026-09-10'
        $script:IQ.IsResume | Should -BeTrue
        Complete-IQRun -Status Completed | Out-Null
    }
}

Describe 'Crash recovery: an Interrupted stage keeps the run resumable (stage-filtered resume)' {
    BeforeAll {
        Set-Clock '2026-10-10T08:00:00Z'
        $script:IQ.BudgetExceeded = $false
        Initialize-IQRun -RunId '2026-10-10' -ResumePolicy Never | Out-Null
        Invoke-IQStage -Name 'Inventory' -Body { Set-IQItemDone -Stage Inventory -ItemKey 'WS1' -Item 'WS1' | Out-Null } | Out-Null
        # a hard kill in the middle of ModelBackup: one item done, one failed, the stage left "Running" on disk
        $entry = Get-IQStageEntry -Name 'ModelBackup'
        $entry.status = 'Running'
        $entry.startedUtc = Get-IQUtcStamp
        Set-IQItemDone -Stage ModelBackup -ItemKey 'm1' -Item 'Model 1' | Out-Null
        Set-IQItemDone -Stage ModelBackup -ItemKey 'm2' -Item 'Model 2' -Status Failed -Message 'XMLA timeout' | Out-Null
        Save-IQManifest
    }
    It 'marks the stage Interrupted on resume; a run that does not re-run it ends Paused (exit 3), never Completed' {
        Set-Clock '2026-10-10T09:00:00Z'
        Initialize-IQRun -RunId '2026-10-10' -ResumePolicy Always | Out-Null
        $script:IQ.IsResume | Should -BeTrue
        $script:IQ.Manifest.stages.ModelBackup.status | Should -Be 'Interrupted'
        Invoke-IQStage -Name 'Assemble' -Body { 'assemble only' } | Should -Be 'Completed'
        Complete-IQRun | Should -Be 'Paused'
        $m = Get-ManifestFromDisk
        $m.status | Should -Be 'Paused'
        Get-IQExitCode -Manifest $script:IQ.Manifest | Should -Be 3
        Get-IQExitCode -Manifest $m | Should -Be 3
    }
    It 'stays resumable: the next start without -RunId adopts it and re-runs the Interrupted stage' {
        (Find-IQResumableRun -MaxAgeDays 3).RunId | Should -Be '2026-10-10'
        Set-Clock '2026-10-11T08:00:00Z'
        Initialize-IQRun -RunId '' -ResumePolicy Auto | Out-Null
        $script:IQ.RunId | Should -Be '2026-10-10'
        $script:IQ.IsResume | Should -BeTrue
        $script:Ran = $false
        Invoke-IQStage -Name 'ModelBackup' -Body {
            $script:Ran = $true
            Set-IQItemDone -Stage ModelBackup -ItemKey 'm2' -Item 'Model 2' | Out-Null
        } | Should -Be 'Completed'
        $script:Ran | Should -BeTrue
        Invoke-IQStage -Name 'Assemble' -Body { 'ok' } | Should -Be 'Completed'
        Complete-IQRun | Should -Be 'Completed'
        Get-IQExitCode -Manifest $script:IQ.Manifest | Should -Be 0
    }
    It 'timestamps loaded from disk are ISO-8601 strings on both hosts (RunSummary / Failures never culture-formatted)' {
        Set-IQItemDone -Stage ModelBackup -ItemKey 'm3' -Item 'Model 3' -Status Failed -Message 'left failed' | Out-Null
        Initialize-IQRun -RunId '2026-10-10' -ResumePolicy Always | Out-Null
        $iso = '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}'
        $script:IQ.Manifest['startedUtc'] | Should -BeOfType [string]
        $script:IQ.Manifest['startedUtc'] | Should -Match $iso
        $script:IQ.Manifest['stages']['Inventory']['endedUtc'] | Should -BeOfType [string]
        $failure = @($script:IQ.Manifest['failures'])[0]
        $failure['timeUtc'] | Should -BeOfType [string]
        $failure['timeUtc'] | Should -Match $iso
        foreach ($row in @(Get-IQRunSummary)) {
            $row.StartedUtc | Should -BeOfType [string]
            if ($row.StartedUtc) { $row.StartedUtc | Should -Match $iso -Because "row $($row.Stage)" }
            if ($row.EndedUtc) { $row.EndedUtc | Should -Match $iso -Because "row $($row.Stage)" }
        }
        Set-IQItemDone -Stage ModelBackup -ItemKey 'm3' -Item 'Model 3' | Out-Null
        Complete-IQRun -Status Completed | Out-Null
    }
}

Describe 'Resume validates checkpoints: a Completed stage with missing outputs is re-run' {
    BeforeAll {
        Set-Clock '2026-10-12T08:00:00Z'
        Initialize-IQRun -RunId '2026-10-12' -ResumePolicy Never | Out-Null
        $script:Pbix = Join-Path $script:IQ.RunPaths.ReportBackups 'r1.pbix'
        Invoke-IQStage -Name 'ReportBackup' -Body {
            New-Item -ItemType File -Path $script:Pbix -Force | Out-Null
            Set-IQItemDone -Stage ReportBackup -ItemKey 'r1' -Item 'Report 1' -Outputs @($script:Pbix) | Out-Null
            Set-IQItemDone -Stage ReportBackup -ItemKey 'r2' -Item 'Report 2' -Status Skipped -Message 'no access' | Out-Null
        } | Out-Null
        Complete-IQRun | Out-Null
    }
    It 'skips a Completed stage while every checkpoint output still exists' {
        Initialize-IQRun -RunId '2026-10-12' -ResumePolicy Always | Out-Null
        Invoke-IQStage -Name 'ReportBackup' -Body { throw 'must be skipped' } | Should -Be 'Completed'
    }
    It 're-runs the stage when a checkpoint output is missing (backups not restored on this agent), intact items still skipped' {
        Remove-Item -LiteralPath $script:Pbix -Force
        Initialize-IQRun -RunId '2026-10-12' -ResumePolicy Always | Out-Null
        $script:Seen = @()
        Invoke-IQStage -Name 'ReportBackup' -Body {
            foreach ($k in @('r1', 'r2')) {
                if (Test-IQItemDone -Stage ReportBackup -ItemKey $k) { continue }
                $script:Seen += $k
                New-Item -ItemType File -Path $script:Pbix -Force | Out-Null
                Set-IQItemDone -Stage ReportBackup -ItemKey $k -Item $k -Outputs @($script:Pbix) | Out-Null
            }
        } | Should -Be 'Completed'
        $script:Seen | Should -Be @('r1')
        Complete-IQRun | Should -Be 'Completed'
    }
}

Describe 'Cross-RunId adoption rules (Find-IQResumableRun, Auto without -RunId)' {
    It 'does not adopt a run that finished as CompletedWithErrors; naming that RunId still resumes it' {
        Set-Clock '2026-10-20T08:00:00Z'
        Initialize-IQRun -RunId '2026-10-20' -ResumePolicy Never | Out-Null
        Complete-IQRun -Status CompletedWithErrors | Out-Null
        Set-Clock '2026-10-21T08:00:00Z'
        Find-IQResumableRun -MaxAgeDays 3 | Should -BeNullOrEmpty
        Initialize-IQRun -RunId '' -ResumePolicy Auto | Out-Null
        $script:IQ.RunId | Should -Be '2026-10-21' -Because 'the daily pipeline must produce a fresh dated snapshot'
        $script:IQ.IsResume | Should -BeFalse
        Complete-IQRun -Status Completed | Out-Null
        Initialize-IQRun -RunId '2026-10-20' -ResumePolicy Auto | Out-Null
        $script:IQ.IsResume | Should -BeTrue
        Complete-IQRun -Status Completed | Out-Null
    }
    It 'does not adopt an ad-hoc (non-dated) RunId; it is resumed only when named' {
        Set-Clock '2026-10-22T08:00:00Z'
        Initialize-IQRun -RunId 'smoke' -ResumePolicy Never | Out-Null
        $script:IQ.Manifest.status = 'Failed'; Save-IQManifest
        Set-Clock '2026-10-23T08:00:00Z'
        Find-IQResumableRun -MaxAgeDays 3 | Should -BeNullOrEmpty
        Initialize-IQRun -RunId '' -ResumePolicy Auto | Out-Null
        $script:IQ.RunId | Should -Be '2026-10-23'
        $script:IQ.IsResume | Should -BeFalse
        Complete-IQRun -Status Completed | Out-Null
        Initialize-IQRun -RunId 'smoke' -ResumePolicy Auto | Out-Null
        $script:IQ.IsResume | Should -BeTrue
        Complete-IQRun -Status Completed | Out-Null
    }
    It 'does not adopt a run recorded for another environment' {
        Set-Clock '2026-10-24T08:00:00Z'
        Initialize-IQRun -RunId '2026-10-24' -ResumePolicy Never | Out-Null
        $script:IQ.Manifest.status = 'Failed'; $script:IQ.Manifest.environment = 'Public'; Save-IQManifest
        Set-Clock '2026-10-25T08:00:00Z'
        Find-IQResumableRun -MaxAgeDays 3 | Should -BeNullOrEmpty
        $script:IQ.Manifest.environment = 'USGov'; Save-IQManifest
        (Find-IQResumableRun -MaxAgeDays 3).RunId | Should -Be '2026-10-24'
        Initialize-IQRun -RunId '2026-10-24' -ResumePolicy Auto | Out-Null
        Complete-IQRun -Status Completed | Out-Null
    }
}

Describe 'Fresh run over a run folder whose manifest is unreadable' {
    It 'archives the unreadable manifest and clears stale checkpoints and inventory files' {
        Set-Clock '2026-10-30T08:00:00Z'
        Initialize-IQRun -RunId '2026-10-30' -ResumePolicy Never | Out-Null
        Set-IQItemDone -Stage Extras -ItemKey 'probe' -Item 'probe' -Status Skipped | Out-Null
        Save-IQInventory -Name 'ws-OLD' -Object @{ WorkspaceId = 'OLD' } | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $script:IQ.RunPath 'manifest.json'), '{ not json')
        Initialize-IQRun -RunId '2026-10-30' -ResumePolicy Auto | Out-Null
        $script:IQ.IsResume | Should -BeFalse
        Test-IQItemDone -Stage Extras -ItemKey 'probe' | Should -BeFalse
        @(Get-IQAllWorkspaceInventories).Count | Should -Be 0
        @(Get-ChildItem -LiteralPath $script:IQ.RunPath -Filter 'manifest.*.json').Count | Should -BeGreaterOrEqual 1
        Complete-IQRun -Status Completed | Out-Null
    }
}
