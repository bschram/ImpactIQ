# Reports.Tests.ps1 - ImpactIQ.Reports.ps1 (brief section 8.1, 8.2): the ReportBackup / ReportDetail bookkeeping that
# must hold on resumed runs (review findings R-01, R-05, R-06). Neither the Power BI service nor Tabular Editor is
# touched: downloads and the csx runs are mocked so the tests are host-independent (Linux CI and Windows PowerShell 5.1).
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Ids = Get-IQTestFixtureJson -Relative 'ids.json'
    Mock Get-IQToken { 'test-token' }

    function New-TestTeResult {
        param([bool]$Success = $true, [int]$ExitCode = 0)
        return @{ ExitCode = $ExitCode; TimedOut = $false; StdOut = ''; StdErr = ''; OutFile = $null; ErrFile = $null; DurationSec = 1; Success = $Success; ErrorLines = @(); FailureReason = '' }
    }
    function Write-TestTxt {
        # What a csx run leaves behind: header-only Visuals.txt / Pages.txt in the folder named by IMPACTIQ_DATE_FOLDER.
        param([string]$Folder)
        if (-not (Test-Path -LiteralPath $Folder)) { New-Item -ItemType Directory -Path $Folder -Force | Out-Null }
        [System.IO.File]::WriteAllText((Join-Path $Folder 'Visuals.txt'), "ReportName`tVisualId`r`n")
        [System.IO.File]::WriteAllText((Join-Path $Folder 'Pages.txt'), "ReportName`tPageName`r`n")
    }
    function Write-TestPbix {
        param([string]$Path)
        [System.IO.File]::WriteAllBytes($Path, [byte[]](1..64))
    }
    function Get-Checkpoint { param([string]$Stage, [string]$Key) return (Get-IQItemCheckpoint -Stage $Stage -ItemKey $Key) }
    function Set-ManifestStageStatus {
        param([string]$Stage, [string]$Status)
        $script:IQ.Manifest['stages'][$Stage]['status'] = $Status
        Save-IQManifest
    }
    # Mocks shared by the ReportDetail runs: Tabular Editor "works" and the csx scripts write into IMPACTIQ_DATE_FOLDER.
    function Enable-TestTabularEditor {
        $script:IQ.Tools = @{ TabularEditorWorks = $true; TabularEditorPath = (Join-Path $script:IQ.BaseFolder 'TabularEditor.exe'); TabularEditorPreflight = $null; PbiToolsWorks = $false; PbiToolsPath = '' }
        $script:TeCalls = 0
        Mock Test-IQReportTabularEditorAvailable { $true }
        Mock Invoke-IQTabularEditor {
            $script:TeCalls++
            Write-TestTxt -Folder $env:IMPACTIQ_DATE_FOLDER
            return (New-TestTeResult)
        }
    }
}

Describe 'ReportBackup retried on resume => ReportDetail re-runs (R-01)' {
    BeforeAll {
        $script:Base = Initialize-IQTestContext -Prefix 'reports-resume' -RunId '2026-09-04'
        $script:RunFolder = Get-IQReportRunFolder
        $script:Reports = @(
            [pscustomobject]@{ ReportId = $script:Ids.r1; ReportName = 'Sales'; ReportType = 'PowerBIReport'; ReportWebUrl = 'https://app.powerbi.com/groups/x/reports/y'; WorkspaceId = $script:Ids.ws1; WorkspaceName = 'Finance'; DatasetId = $script:Ids.d1; DatasetName = 'Model'; DatasetWorkspaceId = $script:Ids.ws1 }
        )
        $script:Workspaces = @([pscustomobject]@{ WorkspaceId = $script:Ids.ws1; WorkspaceName = 'Finance'; WorkspaceIsOnDedicatedCapacity = $true })
        Mock Get-IQSelectedReports { $script:Reports }
        Mock Get-IQSelectedWorkspaces { $script:Workspaces }
        Mock Get-IQSelectedDatasets { @() }
        Mock Get-IQReportsWithSensitivityLabel { @{} }
        Mock Invoke-IQDownload { Write-TestPbix -Path $OutFile; return $true }

        # Run 1 (simulated): the export of r1 failed transiently, ReportDetail completed over the (empty) folder.
        Write-TestTxt -Folder $script:RunFolder
        Set-IQItemDone -Stage ReportDetail -ItemKey 'all' -Item 'Report Detail' -Outputs @((Join-Path $script:RunFolder 'Visuals.txt'), (Join-Path $script:RunFolder 'Pages.txt')) -Method 'TabularEditor' -Message '0 PBIX file(s) processed' -Data @{ PbixCount = 0; PbixFiles = @() } | Out-Null
        Set-ManifestStageStatus -Stage 'ReportDetail' -Status 'Completed'
        Set-IQItemDone -Stage ReportBackup -ItemKey $script:Ids.r1 -Item 'Finance ~ Sales' -Status Failed -Method 'LiveConnect' -Message 'earlier attempt: LiveConnect export failed' | Out-Null
        Set-ManifestStageStatus -Stage 'ReportBackup' -Status 'CompletedWithErrors'
        $script:IQ.IsResume = $true
        $script:BodyRuns = 0
        $script:SkipBefore = Invoke-IQStage -Name 'ReportDetail' -Body { $script:BodyRuns++ }
        # Run 2: -Resume re-enters ReportBackup (CompletedWithErrors) and the retried export now succeeds.
        $script:Summary = Invoke-IQReportBackupStage
    }
    AfterAll { $script:IQ.IsResume = $false; Remove-IQTestFolder -Path $script:Base }

    It 'before the retry the Completed ReportDetail stage is skipped by Invoke-IQStage (the resume path under test)' {
        $script:SkipBefore | Should -Be 'Completed'
        $script:BodyRuns | Should -Be 0
    }
    It 'the retried report is exported and checkpointed Succeeded' {
        $script:Summary.Done | Should -Be 1
        $script:Summary.Failed | Should -Be 0
        (Get-Checkpoint -Stage ReportBackup -Key $script:Ids.r1).status | Should -Be 'Succeeded'
        (Join-Path $script:RunFolder 'Finance ~ Sales.pbix') | Should -Exist
    }
    It 'the new export invalidates the ReportDetail checkpoint and its Completed manifest status' {
        Get-Checkpoint -Stage ReportDetail -Key 'all' | Should -BeNullOrEmpty
        $script:IQ.Manifest['stages']['ReportDetail']['status'] | Should -Be 'Pending'
        Test-IQItemDone -Stage ReportDetail -ItemKey 'all' | Should -BeFalse
    }
    It 'Invoke-IQStage now runs the ReportDetail body instead of skipping it' {
        Invoke-IQStage -Name 'ReportDetail' -Body { $script:BodyRuns++ } | Should -Be 'Completed'
        $script:BodyRuns | Should -Be 1
    }
    It 'ReportDetail then processes the retried PBIX and records the file names it saw' {
        Enable-TestTabularEditor
        $r = Invoke-IQReportDetailStage
        $r.Status | Should -Be 'Succeeded'
        $r.PbixCount | Should -Be 1
        $script:TeCalls | Should -Be 2 -Because 'the PBIR script and the classic script both run'
        $cp = Get-Checkpoint -Stage ReportDetail -Key 'all'
        $cp.status | Should -Be 'Succeeded'
        @($cp.data.PbixFiles) | Should -Be @('finance ~ sales.pbix')
    }
    It 'an unchanged PBIX set is skipped through the checkpoint' {
        Enable-TestTabularEditor
        (Invoke-IQReportDetailStage).Status | Should -Be 'AlreadyDone'
        $script:TeCalls | Should -Be 0
    }
    It 'a PBIX that appears after the extraction (without going through ReportBackup) makes the checkpoint stale' {
        Enable-TestTabularEditor
        Write-TestPbix -Path (Join-Path $script:RunFolder 'Finance ~ Restored.pbix')
        $r = Invoke-IQReportDetailStage
        $r.Status | Should -Be 'Succeeded'
        $script:TeCalls | Should -Be 2
        @((Get-Checkpoint -Stage ReportDetail -Key 'all').data.PbixFiles).Count | Should -Be 2
        # ... and a removed one as well.
        Remove-Item -LiteralPath (Join-Path $script:RunFolder 'Finance ~ Restored.pbix') -Force
        (Invoke-IQReportDetailStage).Status | Should -Be 'Succeeded'
        $script:TeCalls | Should -Be 4
        @((Get-Checkpoint -Stage ReportDetail -Key 'all').data.PbixFiles) | Should -Be @('finance ~ sales.pbix')
    }
    It 'a checkpoint written before PbixFiles existed falls back to the count' {
        $current = Test-IQReportDetailCheckpointCurrent -RunFolder $script:RunFolder
        $current.Current | Should -BeTrue
        $cpPath = Join-Path (Join-Path (Join-Path $script:IQ.RunPath 'done') 'ReportDetail') 'all.json'
        $cp = ConvertFrom-IQJsonFile -Path $cpPath
        $cp.data.PSObject.Properties.Remove('PbixFiles')
        $cp.data.PbixCount = 3
        ConvertTo-IQJsonFile -Object $cp -Path $cpPath
        $stale = Test-IQReportDetailCheckpointCurrent -RunFolder $script:RunFolder
        $stale.Current | Should -BeFalse
        $stale.Reason | Should -Match '3 PBIX/PBIT file\(s\) were processed, 1 are in the run folder now'
    }
}

Describe 'ReportDetail runs without a working-folder link (R-05)' {
    BeforeAll {
        $script:Base = Initialize-IQTestContext -Prefix 'reports-link' -RunId '2026-09-04'
        $script:RunFolder = Get-IQReportRunFolder
        # A newer dated folder exists, so the run folder is not the newest one and a link would normally be created.
        $script:Newer = Join-Path ([string]$script:IQ.Paths.ReportBackups) '2026-09-05'
        New-Item -ItemType Directory -Path $script:Newer -Force | Out-Null
        Write-TestPbix -Path (Join-Path $script:RunFolder 'WS ~ Report.pbix')
        Mock New-IQReportDetailWorkingFolder { $null }
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base }

    It 'falls back to BaseFolder as the working directory and still succeeds (IMPACTIQ_DATE_FOLDER selects the run folder)' {
        Enable-TestTabularEditor
        $r = Invoke-IQReportDetailStage
        $r.Status | Should -Be 'Succeeded'
        $script:TeCalls | Should -Be 2
        (Join-Path $script:RunFolder 'Visuals.txt') | Should -Exist
        $cp = Get-Checkpoint -Stage ReportDetail -Key 'all'
        $cp.status | Should -Be 'Succeeded'
        $cp.data.LinkFallback | Should -BeTrue
        [System.IO.Path]::GetFullPath($cp.data.WorkingDirectory).TrimEnd('/', '\') | Should -Be ([System.IO.Path]::GetFullPath($script:IQ.BaseFolder).TrimEnd('/', '\'))
        Should -Invoke New-IQReportDetailWorkingFolder -Times 1 -Exactly
    }
    It 'detects csx scripts that ignored IMPACTIQ_DATE_FOLDER and wrote into the newest folder instead' {
        Reset-IQReportDetailCheckpoint -Reason 'test' | Should -BeTrue
        Enable-TestTabularEditor
        Mock Invoke-IQTabularEditor {
            $script:TeCalls++
            Write-TestTxt -Folder $script:Newer   # the wrong folder
            return (New-TestTeResult)
        }
        $r = Invoke-IQReportDetailStage
        $r.Status | Should -Be 'Failed'
        $r.Message | Should -Match 'instead of the run folder'
        (Get-Checkpoint -Stage ReportDetail -Key 'all').status | Should -Be 'Failed'
    }
}

Describe 'ReportDetail clears the append-only csx files before a re-run (R-06)' {
    BeforeAll {
        $script:Base = Initialize-IQTestContext -Prefix 'reports-stale' -RunId '2026-09-04'
        $script:RunFolder = Get-IQReportRunFolder
        Write-TestPbix -Path (Join-Path $script:RunFolder 'WS ~ Report.pbix')
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base }

    It 'lists ExtractErrors.txt and ReportObjects_UnusedObjects.txt as stale files but not as required outputs' {
        $stale = @(Get-IQReportDetailStaleFileName)
        $stale | Should -Contain 'ExtractErrors.txt'
        $stale | Should -Contain 'ReportObjects_UnusedObjects.txt'
        @(Get-IQReportDetailTxtName) | Should -Not -Contain 'ExtractErrors.txt'
    }
    It 'removes a stale ExtractErrors.txt (and the classic script leftovers) before the csx scripts run' {
        $errors = Join-Path $script:RunFolder 'ExtractErrors.txt'
        $unused = Join-Path $script:RunFolder 'ReportObjects_UnusedObjects.txt'
        $visuals = Join-Path $script:RunFolder 'Visuals.txt'
        [System.IO.File]::WriteAllText($errors, "ReportName`tScript`tStage`tError`tReportDate`r`nOld`tPBIR`tparse`tstale row`t2026-09-04`r`n")
        [System.IO.File]::WriteAllText($unused, "stale`r`n")
        [System.IO.File]::WriteAllText($visuals, "stale visuals`r`n")
        Enable-TestTabularEditor
        $script:SeenAtRun = @{}
        Mock Invoke-IQTabularEditor {
            $script:TeCalls++
            $script:SeenAtRun['errors'] = Test-Path -LiteralPath (Join-Path $env:IMPACTIQ_DATE_FOLDER 'ExtractErrors.txt')
            $script:SeenAtRun['unused'] = Test-Path -LiteralPath (Join-Path $env:IMPACTIQ_DATE_FOLDER 'ReportObjects_UnusedObjects.txt')
            Write-TestTxt -Folder $env:IMPACTIQ_DATE_FOLDER
            return (New-TestTeResult)
        }
        (Invoke-IQReportDetailStage).Status | Should -Be 'Succeeded'
        $script:SeenAtRun['errors'] | Should -BeFalse -Because 'the csx would append below the stale rows'
        $script:SeenAtRun['unused'] | Should -BeFalse
        $errors | Should -Not -Exist
        [System.IO.File]::ReadAllText($visuals) | Should -Not -Match 'stale'
    }
}
