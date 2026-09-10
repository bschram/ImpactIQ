# Pipeline.Tests.ps1 - pipelines/templates/impactiq-run.yml: the inline PowerShell scripts of the Azure DevOps template.
# The scripts are extracted from the YAML by their step displayName and executed in a child PowerShell process (they use
# `exit`) against a temporary BaseFolder / artifact staging directory, with the pipeline's env: mappings simulated as
# environment variables - including the literal "$(NAME)" text Azure Pipelines leaves for an undefined variable.
# Covered contracts: undefined secrets are dropped before ImpactIQ runs; the state artifact carries the backup folders
# of resumable runs and forgets runs older than stateRetentionDays; the restore step moves those backups next to State;
# exit codes 2/3 map to SucceededWithIssues (step exit 0) and 1 fails the step.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
$script:Template = Join-Path (Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'pipelines') 'templates') 'impactiq-run.yml'
$script:HasTemplate = Test-Path -LiteralPath $script:Template
$script:HasAnalyzer = [bool](Get-Module -ListAvailable -Name PSScriptAnalyzer | Select-Object -First 1)
$script:StepNames = @('Restore run backups next to State', 'Install PowerShell modules', 'Run ImpactIQ', 'Stage artifacts and run summary', 'Commit workbooks to branch', 'Copy workbooks to SharePoint/OneDrive sync folder')

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Template = Join-Path (Join-Path (Join-Path (Get-IQTestRepoRoot) 'pipelines') 'templates') 'impactiq-run.yml'
    $script:HasTemplate = Test-Path -LiteralPath $script:Template
    $script:HasAnalyzer = [bool](Get-Module -ListAvailable -Name PSScriptAnalyzer | Select-Object -First 1)
    $script:PwshExe = (Get-Process -Id $PID).Path
    if ([string]::IsNullOrWhiteSpace($script:PwshExe)) { $script:PwshExe = 'pwsh' }
    $script:TemplateLines = @()
    if ($script:HasTemplate) { $script:TemplateLines = @([System.IO.File]::ReadAllLines($script:Template)) }
    $script:StepNames = @('Restore run backups next to State', 'Install PowerShell modules', 'Run ImpactIQ', 'Stage artifacts and run summary', 'Commit workbooks to branch', 'Copy workbooks to SharePoint/OneDrive sync folder')
    $script:EnvNames = @('IMPACTIQ_P_BASEFOLDER', 'IMPACTIQ_P_ENVIRONMENT', 'IMPACTIQ_P_AUTHMODE', 'IMPACTIQ_P_RUNMODE', 'IMPACTIQ_P_WORKSPACENAMES', 'IMPACTIQ_P_STAGES',
        'IMPACTIQ_P_EXTRAARGS', 'IMPACTIQ_P_TIMEBUDGETMINUTES', 'IMPACTIQ_P_SKIPTOOLUPDATE', 'IMPACTIQ_P_USEPWSH', 'IMPACTIQ_P_PUBLISHBACKUPS', 'IMPACTIQ_P_STATERETENTIONDAYS',
        'BUILD_ARTIFACTSTAGINGDIRECTORY', 'BUILD_BUILDNUMBER', 'TF_BUILD', 'AGENT_NAME', 'IQTEST_EXIT')

    function Get-PipelineScript {
        # The dedented `script: |` block of the first step whose displayName starts with -DisplayName.
        param([string]$DisplayName)
        $start = -1
        for ($i = 0; $i -lt $script:TemplateLines.Count; $i++) {
            if ($script:TemplateLines[$i] -match ('^\s*displayName:\s*' + [regex]::Escape($DisplayName))) { $start = $i; break }
        }
        if ($start -lt 0) { throw "Step '$DisplayName' not found in $($script:Template)" }
        $keyIndent = -1
        $body = New-Object System.Collections.Generic.List[string]
        for ($i = $start + 1; $i -lt $script:TemplateLines.Count; $i++) {
            $line = $script:TemplateLines[$i]
            if ($keyIndent -lt 0) {
                if ($line -match '^(\s*)script:\s*\|\s*$') { $keyIndent = $Matches[1].Length }
                elseif ($line -match '^\s*-\s*(task|template|checkout|\$\{\{)') { throw "Step '$DisplayName' has no inline script" }
                continue
            }
            if ($line.Trim().Length -eq 0) { $body.Add(''); continue }
            $indent = $line.Length - $line.TrimStart().Length
            if ($indent -le $keyIndent) { break }
            $body.Add($line)
        }
        if ($keyIndent -lt 0) { throw "Step '$DisplayName' has no 'script: |' block" }
        $minIndent = [int]::MaxValue
        foreach ($l in $body) { if ($l.Length -gt 0) { $n = $l.Length - $l.TrimStart().Length; if ($n -lt $minIndent) { $minIndent = $n } } }
        $out = New-Object System.Collections.Generic.List[string]
        foreach ($l in $body) { if ($l.Length -gt 0) { $out.Add($l.Substring($minIndent)) } else { $out.Add('') } }
        return ($out -join [Environment]::NewLine)
    }

    function Invoke-PipelineScript {
        # Writes the step script to a file and runs it in a child pwsh with -Environment (name -> value) set; returns
        # @{ ExitCode; Output }. Every variable in $script:EnvNames plus IMPACTIQ_* is cleared afterwards.
        param([string]$DisplayName, [string]$Folder, [hashtable]$Environment)
        $file = Join-Path $Folder ('step-' + ($DisplayName -replace '[^A-Za-z0-9]+', '-').ToLowerInvariant() + '.ps1')
        [System.IO.File]::WriteAllText($file, (Get-PipelineScript -DisplayName $DisplayName), (New-Object System.Text.UTF8Encoding($false)))
        $oldPath = $env:PATH
        try {
            foreach ($k in $Environment.Keys) { [System.Environment]::SetEnvironmentVariable($k, [string]$Environment[$k], 'Process') }
            # the run step starts `pwsh` by name; make sure the child finds the host that runs the tests
            $env:PATH = (Split-Path -Parent $script:PwshExe) + [System.IO.Path]::PathSeparator + $oldPath
            $output = & $script:PwshExe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $file 2>&1 | ForEach-Object { [string]$_ }
            $code = $LASTEXITCODE
        }
        finally {
            $env:PATH = $oldPath
            foreach ($k in $script:EnvNames) { [System.Environment]::SetEnvironmentVariable($k, $null, 'Process') }
            foreach ($k in @($Environment.Keys)) { [System.Environment]::SetEnvironmentVariable($k, $null, 'Process') }
            Reset-IQTestEnvironment
        }
        return @{ ExitCode = $code; Output = ($output -join [Environment]::NewLine) }
    }

    function New-RunFolder {
        # State\runs\<RunId> with a manifest of the given status, updated -DaysAgo days ago.
        param([string]$Base, [string]$RunId, [string]$Status, [double]$DaysAgo)
        $folder = Join-Path (Join-Path (Join-Path $Base 'State') 'runs') $RunId
        New-Item -ItemType Directory -Force -Path (Join-Path $folder 'done') | Out-Null
        $stamp = (Get-Date).ToUniversalTime().AddDays(-$DaysAgo).ToString('o')
        $manifest = [ordered]@{ runId = $RunId; status = $Status; environment = 'USGov'; auth = 'DeviceCode'; startedUtc = $stamp; updatedUtc = $stamp; endedUtc = $null
            stages = [ordered]@{ Inventory = [ordered]@{ status = 'Completed'; itemsDone = 1; itemsFailed = 0; error = $null } }; failures = @() }
        [System.IO.File]::WriteAllText((Join-Path $folder 'manifest.json'), ($manifest | ConvertTo-Json -Depth 20), (New-Object System.Text.UTF8Encoding($false)))
        [System.IO.File]::WriteAllText((Join-Path (Join-Path $folder 'done') 'x.json'), '{"status":"Succeeded"}')
        return $folder
    }

    function New-BackupFile {
        param([string]$Base, [string]$Kind, [string]$RunId, [string]$Name)
        $folder = Join-Path (Join-Path $Base $Kind) $RunId
        New-Item -ItemType Directory -Force -Path $folder | Out-Null
        [System.IO.File]::WriteAllText((Join-Path $folder $Name), "backup $Name")
    }
}

Describe 'impactiq-run.yml inline scripts' -Skip:(-not $script:HasTemplate) {
    It 'declares the stateRetentionDays parameter and passes it to the staging step' {
        $text = Get-Content -LiteralPath $script:Template -Raw
        $text | Should -Match '(?m)^\s*- name: stateRetentionDays\s*$'
        $text | Should -Match 'IMPACTIQ_P_STATERETENTIONDAYS: \$\{\{ parameters\.stateRetentionDays \}\}'
    }
    It 'step "<_>" parses without syntax errors and uses no PowerShell 7-only syntax' -ForEach $script:StepNames {
        $text = Get-PipelineScript -DisplayName $_
        $tokens = $null; $errors = $null
        [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$errors) | Out-Null
        @($errors).Count | Should -Be 0 -Because (@($errors | ForEach-Object { "$($_.Extent.StartLineNumber): $($_.Message)" }) -join '; ')
        $text | Should -Not -Match '\?\?'
        $text | Should -Not -Match '\?\.'
        $text | Should -Not -Match '-Parallel'
        $text | Should -Not -Match '\$IsWindows'
    }
    It 'step "<_>" is clean under PSScriptAnalyzer with the repository settings' -ForEach $script:StepNames -Skip:(-not $script:HasAnalyzer) {
        $folder = New-IQTestBaseFolder -Prefix 'pipe-pssa'
        try {
            $file = Join-Path $folder 'step.ps1'
            [System.IO.File]::WriteAllText($file, (Get-PipelineScript -DisplayName $_), (New-Object System.Text.UTF8Encoding($false)))
            $settings = Join-Path (Get-IQTestRepoRoot) 'PSScriptAnalyzerSettings.psd1'
            $results = @(Invoke-ScriptAnalyzer -Path $file -Settings $settings)
            $blocking = @($results | Where-Object { $_.Severity -eq 'Error' -or $_.RuleName -in @('PSUseCompatibleSyntax', 'PSUseCompatibleCommands') })
            $blocking.Count | Should -Be 0 -Because (@($blocking | ForEach-Object { "$($_.Line): $($_.RuleName) $($_.Message)" }) -join '; ')
        }
        finally { Remove-IQTestFolder -Path $folder }
    }
}

Describe 'Run ImpactIQ step' -Skip:(-not $script:HasTemplate) {
    BeforeAll {
        function New-RunBase {
            # A BaseFolder whose ImpactIQ.ps1 records the IMPACTIQ_* environment it received and exits with $env:IQTEST_EXIT.
            $base = New-IQTestBaseFolder -Prefix 'pipe-run'
            $fake = @(
                'param([string]$BaseFolder, [switch]$NonInteractive, [string]$Environment, [string]$AuthMode, [string]$RunMode, [string]$Resume, [switch]$AllWorkspaces, [string[]]$WorkspaceName, [string[]]$Stages, [int]$TimeBudgetMinutes, [switch]$SkipToolUpdate)',
                '$seen = [ordered]@{}',
                'foreach ($v in (Get-ChildItem Env: | Where-Object { $_.Name -like "IMPACTIQ_*" } | Sort-Object Name)) { $seen[$v.Name] = [string]$v.Value }',
                '$seen["ARGS"] = "BaseFolder=$BaseFolder;Environment=$Environment;AuthMode=$AuthMode;Resume=$Resume;AllWorkspaces=$AllWorkspaces;NonInteractive=$NonInteractive"',
                '[System.IO.File]::WriteAllText((Join-Path $BaseFolder "env-dump.json"), ($seen | ConvertTo-Json -Depth 20))',
                'exit ([int]$env:IQTEST_EXIT)'
            )
            [System.IO.File]::WriteAllText((Join-Path $base 'ImpactIQ.ps1'), ($fake -join [Environment]::NewLine))
            return $base
        }
        function Get-RunEnvironment {
            param([string]$Base, [int]$ExitCode)
            return @{
                IMPACTIQ_P_BASEFOLDER = $Base; IMPACTIQ_P_ENVIRONMENT = 'USGov'; IMPACTIQ_P_AUTHMODE = 'DeviceCode'; IMPACTIQ_P_RUNMODE = 'Workspaces'
                IMPACTIQ_P_WORKSPACENAMES = '*'; IMPACTIQ_P_STAGES = ''; IMPACTIQ_P_EXTRAARGS = ''; IMPACTIQ_P_TIMEBUDGETMINUTES = '0'
                IMPACTIQ_P_SKIPTOOLUPDATE = 'False'; IMPACTIQ_P_USEPWSH = 'True'; IMPACTIQ_ENVIRONMENT = 'USGov'; IQTEST_EXIT = [string]$ExitCode
                # what Azure Pipelines hands over when the secret variables are not defined (macro syntax is left verbatim)
                IMPACTIQ_USERNAME = '$(IMPACTIQ_USERNAME)'; IMPACTIQ_PASSWORD = '$(IMPACTIQ_PASSWORD)'; IMPACTIQ_TENANT_ID = '$(IMPACTIQ_TENANT_ID)'
                IMPACTIQ_CLIENT_ID = '$(IMPACTIQ_CLIENT_ID)'; IMPACTIQ_PBI_TOKEN = '$(IMPACTIQ_PBI_TOKEN)'
                # a defined one
                IMPACTIQ_TOKEN_CACHE_KEY = 'unit-test-cache-key-0123456789abcdef'
            }
        }
    }
    It 'drops undefined "$(NAME)" secrets before ImpactIQ.ps1 runs and keeps the defined ones' {
        $base = New-RunBase
        try {
            $r = Invoke-PipelineScript -DisplayName 'Run ImpactIQ' -Folder $base -Environment (Get-RunEnvironment -Base $base -ExitCode 0)
            $r.ExitCode | Should -Be 0 -Because $r.Output
            $dumpPath = Join-Path $base 'env-dump.json'
            Test-Path -LiteralPath $dumpPath | Should -BeTrue -Because $r.Output
            $seen = Get-Content -LiteralPath $dumpPath -Raw | ConvertFrom-Json
            $names = @($seen.PSObject.Properties.Name)
            foreach ($n in @('IMPACTIQ_USERNAME', 'IMPACTIQ_PASSWORD', 'IMPACTIQ_TENANT_ID', 'IMPACTIQ_CLIENT_ID', 'IMPACTIQ_PBI_TOKEN')) { $names | Should -Not -Contain $n }
            $seen.IMPACTIQ_TOKEN_CACHE_KEY | Should -Be 'unit-test-cache-key-0123456789abcdef'
            $seen.IMPACTIQ_ENVIRONMENT | Should -Be 'USGov'
            $seen.ARGS | Should -Match 'Environment=USGov;AuthMode=DeviceCode;Resume=Auto;AllWorkspaces=True;NonInteractive=True'
            $r.Output | Should -Match 'USERNAME=False'
            $r.Output | Should -Match 'TOKEN_CACHE_KEY=True'
            $r.Output | Should -Match 'Pipeline variables not defined \(ignored\):.*IMPACTIQ_TENANT_ID'
        }
        finally { Remove-IQTestFolder -Path $base }
    }
    It 'maps exit 3 (paused) to SucceededWithIssues and exit 1 to a failed step' {
        $base = New-RunBase
        try {
            $paused = Invoke-PipelineScript -DisplayName 'Run ImpactIQ' -Folder $base -Environment (Get-RunEnvironment -Base $base -ExitCode 3)
            $paused.ExitCode | Should -Be 0 -Because $paused.Output
            $paused.Output | Should -Match 'ImpactIQ exit code: 3'
            $paused.Output | Should -Match 'task\.complete result=SucceededWithIssues'
            $paused.Output | Should -Match 'task\.setvariable variable=ImpactIQ\.ExitCode\]3'
            $failed = Invoke-PipelineScript -DisplayName 'Run ImpactIQ' -Folder $base -Environment (Get-RunEnvironment -Base $base -ExitCode 1)
            $failed.ExitCode | Should -Be 1 -Because $failed.Output
            $failed.Output | Should -Match 'task\.logissue type=error\]ImpactIQ failed \(exit 1\)'
        }
        finally { Remove-IQTestFolder -Path $base }
    }
}

Describe 'Stage artifacts step' -Skip:(-not $script:HasTemplate) {
    BeforeAll {
        function New-StageBase {
            # Three runs: a Completed one 30 days old, a Completed one 2 days old, a Paused one 1 day old; backups for the
            # last two; a DPAPI/AES token cache; one workbook; a log file.
            $base = New-IQTestBaseFolder -Prefix 'pipe-stage'
            New-RunFolder -Base $base -RunId '2026-07-01' -Status 'Completed' -DaysAgo 30 | Out-Null
            New-RunFolder -Base $base -RunId '2026-07-28' -Status 'Completed' -DaysAgo 2 | Out-Null
            New-RunFolder -Base $base -RunId '2026-07-30' -Status 'Paused' -DaysAgo 1 | Out-Null
            New-BackupFile -Base $base -Kind 'Model Backups' -RunId '2026-07-28' -Name 'Finished.bim'
            New-BackupFile -Base $base -Kind 'Model Backups' -RunId '2026-07-30' -Name 'Sales.bim'
            New-BackupFile -Base $base -Kind 'Model Backups' -RunId '2026-07-30' -Name 'Sales.csv'
            New-BackupFile -Base $base -Kind 'Report Backups' -RunId '2026-07-30' -Name 'Sales.txt'
            New-BackupFile -Base $base -Kind 'Dataflow Backups' -RunId '2026-07-30' -Name 'Flow.pq'
            $auth = Join-Path (Join-Path $base 'State') 'auth'
            New-Item -ItemType Directory -Force -Path $auth | Out-Null
            [System.IO.File]::WriteAllText((Join-Path $auth 'token-cache.json'), '{"encrypted":"x"}')
            [System.IO.File]::WriteAllText((Join-Path $base 'Model Detail.xlsx'), 'not really a workbook')
            New-Item -ItemType Directory -Force -Path (Join-Path $base 'Logs') | Out-Null
            [System.IO.File]::WriteAllText((Join-Path (Join-Path $base 'Logs') 'ImpactIQ.log'), 'log')
            New-Item -ItemType Directory -Force -Path (Join-Path $base 'staging') | Out-Null
            return $base
        }
        function Get-StageEnvironment {
            param([string]$Base, [string]$RetentionDays, [string]$CacheKey)
            return @{
                IMPACTIQ_P_BASEFOLDER = $Base; IMPACTIQ_P_PUBLISHBACKUPS = 'False'; IMPACTIQ_P_STATERETENTIONDAYS = $RetentionDays
                BUILD_ARTIFACTSTAGINGDIRECTORY = (Join-Path $Base 'staging'); BUILD_BUILDNUMBER = '20260731.1'; IMPACTIQ_TOKEN_CACHE_KEY = $CacheKey
            }
        }
    }
    It 'stages resumable runs with their backup folders, prunes old Completed runs and ignores an undefined cache key' {
        $base = New-StageBase
        try {
            $r = Invoke-PipelineScript -DisplayName 'Stage artifacts and run summary' -Folder $base -Environment (Get-StageEnvironment -Base $base -RetentionDays '7' -CacheKey '$(IMPACTIQ_TOKEN_CACHE_KEY)')
            $r.ExitCode | Should -Be 0 -Because $r.Output
            $state = Join-Path (Join-Path (Join-Path $base 'staging') 'impactiq') 'state'
            $runs = Join-Path $state 'runs'
            Test-Path -LiteralPath (Join-Path (Join-Path $runs '2026-07-30') 'manifest.json') | Should -BeTrue -Because $r.Output
            Test-Path -LiteralPath (Join-Path (Join-Path $runs '2026-07-28') 'manifest.json') | Should -BeTrue
            Test-Path -LiteralPath (Join-Path $runs '2026-07-01') | Should -BeFalse -Because 'a Completed run updated 30 days ago is outside stateRetentionDays=7'
            $backups = Join-Path $state 'backups'
            foreach ($rel in @(@('Model Backups', 'Sales.bim'), @('Model Backups', 'Sales.csv'), @('Report Backups', 'Sales.txt'), @('Dataflow Backups', 'Flow.pq'))) {
                Test-Path -LiteralPath (Join-Path (Join-Path (Join-Path $backups $rel[0]) '2026-07-30') $rel[1]) | Should -BeTrue -Because "the Paused run needs $($rel[1]) to resume"
            }
            Test-Path -LiteralPath (Join-Path (Join-Path $backups 'Model Backups') '2026-07-28') | Should -BeFalse -Because 'a Completed run is never resumed'
            Test-Path -LiteralPath (Join-Path $state 'auth') | Should -BeFalse -Because 'an unexpanded $(IMPACTIQ_TOKEN_CACHE_KEY) is not a key'
            $r.Output | Should -Match 'Token cache NOT published'
            Test-Path -LiteralPath (Join-Path (Join-Path (Join-Path (Join-Path $base 'staging') 'impactiq') 'outputs') 'Model Detail.xlsx') | Should -BeTrue
            Test-Path -LiteralPath (Join-Path (Join-Path (Join-Path (Join-Path $base 'staging') 'impactiq') 'logs') 'ImpactIQ.log') | Should -BeTrue
            $summary = Join-Path (Join-Path (Join-Path $base 'staging') 'impactiq') 'impactiq-summary.md'
            Test-Path -LiteralPath $summary | Should -BeTrue
            (Get-Content -LiteralPath $summary -Raw) | Should -Match '\| 2026-07-30 \| Paused \|'
            $r.Output | Should -Match 'task\.uploadsummary'
        }
        finally { Remove-IQTestFolder -Path $base }
    }
    It 'keeps every run with stateRetentionDays 0 and publishes the token cache when a key is defined' {
        $base = New-StageBase
        try {
            $r = Invoke-PipelineScript -DisplayName 'Stage artifacts and run summary' -Folder $base -Environment (Get-StageEnvironment -Base $base -RetentionDays '0' -CacheKey 'a-real-key-a-real-key-a-real-key')
            $r.ExitCode | Should -Be 0 -Because $r.Output
            $state = Join-Path (Join-Path (Join-Path $base 'staging') 'impactiq') 'state'
            Test-Path -LiteralPath (Join-Path (Join-Path (Join-Path $state 'runs') '2026-07-01') 'manifest.json') | Should -BeTrue -Because $r.Output
            Test-Path -LiteralPath (Join-Path (Join-Path $state 'auth') 'token-cache.json') | Should -BeTrue
            $r.Output | Should -Match 'Token cache included'
        }
        finally { Remove-IQTestFolder -Path $base }
    }
}

Describe 'Restore run backups step' -Skip:(-not $script:HasTemplate) {
    It 'moves the restored backups\Kind\RunId folders next to State, merging with files already there' {
        $base = New-IQTestBaseFolder -Prefix 'pipe-restore'
        try {
            $restored = Join-Path (Join-Path $base 'State') 'backups'
            $src = Join-Path (Join-Path (Join-Path $restored 'Model Backups') '2026-07-30') 'sub'
            New-Item -ItemType Directory -Force -Path $src | Out-Null
            [System.IO.File]::WriteAllText((Join-Path $src 'Sales.bim'), 'restored')
            $reportSrc = Join-Path (Join-Path $restored 'Report Backups') '2026-07-30'
            New-Item -ItemType Directory -Force -Path $reportSrc | Out-Null
            [System.IO.File]::WriteAllText((Join-Path $reportSrc 'Sales.txt'), 'restored')
            New-BackupFile -Base $base -Kind 'Model Backups' -RunId '2026-07-30' -Name 'Local.bim'
            $r = Invoke-PipelineScript -DisplayName 'Restore run backups next to State' -Folder $base -Environment @{ IMPACTIQ_P_BASEFOLDER = $base }
            $r.ExitCode | Should -Be 0 -Because $r.Output
            $modelRun = Join-Path (Join-Path $base 'Model Backups') '2026-07-30'
            Get-Content -LiteralPath (Join-Path (Join-Path $modelRun 'sub') 'Sales.bim') -Raw | Should -Be 'restored'
            Test-Path -LiteralPath (Join-Path $modelRun 'Local.bim') | Should -BeTrue -Because 'files already on the agent are kept'
            Test-Path -LiteralPath (Join-Path (Join-Path (Join-Path $base 'Report Backups') '2026-07-30') 'Sales.txt') | Should -BeTrue
            Test-Path -LiteralPath $restored | Should -BeFalse -Because 'the staging copy is removed after the merge'
            $r.Output | Should -Match 'Restored 2 backup file'
        }
        finally { Remove-IQTestFolder -Path $base }
    }
    It 'is a no-op when the restored state has no backups' {
        $base = New-IQTestBaseFolder -Prefix 'pipe-restore0'
        try {
            $r = Invoke-PipelineScript -DisplayName 'Restore run backups next to State' -Folder $base -Environment @{ IMPACTIQ_P_BASEFOLDER = $base }
            $r.ExitCode | Should -Be 0 -Because $r.Output
            $r.Output | Should -Match 'No backup folders'
        }
        finally { Remove-IQTestFolder -Path $base }
    }
}
