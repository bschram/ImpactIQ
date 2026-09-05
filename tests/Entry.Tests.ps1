# Entry.Tests.ps1 - ImpactIQ.ps1 (brief section 5.1): parses, refuses a scope-less headless run, assembles from a prepared state.
# The entry point is executed in a child PowerShell process (it uses exit codes); network cmdlets are shadowed there
# with global functions so no real HTTP call can happen.
# Discovery-time values (Pester evaluates -Skip before BeforeAll runs).
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
$script:Entry = Join-Path (Split-Path -Parent $PSScriptRoot) 'ImpactIQ.ps1'
$script:HasEntry = Test-Path -LiteralPath $script:Entry
$script:HasImportExcel = [bool](Get-Module -ListAvailable -Name ImportExcel | Select-Object -First 1)

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    # top-level (discovery) variables are not visible in the run phase: recompute what the run needs
    $script:Entry = Join-Path (Get-IQTestRepoRoot) 'ImpactIQ.ps1'
    $script:HasEntry = Test-Path -LiteralPath $script:Entry
    $script:HasImportExcel = [bool](Get-Module -ListAvailable -Name ImportExcel | Select-Object -First 1)
    $script:Ids = Get-IQTestFixtureJson -Relative 'ids.json'
    $script:PwshExe = (Get-Process -Id $PID).Path
    if ([string]::IsNullOrWhiteSpace($script:PwshExe)) { $script:PwshExe = 'pwsh' }

    function New-EntryBase {
        # A BaseFolder that also carries a copy of Config\Modules so the entry point can load them from either location.
        param([string]$Prefix)
        $base = New-IQTestBaseFolder -Prefix $Prefix
        $modules = Join-Path (Join-Path $base 'Config') 'Modules'
        New-Item -ItemType Directory -Path $modules -Force | Out-Null
        Get-ChildItem -LiteralPath (Join-Path (Get-IQTestRepoRoot) 'Config/Modules') -Filter '*.ps1' | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $modules $_.Name) -Force }
        return $base
    }
    function Invoke-Entry {
        # Runs ImpactIQ.ps1 in a child process with Invoke-WebRequest / Invoke-RestMethod shadowed; returns @{ ExitCode; Output }.
        # -Parameters is splatted onto the entry point (switches as $true) so no command-line quoting is involved.
        param([string]$Base, [hashtable]$Parameters)
        $runner = Join-Path $Base 'run-entry.ps1'
        $splat = New-Object System.Collections.Generic.List[string]
        foreach ($key in $Parameters.Keys) {
            $value = $Parameters[$key]
            if ($value -is [bool]) { $splat.Add(('    {0} = ${1}' -f $key, $value.ToString().ToLowerInvariant())) }
            elseif ($value -is [array]) { $splat.Add(('    {0} = @({1})' -f $key, (($value | ForEach-Object { "'" + ([string]$_ -replace "'", "''") + "'" }) -join ', '))) }
            else { $splat.Add(('    {0} = ''{1}''' -f $key, ([string]$value -replace "'", "''"))) }
        }
        $lines = @(
            '$ErrorActionPreference = ''Continue''',
            'function global:Invoke-WebRequest { param($Uri, $Method, $Headers, $Body, $ContentType, $OutFile, $TimeoutSec, $UserAgent, [switch]$UseBasicParsing, [switch]$PassThru) return [pscustomobject]@{ StatusCode = 200; Headers = @{ ''Content-Type'' = ''application/json'' }; Content = ''{"value":[]}'' } }',
            'function global:Invoke-RestMethod { param($Uri, $Method, $Headers, $Body, $ContentType, $TimeoutSec, [switch]$UseBasicParsing) throw ''test shim: network is disabled'' }',
            '$entryParameters = @{'
        ) + @($splat.ToArray()) + @(
            '}',
            ('& "{0}" @entryParameters' -f $script:Entry),
            'exit $LASTEXITCODE'
        )
        [System.IO.File]::WriteAllText($runner, ($lines -join [Environment]::NewLine))
        $env:IMPACTIQ_PBI_TOKEN = New-IQTestJwt -Claims @{ upn = 'entry@contoso.gov' } -ExpiresInMinutes 90
        try {
            $output = & $script:PwshExe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $runner 2>&1 | ForEach-Object { [string]$_ }
            $code = $LASTEXITCODE
        }
        finally { Reset-IQTestEnvironment }
        return @{ ExitCode = $code; Output = ($output -join [Environment]::NewLine) }
    }
}

Describe 'ImpactIQ.ps1 parses and declares the headless parameters' -Skip:(-not $script:HasEntry) {
    BeforeAll {
        $script:Tokens = $null; $script:Errors = $null
        $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Entry, [ref]$script:Tokens, [ref]$script:Errors)
    }
    It 'has no syntax errors' {
        @($script:Errors).Count | Should -Be 0 -Because (@($script:Errors | ForEach-Object { "$($_.Extent.StartLineNumber): $($_.Message)" }) -join '; ')
    }
    It 'declares [CmdletBinding()] and the section 5.1 parameters' {
        $params = @($script:Ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        foreach ($p in @('BaseFolder', 'Environment', 'AuthMode', 'TenantId', 'ClientId', 'Credential', 'TokenCachePath', 'TokenCacheKey', 'DeviceCodeWebhookUrl', 'NonInteractive', 'RunMode', 'WorkspaceId', 'WorkspaceName', 'AllWorkspaces', 'IncludeMyWorkspace', 'ReportId', 'DatasetId', 'Stages', 'SkipStages', 'RunId', 'Resume', 'ResumeMaxAgeDays', 'Force', 'RefreshInventory', 'ModelDetailMethod', 'MaxParallelExtracts', 'ToolTimeoutMinutes', 'MaxRetries', 'SkipToolUpdate', 'IncludeAdminApis', 'IncludeUsageMetrics', 'ActivityDays', 'LogPath', 'PassThru')) {
            $params | Should -Contain $p
        }
        @($script:Ast.ParamBlock.Attributes | Where-Object { $_.TypeName.Name -eq 'CmdletBinding' }).Count | Should -Be 1
    }
    It 'does not use PowerShell 7-only syntax or $IsWindows' {
        $text = Get-Content -LiteralPath $script:Entry -Raw
        $text | Should -Not -Match '\?\?'
        $text | Should -Not -Match '\?\.\w'
        $text | Should -Not -Match '(?m)^\s*using namespace'
        $text | Should -Not -Match '\$IsWindows'
        $text | Should -Not -Match '-Parallel'
    }
}

Describe 'Headless run without a scope' -Skip:(-not $script:HasEntry) {
    BeforeAll {
        $script:Base1 = New-EntryBase -Prefix 'entry-noscope'
        $script:Run1 = Invoke-Entry -Base $script:Base1 -Parameters @{ BaseFolder = $script:Base1; NonInteractive = $true; Environment = 'Public'; AuthMode = 'AccessToken'; SkipToolUpdate = $true; Stages = @('Inventory'); RunId = 'entry-noscope'; Resume = 'Never' }
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base1 }
    It 'exits with a non-zero code and says that no scope was given' {
        $script:Run1.ExitCode | Should -Not -Be 0 -Because $script:Run1.Output
        $script:Run1.Output | Should -Match '(?i)no scope'
    }
    It 'writes the error to a log file and never leaves a run in Running state' {
        $logs = @(Get-ChildItem -LiteralPath (Join-Path $script:Base1 'Logs') -Filter '*.log' -Recurse)
        $logs.Count | Should -BeGreaterOrEqual 1
        (($logs | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n") | Should -Match '(?i)no scope'
        # the entry point may fail fast before a run/manifest exists; when one exists it must not be left Running
        $manifest = Join-Path (Join-Path (Join-Path (Join-Path $script:Base1 'State') 'runs') 'entry-noscope') 'manifest.json'
        if (Test-Path -LiteralPath $manifest) { (ConvertFrom-IQJsonFile -Path $manifest).status | Should -Match 'Failed|CompletedWithErrors' }
    }
}

Describe '-Stages Assemble on a prepared state folder' -Skip:(-not ($script:HasEntry -and $script:HasImportExcel)) {
    BeforeAll {
        Mock Get-IQToken { 'test-token' }
        Mock Invoke-IQApi { Invoke-IQTestApiFixture -Method $Method -Path $Path -Body $Body -Query $Query -Api $Api -Raw:$Raw }
        $script:Base2 = New-EntryBase -Prefix 'entry-assemble'
        Initialize-IQContext -BaseFolder $script:Base2 -Options @{ Environment = 'Public'; NonInteractive = $true; RunMode = 'Workspaces'; WorkspaceId = @($script:Ids.ws1) } | Out-Null
        $script:IQ.Interactive = $false
        Initialize-IQRun -RunId 'entry-assemble' -ResumePolicy Never | Out-Null
        Invoke-IQStage -Name Inventory -Body { Invoke-IQInventoryStage | Out-Null } | Out-Null
        Copy-IQTestFixtureFolder -Relative 'extracts/model-detail' -Destination $script:IQ.RunPaths.ModelBackups
        Complete-IQRun -Status Completed | Out-Null
        $script:Run2 = Invoke-Entry -Base $script:Base2 -Parameters @{ BaseFolder = $script:Base2; NonInteractive = $true; Environment = 'Public'; AuthMode = 'AccessToken'; SkipToolUpdate = $true; Stages = @('Assemble'); RunId = 'entry-assemble'; Resume = 'Always' }
    }
    AfterAll { Remove-IQTestFolder -Path $script:Base2 }
    It 'exits 0 and produces the four workbooks' {
        $script:Run2.ExitCode | Should -Be 0 -Because $script:Run2.Output
        foreach ($f in @('Power BI Environment Detail.xlsx', 'Report Detail.xlsx', 'Model Detail.xlsx', 'Dataflow Detail.xlsx')) { (Join-Path $script:Base2 $f) | Should -Exist }
    }
    It 'records the outputs in the manifest and marks the run Completed' {
        $manifest = ConvertFrom-IQJsonFile -Path (Join-Path (Join-Path (Join-Path (Join-Path $script:Base2 'State') 'runs') 'entry-assemble') 'manifest.json')
        $manifest.status | Should -Be 'Completed'
        $manifest.outputs.modelWorkbook | Should -Match 'Model Detail\.xlsx$'
        $manifest.stages.Assemble.status | Should -Be 'Completed'
        $manifest.stages.Inventory.status | Should -Be 'Completed' -Because 'the prepared inventory stage is kept on resume'
    }
    It 'the assembled Model Detail workbook contains the prepared CSV rows' {
        @(Import-Excel -Path (Join-Path $script:Base2 'Model Detail.xlsx') -WorksheetName 'Semantic Models').Count | Should -Be 5
    }
}
