<#
.SYNOPSIS
Runs the ImpactIQ quality gate: parse every script, PSScriptAnalyzer, then the Pester 5 suite; exits non-zero on failure.
.DESCRIPTION
Steps (each can be skipped with a switch):
  1. Parse ImpactIQ.ps1, Config\Modules\*.ps1 and tests\*.ps1 with the PowerShell parser and report syntax errors.
  2. Invoke-ScriptAnalyzer with ..\PSScriptAnalyzerSettings.psd1 on the same files. Errors and the
     PSUseCompatibleSyntax / PSUseCompatibleCommands warnings (Windows PowerShell 5.1 + PowerShell 7 compatibility)
     fail the run; other warnings are listed.
  3. Invoke-Pester (5.7.1) on tests\*.Tests.ps1 with detailed (CI style) output and an NUnit XML result file.
Exit code: 0 when everything passed, 1 otherwise. Works on Windows PowerShell 5.1 and PowerShell 7 (Linux/Windows).
.PARAMETER TestName
One or more test file stems (e.g. Common, Http) to run instead of the whole suite.
.PARAMETER SkipParse
Skip the parser step.
.PARAMETER SkipAnalyzer
Skip PSScriptAnalyzer (for example when the module is not installed).
.PARAMETER SkipPester
Skip the Pester run (parse + analyzer only).
.PARAMETER ResultsPath
Folder for the NUnit result file (default tests\TestResults).
.PARAMETER PassThru
Return the Pester result object in addition to writing the summary.
.EXAMPLE
pwsh -NoProfile -File ./tests/Invoke-Tests.ps1
.EXAMPLE
powershell.exe -NoProfile -File .\tests\Invoke-Tests.ps1 -TestName Common,State -SkipAnalyzer
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)][string[]]$TestName,
    [Parameter(Mandatory = $false)][switch]$SkipParse,
    [Parameter(Mandatory = $false)][switch]$SkipAnalyzer,
    [Parameter(Mandatory = $false)][switch]$SkipPester,
    [Parameter(Mandatory = $false)][string]$ResultsPath,
    [Parameter(Mandatory = $false)][switch]$PassThru
)

$ErrorActionPreference = 'Stop'
$testsRoot = $PSScriptRoot
$repoRoot = Split-Path -Path $testsRoot -Parent
$settingsPath = Join-Path $repoRoot 'PSScriptAnalyzerSettings.psd1'
if ([string]::IsNullOrWhiteSpace($ResultsPath)) { $ResultsPath = Join-Path $testsRoot 'TestResults' }
$failed = $false

function Write-Section {
    <#
    .SYNOPSIS
    Prints a section header.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Text)
    Write-Host ''
    Write-Host ('=== ' + $Text + ' ===') -ForegroundColor Cyan
}

function Get-IQTestScriptFile {
    <#
    .SYNOPSIS
    Lists every .ps1 file the gate checks: the entry point, the modules, the launcher-independent test scripts.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Root)
    $files = New-Object System.Collections.Generic.List[string]
    $entry = Join-Path $Root 'ImpactIQ.ps1'
    if (Test-Path -LiteralPath $entry) { $files.Add($entry) }
    $modules = Join-Path (Join-Path $Root 'Config') 'Modules'
    if (Test-Path -LiteralPath $modules) { Get-ChildItem -LiteralPath $modules -Filter '*.ps1' | Sort-Object Name | ForEach-Object { $files.Add($_.FullName) } }
    Get-ChildItem -LiteralPath (Join-Path $Root 'tests') -Filter '*.ps1' | Sort-Object Name | ForEach-Object { $files.Add($_.FullName) }
    return $files.ToArray()
}

$scriptFiles = @(Get-IQTestScriptFile -Root $repoRoot)

# --- 1. parse -------------------------------------------------------------------------------------------------------
if (-not $SkipParse) {
    Write-Section 'Parse'
    $parseErrors = 0
    foreach ($file in $scriptFiles) {
        $tokens = $null
        $errors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$errors) | Out-Null
        $relative = $file.Substring($repoRoot.Length).TrimStart('\', '/')
        if ($errors -and $errors.Count -gt 0) {
            $parseErrors += $errors.Count
            Write-Host ('  FAIL ' + $relative) -ForegroundColor Red
            foreach ($e in $errors) { Write-Host ('       line ' + $e.Extent.StartLineNumber + ': ' + $e.Message) -ForegroundColor Red }
        }
        else { Write-Host ('  ok   ' + $relative) -ForegroundColor Green }
    }
    if ($parseErrors -gt 0) { $failed = $true; Write-Host ("Parse: {0} syntax error(s)" -f $parseErrors) -ForegroundColor Red }
    else { Write-Host ('Parse: {0} file(s) clean' -f $scriptFiles.Count) -ForegroundColor Green }
    if (-not (Test-Path -LiteralPath (Join-Path $repoRoot 'ImpactIQ.ps1'))) { Write-Host '  note: ImpactIQ.ps1 not found - entry-point checks are skipped' -ForegroundColor Yellow }
}

# --- 2. PSScriptAnalyzer ---------------------------------------------------------------------------------------------
if (-not $SkipAnalyzer) {
    Write-Section 'PSScriptAnalyzer'
    $analyzer = Get-Module -ListAvailable -Name PSScriptAnalyzer | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $analyzer) {
        Write-Host 'PSScriptAnalyzer is not installed (Install-Module PSScriptAnalyzer -Scope CurrentUser); step skipped' -ForegroundColor Yellow
    }
    else {
        Import-Module PSScriptAnalyzer -ErrorAction Stop
        $compatRules = @('PSUseCompatibleSyntax', 'PSUseCompatibleCommands')
        $blocking = New-Object System.Collections.Generic.List[object]
        $advisory = New-Object System.Collections.Generic.List[object]
        foreach ($file in $scriptFiles) {
            $results = @()
            if (Test-Path -LiteralPath $settingsPath) { $results = @(Invoke-ScriptAnalyzer -Path $file -Settings $settingsPath -ErrorAction Stop) }
            else { $results = @(Invoke-ScriptAnalyzer -Path $file -Severity Error, Warning -ErrorAction Stop) }
            foreach ($r in $results) {
                if ($r.Severity -eq 'Error' -or ($compatRules -contains $r.RuleName)) { $blocking.Add($r) } else { $advisory.Add($r) }
            }
        }
        foreach ($r in $blocking) {
            Write-Host ('  FAIL {0}:{1} {2} ({3}) {4}' -f (Split-Path -Leaf $r.ScriptPath), $r.Line, $r.RuleName, $r.Severity, $r.Message) -ForegroundColor Red
        }
        foreach ($r in $advisory) {
            Write-Host ('  warn {0}:{1} {2} {3}' -f (Split-Path -Leaf $r.ScriptPath), $r.Line, $r.RuleName, $r.Message) -ForegroundColor Yellow
        }
        if ($blocking.Count -gt 0) { $failed = $true; Write-Host ('PSScriptAnalyzer: {0} blocking finding(s), {1} advisory' -f $blocking.Count, $advisory.Count) -ForegroundColor Red }
        else { Write-Host ('PSScriptAnalyzer: clean ({0} advisory warning(s))' -f $advisory.Count) -ForegroundColor Green }
    }
}

# --- 3. Pester ------------------------------------------------------------------------------------------------------
$pesterResult = $null
if (-not $SkipPester) {
    Write-Section 'Pester'
    $pesterLoaded = $false
    try { Import-Module Pester -RequiredVersion 5.7.1 -ErrorAction Stop; $pesterLoaded = $true }
    catch {
        Write-Host 'Pester 5.7.1 not found; trying any Pester 5.x (Install-Module Pester -RequiredVersion 5.7.1 -Scope CurrentUser -Force)' -ForegroundColor Yellow
        try { Import-Module Pester -MinimumVersion 5.0.0 -ErrorAction Stop; $pesterLoaded = $true } catch { Write-Host ('Pester 5 is not available: ' + $_.Exception.Message) -ForegroundColor Red }
    }
    if (-not $pesterLoaded) { $failed = $true }
    else {
        $pesterVersion = (Get-Module Pester).Version
        Write-Host ('Pester ' + $pesterVersion)
        $paths = @()
        if ($TestName -and $TestName.Count -gt 0) {
            foreach ($n in $TestName) {
                $stem = $n -replace '\.Tests\.ps1$', ''
                $candidate = Join-Path $testsRoot ($stem + '.Tests.ps1')
                if (Test-Path -LiteralPath $candidate) { $paths += $candidate } else { Write-Host ('  unknown test file: ' + $n) -ForegroundColor Yellow }
            }
        }
        else { $paths = @(Get-ChildItem -LiteralPath $testsRoot -Filter '*.Tests.ps1' | Sort-Object Name | ForEach-Object { $_.FullName }) }
        if ($paths.Count -eq 0) { Write-Host 'No test files selected' -ForegroundColor Red; $failed = $true }
        else {
            if (-not (Test-Path -LiteralPath $ResultsPath)) { New-Item -ItemType Directory -Path $ResultsPath -Force | Out-Null }
            $configuration = New-PesterConfiguration
            $configuration.Run.Path = $paths
            $configuration.Run.PassThru = $true
            $configuration.Run.Exit = $false
            $configuration.Output.Verbosity = 'Detailed'
            $configuration.Output.StackTraceVerbosity = 'Filtered'
            $configuration.TestResult.Enabled = $true
            $configuration.TestResult.OutputFormat = 'NUnitXml'
            $configuration.TestResult.OutputPath = Join-Path $ResultsPath 'ImpactIQ.Tests.xml'
            $configuration.Should.ErrorAction = 'Stop'
            $pesterResult = Invoke-Pester -Configuration $configuration
            if ($null -eq $pesterResult) { $failed = $true }
            else {
                Write-Host ''
                Write-Host ('Pester: {0} passed, {1} failed, {2} skipped, {3} not run in {4:N1}s' -f $pesterResult.PassedCount, $pesterResult.FailedCount, $pesterResult.SkippedCount, $pesterResult.NotRunCount, $pesterResult.Duration.TotalSeconds) -ForegroundColor $(if ($pesterResult.FailedCount -gt 0) { 'Red' } else { 'Green' })
                Write-Host ('Result file: ' + $configuration.TestResult.OutputPath.Value)
                if ($pesterResult.FailedCount -gt 0 -or $pesterResult.Result -eq 'Failed') { $failed = $true }
                if ($env:TF_BUILD -eq 'True' -and $pesterResult.FailedCount -gt 0) { Write-Host ('##vso[task.logissue type=error]Pester: {0} test(s) failed' -f $pesterResult.FailedCount) }
            }
        }
    }
}

Write-Section 'Summary'
if ($failed) { Write-Host 'RESULT: FAILED' -ForegroundColor Red } else { Write-Host 'RESULT: PASSED' -ForegroundColor Green }
if ($PassThru -and $null -ne $pesterResult) { $pesterResult }
if ($failed) { exit 1 }
exit 0
