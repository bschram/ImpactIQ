<#
.SYNOPSIS
    Reproduces the recovered build's XMLA export path (MicrosoftPowerBIMgmt token + "Password=<token>" connection string)
    against ONE model, and compares it with the Az.Accounts token ImpactIQ v3 uses, so the difference can be pinned.
.DESCRIPTION
    The recovered build ("Final PS Script.txt", branch bschram_updates) exported models with:
        Connect-PowerBIServiceAccount -Environment USGov
        $token = (Get-PowerBIAccessToken).Authorization -replace 'Bearer ', ''
        TabularEditor.exe "Provider=MSOLAP;Data Source=powerbi://api.powerbigov.us/v1.0/myorg/<ws>;Password=<token>" "<dataset>" -S <rename.cs> -B <out.bim>
    where <ws> only had "[", "]" and spaces percent-encoded.

    This script runs exactly that (step 1), then - unless -SkipAz - mints the same audience with Get-AzAccessToken
    (the v3 token source) and runs the identical Tabular Editor command with it (step 2). Both tokens are decoded
    (aud / appid / tid / scp / ver / upn) and printed side by side, tokens themselves are never printed. Each variant
    also runs in the v3 "User ID=;Password=<token>" form so the connection-string form is ruled in or out.

    Requirements: Windows PowerShell 5.1 or PowerShell 7, MicrosoftPowerBIMgmt (installed for the current user when
    missing), Tabular Editor 2 at Config\TabularEditor\TabularEditor.exe (extracted from Config\TabularEditor.zip when
    only the zip is present), Az.Accounts for step 2 (skipped with a note when it is not installed).
.EXAMPLE
    .\tools\Test-IQXmlaAccess-Legacy.ps1 -Environment USGov -WorkspaceName 'DOH-PBIDashBoards' -DatasetName 'Tobacco Dashboard'
.EXAMPLE
    .\tools\Test-IQXmlaAccess-Legacy.ps1 -WorkspaceName 'DOH-PBIDashBoards' -DatasetName 'Tobacco Dashboard' -SkipAz -ForceLogin
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)][string]$WorkspaceName = 'DOH-PBIDashBoards',
    [Parameter(Mandatory = $false)][string]$DatasetName = 'Tobacco Dashboard',
    [Parameter(Mandatory = $false)][ValidateSet('Public', 'USGov', 'USGovHigh', 'USGovMil', 'Germany', 'China')][string]$Environment = 'USGov',
    [Parameter(Mandatory = $false)][string]$TenantId,
    [Parameter(Mandatory = $false)][string]$TabularEditorPath,
    [Parameter(Mandatory = $false)][string]$OutputFolder,
    [Parameter(Mandatory = $false)][int]$TimeoutSeconds = 180,
    [Parameter(Mandatory = $false)][switch]$SkipAz,
    [Parameter(Mandatory = $false)][switch]$ForceLogin
)
$ErrorActionPreference = 'Stop'
$waitMilliseconds = $TimeoutSeconds * 1000

# ---------------------------------------------------------------- locations
$root = $PSScriptRoot
if (-not (Test-Path -LiteralPath (Join-Path $root 'Config'))) { $root = Split-Path -Parent $root }
$configFolder = Join-Path $root 'Config'
if ([string]::IsNullOrWhiteSpace($TabularEditorPath)) { $TabularEditorPath = Join-Path (Join-Path $configFolder 'TabularEditor') 'TabularEditor.exe' }
if (-not (Test-Path -LiteralPath $TabularEditorPath)) {
    $zip = Join-Path $configFolder 'TabularEditor.zip'
    if (Test-Path -LiteralPath $zip) {
        Write-Host "[INFO] Extracting $zip"
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $teFolder = Split-Path -Parent $TabularEditorPath
        New-Item -ItemType Directory -Path $teFolder -Force | Out-Null
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zip, $teFolder)
    }
}
if (-not (Test-Path -LiteralPath $TabularEditorPath)) { throw "Tabular Editor not found at $TabularEditorPath (pass -TabularEditorPath)." }
if ([string]::IsNullOrWhiteSpace($OutputFolder)) { $OutputFolder = Join-Path ([System.IO.Path]::GetTempPath()) 'ImpactIQ-xmla-legacy-probe' }
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
$teVersion = ''
try { $teVersion = [string](Get-Item -LiteralPath $TabularEditorPath).VersionInfo.FileVersion } catch { $teVersion = '?' }

# ---------------------------------------------------------------- endpoints (copied from the recovered build's Get-PowerBIEndpoints)
$xmlaPrefix = 'powerbi://api.powerbi.com'
$resourceUrl = 'https://analysis.windows.net/powerbi/api'
$azEnvironment = 'AzureCloud'
switch ($Environment) {
    'Germany' { $xmlaPrefix = 'powerbi://api.powerbi.de'; $resourceUrl = 'https://analysis.cloudapi.de/powerbi/api'; $azEnvironment = 'AzureGermanCloud' }
    'China' { $xmlaPrefix = 'powerbi://api.powerbi.cn'; $resourceUrl = 'https://analysis.chinacloudapi.cn/powerbi/api'; $azEnvironment = 'AzureChinaCloud' }
    'USGov' { $xmlaPrefix = 'powerbi://api.powerbigov.us'; $resourceUrl = 'https://analysis.usgovcloudapi.net/powerbi/api'; $azEnvironment = 'AzureCloud' }
    'USGovHigh' { $xmlaPrefix = 'powerbi://api.high.powerbigov.us'; $resourceUrl = 'https://analysis.high.usgovcloudapi.net/powerbi/api'; $azEnvironment = 'AzureUSGovernment' }
    'USGovMil' { $xmlaPrefix = 'powerbi://api.mil.powerbi.us'; $resourceUrl = 'https://analysis.dod.usgovcloudapi.net/powerbi/api'; $azEnvironment = 'AzureUSGovernment' }
}
# Legacy encoding: only [ ] and space (the recovered build); v3 uses full URL escaping.
$legacyWorkspace = $WorkspaceName -replace '\[', '%5B' -replace '\]', '%5D' -replace ' ', '%20'
$dataSource = '{0}/v1.0/myorg/{1}' -f $xmlaPrefix, $legacyWorkspace

# ---------------------------------------------------------------- helpers
function ConvertFrom-JwtPayload {
    param([string]$Token)
    if ([string]::IsNullOrWhiteSpace($Token)) { return $null }
    $parts = $Token.Split('.')
    if ($parts.Count -lt 2) { return $null }
    $p = $parts[1].Replace('-', '+').Replace('_', '/')
    switch ($p.Length % 4) { 2 { $p += '==' } 3 { $p += '=' } }
    try { return ([System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json) } catch { return $null }
}
function Get-ClaimSummary {
    param([string]$Source, [string]$Token)
    $c = ConvertFrom-JwtPayload -Token $Token
    if ($null -eq $c) { return [pscustomobject]@{ Source = $Source; Note = 'not a JWT (length ' + [string]$Token.Length + ')' } }
    $exp = ''
    try { $exp = ([datetime]'1970-01-01Z').AddSeconds([double]$c.exp).ToLocalTime().ToString('HH:mm:ss') } catch { $exp = '' }
    $user = ''
    foreach ($k in 'upn', 'preferred_username', 'unique_name', 'email') { if ($c.PSObject.Properties[$k] -and $c.$k) { $user = [string]$c.$k; break } }
    $app = ''
    foreach ($k in 'appid', 'azp') { if ($c.PSObject.Properties[$k] -and $c.$k) { $app = [string]$c.$k; break } }
    $scp = ''
    if ($c.PSObject.Properties['scp']) { $scp = [string]$c.scp }
    $ver = ''
    if ($c.PSObject.Properties['ver']) { $ver = [string]$c.ver }
    $tid = ''
    if ($c.PSObject.Properties['tid']) { $tid = [string]$c.tid }
    $aud = ''
    if ($c.PSObject.Properties['aud']) { $aud = [string]$c.aud }
    $iss = ''
    if ($c.PSObject.Properties['iss']) { $iss = [string]$c.iss }
    return [pscustomobject]@{ Source = $Source; aud = $aud; appid = $app; tid = $tid; ver = $ver; scp = $scp; user = $user; iss = $iss; expires = $exp }
}
function Get-AppName {
    param([string]$AppId)
    switch ($AppId.ToLowerInvariant()) {
        '1950a258-227b-4e31-a9cf-717495945fc2' { return 'Microsoft Azure PowerShell' }
        '04b07795-8ddb-461a-bbee-02f9e1bf7b46' { return 'Microsoft Azure CLI' }
        '23d8f6bd-1eb0-4cc2-a08c-7bf525c67bcd' { return 'Power BI PowerShell (MicrosoftPowerBIMgmt)' }
        '7f67af8a-fedc-4b08-8b4e-37c4d127b6cf' { return 'Power BI Desktop' }
        'ea0616ba-638b-4df5-95b9-636659ae5121' { return 'Power BI Premium / XMLA' }
        'cf710c6e-dfcc-4fa8-a093-d47294e44c66' { return 'Power BI (Analysis Services client)' }
        default { return '' }
    }
}
$probeIndex = 0
function Invoke-TeProbe {
    param([string]$Label, [string]$Token, [ValidateSet('PasswordOnly', 'UserIdEmpty')][string]$Form)
    $script:probeIndex++
    $n = $script:probeIndex
    $bim = Join-Path $OutputFolder ('legacy-probe-' + $n + '.bim')
    if (Test-Path -LiteralPath $bim) { Remove-Item -LiteralPath $bim -Force }
    # Same rename script the recovered build used (Model.Database.Name = "<ws> ~ <dataset>").
    $cs = Join-Path $OutputFolder ('legacy-probe-' + $n + '.cs')
    $newName = ($WorkspaceName + ' ~ ' + $DatasetName) -replace '"', ''
    [System.IO.File]::WriteAllText($cs, ('Model.Database.Name = "' + $newName + '";' + [Environment]::NewLine), (New-Object System.Text.UTF8Encoding($true)))
    $credential = 'Password=' + $Token
    if ($Form -eq 'UserIdEmpty') { $credential = 'User ID=;Password=' + $Token }
    $arguments = ('"Provider=MSOLAP;Data Source={0};{1}" "{2}" -S "{3}" -B "{4}"' -f $dataSource, $credential, $DatasetName, $cs, $bim)
    $stdout = Join-Path $OutputFolder ('legacy-probe-' + $n + '.out.txt')
    $stderr = Join-Path $OutputFolder ('legacy-probe-' + $n + '.err.txt')
    Write-Host ("  [{0}] {1} / {2} ..." -f $n, $Label, $Form) -NoNewline
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $p = Start-Process -FilePath $TabularEditorPath -ArgumentList $arguments -WorkingDirectory $root -NoNewWindow -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
    $timedOut = -not $p.WaitForExit($waitMilliseconds)
    if ($timedOut) { try { $p.Kill() } catch { $null = $_ } }
    $text = ''
    try { $text = [System.IO.File]::ReadAllText($stdout) + [System.IO.File]::ReadAllText($stderr) } catch { $text = '' }
    $ok = (Test-Path -LiteralPath $bim) -and ((Get-Item -LiteralPath $bim).Length -gt 100)
    $lines = @($text -split "`r?`n" | Where-Object { $_.Trim() -ne '' })
    $firstError = @($lines | Where-Object { $_ -match '^\s*Error|exception|failed' }) | Select-Object -First 1
    if (-not $firstError) { $firstError = @($lines | Select-Object -Last 1) | Select-Object -First 1 }
    $result = 'refused'
    if ($timedOut) { $result = 'timeout' } elseif ($ok) { $result = 'ACCEPTED' }
    Write-Host (' ' + $result + ' (' + [int]$sw.Elapsed.TotalSeconds + ' s)')
    return [pscustomobject]@{
        '#'      = $n
        Token    = $Label
        Form     = $Form
        Result   = $result
        ExitCode = $(if ($timedOut) { '' } else { $p.ExitCode })
        Detail   = $(if ($ok) { '.bim written: ' + $bim } else { [string]$firstError })
        Output   = $stdout
    }
}

Write-Host ''
Write-Host ('Legacy XMLA probe: {0} / ''{1}''' -f $WorkspaceName, $DatasetName)
Write-Host ('  Data Source     : {0}' -f $dataSource)
Write-Host ('  Token audience  : {0}' -f $resourceUrl)
Write-Host ('  Tabular Editor  : {0} (v{1})' -f $TabularEditorPath, $teVersion)
Write-Host ('  Output folder   : {0}' -f $OutputFolder)
Write-Host ''

# ---------------------------------------------------------------- step 1: MicrosoftPowerBIMgmt token (the recovered build's source)
Write-Host '[1] MicrosoftPowerBIMgmt token (Connect-PowerBIServiceAccount / Get-PowerBIAccessToken) - the recovered build''s path'
foreach ($m in 'MicrosoftPowerBIMgmt.Profile') {
    if (-not (Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue)) {
        Write-Host "  [INFO] Installing MicrosoftPowerBIMgmt for the current user (the recovered build did the same)."
        Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
}
$moduleToken = $null
if (-not $ForceLogin) {
    try { $moduleToken = ((Get-PowerBIAccessToken -AsString -ErrorAction Stop) -replace '^(?i)Bearer\s+', '').Trim() } catch { $moduleToken = $null }
}
if ([string]::IsNullOrWhiteSpace($moduleToken)) {
    $connectArgs = @{}
    if ($Environment -ne 'Public') { $connectArgs.Environment = $Environment }
    if ($TenantId) { $connectArgs.Tenant = $TenantId }
    Write-Host ("  Signing in: Connect-PowerBIServiceAccount {0}" -f (($connectArgs.GetEnumerator() | ForEach-Object { '-' + $_.Key + ' ' + $_.Value }) -join ' '))
    Connect-PowerBIServiceAccount @connectArgs | Out-Null
    $moduleToken = ((Get-PowerBIAccessToken -AsString -ErrorAction Stop) -replace '^(?i)Bearer\s+', '').Trim()
}
if ([string]::IsNullOrWhiteSpace($moduleToken)) { throw 'Get-PowerBIAccessToken returned nothing.' }
$claims = @()
$claims += Get-ClaimSummary -Source 'MicrosoftPowerBIMgmt' -Token $moduleToken
$rows = @()
$rows += Invoke-TeProbe -Label 'MicrosoftPowerBIMgmt' -Token $moduleToken -Form PasswordOnly
$rows += Invoke-TeProbe -Label 'MicrosoftPowerBIMgmt' -Token $moduleToken -Form UserIdEmpty

# ---------------------------------------------------------------- step 2: Az.Accounts token (ImpactIQ v3's source)
if ($SkipAz) {
    Write-Host '[2] Az.Accounts comparison skipped (-SkipAz).'
}
elseif (-not (Get-Module -ListAvailable -Name Az.Accounts -ErrorAction SilentlyContinue)) {
    Write-Host '[2] Az.Accounts is not installed on this machine - comparison skipped (Install-Module Az.Accounts -Scope CurrentUser to enable it).'
}
else {
    Write-Host '[2] Az.Accounts token (Get-AzAccessToken -ResourceUrl) - ImpactIQ v3''s path'
    Import-Module Az.Accounts -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
    $ctx = $null
    try { $ctx = Get-AzContext -ErrorAction SilentlyContinue } catch { $ctx = $null }
    if ($ForceLogin -or $null -eq $ctx -or $null -eq $ctx.Account) {
        Write-Host ("  Signing in: Connect-AzAccount -Environment {0}{1}" -f $azEnvironment, $(if ($TenantId) { ' -TenantId ' + $TenantId } else { '' }))
        if ($TenantId) { Connect-AzAccount -Environment $azEnvironment -TenantId $TenantId -WarningAction SilentlyContinue 3>$null | Out-Null }
        else { Connect-AzAccount -Environment $azEnvironment -WarningAction SilentlyContinue 3>$null | Out-Null }
        $ctx = Get-AzContext
    }
    Write-Host ("  Az context: account={0} tenant={1} environment={2}" -f $ctx.Account.Id, $ctx.Tenant.Id, $ctx.Environment.Name)
    $azToken = $null
    try {
        $t = Get-AzAccessToken -ResourceUrl $resourceUrl -ErrorAction Stop -WarningAction SilentlyContinue 3>$null
        $v = $t.Token
        if ($v -is [System.Security.SecureString]) { $v = [System.Net.NetworkCredential]::new('', $v).Password }
        $azToken = [string]$v
    }
    catch {
        Write-Host ('  [WARN] Get-AzAccessToken failed: ' + $_.Exception.Message)
    }
    if ($azToken) {
        $claims += Get-ClaimSummary -Source 'Az.Accounts' -Token $azToken
        $rows += Invoke-TeProbe -Label 'Az.Accounts' -Token $azToken -Form PasswordOnly
        $rows += Invoke-TeProbe -Label 'Az.Accounts' -Token $azToken -Form UserIdEmpty
    }
}

# ---------------------------------------------------------------- report
Write-Host ''
Write-Host '=== Token claims (tokens themselves are not shown) ==='
foreach ($c in $claims) {
    Write-Host ('  ' + $c.Source)
    if ($c.PSObject.Properties['Note']) { Write-Host ('    ' + $c.Note); continue }
    $appName = Get-AppName -AppId $c.appid
    Write-Host ('    aud    : ' + $c.aud)
    Write-Host ('    appid  : ' + $c.appid + $(if ($appName) { '  (' + $appName + ')' } else { '' }))
    Write-Host ('    tid    : ' + $c.tid)
    Write-Host ('    ver    : ' + $c.ver)
    Write-Host ('    scp    : ' + $c.scp)
    Write-Host ('    user   : ' + $c.user)
    Write-Host ('    iss    : ' + $c.iss)
    Write-Host ('    expires: ' + $c.expires)
}
if ($claims.Count -eq 2 -and -not $claims[0].PSObject.Properties['Note'] -and -not $claims[1].PSObject.Properties['Note']) {
    $diff = @()
    foreach ($k in 'aud', 'appid', 'tid', 'ver', 'scp', 'user', 'iss') { if ([string]$claims[0].$k -ne [string]$claims[1].$k) { $diff += $k } }
    Write-Host ('  Claims that differ between the two tokens: ' + $(if ($diff.Count) { $diff -join ', ' } else { 'none' }))
}
Write-Host ''
Write-Host '=== Tabular Editor results ==='
$rows | Select-Object '#', Token, Form, Result, ExitCode, Detail | Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Host
Write-Host ('Full Tabular Editor output per attempt: ' + $OutputFolder + '\legacy-probe-<#>.out.txt')
Write-Host ''
$moduleOk = @($rows | Where-Object { $_.Token -eq 'MicrosoftPowerBIMgmt' -and $_.Result -eq 'ACCEPTED' }).Count -gt 0
$azOk = @($rows | Where-Object { $_.Token -eq 'Az.Accounts' -and $_.Result -eq 'ACCEPTED' }).Count -gt 0
$azTried = @($rows | Where-Object { $_.Token -eq 'Az.Accounts' }).Count -gt 0
if ($moduleOk -and $azTried -and -not $azOk) {
    Write-Host 'VERDICT: the XMLA endpoint accepts the MicrosoftPowerBIMgmt token and refuses the Az.Accounts token for the same audience.'
    Write-Host '         ImpactIQ v3 must take its XMLA token from the Power BI module (the appid/scp lines above show what the endpoint wants).'
}
elseif ($moduleOk -and $azOk) {
    Write-Host 'VERDICT: both tokens are accepted here - the failing runs differ in something else (workspace name encoding, dataset name, or the capacity''s XMLA setting at the time).'
}
elseif (-not $moduleOk -and $azTried -and $azOk) {
    Write-Host 'VERDICT: only the Az.Accounts token is accepted - the Power BI module path is not the answer.'
}
elseif (-not $moduleOk) {
    Write-Host 'VERDICT: the recovered build''s exact path is refused too. The endpoint (not the token source) changed since it last worked:'
    Write-Host '         check Admin portal > Capacity settings > XMLA Endpoint (Read or Read Write) for this capacity, the tenant setting'
    Write-Host '         "Allow XMLA endpoints and Analyze in Excel with on-premises datasets", and that the workspace is still on that capacity.'
}
Write-Host 'Please paste this whole output (it contains no tokens) into the chat.'
