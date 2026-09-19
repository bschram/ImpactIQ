<#
.SYNOPSIS
    Two-minute XMLA access probe: runs Tabular Editor against ONE model with every token audience / connection-string
    variant ImpactIQ knows, and prints which one the endpoint accepts.
.DESCRIPTION
    Use it when ModelBackup reports "Authentication failed for all authenticators" on a workspace that IS on a
    dedicated capacity. It needs an Az.Accounts sign-in (Connect-AzAccount is run when no context exists) and the
    Tabular Editor 2 CLI that ImpactIQ downloads to Config\TabularEditor. Nothing is written except the probe .bim files
    under -OutputFolder (default: the temp folder).

    Variants: audience = the environment's Power BI resource and the commercial one
    (https://analysis.windows.net/powerbi/api); form = "User ID=;Password=<token>" and "Password=<token>".
    A variant that produces a complete .bim is the one to pin: ImpactIQ.ps1 -XmlaTokenResource <audience> (the form is
    chosen automatically by the in-run probe, or reported here).
.EXAMPLE
    .\tools\Test-IQXmlaAccess.ps1 -Environment USGov -WorkspaceName 'DOH-PBIDashBoards' -DatasetName 'Tobacco Dashboard'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$WorkspaceName,
    [Parameter(Mandatory = $true)][string]$DatasetName,
    [Parameter(Mandatory = $false)][ValidateSet('Public', 'Commercial', 'Global', 'Germany', 'USGov', 'GCC', 'China', 'USGovHigh', 'GCCHigh', 'USGovMil', 'DoD')][string]$Environment = 'Public',
    [Parameter(Mandatory = $false)][string]$TenantId,
    [Parameter(Mandatory = $false)][string]$TabularEditorPath,
    [Parameter(Mandatory = $false)][string]$OutputFolder,
    [Parameter(Mandatory = $false)][int]$TimeoutSeconds = 120
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path (Join-Path (Join-Path $root 'Config') 'Modules') 'ImpactIQ.Common.ps1')
$ep = Get-IQEnvironmentSettings -Environment $Environment
if ([string]::IsNullOrWhiteSpace($TabularEditorPath)) { $TabularEditorPath = Join-Path (Join-Path (Join-Path $root 'Config') 'TabularEditor') 'TabularEditor.exe' }
if (-not (Test-Path -LiteralPath $TabularEditorPath)) { throw "Tabular Editor not found at $TabularEditorPath (run ImpactIQ once so it is downloaded, or pass -TabularEditorPath)." }
if ([string]::IsNullOrWhiteSpace($OutputFolder)) { $OutputFolder = Join-Path ([System.IO.Path]::GetTempPath()) 'ImpactIQ-xmla-probe' }
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

Import-Module Az.Accounts -ErrorAction Stop
$ctx = $null
try { $ctx = Get-AzContext -ErrorAction SilentlyContinue } catch { $ctx = $null }
if ($null -eq $ctx -or $null -eq $ctx.Account) {
    $azEnv = 'AzureCloud'
    if ($ep.Name -in @('USGovHigh', 'USGovMil')) { $azEnv = 'AzureUSGovernment' }
    if ($TenantId) { Connect-AzAccount -Environment $azEnv -TenantId $TenantId | Out-Null } else { Connect-AzAccount -Environment $azEnv | Out-Null }
}
function Get-ProbeToken {
    param([string]$ResourceUrl)
    $t = Get-AzAccessToken -ResourceUrl $ResourceUrl -ErrorAction Stop -WarningAction SilentlyContinue 3>$null
    $v = $t.Token
    if ($v -is [System.Security.SecureString]) { $v = [System.Net.NetworkCredential]::new('', $v).Password }
    return [string]$v
}
$audiences = @([string]$ep.PowerBIResource)
if ($ep.PowerBIResource -notlike 'https://analysis.windows.net/*') { $audiences += 'https://analysis.windows.net/powerbi/api' }
$forms = @('UserIdEmpty', 'PasswordOnly')
$dataSource = ('{0}/v1.0/myorg/{1}' -f ([string]$ep.XmlaPrefix).TrimEnd('/'), [System.Uri]::EscapeDataString($WorkspaceName))
$script = Join-Path $OutputFolder 'noop.cs'
[System.IO.File]::WriteAllText($script, '// no-op' + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($true)))
Write-Host ("Probing {0} / '{1}' on {2} ({3} audience(s) x {4} form(s))" -f $WorkspaceName, $DatasetName, $dataSource, $audiences.Count, $forms.Count)
$rows = @()
$n = 0
foreach ($audience in $audiences) {
    $token = $null
    try { $token = Get-ProbeToken -ResourceUrl $audience } catch { $rows += [pscustomobject]@{ Audience = $audience; Form = '-'; Result = 'no token: ' + $_.Exception.Message }; continue }
    foreach ($form in $forms) {
        $n++
        $bim = Join-Path $OutputFolder ('probe-' + $n + '.bim')
        if (Test-Path -LiteralPath $bim) { Remove-Item -LiteralPath $bim -Force }
        $credential = 'User ID=;Password=' + $token
        if ($form -eq 'PasswordOnly') { $credential = 'Password=' + $token }
        $arguments = ('"Provider=MSOLAP;Data Source={0};{1}" "{2}" -S "{3}" -B "{4}"' -f $dataSource, $credential, $DatasetName, $script, $bim)
        $stdout = Join-Path $OutputFolder ('probe-' + $n + '.out.txt')
        $stderr = Join-Path $OutputFolder ('probe-' + $n + '.err.txt')
        $p = Start-Process -FilePath $TabularEditorPath -ArgumentList $arguments -WorkingDirectory $root -NoNewWindow -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr
        if (-not $p.WaitForExit($TimeoutSeconds * 1000)) { try { $p.Kill() } catch { $null = $_ }; $rows += [pscustomobject]@{ Audience = $audience; Form = $form; Result = 'timeout' }; continue }
        $text = ''
        try { $text = [System.IO.File]::ReadAllText($stdout) + [System.IO.File]::ReadAllText($stderr) } catch { $text = '' }
        $ok = (Test-Path -LiteralPath $bim) -and ((Get-Item -LiteralPath $bim).Length -gt 100)
        $firstError = (@($text -split "`r?`n" | Where-Object { $_ -match '^\s*Error' }) | Select-Object -First 1)
        $rows += [pscustomobject]@{ Audience = $audience; Form = $form; Result = $(if ($ok) { 'ACCEPTED (.bim written: ' + $bim + ')' } else { 'refused: exit ' + $p.ExitCode + ' ' + $firstError }) }
    }
}
$rows | Format-Table -AutoSize -Wrap
$accepted = @($rows | Where-Object { $_.Result -like 'ACCEPTED*' })
if ($accepted.Count -gt 0) {
    Write-Host ''
    Write-Host ("Pin the audience for ImpactIQ runs: ImpactIQ.ps1 ... -XmlaTokenResource '{0}'   (or XmlaTokenResource in Config\ImpactIQ.Settings.json)" -f $accepted[0].Audience) -ForegroundColor Green
}
else {
    Write-Host ''
    Write-Host 'No variant was accepted. Remaining causes ImpactIQ cannot see: the capacity XMLA Endpoint setting (Admin portal > Capacity settings > Power BI workloads > XMLA Endpoint: Read or Read Write), the tenant setting "Allow XMLA endpoints and Analyze in Excel with on-premises semantic models", and Build permission on the model.' -ForegroundColor Yellow
}
