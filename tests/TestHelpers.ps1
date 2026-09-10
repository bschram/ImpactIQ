# TestHelpers.ps1 - shared setup for the ImpactIQ Pester 5 tests.
#
# Dot-source this file inside a test file's BeforeAll block:
#     BeforeAll { . (Join-Path $PSScriptRoot 'TestHelpers.ps1') }
# It dot-sources every ImpactIQ module (brief section 1 load order, Interactive excluded because it is WinForms-only)
# into the test file's script scope, so `$script:IQ` inside the modules is the test file's `$script:IQ`, and
# `Mock Invoke-WebRequest` / `Mock Invoke-RestMethod` / `Mock Start-IQHttpSleep` in the test file intercept the calls
# the modules make. Each test file therefore runs in a fresh context; Initialize-IQTestContext creates a temporary
# BaseFolder (with a Config folder holding the csx placeholder files and a copy of Config/SheetContract.json).
#
# Windows PowerShell 5.1 and PowerShell 7 compatible (no ternary, no ??, no ?., no $IsWindows).

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'

$script:IQTestsRoot = $PSScriptRoot
$script:IQTestRepoRoot = Split-Path -Path $PSScriptRoot -Parent
$script:IQTestFixturesRoot = Join-Path $PSScriptRoot 'fixtures'
$script:IQTestModulesRoot = Join-Path $script:IQTestRepoRoot 'Config/Modules'
$script:IQTestModuleOrder = @('Common', 'Auth', 'Http', 'State', 'Tools', 'Inventory', 'Dax', 'Bim', 'Models', 'Reports', 'Dataflows', 'Extras', 'Assemble')
$script:IQTestCsxNames = @('Model Detail Extract Script.csx', 'Measure Dependency Extract Script.csx', 'Report Detail Extract Script.csx', 'Report Detail Extract Script-PBIR.csx')
$script:IQTestApiCalls = New-Object System.Collections.Generic.List[string]
$script:IQTestDaxOverrides = @{}
$script:IQTestApiOverrides = @{}
$script:IQTestModulesLoaded = New-Object System.Collections.Generic.List[string]

# --- load the modules into this (the test file's) script scope -------------------------------------------------------
$script:IQ = $null
foreach ($iqTestModuleName in $script:IQTestModuleOrder) {
    $iqTestModulePath = Join-Path $script:IQTestModulesRoot ('ImpactIQ.' + $iqTestModuleName + '.ps1')
    if (Test-Path -LiteralPath $iqTestModulePath) {
        . $iqTestModulePath
        $script:IQTestModulesLoaded.Add($iqTestModuleName)
    }
}

function Get-IQTestRepoRoot {
    <#
    .SYNOPSIS
    Returns the repository root (the folder that contains ImpactIQ.ps1 and Config\).
    #>
    [CmdletBinding()]
    param()
    return $script:IQTestRepoRoot
}

function Test-IQTestModuleLoaded {
    <#
    .SYNOPSIS
    True when the named ImpactIQ module (Common, Auth, Http, ...) was found and dot-sourced.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Name)
    return ($script:IQTestModulesLoaded -contains $Name)
}

function Reset-IQTestEnvironment {
    <#
    .SYNOPSIS
    Clears every IMPACTIQ_* environment variable (and TF_BUILD/CI) so auth-mode resolution starts from a clean slate.
    #>
    [CmdletBinding()]
    param()
    foreach ($name in @('IMPACTIQ_USERNAME', 'IMPACTIQ_PASSWORD', 'IMPACTIQ_PBI_TOKEN', 'IMPACTIQ_FABRIC_TOKEN', 'IMPACTIQ_TOKEN_CACHE_KEY',
            'IMPACTIQ_TOKEN_CACHE_PATH', 'IMPACTIQ_DEVICECODE_WEBHOOK', 'IMPACTIQ_TENANT_ID', 'IMPACTIQ_CLIENT_ID', 'IMPACTIQ_ENVIRONMENT',
            'IMPACTIQ_OFFLINE', 'IMPACTIQ_DEBUG')) {
        [System.Environment]::SetEnvironmentVariable($name, $null, 'Process')
    }
}

function New-IQTestBaseFolder {
    <#
    .SYNOPSIS
    Creates a temporary BaseFolder with a Config folder (csx placeholders, Blank Model.bim placeholder, SheetContract.json copy).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Prefix = 'iqtest')
    $root = Join-Path ([System.IO.Path]::GetTempPath()) 'ImpactIQ-tests'
    $base = Join-Path $root ($Prefix + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $config = Join-Path $base 'Config'
    New-Item -ItemType Directory -Path $config -Force | Out-Null
    foreach ($csx in $script:IQTestCsxNames) {
        [System.IO.File]::WriteAllText((Join-Path $config $csx), '// placeholder for tests' + [Environment]::NewLine)
    }
    $blankBim = Join-Path $script:IQTestRepoRoot 'Config/Blank Model.bim'
    if (Test-Path -LiteralPath $blankBim) { Copy-Item -LiteralPath $blankBim -Destination (Join-Path $config 'Blank Model.bim') -Force }
    else { [System.IO.File]::WriteAllText((Join-Path $config 'Blank Model.bim'), '{"name":"Blank","model":{}}') }
    $contract = Join-Path $script:IQTestRepoRoot 'Config/SheetContract.json'
    if (Test-Path -LiteralPath $contract) { Copy-Item -LiteralPath $contract -Destination (Join-Path $config 'SheetContract.json') -Force }
    return $base
}

function Initialize-IQTestContext {
    <#
    .SYNOPSIS
    Creates a temp BaseFolder, initialises $script:IQ (headless, Public unless overridden) and, unless -NoRun, a fresh run.
    .DESCRIPTION
    Returns the BaseFolder path. -Options are merged over the defaults (Environment = Public, NonInteractive = $true).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][hashtable]$Options,
        [Parameter(Mandatory = $false)][string]$RunId = '2026-09-04',
        [Parameter(Mandatory = $false)][switch]$NoRun,
        [Parameter(Mandatory = $false)][string]$Prefix = 'iqtest'
    )
    $base = New-IQTestBaseFolder -Prefix $Prefix
    $effective = @{ Environment = 'Public'; NonInteractive = $true }
    if ($Options) { foreach ($k in $Options.Keys) { $effective[$k] = $Options[$k] } }
    Initialize-IQContext -BaseFolder $base -Options $effective | Out-Null
    $script:IQ.Interactive = $false
    $script:IQTestApiCalls.Clear()
    if (-not $NoRun) { Initialize-IQRun -RunId $RunId -ResumePolicy Never | Out-Null }
    return $base
}

function Remove-IQTestFolder {
    <#
    .SYNOPSIS
    Deletes a temporary test folder, ignoring errors.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue }
}

function Get-IQTestFixturePath {
    <#
    .SYNOPSIS
    Absolute path of a file under tests/fixtures (relative path with forward slashes).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Relative)
    return (Join-Path $script:IQTestFixturesRoot ($Relative -replace '/', [System.IO.Path]::DirectorySeparatorChar))
}

function Get-IQTestFixtureText {
    <#
    .SYNOPSIS
    Reads a fixture file as UTF-8 text.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Relative)
    return [System.IO.File]::ReadAllText((Get-IQTestFixturePath -Relative $Relative), [System.Text.Encoding]::UTF8)
}

function Get-IQTestFixtureJson {
    <#
    .SYNOPSIS
    Reads a fixture file and parses it as JSON (ConvertFrom-Json).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Relative)
    return (ConvertFrom-Json -InputObject (Get-IQTestFixtureText -Relative $Relative))
}

function ConvertTo-IQTestFixtureName {
    <#
    .SYNOPSIS
    Maps an API call (Api + relative path) to its fixture file name: <Api>__<path with / replaced by __>.json.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Api,
        [Parameter(Mandatory = $true)][string]$Path
    )
    $p = $Path.Trim().TrimStart('/')
    $q = $p.IndexOf('?')
    if ($q -ge 0) { $p = $p.Substring(0, $q) }
    return ($Api + '__' + ($p -replace '/', '__') + '.json')
}

function Get-IQTestDaxFixtureKey {
    <#
    .SYNOPSIS
    Normalises a DAX query text to its fixture key: 'EVALUATE INFO.VIEW.TABLES()' -> 'info-view-tables'.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Dax)
    $t = ($Dax -replace '(?i)^\s*EVALUATE\s+', '').Trim()
    $t = $t -replace '\(\)\s*$', ''
    $t = $t -replace '[^A-Za-z0-9]+', '-'
    return $t.Trim('-').ToLowerInvariant()
}

function Invoke-IQTestApiFixture {
    <#
    .SYNOPSIS
    Fixture-backed stand-in for Invoke-IQApi: answers from tests/fixtures/api/*.json and tests/fixtures/dax/*.json.
    .DESCRIPTION
    Use it as the body of a Pester mock:
        Mock Invoke-IQApi { Invoke-IQTestApiFixture -Method $Method -Path $Path -Body $Body -Query $Query -Api $Api -Raw:$Raw }
    Rules: a fixture file named <Api>__<path>.json is returned parsed (or as the raw string with -Raw); a fixture whose
    JSON is { "__status": 403 } (or 404/400) yields $null exactly like Invoke-IQApi does for those statuses; a
    list endpoint without a fixture yields { value: [] }; a single-object endpoint (refreshSchedule,
    directQueryRefreshSchedule) without a fixture yields $null; 'groups' with $skip > 0 yields an empty page.
    POST .../executeQueries answers from tests/fixtures/dax/<key>.json keyed by the DAX text (Get-IQTestDaxFixtureKey);
    a missing DAX fixture yields $null (what Invoke-IQApi returns for an HTTP 400 "not supported" answer).
    $script:IQTestApiOverrides[<fixture name>] and $script:IQTestDaxOverrides[<dax key>] (scriptblock or value) win
    over the files. Every call is appended to $script:IQTestApiCalls as "<Api> <METHOD> <path>".
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$Method = 'GET',
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][object]$Body,
        [Parameter(Mandatory = $false)][hashtable]$Query,
        [Parameter(Mandatory = $false)][string]$Api = 'PowerBI',
        [Parameter(Mandatory = $false)][switch]$Raw
    )
    if ([string]::IsNullOrEmpty($Api)) { $Api = 'PowerBI' }
    if ([string]::IsNullOrEmpty($Method)) { $Method = 'GET' }
    $relative = $Path.Trim().TrimStart('/')
    $qi = $relative.IndexOf('?')
    if ($qi -ge 0) { $relative = $relative.Substring(0, $qi) }
    $script:IQTestApiCalls.Add(('{0} {1} {2}' -f $Api, $Method.ToUpperInvariant(), $relative))

    # DAX executeQueries
    if ($relative -match '(?i)/executeQueries$') {
        $dax = ''
        try {
            $queries = $null
            if ($Body -is [System.Collections.IDictionary]) { $queries = $Body['queries'] } elseif ($null -ne $Body) { $queries = $Body.queries }
            $first = @($queries)[0]
            if ($first -is [System.Collections.IDictionary]) { $dax = [string]$first['query'] } elseif ($null -ne $first) { $dax = [string]$first.query }
        }
        catch { $dax = '' }
        $key = Get-IQTestDaxFixtureKey -Dax $dax
        if ($script:IQTestDaxOverrides.ContainsKey($key)) {
            $o = $script:IQTestDaxOverrides[$key]
            if ($o -is [scriptblock]) { return (& $o) }
            return $o
        }
        $daxFile = Join-Path (Join-Path $script:IQTestFixturesRoot 'dax') ($key + '.json')
        if (-not (Test-Path -LiteralPath $daxFile)) { return $null }
        $text = [System.IO.File]::ReadAllText($daxFile, [System.Text.Encoding]::UTF8)
        if ($Raw) { return $text }
        return (ConvertFrom-Json -InputObject $text)
    }

    $name = ConvertTo-IQTestFixtureName -Api $Api -Path $relative
    if ($script:IQTestApiOverrides.ContainsKey($name)) {
        $o = $script:IQTestApiOverrides[$name]
        if ($o -is [scriptblock]) { return (& $o) }
        return $o
    }
    if ($relative -ieq 'groups' -and $Query -and $Query.ContainsKey('$skip') -and [int]$Query['$skip'] -gt 0) {
        return ([pscustomobject]@{ value = @() })
    }
    $file = Join-Path (Join-Path $script:IQTestFixturesRoot 'api') $name
    if (Test-Path -LiteralPath $file) {
        $text = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)
        $parsed = ConvertFrom-Json -InputObject $text
        if ($parsed -is [System.Management.Automation.PSCustomObject] -and $parsed.PSObject.Properties['__status']) { return $null }
        if ($Raw) { return $text }
        return $parsed
    }
    if ($relative -match '(?i)/(refreshSchedule|directQueryRefreshSchedule)$') { return $null }
    if ($Raw) { return $null }
    return ([pscustomobject]@{ value = @() })
}

function New-IQTestJwt {
    <#
    .SYNOPSIS
    Builds an unsigned JWT-shaped token (base64url header.payload.sig) with the given claims; exp defaults to +60 min.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][hashtable]$Claims,
        [Parameter(Mandatory = $false)][int]$ExpiresInMinutes = 60
    )
    $c = @{}
    if ($Claims) { foreach ($k in $Claims.Keys) { $c[$k] = $Claims[$k] } }
    if (-not $c.ContainsKey('exp')) {
        $epoch = New-Object DateTime 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
        $c['exp'] = [int64](([datetime]::UtcNow.AddMinutes($ExpiresInMinutes) - $epoch).TotalSeconds)
    }
    $header = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('{"alg":"RS256","typ":"JWT"}')).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $json = ConvertTo-Json -InputObject $c -Compress -Depth 20
    $payload = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($json)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    return ($header + '.' + $payload + '.testsig')
}

function ConvertTo-IQTestEpoch {
    <#
    .SYNOPSIS
    Unix epoch seconds for a [datetime] (UTC).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][datetime]$Date)
    $epoch = New-Object DateTime 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
    return [int64](($Date.ToUniversalTime() - $epoch).TotalSeconds)
}

function New-IQTestHttpResponse {
    <#
    .SYNOPSIS
    Builds an Invoke-WebRequest-like response object (StatusCode, Headers, Content) for Mock Invoke-WebRequest.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][int]$StatusCode = 200,
        [Parameter(Mandatory = $false)][AllowNull()][object]$Content,
        [Parameter(Mandatory = $false)][hashtable]$Headers
    )
    $h = @{ 'Content-Type' = 'application/json; charset=utf-8' }
    if ($Headers) { foreach ($k in $Headers.Keys) { $h[$k] = $Headers[$k] } }
    $body = $Content
    if ($null -ne $Content -and -not ($Content -is [string]) -and -not ($Content -is [byte[]])) {
        $body = ConvertTo-Json -InputObject $Content -Depth 20 -Compress
    }
    return ([pscustomobject]@{ StatusCode = $StatusCode; Headers = $h; Content = $body })
}

function New-IQTestHttpErrorRecord {
    <#
    .SYNOPSIS
    Builds an ErrorRecord that Get-IQHttpErrorInfo reads as an HTTP error with the given status, headers and body.
    .DESCRIPTION
    Uses Microsoft.PowerShell.Commands.HttpResponseException when the type exists (PowerShell 7); otherwise a plain
    exception carrying a Response note property with StatusCode/Headers (works on Windows PowerShell 5.1 as well).
    The body travels in ErrorRecord.ErrorDetails.Message exactly like Invoke-WebRequest reports it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][int]$StatusCode,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Body = '',
        [Parameter(Mandatory = $false)][hashtable]$Headers
    )
    $ex = $null
    $hreType = 'Microsoft.PowerShell.Commands.HttpResponseException' -as [type]
    if ($null -ne $hreType) {
        $message = New-Object System.Net.Http.HttpResponseMessage ([System.Net.HttpStatusCode]$StatusCode)
        if ($Headers) { foreach ($k in $Headers.Keys) { $message.Headers.TryAddWithoutValidation([string]$k, [string]$Headers[$k]) | Out-Null } }
        $ex = New-Object Microsoft.PowerShell.Commands.HttpResponseException (('Response status code does not indicate success: {0}.' -f $StatusCode), $message)
    }
    else {
        $ex = New-Object System.Exception (('The remote server returned an error: ({0}).' -f $StatusCode))
        $h = @{}
        if ($Headers) { foreach ($k in $Headers.Keys) { $h[[string]$k] = [string]$Headers[$k] } }
        $ex | Add-Member -MemberType NoteProperty -Name 'Response' -Value ([pscustomobject]@{ StatusCode = $StatusCode; Headers = $h })
    }
    $record = New-Object System.Management.Automation.ErrorRecord ($ex, ('WebCmdletWebResponseException,{0}' -f $StatusCode), [System.Management.Automation.ErrorCategory]::InvalidOperation, $null)
    if (-not [string]::IsNullOrEmpty($Body)) { $record.ErrorDetails = New-Object System.Management.Automation.ErrorDetails ($Body) }
    return $record
}

function New-IQTestWebException {
    <#
    .SYNOPSIS
    Builds an ErrorRecord wrapping a System.Net.WebException with the given status (default Timeout, no response).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Status = 'Timeout')
    $ex = New-Object System.Net.WebException ('The operation has timed out', [System.Net.WebExceptionStatus]$Status)
    return (New-Object System.Management.Automation.ErrorRecord ($ex, 'WebCmdletWebResponseException', [System.Management.Automation.ErrorCategory]::InvalidOperation, $null))
}

function Get-IQTestSheetHeader {
    <#
    .SYNOPSIS
    Header cells (row 1) of a worksheet, read raw so header-only sheets (zero data rows) still return their columns.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Sheet
    )
    $rows = @(Import-Excel -Path $Path -WorksheetName $Sheet -NoHeader -WarningAction SilentlyContinue)
    if ($rows.Count -eq 0) { return @() }
    return @($rows[0].PSObject.Properties.Value | ForEach-Object { [string]$_ })
}

function Get-IQTestGen2Definition {
    <#
    .SYNOPSIS
    The object Invoke-IQFabricLro returns for a Gen2 dataflow getDefinition, built from the .pq / queryMetadata fixtures.
    #>
    [CmdletBinding()]
    param()
    $pq = [System.IO.File]::ReadAllBytes((Get-IQTestFixturePath -Relative 'dataflows/gen2-mashup.pq'))
    $meta = [System.IO.File]::ReadAllBytes((Get-IQTestFixturePath -Relative 'dataflows/gen2-queryMetadata.json'))
    return ([pscustomobject]@{
            definition = [pscustomobject]@{
                parts = @(
                    [pscustomobject]@{ path = 'queryMetadata.json'; payloadType = 'InlineBase64'; payload = [Convert]::ToBase64String($meta) },
                    [pscustomobject]@{ path = 'mashup.pq'; payloadType = 'InlineBase64'; payload = [Convert]::ToBase64String($pq) },
                    [pscustomobject]@{ path = '.platform'; payloadType = 'InlineBase64'; payload = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('{"metadata":{"type":"Dataflow"}}')) },
                    [pscustomobject]@{ path = '../evil.txt'; payloadType = 'InlineBase64'; payload = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('x')) }
                )
            }
        })
}

function Copy-IQTestFixtureFolder {
    <#
    .SYNOPSIS
    Copies every file of a fixture folder into a destination folder (created when missing).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Relative,
        [Parameter(Mandatory = $true)][string]$Destination
    )
    $source = Get-IQTestFixturePath -Relative $Relative
    if (-not (Test-Path -LiteralPath $Destination)) { New-Item -ItemType Directory -Path $Destination -Force | Out-Null }
    Get-ChildItem -LiteralPath $source -File | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $Destination $_.Name) -Force }
}
