#Requires -Version 5.1
<#
.SYNOPSIS
    ImpactIQ v3 - Auth module: token providers for Power BI and Fabric without a service principal.

.DESCRIPTION
    Dot-sourced by ImpactIQ.ps1 after ImpactIQ.Common.ps1. Implements brief sections 2.2, 3 and 4:

      Interactive  - MicrosoftPowerBIMgmt (Connect-PowerBIServiceAccount / Get-PowerBIAccessToken), the monolith's
                     Connect-PowerBI retry loop (Final PS Script.txt 455-503, 560-575); Fabric token best-effort via
                     Az.Accounts exactly like the monolith's Get-FabricAccessToken (504-553, SecureString handling).
      DeviceCode   - pure HTTP OAuth 2.0 v2.0 device-code flow + refresh-token grant (no module needed), optional
                     Teams/Slack webhook post of the sign-in message, encrypted refresh-token cache.
      Credential   - HTTP ROPC (password grant) with AADSTS error mapping; falls back to
                     Connect-PowerBIServiceAccount -Credential when the module is available.
      AzContext    - existing Az.Accounts context (Get-AzAccessToken -ResourceUrl ...), never logs in by itself.
      AccessToken  - IMPACTIQ_PBI_TOKEN (+ IMPACTIQ_FABRIC_TOKEN), no refresh possible.

    All HTTP in this module uses Invoke-RestMethod directly (the Http module needs tokens, so it cannot be used here)
    with a small retry on transient errors. Invoke-RestMethod and Start-IQAuthSleep are plain command calls so Pester
    can Mock them. Provider state lives in $script:IQ.Auth (one hashtable). Secrets (refresh tokens, passwords, cache
    keys) are never logged.

    Windows PowerShell 5.1 compatible (no ternary, no "??", no "?.", no "using namespace", no $IsWindows).
#>

# Cross-module functions used (brief section 2.1): Write-IQLog, Test-IQInteractive, Get-IQEnvironmentSettings,
# ConvertTo-IQJsonFile, ConvertFrom-IQJsonFile.
# Private helpers are prefixed *-IQAuth* / *-IQJwt* / *-IQTokenCache* and are not part of the contract.

$script:IQAuthDefaultClientId = '1950a258-227b-4e31-a9cf-717495945fc2'   # Azure PowerShell first-party public client
$script:IQAuthDeviceGrant = 'urn:ietf:params:oauth:grant-type:device_code'
$script:IQAuthRefreshSkewMinutes = 5

# ---------------------------------------------------------------------------------------------------------------------
# Small private helpers (sleep, endpoints, scopes, HTTP, JWT)
# ---------------------------------------------------------------------------------------------------------------------

function Start-IQAuthSleep {
    <#
    .SYNOPSIS
        Sleeps for N seconds (private, mockable by tests so the device-code polling loop runs instantly).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true, Position = 0)][int]$Seconds)
    if ($Seconds -gt 0) { Start-Sleep -Seconds $Seconds }
}

function Get-IQAuthState {
    <#
    .SYNOPSIS
        Returns $script:IQ.Auth, throwing a clear message when Initialize-IQAuth has not run (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][switch]$AllowUninitialized)
    if ($null -eq $script:IQ) { throw 'ImpactIQ context is not initialised (Initialize-IQContext must run before Initialize-IQAuth).' }
    $auth = $script:IQ.Auth
    if ($null -eq $auth -or -not ($auth -is [hashtable]) -or -not $auth.ContainsKey('Initialized') -or -not $auth.Initialized) {
        if ($AllowUninitialized) { return $auth }
        throw 'Authentication is not initialised. Call Initialize-IQAuth before requesting tokens.'
    }
    return $auth
}

function Get-IQAuthEndpoint {
    <#
    .SYNOPSIS
        Builds the OAuth 2.0 v2.0 endpoint URL ({Authority}/{TenantId}/oauth2/v2.0/devicecode|token) for the current or given authority.
    .DESCRIPTION
        When the authority already carries a tenant path segment (e.g. an -AuthorityOverride of
        https://login.microsoftonline.com/contoso.onmicrosoft.com) the tenant id is not appended again.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('DeviceCode', 'Token')][string]$Kind,
        [Parameter(Mandatory = $false)][string]$Authority,
        [Parameter(Mandatory = $false)][string]$TenantId
    )
    $auth = $null
    if ($script:IQ -and $script:IQ.Auth -is [hashtable]) { $auth = $script:IQ.Auth }
    if ([string]::IsNullOrWhiteSpace($Authority)) {
        if ($auth -and $auth.ContainsKey('Authority') -and $auth.Authority) { $Authority = [string]$auth.Authority }
        elseif ($script:IQ -and $script:IQ.Endpoints) { $Authority = [string]$script:IQ.Endpoints.Authority }
    }
    if ([string]::IsNullOrWhiteSpace($Authority)) { throw 'No OAuth authority is configured (environment not resolved).' }
    if ([string]::IsNullOrWhiteSpace($TenantId)) {
        if ($auth -and $auth.ContainsKey('TenantId') -and $auth.TenantId) { $TenantId = [string]$auth.TenantId }
        else { $TenantId = 'organizations' }
    }
    $base = $Authority.Trim().TrimEnd('/')
    # "https://host" has no tenant segment; "https://host/tenant" already has one.
    $hasTenantSegment = ($base -match '^https?://[^/]+/[^/]+')
    if (-not $hasTenantSegment) { $base = $base + '/' + $TenantId.Trim().Trim('/') }
    $leaf = 'token'
    if ($Kind -eq 'DeviceCode') { $leaf = 'devicecode' }
    return ($base + '/oauth2/v2.0/' + $leaf)
}

function Get-IQAuthResourceUrl {
    <#
    .SYNOPSIS
        Returns the OAuth resource (audience) URL for PowerBI or Fabric from the environment table (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateSet('PowerBI', 'Fabric')][string]$Resource)
    if ($null -eq $script:IQ -or $null -eq $script:IQ.Endpoints) { throw 'Environment endpoints are not resolved (Set-IQEnvironment / Initialize-IQAuth -Environment).' }
    $ep = $script:IQ.Endpoints
    $url = $null
    if ($Resource -eq 'PowerBI') { $url = [string]$ep.PowerBIResource }
    else {
        if ($ep.ContainsKey('FabricResource') -and $ep.FabricResource) { $url = [string]$ep.FabricResource } else { $url = [string]$ep.FabricApiPrefix }
    }
    if ([string]::IsNullOrWhiteSpace($url)) { throw "No resource URL configured for $Resource." }
    return $url.TrimEnd('/')
}

function Get-IQAuthScope {
    <#
    .SYNOPSIS
        Builds the v2.0 scope string: "<resource>/.default offline_access" (+ "openid profile" for an initial sign-in).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('PowerBI', 'Fabric')][string]$Resource,
        [Parameter(Mandatory = $false)][switch]$IncludeProfile
    )
    $scope = (Get-IQAuthResourceUrl -Resource $Resource) + '/.default offline_access'
    if ($IncludeProfile) { $scope = $scope + ' openid profile' }
    return $scope
}

function Get-IQAuthHttpErrorInfo {
    <#
    .SYNOPSIS
        Extracts StatusCode, body and transient flag from an Invoke-RestMethod error on both 5.1 (WebException) and 7 (HttpResponseException).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord)
    $info = @{ StatusCode = $null; Body = ''; Message = ''; Transient = $false }
    $ex = $ErrorRecord.Exception
    if ($null -eq $ex) { $info.Message = [string]$ErrorRecord; return $info }
    $info.Message = [string]$ex.Message
    try {
        if ($ErrorRecord.ErrorDetails -and -not [string]::IsNullOrEmpty($ErrorRecord.ErrorDetails.Message)) {
            $info.Body = [string]$ErrorRecord.ErrorDetails.Message
        }
    }
    catch { $info.Body = '' }

    $response = $null
    try { if ($ex.PSObject.Properties['Response']) { $response = $ex.Response } } catch { $response = $null }
    if ($null -ne $response) {
        try { if ($response.PSObject.Properties['StatusCode'] -and $null -ne $response.StatusCode) { $info.StatusCode = [int]$response.StatusCode } } catch { $info.StatusCode = $null }
        # Windows PowerShell 5.1: WebException.Response is an HttpWebResponse; read the body from the stream.
        if ([string]::IsNullOrEmpty($info.Body) -and $response.PSObject.Methods['GetResponseStream']) {
            try {
                $stream = $response.GetResponseStream()
                if ($null -ne $stream) {
                    if ($stream.CanSeek) { $stream.Position = 0 }
                    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
                    try { $info.Body = $reader.ReadToEnd() } finally { $reader.Dispose() }
                }
            }
            catch { $info.Body = '' }
        }
        # PowerShell 7: HttpResponseMessage.Content (rarely needed; ErrorDetails normally carries the body).
        if ([string]::IsNullOrEmpty($info.Body) -and $response.PSObject.Properties['Content'] -and $null -ne $response.Content) {
            try {
                $task = $response.Content.ReadAsStringAsync()
                $task.Wait(5000) | Out-Null
                if ($task.IsCompleted) { $info.Body = [string]$task.Result }
            }
            catch { $info.Body = '' }
        }
    }
    if ($null -eq $info.StatusCode) {
        $webStatus = $null
        try { if ($ex.GetType().FullName -eq 'System.Net.WebException' -and $ex.PSObject.Properties['Status']) { $webStatus = [string]$ex.Status } } catch { $webStatus = $null }
        if ($webStatus -and $webStatus -in @('Timeout', 'ConnectFailure', 'NameResolutionFailure', 'ReceiveFailure', 'SendFailure', 'ConnectionClosed', 'KeepAliveFailure', 'ProxyNameResolutionFailure', 'RequestCanceled', 'UnknownError')) {
            $info.Transient = $true
        }
        $walk = $ex
        $depth = 0
        while ($null -ne $walk -and $depth -lt 6 -and -not $info.Transient) {
            $name = $walk.GetType().FullName
            if ($name -match 'HttpRequestException|TaskCanceledException|OperationCanceledException|SocketException|System\.IO\.IOException|WebException|TimeoutException') { $info.Transient = $true }
            elseif ([string]$walk.Message -match 'timed out|was canceled|forcibly closed|Unable to connect|No such host|actively refused|temporarily unavailable') { $info.Transient = $true }
            $walk = $walk.InnerException
            $depth++
        }
    }
    return $info
}

function ConvertFrom-IQAuthErrorBody {
    <#
    .SYNOPSIS
        Parses an Entra token-endpoint error body into @{Error; ErrorDescription; ErrorCodes; AadCodes} (private, never throws).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Body)
    $out = @{ Error = $null; ErrorDescription = $null; ErrorCodes = @(); AadCodes = @() }
    if ([string]::IsNullOrWhiteSpace($Body)) { return $out }
    try {
        $obj = ConvertFrom-Json -InputObject $Body -ErrorAction Stop
        if ($obj.PSObject.Properties['error']) {
            $e = $obj.error
            # Some endpoints nest {"error":{"code":..,"message":..}}.
            if ($e -is [string]) { $out.Error = [string]$e }
            elseif ($null -ne $e) {
                if ($e.PSObject.Properties['code']) { $out.Error = [string]$e.code }
                if ($e.PSObject.Properties['message']) { $out.ErrorDescription = [string]$e.message }
            }
        }
        if ($obj.PSObject.Properties['error_description']) { $out.ErrorDescription = [string]$obj.error_description }
        if ($obj.PSObject.Properties['error_codes'] -and $null -ne $obj.error_codes) { $out.ErrorCodes = @($obj.error_codes | ForEach-Object { [string]$_ }) }
    }
    catch {
        $out.ErrorDescription = $Body
    }
    $codes = New-Object System.Collections.Generic.List[string]
    foreach ($c in $out.ErrorCodes) { if (-not $codes.Contains($c)) { $codes.Add($c) } }
    foreach ($m in [regex]::Matches([string]$Body, 'AADSTS(\d+)')) { $v = $m.Groups[1].Value; if (-not $codes.Contains($v)) { $codes.Add($v) } }
    $out.AadCodes = $codes.ToArray()
    return $out
}

function Invoke-IQAuthRequest {
    <#
    .SYNOPSIS
        POSTs a form body to an OAuth endpoint with Invoke-RestMethod; returns @{Ok; Response | StatusCode; Error; ErrorDescription; AadCodes; Message} (never throws).
    .DESCRIPTION
        Transient failures (no HTTP status and a network-level exception, or HTTP 5xx/408/429) are retried up to
        MaxAttempts with 2,4,8 s backoff via Start-IQAuthSleep. OAuth errors (400 with an "error" body) are returned
        to the caller for interpretation (authorization_pending, invalid_grant, ...).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][hashtable]$Body,
        [Parameter(Mandatory = $false)][ValidateRange(1, 10)][int]$MaxAttempts = 3,
        [Parameter(Mandatory = $false)][ValidateRange(5, 600)][int]$TimeoutSec = 60,
        [Parameter(Mandatory = $false)][string]$Description = 'token request'
    )
    $result = @{ Ok = $false; Response = $null; StatusCode = $null; Error = $null; ErrorDescription = $null; AadCodes = @(); Message = $null; Body = '' }
    $attempt = 0
    while ($attempt -lt $MaxAttempts) {
        $attempt++
        try {
            $previousProgress = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            try {
                $response = Invoke-RestMethod -Method Post -Uri $Uri -Body $Body -ContentType 'application/x-www-form-urlencoded' `
                    -TimeoutSec $TimeoutSec -UseBasicParsing -ErrorAction Stop
            }
            finally { $ProgressPreference = $previousProgress }
            $result.Ok = $true
            $result.Response = $response
            return $result
        }
        catch {
            $info = Get-IQAuthHttpErrorInfo -ErrorRecord $_
            # A thrown JSON string (e.g. from a test mock) carries the OAuth error in the message, not in a body.
            if ([string]::IsNullOrWhiteSpace($info.Body) -and [string]$info.Message -match '^\s*\{') { $info.Body = [string]$info.Message }
            $parsed = ConvertFrom-IQAuthErrorBody -Body $info.Body
            $result.StatusCode = $info.StatusCode
            $result.Body = $info.Body
            $result.Error = $parsed.Error
            $result.ErrorDescription = $parsed.ErrorDescription
            $result.AadCodes = $parsed.AadCodes
            $result.Message = $info.Message
            $status = 0
            if ($null -ne $info.StatusCode) { $status = [int]$info.StatusCode }
            $retryable = ($status -ge 500) -or ($status -eq 408) -or ($status -eq 429) -or ($status -eq 0 -and $info.Transient)
            if ($retryable -and $attempt -lt $MaxAttempts) {
                $delay = [int][math]::Pow(2, $attempt)
                Write-IQLog -Level Warn -Stage Auth -Message ("{0} failed transiently (HTTP {1}: {2}); retrying in {3} s (attempt {4}/{5})." -f $Description, $status, $info.Message, $delay, $attempt, $MaxAttempts)
                Start-IQAuthSleep -Seconds $delay
                continue
            }
            return $result
        }
    }
    return $result
}

function Test-IQAuthPermanentFailure {
    <#
    .SYNOPSIS
        $true when a failed Invoke-IQAuthRequest result is an OAuth/Entra rejection (4xx or an error code) rather than a transient network/5xx failure.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Result)
    if ($null -ne $Result.StatusCode -and [int]$Result.StatusCode -ge 400 -and [int]$Result.StatusCode -lt 500) { return $true }
    $code = [string]$Result.Error
    if ([string]::IsNullOrWhiteSpace($code)) { return $false }
    if ($code -in @('temporarily_unavailable', 'server_error', 'no_refresh_token')) { return $false }
    return $true
}

function Get-IQAuthResultText {
    <#
    .SYNOPSIS
        One-line human description of a failed Invoke-IQAuthRequest result (no secrets).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Result)
    $parts = New-Object System.Collections.Generic.List[string]
    if ($Result.StatusCode) { $parts.Add('HTTP ' + $Result.StatusCode) }
    if ($Result.Error) { $parts.Add([string]$Result.Error) }
    if ($Result.ErrorDescription) {
        $d = [string]$Result.ErrorDescription
        $nl = $d.IndexOfAny([char[]]@("`r", "`n"))
        if ($nl -gt 0) { $d = $d.Substring(0, $nl) }
        if ($d.Length -gt 300) { $d = $d.Substring(0, 297) + '...' }
        $parts.Add($d)
    }
    elseif ($Result.Message) { $parts.Add([string]$Result.Message) }
    if ($parts.Count -eq 0) { return 'unknown error' }
    return [string]::Join(' - ', $parts.ToArray())
}

function Get-IQJwtPayload {
    <#
    .SYNOPSIS
        Decodes the payload (second segment) of a JWT with base64url handling and padding fix; returns the JSON string or $null.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Token)
    if ([string]::IsNullOrWhiteSpace($Token)) { return $null }
    try {
        $t = $Token.Trim()
        if ($t -match '^(?i)Bearer\s+') { $t = $t -replace '^(?i)Bearer\s+', '' }
        $segments = $t.Split('.')
        if ($segments.Length -lt 2) { return $null }
        $p = $segments[1].Replace('-', '+').Replace('_', '/')
        switch ($p.Length % 4) {
            2 { $p += '==' }
            3 { $p += '=' }
            1 { return $null }
        }
        $bytes = [Convert]::FromBase64String($p)
        return [System.Text.Encoding]::UTF8.GetString($bytes)
    }
    catch { return $null }
}

function Get-IQJwtClaimSet {
    <#
    .SYNOPSIS
        Returns the claims of a JWT as an object (ConvertFrom-Json of the payload), or $null when unparsable.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Token)
    $json = Get-IQJwtPayload -Token $Token
    if ([string]::IsNullOrWhiteSpace($json)) { return $null }
    try { return (ConvertFrom-Json -InputObject $json -ErrorAction Stop) } catch { return $null }
}

function Get-IQJwtExpiry {
    <#
    .SYNOPSIS
        [datetime] UTC from the JWT "exp" claim (base64url decode with padding fix), or [datetime]::MaxValue if unparsable.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false, Position = 0)][AllowNull()][AllowEmptyString()][string]$Token)
    $claims = Get-IQJwtClaimSet -Token $Token
    if ($null -eq $claims) { return [datetime]::MaxValue }
    try {
        if (-not $claims.PSObject.Properties['exp']) { return [datetime]::MaxValue }
        $exp = [int64]$claims.exp
        if ($exp -le 0) { return [datetime]::MaxValue }
        $epoch = New-Object System.DateTime(1970, 1, 1, 0, 0, 0, [System.DateTimeKind]::Utc)
        return $epoch.AddSeconds([double]$exp)
    }
    catch { return [datetime]::MaxValue }
}

function Get-IQJwtAccount {
    <#
    .SYNOPSIS
        Best-effort user identifier from a JWT (preferred_username, upn, unique_name, email) or $null.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Token)
    $claims = Get-IQJwtClaimSet -Token $Token
    if ($null -eq $claims) { return $null }
    foreach ($name in @('preferred_username', 'upn', 'unique_name', 'email')) {
        if ($claims.PSObject.Properties[$name] -and -not [string]::IsNullOrWhiteSpace([string]$claims.$name)) { return [string]$claims.$name }
    }
    return $null
}

function ConvertFrom-IQSecureString {
    <#
    .SYNOPSIS
        Converts a SecureString to plain text via BSTR (works on 5.1 and 7; zeroes the unmanaged copy) - from the monolith's Get-FabricAccessToken.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][System.Security.SecureString]$SecureString)
    $pointer = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try { return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer) }
    finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer) }
}

function ConvertTo-IQPlainTokenValue {
    <#
    .SYNOPSIS
        Normalises a token value that may be a string, a SecureString (Az 14+), or an object with .Token/.Authorization; strips "Bearer ".
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Security.SecureString]) { return (ConvertFrom-IQSecureString -SecureString $Value) }
    if ($Value -is [string]) { return ($Value -replace '^(?i)Bearer\s+', '').Trim() }
    if ($Value -is [System.Collections.IDictionary]) {
        if ($Value.Contains('Authorization')) { return (ConvertTo-IQPlainTokenValue -Value $Value['Authorization']) }
        if ($Value.Contains('Token')) { return (ConvertTo-IQPlainTokenValue -Value $Value['Token']) }
        return $null
    }
    if ($Value.PSObject.Properties['Token']) { return (ConvertTo-IQPlainTokenValue -Value $Value.Token) }
    if ($Value.PSObject.Properties['Authorization']) { return (ConvertTo-IQPlainTokenValue -Value $Value.Authorization) }
    return ([string]$Value -replace '^(?i)Bearer\s+', '').Trim()
}

function Get-IQAuthResponseAccessToken {
    <#
    .SYNOPSIS
        Returns the access_token string from a token-endpoint response object, or $null when absent (strict-mode safe).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Response)
    if ($null -eq $Response) { return $null }
    try {
        if ($Response -is [System.Collections.IDictionary]) {
            if ($Response.Contains('access_token')) { return [string]$Response['access_token'] }
            return $null
        }
        if ($Response.PSObject.Properties['access_token']) { return [string]$Response.access_token }
    }
    catch { return $null }
    return $null
}

function Register-IQAuthSecret {
    <#
    .SYNOPSIS
        On Azure DevOps, registers a value as a secret so the agent masks it in every log line (audit C2-05). No-op elsewhere.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return }
    if ($null -eq $script:IQ -or -not $script:IQ.IsAzureDevOps) { return }
    # Deliberate direct Write-Host: this is an agent logging command consumed by Azure Pipelines, not a log line.
    # It must not pass through Write-IQLog (which would also append the value to the log file).
    try { Write-Host ('##vso[task.setsecret]' + $Value) } catch { $null = $_.Exception }
}

function Test-IQAuthGuid {
    <#
    .SYNOPSIS
        $true when the value is a GUID (used to decide whether a tenant id can be pinned on module cmdlets).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $g = [guid]::Empty
    return [guid]::TryParse($Value.Trim(), [ref]$g)
}

function Import-IQAuthModule {
    <#
    .SYNOPSIS
        Imports a PowerShell module if available; returns $true/$false (never throws; logs at Debug).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Name)
    try {
        if (Get-Module -Name $Name) { return $true }
        Import-Module -Name $Name -ErrorAction Stop -WarningAction SilentlyContinue -Verbose:$false 3>$null | Out-Null
        return $true
    }
    catch {
        Write-IQLog -Level Debug -Stage Auth -Message ("Module '{0}' is not available: {1}" -f $Name, $_.Exception.Message)
        return $false
    }
}

# ---------------------------------------------------------------------------------------------------------------------
# Token bookkeeping
# ---------------------------------------------------------------------------------------------------------------------

function Set-IQAuthToken {
    <#
    .SYNOPSIS
        Stores an access token for a resource in $script:IQ.Auth.Tokens (expiry from the JWT, else expires_in) and returns it (private).
    .DESCRIPTION
        When -Response (a token-endpoint response) carries refresh_token it replaces the stored one and, when
        -PersistCache is set, Save-IQTokenCache is called so the newest refresh token is on disk (brief 4.2).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('PowerBI', 'Fabric')][string]$Resource,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$AccessToken,
        [Parameter(Mandatory = $false)][AllowNull()]$Response,
        [Parameter(Mandatory = $false)][string]$Source,
        [Parameter(Mandatory = $false)][switch]$PersistCache
    )
    $auth = Get-IQAuthState -AllowUninitialized
    $AccessToken = ConvertTo-IQPlainTokenValue -Value $AccessToken
    if ([string]::IsNullOrWhiteSpace($AccessToken)) { throw "The $Resource token provider returned an empty access token." }

    $expires = Get-IQJwtExpiry -Token $AccessToken
    if ($expires -eq [datetime]::MaxValue -and $null -ne $Response) {
        try {
            if ($Response.PSObject.Properties['expires_in'] -and [int64]$Response.expires_in -gt 0) {
                $expires = [datetime]::UtcNow.AddSeconds([double]$Response.expires_in)
            }
        }
        catch { $expires = [datetime]::MaxValue }
    }
    if (-not ($auth.Tokens -is [hashtable])) { $auth.Tokens = @{} }
    $auth.Tokens[$Resource] = @{
        AccessToken = $AccessToken
        ExpiresUtc  = $expires
        ObtainedUtc = [datetime]::UtcNow
        Resource    = (Get-IQAuthResourceUrl -Resource $Resource)
        Source      = $Source
    }
    Register-IQAuthSecret -Value $AccessToken

    if ($null -ne $Response) {
        try {
            if ($Response.PSObject.Properties['refresh_token'] -and -not [string]::IsNullOrWhiteSpace([string]$Response.refresh_token)) {
                $auth.RefreshToken = [string]$Response.refresh_token
                Register-IQAuthSecret -Value $auth.RefreshToken
                if ($PersistCache) { Save-IQTokenCache | Out-Null }
            }
            $account = $null
            if ($Response.PSObject.Properties['id_token']) { $account = Get-IQJwtAccount -Token ([string]$Response.id_token) }
            if ([string]::IsNullOrWhiteSpace($account)) { $account = Get-IQJwtAccount -Token $AccessToken }
            if (-not [string]::IsNullOrWhiteSpace($account)) { $auth.Account = $account }
        }
        catch { Write-IQLog -Level Debug -Stage Auth -Message ('Could not read refresh token / account from the token response: ' + $_.Exception.Message) }
    }
    elseif ([string]::IsNullOrWhiteSpace([string]$auth.Account)) {
        $account = Get-IQJwtAccount -Token $AccessToken
        if (-not [string]::IsNullOrWhiteSpace($account)) { $auth.Account = $account }
    }
    if ([string]::IsNullOrWhiteSpace([string]$auth.TenantIdResolved)) {
        $claims = Get-IQJwtClaimSet -Token $AccessToken
        if ($null -ne $claims -and $claims.PSObject.Properties['tid']) { $auth.TenantIdResolved = [string]$claims.tid }
    }

    $remaining = 'unknown'
    if ($expires -ne [datetime]::MaxValue) { $remaining = [string][int][math]::Round(($expires - [datetime]::UtcNow).TotalMinutes) + ' min' }
    Write-IQLog -Level Debug -Stage Auth -Message ("{0} token obtained via {1}; valid for {2}." -f $Resource, $Source, $remaining)
    return $AccessToken
}

function Get-IQAuthCachedToken {
    <#
    .SYNOPSIS
        Returns the cached access token for a resource when more than the refresh skew remains, else $null (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateSet('PowerBI', 'Fabric')][string]$Resource)
    $auth = Get-IQAuthState -AllowUninitialized
    if ($null -eq $auth -or -not ($auth.Tokens -is [hashtable]) -or -not $auth.Tokens.ContainsKey($Resource)) { return $null }
    $entry = $auth.Tokens[$Resource]
    if ($null -eq $entry -or [string]::IsNullOrWhiteSpace([string]$entry.AccessToken)) { return $null }
    $expires = [datetime]$entry.ExpiresUtc
    if ($expires -gt [datetime]::UtcNow.AddMinutes($script:IQAuthRefreshSkewMinutes)) { return [string]$entry.AccessToken }
    return $null
}

# ---------------------------------------------------------------------------------------------------------------------
# HTTP grants: device code, refresh token, password (ROPC)
# ---------------------------------------------------------------------------------------------------------------------

function Send-IQDeviceCodeWebhook {
    <#
    .SYNOPSIS
        POSTs { "text": "<message>" } to a Teams/Slack incoming webhook so a human learns the device code (never throws).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$Message
    )
    try {
        $payload = ConvertTo-Json -InputObject @{ text = $Message } -Depth 20 -Compress
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
        Invoke-RestMethod -Method Post -Uri $Url -Body $bytes -ContentType 'application/json; charset=utf-8' -TimeoutSec 30 -UseBasicParsing -ErrorAction Stop | Out-Null
        Write-IQLog -Level Info -Stage Auth -Message 'Device-code sign-in message posted to the configured webhook.'
    }
    catch {
        Write-IQLog -Level Warn -Stage Auth -Message ('Could not post the device-code message to the webhook: ' + $_.Exception.Message)
    }
}

function Invoke-IQDeviceCodeFlow {
    <#
    .SYNOPSIS
        Runs the OAuth 2.0 device-code flow (devicecode + polling of the token endpoint) and returns the token response object.
    .DESCRIPTION
        Prints the sign-in message at Warn (visible in pipeline logs), posts it to -WebhookUrl when set, then polls
        every "interval" seconds handling authorization_pending, slow_down (+5 s), expired_token (throw) and
        authorization_declined (throw). Invoke-RestMethod (via Invoke-IQAuthRequest) and Start-IQAuthSleep are mockable.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$Authority,
        [Parameter(Mandatory = $false)][string]$TenantId,
        [Parameter(Mandatory = $false)][string]$ClientId,
        [Parameter(Mandatory = $false)][string]$Scope,
        [Parameter(Mandatory = $false)][string]$WebhookUrl,
        [Parameter(Mandatory = $false)][ValidateRange(1, 3600)][int]$MaxWaitSeconds = 0
    )
    $auth = Get-IQAuthState -AllowUninitialized
    if ([string]::IsNullOrWhiteSpace($ClientId)) { if ($auth -and $auth.ClientId) { $ClientId = [string]$auth.ClientId } else { $ClientId = $script:IQAuthDefaultClientId } }
    if ([string]::IsNullOrWhiteSpace($Scope)) { $Scope = Get-IQAuthScope -Resource PowerBI -IncludeProfile }
    if ([string]::IsNullOrWhiteSpace($WebhookUrl) -and $auth -and $auth.ContainsKey('DeviceCodeWebhookUrl')) { $WebhookUrl = [string]$auth.DeviceCodeWebhookUrl }

    $deviceUrl = Get-IQAuthEndpoint -Kind DeviceCode -Authority $Authority -TenantId $TenantId
    $tokenUrl = Get-IQAuthEndpoint -Kind Token -Authority $Authority -TenantId $TenantId

    Write-IQLog -Level Info -Stage Auth -Message ("Requesting a device code from {0} (client {1})." -f $deviceUrl, $ClientId)
    $start = Invoke-IQAuthRequest -Uri $deviceUrl -Body @{ client_id = $ClientId; scope = $Scope } -Description 'device-code request'
    if (-not $start.Ok) { throw ('Device-code request failed: ' + (Get-IQAuthResultText -Result $start)) }
    $dc = $start.Response
    if ($null -eq $dc -or -not $dc.PSObject.Properties['device_code'] -or [string]::IsNullOrWhiteSpace([string]$dc.device_code)) {
        throw 'Device-code request returned no device_code.'
    }

    $interval = 5
    try { if ($dc.PSObject.Properties['interval'] -and [int]$dc.interval -gt 0) { $interval = [int]$dc.interval } } catch { $interval = 5 }
    $expiresIn = 900
    try { if ($dc.PSObject.Properties['expires_in'] -and [int]$dc.expires_in -gt 0) { $expiresIn = [int]$dc.expires_in } } catch { $expiresIn = 900 }
    if ($MaxWaitSeconds -gt 0 -and $MaxWaitSeconds -lt $expiresIn) { $expiresIn = $MaxWaitSeconds }

    $verificationUri = $null
    if ($dc.PSObject.Properties['verification_uri']) { $verificationUri = [string]$dc.verification_uri }
    $userCode = $null
    if ($dc.PSObject.Properties['user_code']) { $userCode = [string]$dc.user_code }
    $message = $null
    if ($dc.PSObject.Properties['message']) { $message = [string]$dc.message }
    if ([string]::IsNullOrWhiteSpace($message)) { $message = "To sign in, use a web browser to open the page $verificationUri and enter the code $userCode to authenticate." }

    # Always print what the endpoint returned (brief section 3): commercial -> https://microsoft.com/devicelogin, else verification_uri.
    Write-IQLog -Level Warn -Stage Auth -Message ('DEVICE CODE SIGN-IN REQUIRED: ' + $message)
    Write-IQLog -Level Info -Stage Auth -Message ("Waiting up to {0} minutes for a user to complete the sign-in (URL: {1}, code: {2})." -f [int][math]::Ceiling($expiresIn / 60.0), $verificationUri, $userCode)
    if (-not [string]::IsNullOrWhiteSpace($WebhookUrl)) { Send-IQDeviceCodeWebhook -Url $WebhookUrl -Message ('ImpactIQ needs a sign-in: ' + $message) }

    $deadline = [datetime]::UtcNow.AddSeconds($expiresIn)
    $maxPolls = [int][math]::Ceiling($expiresIn / [math]::Max(1, $interval)) + 20
    $polls = 0
    $pollBody = @{ grant_type = $script:IQAuthDeviceGrant; client_id = $ClientId; device_code = [string]$dc.device_code }
    while ($true) {
        $polls++
        if ($polls -gt $maxPolls) { throw 'Device-code sign-in was not completed in time (poll limit reached).' }
        Start-IQAuthSleep -Seconds $interval
        $poll = Invoke-IQAuthRequest -Uri $tokenUrl -Body $pollBody -MaxAttempts 3 -Description 'device-code poll'
        if ($poll.Ok) {
            $accessToken = Get-IQAuthResponseAccessToken -Response $poll.Response
            if ([string]::IsNullOrWhiteSpace($accessToken)) { throw 'Device-code token response contained no access_token.' }
            Write-IQLog -Level Success -Stage Auth -Message 'Device-code sign-in completed.'
            return $poll.Response
        }
        $err = [string]$poll.Error
        if ($err -eq 'authorization_pending') {
            if ([datetime]::UtcNow -gt $deadline) { throw 'Device-code sign-in was not completed before the code expired.' }
        }
        elseif ($err -eq 'slow_down') {
            $interval += 5
            Write-IQLog -Level Debug -Stage Auth -Message ("Token endpoint asked to slow down; polling every {0} s." -f $interval)
        }
        elseif ($err -eq 'expired_token') {
            throw 'The device code expired before anyone completed the sign-in. Re-run and complete the sign-in within the time shown.'
        }
        elseif ($err -eq 'authorization_declined') {
            throw 'The device-code sign-in was declined by the user.'
        }
        elseif ($err -eq 'bad_verification_code') {
            throw ('The token endpoint rejected the device code: ' + (Get-IQAuthResultText -Result $poll))
        }
        else {
            $text = Get-IQAuthResultText -Result $poll
            if ($poll.AadCodes -contains '50076' -or $poll.AadCodes -contains '50079' -or $poll.AadCodes -contains '53003' -or $poll.AadCodes -contains '530036') {
                throw ('Device-code sign-in blocked by Conditional Access / MFA policy: ' + $text)
            }
            throw ('Device-code sign-in failed: ' + $text)
        }
    }
}

function Invoke-IQRefreshTokenGrant {
    <#
    .SYNOPSIS
        Redeems the stored refresh token for an access token of the given resource (refresh_token grant); returns the Invoke-IQAuthRequest result.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('PowerBI', 'Fabric')][string]$Resource,
        [Parameter(Mandatory = $false)][string]$RefreshToken
    )
    $auth = Get-IQAuthState -AllowUninitialized
    if ([string]::IsNullOrWhiteSpace($RefreshToken)) { $RefreshToken = [string]$auth.RefreshToken }
    if ([string]::IsNullOrWhiteSpace($RefreshToken)) {
        return @{ Ok = $false; Response = $null; StatusCode = $null; Error = 'no_refresh_token'; ErrorDescription = 'No refresh token is available.'; AadCodes = @(); Message = 'No refresh token is available.'; Body = '' }
    }
    $body = @{
        grant_type    = 'refresh_token'
        client_id     = [string]$auth.ClientId
        scope         = (Get-IQAuthScope -Resource $Resource)
        refresh_token = $RefreshToken
    }
    $tokenUrl = Get-IQAuthEndpoint -Kind Token
    return (Invoke-IQAuthRequest -Uri $tokenUrl -Body $body -Description ("refresh-token grant ($Resource)"))
}

function Get-IQRopcErrorMessage {
    <#
    .SYNOPSIS
        Maps a failed password-grant result to a clear operator message; returns @{Message; Fatal} (Fatal = do not try the module fallback).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Result)
    $codes = @($Result.AadCodes)
    $text = Get-IQAuthResultText -Result $Result
    $mfaCodes = @('50076', '50079', '53003', '65001', '50074', '530036')
    foreach ($c in $codes) {
        if ($mfaCodes -contains $c) {
            return @{ Fatal = $true; Message = ("MFA/Conditional Access blocks password authentication for this account (AADSTS{0}); use DeviceCode with a cached refresh token or exempt this account from MFA/CA. Details: {1}" -f $c, $text) }
        }
    }
    if ($codes -contains '50126') { return @{ Fatal = $true; Message = ('Invalid username or password (AADSTS50126). Check IMPACTIQ_USERNAME/IMPACTIQ_PASSWORD or -Credential. Details: ' + $text) } }
    if ($codes -contains '50034') { return @{ Fatal = $true; Message = ('The user account does not exist in this tenant (AADSTS50034); check the UPN and -TenantId. Details: ' + $text) } }
    if ($codes -contains '50053') { return @{ Fatal = $true; Message = ('The account is locked (AADSTS50053). Details: ' + $text) } }
    if ($codes -contains '50057') { return @{ Fatal = $true; Message = ('The account is disabled (AADSTS50057). Details: ' + $text) } }
    if ($codes -contains '50055') { return @{ Fatal = $true; Message = ('The password has expired (AADSTS50055). Details: ' + $text) } }
    if ($codes -contains '50056') { return @{ Fatal = $true; Message = ('The account has no password (passwordless/federated); password auth is impossible (AADSTS50056). Details: ' + $text) } }
    if ($codes -contains '50059' -or $codes -contains '90002') { return @{ Fatal = $true; Message = ('Tenant not found for this account; pass -TenantId <tenant guid> (AADSTS50059/90002). Details: ' + $text) } }
    if ($codes -contains '50008' -or $codes -contains '50155') { return @{ Fatal = $true; Message = ('Federated identity: ROPC is not supported for federated accounts; use DeviceCode or AzContext. Details: ' + $text) } }
    if ($codes -contains '7000218' -or $codes -contains '700016' -or $codes -contains '65002') { return @{ Fatal = $false; Message = ('The client application is not usable for password auth in this tenant; try another -ClientId or DeviceCode. Details: ' + $text) } }
    return @{ Fatal = $false; Message = ('Password (ROPC) sign-in failed: ' + $text) }
}

function Invoke-IQPasswordGrant {
    <#
    .SYNOPSIS
        Performs the OAuth 2.0 password (ROPC) grant for a resource; returns the Invoke-IQAuthRequest result (password never logged).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $true)][ValidateSet('PowerBI', 'Fabric')][string]$Resource,
        [Parameter(Mandatory = $false)][switch]$IncludeProfile
    )
    $auth = Get-IQAuthState -AllowUninitialized
    $password = ConvertFrom-IQSecureString -SecureString $Credential.Password
    try {
        $body = @{
            grant_type = 'password'
            client_id  = [string]$auth.ClientId
            scope      = (Get-IQAuthScope -Resource $Resource -IncludeProfile:$IncludeProfile)
            username   = $Credential.UserName
            password   = $password
        }
        $tokenUrl = Get-IQAuthEndpoint -Kind Token
        return (Invoke-IQAuthRequest -Uri $tokenUrl -Body $body -Description ("password grant ($Resource)"))
    }
    finally {
        $password = $null
        if ($null -ne $body) { $body['password'] = $null }
    }
}

# ---------------------------------------------------------------------------------------------------------------------
# Module-based providers (MicrosoftPowerBIMgmt, Az.Accounts)
# ---------------------------------------------------------------------------------------------------------------------

function Connect-IQPowerBIModule {
    <#
    .SYNOPSIS
        Signs in with MicrosoftPowerBIMgmt (Connect-PowerBIServiceAccount) using the monolith's 2-attempt retry loop and returns the access token.
    .DESCRIPTION
        Moved from Final PS Script.txt 455-503: -Environment omitted for Public, -WarningAction SilentlyContinue 3>$null,
        Disconnect-PowerBIServiceAccount + 2 s sleep between attempts. Additions: -Credential (ROPC fallback for the
        Credential mode) and -Tenant when a tenant GUID is configured and the cmdlet supports it.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false)][ValidateRange(1, 5)][int]$MaxAttempts = 2
    )
    $auth = Get-IQAuthState -AllowUninitialized
    if (-not (Import-IQAuthModule -Name 'MicrosoftPowerBIMgmt.Profile')) {
        throw 'The MicrosoftPowerBIMgmt module is not installed. Run: Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser'
    }
    $envName = [string]$script:IQ.Endpoints.MicrosoftPowerBIMgmtEnvironment
    $previousErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Stop'
        for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
            try {
                Write-IQLog -Level Info -Stage Auth -Message ("Connecting to Power BI using environment: {0}" -f $envName)
                $connectArgs = @{ ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
                if (-not [string]::IsNullOrWhiteSpace($envName) -and $envName -ne 'Public') { $connectArgs.Environment = $envName }
                if ($null -ne $Credential) { $connectArgs.Credential = $Credential }
                if ($auth -and (Test-IQAuthGuid -Value ([string]$auth.TenantId))) {
                    $cmd = Get-Command -Name Connect-PowerBIServiceAccount -ErrorAction SilentlyContinue
                    if ($cmd -and $cmd.Parameters.ContainsKey('Tenant')) { $connectArgs.Tenant = [string]$auth.TenantId }
                }
                Connect-PowerBIServiceAccount @connectArgs 3>$null | Out-Null

                $token = Get-PowerBIAccessToken -ErrorAction Stop -WarningAction SilentlyContinue 3>$null
                $plain = ConvertTo-IQPlainTokenValue -Value $token
                if (-not [string]::IsNullOrWhiteSpace($plain)) {
                    Write-IQLog -Level Success -Stage Auth -Message 'Connected to Power BI.'
                    return $plain
                }
                throw 'Power BI sign-in completed without returning an access token.'
            }
            catch {
                try { Disconnect-PowerBIServiceAccount -ErrorAction SilentlyContinue | Out-Null } catch { $null = $_.Exception }
                if ($attempt -ge $MaxAttempts) {
                    Write-IQLog -Level Error -Stage Auth -Message ('Power BI sign-in failed: ' + $_.Exception.Message)
                    throw
                }
                Write-IQLog -Level Warn -Stage Auth -Message ("Power BI sign-in attempt {0} failed: {1}. Retrying..." -f $attempt, $_.Exception.Message)
                Start-IQAuthSleep -Seconds 2
            }
        }
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
}

function Get-IQPowerBIModuleToken {
    <#
    .SYNOPSIS
        Current token from Get-PowerBIAccessToken (MSAL silent refresh); re-connects once on failure (monolith 560-575).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential)
    try {
        $token = Get-PowerBIAccessToken -ErrorAction Stop -WarningAction SilentlyContinue 3>$null
        $plain = ConvertTo-IQPlainTokenValue -Value $token
        if (-not [string]::IsNullOrWhiteSpace($plain)) { return $plain }
        throw 'Get-PowerBIAccessToken returned an empty token.'
    }
    catch {
        Write-IQLog -Level Warn -Stage Auth -Message ('Token fetch failed, re-connecting: ' + $_.Exception.Message)
        if (-not $script:IQ.Interactive -and $null -eq $Credential) {
            throw ('Cannot obtain a Power BI access token from the MicrosoftPowerBIMgmt session and a headless run cannot re-prompt: ' + $_.Exception.Message)
        }
        try {
            return (Connect-IQPowerBIModule -Credential $Credential -MaxAttempts 2)
        }
        catch {
            Write-IQLog -Level Error -Stage Auth -Message ('Re-connect also failed: ' + $_.Exception.Message)
            throw 'Cannot obtain Power BI access token.'
        }
    }
}

function Get-IQAzAccessTokenValue {
    <#
    .SYNOPSIS
        Get-AzAccessToken -ResourceUrl as plain text (handles the SecureString .Token of Az 14+); retries once after 5 s; throws on failure.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$ResourceUrl)
    $lastError = $null
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        try {
            $result = Get-AzAccessToken -ResourceUrl $ResourceUrl -ErrorAction Stop -WarningAction SilentlyContinue 3>$null
            $plain = ConvertTo-IQPlainTokenValue -Value $result
            if ([string]::IsNullOrWhiteSpace($plain)) { throw 'Get-AzAccessToken returned an empty token.' }
            return $plain
        }
        catch {
            $lastError = $_
            if ($attempt -lt 2) {
                Write-IQLog -Level Warn -Stage Auth -Message ("Get-AzAccessToken for {0} failed ({1}); retrying once in 5 s." -f $ResourceUrl, $_.Exception.Message)
                Start-IQAuthSleep -Seconds 5
            }
        }
    }
    throw $lastError
}

function Connect-IQAzForFabric {
    <#
    .SYNOPSIS
        Interactive-only best-effort Az.Accounts sign-in for Fabric tokens (monolith Get-FabricAccessToken 504-553); returns $true when a usable context exists.
    .DESCRIPTION
        Reuses an existing context whose environment matches; otherwise (interactive runs only) disables the
        LoginExperienceV2 picker and runs Connect-AzAccount -Environment <AzEnvironment> -Scope Process
        -SkipContextPopulation. Never blocks a headless run: returns $false instead of prompting.
    #>
    [CmdletBinding()]
    param()
    if (-not (Import-IQAuthModule -Name 'Az.Accounts')) { return $false }
    $azEnvironment = [string]$script:IQ.Endpoints.AzEnvironment
    $azContext = $null
    try { $azContext = Get-AzContext -ErrorAction SilentlyContinue } catch { $azContext = $null }
    if ($azContext -and $azContext.Environment -and $azContext.Environment.Name -eq $azEnvironment) { return $true }
    if (-not $script:IQ.Interactive) {
        Write-IQLog -Level Debug -Stage Auth -Message 'No matching Az context and the run is headless; Fabric tokens are unavailable.'
        return $false
    }
    try {
        Write-IQLog -Level Info -Stage Auth -Message 'Connecting to Microsoft Fabric (Az.Accounts)...'
        if (Get-Command -Name Update-AzConfig -ErrorAction SilentlyContinue) {
            Update-AzConfig -LoginExperienceV2 Off -Scope Process -ErrorAction SilentlyContinue | Out-Null
        }
        $connectArgs = @{ Environment = $azEnvironment; Scope = 'Process'; SkipContextPopulation = $true; ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
        $tenant = [string]$script:IQ.Auth.TenantId
        if (Test-IQAuthGuid -Value $tenant) { $connectArgs.Tenant = $tenant }
        Connect-AzAccount @connectArgs 3>$null | Out-Null
        return $true
    }
    catch {
        Write-IQLog -Level Warn -Stage Auth -Message ('Az.Accounts sign-in for Fabric failed (Fabric data will be skipped): ' + $_.Exception.Message)
        return $false
    }
}

# ---------------------------------------------------------------------------------------------------------------------
# Mode resolution and initial sign-in
# ---------------------------------------------------------------------------------------------------------------------

function Resolve-IQAuthMode {
    <#
    .SYNOPSIS
        Resolves "Auto" to a concrete mode (brief 4.1): Credential > AccessToken > DeviceCode (cache exists) > Interactive > DeviceCode.
    .DESCRIPTION
        AzContext is never auto-selected. -Interactive defaults to $script:IQ.Interactive (or Test-IQInteractive).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Auto', 'Interactive', 'DeviceCode', 'Credential', 'AzContext', 'AccessToken')][string]$Mode,
        [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false)][string]$TokenCachePath,
        [Parameter(Mandatory = $false)][AllowNull()][System.Nullable[bool]]$Interactive
    )
    if ($Mode -ne 'Auto') { return $Mode }
    if ($null -ne $Credential) { return 'Credential' }
    if (-not [string]::IsNullOrWhiteSpace($env:IMPACTIQ_USERNAME) -and -not [string]::IsNullOrEmpty($env:IMPACTIQ_PASSWORD)) { return 'Credential' }
    if (-not [string]::IsNullOrWhiteSpace($env:IMPACTIQ_PBI_TOKEN)) { return 'AccessToken' }
    if (-not [string]::IsNullOrWhiteSpace($TokenCachePath) -and (Test-Path -LiteralPath $TokenCachePath)) { return 'DeviceCode' }
    $isInteractive = $false
    if ($null -ne $Interactive) { $isInteractive = [bool]$Interactive }
    elseif ($script:IQ -and $script:IQ.ContainsKey('Interactive')) { $isInteractive = [bool]$script:IQ.Interactive }
    else { $isInteractive = [bool](Test-IQInteractive) }
    if ($isInteractive) { return 'Interactive' }
    return 'DeviceCode'
}

function Get-IQAuthCredential {
    <#
    .SYNOPSIS
        The PSCredential for Credential mode: -Credential, else IMPACTIQ_USERNAME/IMPACTIQ_PASSWORD (documented env-var path; never logged).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential)
    if ($null -ne $Credential) { return $Credential }
    $user = $env:IMPACTIQ_USERNAME
    $pass = $env:IMPACTIQ_PASSWORD
    if ([string]::IsNullOrWhiteSpace($user) -or [string]::IsNullOrEmpty($pass)) {
        throw 'Credential mode needs -Credential or both IMPACTIQ_USERNAME and IMPACTIQ_PASSWORD environment variables.'
    }
    # Plain-text env var is the documented pipeline secret path (mapped from a secret variable); it is converted immediately.
    $secure = ConvertTo-SecureString -String $pass -AsPlainText -Force
    return (New-Object System.Management.Automation.PSCredential($user, $secure))
}

function Initialize-IQAuth {
    <#
    .SYNOPSIS
        Resolves the auth mode (Auto -> brief 4.1), performs the initial Power BI sign-in, stores provider state in $script:IQ.Auth and returns the resolved mode.
    .PARAMETER Mode
        Auto | Interactive | DeviceCode | Credential | AzContext | AccessToken.
    .PARAMETER Environment
        Power BI cloud name (Public, USGov, USGovHigh, USGovMil, China, Germany); resolved with Get-IQEnvironmentSettings.
    .PARAMETER TokenCachePath
        Encrypted refresh-token cache (default <BaseFolder>\State\auth\token-cache.json or IMPACTIQ_TOKEN_CACHE_PATH).
    .PARAMETER TokenCacheKey
        AES passphrase for the cache (or IMPACTIQ_TOKEN_CACHE_KEY); without it DPAPI is used on Windows, nothing elsewhere.
    .PARAMETER DeviceCodeWebhookUrl
        Teams/Slack incoming webhook that receives the device-code message (or IMPACTIQ_DEVICECODE_WEBHOOK).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][ValidateSet('Auto', 'Interactive', 'DeviceCode', 'Credential', 'AzContext', 'AccessToken')][string]$Mode = 'Auto',
        [Parameter(Mandatory = $true)][string]$Environment,
        [Parameter(Mandatory = $false)][string]$TenantId = 'organizations',
        [Parameter(Mandatory = $false)][string]$ClientId,
        [Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential,
        [Parameter(Mandatory = $false)][string]$TokenCachePath,
        [Parameter(Mandatory = $false)][string]$TokenCacheKey,
        [Parameter(Mandatory = $false)][string]$DeviceCodeWebhookUrl
    )
    if ($null -eq $script:IQ) { throw 'ImpactIQ context is not initialised (Initialize-IQContext must run before Initialize-IQAuth).' }

    # --- environment / endpoints ---
    $endpoints = Get-IQEnvironmentSettings -Environment $Environment
    $script:IQ.Environment = $endpoints.Name
    $script:IQ.Endpoints = $endpoints

    # --- defaults from environment variables (documented in docs/Auth-Options.md) ---
    if ([string]::IsNullOrWhiteSpace($TenantId)) { $TenantId = 'organizations' }
    if ($TenantId -eq 'organizations' -and -not [string]::IsNullOrWhiteSpace($env:IMPACTIQ_TENANT_ID)) { $TenantId = $env:IMPACTIQ_TENANT_ID.Trim() }
    if ([string]::IsNullOrWhiteSpace($ClientId)) {
        if (-not [string]::IsNullOrWhiteSpace($env:IMPACTIQ_CLIENT_ID)) { $ClientId = $env:IMPACTIQ_CLIENT_ID.Trim() } else { $ClientId = $script:IQAuthDefaultClientId }
    }
    if ([string]::IsNullOrWhiteSpace($TokenCachePath)) {
        if (-not [string]::IsNullOrWhiteSpace($env:IMPACTIQ_TOKEN_CACHE_PATH)) { $TokenCachePath = $env:IMPACTIQ_TOKEN_CACHE_PATH }
        else { $TokenCachePath = Join-Path (Join-Path $script:IQ.StatePath 'auth') 'token-cache.json' }
    }
    elseif (-not [System.IO.Path]::IsPathRooted($TokenCachePath)) { $TokenCachePath = Join-Path $script:IQ.BaseFolder $TokenCachePath }
    if ([string]::IsNullOrEmpty($TokenCacheKey) -and -not [string]::IsNullOrEmpty($env:IMPACTIQ_TOKEN_CACHE_KEY)) { $TokenCacheKey = $env:IMPACTIQ_TOKEN_CACHE_KEY }
    if ([string]::IsNullOrWhiteSpace($DeviceCodeWebhookUrl) -and -not [string]::IsNullOrWhiteSpace($env:IMPACTIQ_DEVICECODE_WEBHOOK)) { $DeviceCodeWebhookUrl = $env:IMPACTIQ_DEVICECODE_WEBHOOK.Trim() }

    $resolved = Resolve-IQAuthMode -Mode $Mode -Credential $Credential -TokenCachePath $TokenCachePath
    Write-IQLog -Level Info -Stage Auth -Message ("Authentication mode: {0} (requested {1}); environment {2}; authority {3}; tenant {4}." -f $resolved, $Mode, $endpoints.Name, $endpoints.Authority, $TenantId)
    if ($ClientId -ne $script:IQAuthDefaultClientId) { Write-IQLog -Level Info -Stage Auth -Message ("Using public client id {0}." -f $ClientId) }

    $auth = @{
        Initialized          = $false
        Mode                 = $resolved
        RequestedMode        = $Mode
        Provider             = $null           # Http | Module | Az | Static
        Source               = $null           # e.g. 'cached refresh token', 'fresh device code', 'ROPC'
        Environment          = $endpoints.Name
        Authority            = $endpoints.Authority
        TenantId             = $TenantId
        TenantIdResolved     = $null
        ClientId             = $ClientId
        Account              = $null
        Tokens               = @{}
        RefreshToken         = $null
        Credential           = $null
        StaticTokens         = @{}
        TokenCachePath       = $TokenCachePath
        TokenCacheKey        = $TokenCacheKey
        DeviceCodeWebhookUrl = $DeviceCodeWebhookUrl
        FabricUnavailable    = $false
        FabricWarned         = $false
        StaticExpiryWarned   = $false
        CacheWarned          = $false
        Description          = $null
        SignedInUtc          = $null
    }
    $script:IQ.Auth = $auth

    switch ($resolved) {
        'Interactive' { Initialize-IQAuthInteractive }
        'DeviceCode' { Initialize-IQAuthDeviceCode }
        'Credential' { Initialize-IQAuthCredential -Credential $Credential }
        'AzContext' { Initialize-IQAuthAzContext }
        'AccessToken' { Initialize-IQAuthAccessToken }
    }

    $auth.Initialized = $true
    $auth.SignedInUtc = [datetime]::UtcNow
    $auth.Description = Get-IQAuthDescription
    $tenantText = ''
    if ($auth.TenantIdResolved) { $tenantText = ' (tenant ' + $auth.TenantIdResolved + ')' }
    Write-IQLog -Level Success -Stage Auth -Message ('Signed in: ' + $auth.Description + $tenantText)
    return $resolved
}

function Initialize-IQAuthInteractive {
    <#
    .SYNOPSIS
        Interactive mode: MicrosoftPowerBIMgmt sign-in (monolith Connect-PowerBI) + best-effort Az.Accounts for Fabric (private).
    #>
    [CmdletBinding()]
    param()
    $auth = Get-IQAuthState -AllowUninitialized
    if (-not $script:IQ.Interactive) {
        throw 'AuthMode Interactive needs an interactive desktop session (browser sign-in). Use -AuthMode DeviceCode, Credential, AzContext or AccessToken for headless runs.'
    }
    $auth.Provider = 'Module'
    $auth.Source = 'MicrosoftPowerBIMgmt'
    $token = Connect-IQPowerBIModule -MaxAttempts 2
    Set-IQAuthToken -Resource PowerBI -AccessToken $token -Source 'MicrosoftPowerBIMgmt' | Out-Null

    # Fabric token: best effort, never blocks the run (monolith 555; audit C2-02: no re-login on token failure).
    if (Connect-IQAzForFabric) {
        try {
            $fabric = Get-IQAzAccessTokenValue -ResourceUrl (Get-IQAuthResourceUrl -Resource Fabric)
            Set-IQAuthToken -Resource Fabric -AccessToken $fabric -Source 'Az.Accounts' | Out-Null
            Write-IQLog -Level Success -Stage Auth -Message 'Connected to Microsoft Fabric.'
        }
        catch {
            Write-IQLog -Level Warn -Stage Auth -Message ('Fabric token unavailable (Fabric collections will be skipped): ' + $_.Exception.Message)
        }
    }
    else {
        Write-IQLog -Level Warn -Stage Auth -Message 'Az.Accounts is not available or the Fabric sign-in was skipped; Fabric collections will be skipped.'
    }
}

function Initialize-IQAuthDeviceCode {
    <#
    .SYNOPSIS
        DeviceCode mode: silent refresh from the encrypted cache when possible, else a fresh device-code prompt; persists the newest refresh token (private).
    #>
    [CmdletBinding()]
    param()
    $auth = Get-IQAuthState -AllowUninitialized
    $auth.Provider = 'Http'
    $signedIn = $false

    $cache = Restore-IQTokenCache
    if ($null -ne $cache) {
        $auth.RefreshToken = [string]$cache.refreshToken
        if (-not [string]::IsNullOrWhiteSpace([string]$cache.account)) { $auth.Account = [string]$cache.account }
        Write-IQLog -Level Info -Stage Auth -Message ("Token cache found (saved {0}); attempting a silent refresh." -f $cache.savedUtc)
        $r = Invoke-IQRefreshTokenGrant -Resource PowerBI
        if ($r.Ok) {
            Set-IQAuthToken -Resource PowerBI -AccessToken (Get-IQAuthResponseAccessToken -Response $r.Response) -Response $r.Response -Source 'cached refresh token' -PersistCache | Out-Null
            $auth.Source = 'cached refresh token'
            $signedIn = $true
        }
        else {
            Write-IQLog -Level Warn -Stage Auth -Message ('The cached refresh token was rejected (' + (Get-IQAuthResultText -Result $r) + '); falling back to a fresh device-code sign-in.')
            $auth.RefreshToken = $null
            $auth.Account = $null
        }
    }
    if (-not $signedIn) {
        $response = Invoke-IQDeviceCodeFlow
        Set-IQAuthToken -Resource PowerBI -AccessToken (Get-IQAuthResponseAccessToken -Response $response) -Response $response -Source 'device code' -PersistCache | Out-Null
        $auth.Source = 'fresh device code'
    }
}

function Initialize-IQAuthCredential {
    <#
    .SYNOPSIS
        Credential mode: HTTP ROPC with AADSTS mapping, falling back to Connect-PowerBIServiceAccount -Credential when possible (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][System.Management.Automation.PSCredential]$Credential)
    $auth = Get-IQAuthState -AllowUninitialized
    $cred = Get-IQAuthCredential -Credential $Credential
    $auth.Credential = $cred
    $auth.Account = $cred.UserName
    Write-IQLog -Level Info -Stage Auth -Message ("Signing in with username/password (ROPC) as {0}." -f $cred.UserName)

    $r = Invoke-IQPasswordGrant -Credential $cred -Resource PowerBI -IncludeProfile
    if ($r.Ok) {
        $auth.Provider = 'Http'
        $auth.Source = 'ROPC'
        Set-IQAuthToken -Resource PowerBI -AccessToken (Get-IQAuthResponseAccessToken -Response $r.Response) -Response $r.Response -Source 'ROPC' | Out-Null
        return
    }
    $mapped = Get-IQRopcErrorMessage -Result $r
    if ($mapped.Fatal) { throw $mapped.Message }

    # Fallback: the MicrosoftPowerBIMgmt module (its own MSAL ROPC path; Windows only, no Fabric token).
    $moduleAvailable = $false
    if ($script:IQ.IsWindows) { $moduleAvailable = [bool](Get-Module -ListAvailable -Name 'MicrosoftPowerBIMgmt.Profile' -ErrorAction SilentlyContinue) }
    if (-not $moduleAvailable) { throw ($mapped.Message + ' (MicrosoftPowerBIMgmt fallback not available on this host.)') }
    Write-IQLog -Level Warn -Stage Auth -Message ($mapped.Message + ' Trying Connect-PowerBIServiceAccount -Credential as a fallback.')
    try {
        $token = Connect-IQPowerBIModule -Credential $cred -MaxAttempts 2
        $auth.Provider = 'Module'
        $auth.Source = 'MicrosoftPowerBIMgmt fallback'
        Set-IQAuthToken -Resource PowerBI -AccessToken $token -Source 'MicrosoftPowerBIMgmt -Credential' | Out-Null
    }
    catch {
        throw ($mapped.Message + ' Module fallback also failed: ' + $_.Exception.Message)
    }
}

function Initialize-IQAuthAzContext {
    <#
    .SYNOPSIS
        AzContext mode: requires an existing Az.Accounts context (never signs in); tokens via Get-AzAccessToken (private).
    #>
    [CmdletBinding()]
    param()
    $auth = Get-IQAuthState -AllowUninitialized
    $azEnvironment = [string]$script:IQ.Endpoints.AzEnvironment
    $setup = ("One-time setup as the account that runs ImpactIQ: Connect-AzAccount -UseDeviceAuthentication -Environment {0}{1}; Enable-AzContextAutosave -Scope CurrentUser" -f $azEnvironment, $(if (Test-IQAuthGuid -Value ([string]$auth.TenantId)) { ' -Tenant ' + $auth.TenantId } else { '' }))
    if (-not (Import-IQAuthModule -Name 'Az.Accounts')) {
        throw ('AuthMode AzContext needs the Az.Accounts module (Install-Module Az.Accounts -Scope CurrentUser). ' + $setup)
    }
    $ctx = $null
    try { $ctx = Get-AzContext -ErrorAction Stop } catch { $ctx = $null }
    if ($null -eq $ctx -or $null -eq $ctx.Account) {
        throw ('AuthMode AzContext: no Az context is available for this user. ' + $setup)
    }
    if ($ctx.Environment -and $ctx.Environment.Name -ne $azEnvironment) {
        Write-IQLog -Level Warn -Stage Auth -Message ("The Az context environment is {0} but {1} expects {2}; token requests may fail." -f $ctx.Environment.Name, $script:IQ.Environment, $azEnvironment)
    }
    $auth.Provider = 'Az'
    $auth.Source = 'Az.Accounts context'
    if ($ctx.Account -and $ctx.Account.Id) { $auth.Account = [string]$ctx.Account.Id }
    try {
        $token = Get-IQAzAccessTokenValue -ResourceUrl (Get-IQAuthResourceUrl -Resource PowerBI)
    }
    catch {
        throw ('AuthMode AzContext: Get-AzAccessToken for the Power BI resource failed (' + $_.Exception.Message + '). The cached sign-in may have expired or been revoked. ' + $setup)
    }
    Set-IQAuthToken -Resource PowerBI -AccessToken $token -Source 'Az.Accounts' | Out-Null
}

function Initialize-IQAuthAccessToken {
    <#
    .SYNOPSIS
        AccessToken mode: IMPACTIQ_PBI_TOKEN (+ optional IMPACTIQ_FABRIC_TOKEN); no refresh possible, expiry logged at Warn (private).
    #>
    [CmdletBinding()]
    param()
    $auth = Get-IQAuthState -AllowUninitialized
    $pbi = ConvertTo-IQPlainTokenValue -Value ([string]$env:IMPACTIQ_PBI_TOKEN)
    if ([string]::IsNullOrWhiteSpace($pbi)) { throw 'AuthMode AccessToken needs the IMPACTIQ_PBI_TOKEN environment variable (a bearer token for the Power BI resource).' }
    $auth.Provider = 'Static'
    $auth.Source = 'IMPACTIQ_PBI_TOKEN'
    $auth.StaticTokens = @{ PowerBI = $pbi }
    $fabric = ConvertTo-IQPlainTokenValue -Value ([string]$env:IMPACTIQ_FABRIC_TOKEN)
    if (-not [string]::IsNullOrWhiteSpace($fabric)) { $auth.StaticTokens['Fabric'] = $fabric }

    $expiry = Get-IQJwtExpiry -Token $pbi
    if ($expiry -ne [datetime]::MaxValue) {
        if ($expiry -le [datetime]::UtcNow) { throw ("IMPACTIQ_PBI_TOKEN already expired at {0:u}." -f $expiry) }
        Write-IQLog -Level Warn -Stage Auth -Message ("AccessToken mode: the token cannot be refreshed and expires at {0:u} (in {1} min). The run fails if it is still going then." -f $expiry, [int]($expiry - [datetime]::UtcNow).TotalMinutes)
    }
    else {
        Write-IQLog -Level Warn -Stage Auth -Message 'AccessToken mode: the token expiry could not be read; it cannot be refreshed.'
    }
    Set-IQAuthToken -Resource PowerBI -AccessToken $pbi -Source 'IMPACTIQ_PBI_TOKEN' | Out-Null
    if ($auth.StaticTokens.ContainsKey('Fabric')) { Set-IQAuthToken -Resource Fabric -AccessToken $auth.StaticTokens['Fabric'] -Source 'IMPACTIQ_FABRIC_TOKEN' | Out-Null }
}

# ---------------------------------------------------------------------------------------------------------------------
# Get-IQToken and per-mode minting
# ---------------------------------------------------------------------------------------------------------------------

function Set-IQAuthFabricUnavailable {
    <#
    .SYNOPSIS
        Marks Fabric tokens as unavailable (logged once at Debug) and returns $null (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Reason, [Parameter(Mandatory = $false)][switch]$Permanent)
    $auth = Get-IQAuthState -AllowUninitialized
    if ($Permanent) { $auth.FabricUnavailable = $true }
    if (-not $auth.FabricWarned) {
        $auth.FabricWarned = $true
        Write-IQLog -Level Debug -Stage Auth -Message ('Fabric tokens are not available from this provider; Fabric API calls will be skipped. ' + $Reason)
    }
    return $null
}

function Update-IQAuthToken {
    <#
    .SYNOPSIS
        Mints or refreshes the token for a resource according to the resolved mode; throws for PowerBI, returns $null for Fabric when impossible (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateSet('PowerBI', 'Fabric')][string]$Resource)
    $auth = Get-IQAuthState
    if ($Resource -eq 'Fabric' -and $auth.FabricUnavailable) { return (Set-IQAuthFabricUnavailable) }

    switch ($auth.Mode) {
        'Interactive' {
            if ($Resource -eq 'PowerBI') {
                $t = Get-IQPowerBIModuleToken
                return (Set-IQAuthToken -Resource PowerBI -AccessToken $t -Source 'MicrosoftPowerBIMgmt')
            }
            if (-not (Connect-IQAzForFabric)) { return (Set-IQAuthFabricUnavailable -Reason 'No Az.Accounts context.' -Permanent) }
            try {
                $t = Get-IQAzAccessTokenValue -ResourceUrl (Get-IQAuthResourceUrl -Resource Fabric)
                return (Set-IQAuthToken -Resource Fabric -AccessToken $t -Source 'Az.Accounts')
            }
            catch {
                Write-IQLog -Level Warn -Stage Auth -Message ('Fabric token refresh via Az.Accounts failed: ' + $_.Exception.Message)
                return (Set-IQAuthFabricUnavailable -Reason $_.Exception.Message)
            }
        }
        'DeviceCode' {
            $r = Invoke-IQRefreshTokenGrant -Resource $Resource
            if ($r.Ok) { return (Set-IQAuthToken -Resource $Resource -AccessToken (Get-IQAuthResponseAccessToken -Response $r.Response) -Response $r.Response -Source 'refresh token' -PersistCache) }
            $text = Get-IQAuthResultText -Result $r
            if ($Resource -eq 'Fabric') {
                Write-IQLog -Level Debug -Stage Auth -Message ('Fabric refresh-token grant failed: ' + $text)
                return (Set-IQAuthFabricUnavailable -Reason $text -Permanent:(Test-IQAuthPermanentFailure -Result $r))
            }
            if ($r.Error -eq 'invalid_grant' -or $r.Error -eq 'interaction_required') {
                # The refresh token was revoked/expired (password change, CA policy, 90-day inactivity): the only way
                # forward is a new device-code sign-in; the webhook/log tells a human. Throws if nobody completes it.
                Write-IQLog -Level Warn -Stage Auth -Message ('The refresh token was rejected (' + $text + '); starting a new device-code sign-in.')
                $auth.RefreshToken = $null
                $response = Invoke-IQDeviceCodeFlow
                $auth.Source = 'fresh device code'
                $auth.Description = $null
                $t = Set-IQAuthToken -Resource PowerBI -AccessToken (Get-IQAuthResponseAccessToken -Response $response) -Response $response -Source 'device code' -PersistCache
                $auth.Description = Get-IQAuthDescription
                return $t
            }
            throw ('Could not refresh the Power BI access token: ' + $text)
        }
        'Credential' {
            if ($auth.Provider -eq 'Module') {
                if ($Resource -eq 'Fabric') { return (Set-IQAuthFabricUnavailable -Reason 'MicrosoftPowerBIMgmt cannot mint Fabric tokens.' -Permanent) }
                $t = Get-IQPowerBIModuleToken -Credential $auth.Credential
                return (Set-IQAuthToken -Resource PowerBI -AccessToken $t -Source 'MicrosoftPowerBIMgmt -Credential')
            }
            # HTTP provider: prefer the refresh token from the first ROPC response, else repeat the password grant.
            $r = $null
            if (-not [string]::IsNullOrWhiteSpace([string]$auth.RefreshToken)) {
                $r = Invoke-IQRefreshTokenGrant -Resource $Resource
                if (-not $r.Ok) { Write-IQLog -Level Debug -Stage Auth -Message ("Refresh-token grant failed ({0}); repeating the password grant." -f (Get-IQAuthResultText -Result $r)) }
            }
            if ($null -eq $r -or -not $r.Ok) { $r = Invoke-IQPasswordGrant -Credential $auth.Credential -Resource $Resource }
            if ($r.Ok) { return (Set-IQAuthToken -Resource $Resource -AccessToken (Get-IQAuthResponseAccessToken -Response $r.Response) -Response $r.Response -Source 'ROPC') }
            if ($Resource -eq 'Fabric') {
                $text = Get-IQAuthResultText -Result $r
                return (Set-IQAuthFabricUnavailable -Reason $text -Permanent:(Test-IQAuthPermanentFailure -Result $r))
            }
            $mapped = Get-IQRopcErrorMessage -Result $r
            throw ('Could not refresh the Power BI access token: ' + $mapped.Message)
        }
        'AzContext' {
            try {
                $t = Get-IQAzAccessTokenValue -ResourceUrl (Get-IQAuthResourceUrl -Resource $Resource)
                return (Set-IQAuthToken -Resource $Resource -AccessToken $t -Source 'Az.Accounts')
            }
            catch {
                if ($Resource -eq 'Fabric') {
                    Write-IQLog -Level Debug -Stage Auth -Message ('Fabric token via Az.Accounts failed: ' + $_.Exception.Message)
                    return (Set-IQAuthFabricUnavailable -Reason $_.Exception.Message -Permanent)
                }
                throw ('Could not obtain a Power BI token from the Az context (' + $_.Exception.Message + '). Re-run the one-time setup: Connect-AzAccount -UseDeviceAuthentication; Enable-AzContextAutosave -Scope CurrentUser')
            }
        }
        'AccessToken' {
            if (-not $auth.StaticTokens.ContainsKey($Resource)) {
                if ($Resource -eq 'Fabric') { return (Set-IQAuthFabricUnavailable -Reason 'IMPACTIQ_FABRIC_TOKEN is not set.' -Permanent) }
                throw 'IMPACTIQ_PBI_TOKEN is not available.'
            }
            $t = [string]$auth.StaticTokens[$Resource]
            $expiry = Get-IQJwtExpiry -Token $t
            if ($expiry -le [datetime]::UtcNow) {
                if ($Resource -eq 'Fabric') { return (Set-IQAuthFabricUnavailable -Reason ('IMPACTIQ_FABRIC_TOKEN expired at ' + $expiry.ToString('u')) -Permanent) }
                throw ("The IMPACTIQ_PBI_TOKEN access token expired at {0:u} and AccessToken mode cannot refresh it. Provide a fresh token (or use DeviceCode/Credential/AzContext) and resume the run." -f $expiry)
            }
            if ($expiry -ne [datetime]::MaxValue -and $expiry -le [datetime]::UtcNow.AddMinutes($script:IQAuthRefreshSkewMinutes) -and -not $auth.StaticExpiryWarned) {
                $auth.StaticExpiryWarned = $true
                Write-IQLog -Level Warn -Stage Auth -Message ("The static {0} token expires at {1:u}; the run will fail after that." -f $Resource, $expiry)
            }
            return (Set-IQAuthToken -Resource $Resource -AccessToken $t -Source ('IMPACTIQ_' + $(if ($Resource -eq 'Fabric') { 'FABRIC' } else { 'PBI' }) + '_TOKEN'))
        }
        default { throw ("Unknown authentication mode '{0}'." -f $auth.Mode) }
    }
}

function Get-IQToken {
    <#
    .SYNOPSIS
        Bearer token string (without "Bearer ") for PowerBI or Fabric; refreshes silently when < 5 minutes remain or -Force is set.
    .DESCRIPTION
        Throws if a Power BI token cannot be obtained. For Fabric returns $null (logged once at Debug) when the provider
        cannot mint Fabric tokens so Fabric calls degrade gracefully. -Force is used by the Http module after a 401.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('PowerBI', 'Fabric')][string]$Resource,
        [Parameter(Mandatory = $false)][switch]$Force
    )
    $auth = Get-IQAuthState
    if (-not $Force) {
        $cached = Get-IQAuthCachedToken -Resource $Resource
        if ($null -ne $cached) { return $cached }
    }
    else {
        # A forced refresh (401) means the token we hold is bad; make sure the static-token path re-validates too.
        if ($auth.Tokens -is [hashtable] -and $auth.Tokens.ContainsKey($Resource)) { $auth.Tokens.Remove($Resource) }
        if ($Resource -eq 'Fabric' -and $auth.FabricUnavailable) { return (Set-IQAuthFabricUnavailable) }
    }
    $token = Update-IQAuthToken -Resource $Resource
    if ($Resource -eq 'PowerBI' -and [string]::IsNullOrWhiteSpace($token)) { throw 'Could not obtain a Power BI access token.' }
    return $token
}

function Get-IQAuthDescription {
    <#
    .SYNOPSIS
        One-line description for the manifest, e.g. "DeviceCode (cached refresh token) as user@contoso.gov".
    #>
    [CmdletBinding()]
    param()
    $auth = $null
    if ($script:IQ -and $script:IQ.Auth -is [hashtable]) { $auth = $script:IQ.Auth }
    if ($null -eq $auth -or [string]::IsNullOrWhiteSpace([string]$auth.Mode)) { return 'not signed in' }
    if ($auth.ContainsKey('Description') -and -not [string]::IsNullOrWhiteSpace([string]$auth.Description)) { return [string]$auth.Description }
    $text = [string]$auth.Mode
    $source = $null
    if ($auth.ContainsKey('Source')) { $source = [string]$auth.Source }
    if (-not [string]::IsNullOrWhiteSpace($source)) { $text = $text + ' (' + $source + ')' }
    $account = $null
    if ($auth.ContainsKey('Account')) { $account = [string]$auth.Account }
    if ([string]::IsNullOrWhiteSpace($account) -and $auth.Tokens -is [hashtable] -and $auth.Tokens.ContainsKey('PowerBI')) {
        $account = Get-IQJwtAccount -Token ([string]$auth.Tokens['PowerBI'].AccessToken)
    }
    if (-not [string]::IsNullOrWhiteSpace($account)) { $text = $text + ' as ' + $account }
    return $text
}

# ---------------------------------------------------------------------------------------------------------------------
# Token cache (brief 4.3)
# ---------------------------------------------------------------------------------------------------------------------

function Get-IQTokenCacheKey {
    <#
    .SYNOPSIS
        SHA-256 of the UTF-8 cache key string = the AES-256 key (private).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$KeyString)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($KeyString)) }
    finally { $sha.Dispose() }
}

function Protect-IQTokenCacheData {
    <#
    .SYNOPSIS
        Encrypts bytes: AES-256-CBC (random IV prepended) when a key string is given, else DPAPI CurrentUser on Windows; returns @{Format; Data} (base64).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$KeyString
    )
    if (-not [string]::IsNullOrEmpty($KeyString)) {
        $aes = New-Object System.Security.Cryptography.AesCryptoServiceProvider
        try {
            $aes.KeySize = 256
            $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
            $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
            $aes.Key = Get-IQTokenCacheKey -KeyString $KeyString
            $aes.GenerateIV()
            $encryptor = $aes.CreateEncryptor()
            try { $cipher = $encryptor.TransformFinalBlock($Bytes, 0, $Bytes.Length) } finally { $encryptor.Dispose() }
            $out = New-Object byte[] ($aes.IV.Length + $cipher.Length)
            [Array]::Copy($aes.IV, 0, $out, 0, $aes.IV.Length)
            [Array]::Copy($cipher, 0, $out, $aes.IV.Length, $cipher.Length)
            return @{ Format = 'aes256'; Data = [Convert]::ToBase64String($out) }
        }
        finally { $aes.Dispose() }
    }
    if ($script:IQ -and $script:IQ.IsWindows) {
        try { Add-Type -AssemblyName System.Security -ErrorAction Stop } catch { $null = $_.Exception }
        $type = 'System.Security.Cryptography.ProtectedData' -as [type]
        if ($null -eq $type) { throw 'DPAPI (System.Security.Cryptography.ProtectedData) is not available on this host.' }
        $protected = $type::Protect($Bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return @{ Format = 'dpapi-user'; Data = [Convert]::ToBase64String($protected) }
    }
    throw 'No token-cache key (-TokenCacheKey / IMPACTIQ_TOKEN_CACHE_KEY) and DPAPI is only available on Windows.'
}

function Unprotect-IQTokenCacheData {
    <#
    .SYNOPSIS
        Decrypts a cache payload written by Protect-IQTokenCacheData; throws on wrong key / other user / unsupported format.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Format,
        [Parameter(Mandatory = $true)][string]$Data,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$KeyString
    )
    $raw = [Convert]::FromBase64String($Data)
    switch ($Format) {
        'aes256' {
            if ([string]::IsNullOrEmpty($KeyString)) { throw 'The token cache is AES-encrypted but no -TokenCacheKey / IMPACTIQ_TOKEN_CACHE_KEY was provided.' }
            if ($raw.Length -le 16) { throw 'The token cache payload is too short.' }
            $aes = New-Object System.Security.Cryptography.AesCryptoServiceProvider
            try {
                $aes.KeySize = 256
                $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
                $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
                $aes.Key = Get-IQTokenCacheKey -KeyString $KeyString
                $iv = New-Object byte[] 16
                [Array]::Copy($raw, 0, $iv, 0, 16)
                $aes.IV = $iv
                $decryptor = $aes.CreateDecryptor()
                try { return $decryptor.TransformFinalBlock($raw, 16, $raw.Length - 16) } finally { $decryptor.Dispose() }
            }
            finally { $aes.Dispose() }
        }
        'dpapi-user' {
            if (-not ($script:IQ -and $script:IQ.IsWindows)) { throw 'The token cache is DPAPI-protected and can only be read on the Windows machine/user that wrote it.' }
            try { Add-Type -AssemblyName System.Security -ErrorAction Stop } catch { $null = $_.Exception }
            $type = 'System.Security.Cryptography.ProtectedData' -as [type]
            if ($null -eq $type) { throw 'DPAPI (System.Security.Cryptography.ProtectedData) is not available on this host.' }
            return $type::Unprotect($raw, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        }
        default { throw ("Unsupported token cache format '{0}'." -f $Format) }
    }
}

function Save-IQTokenCache {
    <#
    .SYNOPSIS
        Persists the current refresh token, encrypted (AES-256 with -TokenCacheKey, else DPAPI on Windows), to the token cache path; returns $true when written.
    .DESCRIPTION
        Payload: { schemaVersion, authority, tenantId, clientId, environment, refreshToken, account, savedUtc } wrapped as
        { "format": "aes256" | "dpapi-user", "data": "<base64>" }. Never written into the run folder. Logs a Warn (once)
        and returns $false when nothing can be persisted; never throws.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Path)
    $auth = Get-IQAuthState -AllowUninitialized
    if ($null -eq $auth) { return $false }
    if ([string]::IsNullOrWhiteSpace([string]$auth.RefreshToken)) {
        Write-IQLog -Level Debug -Stage Auth -Message 'No refresh token to persist.'
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($Path)) { $Path = [string]$auth.TokenCachePath }
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $keyString = $null
    if ($auth.ContainsKey('TokenCacheKey')) { $keyString = [string]$auth.TokenCacheKey }
    if ([string]::IsNullOrEmpty($keyString) -and -not ($script:IQ -and $script:IQ.IsWindows)) {
        if (-not $auth.CacheWarned) {
            $auth.CacheWarned = $true
            Write-IQLog -Level Warn -Stage Auth -Message 'Refresh token not persisted: set -TokenCacheKey / IMPACTIQ_TOKEN_CACHE_KEY (DPAPI is only available on Windows). The next run will prompt for a device code again.'
        }
        return $false
    }
    try {
        $payload = [ordered]@{
            schemaVersion = 1
            authority     = [string]$auth.Authority
            tenantId      = [string]$auth.TenantId
            clientId      = [string]$auth.ClientId
            environment   = [string]$auth.Environment
            refreshToken  = [string]$auth.RefreshToken
            account       = [string]$auth.Account
            savedUtc      = [datetime]::UtcNow.ToString('o')
        }
        $json = ConvertTo-Json -InputObject $payload -Depth 20 -Compress
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
        $protected = Protect-IQTokenCacheData -Bytes $bytes -KeyString $keyString
        $file = [ordered]@{ format = $protected.Format; data = $protected.Data }
        ConvertTo-IQJsonFile -Object $file -Path $Path
        Write-IQLog -Level Debug -Stage Auth -Message ("Token cache saved ({0}) to {1}." -f $protected.Format, $Path)
        return $true
    }
    catch {
        Write-IQLog -Level Warn -Stage Auth -Message ('Could not save the token cache: ' + $_.Exception.Message)
        return $false
    }
}

function Restore-IQTokenCache {
    <#
    .SYNOPSIS
        Reads and decrypts the token cache; returns a hashtable (refreshToken, account, ...) or $null on any failure or mismatch (other environment/client/authority).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Path)
    $auth = Get-IQAuthState -AllowUninitialized
    if ([string]::IsNullOrWhiteSpace($Path) -and $auth) { $Path = [string]$auth.TokenCachePath }
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        $file = ConvertFrom-IQJsonFile -Path $Path
        if ($null -eq $file -or -not $file.PSObject.Properties['format'] -or -not $file.PSObject.Properties['data']) {
            Write-IQLog -Level Warn -Stage Auth -Message ('Token cache at ' + $Path + ' has an unexpected shape; ignoring it.')
            return $null
        }
        $keyString = $null
        if ($auth -and $auth.ContainsKey('TokenCacheKey')) { $keyString = [string]$auth.TokenCacheKey }
        $bytes = Unprotect-IQTokenCacheData -Format ([string]$file.format) -Data ([string]$file.data) -KeyString $keyString
        $json = [System.Text.Encoding]::UTF8.GetString($bytes)
        $payload = ConvertFrom-Json -InputObject $json -ErrorAction Stop
        $cache = @{}
        foreach ($p in $payload.PSObject.Properties) { $cache[$p.Name] = $p.Value }
        if (-not $cache.ContainsKey('refreshToken') -or [string]::IsNullOrWhiteSpace([string]$cache['refreshToken'])) {
            Write-IQLog -Level Warn -Stage Auth -Message 'Token cache contains no refresh token; ignoring it.'
            return $null
        }
        if ([int]$cache['schemaVersion'] -ne 1) {
            Write-IQLog -Level Warn -Stage Auth -Message ("Token cache schema version {0} is not supported; ignoring it." -f $cache['schemaVersion'])
            return $null
        }
        if ($auth) {
            $mismatch = New-Object System.Collections.Generic.List[string]
            if ([string]$cache['authority'] -ne [string]$auth.Authority) { $mismatch.Add('authority') }
            if ([string]$cache['clientId'] -ne [string]$auth.ClientId) { $mismatch.Add('clientId') }
            if ([string]$cache['environment'] -ne [string]$auth.Environment) { $mismatch.Add('environment') }
            if ([string]$cache['tenantId'] -ne [string]$auth.TenantId) {
                # Refresh tokens are not bound to a tenant id (organizations vs. a GUID); keep using the cache.
                Write-IQLog -Level Debug -Stage Auth -Message ("Token cache was saved with tenant '{0}' (current '{1}'); using it anyway." -f $cache['tenantId'], $auth.TenantId)
            }
            if ($mismatch.Count -gt 0) {
                Write-IQLog -Level Warn -Stage Auth -Message ('Token cache was created for a different ' + [string]::Join('/', $mismatch.ToArray()) + '; ignoring it.')
                return $null
            }
        }
        return $cache
    }
    catch {
        Write-IQLog -Level Warn -Stage Auth -Message ('Token cache could not be read (wrong key, other user/machine, or corrupt): ' + $_.Exception.Message)
        return $null
    }
}
