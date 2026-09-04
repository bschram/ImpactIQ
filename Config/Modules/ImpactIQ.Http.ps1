# ImpactIQ.Http.ps1 - the single HTTP wrapper for every Power BI / Fabric REST call.
#
# Contract: brief section 2.3. Everything that talks to the service goes through Invoke-IQApi
# (paging, retry/backoff/429 handling, 401 refresh-once, 403/404 -> $null, -Raw, -OutFile),
# Invoke-IQFabricLro (202 long-running operations) or Invoke-IQDownload.
#
# Windows PowerShell 5.1 and PowerShell 7 compatible. Loaded by dot-sourcing from ImpactIQ.ps1, so
# $script:IQ is the shared context created by Initialize-IQContext (ImpactIQ.Common.ps1).
#
# Cross-module functions used (brief section 2): Write-IQLog, Get-IQToken.
# Private helpers are prefixed Start-IQHttp* / Get-IQHttp* / Invoke-IQHttp* and are not part of the contract.

function Start-IQHttpSleep {
    <#
    .SYNOPSIS
    Sleeps for the given number of seconds (private; mocked by tests so retry waits are instant).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [double]$Seconds
    )
    if ($Seconds -le 0) { return }
    Start-Sleep -Milliseconds ([int][math]::Ceiling($Seconds * 1000))
}

function Initialize-IQHttpDefault {
    <#
    .SYNOPSIS
    One-time process defaults for HTTP: TLS 1.2 OR-ed into the protocol set and system-proxy credentials (private).
    #>
    [CmdletBinding()]
    param()
    if ($script:IQ -and $script:IQ.ContainsKey('HttpDefaultsApplied') -and $script:IQ.HttpDefaultsApplied) { return }
    try {
        # Audit C1-06: OR Tls12 into the existing set instead of replacing it (Windows PowerShell 5.1 defaults to SSL3/TLS1.0).
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }
    catch {
        Write-IQLog -Level Debug -Stage Http -Message ("Could not adjust SecurityProtocol: " + $_.Exception.Message)
    }
    try {
        # Audit C1-17: agents behind an authenticated corporate proxy need default credentials on the system proxy.
        $proxy = [System.Net.WebRequest]::GetSystemWebProxy()
        if ($proxy) {
            $proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
            [System.Net.WebRequest]::DefaultWebProxy = $proxy
        }
    }
    catch {
        Write-IQLog -Level Debug -Stage Http -Message ("Could not configure the system web proxy: " + $_.Exception.Message)
    }
    if ($script:IQ) { $script:IQ.HttpDefaultsApplied = $true }
}

function Get-IQHttpHeaderValue {
    <#
    .SYNOPSIS
    Returns the first value of a response header by case-insensitive name from a 5.1 or 7 header dictionary (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Headers,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )
    if ($null -eq $Headers) { return $null }
    try {
        # HttpResponseHeaders (PowerShell 7 HttpResponseException) expose TryGetValues.
        if ($Headers.PSObject.Methods['TryGetValues']) {
            $values = $null
            if ($Headers.TryGetValues($Name, [ref]$values) -and $values) {
                return [string](@($values)[0])
            }
            return $null
        }
        # WebHeaderCollection (5.1 HttpWebResponse) and Dictionary[string,string]/[string,string[]] (Invoke-WebRequest).
        $keys = @()
        if ($Headers.PSObject.Properties['AllKeys']) { $keys = @($Headers.AllKeys) }
        elseif ($Headers.PSObject.Properties['Keys']) { $keys = @($Headers.Keys) }
        foreach ($key in $keys) {
            if ([string]$key -ieq $Name) {
                $value = $Headers[$key]
                if ($null -eq $value) { return $null }
                if ($value -is [string]) { return $value }
                return [string](@($value)[0])
            }
        }
    }
    catch {
        return $null
    }
    return $null
}

function ConvertTo-IQRetryAfterDelay {
    <#
    .SYNOPSIS
    Parses a Retry-After header value (delta-seconds or HTTP-date) into whole seconds, or $null when unparsable (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $text = $Value.Trim()
    $seconds = 0
    if ([int]::TryParse($text, [ref]$seconds)) {
        if ($seconds -lt 0) { return 0 }
        return $seconds
    }
    $when = [datetime]::MinValue
    $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
    if ([datetime]::TryParse($text, [System.Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$when)) {
        $delta = ($when - [datetime]::UtcNow).TotalSeconds
        if ($delta -lt 0) { return 0 }
        return [int][math]::Ceiling($delta)
    }
    return $null
}

function Get-IQHttpErrorInfo {
    <#
    .SYNOPSIS
    Extracts StatusCode, Retry-After, response body and message from an Invoke-WebRequest/Invoke-RestMethod error on 5.1 and 7.
    .DESCRIPTION
    Handles both Windows PowerShell 5.1 (System.Net.WebException, body read from
    $_.Exception.Response.GetResponseStream()) and PowerShell 7
    (Microsoft.PowerShell.Commands.HttpResponseException, body in $_.ErrorDetails.Message).
    Returns @{ StatusCode = <int or $null>; RetryAfter = <int seconds or $null>; Body = <string>; Message = <string>;
    Transient = <bool: network-level failure worth retrying>; WebStatus = <string> }.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    $info = @{ StatusCode = $null; RetryAfter = $null; Body = ''; Message = ''; Transient = $false; WebStatus = $null }
    $ex = $ErrorRecord.Exception
    if ($null -eq $ex) {
        $info.Message = [string]$ErrorRecord
        return $info
    }
    $info.Message = $ex.Message

    # ErrorDetails.Message carries the response body on PowerShell 7 (and usually on 5.1 as well).
    try {
        if ($ErrorRecord.ErrorDetails -and -not [string]::IsNullOrEmpty($ErrorRecord.ErrorDetails.Message)) {
            $info.Body = [string]$ErrorRecord.ErrorDetails.Message
        }
    }
    catch { $info.Body = '' }

    $response = $null
    try { if ($ex.PSObject.Properties['Response']) { $response = $ex.Response } } catch { $response = $null }

    $typeName = $ex.GetType().FullName
    try { if ($ex.PSObject.Properties['Status'] -and $typeName -eq 'System.Net.WebException') { $info.WebStatus = [string]$ex.Status } } catch { $info.WebStatus = $null }

    if ($null -ne $response) {
        # Status code: HttpWebResponse.StatusCode (5.1) and HttpResponseMessage.StatusCode (7) both cast to int.
        try { if ($response.PSObject.Properties['StatusCode'] -and $null -ne $response.StatusCode) { $info.StatusCode = [int]$response.StatusCode } } catch { $info.StatusCode = $null }

        # Retry-After header.
        $retryAfterText = $null
        try {
            if ($response.PSObject.Properties['Headers'] -and $null -ne $response.Headers) {
                $headers = $response.Headers
                if ($headers.PSObject.Properties['RetryAfter'] -and $null -ne $headers.RetryAfter) {
                    # PowerShell 7: RetryConditionHeaderValue with Delta (TimeSpan) or Date (DateTimeOffset).
                    $ra = $headers.RetryAfter
                    if ($null -ne $ra.Delta) { $retryAfterText = [string][int][math]::Ceiling($ra.Delta.TotalSeconds) }
                    elseif ($null -ne $ra.Date) { $retryAfterText = $ra.Date.UtcDateTime.ToString('r') }
                }
                if ($null -eq $retryAfterText) { $retryAfterText = Get-IQHttpHeaderValue -Headers $headers -Name 'Retry-After' }
            }
        }
        catch { $retryAfterText = $null }
        if ($null -ne $retryAfterText) { $info.RetryAfter = ConvertTo-IQRetryAfterDelay -Value $retryAfterText }

        # 5.1: read the body from the response stream when ErrorDetails did not carry it.
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
        # 7: HttpResponseMessage.Content can be read when ErrorDetails is empty (rare).
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
        # No HTTP response: decide whether this is a transient network failure.
        $transientWebStatuses = @('Timeout', 'ConnectFailure', 'NameResolutionFailure', 'ReceiveFailure', 'SendFailure',
            'ConnectionClosed', 'KeepAliveFailure', 'PipelineFailure', 'ProxyNameResolutionFailure', 'RequestCanceled', 'UnknownError')
        if ($info.WebStatus -and $transientWebStatuses -contains $info.WebStatus) { $info.Transient = $true }
        $walk = $ex
        $depth = 0
        while ($null -ne $walk -and $depth -lt 6 -and -not $info.Transient) {
            $name = $walk.GetType().FullName
            if ($name -match 'HttpRequestException|TaskCanceledException|OperationCanceledException|SocketException|System\.IO\.IOException|WebException|TimeoutException') {
                $info.Transient = $true
            }
            elseif ($walk.Message -match 'operation has timed out|operation was canceled|connection was forcibly closed|Unable to connect|No such host|actively refused|Resource temporarily unavailable') {
                $info.Transient = $true
            }
            $walk = $walk.InnerException
            $depth++
        }
    }
    if ($info.StatusCode -and [string]::IsNullOrEmpty($info.Message)) { $info.Message = "HTTP $($info.StatusCode)" }
    return $info
}

function Get-IQApiUrl {
    <#
    .SYNOPSIS
    Builds the absolute request URL for a relative Power BI / Fabric path plus an optional query hashtable (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $false)]
        [hashtable]$Query,
        [Parameter(Mandatory = $false)]
        [ValidateSet('PowerBI', 'Fabric')]
        [string]$Api = 'PowerBI'
    )
    $url = $null
    if ($Path -match '^https?://') {
        $url = $Path
    }
    else {
        if (-not $script:IQ -or -not $script:IQ.Endpoints) {
            throw "ImpactIQ context has no Endpoints; call Initialize-IQContext before Invoke-IQApi."
        }
        $relative = $Path.TrimStart('/')
        if ($Api -eq 'Fabric') {
            $prefix = [string]$script:IQ.Endpoints.FabricApiPrefix
            if ([string]::IsNullOrWhiteSpace($prefix)) { throw "No FabricApiPrefix is configured for environment '$($script:IQ.Environment)'." }
            $url = $prefix.TrimEnd('/') + '/v1/' + $relative
        }
        else {
            $prefix = [string]$script:IQ.Endpoints.ApiPrefix
            if ([string]::IsNullOrWhiteSpace($prefix)) { throw "No ApiPrefix is configured for environment '$($script:IQ.Environment)'." }
            $url = $prefix.TrimEnd('/') + '/v1.0/myorg/' + $relative
        }
    }
    if ($Query -and $Query.Count -gt 0) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($key in @($Query.Keys | Sort-Object)) {
            $value = $Query[$key]
            if ($null -eq $value) { continue }
            if ($value -is [bool]) { $value = $value.ToString().ToLowerInvariant() }
            $parts.Add(([string]$key) + '=' + [System.Uri]::EscapeDataString([string]$value))
        }
        if ($parts.Count -gt 0) {
            $separator = '?'
            if ($url.Contains('?')) { $separator = '&' }
            $url = $url + $separator + ($parts -join '&')
        }
    }
    return $url
}

function Get-IQHttpDisplayPath {
    <#
    .SYNOPSIS
    Returns a short, secret-free form of a URL for log lines (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )
    $display = $Url -replace '(?i)(access_token|token|sig|password)=[^&]+', '$1=***'
    if ($script:IQ -and $script:IQ.Endpoints) {
        foreach ($name in @('ApiPrefix', 'FabricApiPrefix')) {
            $prefix = [string]$script:IQ.Endpoints[$name]
            if ($prefix -and $display.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                $display = $display.Substring($prefix.Length)
                break
            }
        }
    }
    if ($display.Length -gt 220) { $display = $display.Substring(0, 217) + '...' }
    return $display
}

function Get-IQHttpBearerToken {
    <#
    .SYNOPSIS
    Gets the bearer token for an API, forcing a refresh when asked (private; wraps Get-IQToken from ImpactIQ.Auth.ps1).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('PowerBI', 'Fabric')]
        [string]$Api,
        [Parameter(Mandatory = $false)]
        [switch]$ForceRefresh
    )
    if ($ForceRefresh) {
        # Get-IQToken refreshes when the cached token is close to expiry. To force a refresh after a 401 we either
        # pass -Force when the auth module supports it or drop the cached entry so the provider re-mints it.
        $command = Get-Command -Name Get-IQToken -ErrorAction SilentlyContinue
        if ($command -and $command.Parameters.ContainsKey('Force')) {
            return (Get-IQToken -Resource $Api -Force)
        }
        try {
            if ($script:IQ.Auth -and $script:IQ.Auth.ContainsKey('Tokens') -and $script:IQ.Auth.Tokens -is [hashtable]) {
                $script:IQ.Auth.Tokens.Remove($Api)
            }
        }
        catch {
            Write-IQLog -Level Debug -Stage Http -Message ("Could not clear the cached $Api token: " + $_.Exception.Message)
        }
    }
    return (Get-IQToken -Resource $Api)
}

function Invoke-IQHttpRequest {
    <#
    .SYNOPSIS
    Executes one HTTP request with the full ImpactIQ retry matrix and returns status/headers/content (private core).
    .DESCRIPTION
    Returns $null for 400/403/404 (logged), otherwise @{ StatusCode; Headers; Content; ContentType; ElapsedMs;
    OutFile }. Throws for statuses that are not retryable or after the retry budget is exhausted.
    Retry policy (brief section 2.3): 429 -> Retry-After seconds (default 30, max 300), up to 8 times;
    5xx/408 and transient network errors -> 2,4,8,16,32,60 s backoff up to $script:IQ.Options.MaxRetries attempts
    (default 5); 401 -> force one token refresh and retry once.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('GET', 'POST', 'PUT', 'DELETE', 'PATCH')]
        [string]$Method,
        [Parameter(Mandatory = $true)]
        [string]$Url,
        [Parameter(Mandatory = $false)]
        [object]$Body,
        [Parameter(Mandatory = $false)]
        [string]$ContentType = 'application/json',
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers,
        [Parameter(Mandatory = $false)]
        [ValidateSet('PowerBI', 'Fabric')]
        [string]$Api = 'PowerBI',
        [Parameter(Mandatory = $false)]
        [switch]$NoAuth,
        [Parameter(Mandatory = $false)]
        [switch]$AllowNotFound,
        [Parameter(Mandatory = $false)]
        [string]$OutFile,
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSec = 300,
        [Parameter(Mandatory = $false)]
        [string]$Stage = 'Http'
    )
    Initialize-IQHttpDefault
    $ProgressPreference = 'SilentlyContinue'   # Audit C1-07: the 5.1 progress bar makes -OutFile downloads very slow.

    $maxAttempts = 5
    try { if ($script:IQ.Options -and $null -ne $script:IQ.Options.MaxRetries -and [int]$script:IQ.Options.MaxRetries -gt 0) { $maxAttempts = [int]$script:IQ.Options.MaxRetries } } catch { $maxAttempts = 5 }
    $maxRateLimitRetries = 8
    $displayPath = Get-IQHttpDisplayPath -Url $Url

    # Prepare the body once.
    $requestBody = $null
    if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
        if ($Body -is [byte[]]) { $requestBody = $Body }
        elseif ($Body -is [string]) {
            if ($ContentType -match '(?i)json') { $requestBody = [System.Text.Encoding]::UTF8.GetBytes([string]$Body) }
            else { $requestBody = [string]$Body }
        }
        elseif (($Body -is [hashtable] -or $Body -is [System.Collections.IDictionary]) -and $ContentType -match '(?i)x-www-form-urlencoded') {
            $requestBody = $Body      # Invoke-WebRequest form-encodes dictionaries.
        }
        else {
            $json = ConvertTo-Json -InputObject $Body -Depth 20 -Compress
            $requestBody = [System.Text.Encoding]::UTF8.GetBytes($json)
        }
        if ($ContentType -match '(?i)^application/json$') { $ContentType = 'application/json; charset=utf-8' }
    }

    $attempt = 0
    $rateLimitRetries = 0
    $authRetried = $false
    $forceRefresh = $false
    $fatalMessage = $null
    $fatalInner = $null
    # NOTE: the loop exits with 'break' and throws afterwards. A 'throw' inside the catch block would be swallowed
    # when a caller passes -ErrorAction SilentlyContinue (PowerShell then continues after the throw) and the loop
    # would spin forever.
    while ($true) {
        $attempt++
        # User-Agent goes through -UserAgent (a restricted header on .NET Framework / Windows PowerShell 5.1).
        $requestHeaders = @{}
        if ($Headers) { foreach ($key in $Headers.Keys) { $requestHeaders[$key] = $Headers[$key] } }
        if (-not $NoAuth) {
            $token = Get-IQHttpBearerToken -Api $Api -ForceRefresh:$forceRefresh
            $forceRefresh = $false
            if ([string]::IsNullOrEmpty($token)) {
                if ($Api -eq 'Fabric') {
                    Write-IQLog -Level Debug -Stage $Stage -Message "No Fabric token available; skipping $Method $displayPath"
                    return $null
                }
                $fatalMessage = "No Power BI access token is available for $Method $displayPath."
                break
            }
            $requestHeaders['Authorization'] = "Bearer $token"
        }

        $params = @{ Uri = $Url; Method = $Method; Headers = $requestHeaders; UseBasicParsing = $true; TimeoutSec = $TimeoutSec; UserAgent = 'ImpactIQ/3.0'; ErrorAction = 'Stop' }
        if ($null -ne $requestBody) {
            $params.Body = $requestBody
            $params.ContentType = $ContentType
        }
        if ($OutFile) {
            $params.OutFile = $OutFile
            $params.PassThru = $true
            $outDir = Split-Path -Path $OutFile -Parent
            if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -Path $outDir -ItemType Directory -Force | Out-Null }
        }

        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        if ($script:IQ -and $script:IQ.Stats) { $script:IQ.Stats.ApiCalls = [int]$script:IQ.Stats.ApiCalls + 1 }
        try {
            $response = Invoke-WebRequest @params
            $stopwatch.Stop()
            $statusCode = 200
            try { if ($null -ne $response -and $null -ne $response.StatusCode) { $statusCode = [int]$response.StatusCode } } catch { $statusCode = 200 }
            Write-IQLog -Level Debug -Stage $Stage -Message ("{0} {1} -> {2} ({3} ms)" -f $Method, $displayPath, $statusCode, $stopwatch.ElapsedMilliseconds)

            $result = @{ StatusCode = $statusCode; Headers = $null; Content = $null; ContentType = $null; ElapsedMs = $stopwatch.ElapsedMilliseconds; OutFile = $OutFile }
            if ($null -ne $response) {
                try { $result.Headers = $response.Headers } catch { $result.Headers = $null }
                $result.ContentType = Get-IQHttpHeaderValue -Headers $result.Headers -Name 'Content-Type'
            }
            if ($OutFile) {
                if ((Test-Path -LiteralPath $OutFile) -and (Get-Item -LiteralPath $OutFile).Length -gt 0) {
                    $result.Content = $null
                }
                else {
                    if (Test-Path -LiteralPath $OutFile) { Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue }
                    Write-IQLog -Level Warn -Stage $Stage -Message "$Method $displayPath returned an empty body; no file written to $OutFile"
                    $result.StatusCode = $statusCode
                    $result.OutFile = $null
                }
                return $result
            }
            if ($null -ne $response) {
                $content = $null
                $charsetMissing = ($null -eq $result.ContentType) -or ($result.ContentType -notmatch '(?i)charset=')
                if ($charsetMissing -and $response.PSObject.Properties['RawContentStream'] -and $null -ne $response.RawContentStream -and $response.RawContentStream.PSObject.Methods['ToArray']) {
                    # 5.1 decodes bodies without a charset as ISO-8859-1; re-decode as UTF-8 so names survive.
                    try { $content = [System.Text.Encoding]::UTF8.GetString($response.RawContentStream.ToArray()) } catch { $content = $null }
                }
                if ($null -eq $content -and $response.PSObject.Properties['Content']) {
                    $raw = $response.Content
                    if ($raw -is [byte[]]) { $content = [System.Text.Encoding]::UTF8.GetString($raw) }
                    elseif ($null -ne $raw) { $content = [string]$raw }
                }
                if ($null -ne $content -and $content.Length -gt 0 -and $content[0] -eq [char]0xFEFF) { $content = $content.Substring(1) }
                $result.Content = $content
            }
            return $result
        }
        catch {
            $stopwatch.Stop()
            $info = Get-IQHttpErrorInfo -ErrorRecord $_
            $status = $info.StatusCode
            $bodySnippet = ''
            if ($info.Body) { $bodySnippet = ($info.Body -replace '\s+', ' ').Trim(); if ($bodySnippet.Length -gt 400) { $bodySnippet = $bodySnippet.Substring(0, 400) + '...' } }
            if ($OutFile -and (Test-Path -LiteralPath $OutFile)) { Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue }
            Write-IQLog -Level Debug -Stage $Stage -Message ("{0} {1} -> {2} ({3} ms) attempt {4}: {5}" -f $Method, $displayPath, $(if ($null -ne $status) { $status } else { 'no-response' }), $stopwatch.ElapsedMilliseconds, $attempt, $info.Message)

            if ($status -eq 429) {
                if ($rateLimitRetries -lt $maxRateLimitRetries) {
                    $rateLimitRetries++
                    $wait = 30
                    if ($null -ne $info.RetryAfter) { $wait = [int]$info.RetryAfter }
                    if ($wait -lt 1) { $wait = 1 }
                    if ($wait -gt 300) { $wait = 300 }
                    if ($script:IQ -and $script:IQ.Stats) { $script:IQ.Stats.Retries = [int]$script:IQ.Stats.Retries + 1 }
                    Write-IQLog -Level Warn -Stage $Stage -Message "429 Too Many Requests for $Method $displayPath; waiting $wait s (retry $rateLimitRetries/$maxRateLimitRetries)"
                    Start-IQHttpSleep -Seconds $wait
                    continue
                }
                $fatalMessage = "HTTP 429 for $Method $displayPath persisted after $maxRateLimitRetries retries. $bodySnippet"
                $fatalInner = $_.Exception
                break
            }
            if ($status -eq 401) {
                if (-not $authRetried -and -not $NoAuth) {
                    $authRetried = $true
                    $forceRefresh = $true
                    if ($script:IQ -and $script:IQ.Stats) { $script:IQ.Stats.Retries = [int]$script:IQ.Stats.Retries + 1 }
                    Write-IQLog -Level Warn -Stage $Stage -Message "401 Unauthorized for $Method $displayPath; refreshing the token and retrying once"
                    continue
                }
                $fatalMessage = "HTTP 401 for $Method $displayPath after a token refresh. $bodySnippet"
                $fatalInner = $_.Exception
                break
            }
            if ($status -eq 403 -or $status -eq 404) {
                $level = 'Warn'
                if ($AllowNotFound) { $level = 'Debug' }
                Write-IQLog -Level $level -Stage $Stage -Message ("HTTP {0} for {1} {2}. {3}" -f $status, $Method, $displayPath, $bodySnippet)
                return $null
            }
            if ($status -eq 400) {
                Write-IQLog -Level Warn -Stage $Stage -Message ("HTTP 400 for {0} {1}. {2}" -f $Method, $displayPath, $bodySnippet)
                return $null
            }
            $retryable = ($null -eq $status -and $info.Transient) -or ($status -eq 408) -or ($null -ne $status -and $status -ge 500 -and $status -le 599)
            if ($retryable -and $attempt -lt $maxAttempts) {
                $wait = [math]::Min(60, [math]::Pow(2, $attempt))
                if ($null -ne $info.RetryAfter -and $info.RetryAfter -gt $wait) { $wait = [math]::Min(300, $info.RetryAfter) }
                if ($script:IQ -and $script:IQ.Stats) { $script:IQ.Stats.Retries = [int]$script:IQ.Stats.Retries + 1 }
                Write-IQLog -Level Warn -Stage $Stage -Message ("{0} for {1} {2}; retrying in {3} s (attempt {4}/{5}): {6}" -f $(if ($null -ne $status) { "HTTP $status" } else { 'Network error' }), $Method, $displayPath, [int]$wait, $attempt, $maxAttempts, $info.Message)
                Start-IQHttpSleep -Seconds $wait
                continue
            }
            $statusText = 'no response'
            if ($null -ne $status) { $statusText = "HTTP $status" }
            $fatalMessage = "$statusText for $Method $displayPath after $attempt attempt(s): $($info.Message)"
            if ($bodySnippet) { $fatalMessage += " Body: $bodySnippet" }
            $fatalInner = $_.Exception
            break
        }
    }
    $fatalError = New-Object System.Exception($fatalMessage, $fatalInner)
    throw $fatalError
}

function ConvertFrom-IQHttpJson {
    <#
    .SYNOPSIS
    Parses a response body as JSON; returns $null for empty bodies and the raw string when parsing fails (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Content,
        [Parameter(Mandatory = $false)]
        [string]$Stage = 'Http'
    )
    if ([string]::IsNullOrWhiteSpace($Content)) { return $null }
    try {
        return (ConvertFrom-Json -InputObject $Content)
    }
    catch {
        Write-IQLog -Level Debug -Stage $Stage -Message ("Response body is not JSON (" + $_.Exception.Message + "); returning the raw string")
        return $Content
    }
}

function Get-IQContinuationUrl {
    <#
    .SYNOPSIS
    Returns the URL of the next page from @odata.nextLink / continuationUri / continuationToken, or $null (private).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Response,
        [Parameter(Mandatory = $true)]
        [string]$RequestUrl
    )
    if ($null -eq $Response -or -not ($Response -is [System.Management.Automation.PSCustomObject])) { return $null }
    $props = $Response.PSObject.Properties
    $next = $null
    if ($props['@odata.nextLink'] -and -not [string]::IsNullOrWhiteSpace([string]$props['@odata.nextLink'].Value)) {
        $next = [string]$props['@odata.nextLink'].Value
    }
    elseif ($props['continuationUri'] -and -not [string]::IsNullOrWhiteSpace([string]$props['continuationUri'].Value)) {
        $next = [string]$props['continuationUri'].Value
    }
    elseif ($props['continuationToken'] -and -not [string]::IsNullOrWhiteSpace([string]$props['continuationToken'].Value)) {
        # Fabric style: the token is re-sent as ?continuationToken=<token> on the original URL (replacing any previous token).
        $token = [string]$props['continuationToken'].Value
        $baseUrl = $RequestUrl -replace '([?&])continuationToken=[^&]*(&|$)', '$1'
        $baseUrl = $baseUrl.TrimEnd('?', '&')
        $separator = '?'
        if ($baseUrl.Contains('?')) { $separator = '&' }
        $next = $baseUrl + $separator + 'continuationToken=' + [System.Uri]::EscapeDataString($token)
    }
    if ($next -and ($next -eq $RequestUrl)) { return $null }   # guard against a server echoing the same page forever
    return $next
}

function Invoke-IQApi {
    <#
    .SYNOPSIS
    Calls the Power BI or Fabric REST API with retries, paging and consistent error handling; returns parsed JSON.
    .DESCRIPTION
    -Path is relative to <ApiPrefix>/v1.0/myorg/ (PowerBI) or <FabricApiPrefix>/v1/ (Fabric); absolute https:// paths
    are used verbatim. -Query values are URL-encoded with [System.Uri]::EscapeDataString. When the response has a
    'value' array and a continuation (@odata.nextLink, continuationUri or continuationToken) all pages are followed
    (unless -NoPaging) and the returned object's 'value' is the concatenation (10 000 page guard).
    -Raw returns the body string; -OutFile streams the body to disk and returns $true only when the file exists and
    is non-empty. 403/404 return $null (Warn, or Debug with -AllowNotFound); 400 returns $null with a Warn.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('GET', 'POST', 'PUT', 'DELETE', 'PATCH')]
        [string]$Method = 'GET',
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $false)]
        [object]$Body,
        [Parameter(Mandatory = $false)]
        [hashtable]$Query,
        [Parameter(Mandatory = $false)]
        [switch]$Raw,
        [Parameter(Mandatory = $false)]
        [switch]$AllowNotFound,
        [Parameter(Mandatory = $false)]
        [switch]$NoPaging,
        [Parameter(Mandatory = $false)]
        [string]$OutFile,
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSec = 300,
        [Parameter(Mandatory = $false)]
        [ValidateSet('PowerBI', 'Fabric')]
        [string]$Api = 'PowerBI',
        [Parameter(Mandatory = $false)]
        [string]$ContentType = 'application/json',
        [Parameter(Mandatory = $false)]
        [hashtable]$Headers,
        [Parameter(Mandatory = $false)]
        [switch]$NoAuth,
        [Parameter(Mandatory = $false)]
        [string]$Stage = 'Http'
    )
    $url = Get-IQApiUrl -Path $Path -Query $Query -Api $Api
    $requestParams = @{ Method = $Method; Url = $url; Api = $Api; TimeoutSec = $TimeoutSec; AllowNotFound = $AllowNotFound; NoAuth = $NoAuth; Stage = $Stage; ContentType = $ContentType }
    if ($PSBoundParameters.ContainsKey('Body')) { $requestParams.Body = $Body }
    if ($Headers) { $requestParams.Headers = $Headers }

    if ($OutFile) {
        $requestParams.OutFile = $OutFile
        $result = Invoke-IQHttpRequest @requestParams
        if ($null -eq $result) { return $false }
        return [bool]((Test-Path -LiteralPath $OutFile) -and (Get-Item -LiteralPath $OutFile).Length -gt 0)
    }

    $result = Invoke-IQHttpRequest @requestParams
    if ($null -eq $result) { return $null }
    if ($Raw) { return $result.Content }

    $first = ConvertFrom-IQHttpJson -Content $result.Content -Stage $Stage
    if ($NoPaging -or $null -eq $first -or -not ($first -is [System.Management.Automation.PSCustomObject]) -or -not $first.PSObject.Properties['value']) {
        return $first
    }

    $nextUrl = Get-IQContinuationUrl -Response $first -RequestUrl $url
    if (-not $nextUrl) { return $first }

    $all = New-Object System.Collections.Generic.List[object]
    if ($null -ne $first.value) { $all.AddRange([object[]]@($first.value)) }
    $pages = 1
    $maxPages = 10000
    $currentUrl = $nextUrl
    while ($currentUrl -and $pages -lt $maxPages) {
        $pages++
        $pageParams = @{ Method = 'GET'; Url = $currentUrl; Api = $Api; TimeoutSec = $TimeoutSec; AllowNotFound = $AllowNotFound; NoAuth = $NoAuth; Stage = $Stage }
        if ($Headers) { $pageParams.Headers = $Headers }
        $pageResult = Invoke-IQHttpRequest @pageParams
        if ($null -eq $pageResult) {
            Write-IQLog -Level Warn -Stage $Stage -Message "Paging stopped at page $pages for $(Get-IQHttpDisplayPath -Url $url); returning $($all.Count) rows collected so far"
            break
        }
        $page = ConvertFrom-IQHttpJson -Content $pageResult.Content -Stage $Stage
        if ($null -eq $page -or -not ($page -is [System.Management.Automation.PSCustomObject])) { break }
        if ($page.PSObject.Properties['value'] -and $null -ne $page.value) { $all.AddRange([object[]]@($page.value)) }
        $currentUrl = Get-IQContinuationUrl -Response $page -RequestUrl $currentUrl
    }
    if ($pages -ge $maxPages) { Write-IQLog -Level Warn -Stage $Stage -Message "Paging guard hit ($maxPages pages) for $(Get-IQHttpDisplayPath -Url $url)" }
    Write-IQLog -Level Debug -Stage $Stage -Message ("Paged {0} pages / {1} rows for {2}" -f $pages, $all.Count, (Get-IQHttpDisplayPath -Url $url))

    $first | Add-Member -MemberType NoteProperty -Name 'value' -Value $all.ToArray() -Force
    foreach ($name in @('@odata.nextLink', 'continuationUri', 'continuationToken')) {
        if ($first.PSObject.Properties[$name]) { $first.PSObject.Properties.Remove($name) }
    }
    return $first
}

function Invoke-IQFabricLro {
    <#
    .SYNOPSIS
    Calls a Fabric long-running-operation endpoint (e.g. getDefinition) and returns the final result object or $null.
    .DESCRIPTION
    200 -> parsed body. 202 -> polls the Location header (or operations/<x-ms-operation-id>) every Retry-After
    seconds (default 5) until status is Succeeded, then GETs <operationUrl>/result. Failed/timeout/403/404 -> $null
    with a Warn. Never throws for operation failures; throws only for unrecoverable HTTP errors.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('GET', 'POST')]
        [string]$Method = 'POST',
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $false)]
        [object]$Body,
        [Parameter(Mandatory = $false)]
        [int]$TimeoutMinutes = 10,
        [Parameter(Mandatory = $false)]
        [hashtable]$Query,
        [Parameter(Mandatory = $false)]
        [string]$Stage = 'Http'
    )
    $url = Get-IQApiUrl -Path $Path -Query $Query -Api Fabric
    $displayPath = Get-IQHttpDisplayPath -Url $url
    $requestParams = @{ Method = $Method; Url = $url; Api = 'Fabric'; Stage = $Stage; AllowNotFound = $true }
    if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) { $requestParams.Body = $Body }
    elseif ($Method -eq 'POST') { $requestParams.Body = '{}' }   # Fabric rejects a POST without a JSON body on some endpoints

    $result = Invoke-IQHttpRequest @requestParams
    if ($null -eq $result) {
        Write-IQLog -Level Warn -Stage $Stage -Message "Fabric operation $Method $displayPath was not accepted (see previous message)"
        return $null
    }
    if ($result.StatusCode -ne 202) {
        return (ConvertFrom-IQHttpJson -Content $result.Content -Stage $Stage)
    }

    $location = Get-IQHttpHeaderValue -Headers $result.Headers -Name 'Location'
    $operationId = Get-IQHttpHeaderValue -Headers $result.Headers -Name 'x-ms-operation-id'
    $operationUrl = $null
    if (-not [string]::IsNullOrWhiteSpace($location)) {
        $operationUrl = $location
        if ($operationUrl -notmatch '^https?://') { $operationUrl = Get-IQApiUrl -Path $operationUrl -Api Fabric }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($operationId)) {
        $operationUrl = Get-IQApiUrl -Path ("operations/" + $operationId) -Api Fabric
    }
    if (-not $operationUrl) {
        Write-IQLog -Level Warn -Stage $Stage -Message "Fabric returned 202 for $Method $displayPath without a Location or x-ms-operation-id header"
        return $null
    }
    $operationUrl = $operationUrl.TrimEnd('/')
    $retryAfter = ConvertTo-IQRetryAfterDelay -Value (Get-IQHttpHeaderValue -Headers $result.Headers -Name 'Retry-After')
    if ($null -eq $retryAfter -or $retryAfter -lt 1) { $retryAfter = 5 }
    if ($retryAfter -gt 60) { $retryAfter = 60 }
    Write-IQLog -Level Debug -Stage $Stage -Message "Fabric LRO accepted for $Method $displayPath; polling $(Get-IQHttpDisplayPath -Url $operationUrl) every $retryAfter s"

    $deadline = [datetime]::UtcNow.AddMinutes([math]::Max(1, $TimeoutMinutes))
    $status = $null
    $polls = 0
    while ([datetime]::UtcNow -lt $deadline) {
        Start-IQHttpSleep -Seconds $retryAfter
        $polls++
        $pollResult = Invoke-IQHttpRequest -Method GET -Url $operationUrl -Api Fabric -Stage $Stage
        if ($null -eq $pollResult) {
            Write-IQLog -Level Warn -Stage $Stage -Message "Fabric operation status for $displayPath could not be read (poll $polls)"
            return $null
        }
        $pollRetry = ConvertTo-IQRetryAfterDelay -Value (Get-IQHttpHeaderValue -Headers $pollResult.Headers -Name 'Retry-After')
        if ($null -ne $pollRetry -and $pollRetry -ge 1 -and $pollRetry -le 60) { $retryAfter = $pollRetry }
        $state = ConvertFrom-IQHttpJson -Content $pollResult.Content -Stage $Stage
        $status = $null
        if ($state -is [System.Management.Automation.PSCustomObject] -and $state.PSObject.Properties['status']) { $status = [string]$state.status }
        if ($status -ieq 'Succeeded') {
            $finalResult = Invoke-IQHttpRequest -Method GET -Url ($operationUrl + '/result') -Api Fabric -Stage $Stage
            if ($null -eq $finalResult) {
                Write-IQLog -Level Warn -Stage $Stage -Message "Fabric operation for $displayPath succeeded but its result could not be fetched"
                return $null
            }
            return (ConvertFrom-IQHttpJson -Content $finalResult.Content -Stage $Stage)
        }
        if ($status -ieq 'Failed' -or $status -ieq 'Cancelled' -or $status -ieq 'Canceled') {
            $errorText = ''
            try { if ($state.PSObject.Properties['error'] -and $state.error) { $errorText = (ConvertTo-Json -InputObject $state.error -Depth 20 -Compress) } } catch { $errorText = '' }
            if ($errorText.Length -gt 400) { $errorText = $errorText.Substring(0, 400) + '...' }
            Write-IQLog -Level Warn -Stage $Stage -Message "Fabric operation for $displayPath ended with status '$status'. $errorText"
            return $null
        }
        # NotStarted / Running / Undefined: keep polling.
    }
    Write-IQLog -Level Warn -Stage $Stage -Message "Fabric operation for $displayPath did not finish within $TimeoutMinutes minute(s) (last status: '$status', $polls polls)"
    return $null
}

function Invoke-IQDownload {
    <#
    .SYNOPSIS
    Downloads a URL to a file with the standard retry policy; returns $true when the file exists and is non-empty.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url,
        [Parameter(Mandatory = $true)]
        [string]$OutFile,
        [Parameter(Mandatory = $false)]
        [ValidateSet('PowerBI', 'Fabric')]
        [string]$Api = 'PowerBI',
        [Parameter(Mandatory = $false)]
        [switch]$NoAuth,
        [Parameter(Mandatory = $false)]
        [int]$TimeoutSec = 600,
        [Parameter(Mandatory = $false)]
        [string]$Stage = 'Http'
    )
    try {
        $fullUrl = Get-IQApiUrl -Path $Url -Api $Api
        $result = Invoke-IQHttpRequest -Method GET -Url $fullUrl -Api $Api -OutFile $OutFile -TimeoutSec $TimeoutSec -NoAuth:$NoAuth -Stage $Stage
        if ($null -eq $result) { return $false }
        if ((Test-Path -LiteralPath $OutFile) -and (Get-Item -LiteralPath $OutFile).Length -gt 0) { return $true }
        if (Test-Path -LiteralPath $OutFile) { Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue }
        return $false
    }
    catch {
        Write-IQLog -Level Warn -Stage $Stage -Message ("Download failed for {0}: {1}" -f (Get-IQHttpDisplayPath -Url $Url), $_.Exception.Message)
        if (Test-Path -LiteralPath $OutFile) { Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue }
        return $false
    }
}
