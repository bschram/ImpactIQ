# Http.Tests.ps1 - ImpactIQ.Http.ps1 (brief section 2.3): URL building, paging, retry matrix, error info, LRO, downloads.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Base = Initialize-IQTestContext -Options @{ Environment = 'USGov'; MaxRetries = 3 } -NoRun -Prefix 'http'
    $script:Requests = New-Object System.Collections.Generic.List[object]
    $script:Sleeps = New-Object System.Collections.Generic.List[double]
    $script:Responses = New-Object System.Collections.Generic.List[object]
    function Add-HttpResponse {
        # queue an item: a response object (returned) or an ErrorRecord (thrown)
        param([object]$Item)
        $script:Responses.Add($Item)
    }
    function Pop-HttpResponse {
        param($Uri, $Method, $Headers, $Body, $OutFile, $UserAgent, $ContentType)
        $uriText = [string]$Uri
        if ($Uri -is [uri]) { $uriText = $Uri.OriginalString }   # [uri]::ToString() un-escapes %20 etc.
        $script:Requests.Add(@{ Uri = $uriText; Method = "$Method"; Headers = $Headers; Body = $Body; OutFile = $OutFile; UserAgent = $UserAgent; ContentType = $ContentType })
        if ($script:Responses.Count -eq 0) { throw "test mock: no HTTP response queued for $Method $Uri" }
        $item = $script:Responses[0]
        $script:Responses.RemoveAt(0)
        if ($item -is [scriptblock]) { $item = & $item $Uri }
        if ($item -is [System.Management.Automation.ErrorRecord]) { throw $item }
        if ($OutFile -and $null -ne $item -and $null -ne $item.Content) {
            $dir = Split-Path -Parent $OutFile
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            [System.IO.File]::WriteAllText($OutFile, [string]$item.Content)
        }
        return $item
    }
    function Reset-Http {
        $script:Requests.Clear(); $script:Responses.Clear(); $script:Sleeps.Clear()
        $script:IQ.Stats.ApiCalls = 0; $script:IQ.Stats.Retries = 0
    }
    Mock Get-IQToken { if ($Resource -eq 'Fabric') { return 'fabric-token' } return 'pbi-token' }
    Mock Invoke-WebRequest { Pop-HttpResponse -Uri $Uri -Method $Method -Headers $Headers -Body $Body -OutFile $OutFile -UserAgent $UserAgent -ContentType $ContentType }
    Mock Start-IQHttpSleep { $script:Sleeps.Add([double]$Seconds) }
    Mock Start-Sleep { }
}
AfterAll { Remove-IQTestFolder -Path $script:Base }

Describe 'Invoke-IQApi URL building and headers' {
    BeforeEach { Reset-Http }
    It 'appends a relative Power BI path to ApiPrefix/v1.0/myorg/ and strips a leading slash' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @() })
        Invoke-IQApi -Method GET -Path '/groups' | Out-Null
        $script:Requests[0].Uri | Should -Be 'https://api.powerbigov.us/v1.0/myorg/groups'
    }
    It 'uses FabricApiPrefix/v1/ for -Api Fabric' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @() })
        Invoke-IQApi -Method GET -Path 'workspaces' -Api Fabric | Out-Null
        $script:Requests[0].Uri | Should -Be ((Get-IQContext).Endpoints.FabricApiPrefix + '/v1/workspaces')
    }
    It 'uses absolute https:// paths verbatim' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @() })
        Invoke-IQApi -Method GET -Path 'https://example.gov/api/x?y=1' | Out-Null
        $script:Requests[0].Uri | Should -Be 'https://example.gov/api/x?y=1'
    }
    It 'encodes -Query values' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @() })
        Invoke-IQApi -Method GET -Path 'groups' -Query @{ '$top' = 5000; '$filter' = "name eq 'A B'" } | Out-Null
        $script:Requests[0].Uri | Should -Match '\$top=5000'
        $script:Requests[0].Uri | Should -Match 'A%20B'
    }
    It 'sends the Authorization bearer header, the ImpactIQ user agent and a JSON body' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ ok = 1 })
        Invoke-IQApi -Method POST -Path 'datasets/x/executeQueries' -Body @{ queries = @(@{ query = 'EVALUATE T' }) } -NoPaging | Out-Null
        $r = $script:Requests[0]
        $r.Headers['Authorization'] | Should -Be 'Bearer pbi-token'
        $r.UserAgent | Should -Be 'ImpactIQ/3.0'
        $r.ContentType | Should -Match '(?i)application/json'
        $bodyText = $r.Body
        if ($bodyText -is [byte[]]) { $bodyText = [System.Text.Encoding]::UTF8.GetString($bodyText) }
        [string]$bodyText | Should -Match 'EVALUATE T'
        $r.Method | Should -Be 'POST'
    }
    It 'returns the parsed JSON object and counts the API call' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 'a' }, @{ id = 'b' }) })
        $res = Invoke-IQApi -Method GET -Path 'groups'
        @($res.value).Count | Should -Be 2
        (Get-IQContext).Stats.ApiCalls | Should -Be 1
    }
    It '-Raw returns the body string' {
        Add-HttpResponse (New-IQTestHttpResponse -Content '{"raw":true}')
        Invoke-IQApi -Method GET -Path 'groups/x/dataflows/y' -Raw | Should -Be '{"raw":true}'
    }
    It 'returns $null with a Debug line when no Fabric token is available' {
        Mock Get-IQToken { if ($Resource -eq 'Fabric') { return $null } return 'pbi-token' }
        Invoke-IQApi -Method GET -Path 'workspaces' -Api Fabric | Should -BeNullOrEmpty
        $script:Requests.Count | Should -Be 0
    }
}

Describe 'Paging' {
    BeforeEach { Reset-Http }
    It 'follows @odata.nextLink and continuationUri and concatenates value' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 1 }, @{ id = 2 }); '@odata.nextLink' = 'https://api.powerbigov.us/v1.0/myorg/groups?$skip=2' })
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 3 }); continuationUri = 'https://api.powerbigov.us/v1.0/myorg/groups?$skip=3' })
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 4 }) })
        $res = Invoke-IQApi -Method GET -Path 'groups'
        @($res.value).Count | Should -Be 4
        @($res.value)[3].id | Should -Be 4
        $script:Requests.Count | Should -Be 3
        $script:Requests[1].Uri | Should -Match '\$skip=2'
        $res.PSObject.Properties['@odata.nextLink'] | Should -BeNullOrEmpty
        $res.PSObject.Properties['continuationUri'] | Should -BeNullOrEmpty
    }
    It 're-sends a Fabric continuationToken as a query parameter' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 1 }); continuationToken = 'tok en+1' })
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 2 }); continuationToken = $null })
        $res = Invoke-IQApi -Method GET -Path 'workspaces' -Api Fabric
        @($res.value).Count | Should -Be 2
        $script:Requests[1].Uri | Should -Match 'continuationToken=tok%20en%2B1'
    }
    It 'does not page with -NoPaging' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 1 }); '@odata.nextLink' = 'https://api.powerbigov.us/v1.0/myorg/groups?$skip=1' })
        $res = Invoke-IQApi -Method GET -Path 'groups' -NoPaging
        @($res.value).Count | Should -Be 1
        $script:Requests.Count | Should -Be 1
    }
    It 'throws when a later page fails with 404 instead of returning a truncated list as complete' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 1 }); '@odata.nextLink' = 'https://api.powerbigov.us/v1.0/myorg/groups?$skip=1' })
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 404 -Body '{"error":{"code":"NotFound"}}')
        { Invoke-IQApi -Method GET -Path 'groups' } | Should -Throw -ExpectedMessage '*page 2*404*'
        $script:Requests.Count | Should -Be 2
    }
    It 'throws when a Fabric continuationToken is rejected with 400 even with -AllowNotFound' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 1 }); continuationToken = 'expired' })
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 400 -Body '{"error":{"code":"InvalidContinuationToken"}}')
        { Invoke-IQApi -Method GET -Path 'workspaces' -Api Fabric -AllowNotFound } | Should -Throw -ExpectedMessage '*page 2*400*'
    }
    It 'with -AllowNotFound a 404 on a later page returns the rows so far stamped IQPartial' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 1 }); '@odata.nextLink' = 'https://api.powerbigov.us/v1.0/myorg/groups?$skip=1' })
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 404 -Body '{"error":{"code":"NotFound"}}')
        $res = Invoke-IQApi -Method GET -Path 'groups' -AllowNotFound
        @($res.value).Count | Should -Be 1
        $res.IQPartial | Should -BeTrue
        $res.IQPagesFetched | Should -Be 1
        (Get-Content -LiteralPath $script:IQ.LogFile -Raw) | Should -Match '\[WARN\].*Paging stopped at page 2.*partial'
    }
    It 'a complete paged result carries no IQPartial stamp' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 1 }); '@odata.nextLink' = 'https://api.powerbigov.us/v1.0/myorg/groups?$skip=1' })
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 2 }) })
        $res = Invoke-IQApi -Method GET -Path 'groups' -AllowNotFound
        @($res.value).Count | Should -Be 2
        $res.PSObject.Properties['IQPartial'] | Should -BeNullOrEmpty
    }
}

Describe 'Retry matrix' {
    BeforeEach { Reset-Http }
    It '429: waits Retry-After seconds and retries' {
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 429 -Body '{"error":{"code":"TooManyRequests"}}' -Headers @{ 'Retry-After' = '7' })
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 1 }) })
        $res = Invoke-IQApi -Method GET -Path 'groups'
        @($res.value).Count | Should -Be 1
        $script:Sleeps.Count | Should -Be 1
        $script:Sleeps[0] | Should -Be 7
        (Get-IQContext).Stats.Retries | Should -Be 1
    }
    It '429 without Retry-After waits the 30 s default; values above 300 are capped' {
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 429)
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 429 -Headers @{ 'Retry-After' = '9999' })
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @() })
        Invoke-IQApi -Method GET -Path 'groups' | Out-Null
        $script:Sleeps[0] | Should -Be 30
        $script:Sleeps[1] | Should -Be 300
    }
    It '401: forces one token refresh and retries once' {
        Mock Get-IQToken { if ($Force) { return 'pbi-token-refreshed' } return 'pbi-token' }
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 401 -Body 'Unauthorized')
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 1 }) })
        $res = Invoke-IQApi -Method GET -Path 'groups'
        @($res.value).Count | Should -Be 1
        $script:Requests.Count | Should -Be 2
        $script:Requests[1].Headers['Authorization'] | Should -Be 'Bearer pbi-token-refreshed'
        Should -Invoke Get-IQToken -ParameterFilter { $Force } -Times 1
    }
    It '401 twice: throws after the single refresh' {
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 401)
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 401)
        { Invoke-IQApi -Method GET -Path 'groups' } | Should -Throw -ExpectedMessage '*401*'
        $script:Requests.Count | Should -Be 2
    }
    It '404 and 403: no retry, returns $null' {
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 404 -Body '{"error":{"code":"ItemNotFound"}}')
        Invoke-IQApi -Method GET -Path 'groups/x/datasets/y' -AllowNotFound | Should -BeNullOrEmpty
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 403 -Body 'Forbidden')
        Invoke-IQApi -Method GET -Path 'groups/x/users' | Should -BeNullOrEmpty
        $script:Requests.Count | Should -Be 2
        $script:Sleeps.Count | Should -Be 0
    }
    It '400: no retry, returns $null and logs the body at Warn' {
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 400 -Body '{"error":{"code":"DaxQueryFailure","message":"INFO functions are not supported"}}')
        Invoke-IQApi -Method POST -Path 'datasets/x/executeQueries' -Body @{ q = 1 } -NoPaging | Should -BeNullOrEmpty
        $script:Requests.Count | Should -Be 1
        (Get-Content -LiteralPath $script:IQ.LogFile -Raw) | Should -Match '\[WARN\].*HTTP 400.*INFO functions'
    }
    It '400: records StatusCode and the body in $script:IQ.LastHttpError for the Dax module' {
        $body = '{"error":{"code":"DaxQueryFailure","message":"INFO functions are not supported"}}'
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 400 -Body $body)
        Invoke-IQApi -Method POST -Path 'datasets/x/executeQueries' -Body @{ q = 1 } -NoPaging | Should -BeNullOrEmpty
        $script:IQ.ContainsKey('LastHttpError') | Should -BeTrue
        $script:IQ.LastHttpError.StatusCode | Should -Be 400
        $script:IQ.LastHttpError.Body | Should -Be $body
        $script:IQ.LastHttpError.Method | Should -Be 'POST'
        $script:IQ.LastHttpError.Url | Should -Match 'executeQueries'
    }
    It '403 records LastHttpError too, and a following successful request clears it' {
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 403 -Body 'Forbidden')
        Invoke-IQApi -Method GET -Path 'groups/x/users' | Should -BeNullOrEmpty
        $script:IQ.LastHttpError.StatusCode | Should -Be 403
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @() })
        Invoke-IQApi -Method GET -Path 'groups' | Out-Null
        $script:IQ.LastHttpError | Should -BeNullOrEmpty
    }
    It 'no Fabric token: records a LastHttpError without a status' {
        Mock Get-IQToken { if ($Resource -eq 'Fabric') { return $null } return 'pbi-token' }
        Invoke-IQApi -Method GET -Path 'workspaces' -Api Fabric | Should -BeNullOrEmpty
        $script:IQ.LastHttpError | Should -Not -BeNullOrEmpty
        $script:IQ.LastHttpError.StatusCode | Should -BeNullOrEmpty
        $script:IQ.LastHttpError.Message | Should -Match 'Fabric token'
    }
    It '429 retries do not consume the 5xx budget: three 429s, a 503 and then 200 succeeds' {
        foreach ($i in 1..3) { Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 429 -Headers @{ 'Retry-After' = '1' }) }
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 503)
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 1 }) })
        $res = Invoke-IQApi -Method GET -Path 'groups'
        @($res.value).Count | Should -Be 1
        $script:Requests.Count | Should -Be 5
        $script:Sleeps.Count | Should -Be 4
        $script:Sleeps[3] | Should -Be 2 -Because 'the first 5xx backoff is 2 s regardless of earlier 429 waits'
    }
    It 'a 401 refresh does not consume the 5xx budget either' {
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 401)
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 500)
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 500)
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 1 }) })
        $res = Invoke-IQApi -Method GET -Path 'groups'
        @($res.value).Count | Should -Be 1
        $script:Sleeps.Count | Should -Be 2
        $script:Sleeps[0] | Should -Be 2
        $script:Sleeps[1] | Should -Be 4
    }
    It 'stops retrying (throws, no sleep) when the run time budget is used up' {
        $script:IQ.Options['TimeBudgetMinutes'] = 5
        $script:IQ.StartedUtc = [datetime]::UtcNow.AddMinutes(-10)
        $script:IQ.BudgetExceeded = $false
        try {
            Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 429 -Headers @{ 'Retry-After' = '60' })
            { Invoke-IQApi -Method GET -Path 'groups' } | Should -Throw -ExpectedMessage '*Time budget*'
            $script:Sleeps.Count | Should -Be 0
            $script:Requests.Count | Should -Be 1
            $script:IQ.BudgetExceeded = $false
            Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 503)
            { Invoke-IQApi -Method GET -Path 'groups' } | Should -Throw -ExpectedMessage '*Time budget*503*'
            $script:Sleeps.Count | Should -Be 0
        }
        finally {
            $script:IQ.Options.Remove('TimeBudgetMinutes')
            $script:IQ.StartedUtc = [datetime]::UtcNow
            $script:IQ.BudgetExceeded = $false
        }
    }
    It 'caps a long Retry-After to the seconds left in the time budget' {
        $script:IQ.Options['TimeBudgetMinutes'] = 10
        $script:IQ.StartedUtc = [datetime]::UtcNow.AddMinutes(-7.5)   # 8 min usable -> about 30 s left
        $script:IQ.BudgetExceeded = $false
        try {
            Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 429 -Headers @{ 'Retry-After' = '200' })
            Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @() })
            Invoke-IQApi -Method GET -Path 'groups' | Out-Null
            $script:Sleeps.Count | Should -Be 1
            $script:Sleeps[0] | Should -BeLessOrEqual 30
            $script:Sleeps[0] | Should -BeGreaterOrEqual 1
        }
        finally {
            $script:IQ.Options.Remove('TimeBudgetMinutes')
            $script:IQ.StartedUtc = [datetime]::UtcNow
            $script:IQ.BudgetExceeded = $false
        }
    }
    It '503: exponential backoff 2,4,... then success' {
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 503 -Body 'Service Unavailable')
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 502)
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @(@{ id = 1 }) })
        $res = Invoke-IQApi -Method GET -Path 'groups'
        @($res.value).Count | Should -Be 1
        $script:Sleeps.Count | Should -Be 2
        $script:Sleeps[0] | Should -Be 2
        $script:Sleeps[1] | Should -Be 4
    }
    It '5xx: throws after $IQ.Options.MaxRetries attempts' {
        foreach ($i in 1..3) { Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 500 -Body 'boom') }
        { Invoke-IQApi -Method GET -Path 'groups' } | Should -Throw -ExpectedMessage '*500*'
        $script:Requests.Count | Should -Be 3
    }
    It 'network timeouts (WebException Timeout) are retried' {
        Add-HttpResponse (New-IQTestWebException -Status 'Timeout')
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ value = @() })
        { Invoke-IQApi -Method GET -Path 'groups' } | Should -Not -Throw
        $script:Sleeps.Count | Should -Be 1
    }
    It 'certificate failures (WebException TrustFailure / SecureChannelFailure) are not retried' {
        Add-HttpResponse (New-IQTestWebException -Status 'TrustFailure')
        { Invoke-IQApi -Method GET -Path 'groups' } | Should -Throw
        $script:Sleeps.Count | Should -Be 0
        $script:Requests.Count | Should -Be 1
        Add-HttpResponse (New-IQTestWebException -Status 'SecureChannelFailure')
        { Invoke-IQApi -Method GET -Path 'groups' } | Should -Throw
        $script:Sleeps.Count | Should -Be 0
    }
    It 'a non-JSON body (proxy HTML page) returns $null with a Warn instead of a raw string' {
        Add-HttpResponse (New-IQTestHttpResponse -Content '<html><body>Access denied by proxy</body></html>' -Headers @{ 'Content-Type' = 'text/html' })
        $res = Invoke-IQApi -Method GET -Path 'groups'
        $res | Should -BeNullOrEmpty
        (Get-Content -LiteralPath $script:IQ.LogFile -Raw) | Should -Match '\[WARN\].*not JSON.*text/html.*Access denied'
    }
    It 'other 4xx (409) throws without retry' {
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 409 -Body 'Conflict')
        { Invoke-IQApi -Method GET -Path 'groups' } | Should -Throw
        $script:Requests.Count | Should -Be 1
    }
}

Describe 'Get-IQHttpErrorInfo' {
    It 'reads a synthetic WebException (timeout, no response) as transient' {
        $record = New-IQTestWebException -Status 'Timeout'
        $info = $null
        try { throw $record } catch { $info = Get-IQHttpErrorInfo -ErrorRecord $_ }
        $info.StatusCode | Should -BeNullOrEmpty
        $info.Transient | Should -BeTrue
        $info.WebStatus | Should -Be 'Timeout'
        $info.Message | Should -Not -BeNullOrEmpty
    }
    It 'reads an HttpResponseException (PowerShell 7) with status, Retry-After and body' -Skip:($null -eq ('Microsoft.PowerShell.Commands.HttpResponseException' -as [type])) {
        $msg = New-Object System.Net.Http.HttpResponseMessage ([System.Net.HttpStatusCode]::ServiceUnavailable)
        $msg.Headers.TryAddWithoutValidation('Retry-After', '12') | Out-Null
        $ex = New-Object Microsoft.PowerShell.Commands.HttpResponseException ('503', $msg)
        $record = New-Object System.Management.Automation.ErrorRecord ($ex, 'x', 'InvalidOperation', $null)
        $record.ErrorDetails = New-Object System.Management.Automation.ErrorDetails ('{"error":"server_error"}')
        $info = $null
        try { throw $record } catch { $info = Get-IQHttpErrorInfo -ErrorRecord $_ }
        $info.StatusCode | Should -Be 503
        $info.RetryAfter | Should -Be 12
        $info.Body | Should -Be '{"error":"server_error"}'
    }
    It 'reads a Windows PowerShell style response object (Response.StatusCode / Headers) and an HTTP-date Retry-After' {
        $ex = New-Object System.Exception ('The remote server returned an error: (429).')
        $future = [datetime]::UtcNow.AddSeconds(90).ToString('r')
        $ex | Add-Member -MemberType NoteProperty -Name 'Response' -Value ([pscustomobject]@{ StatusCode = 429; Headers = @{ 'Retry-After' = $future } })
        $record = New-Object System.Management.Automation.ErrorRecord ($ex, 'x', 'InvalidOperation', $null)
        $record.ErrorDetails = New-Object System.Management.Automation.ErrorDetails ('throttled')
        $info = $null
        try { throw $record } catch { $info = Get-IQHttpErrorInfo -ErrorRecord $_ }
        $info.StatusCode | Should -Be 429
        $info.RetryAfter | Should -BeGreaterThan 60
        $info.RetryAfter | Should -BeLessOrEqual 91
        $info.Body | Should -Be 'throttled'
    }
    It 'reads a WebException TrustFailure as permanent (not transient)' {
        $record = New-IQTestWebException -Status 'TrustFailure'
        $info = $null
        try { throw $record } catch { $info = Get-IQHttpErrorInfo -ErrorRecord $_ }
        $info.StatusCode | Should -BeNullOrEmpty
        $info.Transient | Should -BeFalse
        $info.WebStatus | Should -Be 'TrustFailure'
    }
    It 'reads a WebException ConnectFailure as transient' {
        $record = New-IQTestWebException -Status 'ConnectFailure'
        $info = $null
        try { throw $record } catch { $info = Get-IQHttpErrorInfo -ErrorRecord $_ }
        $info.Transient | Should -BeTrue
    }
    It 'returns a Message for a plain exception without response' {
        $record = New-Object System.Management.Automation.ErrorRecord ((New-Object System.Exception 'plain failure'), 'x', 'InvalidOperation', $null)
        $info = Get-IQHttpErrorInfo -ErrorRecord $record
        $info.StatusCode | Should -BeNullOrEmpty
        $info.Message | Should -Be 'plain failure'
    }
}

Describe 'Invoke-IQFabricLro' {
    BeforeEach { Reset-Http }
    It 'returns the body directly on 200' {
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ definition = @{ parts = @() } })
        $r = Invoke-IQFabricLro -Method POST -Path 'workspaces/w/dataflows/d/getDefinition'
        $r.definition | Should -Not -BeNullOrEmpty
        $script:Requests[0].Method | Should -Be 'POST'
    }
    It 'polls Location until Succeeded and GETs <operation>/result on 202' {
        Add-HttpResponse (New-IQTestHttpResponse -StatusCode 202 -Content '' -Headers @{ 'Location' = 'https://api.fabric.microsoft.us/v1/operations/op-1'; 'Retry-After' = '1'; 'x-ms-operation-id' = 'op-1' })
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ status = 'Running' })
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ status = 'Succeeded' })
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ definition = @{ parts = @(@{ path = 'mashup.pq' }) } })
        $r = Invoke-IQFabricLro -Method POST -Path 'workspaces/w/dataflows/d/getDefinition' -TimeoutMinutes 1
        @($r.definition.parts).Count | Should -Be 1
        $script:Requests.Count | Should -Be 4
        $script:Requests[1].Uri | Should -Be 'https://api.fabric.microsoft.us/v1/operations/op-1'
        $script:Requests[3].Uri | Should -Be 'https://api.fabric.microsoft.us/v1/operations/op-1/result'
        $script:Sleeps.Count | Should -Be 2
    }
    It 'returns $null when the operation fails' {
        Add-HttpResponse (New-IQTestHttpResponse -StatusCode 202 -Content '' -Headers @{ 'Location' = 'https://api.fabric.microsoft.us/v1/operations/op-2'; 'Retry-After' = '1' })
        Add-HttpResponse (New-IQTestHttpResponse -Content @{ status = 'Failed'; error = @{ errorCode = 'X' } })
        Invoke-IQFabricLro -Method POST -Path 'workspaces/w/dataflows/d/getDefinition' -TimeoutMinutes 1 | Should -BeNullOrEmpty
    }
    It 'returns $null on 403/404 with the status in the single Warn line' {
        Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 404 -Body '{"error":{"code":"ItemNotFound"}}')
        Invoke-IQFabricLro -Method POST -Path 'workspaces/w/dataflows/d/getDefinition' | Should -BeNullOrEmpty
        $log = Get-Content -LiteralPath $script:IQ.LogFile -Raw
        $log | Should -Match '\[WARN\].*HTTP 404 for POST.*getDefinition.*ItemNotFound'
        $log | Should -Not -Match 'was not accepted'
    }
    It 'returns $null with one Debug line and no request when no Fabric token is available' {
        Mock Get-IQToken { if ($Resource -eq 'Fabric') { return $null } return 'pbi-token' }
        Invoke-IQFabricLro -Method POST -Path 'workspaces/w/semanticModels/m/getDefinition' | Should -BeNullOrEmpty
        $script:Requests.Count | Should -Be 0
        (Get-Content -LiteralPath $script:IQ.LogFile -Raw) | Should -Not -Match '\[WARN\].*semanticModels/m/getDefinition'
    }
    It 'stops polling with $null when the run time budget is used up' {
        Add-HttpResponse (New-IQTestHttpResponse -StatusCode 202 -Content '' -Headers @{ 'Location' = 'https://api.fabric.microsoft.us/v1/operations/op-3'; 'Retry-After' = '1' })
        $script:IQ.Options['TimeBudgetMinutes'] = 5
        $script:IQ.StartedUtc = [datetime]::UtcNow.AddMinutes(-10)
        $script:IQ.BudgetExceeded = $false
        try {
            Invoke-IQFabricLro -Method POST -Path 'workspaces/w/dataflows/d/getDefinition' -TimeoutMinutes 1 | Should -BeNullOrEmpty
            $script:Requests.Count | Should -Be 1
            $script:Sleeps.Count | Should -Be 0
            (Get-Content -LiteralPath $script:IQ.LogFile -Raw) | Should -Match 'Time budget reached while polling'
        }
        finally {
            $script:IQ.Options.Remove('TimeBudgetMinutes')
            $script:IQ.StartedUtc = [datetime]::UtcNow
            $script:IQ.BudgetExceeded = $false
        }
    }
}

Describe 'Downloads (-OutFile / Invoke-IQDownload)' {
    BeforeEach { Reset-Http }
    It 'streams to disk and returns $true only for a non-empty file' {
        $target = Join-Path $script:Base 'dl/report.pbix'
        Add-HttpResponse (New-IQTestHttpResponse -Content 'PBIX-BYTES' -Headers @{ 'Content-Type' = 'application/octet-stream' })
        Invoke-IQApi -Method GET -Path 'groups/w/reports/r/Export' -OutFile $target | Should -BeTrue
        $target | Should -Exist
    }
    It 'streams with -OutFile only (no -PassThru, which would buffer the whole download in memory on 5.1)' {
        $target = Join-Path $script:Base 'dl/big.pbix'
        Add-HttpResponse (New-IQTestHttpResponse -Content 'PBIX-BYTES' -Headers @{ 'Content-Type' = 'application/octet-stream' })
        Invoke-IQApi -Method GET -Path 'groups/w/reports/r/Export' -OutFile $target | Should -BeTrue
        Should -Invoke Invoke-WebRequest -Times 1 -Exactly -ParameterFilter { $OutFile -eq $target -and -not $PassThru }
    }
    It 'deletes a zero-byte file and returns $false' {
        $target = Join-Path $script:Base 'dl/empty.pbix'
        Add-HttpResponse (New-IQTestHttpResponse -Content '' -Headers @{ 'Content-Type' = 'application/octet-stream' })
        Invoke-IQApi -Method GET -Path 'groups/w/reports/r/Export' -OutFile $target | Should -BeFalse
        $target | Should -Not -Exist
    }
    It 'Invoke-IQDownload never throws and returns $false on a hard failure' {
        $target = Join-Path $script:Base 'dl/fail.pbix'
        foreach ($i in 1..3) { Add-HttpResponse (New-IQTestHttpErrorRecord -StatusCode 500) }
        $ok = $true
        { $script:DlResult = Invoke-IQDownload -Url 'groups/w/reports/r/Export' -OutFile $target } | Should -Not -Throw
        $script:DlResult | Should -BeFalse
        $target | Should -Not -Exist
        $ok | Should -BeTrue
    }
    It 'Invoke-IQDownload returns $true for a good download' {
        $target = Join-Path $script:Base 'dl/good.pbix'
        Add-HttpResponse (New-IQTestHttpResponse -Content 'DATA' -Headers @{ 'Content-Type' = 'application/octet-stream' })
        Invoke-IQDownload -Url 'https://api.powerbigov.us/v1.0/myorg/groups/w/reports/r/Export' -OutFile $target | Should -BeTrue
    }
}
