# Auth.Tests.ps1 - ImpactIQ.Auth.ps1 (brief section 2.2, 4): JWT expiry, token cache, Auto resolution, device code, ROPC.
# Discovery-time values (Pester evaluates -Skip / -TestCases before BeforeAll runs).
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
$script:OnWindows = ($env:OS -eq 'Windows_NT')
$script:HasResolver = [bool](Select-String -Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'Config/Modules/ImpactIQ.Auth.ps1') -Pattern '^function Resolve-IQAuthMode' -Quiet -ErrorAction SilentlyContinue)

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:OnWindows = ($env:OS -eq 'Windows_NT')
    Reset-IQTestEnvironment
    $script:Base = Initialize-IQTestContext -Options @{ Environment = 'USGov' } -NoRun -Prefix 'auth'
    $script:CachePath = Join-Path $script:Base 'State/auth/token-cache.json'

    # Response queue for the mocked Invoke-RestMethod. Keys: devicecode | token | webhook.
    # An item can be an object (returned), a string (thrown, e.g. an OAuth JSON error body) or a scriptblock.
    $script:AuthQueue = @{}
    $script:AuthCalls = New-Object System.Collections.Generic.List[object]
    $script:AuthSleeps = New-Object System.Collections.Generic.List[int]
    function Set-AuthQueue {
        param([string]$Key, [object[]]$Items)
        $list = New-Object System.Collections.Generic.List[object]
        foreach ($i in $Items) { $list.Add($i) }
        $script:AuthQueue[$Key] = $list
    }
    function Pop-AuthResponse {
        param($Uri, $Body, $ContentType)
        $key = 'token'
        if ("$Uri" -like '*/devicecode*') { $key = 'devicecode' } elseif ("$Uri" -like '*hooks*') { $key = 'webhook' }
        $script:AuthCalls.Add(@{ Key = $key; Uri = "$Uri"; Body = $Body; ContentType = $ContentType })
        if (-not $script:AuthQueue.ContainsKey($key) -or $script:AuthQueue[$key].Count -eq 0) { throw "test mock: no response queued for $key ($Uri)" }
        $item = $script:AuthQueue[$key][0]
        $script:AuthQueue[$key].RemoveAt(0)
        if ($item -is [scriptblock]) { return (& $item) }
        if ($item -is [string]) { throw $item }
        return $item
    }
    function Reset-AuthMock {
        $script:AuthQueue = @{}
        $script:AuthCalls.Clear()
        $script:AuthSleeps.Clear()
    }
}
AfterAll {
    Reset-IQTestEnvironment
    Remove-IQTestFolder -Path $script:Base
}

Describe 'Get-IQJwtExpiry' {
    It 'parses the exp claim as a UTC datetime' {
        $exp = [datetime]::UtcNow.AddHours(1)
        $token = New-IQTestJwt -Claims @{ exp = (ConvertTo-IQTestEpoch -Date $exp); upn = 'user@contoso.gov' }
        $got = Get-IQJwtExpiry -Token $token
        [math]::Abs(($got - $exp).TotalSeconds) | Should -BeLessThan 2
        $got.Kind | Should -Be 'Utc'
    }
    It 'handles base64url payloads of every padding length (<Pad> filler chars)' -TestCases @(@{ Pad = 1 }, @{ Pad = 2 }, @{ Pad = 3 }, @{ Pad = 4 }) {
        param($Pad)
        $token = New-IQTestJwt -Claims @{ exp = (ConvertTo-IQTestEpoch -Date ([datetime]::UtcNow.AddMinutes(30))); x = ('y' * $Pad) }
        Get-IQJwtExpiry -Token $token | Should -Not -Be ([datetime]::MaxValue)
    }
    It 'tolerates a "Bearer " prefix' {
        $token = New-IQTestJwt -ExpiresInMinutes 20
        Get-IQJwtExpiry -Token ('Bearer ' + $token) | Should -Be (Get-IQJwtExpiry -Token $token)
    }
    It 'returns [datetime]::MaxValue for <Label>' -TestCases @(
        @{ Label = 'garbage'; Token = 'garbage' }
        @{ Label = 'empty'; Token = '' }
        @{ Label = 'bad base64'; Token = 'a.!!!.c' }
        @{ Label = 'no exp claim'; Token = 'NOEXP' }
    ) {
        param($Token)
        if ($Token -eq 'NOEXP') { $Token = New-IQTestJwt -Claims @{ exp = $null; upn = 'x' } }
        Get-IQJwtExpiry -Token $Token | Should -Be ([datetime]::MaxValue)
    }
}

Describe 'Token cache (brief section 4.3)' {
    BeforeEach {
        if (Test-Path -LiteralPath $script:CachePath) { Remove-Item -LiteralPath $script:CachePath -Force }
        $script:IQ.Auth = @{ Mode = 'DeviceCode'; TokenCachePath = $script:CachePath; TokenCacheKey = 'unit-test-key'; RefreshToken = 'RT-secret-1'; Authority = 'https://login.microsoftonline.com'; ClientId = '1950a258-227b-4e31-a9cf-717495945fc2'; Environment = 'USGov'; TenantId = 'organizations'; Account = 'svc@contoso.gov'; CacheWarned = $false; Initialized = $true }
    }
    It 'round-trips through AES-256 with a key and never stores the refresh token in clear text' {
        Save-IQTokenCache | Should -BeTrue
        $script:CachePath | Should -Exist
        $file = Get-Content -LiteralPath $script:CachePath -Raw | ConvertFrom-Json
        $file.format | Should -Be 'aes256'
        (Get-Content -LiteralPath $script:CachePath -Raw) | Should -Not -Match 'RT-secret-1'
        $restored = Restore-IQTokenCache
        $restored.refreshToken | Should -Be 'RT-secret-1'
        $restored.account | Should -Be 'svc@contoso.gov'
        $restored.environment | Should -Be 'USGov'
        [int]$restored.schemaVersion | Should -Be 1
    }
    It 'returns $null with the wrong key' {
        Save-IQTokenCache | Out-Null
        $script:IQ.Auth.TokenCacheKey = 'WRONG'
        Restore-IQTokenCache | Should -BeNullOrEmpty
    }
    It 'returns $null when the cache belongs to another environment or client' {
        Save-IQTokenCache | Out-Null
        $script:IQ.Auth.Environment = 'Public'
        Restore-IQTokenCache | Should -BeNullOrEmpty
        $script:IQ.Auth.Environment = 'USGov'
        $script:IQ.Auth.ClientId = '04b07795-8ddb-461a-bbee-02f9e1bf7b46'
        Restore-IQTokenCache | Should -BeNullOrEmpty
    }
    It 'tolerates a different tenant id (refresh tokens are not tenant bound)' {
        Save-IQTokenCache | Out-Null
        $script:IQ.Auth.TenantId = 'aaaaaaaa-1111-4111-8111-111111111111'
        (Restore-IQTokenCache).refreshToken | Should -Be 'RT-secret-1'
    }
    It 'returns $null for corrupt or non-JSON content' {
        Set-Content -LiteralPath $script:CachePath -Value '{"format":"aes256","data":"!!!not-base64!!!"}'
        Restore-IQTokenCache | Should -BeNullOrEmpty
        Set-Content -LiteralPath $script:CachePath -Value 'not json at all'
        Restore-IQTokenCache | Should -BeNullOrEmpty
    }
    It 'returns $null when the file is missing' {
        Restore-IQTokenCache | Should -BeNullOrEmpty
    }
    It 'does not persist without a key on non-Windows hosts (warns once)' -Skip:($script:OnWindows) {
        $script:IQ.Auth.TokenCacheKey = $null
        Save-IQTokenCache | Should -BeFalse
        $script:CachePath | Should -Not -Exist
        (Get-Content -LiteralPath $script:IQ.LogFile -Raw) | Should -Match '(?i)not persisted'
    }
    It 'uses DPAPI without a key on Windows' -Skip:(-not $script:OnWindows) {
        $script:IQ.Auth.TokenCacheKey = $null
        Save-IQTokenCache | Should -BeTrue
        (Get-Content -LiteralPath $script:CachePath -Raw | ConvertFrom-Json).format | Should -Be 'dpapi-user'
        (Restore-IQTokenCache).refreshToken | Should -Be 'RT-secret-1'
    }
}

Describe 'Auto mode resolution (brief section 4.1)' {
    BeforeAll {
        $script:Cred = New-Object System.Management.Automation.PSCredential ('user@contoso.gov', (ConvertTo-SecureString -String 'p@ss' -AsPlainText -Force))
    }
    BeforeEach {
        Reset-IQTestEnvironment
        if (Test-Path -LiteralPath $script:CachePath) { Remove-Item -LiteralPath $script:CachePath -Force }
        $script:IQ.Interactive = $false
    }
    AfterEach { Reset-IQTestEnvironment }
    It 'prefers Credential when -Credential is given' -Skip:(-not $script:HasResolver) {
        Resolve-IQAuthMode -Mode Auto -Credential $script:Cred | Should -Be 'Credential'
    }
    It 'prefers Credential when IMPACTIQ_USERNAME and IMPACTIQ_PASSWORD are set' -Skip:(-not $script:HasResolver) {
        $env:IMPACTIQ_USERNAME = 'u@contoso.gov'; $env:IMPACTIQ_PASSWORD = 'p'
        Resolve-IQAuthMode -Mode Auto | Should -Be 'Credential'
    }
    It 'uses AccessToken when IMPACTIQ_PBI_TOKEN is set' -Skip:(-not $script:HasResolver) {
        $env:IMPACTIQ_PBI_TOKEN = 'tok'
        Resolve-IQAuthMode -Mode Auto | Should -Be 'AccessToken'
    }
    It 'uses DeviceCode when a token cache exists, even when interactive' -Skip:(-not $script:HasResolver) {
        New-Item -ItemType Directory -Path (Split-Path -Parent $script:CachePath) -Force | Out-Null
        Set-Content -LiteralPath $script:CachePath -Value '{}'
        Resolve-IQAuthMode -Mode Auto -TokenCachePath $script:CachePath -Interactive $true | Should -Be 'DeviceCode'
    }
    It 'uses Interactive when interactive and nothing else applies' -Skip:(-not $script:HasResolver) {
        Resolve-IQAuthMode -Mode Auto -Interactive $true | Should -Be 'Interactive'
    }
    It 'falls back to DeviceCode when headless' -Skip:(-not $script:HasResolver) {
        Resolve-IQAuthMode -Mode Auto -Interactive $false | Should -Be 'DeviceCode'
    }
    It 'never auto-selects AzContext but honours an explicit request' -Skip:(-not $script:HasResolver) {
        Resolve-IQAuthMode -Mode AzContext | Should -Be 'AzContext'
        Resolve-IQAuthMode -Mode Auto -Interactive $false | Should -Not -Be 'AzContext'
    }
    It 'Initialize-IQAuth -Mode Auto resolves AccessToken from the environment variable' {
        $env:IMPACTIQ_PBI_TOKEN = New-IQTestJwt -Claims @{ upn = 'static@contoso.gov' } -ExpiresInMinutes 45
        Initialize-IQAuth -Mode Auto -Environment 'USGov' | Should -Be 'AccessToken'
        Get-IQToken -Resource PowerBI | Should -Be $env:IMPACTIQ_PBI_TOKEN
    }
}

Describe 'DeviceCode flow with a mocked token endpoint' {
    BeforeAll {
        Reset-IQTestEnvironment
        Reset-AuthMock
        Mock Invoke-RestMethod { Pop-AuthResponse -Uri $Uri -Body $Body -ContentType $ContentType }
        Mock Start-IQAuthSleep { $script:AuthSleeps.Add([int]$Seconds) }
        if (Test-Path -LiteralPath $script:CachePath) { Remove-Item -LiteralPath $script:CachePath -Force }
        $script:PbiJwt = New-IQTestJwt -Claims @{ upn = 'svc@contoso.gov'; tid = 'aaaaaaaa-1111-4111-8111-111111111111' } -ExpiresInMinutes 70
        $script:IdJwt = New-IQTestJwt -Claims @{ preferred_username = 'svc@contoso.gov'; tid = 'aaaaaaaa-1111-4111-8111-111111111111' }
        Set-AuthQueue 'devicecode' @([pscustomobject]@{ device_code = 'DEV-CODE-1'; user_code = 'ABCD1234'; verification_uri = 'https://microsoft.com/devicelogin'; expires_in = 900; interval = 5; message = 'To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code ABCD1234 to authenticate.' })
        Set-AuthQueue 'token' @(
            '{"error":"authorization_pending","error_description":"AADSTS70016: Pending end-user authorization."}',
            '{"error":"slow_down","error_description":"AADSTS70016: slow down"}',
            '{"error":"authorization_pending","error_description":"AADSTS70016: Pending end-user authorization."}',
            [pscustomobject]@{ token_type = 'Bearer'; scope = 'x'; expires_in = 4200; access_token = $script:PbiJwt; refresh_token = 'RT-1'; id_token = $script:IdJwt }
        )
        Set-AuthQueue 'webhook' @([pscustomobject]@{ ok = 1 })
        $script:Mode = Initialize-IQAuth -Mode DeviceCode -Environment 'USGov' -TokenCacheKey 'unit-test-key' -TokenCachePath $script:CachePath -DeviceCodeWebhookUrl 'https://contoso.webhook.office.com/webhookb2/hooks/abc'
    }
    It 'resolves to DeviceCode and signs in after authorization_pending -> slow_down -> success' {
        $script:Mode | Should -Be 'DeviceCode'
        @($script:AuthCalls | Where-Object { $_.Key -eq 'token' }).Count | Should -Be 4
        Get-IQToken -Resource PowerBI | Should -Be $script:PbiJwt
    }
    It 'posts the device-code request with the v2 scope for the environment' {
        $dc = @($script:AuthCalls | Where-Object { $_.Key -eq 'devicecode' })[0]
        $dc.Uri | Should -Be 'https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode'
        $dc.Body.client_id | Should -Be '1950a258-227b-4e31-a9cf-717495945fc2'
        $dc.Body.scope | Should -Match '^https://analysis\.usgovcloudapi\.net/powerbi/api/\.default offline_access'
        $dc.ContentType | Should -Be 'application/x-www-form-urlencoded'
    }
    It 'polls with the device_code grant and honours interval / slow_down (+5 s)' {
        $poll = @($script:AuthCalls | Where-Object { $_.Key -eq 'token' })[0]
        $poll.Body.grant_type | Should -Be 'urn:ietf:params:oauth:grant-type:device_code'
        $poll.Body.device_code | Should -Be 'DEV-CODE-1'
        $script:AuthSleeps.Count | Should -BeGreaterOrEqual 3
        $script:AuthSleeps[0] | Should -Be 5
        ($script:AuthSleeps | Select-Object -Last 1) | Should -BeGreaterOrEqual 10 -Because 'slow_down adds 5 s to the interval'
    }
    It 'prints the device-code message at Warn and posts it to the webhook' {
        (Get-Content -LiteralPath $script:IQ.LogFile -Raw) | Should -Match '\[WARN\].*ABCD1234'
        $wh = @($script:AuthCalls | Where-Object { $_.Key -eq 'webhook' })
        $wh.Count | Should -Be 1
        $text = $wh[0].Body
        if ($text -is [byte[]]) { $text = [System.Text.Encoding]::UTF8.GetString($text) }
        [string]$text | Should -Match 'ABCD1234'
    }
    It 'stores the refresh token, persists the cache (AES) and reports the account' {
        $script:IQ.Auth.RefreshToken | Should -Be 'RT-1'
        $script:CachePath | Should -Exist
        (Get-Content -LiteralPath $script:CachePath -Raw | ConvertFrom-Json).format | Should -Be 'aes256'
        Get-IQAuthDescription | Should -Match 'DeviceCode'
        Get-IQAuthDescription | Should -Match 'svc@contoso.gov'
        # The very first cache file (written during the initial sign-in) must already carry the account (brief 4.3).
        (Restore-IQTokenCache).account | Should -Be 'svc@contoso.gov'
    }
    It 'applies the process HTTP defaults (proxy credentials) before the first OAuth call' {
        $script:IQAuthHttpDefaultsApplied | Should -BeTrue
        $proxy = [System.Net.WebRequest]::DefaultWebProxy
        if ($null -ne $proxy) { $proxy.Credentials | Should -Not -BeNullOrEmpty }
    }
    It 'refreshes proactively when less than 5 minutes remain and rotates the refresh token' {
        Reset-AuthMock
        $script:IQ.Auth.Tokens['PowerBI'].ExpiresUtc = [datetime]::UtcNow.AddMinutes(3)
        $newJwt = New-IQTestJwt -Claims @{ upn = 'svc@contoso.gov' } -ExpiresInMinutes 65
        Set-AuthQueue 'token' @([pscustomobject]@{ access_token = $newJwt; refresh_token = 'RT-2'; expires_in = 3900 })
        Get-IQToken -Resource PowerBI | Should -Be $newJwt
        $script:AuthCalls.Count | Should -Be 1
        $script:AuthCalls[0].Body.grant_type | Should -Be 'refresh_token'
        $script:AuthCalls[0].Body.refresh_token | Should -Be 'RT-1'
        $script:IQ.Auth.RefreshToken | Should -Be 'RT-2'
        (Restore-IQTokenCache).refreshToken | Should -Be 'RT-2'
    }
    It 'redeems the refresh token for Fabric and degrades to $null when that fails' {
        Reset-AuthMock
        $fabJwt = New-IQTestJwt -Claims @{ aud = 'fabric' } -ExpiresInMinutes 60
        Set-AuthQueue 'token' @([pscustomobject]@{ access_token = $fabJwt; refresh_token = 'RT-3'; expires_in = 3600 })
        Get-IQToken -Resource Fabric | Should -Be $fabJwt
        $script:AuthCalls[0].Body.scope | Should -Match '^https://api\.fabric\..*/\.default'
        Reset-AuthMock
        $script:IQ.Auth.Tokens.Remove('Fabric')
        Set-AuthQueue 'token' @('{"error":"invalid_scope","error_description":"AADSTS65001: The user or administrator has not consented"}')
        Get-IQToken -Resource Fabric | Should -BeNullOrEmpty
        { Get-IQToken -Resource Fabric } | Should -Not -Throw
    }
    It 'throws on expired_token and authorization_declined' {
        foreach ($case in @('expired_token', 'authorization_declined')) {
            Reset-AuthMock
            Set-AuthQueue 'devicecode' @([pscustomobject]@{ device_code = 'DEV-2'; user_code = 'Q'; verification_uri = 'https://microsoft.com/devicelogin'; expires_in = 60; interval = 1; message = 'code Q' })
            Set-AuthQueue 'token' @(('{"error":"' + $case + '","error_description":"' + $case + '"}'))
            { Invoke-IQDeviceCodeFlow } | Should -Throw
        }
    }
    It 'never writes secrets to the log file' {
        $log = Get-Content -LiteralPath $script:IQ.LogFile -Raw
        $log | Should -Not -Match 'RT-1'
        $log | Should -Not -Match 'RT-2'
        $log | Should -Not -Match 'unit-test-key'
        $log | Should -Not -Match ([regex]::Escape($script:PbiJwt))
    }
    It 'signs in silently from the cache in a fresh context (no device code prompt)' {
        Reset-AuthMock
        Initialize-IQContext -BaseFolder $script:Base -Options @{ Environment = 'USGov'; NonInteractive = $true } | Out-Null
        $script:IQ.Interactive = $false
        $jwt = New-IQTestJwt -Claims @{ upn = 'svc@contoso.gov' } -ExpiresInMinutes 65
        Set-AuthQueue 'token' @([pscustomobject]@{ access_token = $jwt; refresh_token = 'RT-9' })
        Initialize-IQAuth -Mode Auto -Environment 'USGov' -TokenCacheKey 'unit-test-key' -TokenCachePath $script:CachePath | Should -Be 'DeviceCode'
        @($script:AuthCalls | Where-Object { $_.Key -eq 'devicecode' }).Count | Should -Be 0
        $script:AuthCalls[0].Body.grant_type | Should -Be 'refresh_token'
        Get-IQAuthDescription | Should -Match '(?i)cached'
    }
    It 'keeps polling through a transient token-endpoint failure while the device code is still valid' {
        Reset-AuthMock
        $jwt = New-IQTestJwt -Claims @{ upn = 'svc@contoso.gov' } -ExpiresInMinutes 60
        Set-AuthQueue 'devicecode' @([pscustomobject]@{ device_code = 'DEV-3'; user_code = 'T'; verification_uri = 'https://microsoft.com/devicelogin'; expires_in = 900; interval = 5; message = 'code T' })
        $blip = { throw (New-IQTestWebException -Status Timeout) }
        Set-AuthQueue 'token' @(
            $blip, $blip, $blip,                                                     # one poll: 3 transient attempts, all fail
            '{"error":"authorization_pending","error_description":"AADSTS70016"}',
            [pscustomobject]@{ access_token = $jwt; refresh_token = 'RT-T'; expires_in = 3600 }
        )
        $response = Invoke-IQDeviceCodeFlow
        $response.access_token | Should -Be $jwt
        @($script:AuthCalls | Where-Object { $_.Key -eq 'token' }).Count | Should -Be 5
        (Get-Content -LiteralPath $script:IQ.LogFile -Raw) | Should -Match '\[WARN\].*(?i)poll failed transiently'
    }
    It 'gives up on a transient failure only once the device code has expired' {
        Reset-AuthMock
        Set-AuthQueue 'devicecode' @([pscustomobject]@{ device_code = 'DEV-4'; user_code = 'X'; verification_uri = 'https://microsoft.com/devicelogin'; expires_in = 1; interval = 1; message = 'code X' })
        $blip = { throw (New-IQTestWebException -Status ConnectFailure) }
        $slowBlip = { Start-Sleep -Milliseconds 1300; throw (New-IQTestWebException -Status ConnectFailure) }
        Set-AuthQueue 'token' @($blip, $blip, $blip, $slowBlip, $blip, $blip)
        { Invoke-IQDeviceCodeFlow } | Should -Throw -ExpectedMessage '*not completed before the code expired*'
        @($script:AuthCalls | Where-Object { $_.Key -eq 'token' }).Count | Should -Be 6
    }
    It 'survives a longer outage when refreshing mid-run (6 attempts, capped backoff)' {
        Reset-AuthMock
        $script:IQ.Auth.Tokens['PowerBI'].ExpiresUtc = [datetime]::UtcNow.AddMinutes(1)
        $newJwt = New-IQTestJwt -Claims @{ upn = 'svc@contoso.gov' } -ExpiresInMinutes 65
        $blip = { throw (New-IQTestWebException -Status NameResolutionFailure) }
        Set-AuthQueue 'token' @($blip, $blip, $blip, $blip, [pscustomobject]@{ access_token = $newJwt; refresh_token = 'RT-4'; expires_in = 3900 })
        Get-IQToken -Resource PowerBI | Should -Be $newJwt
        $script:AuthCalls.Count | Should -Be 5
        @($script:AuthSleeps) | Should -Be @(2, 4, 8, 16)
    }
    It 'honours Retry-After on a throttled token endpoint and does not treat 429 as permanent' {
        Reset-AuthMock
        Set-AuthQueue 'token' @(
            { throw (New-IQTestHttpErrorRecord -StatusCode 429 -Body '{"error":"temporarily_unavailable"}' -Headers @{ 'Retry-After' = '7' }) },
            { throw (New-IQTestHttpErrorRecord -StatusCode 429 -Headers @{ 'Retry-After' = '9999' }) }
        )
        $r = Invoke-IQAuthRequest -Uri 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token' -Body @{ grant_type = 'x' } -MaxAttempts 2
        $r.Ok | Should -BeFalse
        [int]$r.StatusCode | Should -Be 429
        @($script:AuthSleeps) | Should -Be @(7) -Because 'the first Retry-After is honoured; the last attempt does not sleep'
        Test-IQAuthPermanentFailure -Result $r | Should -BeFalse
        Test-IQAuthPermanentFailure -Result @{ StatusCode = 408; Error = $null } | Should -BeFalse
        Test-IQAuthPermanentFailure -Result @{ StatusCode = 400; Error = 'invalid_grant' } | Should -BeTrue
        Test-IQAuthPermanentFailure -Result @{ StatusCode = $null; Error = $null } | Should -BeFalse
    }
    It 'caps a huge Retry-After at 300 s' {
        Reset-AuthMock
        Set-AuthQueue 'token' @(
            { throw (New-IQTestHttpErrorRecord -StatusCode 503 -Headers @{ 'Retry-After' = '9999' }) },
            [pscustomobject]@{ access_token = 'x'; expires_in = 60 }
        )
        (Invoke-IQAuthRequest -Uri 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token' -Body @{ grant_type = 'x' } -MaxAttempts 2).Ok | Should -BeTrue
        @($script:AuthSleeps) | Should -Be @(300)
    }
    It 'Fabric tokens are not permanently disabled by a throttled refresh grant' {
        Reset-AuthMock
        if ($script:IQ.Auth.Tokens.ContainsKey('Fabric')) { $script:IQ.Auth.Tokens.Remove('Fabric') }
        $script:IQ.Auth.FabricUnavailable = $false
        $throttle = { throw (New-IQTestHttpErrorRecord -StatusCode 429 -Headers @{ 'Retry-After' = '1' }) }
        Set-AuthQueue 'token' @($throttle, $throttle, $throttle, $throttle, $throttle, $throttle)
        Get-IQToken -Resource Fabric | Should -BeNullOrEmpty
        $script:IQ.Auth.FabricUnavailable | Should -BeFalse
        Reset-AuthMock
        $fabJwt = New-IQTestJwt -Claims @{ aud = 'fabric' } -ExpiresInMinutes 60
        Set-AuthQueue 'token' @([pscustomobject]@{ access_token = $fabJwt; expires_in = 3600 })
        Get-IQToken -Resource Fabric | Should -Be $fabJwt
    }
}

Describe 'DeviceCode start-up with a cached refresh token' {
    BeforeAll {
        Reset-IQTestEnvironment
        Reset-AuthMock
        Mock Invoke-RestMethod { Pop-AuthResponse -Uri $Uri -Body $Body -ContentType $ContentType }
        Mock Start-IQAuthSleep { $script:AuthSleeps.Add([int]$Seconds) }
        $script:StartCachePath = Join-Path $script:Base 'State/auth/startup-cache.json'
    }
    BeforeEach {
        Reset-AuthMock
        Initialize-IQContext -BaseFolder $script:Base -Options @{ Environment = 'USGov'; NonInteractive = $true } | Out-Null
        $script:IQ.Interactive = $false
        if (Test-Path -LiteralPath $script:StartCachePath) { Remove-Item -LiteralPath $script:StartCachePath -Force }
        $script:IQ.Auth = @{ Mode = 'DeviceCode'; TokenCachePath = $script:StartCachePath; TokenCacheKey = 'unit-test-key'; RefreshToken = 'RT-CACHED'; Authority = 'https://login.microsoftonline.com'; ClientId = '1950a258-227b-4e31-a9cf-717495945fc2'; Environment = 'USGov'; TenantId = 'organizations'; Account = 'svc@contoso.gov'; CacheWarned = $false; Initialized = $true }
        Save-IQTokenCache | Should -BeTrue
    }
    AfterAll { Reset-IQTestEnvironment }
    It 'throws and leaves the cache untouched when the token endpoint is unreachable (no device-code prompt)' {
        $blip = { throw (New-IQTestWebException -Status Timeout) }
        Set-AuthQueue 'token' @($blip, $blip, $blip, $blip, $blip, $blip)
        $before = Get-Content -LiteralPath $script:StartCachePath -Raw
        { Initialize-IQAuth -Mode DeviceCode -Environment 'USGov' -TokenCacheKey 'unit-test-key' -TokenCachePath $script:StartCachePath } | Should -Throw -ExpectedMessage '*left untouched*'
        @($script:AuthCalls | Where-Object { $_.Key -eq 'devicecode' }).Count | Should -Be 0
        @($script:AuthCalls | Where-Object { $_.Key -eq 'token' }).Count | Should -Be 6 -Because 'the start-up refresh gets the full retry budget'
        (Get-Content -LiteralPath $script:StartCachePath -Raw) | Should -Be $before
        $script:IQ.Auth.TokenCacheKey = 'unit-test-key'
        (Restore-IQTokenCache -Path $script:StartCachePath).refreshToken | Should -Be 'RT-CACHED'
    }
    It 'falls back to a fresh device-code sign-in only when the refresh token is permanently rejected' {
        $jwt = New-IQTestJwt -Claims @{ upn = 'svc@contoso.gov' } -ExpiresInMinutes 60
        Set-AuthQueue 'token' @(
            '{"error":"invalid_grant","error_description":"AADSTS70008: The refresh token has expired","error_codes":[70008]}',
            [pscustomobject]@{ access_token = $jwt; refresh_token = 'RT-NEW'; expires_in = 3600 }
        )
        Set-AuthQueue 'devicecode' @([pscustomobject]@{ device_code = 'DEV-5'; user_code = 'Z'; verification_uri = 'https://microsoft.com/devicelogin'; expires_in = 900; interval = 5; message = 'code Z' })
        Initialize-IQAuth -Mode DeviceCode -Environment 'USGov' -TokenCacheKey 'unit-test-key' -TokenCachePath $script:StartCachePath | Should -Be 'DeviceCode'
        @($script:AuthCalls | Where-Object { $_.Key -eq 'devicecode' }).Count | Should -Be 1
        $script:IQ.Auth.RefreshToken | Should -Be 'RT-NEW'
        (Restore-IQTokenCache -Path $script:StartCachePath).refreshToken | Should -Be 'RT-NEW'
    }
}

Describe 'Interactive mode: Fabric minting via Az.Accounts gives up instead of retrying on every call' {
    BeforeAll {
        Reset-IQTestEnvironment
        Initialize-IQContext -BaseFolder $script:Base -Options @{ Environment = 'USGov'; NonInteractive = $true } | Out-Null
        Set-IQEnvironment -Environment 'USGov' | Out-Null
        Mock Start-IQAuthSleep { $script:AuthSleeps.Add([int]$Seconds) }
        Mock Connect-IQAzForFabric { $true }
        $script:AzCalls = 0
    }
    BeforeEach {
        Reset-AuthMock
        $script:AzCalls = 0
        $script:IQ.Auth = @{ Initialized = $true; Mode = 'Interactive'; Provider = 'Module'; Tokens = @{}; FabricUnavailable = $false; FabricWarned = $false; FabricFailures = 0; TenantId = 'organizations'; Account = 'user@contoso.gov' }
    }
    AfterAll { Reset-IQTestEnvironment }
    It 'marks Fabric unavailable immediately when Get-AzAccessToken rejects the resource' {
        Mock Get-IQAzAccessTokenValue { $script:AzCalls++; throw 'Get-AzAccessToken: the resource https://api.fabric.microsoft.us is not supported for this environment' }
        Get-IQToken -Resource Fabric | Should -BeNullOrEmpty
        Get-IQToken -Resource Fabric | Should -BeNullOrEmpty
        Get-IQToken -Resource Fabric | Should -BeNullOrEmpty
        $script:AzCalls | Should -Be 1
        $script:IQ.Auth.FabricUnavailable | Should -BeTrue
        $log = Get-Content -LiteralPath $script:IQ.LogFile -Raw
        ([regex]::Matches($log, '\[WARN\].*Fabric token refresh via Az\.Accounts failed: Get-AzAccessToken: the resource')).Count | Should -Be 1
    }
    It 'gives a transient error a second chance, then stops retrying and stops warning' {
        Mock Get-IQAzAccessTokenValue { $script:AzCalls++; throw 'The operation has timed out' }
        Get-IQToken -Resource Fabric | Should -BeNullOrEmpty
        $script:IQ.Auth.FabricUnavailable | Should -BeFalse
        Get-IQToken -Resource Fabric | Should -BeNullOrEmpty
        $script:IQ.Auth.FabricUnavailable | Should -BeTrue
        Get-IQToken -Resource Fabric | Should -BeNullOrEmpty
        Get-IQToken -Resource Fabric | Should -BeNullOrEmpty
        $script:AzCalls | Should -Be 2
        $log = Get-Content -LiteralPath $script:IQ.LogFile -Raw
        ([regex]::Matches($log, '\[WARN\].*Fabric token refresh via Az\.Accounts failed: The operation has timed out')).Count | Should -Be 1
    }
    It 'resets the failure counter after a success' {
        $fabJwt = New-IQTestJwt -Claims @{ aud = 'fabric' } -ExpiresInMinutes 60
        $script:IQ.Auth.FabricFailures = 1
        Mock Get-IQAzAccessTokenValue { $script:AzCalls++; return $fabJwt }
        Get-IQToken -Resource Fabric | Should -Be $fabJwt
        $script:IQ.Auth.FabricFailures | Should -Be 0
    }
}

Describe 'Connect-PowerBIServiceAccount -Tenant is only passed when the user parameter set declares it' {
    BeforeAll {
        Reset-IQTestEnvironment
        Initialize-IQContext -BaseFolder $script:Base -Options @{ Environment = 'USGov'; NonInteractive = $true } | Out-Null
        Set-IQEnvironment -Environment 'USGov' | Out-Null
        Mock Start-IQAuthSleep { $script:AuthSleeps.Add([int]$Seconds) }
        Mock Import-IQAuthModule { $true }
        $script:PbiJwt2 = New-IQTestJwt -Claims @{ upn = 'user@contoso.gov' } -ExpiresInMinutes 60
        $script:ConnectCalls = New-Object System.Collections.Generic.List[object]
        function Get-PowerBIAccessToken { [CmdletBinding()] param() return $script:PbiJwt2 }
        function Disconnect-PowerBIServiceAccount { [CmdletBinding()] param() }
        # Older module build: -Tenant lives only in the service-principal parameter sets.
        $script:OldShape = {
            [CmdletBinding(DefaultParameterSetName = 'User')]
            param(
                [Parameter(ParameterSetName = 'User')][Parameter(ParameterSetName = 'UserAndCredential')][Parameter(ParameterSetName = 'ServicePrincipal')][string]$Environment,
                [Parameter(ParameterSetName = 'UserAndCredential')][System.Management.Automation.PSCredential]$Credential,
                [Parameter(ParameterSetName = 'ServicePrincipal')][switch]$ServicePrincipal,
                [Parameter(ParameterSetName = 'ServicePrincipal')][string]$Tenant
            )
            $script:ConnectCalls.Add(@{ Set = $PSCmdlet.ParameterSetName; Bound = @($PSBoundParameters.Keys) })
            if ($PSCmdlet.ParameterSetName -eq 'ServicePrincipal' -and -not $ServicePrincipal) { throw 'fake: -Tenant selected the ServicePrincipal parameter set without -ServicePrincipal' }
        }
        # Newer module build: -Tenant is available for user sign-in as well.
        $script:NewShape = {
            [CmdletBinding(DefaultParameterSetName = 'User')]
            param(
                [Parameter(ParameterSetName = 'User')][Parameter(ParameterSetName = 'UserAndCredential')][Parameter(ParameterSetName = 'ServicePrincipal')][string]$Environment,
                [Parameter(ParameterSetName = 'UserAndCredential')][Parameter(ParameterSetName = 'ServicePrincipal')][System.Management.Automation.PSCredential]$Credential,
                [Parameter(ParameterSetName = 'ServicePrincipal')][switch]$ServicePrincipal,
                [Parameter(ParameterSetName = 'User')][Parameter(ParameterSetName = 'UserAndCredential')][Parameter(ParameterSetName = 'ServicePrincipal')][string]$Tenant
            )
            $script:ConnectCalls.Add(@{ Set = $PSCmdlet.ParameterSetName; Bound = @($PSBoundParameters.Keys) })
        }
        # A build whose metadata advertises -Tenant for users but whose binder rejects it: retried once without -Tenant.
        $script:BindingErrorShape = {
            [CmdletBinding(DefaultParameterSetName = 'User')]
            param(
                [Parameter(ParameterSetName = 'User')][string]$Environment,
                [Parameter(ParameterSetName = 'User')][string]$Tenant
            )
            $script:ConnectCalls.Add(@{ Set = $PSCmdlet.ParameterSetName; Bound = @($PSBoundParameters.Keys) })
            if ($PSBoundParameters.ContainsKey('Tenant')) { throw (New-Object System.Management.Automation.ParameterBindingException 'Parameter set cannot be resolved using the specified named parameters.') }
        }
        $script:Cred2 = New-Object System.Management.Automation.PSCredential ('user@contoso.gov', (ConvertTo-SecureString -String 'p@ss' -AsPlainText -Force))
    }
    BeforeEach {
        Reset-AuthMock
        $script:ConnectCalls.Clear()
        $script:IQ.Auth = @{ Initialized = $false; Mode = 'Interactive'; TenantId = 'aaaaaaaa-1111-4111-8111-111111111111'; Tokens = @{} }
    }
    AfterEach { Remove-Item -Path 'function:Connect-PowerBIServiceAccount' -ErrorAction SilentlyContinue }
    AfterAll {
        Remove-Item -Path 'function:Get-PowerBIAccessToken', 'function:Disconnect-PowerBIServiceAccount' -ErrorAction SilentlyContinue
        Reset-IQTestEnvironment
    }
    It 'omits -Tenant when only the service-principal sets declare it (user sign-in)' {
        Set-Item -Path 'function:Connect-PowerBIServiceAccount' -Value $script:OldShape
        Connect-IQPowerBIModule -MaxAttempts 1 | Should -Be $script:PbiJwt2
        $script:ConnectCalls.Count | Should -Be 1
        $script:ConnectCalls[0].Set | Should -Be 'User'
        @($script:ConnectCalls[0].Bound) | Should -Not -Contain 'Tenant'
        @($script:ConnectCalls[0].Bound) | Should -Contain 'Environment'
    }
    It 'omits -Tenant when only the service-principal sets declare it (credential sign-in)' {
        Set-Item -Path 'function:Connect-PowerBIServiceAccount' -Value $script:OldShape
        Connect-IQPowerBIModule -Credential $script:Cred2 -MaxAttempts 1 | Should -Be $script:PbiJwt2
        $script:ConnectCalls[0].Set | Should -Be 'UserAndCredential'
        @($script:ConnectCalls[0].Bound) | Should -Not -Contain 'Tenant'
    }
    It 'passes -Tenant when the user parameter set declares it' {
        Set-Item -Path 'function:Connect-PowerBIServiceAccount' -Value $script:NewShape
        Connect-IQPowerBIModule -MaxAttempts 1 | Should -Be $script:PbiJwt2
        @($script:ConnectCalls[0].Bound) | Should -Contain 'Tenant'
        $script:ConnectCalls[0].Set | Should -Be 'User'
    }
    It 'does not pass -Tenant when the tenant id is not a GUID' {
        Set-Item -Path 'function:Connect-PowerBIServiceAccount' -Value $script:NewShape
        $script:IQ.Auth.TenantId = 'contoso.onmicrosoft.com'
        Connect-IQPowerBIModule -MaxAttempts 1 | Should -Be $script:PbiJwt2
        @($script:ConnectCalls[0].Bound) | Should -Not -Contain 'Tenant'
    }
    It 'retries once without -Tenant when binding fails' {
        Set-Item -Path 'function:Connect-PowerBIServiceAccount' -Value $script:BindingErrorShape
        Connect-IQPowerBIModule -MaxAttempts 1 | Should -Be $script:PbiJwt2
        $script:ConnectCalls.Count | Should -Be 2
        @($script:ConnectCalls[0].Bound) | Should -Contain 'Tenant'
        @($script:ConnectCalls[1].Bound) | Should -Not -Contain 'Tenant'
    }
}

Describe 'Credential (ROPC) error mapping' {
    BeforeAll {
        Reset-IQTestEnvironment
        Reset-AuthMock
        Mock Invoke-RestMethod { Pop-AuthResponse -Uri $Uri -Body $Body -ContentType $ContentType }
        Mock Start-IQAuthSleep { $script:AuthSleeps.Add([int]$Seconds) }
        Initialize-IQContext -BaseFolder $script:Base -Options @{ Environment = 'Public'; NonInteractive = $true } | Out-Null
        $script:IQ.Interactive = $false
        $script:Cred = New-Object System.Management.Automation.PSCredential ('user@contoso.gov', (ConvertTo-SecureString -String 'Secret#Pass1' -AsPlainText -Force))
    }
    It 'maps <Code> to a clear MFA / Conditional Access message' -TestCases @(@{ Code = '50076' }, @{ Code = '50079' }, @{ Code = '53003' }, @{ Code = '65001' }) {
        param($Code)
        Reset-AuthMock
        Set-AuthQueue 'token' @(('{"error":"invalid_grant","error_description":"AADSTS' + $Code + ': blocked","error_codes":[' + $Code + ']}'))
        { Initialize-IQAuth -Mode Credential -Environment Public -Credential $script:Cred } | Should -Throw -ExpectedMessage '*MFA/Conditional Access blocks password auth*'
    }
    It 'reports an invalid password (AADSTS50126) clearly' {
        Reset-AuthMock
        Set-AuthQueue 'token' @('{"error":"invalid_grant","error_description":"AADSTS50126: Error validating credentials due to invalid username or password.","error_codes":[50126]}')
        { Initialize-IQAuth -Mode Credential -Environment Public -Credential $script:Cred } | Should -Throw
    }
    It 'signs in with the password grant and never logs the password' {
        Reset-AuthMock
        $jwt = New-IQTestJwt -Claims @{ upn = 'user@contoso.gov' } -ExpiresInMinutes 60
        Set-AuthQueue 'token' @([pscustomobject]@{ access_token = $jwt; refresh_token = 'RT-R'; expires_in = 3600 })
        Initialize-IQAuth -Mode Credential -Environment Public -Credential $script:Cred | Should -Be 'Credential'
        Get-IQToken -Resource PowerBI | Should -Be $jwt
        $call = $script:AuthCalls[0]
        $call.Body.grant_type | Should -Be 'password'
        $call.Body.username | Should -Be 'user@contoso.gov'
        $call.Body.scope | Should -Match '^https://analysis\.windows\.net/powerbi/api/\.default'
        (Get-Content -LiteralPath $script:IQ.LogFile -Raw) | Should -Not -Match 'Secret#Pass1'
        Get-IQAuthDescription | Should -Match 'Credential'
    }
    It 'reads the credential from IMPACTIQ_USERNAME / IMPACTIQ_PASSWORD in Auto mode' {
        Reset-AuthMock
        $env:IMPACTIQ_USERNAME = 'env@contoso.gov'; $env:IMPACTIQ_PASSWORD = 'EnvSecret!9'
        try {
            $jwt = New-IQTestJwt -Claims @{ upn = 'env@contoso.gov' } -ExpiresInMinutes 60
            Set-AuthQueue 'token' @([pscustomobject]@{ access_token = $jwt; refresh_token = 'RT-E'; expires_in = 3600 })
            Initialize-IQAuth -Mode Auto -Environment Public | Should -Be 'Credential'
            $script:AuthCalls[0].Body.username | Should -Be 'env@contoso.gov'
            (Get-Content -LiteralPath $script:IQ.LogFile -Raw) | Should -Not -Match 'EnvSecret!9'
            $manifestJson = ConvertTo-Json -InputObject $script:IQ.Options -Depth 20
            $manifestJson | Should -Not -Match 'EnvSecret!9'
        }
        finally { Reset-IQTestEnvironment }
    }
}

Describe 'AccessToken mode' {
    BeforeAll {
        Reset-IQTestEnvironment
        Initialize-IQContext -BaseFolder $script:Base -Options @{ Environment = 'Public'; NonInteractive = $true } | Out-Null
        $script:IQ.Interactive = $false
    }
    AfterAll { Reset-IQTestEnvironment }
    It 'strips a "Bearer " prefix, warns about the expiry and returns $null for Fabric' {
        $jwt = New-IQTestJwt -Claims @{ upn = 'static@contoso.gov' } -ExpiresInMinutes 30
        $env:IMPACTIQ_PBI_TOKEN = 'Bearer ' + $jwt
        Initialize-IQAuth -Mode AccessToken -Environment Public | Should -Be 'AccessToken'
        Get-IQToken -Resource PowerBI | Should -Be $jwt
        Get-IQToken -Resource Fabric | Should -BeNullOrEmpty
        (Get-Content -LiteralPath $script:IQ.LogFile -Raw) | Should -Match '\[WARN\].*(?i)expire'
    }
    It 'throws at start when the static token is already expired' {
        $env:IMPACTIQ_PBI_TOKEN = New-IQTestJwt -ExpiresInMinutes -5
        { Initialize-IQAuth -Mode AccessToken -Environment Public } | Should -Throw
    }
    It 'throws with a clear message when IMPACTIQ_PBI_TOKEN is missing' {
        Reset-IQTestEnvironment
        { Initialize-IQAuth -Mode AccessToken -Environment Public } | Should -Throw -ExpectedMessage '*IMPACTIQ_PBI_TOKEN*'
    }
}

Describe 'Headless guards' {
    BeforeAll {
        Reset-IQTestEnvironment
        Initialize-IQContext -BaseFolder $script:Base -Options @{ Environment = 'Public'; NonInteractive = $true } | Out-Null
        $script:IQ.Interactive = $false
    }
    It 'refuses Interactive mode when not interactive' {
        { Initialize-IQAuth -Mode Interactive -Environment Public } | Should -Throw
    }
    It 'Get-IQToken throws before Initialize-IQAuth' {
        { Get-IQToken -Resource PowerBI } | Should -Throw
    }
}
