# Common.Tests.ps1 - ImpactIQ.Common.ps1 (brief section 2.1, 3, 5.1 exit codes).
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Base = Initialize-IQTestContext -Options @{ Environment = 'USGov'; MaxRetries = 5 } -NoRun -Prefix 'common'
}
AfterAll {
    Remove-IQTestFolder -Path $script:Base
}

Describe 'Initialize-IQContext / Get-IQContext' {
    It 'creates the context hashtable and the Config/State/Logs folders' {
        $ctx = Get-IQContext
        $ctx | Should -Not -BeNullOrEmpty
        $ctx.BaseFolder | Should -Be $script:Base
        (Join-Path $script:Base 'State') | Should -Exist
        (Join-Path $script:Base 'Logs') | Should -Exist
        (Join-Path $script:Base 'Config') | Should -Exist
    }
    It 'exposes the section 2.9 keys' {
        $ctx = Get-IQContext
        foreach ($key in @('BaseFolder', 'ConfigFolder', 'StatePath', 'LogsPath', 'LogFile', 'IsWindows', 'IsAzureDevOps', 'Interactive', 'Options', 'Environment', 'Endpoints', 'Auth', 'Tools', 'Paths', 'Stats')) {
            $ctx.ContainsKey($key) | Should -BeTrue -Because "context key $key"
        }
        $ctx.IsWindows | Should -Be ($env:OS -eq 'Windows_NT')
        $ctx.Stats.ApiCalls | Should -Be 0
    }
    It 'is headless when NonInteractive is set' {
        (Get-IQContext).Interactive | Should -BeFalse
        Test-IQInteractive | Should -BeFalse
    }
    It 'resolves the environment given in Options' {
        (Get-IQContext).Environment | Should -Be 'USGov'
        (Get-IQContext).Endpoints.ApiPrefix | Should -Be 'https://api.powerbigov.us'
    }
    It 'resolves a relative BaseFolder against $PWD, not the process working directory' {
        $saved = $script:IQ
        $parent = Join-Path $script:Base 'relbase'
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
        Push-Location -LiteralPath $parent
        try {
            $ctx = Initialize-IQContext -BaseFolder 'sub' -Options @{ NonInteractive = $true }
            $ctx.BaseFolder | Should -Be ([System.IO.Path]::GetFullPath((Join-Path $parent 'sub')))
            (Join-Path (Join-Path $parent 'sub') 'State') | Should -Exist
        }
        finally { Pop-Location; $script:IQ = $saved }
    }
    It 'names the default log file with the Gregorian calendar whatever the current culture' {
        $saved = $script:IQ
        $culture = [System.Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('th-TH')
            $ctx = Initialize-IQContext -BaseFolder (Join-Path $script:Base 'culture') -Options @{ NonInteractive = $true }
            $year = [datetime]::Now.ToString('yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
            Split-Path -Leaf $ctx.LogFile | Should -Match ('^ImpactIQ_' + $year + '\d{4}_\d{6}\.log$')
        }
        finally { [System.Threading.Thread]::CurrentThread.CurrentCulture = $culture; $script:IQ = $saved }
    }
}

Describe 'Get-IQCleanName (monolith parity)' {
    It 'matches the original sanitiser for <Name>' -TestCases @(
        @{ Name = 'Finance [Prod]' }
        @{ Name = ' Sales / EU_2024 & co, v1.2-x' }
        @{ Name = 'Ärzte Übersicht (2024)' }
        @{ Name = 'a:b*c?d"e<f>g|h' }
        @{ Name = '  leading and trailing  ' }
        @{ Name = 'Report [Q1] & [Q2].pbix' }
        @{ Name = '' }
    ) {
        param($Name)
        # The monolith: -replace '\[','(' -replace '\]',')' then -replace "[^a-zA-Z0-9\(\)&,.-]", " " then .TrimStart()
        $expected = (($Name -replace '\[', '(') -replace '\]', ')') -replace "[^a-zA-Z0-9\(\)&,.-]", " "
        $expected = $expected.TrimStart()
        Get-IQCleanName -Name $Name | Should -BeExactly $expected
    }
    It 'keeps trailing spaces (TrimStart only, like the monolith)' {
        Get-IQCleanName -Name 'A B ' | Should -BeExactly 'A B '
    }
    It 'returns an empty string for $null' {
        Get-IQCleanName -Name $null | Should -Be ''
    }
    It 'produces the documented "<CleanWs> ~ <CleanModel>" pieces' {
        (Get-IQCleanName -Name 'Finance [Prod]') + ' ~ ' + (Get-IQCleanName -Name 'Sales/Model') | Should -Be 'Finance (Prod) ~ Sales Model'
    }
}

Describe 'Get-IQSafeKey' {
    It 'lower-cases and replaces unsafe characters with _' {
        Get-IQSafeKey -Value 'AbC-123 x/y\z' | Should -Be 'abc-123_x_y_z'
    }
    It 'leaves GUIDs unchanged' {
        Get-IQSafeKey -Value 'AAAAAAAA-1111-4111-8111-111111111111' | Should -Be 'aaaaaaaa-1111-4111-8111-111111111111'
    }
    It 'caps the length at 120 characters' {
        (Get-IQSafeKey -Value ('z' * 300)).Length | Should -Be 120
    }
    It 'keeps dots, dashes and underscores' {
        Get-IQSafeKey -Value 'ws-my_workspace.v2' | Should -Be 'ws-my_workspace.v2'
    }
}

Describe 'JSON file helpers' {
    BeforeAll { $script:JsonPath = Join-Path $script:Base 'roundtrip.json' }
    It 'round-trips an object with nested values and non-ASCII text without a BOM' {
        $obj = [ordered]@{ Name = 'Ärzte "quoted"'; Count = 3; Flag = $true; Nested = @{ List = @(1, 2, 3); Inner = @{ Deep = 'x' } } }
        ConvertTo-IQJsonFile -Object $obj -Path $script:JsonPath
        $bytes = [System.IO.File]::ReadAllBytes($script:JsonPath)
        (($bytes[0] -eq 0xEF) -and ($bytes[1] -eq 0xBB)) | Should -BeFalse -Because 'UTF-8 without BOM'
        ($script:JsonPath + '.tmp') | Should -Not -Exist -Because 'atomic write leaves no temp file'
        $back = ConvertFrom-IQJsonFile -Path $script:JsonPath
        $back.Name | Should -Be 'Ärzte "quoted"'
        $back.Count | Should -Be 3
        $back.Flag | Should -BeTrue
        @($back.Nested.List).Count | Should -Be 3
        $back.Nested.Inner.Deep | Should -Be 'x'
    }
    It 'serialises deep structures (Depth 20) without truncation' {
        $deep = @{ l1 = @{ l2 = @{ l3 = @{ l4 = @{ l5 = @{ l6 = @{ l7 = @{ l8 = @{ l9 = @{ l10 = @{ l11 = @{ l12 = 'bottom' } } } } } } } } } } } }
        ConvertTo-IQJsonFile -Object $deep -Path $script:JsonPath
        (ConvertFrom-IQJsonFile -Path $script:JsonPath).l1.l2.l3.l4.l5.l6.l7.l8.l9.l10.l11.l12 | Should -Be 'bottom'
    }
    It 'writes arrays (including one-element arrays) as JSON arrays' {
        ConvertTo-IQJsonFile -Object @(@{ a = 1 }) -Path $script:JsonPath
        (Get-Content -LiteralPath $script:JsonPath -Raw).TrimStart() | Should -Match '^\['
        @(ConvertFrom-IQJsonFile -Path $script:JsonPath).Count | Should -Be 1
    }
    It 'returns $null for a missing file' {
        ConvertFrom-IQJsonFile -Path (Join-Path $script:Base 'does-not-exist.json') | Should -BeNullOrEmpty
    }
    It 'replaces an existing file atomically (Move-Item -Force)' {
        ConvertTo-IQJsonFile -Object @{ v = 1 } -Path $script:JsonPath
        ConvertTo-IQJsonFile -Object @{ v = 2 } -Path $script:JsonPath
        (ConvertFrom-IQJsonFile -Path $script:JsonPath).v | Should -Be 2
    }
    It 'resolves a relative -Path against $PWD for both write and read' {
        Push-Location -LiteralPath $script:Base
        try {
            ConvertTo-IQJsonFile -Object @{ rel = $true } -Path 'relative.json'
            (Join-Path $script:Base 'relative.json') | Should -Exist
            (Join-Path $script:Base 'relative.json.tmp') | Should -Not -Exist
            (ConvertFrom-IQJsonFile -Path 'relative.json').rel | Should -BeTrue
        }
        finally { Pop-Location }
    }
    It 'writes and replaces files under a folder whose name contains brackets' {
        $p = Join-Path (Join-Path $script:Base 'ws [Prod]') 'state.json'
        ConvertTo-IQJsonFile -Object @{ v = 1 } -Path $p
        ConvertTo-IQJsonFile -Object @{ v = 2 } -Path $p
        [System.IO.File]::Exists($p) | Should -BeTrue
        [System.IO.File]::Exists($p + '.tmp') | Should -BeFalse
        (ConvertFrom-IQJsonFile -Path $p).v | Should -Be 2
    }
}

Describe 'Invoke-IQWithRetry' {
    BeforeEach {
        Mock Start-Sleep { }
        $script:Attempts = 0
    }
    It 'retries and returns the result once the block succeeds' {
        $r = Invoke-IQWithRetry -ScriptBlock { $script:Attempts++; if ($script:Attempts -lt 3) { throw 'boom' }; 'done' } -MaxAttempts 5 -InitialDelaySeconds 1 -Description 'test'
        $r | Should -Be 'done'
        $script:Attempts | Should -Be 3
        Should -Invoke Start-Sleep -Times 2 -Exactly
    }
    It 'uses exponential backoff capped at 60 s' {
        $script:Delays = New-Object System.Collections.Generic.List[int]
        Mock Start-Sleep { $script:Delays.Add([int]$Seconds) }
        { Invoke-IQWithRetry -ScriptBlock { throw 'always' } -MaxAttempts 8 -InitialDelaySeconds 2 } | Should -Throw
        $script:Delays.Count | Should -Be 7
        $script:Delays[0] | Should -Be 2
        $script:Delays[1] | Should -Be 4
        ($script:Delays | Measure-Object -Maximum).Maximum | Should -BeLessOrEqual 60
    }
    It 'does not retry when -RetryOn returns $false' {
        { Invoke-IQWithRetry -ScriptBlock { $script:Attempts++; throw 'nope' } -MaxAttempts 4 -InitialDelaySeconds 1 -RetryOn { $_.Exception.Message -eq 'other' } } | Should -Throw
        $script:Attempts | Should -Be 1
    }
    It 'rethrows the original error after the last attempt' {
        { Invoke-IQWithRetry -ScriptBlock { $script:Attempts++; throw 'final failure' } -MaxAttempts 3 -InitialDelaySeconds 1 } | Should -Throw -ExpectedMessage '*final failure*'
        $script:Attempts | Should -Be 3
    }
    It 'counts retries in $IQ.Stats.Retries' {
        $before = [int](Get-IQContext).Stats.Retries
        Invoke-IQWithRetry -ScriptBlock { $script:Attempts++; if ($script:Attempts -lt 2) { throw 'x' }; 1 } -MaxAttempts 3 -InitialDelaySeconds 1 | Out-Null
        [int](Get-IQContext).Stats.Retries | Should -Be ($before + 1)
    }
    It 'passes the ErrorRecord to -RetryOn as $_, $args[0] and a param() argument' {
        foreach ($filter in @({ $_.Exception.Message -eq 'transient' }, { $args[0].Exception.Message -eq 'transient' }, { param($rec) $rec.Exception.Message -eq 'transient' })) {
            $script:Attempts = 0
            $r = Invoke-IQWithRetry -ScriptBlock { $script:Attempts++; if ($script:Attempts -lt 2) { throw 'transient' }; 'ok' } -MaxAttempts 3 -InitialDelaySeconds 1 -RetryOn $filter
            $r | Should -Be 'ok'
            $script:Attempts | Should -Be 2 -Because "filter $filter must see the ErrorRecord"
        }
    }
    It 'does not retry, and logs at Debug, when the -RetryOn filter itself throws' {
        { Invoke-IQWithRetry -ScriptBlock { $script:Attempts++; throw 'nope' } -MaxAttempts 4 -InitialDelaySeconds 1 -RetryOn { throw 'filter broke' } -Description 'flaky' } | Should -Throw -ExpectedMessage '*nope*'
        $script:Attempts | Should -Be 1
        (Get-Content -LiteralPath (Get-IQContext).LogFile -Raw) | Should -Match 'flaky: RetryOn filter threw \(filter broke\)'
    }
}

Describe 'Get-IQEnvironmentSettings (brief section 3)' {
    It 'returns the corrected endpoint table for <Env>' -TestCases @(
        @{ Env = 'Public'; Api = 'https://api.powerbi.com'; Authority = 'https://login.microsoftonline.com'; Resource = 'https://analysis.windows.net/powerbi/api'; Xmla = 'powerbi://api.powerbi.com'; Az = 'AzureCloud' }
        @{ Env = 'USGov'; Api = 'https://api.powerbigov.us'; Authority = 'https://login.microsoftonline.com'; Resource = 'https://analysis.usgovcloudapi.net/powerbi/api'; Xmla = 'powerbi://api.powerbigov.us'; Az = 'AzureCloud' }
        @{ Env = 'USGovHigh'; Api = 'https://api.high.powerbigov.us'; Authority = 'https://login.microsoftonline.us'; Resource = 'https://high.analysis.usgovcloudapi.net/powerbi/api'; Xmla = 'powerbi://api.high.powerbigov.us'; Az = 'AzureUSGovernment' }
        @{ Env = 'USGovMil'; Api = 'https://api.mil.powerbigov.us'; Authority = 'https://login.microsoftonline.us'; Resource = 'https://mil.analysis.usgovcloudapi.net/powerbi/api'; Xmla = 'powerbi://api.mil.powerbigov.us'; Az = 'AzureUSGovernment' }
        @{ Env = 'China'; Api = 'https://api.powerbi.cn'; Authority = 'https://login.chinacloudapi.cn'; Resource = 'https://analysis.chinacloudapi.cn/powerbi/api'; Xmla = 'powerbi://api.powerbi.cn'; Az = 'AzureChinaCloud' }
        @{ Env = 'Germany'; Api = 'https://api.powerbi.de'; Authority = 'https://login.microsoftonline.de'; Resource = 'https://analysis.cloudapi.de/powerbi/api'; Xmla = 'powerbi://api.powerbi.de'; Az = 'AzureGermanCloud' }
    ) {
        param($Env, $Api, $Authority, $Resource, $Xmla, $Az)
        $s = Get-IQEnvironmentSettings -Environment $Env
        $s.ApiPrefix | Should -Be $Api
        $s.Authority | Should -Be $Authority
        $s.PowerBIResource | Should -Be $Resource
        $s.XmlaPrefix | Should -Be $Xmla
        $s.AzEnvironment | Should -Be $Az
        $s.MicrosoftPowerBIMgmtEnvironment | Should -Be $Env
        $s.FabricApiPrefix | Should -Match '^https://'
        $s.FabricResource | Should -Be $s.FabricApiPrefix
    }
    It 'honours the FabricApiPrefix and Authority overrides' {
        $s = Get-IQEnvironmentSettings -Environment 'USGov' -FabricApiPrefixOverride 'https://fabric.example.gov/' -AuthorityOverride 'https://login.example.gov'
        $s.FabricApiPrefix | Should -Be 'https://fabric.example.gov'
        $s.FabricResource | Should -Be 'https://fabric.example.gov'
        $s.Authority | Should -Be 'https://login.example.gov'
    }
    It 'throws for an unknown environment' {
        { Get-IQEnvironmentSettings -Environment 'Mars' } | Should -Throw
    }
}

Describe 'Write-IQLog' {
    It 'writes "[HH:mm:ss] [LEVEL] [Stage] [Item] message" to the log file' {
        Write-IQLog -Message 'hello world' -Level Warn -Stage 'TestStage' -Item 'Item1'
        $log = Get-Content -LiteralPath (Get-IQContext).LogFile -Raw
        $log | Should -Match '\[\d{2}:\d{2}:\d{2}\] \[WARN\] \[TestStage\] \[Item1\] hello world'
    }
    It 'redacts connection-string passwords and bearer tokens' {
        Write-IQLog -Message 'Provider=MSOLAP;Data Source=x;Password=SuperSecret123;Other=1 Bearer eyJabc.def.ghi' -Level Info
        $log = Get-Content -LiteralPath (Get-IQContext).LogFile -Raw
        $log | Should -Not -Match 'SuperSecret123'
        $log | Should -Not -Match 'eyJabc\.def\.ghi'
    }
    It 'never throws (bad input, exception parameter)' {
        { Write-IQLog -Message $null -Level Error -Exception (New-Object System.Exception 'inner') } | Should -Not -Throw
        { Write-IQLog -Message 'x' -Level Debug -Stage $null -Item $null } | Should -Not -Throw
    }
    It 'masks whole secret values (special characters included), raw JWTs and device codes' {
        Write-IQLog -Message 'password: Pa$$w0rd!x device_code=ABCD-1234 token eyJhbGciOiJSUzI1NiJ9.eyJleHAiOjF9.abc-def_ghi client_secret="s3cr3t~!" tail' -Level Info
        $log = Get-Content -LiteralPath (Get-IQContext).LogFile -Raw
        $log | Should -Not -Match 'w0rd'
        $log | Should -Not -Match 'ABCD-1234'
        $log | Should -Not -Match 'eyJleHAiOjF9'
        $log | Should -Not -Match 's3cr3t'
        $log | Should -Match 'client_secret="\*\*\*" tail'
    }
    It 'accepts an ErrorRecord (or anything else) for -Exception without throwing' {
        $rec = $null
        try { throw 'from catch block' } catch { $rec = $_ }
        $rec | Should -BeOfType [System.Management.Automation.ErrorRecord]
        { Write-IQLog -Message 'wrapped' -Level Error -Exception $rec } | Should -Not -Throw
        { Write-IQLog -Message 'weird' -Level Warn -Exception 'just a string' } | Should -Not -Throw
        $log = Get-Content -LiteralPath (Get-IQContext).LogFile -Raw
        $log | Should -Match 'wrapped :: from catch block'
        $log | Should -Match 'weird :: just a string'
    }
}

Describe 'Get-IQDateFolder' {
    It 'returns the newest yyyy-MM-dd sub folder and ignores junk' {
        $root = Join-Path $script:Base 'Model Backups'
        foreach ($d in @('2026-01-05', '2025-12-31', 'junk', '2026-13-01', 'manifest')) { New-Item -ItemType Directory -Path (Join-Path $root $d) -Force | Out-Null }
        Split-Path -Leaf (Get-IQDateFolder -Root $root) | Should -Be '2026-01-05'
    }
    It 'returns $null when the root does not exist' {
        Get-IQDateFolder -Root (Join-Path $script:Base 'nothing-here') | Should -BeNullOrEmpty
    }
}

Describe 'Test-IQTimeBudget (time budget for capped agents)' {
    BeforeEach { $script:IQ.BudgetExceeded = $false; $script:IQ.StartedUtc = [datetime]::UtcNow }
    AfterAll { $script:IQ.Options.Remove('TimeBudgetMinutes'); $script:IQ.BudgetExceeded = $false; $script:IQ.StartedUtc = [datetime]::UtcNow }
    It 'is false when no budget is configured (0 = unlimited)' {
        $script:IQ.Options['TimeBudgetMinutes'] = 0
        $script:IQ.StartedUtc = [datetime]::UtcNow.AddHours(-5)
        Test-IQTimeBudget -Stage 'Inventory' | Should -BeFalse
        $script:IQ.BudgetExceeded | Should -BeFalse
    }
    It 'is false while budget minus the 2-minute grace has not elapsed' {
        $script:IQ.Options['TimeBudgetMinutes'] = 60
        $script:IQ.StartedUtc = [datetime]::UtcNow.AddMinutes(-30)
        Test-IQTimeBudget -Stage 'Inventory' | Should -BeFalse
        $script:IQ.BudgetExceeded | Should -BeFalse
    }
    It 'is true once budget minus grace has elapsed, sets BudgetExceeded and logs once' {
        $script:IQ.Options['TimeBudgetMinutes'] = 60
        $script:IQ.StartedUtc = [datetime]::UtcNow.AddMinutes(-58.5)
        Test-IQTimeBudget -Stage 'ModelBackup' -Item 'WS ~ Model' | Should -BeTrue
        $script:IQ.BudgetExceeded | Should -BeTrue
        Test-IQTimeBudget -Stage 'ModelBackup' | Should -BeTrue -Because 'the flag is sticky for the rest of the process'
        $lines = @(Get-Content -LiteralPath $script:IQ.LogFile | Where-Object { $_ -match 'Time budget of 60 min reached' })
        $lines.Count | Should -Be 1
    }
    It 'stays true after the flag was set even when the clock says otherwise' {
        $script:IQ.Options['TimeBudgetMinutes'] = 60
        $script:IQ.BudgetExceeded = $true
        Test-IQTimeBudget | Should -BeTrue
    }
    It 'honours Options.TimeBudgetGraceMinutes instead of the 2-minute default' {
        $script:IQ.Options['TimeBudgetMinutes'] = 60
        $script:IQ.Options['TimeBudgetGraceMinutes'] = 25
        try {
            $script:IQ.StartedUtc = [datetime]::UtcNow.AddMinutes(-30)
            Test-IQTimeBudget -Stage 'Inventory' | Should -BeFalse
            $script:IQ.StartedUtc = [datetime]::UtcNow.AddMinutes(-36)
            Test-IQTimeBudget -Stage 'Inventory' | Should -BeTrue
            (Get-Content -LiteralPath $script:IQ.LogFile -Raw) | Should -Match '25 min grace'
        }
        finally { $script:IQ.Options.Remove('TimeBudgetGraceMinutes') }
    }
    It 'uses IMPACTIQ_BUDGET_START_UTC / Options.BudgetStartUtc as the clock origin when given' {
        $script:IQ.Options['TimeBudgetMinutes'] = 60
        $script:IQ.StartedUtc = [datetime]::UtcNow
        $env:IMPACTIQ_BUDGET_START_UTC = [datetime]::UtcNow.AddMinutes(-59).ToString('o')
        try {
            Test-IQTimeBudget -Stage 'Inventory' | Should -BeTrue -Because 'the agent clock started 59 minutes ago'
            $script:IQ.BudgetExceeded = $false
            $script:IQ.Options['BudgetStartUtc'] = [datetime]::UtcNow.AddMinutes(-10)
            Test-IQTimeBudget -Stage 'Inventory' | Should -BeFalse -Because 'the Options value wins over the environment variable'
        }
        finally { $env:IMPACTIQ_BUDGET_START_UTC = $null; $script:IQ.Options.Remove('BudgetStartUtc') }
    }
}

Describe 'Get-IQExitCode (brief section 5.1)' {
    It 'maps <Status> to exit code <Expected>' -TestCases @(
        @{ Status = 'Completed'; Failures = @(); Stages = @{}; Expected = 0 }
        @{ Status = 'CompletedWithErrors'; Failures = @(); Stages = @{}; Expected = 2 }
        @{ Status = 'Failed'; Failures = @(); Stages = @{}; Expected = 1 }
        @{ Status = 'Running'; Failures = @(); Stages = @{}; Expected = 1 }
        @{ Status = 'Paused'; Failures = @(); Stages = @{}; Expected = 3 }
        @{ Status = 'Paused'; Failures = @(@{ stage = 'x' }); Stages = @{ Dataflows = @{ status = 'Paused' } }; Expected = 3 }
        @{ Status = 'Completed'; Failures = @(@{ stage = 'x' }); Stages = @{}; Expected = 2 }
        @{ Status = 'Completed'; Failures = @(); Stages = @{ Dataflows = @{ status = 'Failed' } }; Expected = 2 }
        @{ Status = 'Completed'; Failures = $null; Stages = @{}; Expected = 0 }
        @{ Status = 'Completed'; Failures = @(); Stages = @{ ModelBackup = @{ status = 'Paused' } }; Expected = 3 }
        @{ Status = 'Completed'; Failures = @(); Stages = @{ Inventory = @{ status = 'Failed' }; ModelBackup = @{ status = 'Paused' } }; Expected = 3 }
    ) {
        param($Status, $Failures, $Stages, $Expected)
        Get-IQExitCode -Manifest @{ status = $Status; failures = $Failures; stages = $Stages } | Should -Be $Expected
    }
    It 'returns 0 for a clean Completed manifest without a failures member (hashtable and object form)' {
        Get-IQExitCode -Manifest ([ordered]@{ status = 'Completed'; stages = [ordered]@{} }) | Should -Be 0
        Get-IQExitCode -Manifest ([pscustomobject]@{ status = 'Completed'; failures = $null }) | Should -Be 0
    }
    It 'returns 1 for a missing manifest' {
        Get-IQExitCode -Manifest $null | Should -Be 1
    }
}
