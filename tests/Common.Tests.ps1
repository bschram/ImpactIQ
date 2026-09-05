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

Describe 'Get-IQExitCode (brief section 5.1)' {
    It 'maps <Status> to exit code <Expected>' -TestCases @(
        @{ Status = 'Completed'; Failures = @(); Stages = @{}; Expected = 0 }
        @{ Status = 'CompletedWithErrors'; Failures = @(); Stages = @{}; Expected = 2 }
        @{ Status = 'Failed'; Failures = @(); Stages = @{}; Expected = 1 }
        @{ Status = 'Running'; Failures = @(); Stages = @{}; Expected = 1 }
        @{ Status = 'Completed'; Failures = @(@{ stage = 'x' }); Stages = @{}; Expected = 2 }
        @{ Status = 'Completed'; Failures = @(); Stages = @{ Dataflows = @{ status = 'Failed' } }; Expected = 2 }
    ) {
        param($Status, $Failures, $Stages, $Expected)
        Get-IQExitCode -Manifest @{ status = $Status; failures = $Failures; stages = $Stages } | Should -Be $Expected
    }
    It 'returns 1 for a missing manifest' {
        Get-IQExitCode -Manifest $null | Should -Be 1
    }
}
