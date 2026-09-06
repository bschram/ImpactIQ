# Extras.Tests.ps1 - ImpactIQ.Extras.ps1 (brief section 8.4): admin probe outcomes, Scanner API batch cache / time
# budget / Retry-After / timeout handling, activity-day paging, the activity window, and the sheet flattening. The
# Power BI service is never touched: Invoke-IQApi, Invoke-IQHttpRequest and Start-IQExtrasSleep are mocked
# (host-independent: Linux CI and Windows PowerShell 5.1).
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Ids = Get-IQTestFixtureJson -Relative 'ids.json'
    Mock Get-IQToken { 'test-token' }
    Mock Get-IQSelectedWorkspaces { @() }

    # Scanner API stand-ins. $script:ScanIds is what admin/workspaces/modified answers (raw JSON built from it);
    # every getInfo is counted in $script:GetInfoCalls; scanStatus answers come from Invoke-IQHttpRequest (mocked
    # below) with a Retry-After header; scanResult returns one workspace row per id of the batch.
    $script:ScanIds = @()
    $script:GetInfoCalls = 0
    $script:Sleeps = New-Object System.Collections.Generic.List[int]
    $script:PollStatuses = @('Succeeded')
    $script:PollIndex = 0
    $script:RetryAfter = '30'
    $script:PollHook = $null
    $script:ScanBodies = @{}
    $script:ActivityPages = @{}
    $script:GroupPages = @{}

    Mock Start-IQExtrasSleep { $script:Sleeps.Add([int]$Seconds) }
    Mock Invoke-IQHttpRequest {
        $script:IQTestApiCalls.Add(('PowerBI GET ' + $Url))
        if ($null -ne $script:PollHook) { & $script:PollHook }
        $status = $script:PollStatuses[[math]::Min($script:PollIndex, $script:PollStatuses.Count - 1)]
        $script:PollIndex++
        $headers = @{ 'Content-Type' = 'application/json; charset=utf-8' }
        if ($null -ne $script:RetryAfter) { $headers['Retry-After'] = $script:RetryAfter }
        return @{ StatusCode = 200; Headers = $headers; Content = ('{"id":"scan-x","status":"' + $status + '"}'); ContentType = 'application/json'; ElapsedMs = 1; OutFile = $null }
    }
    Mock Invoke-IQApi {
        $relative = $Path.Trim().TrimStart('/')
        $script:IQTestApiCalls.Add(('{0} {1} {2}' -f $Api, $Method.ToUpperInvariant(), $relative))
        switch -Regex ($relative) {
            '^admin/capacities$' {
                if ($script:ProbeMode -eq 'throw') { throw 'HTTP 503 after 5 attempts' }
                if ($script:ProbeMode -eq 'null') { return $null }
                return ([pscustomobject]@{ value = @([pscustomobject]@{ id = $script:Ids.cap1 }) })
            }
            '^admin/workspaces/modified$' {
                if ($null -eq $script:ScanIds) { return $null }
                $text = ConvertTo-Json -InputObject @($script:ScanIds | ForEach-Object { [ordered]@{ id = $_ } }) -Compress -Depth 5
                if (@($script:ScanIds).Count -eq 0) { $text = '[]' }
                if ($Raw) { return $text }
                return (ConvertFrom-Json -InputObject $text)
            }
            '^admin/workspaces/getInfo$' {
                $script:GetInfoCalls++
                $ids = @($Body.workspaces)
                $script:ScanBodies['scan-' + $script:GetInfoCalls] = $ids
                return ([pscustomobject]@{ id = ('scan-' + $script:GetInfoCalls); status = 'NotStarted'; createdDateTime = '2026-09-04T10:00:00Z' })
            }
            '^admin/workspaces/scanResult/(.+)$' {
                $ids = @()
                $key = $Matches[1]
                if ($script:ScanBodies.ContainsKey($key)) { $ids = @($script:ScanBodies[$key]) }
                $text = ConvertTo-Json -InputObject ([ordered]@{
                        workspaces          = @($ids | ForEach-Object { [ordered]@{ id = $_; name = ('WS ' + $_.Substring(0, 8)); type = 'Workspace'; state = 'Active'; datasets = @([ordered]@{ id = ('d-' + $_.Substring(0, 8)); name = 'Model'; tables = @([ordered]@{ name = 'T'; columns = @([ordered]@{ name = 'C'; dataType = 'String' }) }) }) } })
                        datasourceInstances = @()
                    }) -Depth 20
                if ($Raw) { return $text }
                return (ConvertFrom-Json -InputObject $text)
            }
            '^admin/groups$' {
                $skip = 0
                if ($Query -and $Query.ContainsKey('$skip')) { $skip = [int]$Query['$skip'] }
                if (-not $script:GroupPages.ContainsKey($skip)) { return $null }
                $text = $script:GroupPages[$skip]
                if ($null -eq $text) { return $null }
                if ($Raw) { return $text }
                return (ConvertFrom-Json -InputObject $text)
            }
            '^admin/activityevents$' { return $script:ActivityPages['first'] }
            '^https://.*continuation' { return $script:ActivityPages['next'] }
            default { return $null }
        }
    }

    function Reset-ScanState {
        $script:GetInfoCalls = 0
        $script:Sleeps.Clear()
        $script:PollIndex = 0
        $script:PollStatuses = @('Succeeded')
        $script:RetryAfter = '30'
        $script:PollHook = $null
        $script:ScanBodies = @{}
        $script:IQTestApiCalls.Clear()
    }
    function Reset-Budget {
        $script:IQ.Options['TimeBudgetMinutes'] = 0
        $script:IQ['BudgetExceeded'] = $false
        $script:IQ['StartedUtc'] = [datetime]::UtcNow
    }
    function Get-ScanFile { return @(Get-ChildItem -LiteralPath (Get-IQExtrasFolder -Name 'admin') -Filter 'scan-0*.json' -File | Sort-Object Name) }
}
AfterAll { Remove-IQTestFolder -Path $script:Base }

Describe 'Test-IQAdminAccess and how the stage records the probe outcome (EX-01)' {
    BeforeAll {
        $script:Base = Initialize-IQTestContext -Options @{ IncludeAdminApis = $true; ExtrasPollSeconds = 0 } -Prefix 'extras-probe'
    }
    It 'a $null probe response (401/403/404) means not an administrator, not transient' {
        $script:ProbeMode = 'null'
        $p = Test-IQAdminAccess
        $p.IsAdmin | Should -BeFalse
        $p.Transient | Should -BeFalse
        $p.Message | Should -Match 'not a Fabric administrator'
    }
    It 'a probe exception (5xx after retries, network) is transient' {
        $script:ProbeMode = 'throw'
        $p = Test-IQAdminAccess
        $p.IsAdmin | Should -BeFalse
        $p.Transient | Should -BeTrue
        $p.Message | Should -Match 'HTTP 503'
    }
    It 'a transient probe failure checkpoints admin-groups and admin-scan as Failed, so the stage is CompletedWithErrors and retried on resume' {
        $script:ProbeMode = 'throw'
        $status = Invoke-IQStage -Name 'Extras' -Body { Invoke-IQExtrasStage | Out-Null }
        $status | Should -Be 'CompletedWithErrors'
        foreach ($key in @('admin-groups', 'admin-scan')) {
            $cp = Get-IQItemCheckpoint -Stage Extras -ItemKey $key
            $cp.status | Should -Be 'Failed'
            $cp.message | Should -Match 'Admin API probe failed'
            Test-IQItemDone -Stage Extras -ItemKey $key | Should -BeFalse
        }
        @($script:IQ.Manifest.failures | Where-Object { $_.stage -eq 'Extras' }).Count | Should -Be 2
        [int]$script:IQ.Manifest.stages.Extras.itemsFailed | Should -Be 2
    }
    It 'a definite "not an administrator" answer checkpoints them Skipped (success class) and clears the earlier failures' {
        $script:ProbeMode = 'null'
        $status = Invoke-IQStage -Name 'Extras' -Body { Invoke-IQExtrasStage | Out-Null }
        $status | Should -Be 'Completed'
        foreach ($key in @('admin-groups', 'admin-scan')) {
            (Get-IQItemCheckpoint -Stage Extras -ItemKey $key).status | Should -Be 'Skipped'
            Test-IQItemDone -Stage Extras -ItemKey $key | Should -BeTrue
        }
        @($script:IQ.Manifest.failures | Where-Object { $_.stage -eq 'Extras' }).Count | Should -Be 0
    }
    It 'Invoke-IQExtrasCollector does not checkpoint a collector that reports Paused' {
        Reset-Budget
        $r = Invoke-IQExtrasCollector -ItemKey 'paused-probe' -Item 'paused' -Method 'ScannerApi' -Body { @{ Success = $false; Paused = $true; Message = 'budget'; Outputs = @() } }
        $r | Should -Be 'Paused'
        Get-IQItemCheckpoint -Stage Extras -ItemKey 'paused-probe' | Should -BeNullOrEmpty
    }
}

Describe 'Get-IQAdminGroupInventory (EX-04, EX-05)' {
    BeforeAll {
        $script:Base = Initialize-IQTestContext -Options @{ AdminGroupsPageSize = 2 } -Prefix 'extras-groups'
        $script:Page1 = '{"value":[{"id":"' + $script:Ids.ws1 + '","name":"Finance","type":"Workspace","users":[{"emailAddress":"a@contoso.gov","groupUserAccessRight":"Admin"}],"reports":[{"id":"' + $script:Ids.r1 + '"}]},{"id":"' + $script:Ids.ws2 + '","name":"Sales","type":"Workspace","users":[],"reports":[]}]}'
        $script:Page2 = '{"value":[{"id":"' + $script:Ids.ws3 + '","name":"HR","type":"Workspace","users":[],"reports":[]}]}'
    }
    It 'a continuation page that returns nothing fails the collector instead of checkpointing a truncated tenant' {
        $script:GroupPages = @{ 0 = $script:Page1; 2 = $null }
        $r = Get-IQAdminGroupInventory
        $r.Success | Should -BeFalse
        $r.Message | Should -Match 'page 2'
        $r.Message | Should -Match 'retried on the next start'
    }
    It 'writes every page body verbatim (no ConvertTo-Json round-trip) and builds the two sheets' {
        $script:GroupPages = @{ 0 = $script:Page1; 2 = $script:Page2 }
        $r = Get-IQAdminGroupInventory
        $r.Success | Should -BeTrue
        $r.WorkspaceCount | Should -Be 3
        $r.UserCount | Should -Be 1
        $r.Pages | Should -Be 2
        $raw1 = Join-Path (Get-IQExtrasFolder -Name 'admin') 'groups-001.json'
        [System.IO.File]::ReadAllText($raw1) | Should -Be $script:Page1
        $sheet = ConvertFrom-IQJsonFile -Path (Get-IQExtrasSheetPath -SheetName 'AdminWorkspaces')
        $sheet.RowCount | Should -Be 3
        @($sheet.Rows | Where-Object { $_.WorkspaceId -eq $script:Ids.ws1 })[0].WorkspaceReportIds | Should -Be $script:Ids.r1
        @($sheet.Columns)[0] | Should -Be 'WorkspaceId'
    }
}

Describe 'Scanner API: deterministic batches, cache reuse, empty tenant (EX-02, EX-06)' {
    BeforeAll {
        $script:Base = Initialize-IQTestContext -Options @{ AdminScanBatchSize = 2; ExtrasPollSeconds = 1 } -Prefix 'extras-scan'
        Reset-Budget
    }
    It 'Get-IQAdminScanWorkspaceId returns @() for "[]" but throws for no response (403/404)' {
        $script:ScanIds = @()
        @(Get-IQAdminScanWorkspaceId).Count | Should -Be 0
        $script:ScanIds = $null
        { Get-IQAdminScanWorkspaceId } | Should -Throw '*returned no response*'
    }
    It 'sorts the ids so the batch layout does not depend on the API order' {
        $script:ScanIds = @($script:Ids.ws3, $script:Ids.ws1, $script:Ids.ws2)
        @(Get-IQAdminScanWorkspaceId) | Should -Be @($script:Ids.ws1, $script:Ids.ws2, $script:Ids.ws3)
    }
    It 'an empty modified list yields a successful item with empty Scan* sheets' {
        Reset-ScanState
        $script:ScanIds = @()
        $r = Get-IQAdminScanInventory
        $r.Success | Should -BeTrue
        $r.BatchCount | Should -Be 0
        $script:GetInfoCalls | Should -Be 0
        (ConvertFrom-IQJsonFile -Path (Get-IQExtrasSheetPath -SheetName 'ScanDatasets')).RowCount | Should -Be 0
    }
    It 'scans 4 workspaces in 2 batches, writes the scanResult bodies verbatim and honours Retry-After over ExtrasPollSeconds' {
        Reset-ScanState
        $script:PollStatuses = @('Running', 'Succeeded')
        $script:ScanIds = @($script:Ids.ws4, $script:Ids.ws2, $script:Ids.ws1, $script:Ids.ws3)
        $r = Get-IQAdminScanInventory
        $r.Success | Should -BeTrue
        $r.BatchCount | Should -Be 2
        $r.BatchesReused | Should -Be 0
        $script:GetInfoCalls | Should -Be 2
        @(Get-ScanFile).Count | Should -Be 2
        # first poll waits ExtrasPollSeconds (1), the next waits the 30 s the service asked for
        $script:Sleeps[0] | Should -Be 1
        $script:Sleeps | Should -Contain 30
        (ConvertFrom-IQJsonFile -Path (Get-IQExtrasSheetPath -SheetName 'ScanWorkspaces')).RowCount | Should -Be 4
        (ConvertFrom-IQJsonFile -Path (Get-IQExtrasSheetPath -SheetName 'ScanColumns')).RowCount | Should -Be 4
        $script:ScanBodies['scan-1'] | Should -Be @($script:Ids.ws1, $script:Ids.ws2)
    }
    It 'a retry with the ids in another order reuses every cached batch (no getInfo call)' {
        Reset-ScanState
        $script:ScanIds = @($script:Ids.ws1, $script:Ids.ws3, $script:Ids.ws4, $script:Ids.ws2)
        $r = Get-IQAdminScanInventory
        $r.Success | Should -BeTrue
        $r.BatchesReused | Should -Be 2
        $script:GetInfoCalls | Should -Be 0
        @(Get-ScanFile).Count | Should -Be 2
    }
    It 'a cached batch whose number moved is still found by its hash' {
        Reset-ScanState
        $script:ScanIds = @($script:Ids.ws4, $script:Ids.ws3)   # the former batch 2 is now batch 1
        $r = Get-IQAdminScanInventory
        $r.Success | Should -BeTrue
        $r.BatchCount | Should -Be 1
        $r.BatchesReused | Should -Be 1
        $script:GetInfoCalls | Should -Be 0
    }
}

Describe 'Scanner API: time budget and timeouts (EX-03, EX-08)' {
    BeforeAll {
        $script:Base = Initialize-IQTestContext -Options @{ AdminScanBatchSize = 1; ExtrasPollSeconds = 1 } -Prefix 'extras-budget'
        $script:ScanIds = @($script:Ids.ws1, $script:Ids.ws2, $script:Ids.ws3)
    }
    AfterEach { Reset-Budget }
    It 'stops before the first batch when the budget is used up: Paused, nothing submitted, no checkpoint' {
        Reset-ScanState
        $script:IQ.Options['TimeBudgetMinutes'] = 1
        $script:IQ['StartedUtc'] = [datetime]::UtcNow.AddMinutes(-5)
        $r = Get-IQAdminScanInventory
        $r.Paused | Should -BeTrue
        $r.Success | Should -BeFalse
        $script:GetInfoCalls | Should -Be 0
        Test-Path -LiteralPath (Get-IQExtrasSheetPath -SheetName 'ScanWorkspaces') | Should -BeFalse
        $status = Invoke-IQExtrasCollector -ItemKey 'admin-scan' -Item 'Scanner API' -Method 'ScannerApi' -Body { Get-IQAdminScanInventory }
        $status | Should -Be 'Paused'
        Get-IQItemCheckpoint -Stage Extras -ItemKey 'admin-scan' | Should -BeNullOrEmpty
    }
    It 'a budget that runs out inside the scanStatus poll pauses the collector and keeps the finished batches' {
        Reset-ScanState
        $script:IQ.Options['TimeBudgetMinutes'] = 1
        # batch 1 succeeds on its first poll; batch 2 is still Running when the budget runs out (set by the poll hook)
        $script:PollStatuses = @('Succeeded', 'Running', 'Running', 'Succeeded')
        $script:PollHook = { if ($script:GetInfoCalls -eq 2) { $script:IQ['StartedUtc'] = [datetime]::UtcNow.AddMinutes(-5) } }
        $r = Get-IQAdminScanInventory
        $r.Paused | Should -BeTrue
        $r.Message | Should -Match 'time budget reached while polling scan scan-2'
        $script:GetInfoCalls | Should -Be 2
        @(Get-ScanFile).Count | Should -Be 1
    }
    It 'the stage is marked Paused (not Completed) when the scan stops on the budget' {
        Reset-ScanState
        Remove-Item -LiteralPath (Join-Path $script:IQ.RunPath 'done') -Recurse -Force -ErrorAction SilentlyContinue
        $script:IQ.Options['TimeBudgetMinutes'] = 1
        $script:IQ['StartedUtc'] = [datetime]::UtcNow.AddMinutes(-5)
        $status = Invoke-IQStage -Name 'Extras' -Body {
            Invoke-IQExtrasCollector -ItemKey 'admin-scan' -Item 'Scanner API' -Method 'ScannerApi' -Body { Get-IQAdminScanInventory } | Out-Null
        }
        $status | Should -Be 'Paused'
        $script:IQ.Manifest.stages.Extras.status | Should -Be 'Paused'
    }
    It 'a batch timeout fails that batch, stops submitting the rest of this run and leaves the item Failed for the resume' {
        Reset-ScanState
        Remove-Item -LiteralPath (Get-IQExtrasFolder -Name 'admin') -Recurse -Force -ErrorAction SilentlyContinue
        Mock Invoke-IQAdminScanBatch {
            $script:GetInfoCalls++
            if (@($Ids) -contains $script:Ids.ws1) { throw (New-Object System.TimeoutException('scan scan-1 did not finish within 30 minute(s) (last status NotStarted)')) }
            return ([pscustomobject]@{ workspaces = @(); datasourceInstances = @() })
        }
        $r = Get-IQAdminScanInventory
        $r.Success | Should -BeFalse
        $r.Paused | Should -BeFalse
        $r.BatchesFailed | Should -Be 1
        $script:GetInfoCalls | Should -Be 1 -Because 'batches 2 and 3 must not be submitted after a timeout'
        $r.Message | Should -Match '2 batch\(es\) not submitted after a timeout'
        $status = Invoke-IQExtrasCollector -ItemKey 'admin-scan' -Item 'Scanner API' -Method 'ScannerApi' -Body { Get-IQAdminScanInventory }
        $status | Should -Be 'Failed'
        $script:GetInfoCalls | Should -Be 2 -Because 'the second run submits only the timed-out batch again'
        (Get-IQItemCheckpoint -Stage Extras -ItemKey 'admin-scan').status | Should -Be 'Failed'
    }
    It 'a Retry-After below ExtrasPollSeconds does not shorten the poll interval' {
        Reset-ScanState
        Remove-Item -LiteralPath (Get-IQExtrasFolder -Name 'admin') -Recurse -Force -ErrorAction SilentlyContinue
        $script:IQ.Options['ExtrasPollSeconds'] = 5
        $script:RetryAfter = '2'
        $script:PollStatuses = @('Running', 'Succeeded')
        $script:ScanIds = @($script:Ids.ws1)
        try {
            $r = Get-IQAdminScanInventory
            $r.Success | Should -BeTrue
            @($script:Sleeps | Select-Object -Unique) | Should -Be @(5)
        }
        finally { $script:IQ.Options['ExtrasPollSeconds'] = 1 }
    }
}

Describe 'Activity events: continuation failure and the day window (EX-04, EX-07)' {
    BeforeAll {
        $script:Base = Initialize-IQTestContext -Options @{ NowUtc = '2026-09-04T10:00:00Z'; ActivityDays = 30 } -Prefix 'extras-activity'
        Reset-Budget
        $script:Day = [datetime]::new(2026, 9, 3, 0, 0, 0, [System.DateTimeKind]::Utc)
    }
    It 'a continuation page that returns nothing fails the day (re-read on resume) instead of keeping a truncated day' {
        $script:ActivityPages = @{
            first = [pscustomobject]@{ activityEventEntities = @([pscustomobject]@{ Id = 'e1'; Activity = 'ViewReport' }); continuationUri = 'https://api.powerbi.com/v1.0/myorg/admin/activityevents?continuationToken=abc' }
            next  = $null
        }
        $r = Get-IQAdminActivityDay -Day $script:Day
        $r.Success | Should -BeFalse
        $r.Message | Should -Match 'continuation page 2'
        Test-Path -LiteralPath $r.Path | Should -BeFalse
    }
    It 'a complete day is written with every page and is not partial (the clock says 2026-09-04)' {
        $script:ActivityPages = @{
            first = [pscustomobject]@{ activityEventEntities = @([pscustomobject]@{ Id = 'e1'; Activity = 'ViewReport'; Nested = [pscustomobject]@{ a = 1 } }); continuationUri = 'https://api.powerbi.com/v1.0/myorg/admin/activityevents?continuationToken=abc' }
            next  = [pscustomobject]@{ activityEventEntities = @([pscustomobject]@{ Id = 'e2'; Activity = 'Export' }, [pscustomobject]@{ Id = 'e3'; Activity = 'Export' }); lastResultSet = $true }
        }
        $r = Get-IQAdminActivityDay -Day $script:Day
        $r.Success | Should -BeTrue
        $r.EventCount | Should -Be 3
        $r.Pages | Should -Be 2
        $r.Partial | Should -BeFalse
        (Get-IQAdminActivityDay -Day $script:Day.AddDays(1)).Partial | Should -BeTrue
        $path = Save-IQActivitySheet
        $sheet = ConvertFrom-IQJsonFile -Path $path
        $sheet.RowCount | Should -Be 6 -Because 'three events of 2026-09-03 plus three of the partial day 2026-09-04'
        @($sheet.Columns)[0] | Should -Be 'Id'
        $sheet.Columns | Should -Contain 'ActivityDate'
        @($sheet.Rows | Where-Object { $_.Id -eq 'e1' })[0].Nested | Should -Be '{"a":1}'
        @($sheet.Rows | Where-Object { $_.Id -eq 'e2' })[0].Nested | Should -BeNullOrEmpty
    }
    It 'the default -ActivityDays 30 is clamped to 28 days ending on the run clock without a Warn line' {
        Mock Write-IQLog { }
        $days = @(Get-IQAdminActivityDayList)
        $days.Count | Should -Be 28
        $days[-1].ToString('yyyy-MM-dd') | Should -Be '2026-09-04'
        $days[0].ToString('yyyy-MM-dd') | Should -Be '2026-08-08'
        Should -Invoke Write-IQLog -Times 0 -ParameterFilter { $Level -eq 'Warn' }
        Should -Invoke Write-IQLog -Times 1 -ParameterFilter { $Level -eq 'Info' -and $Message -like '*exceeds the activity log window*' }
    }
    It 'an explicit value above the window is clamped with a Warn' {
        Mock Write-IQLog { }
        $script:IQ.Options['ActivityDays'] = 45
        try { @(Get-IQAdminActivityDayList).Count | Should -Be 28 } finally { $script:IQ.Options['ActivityDays'] = 30 }
        Should -Invoke Write-IQLog -Times 1 -ParameterFilter { $Level -eq 'Warn' -and $Message -like '*-ActivityDays 45*' }
    }
    It 'a resumed run keeps the window of its original start date' {
        $script:IQ.Options['ActivityDays'] = 3
        $script:IQ.IsResume = $true
        $script:IQ.Manifest['startedUtc'] = '2026-09-01T08:00:00.0000000Z'
        try {
            $days = @(Get-IQAdminActivityDayList)
            @($days | ForEach-Object { $_.ToString('yyyy-MM-dd') }) | Should -Be @('2026-08-30', '2026-08-31', '2026-09-01')
        }
        finally {
            $script:IQ.IsResume = $false
            $script:IQ.Options['ActivityDays'] = 30
        }
    }
}

Describe 'Sheet flattening (EX-05)' {
    It 'ConvertTo-IQExtrasRow keeps scalars, flattens nested values to compact JSON and drops empty arrays' {
        $row = ConvertTo-IQExtrasRow -Object ([pscustomobject]@{ id = 'x'; name = 'N'; n = 3; flag = $true; nested = [pscustomobject]@{ a = 1 }; list = @(1, 2); empty = @() }) -Map ([ordered]@{ id = 'WorkspaceId' })
        @($row.Keys)[0] | Should -Be 'WorkspaceId'
        $row['WorkspaceId'] | Should -Be 'x'
        $row.Contains('id') | Should -BeFalse
        $row['n'] | Should -Be 3
        $row['flag'] | Should -BeTrue
        $row['nested'] | Should -Be '{"a":1}'
        $row['list'] | Should -Be '[1,2]'
        $row['empty'] | Should -BeNullOrEmpty
    }
    It 'New-IQExtrasSheet gives every row every column, preferred columns first, and still flattens object rows' {
        $rows = @(
            [ordered]@{ B = 'b1'; A = 'a1' },
            [pscustomobject]@{ A = 'a2'; C = [pscustomobject]@{ deep = $true }; D = @() }
        )
        $sheet = New-IQExtrasSheet -SheetName 'T' -Collector 'test' -Rows $rows -PreferredColumns @('A', 'Z')
        @($sheet.Columns) | Should -Be @('A', 'B', 'C', 'D')
        $sheet.RowCount | Should -Be 2
        $r1 = $sheet.Rows[0]
        $r1.A | Should -Be 'a1'
        $r1.PSObject.Properties['C'] | Should -Not -BeNullOrEmpty
        $r1.C | Should -BeNullOrEmpty
        $sheet.Rows[1].C | Should -Be '{"deep":true}'
        $sheet.Rows[1].D | Should -BeNullOrEmpty
        $sheet.Rows[1].B | Should -BeNullOrEmpty
        $kept = New-IQExtrasSheet -SheetName 'T' -Collector 'test' -Rows $rows -PreferredColumns @('A', 'Z') -KeepPreferred
        @($kept.Columns) | Should -Be @('A', 'Z', 'B', 'C', 'D')
    }
}
