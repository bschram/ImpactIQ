# Tools.Tests.ps1 - ImpactIQ.Tools.ps1 (brief section 2.5): process runner, batch pool cleanup, tool install / update logic.
# Real processes are the current PowerShell host itself (works on Windows PowerShell 5.1, PowerShell 7, Linux);
# GitHub / registry / Windows-only paths are exercised through their pure helper functions or skipped off Windows.
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCommands', '', Justification = 'Pester 5 Should/Describe parameters are registered at run time; the 5.1/7.0 compatibility profiles only know Pester 3.4')]
param()
BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    $script:Base = Initialize-IQTestContext -Prefix 'tools'
    $script:OnWindows = ($env:OS -eq 'Windows_NT')
    $script:HostExe = (Get-Process -Id $PID).Path
    function New-HostJob {
        # A process job that runs a PowerShell command in the current host executable (cross-platform).
        param([string]$Command, [string]$ItemKey = 'job', [string]$Item = 'job', [string]$LogName)
        return @{ ItemKey = $ItemKey; Item = $Item; FilePath = $script:HostExe; ArgumentList = ('-NoProfile -NonInteractive -Command "' + $Command + '"'); WorkingDirectory = $script:Base; LogName = $LogName }
    }
    function New-ToolZip {
        # Builds <ZipPath> containing <TopFolder>\<ExeName> plus a lib\readme.txt next to it.
        param([string]$ZipPath, [string]$ExeName, [string]$TopFolder = 'pbi-tools-1.0')
        $staging = Join-Path $script:Base ('zip-staging-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $inner = [System.IO.Directory]::CreateDirectory((Join-Path $staging $TopFolder)).FullName
        [System.IO.File]::WriteAllText((Join-Path $inner $ExeName), 'fake exe ' + $ExeName)
        $lib = [System.IO.Directory]::CreateDirectory((Join-Path $inner 'lib')).FullName
        [System.IO.File]::WriteAllText((Join-Path $lib 'readme.txt'), 'lib file')
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        if (Test-Path -LiteralPath $ZipPath) { Remove-Item -LiteralPath $ZipPath -Force }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($staging, $ZipPath)
        Remove-Item -LiteralPath $staging -Recurse -Force
        return $ZipPath
    }
}
AfterAll { Remove-IQTestFolder -Path $script:Base }

Describe 'Argument redaction and Tabular Editor error detection' {
    It 'masks Password=, Bearer tokens and access_token= in the logged argument string' {
        $r = Get-IQRedactedArgumentList -ArgumentList '"Provider=MSOLAP;Data Source=x;User ID=u;Password=s3cret;" -S "Authorization: Bearer abc.def" x?access_token=tok&y=1'
        $r | Should -Match 'Password=\*\*\*;'
        $r | Should -Match 'Bearer \*\*\*'
        $r | Should -Match 'access_token=\*\*\*&y=1'
        $r | Should -Not -Match 's3cret|abc\.def|tok&'
    }
    It 'Select-IQTabularEditorError returns Error lines with their follow-up line' {
        $lines = @(Select-IQTabularEditorError -Output ("Loading model...`nError CS1002: ; expected`n   at line 3`nDone"))
        $lines.Count | Should -Be 2
        $lines[0] | Should -Match 'CS1002'
    }
}

Describe 'GitHub release helpers (T-03, T-04, T-08)' {
    It 'Get-IQTagFromRedirect reads the tag from a releases/tag redirect and rejects other locations' {
        Get-IQTagFromRedirect -Location 'https://github.com/TabularEditor/TabularEditor/releases/tag/2.26.0' | Should -Be '2.26.0'
        Get-IQTagFromRedirect -Location 'https://github.com/TabularEditor/TabularEditor/releases/tag/v1.2.3/' | Should -Be 'v1.2.3'
        Get-IQTagFromRedirect -Location 'https://github.com/TabularEditor/TabularEditor/releases' | Should -BeNullOrEmpty
        Get-IQTagFromRedirect -Location '' | Should -BeNullOrEmpty
        Get-IQTagFromRedirect -Location $null | Should -BeNullOrEmpty
    }
    It 'Test-IQToolVersionCurrent compares version segments, honours the stamp and never matches an empty tag' {
        Test-IQToolVersionCurrent -InstalledVersion '2.26.0.0' -Tag '2.26.0' | Should -BeTrue
        Test-IQToolVersionCurrent -InstalledVersion '1.2.0.123' -Tag 'v1.2.0' | Should -BeTrue
        Test-IQToolVersionCurrent -InstalledVersion '1.0.0.0' -Tag '1.0.0-beta.6' | Should -BeTrue
        Test-IQToolVersionCurrent -InstalledVersion '2.260.0' -Tag '2.26.0' | Should -BeFalse -Because 'a StartsWith comparison would accept this'
        Test-IQToolVersionCurrent -InstalledVersion '2.25.1' -Tag '2.26.0' | Should -BeFalse
        Test-IQToolVersionCurrent -InstalledVersion $null -Tag '2.26.0' -StampTag '2.26.0' | Should -BeTrue -Because 'the persisted stamp recorded that tag'
        Test-IQToolVersionCurrent -InstalledVersion '2.26.0' -Tag '' | Should -BeFalse
        Test-IQToolVersionCurrent -InstalledVersion 'abc' -Tag '2.26.0' | Should -BeFalse
    }
    It 'Select-IQPbiToolsAsset picks the Desktop zip regardless of asset order and skips the .core. builds' {
        $assets = @(
            [pscustomobject]@{ name = 'pbi-tools.core.1.2.0.win-x64.zip'; browser_download_url = 'https://x/core-win' },
            [pscustomobject]@{ name = 'pbi-tools.core.1.2.0.linux-x64.zip'; browser_download_url = 'https://x/core-linux' },
            [pscustomobject]@{ name = 'pbi-tools.1.2.0.zip'; browser_download_url = 'https://x/desktop' },
            [pscustomobject]@{ name = 'pbi-tools.1.2.0.zip.sha256'; browser_download_url = 'https://x/sha' }
        )
        (Select-IQPbiToolsAsset -Assets $assets).name | Should -Be 'pbi-tools.1.2.0.zip'
        (Select-IQPbiToolsAsset -Assets @($assets[2], $assets[0])).name | Should -Be 'pbi-tools.1.2.0.zip'
        (Select-IQPbiToolsAsset -Assets @([pscustomobject]@{ name = 'pbi-tools.1.0.0-beta.6.zip' })).name | Should -Be 'pbi-tools.1.0.0-beta.6.zip'
    }
    It 'Select-IQPbiToolsAsset fails closed when only core builds (or nothing) are published' {
        Select-IQPbiToolsAsset -Assets @([pscustomobject]@{ name = 'pbi-tools.core.1.2.0.win-x64.zip' }) | Should -BeNullOrEmpty
        Select-IQPbiToolsAsset -Assets @() | Should -BeNullOrEmpty
        Select-IQPbiToolsAsset -Assets $null | Should -BeNullOrEmpty
    }
    It 'the version stamp lives under State\tools (persisted by the pipeline artifact), not under Config\' {
        $path = Get-IQToolVersionStampPath
        $path | Should -Be (Join-Path (Join-Path $script:IQ.StatePath 'tools') 'tool-versions.json')
        Test-Path -LiteralPath (Split-Path -Path $path -Parent) | Should -BeTrue
        Save-IQToolVersionStamp -Path $path -Stamp @{ pbiTools = '1.2.0'; tabularEditor = '2.26.0' }
        $stamp = Get-IQToolVersionStamp -Path $path
        $stamp.pbiTools | Should -Be '1.2.0'
        $stamp.tabularEditor | Should -Be '2.26.0'
    }
}

Describe 'Power BI Desktop detection from pbi-tools info (T-02)' {
    It 'parses pbiInstalls[].location even when the token PBIDesktop never appears' {
        $json = '{"version":"1.2.0","edition":"Desktop","pbiInstalls":[{"productVersion":"2.130.754.0","location":"D:\\Apps\\Power BI\\bin","is64Bit":true}],"effectivePbiInstallDir":"D:\\Apps\\Power BI\\bin"}'
        $d = Get-IQPbiDesktopDetection -ProbeOutput $json
        $d.Found | Should -BeTrue
        $d.Source | Should -Be 'pbi-tools'
        $d.Location | Should -Be 'D:\Apps\Power BI\bin'
        $d.Edition | Should -Be 'Desktop'
    }
    It 'falls back to effectivePbiInstallDir and tolerates banner text around the JSON' {
        $out = "pbi-tools 1.2.0`n{""edition"":""Desktop"",""pbiInstalls"":[],""effectivePbiInstallDir"":""C:\\PBI\\bin""}`n"
        $d = Get-IQPbiDesktopDetection -ProbeOutput $out
        $d.Found | Should -BeTrue
        $d.Location | Should -Be 'C:\PBI\bin'
    }
    It 'reports the core edition and does not claim a Desktop install from an empty list or non-JSON output' {
        $d1 = Get-IQPbiDesktopDetection -ProbeOutput '{"edition":"Core","pbiInstalls":[],"effectivePbiInstallDir":null}'
        $d1.Edition | Should -Be 'Core'
        $d1.Source | Should -Not -Be 'pbi-tools'
        $d2 = Get-IQPbiDesktopDetection -ProbeOutput 'not json at all'
        $d2.Source | Should -Not -Be 'pbi-tools'
        if (-not $script:OnWindows) { $d1.Found | Should -BeFalse; $d2.Found | Should -BeFalse }
    }
}

Describe 'Invoke-IQProcess with a real process' {
    It 'captures the exit code and stdout into tool-logs\<Stage>' {
        $r = Invoke-IQProcess -FilePath $script:HostExe -ArgumentList '-NoProfile -NonInteractive -Command "Write-Output hello-from-child; exit 3"' -WorkingDirectory $script:Base -TimeoutMinutes 1 -LogName 'exit-three' -Stage 'ToolsTest' -Item 'exit three'
        $r.ExitCode | Should -Be 3
        $r.TimedOut | Should -BeFalse
        $r.StdOut | Should -Match 'hello-from-child'
        $r.OutFile | Should -Be (Join-Path (Get-IQToolLogFolder -Stage 'ToolsTest') 'exit-three.out.txt')
        Test-Path -LiteralPath $r.OutFile | Should -BeTrue
    }
    It 'returns a StartError result (never throws) when the executable does not exist' {
        $r = Invoke-IQProcess -FilePath (Join-Path $script:Base 'missing.exe') -ArgumentList 'x' -TimeoutMinutes 1 -LogName 'missing' -Stage 'ToolsTest'
        $r.ExitCode | Should -Be -1
        $r.StartError | Should -Match 'not found'
        $r.StdErr | Should -Match 'Start failure'
    }
    It 'Invoke-IQTabularEditor reports missing-exe without running anything' {
        $script:IQ.Tools = @{ TabularEditorPath = (Join-Path $script:Base 'nope\TabularEditor.exe') }
        $r = Invoke-IQTabularEditor -ArgumentList '"m.bim" -S "s.cs"' -Stage 'ToolsTest'
        $r.Success | Should -BeFalse
        $r.FailureReason | Should -Be 'missing-exe'
    }
}

Describe 'Complete-IQProcessJob exit-code handling (T-06)' {
    BeforeAll {
        function New-FakeProcessJob {
            param($ExitCode)
            $process = [pscustomobject]@{ ExitCode = $ExitCode }
            $process | Add-Member -MemberType ScriptMethod -Name Dispose -Value { } -Force
            $folder = Get-IQToolLogFolder -Stage 'ToolsTest'
            return @{
                ItemKey = 'k'; Item = 'fake'; FilePath = 'C:\Tools\TabularEditor.exe'; ArgumentList = ''; WorkingDirectory = $null
                LogName = 'fake'; Stage = 'ToolsTest'; OutFile = (Join-Path $folder 'fake.out.txt'); ErrFile = (Join-Path $folder 'fake.err.txt'); TimeoutMs = 60000
                Process = $process; Started = [datetime]::UtcNow; StartError = $null; TimedOut = $false; Completed = $false; Result = $null; ExitCodeUnknown = $false
            }
        }
    }
    It 'a $null ExitCode (handle not cached on Windows PowerShell 5.1) is reported as -1 with a StartError note, not as success' {
        $r = Complete-IQProcessJob -Job (New-FakeProcessJob -ExitCode $null)
        $r.ExitCode | Should -Be -1
        $r.StartError | Should -Match 'exit code unavailable'
        $r.StdErr | Should -Match 'exit code unavailable'
    }
    It 'a real 0 stays 0 and a non-zero code is kept' {
        (Complete-IQProcessJob -Job (New-FakeProcessJob -ExitCode 0)).ExitCode | Should -Be 0
        (Complete-IQProcessJob -Job (New-FakeProcessJob -ExitCode 7)).ExitCode | Should -Be 7
        (Complete-IQProcessJob -Job (New-FakeProcessJob -ExitCode 0)).StartError | Should -BeNullOrEmpty
    }
    It 'Invoke-IQTabularEditor classifies the unknown exit code as a failure' {
        $script:IQ.Tools = @{ TabularEditorPath = $script:HostExe }
        Mock Invoke-IQProcess { return @{ ExitCode = -1; TimedOut = $false; StdOut = 'Loading model'; StdErr = ''; OutFile = $null; ErrFile = $null; DurationSec = 1; StartError = 'exit code unavailable (process handle not cached)'; FilePath = $FilePath } }
        $r = Invoke-IQTabularEditor -ArgumentList 'x' -Stage 'ToolsTest'
        $r.Success | Should -BeFalse
        $r.FailureReason | Should -Be 'start-failure'
    }
}

Describe 'Invoke-IQProcessBatch pool' {
    It 'runs the jobs through the pool and returns results in job order' {
        $jobs = @(
            (New-HostJob -Command 'exit 0' -ItemKey 'a' -Item 'A' -LogName 'batch-a'),
            (New-HostJob -Command 'Write-Output second; exit 2' -ItemKey 'b' -Item 'B' -LogName 'batch-b'),
            (New-HostJob -Command 'exit 0' -ItemKey 'c' -Item 'C' -LogName 'batch-c')
        )
        $script:BatchDone = @()
        $results = @(Invoke-IQProcessBatch -Jobs $jobs -MaxParallel 2 -TimeoutMinutes 1 -Stage 'ToolsTest' -OnJobComplete { param($e) $script:BatchDone += , $e.ItemKey })
        if ($results.Count -eq 1 -and $results[0] -is [array]) { $results = @($results[0]) }
        $results.Count | Should -Be 3
        @($results | ForEach-Object { $_.ItemKey }) | Should -Be @('a', 'b', 'c')
        $results[1].Result.ExitCode | Should -Be 2
        $results[1].Result.StdOut | Should -Match 'second'
        $results[0].Result.ExitCode | Should -Be 0
        @($script:BatchDone).Count | Should -Be 3 -Because 'OnJobComplete runs once per job'
    }
    It 'a log-folder failure becomes a StartError result instead of an exception out of the pool (T-05)' {
        Mock Get-IQToolLogFolder { throw 'disk full' }
        $jobs = @((New-HostJob -Command 'exit 0' -ItemKey 'a' -Item 'A' -LogName 'nolog-a'), (New-HostJob -Command 'exit 0' -ItemKey 'b' -Item 'B' -LogName 'nolog-b'))
        $results = $null
        { $script:PoolResults = @(Invoke-IQProcessBatch -Jobs $jobs -MaxParallel 2 -TimeoutMinutes 1 -Stage 'ToolsTest') } | Should -Not -Throw
        $results = $script:PoolResults
        if ($results.Count -eq 1 -and $results[0] -is [array]) { $results = @($results[0]) }
        $results.Count | Should -Be 2
        $results[0].Result.StartError | Should -Match 'disk full'
        $results[1].Result.ExitCode | Should -Be -1
    }
}

Describe 'Invoke-IQProcessBatch cleanup on interruption (T-05)' {
    BeforeAll {
        # The pool polls with Start-Sleep 500 ms; an exception there stands in for Ctrl+C / a pipeline cancel.
        Mock Start-Sleep { throw 'simulated cancel' }
        # Record the PIDs the pool starts (the children are killed before they could print anything themselves).
        $script:RealStartJob = ${function:Start-IQProcessJob}
        $script:SleeperPids = @()
        Mock Start-IQProcessJob {
            $j = & $script:RealStartJob -FilePath $FilePath -ArgumentList $ArgumentList -WorkingDirectory $WorkingDirectory -TimeoutMinutes $TimeoutMinutes -LogName $LogName -Stage $Stage -ItemKey $ItemKey -Item $Item
            if ($j.Process) { $script:SleeperPids += [int]$j.Process.Id }
            return $j
        }
        $script:Sleepers = @(
            (New-HostJob -Command 'Start-Sleep -Seconds 40' -ItemKey 's1' -Item 'sleeper 1' -LogName 'cancel-s1'),
            (New-HostJob -Command 'Start-Sleep -Seconds 40' -ItemKey 's2' -Item 'sleeper 2' -LogName 'cancel-s2')
        )
        $script:CancelError = $null
        try { Invoke-IQProcessBatch -Jobs $script:Sleepers -MaxParallel 2 -TimeoutMinutes 1 -Stage 'CancelTest' | Out-Null } catch { $script:CancelError = $_ }
    }
    It 'the exception still propagates (it is not a per-job failure)' {
        $script:CancelError | Should -Not -BeNullOrEmpty
        $script:CancelError.Exception.Message | Should -Match 'simulated cancel'
    }
    It 'both sleeper processes were started and are no longer running afterwards' {
        $script:SleeperPids.Count | Should -Be 2 -Because 'the pool started both jobs before the interruption'
        foreach ($sleeperPid in $script:SleeperPids) {
            $alive = $false
            $deadline = [datetime]::UtcNow.AddSeconds(10)
            do {
                $p = Get-Process -Id $sleeperPid -ErrorAction SilentlyContinue
                $alive = ($null -ne $p -and -not $p.HasExited)
                if ($alive) { [System.Threading.Thread]::Sleep(250) }
            } while ($alive -and [datetime]::UtcNow -lt $deadline)
            $alive | Should -BeFalse -Because "PID $sleeperPid must have been killed by the pool's finally block"
        }
    }
}

Describe 'Install-IQToolFromZip (T-01, T-09)' {
    BeforeAll {
        # A tool root with [ ] in its path: every file operation must use literal paths.
        $script:ToolRoot = [System.IO.Directory]::CreateDirectory((Join-Path $script:Base 'Tools [v2]')).FullName
        $script:Target = Join-Path $script:ToolRoot 'PBI Tools'
        $script:TargetExe = Join-Path $script:Target 'pbi-tools.exe'
        Mock Stop-Process { throw 'Stop-Process must never be called by the installer' }
    }
    It 'extracts the nested folder into the target, keeps sub-folders, removes the zip and leaves no .old folder' {
        $zip = New-ToolZip -ZipPath (Join-Path $script:ToolRoot 'PBI Tools [manual].zip') -ExeName 'pbi-tools.exe'
        Install-IQToolFromZip -ZipPath $zip -TargetFolder $script:Target -ExeName 'pbi-tools.exe' -DisplayName 'PBI Tools' | Should -BeTrue
        Test-Path -LiteralPath $script:TargetExe | Should -BeTrue
        Test-Path -LiteralPath (Join-Path (Join-Path $script:Target 'lib') 'readme.txt') | Should -BeTrue
        Test-Path -LiteralPath $zip | Should -BeFalse
        Test-Path -LiteralPath ($script:Target + '.old') | Should -BeFalse
        Should -Invoke Stop-Process -Times 0
    }
    It 'replaces the previous folder instead of overlaying it (stale files disappear)' {
        [System.IO.File]::WriteAllText((Join-Path $script:Target 'stale.dll'), 'old')
        $zip = New-ToolZip -ZipPath (Join-Path $script:ToolRoot 'PBI Tools.zip') -ExeName 'pbi-tools.exe' -TopFolder 'pbi-tools-1.1'
        Install-IQToolFromZip -ZipPath $zip -TargetFolder $script:Target -ExeName 'pbi-tools.exe' -DisplayName 'PBI Tools' | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:Target 'stale.dll') | Should -BeFalse
        Test-Path -LiteralPath $script:TargetExe | Should -BeTrue
    }
    It 'keeps the existing copy (and the zip) when an instance is running FROM the target folder, without killing it' {
        Mock Get-Process { [pscustomobject]@{ Id = 4242; Path = $script:TargetExe; ProcessName = 'pbi-tools' } }
        [System.IO.File]::WriteAllText($script:TargetExe, 'installed copy')
        $zip = New-ToolZip -ZipPath (Join-Path $script:ToolRoot 'PBI Tools.zip') -ExeName 'pbi-tools.exe'
        Install-IQToolFromZip -ZipPath $zip -TargetFolder $script:Target -ExeName 'pbi-tools.exe' -DisplayName 'PBI Tools' | Should -BeFalse
        [System.IO.File]::ReadAllText($script:TargetExe) | Should -Be 'installed copy'
        Test-Path -LiteralPath $zip | Should -BeTrue -Because 'the update is retried on the next run'
        Should -Invoke Stop-Process -Times 0
        Remove-Item -LiteralPath $zip -Force
    }
    It 'ignores a same-named process that runs from somewhere else (the user''s own installation)' {
        Mock Get-Process { [pscustomobject]@{ Id = 4343; Path = (Join-Path $script:Base 'Elsewhere\pbi-tools.exe'); ProcessName = 'pbi-tools' } }
        @(Get-IQToolLockingProcess -ExePath $script:TargetExe).Count | Should -Be 0
        $zip = New-ToolZip -ZipPath (Join-Path $script:ToolRoot 'PBI Tools.zip') -ExeName 'pbi-tools.exe'
        Install-IQToolFromZip -ZipPath $zip -TargetFolder $script:Target -ExeName 'pbi-tools.exe' -DisplayName 'PBI Tools' | Should -BeTrue
        [System.IO.File]::ReadAllText($script:TargetExe) | Should -Match 'fake exe'
        Should -Invoke Stop-Process -Times 0
    }
    It 'reports a zip without the expected exe and leaves the target untouched' {
        $zip = New-ToolZip -ZipPath (Join-Path $script:ToolRoot 'TabularEditor.zip') -ExeName 'Other.exe'
        Install-IQToolFromZip -ZipPath $zip -TargetFolder $script:Target -ExeName 'pbi-tools.exe' -DisplayName 'PBI Tools' | Should -BeFalse
        Test-Path -LiteralPath $script:TargetExe | Should -BeTrue
    }
}

Describe 'WER DontShowUI handling (T-07)' {
    It 'Restore-IQWerSetting is a no-op when nothing was changed' {
        $script:IQ.Tools = @{ WerChanged = $false }
        { Restore-IQWerSetting } | Should -Not -Throw
    }
    It 'Initialize-IQTools leaves the setting alone in an interactive session' -Skip:(-not $script:OnWindows) {
        Mock Set-IQWerDontShowUI { }
        $script:IQ.Interactive = $true
        try { Initialize-IQTools -SkipToolUpdate | Out-Null } finally { $script:IQ.Interactive = $false }
        Should -Invoke Set-IQWerDontShowUI -Times 0
        $script:IQ.Tools.WerChanged | Should -BeFalse
    }
    It 'a headless run sets DontShowUI=1, remembers the previous value and Restore-IQWerSetting puts it back' -Skip:(-not $script:OnWindows) {
        $key = 'HKCU:\Software\Microsoft\Windows\Windows Error Reporting'
        $before = $null
        $had = $false
        if (Test-Path -LiteralPath $key) { $item = Get-ItemProperty -LiteralPath $key -Name DontShowUI -ErrorAction SilentlyContinue; if ($item) { $before = $item.DontShowUI; $had = $true } }
        if ($had -and [int]$before -eq 1) { Set-ItemProperty -LiteralPath $key -Name DontShowUI -Value 0 }
        try {
            $script:IQ.Tools = @{ WerChanged = $false; WerPrevious = $null; WerHadValue = $false }
            Set-IQWerDontShowUI
            (Get-ItemProperty -LiteralPath $key -Name DontShowUI).DontShowUI | Should -Be 1
            $script:IQ.Tools.WerChanged | Should -BeTrue
            Restore-IQWerSetting
            $after = Get-ItemProperty -LiteralPath $key -Name DontShowUI -ErrorAction SilentlyContinue
            if ($had) { [int]$after.DontShowUI | Should -Be ([int]$(if ([int]$before -eq 1) { 0 } else { $before })) } else { $after | Should -BeNullOrEmpty }
        }
        finally {
            if ($had) { Set-ItemProperty -LiteralPath $key -Name DontShowUI -Value $before } else { Remove-ItemProperty -LiteralPath $key -Name DontShowUI -ErrorAction SilentlyContinue }
        }
    }
}

Describe 'Initialize-IQTools off Windows' -Skip:$script:OnWindows {
    It 'reports not-windows for both tools and never downloads' {
        Mock Invoke-RestMethod { throw 'no network in tests' }
        Mock Invoke-WebRequest { throw 'no network in tests' }
        $tools = Initialize-IQTools
        $tools.TabularEditorWorks | Should -BeFalse
        $tools.PbiToolsWorks | Should -BeFalse
        $tools.TabularEditorPreflight | Should -Be 'not-windows'
        $tools.ToolUpdateSkipped | Should -BeTrue
        $tools.WerChanged | Should -BeFalse
        $script:IQ.Tools.PbiToolsPath | Should -Be (Join-Path (Join-Path $script:IQ.ConfigFolder 'PBI Tools') 'pbi-tools.exe')
        Should -Invoke Invoke-RestMethod -Times 0
    }
}
