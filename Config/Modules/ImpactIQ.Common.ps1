#Requires -Version 5.1
<#
.SYNOPSIS
    ImpactIQ v3 - Common module: shared context, logging, names, retry helper, JSON helpers, environments.

.DESCRIPTION
    Dot-sourced by ImpactIQ.ps1 (and by the Pester tests) before every other module. All shared state lives in
    ONE hashtable, $script:IQ, created by Initialize-IQContext. Because the modules are dot-sourced, "$script:"
    here is the scope of the script that dot-sourced them (ImpactIQ.ps1 or a test file).

    Windows PowerShell 5.1 compatible (no ternary, no "??", no "?.", no "using namespace").
    Never uses $IsWindows (undefined on 5.1) - use $script:IQ.IsWindows instead.
#>

# The shared context lives in the dot-sourcing script's scope; make sure the variable exists (strict-mode safe).
if (-not (Get-Variable -Name IQ -Scope Script -ErrorAction SilentlyContinue)) { $script:IQ = $null }

function Initialize-IQContext {
    <#
    .SYNOPSIS
        Creates the shared $script:IQ context hashtable and the Config, State and Logs folders.
    .PARAMETER BaseFolder
        Root folder of the deployment (contains Config\, and receives Model/Report/Dataflow Backups, State, Logs).
    .PARAMETER Options
        Hashtable of all entry-point parameters by name (no secrets are ever logged from it). Recognised keys used
        here: Environment, LogPath, NonInteractive, Verbose/Debug (host echo of Debug lines), AuthorityOverride,
        FabricApiPrefixOverride, NowUtc (test-only clock override, [datetime] or ISO string).
    .OUTPUTS
        The context hashtable (also stored in $script:IQ).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$BaseFolder,
        [Parameter(Mandatory = $false)][hashtable]$Options
    )

    if ($null -eq $Options) { $Options = @{} }
    $BaseFolder = [System.IO.Path]::GetFullPath($BaseFolder)

    $configFolder = Join-Path $BaseFolder 'Config'
    $statePath = Join-Path $BaseFolder 'State'
    $logsPath = Join-Path $BaseFolder 'Logs'
    foreach ($folder in @($BaseFolder, $configFolder, $statePath, $logsPath)) {
        if (-not (Test-Path -LiteralPath $folder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        }
    }

    $logFile = $null
    if (-not [string]::IsNullOrWhiteSpace([string]$Options['LogPath'])) {
        $logFile = [string]$Options['LogPath']
        if (-not [System.IO.Path]::IsPathRooted($logFile)) { $logFile = Join-Path $BaseFolder $logFile }
        $logDir = Split-Path -Path $logFile -Parent
        if ($logDir -and -not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    }
    else {
        $logFile = Join-Path $logsPath ('ImpactIQ_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.log')
    }

    $onWindows = ($env:OS -eq 'Windows_NT')
    $isAzureDevOps = ($env:TF_BUILD -eq 'True')

    $script:IQ = @{
        BaseFolder    = $BaseFolder
        ConfigFolder  = $configFolder
        StatePath     = $statePath
        LogsPath      = $logsPath
        LogFile       = $logFile
        IsWindows     = $onWindows
        IsAzureDevOps = $isAzureDevOps
        Interactive   = $false
        Options       = $Options
        Environment   = $null
        Endpoints     = $null
        Auth          = @{ Mode = $null }
        Tools         = @{ TabularEditorPath = $null; PbiToolsPath = $null; TabularEditorWorks = $false; PbiToolsWorks = $false }
        RunId         = $null
        RunPath       = $null
        IsResume      = $false
        Manifest      = $null
        Scope         = $null
        CurrentStage  = $null
        Paths         = @{
            ModelBackups    = Join-Path $BaseFolder 'Model Backups'
            ReportBackups   = Join-Path $BaseFolder 'Report Backups'
            DataflowBackups = Join-Path $BaseFolder 'Dataflow Backups'
            TempExtract     = Join-Path $configFolder 'Temp'
        }
        Stats         = @{ ApiCalls = 0; Retries = 0 }
    }

    $script:IQ.Interactive = Test-IQInteractive

    $envName = [string]$Options['Environment']
    if (-not [string]::IsNullOrWhiteSpace($envName)) {
        Set-IQEnvironment -Environment $envName | Out-Null
    }

    Write-IQLog -Level Debug -Message ("Context initialised. BaseFolder='{0}' IsWindows={1} IsAzureDevOps={2} Interactive={3} PS={4}" -f `
            $BaseFolder, $onWindows, $isAzureDevOps, $script:IQ.Interactive, $PSVersionTable.PSVersion)
    return $script:IQ
}

function Get-IQContext {
    <#
    .SYNOPSIS
        Returns the shared $script:IQ context hashtable ($null before Initialize-IQContext).
    #>
    [CmdletBinding()]
    param()
    return $script:IQ
}

function Set-IQEnvironment {
    <#
    .SYNOPSIS
        Sets $script:IQ.Environment and $script:IQ.Endpoints from the environment table (convenience for the entry point).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Environment)
    $settings = Get-IQEnvironmentSettings -Environment $Environment
    if ($script:IQ) {
        $script:IQ.Environment = $settings.Name
        $script:IQ.Endpoints = $settings
    }
    return $settings
}

function Get-IQNowUtc {
    <#
    .SYNOPSIS
        Current UTC time, or the injectable test clock ($script:IQ.Options.NowUtc) when set.
    #>
    [CmdletBinding()]
    param()
    $override = $null
    if ($script:IQ -and $script:IQ.Options -and $script:IQ.Options.Contains('NowUtc')) { $override = $script:IQ.Options['NowUtc'] }
    if ($null -ne $override -and -not ($override -is [string] -and [string]::IsNullOrWhiteSpace($override))) {
        if ($override -is [datetime]) {
            return [datetime]::SpecifyKind($override, [System.DateTimeKind]::Utc)
        }
        $parsed = [datetime]::Parse([string]$override, [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AdjustToUniversal -bor [System.Globalization.DateTimeStyles]::AssumeUniversal)
        return [datetime]::SpecifyKind($parsed, [System.DateTimeKind]::Utc)
    }
    return [datetime]::UtcNow
}

function ConvertTo-IQRedactedText {
    <#
    .SYNOPSIS
        Masks secrets (connection-string passwords, bearer tokens, client secrets) in a string before it is logged.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    $t = $Text
    $t = [regex]::Replace($t, '(?i)(Password\s*=)[^;"]+', '$1***')
    $t = [regex]::Replace($t, '(?i)(Bearer\s+)[A-Za-z0-9\-_\.=]+', '$1***')
    $t = [regex]::Replace($t, '(?i)((?:refresh_token|access_token|id_token|client_secret|password)["'']?\s*[=:]\s*["'']?)[A-Za-z0-9\-_\.=%+/]+', '$1***')
    return $t
}

function Write-IQLog {
    <#
    .SYNOPSIS
        Writes "[HH:mm:ss] [LEVEL] [Stage] [Item] message" to the host (coloured) and appends it to the run log. Never throws.
    .DESCRIPTION
        Debug lines always go to the log file; they are echoed to the host only when Options.Verbose/Options.Debug or
        $env:IMPACTIQ_DEBUG=1. On Azure DevOps ($env:TF_BUILD -eq 'True') Warn/Error lines also emit
        ##vso[task.logissue type=warning|error] so they appear in the pipeline issues summary.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)][AllowNull()][AllowEmptyString()][string]$Message,
        [Parameter(Mandatory = $false)][ValidateSet('Info', 'Warn', 'Error', 'Debug', 'Success')][string]$Level = 'Info',
        [Parameter(Mandatory = $false)][string]$Stage,
        [Parameter(Mandatory = $false)][string]$Item,
        [Parameter(Mandatory = $false)][System.Exception]$Exception
    )
    try {
        $ctx = $script:IQ
        if ([string]::IsNullOrEmpty($Stage) -and $ctx -and $ctx.CurrentStage) { $Stage = [string]$ctx.CurrentStage }

        $text = ConvertTo-IQRedactedText -Text ([string]$Message)
        $exceptionText = $null
        if ($null -ne $Exception) {
            $exceptionText = ConvertTo-IQRedactedText -Text $Exception.Message
            if (-not [string]::IsNullOrEmpty($exceptionText)) { $text = $text + ' :: ' + $exceptionText }
        }

        $parts = New-Object System.Collections.Generic.List[string]
        $parts.Add('[' + (Get-Date -Format 'HH:mm:ss') + ']')
        $parts.Add('[' + $Level.ToUpperInvariant() + ']')
        if (-not [string]::IsNullOrEmpty($Stage)) { $parts.Add('[' + $Stage + ']') }
        if (-not [string]::IsNullOrEmpty($Item)) { $parts.Add('[' + $Item + ']') }
        $parts.Add($text)
        $line = [string]::Join(' ', $parts.ToArray())

        # --- log file (always, including Debug) ---
        $logFile = $null
        if ($ctx) { $logFile = $ctx.LogFile }
        if (-not [string]::IsNullOrEmpty($logFile)) {
            try {
                $fileLine = $line
                if ($null -ne $Exception -and $Level -in @('Error', 'Warn')) {
                    $detail = $Exception.GetType().FullName
                    if ($Exception.StackTrace) { $detail = $detail + [Environment]::NewLine + $Exception.StackTrace }
                    $fileLine = $fileLine + [Environment]::NewLine + '    ' + (ConvertTo-IQRedactedText -Text $detail)
                }
                $enc = New-Object System.Text.UTF8Encoding($false)
                [System.IO.File]::AppendAllText($logFile, $fileLine + [Environment]::NewLine, $enc)
            }
            catch { $null = $_.Exception }
        }

        # --- host ---
        $echoDebug = $false
        if ($Level -eq 'Debug') {
            if ($env:IMPACTIQ_DEBUG -eq '1') { $echoDebug = $true }
            elseif ($ctx -and $ctx.Options) {
                if ($ctx.Options['Verbose'] -or $ctx.Options['Debug'] -or $ctx.Options['LogDebug']) { $echoDebug = $true }
            }
        }
        if ($Level -ne 'Debug' -or $echoDebug) {
            $colour = 'White'
            switch ($Level) {
                'Warn' { $colour = 'Yellow' }
                'Error' { $colour = 'Red' }
                'Debug' { $colour = 'DarkGray' }
                'Success' { $colour = 'Green' }
                default { $colour = 'White' }
            }
            try { Write-Host $line -ForegroundColor $colour } catch { try { Write-Host $line } catch { $null = $_.Exception } }
        }

        # --- Azure DevOps logging commands ---
        $isAdo = $false
        if ($ctx) { $isAdo = [bool]$ctx.IsAzureDevOps } else { $isAdo = ($env:TF_BUILD -eq 'True') }
        if ($isAdo -and ($Level -eq 'Warn' -or $Level -eq 'Error')) {
            $vsoType = 'warning'
            if ($Level -eq 'Error') { $vsoType = 'error' }
            $vsoText = $text -replace '\r', '%0D' -replace '\n', '%0A' -replace ';', '%3B' -replace ']', '%5D'
            $vsoPrefix = ''
            if (-not [string]::IsNullOrEmpty($Stage)) { $vsoPrefix = '[' + $Stage + '] ' }
            if (-not [string]::IsNullOrEmpty($Item)) { $vsoPrefix = $vsoPrefix + '[' + $Item + '] ' }
            try { Write-Host ('##vso[task.logissue type=' + $vsoType + ']' + $vsoPrefix + $vsoText) } catch { $null = $_.Exception }
        }
    }
    catch {
        # Logging must never take the run down.
        $null = $_.Exception
    }
}

function Get-IQCleanName {
    <#
    .SYNOPSIS
        The exact original file-name sanitiser: [ ] -> ( ), then every char outside [a-zA-Z0-9()&,.-] -> space, then TrimStart().
    .DESCRIPTION
        Used for every backup file name ("<CleanWs> ~ <CleanModel>.bim", "<CleanWs> ~ <CleanReport>.pbix", ...).
        Must not change: the PBIT and the csx scripts depend on it.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false, Position = 0)][AllowNull()][AllowEmptyString()][string]$Name)
    if ($null -eq $Name) { return '' }
    $clean = $Name -replace '\[', '(' -replace '\]', ')'
    $clean = $clean -replace "[^a-zA-Z0-9\(\)&,.-]", " "
    $clean = $clean.TrimStart()
    return $clean
}

function Get-IQSafeKey {
    <#
    .SYNOPSIS
        File-name-safe key for checkpoint/inventory files: lowercase, [^a-z0-9._-] -> "_", max 120 chars.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false, Position = 0)][AllowNull()][AllowEmptyString()][string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '_' }
    $key = $Value.ToLowerInvariant() -replace '[^a-z0-9._-]', '_'
    if ($key.Length -gt 120) { $key = $key.Substring(0, 120) }
    if ([string]::IsNullOrEmpty($key.Trim('.', ' '))) { $key = '_' }
    return $key
}

function Invoke-IQWithRetry {
    <#
    .SYNOPSIS
        Runs a script block, retrying on exception with exponential backoff (capped at 60 s); rethrows after the last attempt.
    .PARAMETER RetryOn
        Optional filter: receives the ErrorRecord as $_ (and $args[0]); return $true to retry, $false to rethrow immediately.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)][scriptblock]$ScriptBlock,
        [Parameter(Mandatory = $false)][ValidateRange(1, 100)][int]$MaxAttempts = 5,
        [Parameter(Mandatory = $false)][ValidateRange(0, 3600)][int]$InitialDelaySeconds = 2,
        [Parameter(Mandatory = $false)][scriptblock]$RetryOn,
        [Parameter(Mandatory = $false)][string]$Description = 'operation'
    )
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return (& $ScriptBlock)
        }
        catch {
            $err = $_
            $shouldRetry = $true
            if ($null -ne $RetryOn) {
                try {
                    $verdict = @($err | ForEach-Object -Process $RetryOn) | Select-Object -Last 1
                    $shouldRetry = [bool]$verdict
                }
                catch { $shouldRetry = $false }
            }
            if (-not $shouldRetry -or $attempt -ge $MaxAttempts) {
                if ($attempt -ge $MaxAttempts -and $shouldRetry) {
                    Write-IQLog -Level Warn -Message ("{0} failed after {1} attempt(s): {2}" -f $Description, $attempt, $err.Exception.Message)
                }
                throw $err
            }
            $delay = [double]$InitialDelaySeconds * [math]::Pow(2, ($attempt - 1))
            if ($delay -gt 60) { $delay = 60 }
            if ($script:IQ -and $script:IQ.Stats) { $script:IQ.Stats.Retries = [int]$script:IQ.Stats.Retries + 1 }
            Write-IQLog -Level Warn -Message ("{0} failed (attempt {1}/{2}): {3}. Retrying in {4} s." -f $Description, $attempt, $MaxAttempts, $err.Exception.Message, [int]$delay)
            if ($delay -gt 0) { Start-Sleep -Seconds ([int]$delay) }
        }
    }
}

function ConvertTo-IQJsonFile {
    <#
    .SYNOPSIS
        Atomically writes an object as JSON (UTF-8 without BOM, -Depth 20): writes "<path>.tmp" then Move-Item -Force.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)][AllowNull()]$Object,
        [Parameter(Mandatory = $true, Position = 1)][string]$Path
    )
    $dir = Split-Path -Path $Path -Parent
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $json = $null
    if ($null -eq $Object) { $json = 'null' }
    else { $json = ConvertTo-Json -InputObject $Object -Depth 20 }
    $tmp = $Path + '.tmp'
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($tmp, $json, $enc)
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}

function ConvertFrom-IQJsonFile {
    <#
    .SYNOPSIS
        Reads a JSON file into objects; returns $null when the file is missing or empty.
    .DESCRIPTION
        A top-level JSON array is emitted element by element (standard PowerShell semantics on both 5.1 and 7), so
        wrap the call in @( ) when you need an array: $rows = @(ConvertFrom-IQJsonFile -Path $p).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true, Position = 0)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $text = [System.IO.File]::ReadAllText($Path)
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $obj = ConvertFrom-Json -InputObject $text
    return $obj
}

function Get-IQDateFolder {
    <#
    .SYNOPSIS
        Full path of the newest "yyyy-MM-dd" sub-folder under Root (the monolith's "latest dated folder" rule), or $null.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true, Position = 0)][string]$Root)
    if (-not (Test-Path -LiteralPath $Root)) { return $null }
    $candidates = @(Get-ChildItem -LiteralPath $Root -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}-\d{2}-\d{2}$' })
    $best = $null
    $bestDate = [datetime]::MinValue
    foreach ($c in $candidates) {
        $d = [datetime]::MinValue
        if ([datetime]::TryParseExact($c.Name, 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None, [ref]$d)) {
            if ($d -gt $bestDate) { $bestDate = $d; $best = $c }
        }
    }
    if ($null -eq $best) { return $null }
    return $best.FullName
}

function Test-IQInteractive {
    <#
    .SYNOPSIS
        $true only when a human can respond: not -NonInteractive, UserInteractive, not TF_BUILD/CI, and a console host.
    #>
    [CmdletBinding()]
    param()
    if ($script:IQ -and $script:IQ.Options -and $script:IQ.Options['NonInteractive']) { return $false }
    if (-not [Environment]::UserInteractive) { return $false }
    if (-not [string]::IsNullOrEmpty($env:TF_BUILD)) { return $false }
    $ci = [string]$env:CI
    if (-not [string]::IsNullOrEmpty($ci) -and $ci -notin @('false', '0', 'False', 'FALSE')) { return $false }
    if ($null -eq $Host -or $Host.Name -eq 'ServerRemoteHost') { return $false }
    return $true
}

function Get-IQEnvironmentSettings {
    <#
    .SYNOPSIS
        Endpoint table for a Power BI cloud (Public, USGov, USGovHigh, USGovMil, China, Germany), with optional overrides.
    .DESCRIPTION
        Corrections versus the monolith: USGovHigh/USGovMil PowerBI resource hosts are high.analysis.usgovcloudapi.net /
        mil.analysis.usgovcloudapi.net (the monolith had the host order wrong), USGovMil API host is api.mil.powerbigov.us,
        China authority is login.chinacloudapi.cn, China WebPrefix typo fixed. Overrides: -AuthorityOverride /
        -FabricApiPrefixOverride parameters, else $script:IQ.Options.AuthorityOverride / FabricApiPrefixOverride.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)][string]$Environment,
        [Parameter(Mandatory = $false)][string]$AuthorityOverride,
        [Parameter(Mandatory = $false)][string]$FabricApiPrefixOverride
    )

    $name = $null
    switch -Regex ($Environment.Trim()) {
        '^(Public|Commercial|Global)$' { $name = 'Public' }
        '^(USGov|GCC|USGovernment)$' { $name = 'USGov' }
        '^(USGovHigh|GCCHigh|USGovDoDCon|DoDCON)$' { $name = 'USGovHigh' }
        '^(USGovMil|USGovDoD|DoD)$' { $name = 'USGovMil' }
        '^China$' { $name = 'China' }
        '^Germany$' { $name = 'Germany' }
        default { $name = $null }
    }
    if ($null -eq $name) {
        throw "Unknown Power BI environment '$Environment'. Valid values: Public, USGov, USGovHigh, USGovMil, China, Germany."
    }

    $s = @{
        Name            = $name
        ApiPrefix       = 'https://api.powerbi.com'
        Authority       = 'https://login.microsoftonline.com'
        PowerBIResource = 'https://analysis.windows.net/powerbi/api'
        XmlaPrefix      = 'powerbi://api.powerbi.com'
        FabricApiPrefix = 'https://api.fabric.microsoft.com'
        AzEnvironment   = 'AzureCloud'
        WebPrefix       = 'https://app.powerbi.com'
        FabricVerified  = $true
    }
    switch ($name) {
        'USGov' {
            $s.ApiPrefix = 'https://api.powerbigov.us'
            $s.Authority = 'https://login.microsoftonline.com'          # GCC tenants live in commercial Entra
            $s.PowerBIResource = 'https://analysis.usgovcloudapi.net/powerbi/api'
            $s.XmlaPrefix = 'powerbi://api.powerbigov.us'
            $s.FabricApiPrefix = 'https://api.fabric.microsoft.us'      # unverified; Fabric not GA in GCC
            $s.AzEnvironment = 'AzureCloud'
            $s.WebPrefix = 'https://app.powerbigov.us'
            $s.FabricVerified = $false
        }
        'USGovHigh' {
            $s.ApiPrefix = 'https://api.high.powerbigov.us'
            $s.Authority = 'https://login.microsoftonline.us'
            $s.PowerBIResource = 'https://high.analysis.usgovcloudapi.net/powerbi/api'
            $s.XmlaPrefix = 'powerbi://api.high.powerbigov.us'
            $s.FabricApiPrefix = 'https://api.fabric.high.microsoft.us' # unverified
            $s.AzEnvironment = 'AzureUSGovernment'
            $s.WebPrefix = 'https://app.high.powerbigov.us'
            $s.FabricVerified = $false
        }
        'USGovMil' {
            $s.ApiPrefix = 'https://api.mil.powerbigov.us'
            $s.Authority = 'https://login.microsoftonline.us'
            $s.PowerBIResource = 'https://mil.analysis.usgovcloudapi.net/powerbi/api'
            $s.XmlaPrefix = 'powerbi://api.mil.powerbigov.us'
            $s.FabricApiPrefix = 'https://api.fabric.mil.microsoft.us'  # unverified
            $s.AzEnvironment = 'AzureUSGovernment'
            $s.WebPrefix = 'https://app.mil.powerbigov.us'
            $s.FabricVerified = $false
        }
        'China' {
            $s.ApiPrefix = 'https://api.powerbi.cn'
            $s.Authority = 'https://login.chinacloudapi.cn'
            $s.PowerBIResource = 'https://analysis.chinacloudapi.cn/powerbi/api'
            $s.XmlaPrefix = 'powerbi://api.powerbi.cn'
            $s.FabricApiPrefix = 'https://api.fabric.microsoft.cn'
            $s.AzEnvironment = 'AzureChinaCloud'
            $s.WebPrefix = 'https://app.powerbi.cn'
        }
        'Germany' {
            $s.ApiPrefix = 'https://api.powerbi.de'
            $s.Authority = 'https://login.microsoftonline.de'
            $s.PowerBIResource = 'https://analysis.cloudapi.de/powerbi/api'
            $s.XmlaPrefix = 'powerbi://api.powerbi.de'
            $s.FabricApiPrefix = 'https://api.fabric.microsoft.de'
            $s.AzEnvironment = 'AzureGermanCloud'
            $s.WebPrefix = 'https://app.powerbi.de'
        }
    }

    # Overrides: explicit parameters win, then the entry-point options.
    $authOverride = $AuthorityOverride
    $fabricOverride = $FabricApiPrefixOverride
    if ([string]::IsNullOrWhiteSpace($authOverride) -and $script:IQ -and $script:IQ.Options) { $authOverride = [string]$script:IQ.Options['AuthorityOverride'] }
    if ([string]::IsNullOrWhiteSpace($fabricOverride) -and $script:IQ -and $script:IQ.Options) { $fabricOverride = [string]$script:IQ.Options['FabricApiPrefixOverride'] }
    if (-not [string]::IsNullOrWhiteSpace($authOverride)) { $s.Authority = $authOverride.TrimEnd('/') }
    if (-not [string]::IsNullOrWhiteSpace($fabricOverride)) { $s.FabricApiPrefix = $fabricOverride.TrimEnd('/') }

    # Derived / legacy names (the monolith's Get-PowerBIEndpoints keys, kept for moved code).
    $s.MicrosoftPowerBIMgmtEnvironment = $name
    $s.FabricResource = $s.FabricApiPrefix
    $s.LoginUrl = $s.Authority
    $s.ResourceUrl = $s.PowerBIResource
    $s.EmbedUrl = $s.WebPrefix + '/reportEmbed'
    $s.IsCommercialAuthority = ($s.Authority -like 'https://login.microsoftonline.com*')
    $s.DeviceLoginUrl = $null
    if ($s.IsCommercialAuthority) { $s.DeviceLoginUrl = 'https://microsoft.com/devicelogin' }
    return $s
}

function Get-IQExitCode {
    <#
    .SYNOPSIS
        Maps a run manifest to the process exit code: 0 Completed, 2 CompletedWithErrors (item/non-fatal stage failures), 1 fatal.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false, Position = 0)][AllowNull()]$Manifest)
    if ($null -eq $Manifest) { return 1 }
    $status = [string](Get-IQMemberValue -Object $Manifest -Name 'status')
    switch ($status) {
        'Completed' {
            # Defensive: a "Completed" run that still carries failures or non-completed stages is an errors run.
            $failures = @(Get-IQMemberValue -Object $Manifest -Name 'failures')
            if ($failures.Count -gt 0) { return 2 }
            $stages = Get-IQMemberValue -Object $Manifest -Name 'stages'
            foreach ($st in @(Get-IQMemberValueList -Object $stages)) {
                $sStatus = [string](Get-IQMemberValue -Object $st -Name 'status')
                if ($sStatus -in @('Failed', 'CompletedWithErrors')) { return 2 }
            }
            return 0
        }
        'CompletedWithErrors' { return 2 }
        default { return 1 }
    }
}

function Get-IQMemberValue {
    <#
    .SYNOPSIS
        Reads a named member from a hashtable/dictionary or an object property (works for manifests loaded either way).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Object,
        [Parameter(Mandatory = $true)][string]$Name
    )
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -ne $prop) { return $prop.Value }
    return $null
}

function Get-IQMemberValueList {
    <#
    .SYNOPSIS
        Emits the values of a hashtable/dictionary or of a PSCustomObject's properties (nothing when null); wrap in @( ).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Object)
    $out = @()
    if ($null -eq $Object) { return $out }
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($v in $Object.Values) { $out += , $v }
        return $out
    }
    foreach ($p in $Object.PSObject.Properties) { $out += , $p.Value }
    return $out
}

function ConvertTo-IQHashtable {
    <#
    .SYNOPSIS
        Recursively converts PSCustomObjects (e.g. from ConvertFrom-Json) into ordered hashtables; arrays stay arrays.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false, Position = 0)][AllowNull()]$InputObject)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [string] -or $InputObject.GetType().IsPrimitive -or $InputObject -is [datetime] -or $InputObject -is [decimal]) { return $InputObject }
    if ($InputObject -is [System.Collections.IDictionary]) {
        $h = [ordered]@{}
        foreach ($k in @($InputObject.Keys)) { $h[[string]$k] = ConvertTo-IQHashtable -InputObject $InputObject[$k] }
        return $h
    }
    if ($InputObject -is [System.Collections.IEnumerable]) {
        $list = @()
        foreach ($i in $InputObject) { $list += , (ConvertTo-IQHashtable -InputObject $i) }
        return , $list
    }
    if ($InputObject -is [System.Management.Automation.PSCustomObject]) {
        $h = [ordered]@{}
        foreach ($p in $InputObject.PSObject.Properties) { $h[$p.Name] = ConvertTo-IQHashtable -InputObject $p.Value }
        return $h
    }
    return $InputObject
}
