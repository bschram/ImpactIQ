#Requires -Version 5.1
<#
.SYNOPSIS
    ImpactIQ v3 entry point: backs up and documents Power BI / Fabric workspaces, models, reports and dataflows -
    headless-safe, resumable and modular (Windows PowerShell 5.1 and PowerShell 7).

.DESCRIPTION
    Runs the stages Inventory, ModelBackup, ReportBackup, ReportDetail, ModelDetail, Dataflows, Extras and Assemble
    (in that order) against the workspaces / reports / models selected by the scope parameters, writing every unit of
    work to a checkpoint under <BaseFolder>\State\runs\<RunId>\ the moment it completes so an interrupted run can be
    resumed (-Resume Auto is the default). The four workbooks ("Power BI Environment Detail.xlsx", "Report Detail.xlsx",
    "Model Detail.xlsx", "Dataflow Detail.xlsx") and the backup folders keep their v2 names and layout.

    Interactive runs (a console session, no -NonInteractive, not under Azure DevOps / CI) keep the v2 experience:
    environment dialog, browser sign-in, run-mode dialog and workspace / report / model pickers. Headless runs never
    open a dialog: every prompt has a parameter or IMPACTIQ_* environment-variable equivalent (docs/Headless-and-Resume.md).

    Exit codes: 0 = every stage Completed (or the user cancelled a dialog); 2 = finished with item failures or a
    non-fatal stage failure (see manifest.failures / the Failures sheet); 3 = Paused because -TimeBudgetMinutes was
    reached (the next start resumes it); 1 = fatal (authentication, no scope, Inventory failure, unhandled
    exception). The manifest (State\runs\<RunId>\manifest.json) describes what succeeded, what failed and why.

.PARAMETER BaseFolder
    Root folder: contains Config\ (csx scripts, Blank Model.bim, TabularEditor\, PBI Tools\, Modules\) and receives
    Model Backups\, Report Backups\, Dataflow Backups\, State\, Logs\ and the workbooks. Default: the folder of this
    script when it contains Config\, else 'C:\Power BI Backups'.
.PARAMETER Environment
    Power BI cloud: Public, USGov (GCC), USGovHigh, USGovMil, China, Germany. Default: IMPACTIQ_ENVIRONMENT, else the
    interactive dialog, else Public.
.PARAMETER AuthMode
    Auto | Interactive | DeviceCode | Credential | AzContext | AccessToken (docs/Auth-Options.md). Auto picks
    Credential (IMPACTIQ_USERNAME/PASSWORD or -Credential) -> AccessToken (IMPACTIQ_PBI_TOKEN) -> DeviceCode with a
    cached refresh token -> Interactive (console session) -> DeviceCode.
.PARAMETER TenantId
    Tenant id or domain for device-code / password sign-in (default 'organizations').
.PARAMETER ClientId
    Public client id for device-code / password sign-in (default: Azure PowerShell 1950a258-227b-4e31-a9cf-717495945fc2).
.PARAMETER Credential
    User name / password for Credential mode (or IMPACTIQ_USERNAME / IMPACTIQ_PASSWORD). MFA-bound accounts cannot use it.
.PARAMETER TokenCachePath
    Encrypted refresh-token cache for DeviceCode mode (default <BaseFolder>\State\auth\token-cache.json).
.PARAMETER TokenCacheKey
    AES-256 pass phrase for the token cache (or IMPACTIQ_TOKEN_CACHE_KEY); without it Windows uses DPAPI.
.PARAMETER DeviceCodeWebhookUrl
    Teams / Slack incoming webhook that receives the device-code message (or IMPACTIQ_DEVICECODE_WEBHOOK).
.PARAMETER AuthorityOverride
    OAuth authority URL override (sovereign clouds / testing).
.PARAMETER FabricApiPrefixOverride
    Fabric REST API prefix override (sovereign clouds where the Fabric endpoint is unverified).
.PARAMETER NonInteractive
    Never show a dialog, never call Read-Host, never open a browser. Automatically on under TF_BUILD / CI.
.PARAMETER RunMode
    Workspaces (default), Reports (-ReportId required) or Models (-DatasetId required).
.PARAMETER WorkspaceId
    Workspace ids to include (Workspaces mode; in Reports / Models mode the workspaces to scan).
.PARAMETER WorkspaceName
    Workspace names or -like wildcards to include.
.PARAMETER AllWorkspaces
    Include every accessible workspace (Workspaces mode).
.PARAMETER IncludeMyWorkspace
    Also process the personal "My Workspace".
.PARAMETER ReportId
    Report ids (Reports mode). Their datasets and dataset workspaces are included automatically.
.PARAMETER DatasetId
    Semantic model ids (Models mode). Every accessible report using them is included automatically.
.PARAMETER Stages
    Only run these stages (canonical order is kept): Inventory, ModelBackup, ReportBackup, ReportDetail, ModelDetail,
    Dataflows, Extras, Assemble - as an array or one comma/semicolon-separated string (powershell.exe -File form).
    Default: all. A list without Inventory runs on top of an existing run: today's run, else (without -RunId) the
    newest run that has a manifest; when none exists the script stops before touching any folder.
.PARAMETER SkipStages
    Stages to skip.
.PARAMETER RunId
    Run / backup folder name (default today's yyyy-MM-dd).
.PARAMETER Resume
    Auto (default: resume an unfinished run), Always (resume <RunId> if a manifest exists), Never (fresh run).
.PARAMETER ResumeMaxAgeDays
    How far back -Resume Auto looks for an unfinished run when -RunId is not given (default 3).
.PARAMETER Force
    Delete State\runs\<RunId> and the three <RunId> backup folders, then start fresh.
.PARAMETER RefreshInventory
    On a resumed run, re-run the Inventory stage even though it completed.
.PARAMETER ModelDetailMethod
    Auto (default) | TabularEditor | Bim | Dax | Both - how Model Detail is produced. Auto tries the Tabular Editor csx
    scripts (when TE2 works and a .bim exists), then the built-in .bim (TMSL) parser, then DAX INFO.* over
    executeQueries; Bim = parser only; Dax = DAX only; Both = Tabular Editor then the fallbacks on failure.
.PARAMETER MaxParallelExtracts
    Concurrent Tabular Editor / pbi-tools processes (default 2).
.PARAMETER ToolTimeoutMinutes
    Timeout per external process (default 20; ReportDetail uses three times this).
.PARAMETER MaxRetries
    HTTP retries for 5xx / network errors (default 5).
.PARAMETER DefinitionTimeoutMinutes
    Timeout for Fabric getDefinition long-running operations (report / dataflow / semantic-model definitions; default 10).
.PARAMETER TimeBudgetMinutes
    Stop cleanly after this many minutes (0 = unlimited). Between items every stage checks the elapsed time of this
    process; when the budget minus a 2-minute grace is reached the stage stops, later stages are deferred, Assemble
    still builds the (partial) workbooks, the manifest is marked Paused and the exit code is 3. The next start resumes
    the run. Use 55 on free Microsoft-hosted Azure DevOps agents (60-minute job cap).
.PARAMETER SkipToolUpdate
    Do not download Tabular Editor 2 / pbi-tools updates (same as IMPACTIQ_OFFLINE=1).
.PARAMETER IncludeAdminApis
    Extras stage: admin groups, Scanner API and activity events (Fabric administrators only; probed, else skipped).
.PARAMETER IncludeUsageMetrics
    Extras stage: per-workspace usage metrics via DAX.
.PARAMETER ActivityDays
    Days of admin/activityevents to collect with -IncludeAdminApis (default 30, API window is 28).
.PARAMETER LogPath
    Log file (default <BaseFolder>\Logs\ImpactIQ_yyyyMMdd_HHmmss.log).
.PARAMETER PassThru
    Also return the run manifest object.

.EXAMPLE
    .\ImpactIQ.ps1
    Interactive run with the v2 dialogs (environment, sign-in, run mode, pickers).
.EXAMPLE
    .\ImpactIQ.ps1 -NonInteractive -Environment USGov -AuthMode DeviceCode -AllWorkspaces -IncludeMyWorkspace
    Unattended run against everything the account can see; resumes automatically when yesterday's run died.
.EXAMPLE
    .\ImpactIQ.ps1 -NonInteractive -Environment USGov -RunMode Reports -ReportId <guid> -Stages Inventory,ReportBackup,ReportDetail,Assemble
.EXAMPLE
    .\ImpactIQ.ps1 -NonInteractive -Resume Always -Stages Assemble
    Rebuild the four workbooks from the latest run's checkpoints (no sign-in needed).

.NOTES
    Automation (Azure DevOps, Task Scheduler): call this script directly; see docs/Automation.md. The interactive
    launcher "Final PS Script.txt" only wraps it with -AuthMode Interactive.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)][string]$BaseFolder,
    [Parameter(Mandatory = $false)][ValidateSet('Public', 'Germany', 'USGov', 'China', 'USGovHigh', 'USGovMil')][string]$Environment,
    [Parameter(Mandatory = $false)][ValidateSet('Auto', 'Interactive', 'DeviceCode', 'Credential', 'AzContext', 'AccessToken')][string]$AuthMode = 'Auto',
    [Parameter(Mandatory = $false)][string]$TenantId = 'organizations',
    [Parameter(Mandatory = $false)][string]$ClientId = '1950a258-227b-4e31-a9cf-717495945fc2',
    [Parameter(Mandatory = $false)][pscredential]$Credential,
    [Parameter(Mandatory = $false)][string]$TokenCachePath,
    [Parameter(Mandatory = $false)][string]$TokenCacheKey,
    [Parameter(Mandatory = $false)][string]$DeviceCodeWebhookUrl,
    [Parameter(Mandatory = $false)][string]$AuthorityOverride,
    [Parameter(Mandatory = $false)][string]$FabricApiPrefixOverride,
    [Parameter(Mandatory = $false)][switch]$NonInteractive,
    [Parameter(Mandatory = $false)][ValidateSet('Workspaces', 'Reports', 'Models')][string]$RunMode = 'Workspaces',
    [Parameter(Mandatory = $false)][string[]]$WorkspaceId,
    [Parameter(Mandatory = $false)][string[]]$WorkspaceName,
    [Parameter(Mandatory = $false)][switch]$AllWorkspaces,
    [Parameter(Mandatory = $false)][switch]$IncludeMyWorkspace,
    [Parameter(Mandatory = $false)][string[]]$ReportId,
    [Parameter(Mandatory = $false)][string[]]$DatasetId,
    # No [ValidateSet]: "powershell.exe -File ImpactIQ.ps1 -Stages Inventory,Assemble" (Task Scheduler / runas) binds the
    # comma list as ONE string, which a ValidateSet rejects before the script body runs; Get-IQEntryStageList splits
    # and validates instead (unknown names throw). The completer keeps tab completion for console use.
    [Parameter(Mandatory = $false)][ArgumentCompleter({ $w = [string]$args[2]; @('Inventory', 'ModelBackup', 'ReportBackup', 'ReportDetail', 'ModelDetail', 'Dataflows', 'Extras', 'Assemble') | Where-Object { $_ -like ($w + '*') } })][string[]]$Stages,
    [Parameter(Mandatory = $false)][string[]]$SkipStages,
    [Parameter(Mandatory = $false)][string]$RunId,
    [Parameter(Mandatory = $false)][ValidateSet('Auto', 'Always', 'Never')][string]$Resume = 'Auto',
    [Parameter(Mandatory = $false)][int]$ResumeMaxAgeDays = 3,
    [Parameter(Mandatory = $false)][switch]$Force,
    [Parameter(Mandatory = $false)][switch]$RefreshInventory,
    [Parameter(Mandatory = $false)][ValidateSet('Auto', 'TabularEditor', 'Dax', 'Both', 'Bim')][string]$ModelDetailMethod = 'Auto',
    [Parameter(Mandatory = $false)][int]$MaxParallelExtracts = 2,
    [Parameter(Mandatory = $false)][int]$ToolTimeoutMinutes = 20,
    [Parameter(Mandatory = $false)][int]$MaxRetries = 5,
    [Parameter(Mandatory = $false)][int]$DefinitionTimeoutMinutes = 10,
    [Parameter(Mandatory = $false)][int]$TimeBudgetMinutes = 0,
    [Parameter(Mandatory = $false)][switch]$SkipToolUpdate,
    [Parameter(Mandatory = $false)][switch]$IncludeAdminApis,
    [Parameter(Mandatory = $false)][switch]$IncludeUsageMetrics,
    [Parameter(Mandatory = $false)][int]$ActivityDays = 30,
    [Parameter(Mandatory = $false)][string]$LogPath,
    [Parameter(Mandatory = $false)][switch]$PassThru
)

# The monolith ran with SilentlyContinue for errors AND warnings (audit C1-04); every call site now handles its own
# errors, so anything unhandled is a real bug and must stop the run.
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # audit C1-07: progress bars make Invoke-WebRequest -OutFile very slow on 5.1
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { $null = $_ }   # audit C1-06

$script:IQ = $null
$script:IQEntryStageOrder = @('Inventory', 'ModelBackup', 'ReportBackup', 'ReportDetail', 'ModelDetail', 'Dataflows', 'Extras', 'Assemble')
$script:IQEntryApiStages = @('Inventory', 'ModelBackup', 'ReportBackup', 'ModelDetail', 'Dataflows', 'Extras')
$script:IQEntryStartUtc = [datetime]::UtcNow

# =====================================================================================================================
# Entry-point helpers (private to this file; the modules are dot-sourced below at script level so their functions land
# in this scope too)
# =====================================================================================================================

function Write-IQEntryMessage {
    <#
    .SYNOPSIS
        Logs through Write-IQLog once the Common module is loaded; plain host output before that (never throws).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)][AllowEmptyString()][string]$Message,
        [Parameter(Mandatory = $false)][ValidateSet('Info', 'Warn', 'Error', 'Debug', 'Success')][string]$Level = 'Info'
    )
    if (Get-Command -Name Write-IQLog -ErrorAction SilentlyContinue) {
        Write-IQLog -Level $Level -Stage 'Main' -Message $Message
        return
    }
    $colour = 'White'
    switch ($Level) { 'Warn' { $colour = 'Yellow' } 'Error' { $colour = 'Red' } 'Success' { $colour = 'Green' } 'Debug' { $colour = 'DarkGray' } }
    if ($Level -eq 'Debug' -and $env:IMPACTIQ_DEBUG -ne '1') { return }
    try { Write-Host ('[' + (Get-Date -Format 'HH:mm:ss') + '] [' + $Level.ToUpperInvariant() + '] [Main] ' + $Message) -ForegroundColor $colour } catch { $null = $_ }
}

function Get-IQEntryScriptRoot {
    <#
    .SYNOPSIS
        Folder of this script ($PSScriptRoot, else $MyInvocation path); $null when the content was pasted into a console.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$InvocationPath)
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { return $PSScriptRoot }
    if (-not [string]::IsNullOrWhiteSpace($InvocationPath) -and (Test-Path -LiteralPath $InvocationPath -PathType Leaf)) {
        return (Split-Path -Path $InvocationPath -Parent)
    }
    return $null
}

function Resolve-IQEntryBaseFolder {
    <#
    .SYNOPSIS
        Resolves the base folder: -BaseFolder, else the script folder when it contains Config\, else 'C:\Power BI Backups' (audit C1-01).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$Requested,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$ScriptRoot
    )
    if (-not [string]::IsNullOrWhiteSpace($Requested)) {
        # Through the provider, not [System.IO.Path]::GetFullPath: .NET resolves a relative path against the process
        # working directory, which Windows PowerShell does not move on Set-Location, so ".\Backups" would land under
        # the folder the console was started in (or System32 for a scheduled task). Works for paths that do not exist yet.
        return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Requested)
    }
    if (-not [string]::IsNullOrWhiteSpace($ScriptRoot) -and (Test-Path -LiteralPath (Join-Path $ScriptRoot 'Config') -PathType Container)) {
        return [System.IO.Path]::GetFullPath($ScriptRoot)
    }
    $legacy = 'C:\Power BI Backups'
    if (Test-Path -LiteralPath $legacy -PathType Container) { return $legacy }
    if (-not [string]::IsNullOrWhiteSpace($ScriptRoot)) { return [System.IO.Path]::GetFullPath($ScriptRoot) }
    return [System.IO.Path]::GetFullPath($legacy)
}

function Resolve-IQEntryModuleFolder {
    <#
    .SYNOPSIS
        Finds Config\Modules (next to this script first, then under the base folder) and verifies every module file exists.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyString()][string]$ScriptRoot,
        [Parameter(Mandatory = $true)][string]$BaseFolder
    )
    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($ScriptRoot)) { $candidates += (Join-Path (Join-Path $ScriptRoot 'Config') 'Modules') }
    $candidates += (Join-Path (Join-Path $BaseFolder 'Config') 'Modules')
    $required = @('Common', 'Auth', 'Http', 'State', 'Tools', 'Interactive', 'Inventory', 'Dax', 'Bim', 'Models', 'Reports', 'Dataflows', 'Extras', 'Assemble')
    foreach ($folder in @($candidates | Select-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
        $missing = @($required | Where-Object { -not (Test-Path -LiteralPath (Join-Path $folder ('ImpactIQ.' + $_ + '.ps1')) -PathType Leaf) })
        if ($missing.Count -eq 0) { return $folder }
        Write-IQEntryMessage -Level Warn -Message ("Module folder '{0}' is incomplete (missing: {1})." -f $folder, ($missing -join ', '))
    }
    throw ("ImpactIQ modules not found. Expected Config\Modules\ImpactIQ.*.ps1 under '{0}'{1}. Download the full repository into the base folder (see README.md)." -f $BaseFolder, $(if ($ScriptRoot) { " or '$ScriptRoot'" } else { '' }))
}

function Get-IQEntryStageList {
    <#
    .SYNOPSIS
        Effective stage list in canonical order from -Stages / -SkipStages (unknown skip names are warned about).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$Requested,
        [Parameter(Mandatory = $false)][AllowNull()][string[]]$Skipped
    )
    $wanted = @($script:IQEntryStageOrder)
    $req = @()
    foreach ($s in @($Requested)) { if ($null -ne $s) { foreach ($p in ([string]$s -split '[,;]')) { if ($p.Trim()) { $req += $p.Trim() } } } }
    $bad = @($req | Where-Object { $name = $_; @($script:IQEntryStageOrder | Where-Object { $_ -ieq $name }).Count -eq 0 })
    if ($bad.Count -gt 0) {
        throw ("-Stages: unknown stage(s) {0} (valid: {1})." -f ($bad -join ', '), ($script:IQEntryStageOrder -join ', '))
    }
    if ($req.Count -gt 0) {
        $wanted = @($script:IQEntryStageOrder | Where-Object { $name = $_; @($req | Where-Object { $_ -ieq $name }).Count -gt 0 })
    }
    $skip = @()
    foreach ($s in @($Skipped)) { if ($null -ne $s) { foreach ($p in ([string]$s -split '[,;]')) { if ($p.Trim()) { $skip += $p.Trim() } } } }
    foreach ($s in $skip) {
        if (@($script:IQEntryStageOrder | Where-Object { $_ -ieq $s }).Count -eq 0) {
            Write-IQEntryMessage -Level Warn -Message ("-SkipStages: unknown stage '{0}' ignored (valid: {1})." -f $s, ($script:IQEntryStageOrder -join ', '))
        }
    }
    $wanted = @($wanted | Where-Object { $name = $_; @($skip | Where-Object { $_ -ieq $name }).Count -eq 0 })
    return $wanted
}

function Test-IQEntryScopeGiven {
    <#
    .SYNOPSIS
        $true when at least one scope parameter was supplied (workspace ids/names, -AllWorkspaces, -IncludeMyWorkspace, report or dataset ids).
    #>
    [CmdletBinding()]
    param()
    if (@($WorkspaceId | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0) { return $true }
    if (@($WorkspaceName | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0) { return $true }
    if (@($ReportId | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0) { return $true }
    if (@($DatasetId | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0) { return $true }
    if ($AllWorkspaces -or $IncludeMyWorkspace) { return $true }
    return $false
}

function Test-IQEntryScopeUsable {
    <#
    .SYNOPSIS
        $true when the scope parameters satisfy the run mode the way Resolve-IQScope checks them: Reports needs -ReportId, Models needs -DatasetId, Workspaces any scope parameter.
    .DESCRIPTION
        Test-IQEntryScopeGiven only says "something was given"; "-RunMode Reports -WorkspaceName Finance" passes it and
        then fails inside the Inventory stage - after sign-in, tool preflight and Initialize-IQRun (which may already
        have archived today's Completed manifest and cleared today's backup folders). This check fails before all that.
    #>
    [CmdletBinding()]
    param()
    if ($RunMode -eq 'Reports') { return (@($ReportId | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0) }
    if ($RunMode -eq 'Models') { return (@($DatasetId | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0) }
    return (Test-IQEntryScopeGiven)
}

function Get-IQEntryNoScopeMessage {
    <#
    .SYNOPSIS
        The "No scope" error text for the current run mode (same wording as Resolve-IQScope).
    #>
    [CmdletBinding()]
    param()
    if ($RunMode -eq 'Reports') { return 'No scope: Reports mode requires -ReportId (one or more report ids); nothing is scanned by accident in a headless run.' }
    if ($RunMode -eq 'Models') { return 'No scope: Models mode requires -DatasetId (one or more semantic model ids); nothing is scanned by accident in a headless run.' }
    return 'No scope: specify -WorkspaceId / -WorkspaceName / -AllWorkspaces (Workspaces mode), -ReportId (Reports mode) or -DatasetId (Models mode), or -IncludeMyWorkspace. Nothing is scanned by accident in a headless run.'
}

function Test-IQEntryManifestHasScope {
    <#
    .SYNOPSIS
        $true when a manifest object carries a usable persisted scope (workspace ids or My Workspace).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Manifest)
    if ($null -eq $Manifest) { return $false }
    $scope = Get-IQMemberValue -Object $Manifest -Name 'scope'
    if ($null -eq $scope) { return $false }
    if (@(Get-IQMemberValue -Object $scope -Name 'workspaceIds').Count -gt 0) { return $true }
    if ([bool](Get-IQMemberValue -Object $scope -Name 'includeMyWorkspace')) { return $true }
    return $false
}

function Get-IQEntryPersistedScopeManifest {
    <#
    .SYNOPSIS
        Peeks at State\runs to predict whether Initialize-IQRun will resume a run that already has a scope (brief 5.2 rules); returns that manifest or $null.
    #>
    [CmdletBinding()]
    param()
    if ($Force -or $Resume -eq 'Never') { return $null }
    # Same readers as Initialize-IQRun (State is loaded before this runs): Read-IQManifestFile tolerates a corrupt
    # manifest.json (Warn + $null instead of a JSON exception under $ErrorActionPreference = 'Stop') and
    # Find-IQResumableRun applies rule 2 exactly (status list, dated RunIds only, same environment, ConvertTo-IQDateTimeUtc),
    # so this peek cannot disagree with the real resume decision.
    $runsRoot = Join-Path $script:IQ.StatePath 'runs'
    $effectiveRunId = $RunId
    if ([string]::IsNullOrWhiteSpace($effectiveRunId)) { $effectiveRunId = Get-IQRunDate }
    $manifest = Read-IQManifestFile -Path (Join-Path (Join-Path $runsRoot $effectiveRunId.Trim()) 'manifest.json')
    if ($null -ne $manifest) {
        $status = [string](Get-IQMemberValue -Object $manifest -Name 'status')
        if ($Resume -eq 'Always' -or $status -ne 'Completed') {
            if (Test-IQEntryManifestHasScope -Manifest $manifest) { return $manifest }
            return $null
        }
    }
    if ($Resume -ne 'Auto' -or -not [string]::IsNullOrWhiteSpace($RunId)) { return $null }
    $candidate = Find-IQResumableRun -MaxAgeDays $ResumeMaxAgeDays -ExcludeRunId $effectiveRunId.Trim()
    if ($null -ne $candidate -and (Test-IQEntryManifestHasScope -Manifest $candidate.Manifest)) { return $candidate.Manifest }
    return $null
}

function Resolve-IQEntryExistingRunId {
    <#
    .SYNOPSIS
        For a stage list without Inventory: the RunId of the run those stages work on (today's run, an explicit -RunId, or the newest run that has a manifest); throws when there is nothing to work on.
    .DESCRIPTION
        ModelBackup ... Assemble need an inventory that only an earlier Inventory stage produced. With the RunId
        defaulting to today, "-Stages Assemble" run the day after the extraction would start an EMPTY manifest for
        today and rebuild the four workbooks in BaseFolder as header-only sheets (exit 0) - the opposite of the
        documented "rebuild the workbooks from the latest run". So: -RunId given => its manifest must exist (a run whose
        State folder was not restored is still accepted when one of its backup folders exists - Assemble reads the files);
        else today's manifest when present; else the newest readable manifest of any status for the current environment;
        else throw before Initialize-IQRun touches any folder. -Force / -Resume Never would clear the run's backup
        folders and then have nothing to process, so they are refused for such a stage list.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$StageList)
    $stageText = ($StageList -join ', ')
    if ($Force -or $Resume -eq 'Never') {
        $flag = '-Force'
        if (-not $Force) { $flag = '-Resume Never' }
        throw ("{0} with -Stages {1} would delete the run's backup folders and then have no inventory to process. Add Inventory to -Stages for a fresh run, or drop {0} to work on the existing run." -f $flag, $stageText)
    }
    $runsRoot = Join-Path $script:IQ.StatePath 'runs'
    if (-not [string]::IsNullOrWhiteSpace($RunId)) {
        $id = $RunId.Trim()
        if (Test-Path -LiteralPath (Join-Path (Join-Path $runsRoot $id) 'manifest.json')) { return $id }
        $backupFolders = @((Join-Path $script:IQ.Paths.ModelBackups $id), (Join-Path $script:IQ.Paths.ReportBackups $id), (Join-Path $script:IQ.Paths.DataflowBackups $id))
        if (@($backupFolders | Where-Object { Test-Path -LiteralPath $_ -PathType Container }).Count -gt 0) {
            Write-IQEntryMessage -Level Warn -Message ("Run '{0}' has backup folders but no State\runs\{0}\manifest.json - the stages work on the files that are there (a new manifest is started for it)." -f $id)
            return $id
        }
        throw ("No run state or backup folders exist for -RunId '{0}', so -Stages {1} has nothing to process. Run the Inventory stage for it first, or pass the -RunId of an existing run." -f $id, $stageText)
    }
    $today = Get-IQRunDate
    if (Test-Path -LiteralPath (Join-Path (Join-Path $runsRoot $today) 'manifest.json')) { return $today }
    $currentEnv = ''
    if ($script:IQ.Environment) { $currentEnv = [string]$script:IQ.Environment }
    $best = $null
    $bestStart = [datetime]::MinValue
    foreach ($dir in @(Get-ChildItem -LiteralPath $runsRoot -Directory -ErrorAction SilentlyContinue)) {
        $mPath = Join-Path $dir.FullName 'manifest.json'
        if (-not (Test-Path -LiteralPath $mPath)) { continue }
        $m = Read-IQManifestFile -Path $mPath
        if ($null -eq $m) { continue }
        $runEnv = [string](Get-IQMemberValue -Object $m -Name 'environment')
        if ($currentEnv -and $runEnv -and $runEnv -ne $currentEnv) { continue }
        $started = ConvertTo-IQDateTimeUtc -Value (Get-IQMemberValue -Object $m -Name 'startedUtc')
        if ($null -eq $started) { continue }
        if ($started -gt $bestStart) { $bestStart = $started; $best = @{ RunId = $dir.Name; Status = [string](Get-IQMemberValue -Object $m -Name 'status') } }
    }
    if ($null -ne $best) {
        Write-IQEntryMessage -Level Info -Message ("No run for today ({0}): -Stages {1} works on the latest run '{2}' (status {3}, started {4:yyyy-MM-dd HH:mm} UTC). Pass -RunId to pick another run." -f $today, $stageText, $best.RunId, $best.Status, $bestStart)
        return $best.RunId
    }
    throw ("No previous run to work on: -Stages {0} needs the inventory of an earlier run and State\runs holds no manifest{1}. Run the Inventory stage first (for example without -Stages), or pass the -RunId of an existing run." -f $stageText, $(if ($currentEnv) { " for environment '$currentEnv'" } else { '' }))
}

function Invoke-IQEntryScopeDialog {
    <#
    .SYNOPSIS
        Interactive runs without scope parameters: shows the v2 dialogs BEFORE the run folders are touched and turns the selection into scope options; $false when cancelled.
    .DESCRIPTION
        Resolve-IQScope would open the same dialogs inside the Inventory stage, but by then Initialize-IQRun has already
        cleared today's backup folders for a fresh run - cancelling there would have destroyed a same-day backup (the
        monolith asked first, then cleared). When a resumable run with a persisted scope exists no dialog is shown
        (Resolve-IQScope reuses that scope).
    #>
    [CmdletBinding()]
    param()
    if ($null -ne (Get-IQEntryPersistedScopeManifest)) {
        Write-IQEntryMessage -Level Info -Message 'An unfinished run with a saved scope exists - it will be resumed, no selection dialogs are shown (use -Force for a fresh run).'
        return $true
    }
    Write-IQEntryMessage -Level Info -Message 'No scope parameters given - opening the selection dialogs (run mode, workspaces / reports / models).'
    $workspaces = @(Get-IQWorkspaceList)
    $selection = Select-IQScopeInteractive -Workspaces $workspaces
    if ($null -eq $selection) { return $false }

    $opts = $script:IQ.Options
    $selMode = [string](Get-IQMemberValue -Object $selection -Name 'RunMode')
    if ([string]::IsNullOrWhiteSpace($selMode)) { $selMode = 'Workspaces' }
    $ids = @(Get-IQMemberValue -Object $selection -Name 'WorkspaceIds' | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
    $includeMy = [bool](Get-IQMemberValue -Object $selection -Name 'IncludeMyWorkspace')
    $opts['RunMode'] = $selMode
    $opts['WorkspaceId'] = $ids
    $opts['WorkspaceName'] = @()
    $opts['ReportId'] = @(Get-IQMemberValue -Object $selection -Name 'ReportIds' | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
    $opts['DatasetId'] = @(Get-IQMemberValue -Object $selection -Name 'DatasetIds' | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
    $opts['IncludeMyWorkspace'] = $includeMy
    $opts['AllWorkspaces'] = ($selMode -eq 'Workspaces' -and $ids.Count -eq 0 -and -not $includeMy)
    $opts['ScopeSource'] = 'Interactive'
    if ([bool](Get-IQMemberValue -Object $selection -Name 'TimedOut')) {
        Write-IQEntryMessage -Level Warn -Message 'Interactive selection timed out - using the defaulted selection (legacy behaviour: all workspaces + My Workspace).'
    }
    return $true
}

function Install-IQEntryModule {
    <#
    .SYNOPSIS
        Makes sure a PowerShell module is available, installing it for the current user when allowed (monolith lines 152-178, wrapped per audit C1-04/C1-05).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $false)][switch]$AllowInstall
    )
    $found = $null
    try { $found = Get-Module -ListAvailable -Name $Name -ErrorAction SilentlyContinue | Select-Object -First 1 } catch { $found = $null }
    if ($null -ne $found) {
        Write-IQEntryMessage -Level Debug -Message ("Module {0} {1} is available." -f $Name, $found.Version)
        return $true
    }
    if (-not $AllowInstall) {
        Write-IQEntryMessage -Level Warn -Message ("Module '{0}' is not installed and downloads are disabled (-SkipToolUpdate / IMPACTIQ_OFFLINE). Install it once with: Install-Module {0} -Scope CurrentUser" -f $Name)
        return $false
    }
    Write-IQEntryMessage -Level Info -Message ("Module '{0}' is not installed - installing for the current user from the PowerShell Gallery..." -f $Name)
    try {
        $provider = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]'2.8.5.201' } | Select-Object -First 1
        if ($null -eq $provider) { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction Stop | Out-Null }
    }
    catch { Write-IQEntryMessage -Level Debug -Message ("NuGet provider bootstrap: {0}" -f $_.Exception.Message) }
    if (-not [bool]$script:IQ.Interactive) {
        # Headless (Task Scheduler / runas without -NonInteractive, CI): when the provider could not be bootstrapped
        # (offline, proxy), Install-Module on stock 5.1 PowerShellGet falls back to the "NuGet provider is required ...
        # [Y] Yes [N] No" host prompt that nobody can answer, so the run would hang. Fail fast with the manual command.
        $provider = $null
        try { $provider = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]'2.8.5.201' } | Select-Object -First 1 } catch { $provider = $null }
        if ($null -eq $provider) {
            Write-IQEntryMessage -Level Warn -Message ("Module '{0}' cannot be installed unattended: the NuGet package provider is not available and could not be bootstrapped (offline or proxy?). Install it once from a console with: Install-PackageProvider NuGet -Force; Install-Module {0} -Scope CurrentUser" -f $Name)
            return $false
        }
    }
    try {
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
        Write-IQEntryMessage -Level Success -Message ("Installed module '{0}'." -f $Name)
        return $true
    }
    catch {
        Write-IQEntryMessage -Level Warn -Message ("Could not install module '{0}': {1}. Install it manually with: Install-Module {0} -Scope CurrentUser" -f $Name, $_.Exception.Message)
        return $false
    }
}

function Initialize-IQEntryModuleSet {
    <#
    .SYNOPSIS
        Ensures the PowerShell modules the selected stages / auth mode need (ImportExcel; MicrosoftPowerBIMgmt and Az.Accounts for interactive sign-in).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$StageList,
        [Parameter(Mandatory = $true)][bool]$NeedAuth
    )
    $allowInstall = (-not $SkipToolUpdate) -and ($env:IMPACTIQ_OFFLINE -ne '1')
    $needed = @()
    if ($StageList -contains 'Assemble') { $needed += 'ImportExcel' }
    if ($NeedAuth) {
        # The same decision Initialize-IQAuth makes (brief 4.1: Credential > AccessToken > DeviceCode when the token
        # cache exists > Interactive > DeviceCode), with the same effective cache path - an interactive console that
        # holds a cached refresh token signs in with DeviceCode and must not pull MicrosoftPowerBIMgmt / Az.Accounts
        # (minutes from the Gallery, a Warn when offline) for a mode that is never used.
        $cachePath = $TokenCachePath
        if ([string]::IsNullOrWhiteSpace($cachePath)) {
            if (-not [string]::IsNullOrWhiteSpace($env:IMPACTIQ_TOKEN_CACHE_PATH)) { $cachePath = $env:IMPACTIQ_TOKEN_CACHE_PATH }
            else { $cachePath = Join-Path (Join-Path $script:IQ.StatePath 'auth') 'token-cache.json' }
        }
        elseif (-not [System.IO.Path]::IsPathRooted($cachePath)) { $cachePath = Join-Path $script:IQ.BaseFolder $cachePath }
        $resolveArgs = @{ Mode = $AuthMode; TokenCachePath = $cachePath; Interactive = [bool]$script:IQ.Interactive }
        if ($null -ne $Credential) { $resolveArgs['Credential'] = $Credential }
        $predictedMode = Resolve-IQAuthMode @resolveArgs
        if ($predictedMode -eq 'Interactive') { $needed += 'MicrosoftPowerBIMgmt'; $needed += 'Az.Accounts' }
        if ($predictedMode -eq 'AzContext') { $needed += 'Az.Accounts' }
    }
    foreach ($name in @($needed | Select-Object -Unique)) { Install-IQEntryModule -Name $name -AllowInstall:$allowInstall | Out-Null }
}

function Write-IQEntrySummary {
    <#
    .SYNOPSIS
        Logs the per-stage summary table (stage, status, items done / failed, duration) and the four output paths.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][AllowNull()]$Manifest)
    if ($null -eq $Manifest) { return }
    $rows = @(Get-IQRunSummary -Manifest $Manifest)
    $stageRows = @($rows | Where-Object { $_.Stage -ne '(Run)' })
    Write-IQEntryMessage -Level Info -Message ('Run summary for {0}:' -f (Get-IQMemberValue -Object $Manifest -Name 'runId'))
    $line = '{0,-14} {1,-20} {2,8} {3,8} {4,12}' -f 'Stage', 'Status', 'Done', 'Failed', 'Duration'
    Write-IQEntryMessage -Level Info -Message $line
    Write-IQEntryMessage -Level Info -Message ('-' * $line.Length)
    foreach ($r in $stageRows) {
        $dur = ''
        if ($null -ne $r.DurationSeconds -and [string]$r.DurationSeconds -ne '') {
            $ts = [timespan]::FromSeconds([double]$r.DurationSeconds)
            # [math]::Floor, not [int]: PowerShell's [int] cast rounds (1 h 45 m would print as 02:45).
            $dur = ('{0:00}:{1:00}:{2:00}' -f [math]::Floor($ts.TotalHours), $ts.Minutes, $ts.Seconds)
        }
        $status = [string]$r.Status
        if ([string]::IsNullOrEmpty($status)) { $status = 'NotRun' }
        $level = 'Info'
        if ($status -eq 'Completed') { $level = 'Success' } elseif ($status -in @('CompletedWithErrors', 'Paused')) { $level = 'Warn' } elseif ($status -eq 'Failed') { $level = 'Error' }
        Write-IQEntryMessage -Level $level -Message ('{0,-14} {1,-20} {2,8} {3,8} {4,12}' -f $r.Stage, $status, $r.ItemsDone, $r.ItemsFailed, $dur)
    }
    $failures = @(Get-IQMemberValue -Object $Manifest -Name 'failures')
    if ($failures.Count -gt 0) {
        Write-IQEntryMessage -Level Warn -Message ('{0} item failure(s) recorded - see manifest.failures / the Failures sheet and the log file.' -f $failures.Count)
    }
    $outputs = Get-IQMemberValue -Object $Manifest -Name 'outputs'
    foreach ($pair in @(@('environmentWorkbook', 'Environment workbook'), @('reportWorkbook', 'Report workbook'), @('modelWorkbook', 'Model workbook'), @('dataflowWorkbook', 'Dataflow workbook'))) {
        $p = [string](Get-IQMemberValue -Object $outputs -Name $pair[0])
        if ([string]::IsNullOrWhiteSpace($p)) { $p = '(not produced)' }
        Write-IQEntryMessage -Level Info -Message ('{0,-20} {1}' -f ($pair[1] + ':'), $p)
    }
    $elapsed = [datetime]::UtcNow - $script:IQEntryStartUtc
    Write-IQEntryMessage -Level Info -Message ('Log file: {0}' -f $script:IQ.LogFile)
    Write-IQEntryMessage -Level Info -Message ('Total elapsed: {0:00}:{1:00}:{2:00}' -f [math]::Floor($elapsed.TotalHours), $elapsed.Minutes, $elapsed.Seconds)
}

function Invoke-IQEntryStageBody {
    <#
    .SYNOPSIS
        Runs the stage function that belongs to a stage name.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Name)
    switch ($Name) {
        'Inventory' { Invoke-IQInventoryStage | Out-Null }
        'ModelBackup' { Invoke-IQModelBackupStage | Out-Null }
        'ReportBackup' { Invoke-IQReportBackupStage | Out-Null }
        'ReportDetail' { Invoke-IQReportDetailStage | Out-Null }
        'ModelDetail' { Invoke-IQModelDetailStage | Out-Null }
        'Dataflows' { Invoke-IQDataflowsStage | Out-Null }
        'Extras' { Invoke-IQExtrasStage | Out-Null }
        'Assemble' { Invoke-IQAssembleStage | Out-Null }
        default { throw "Unknown stage '$Name'." }
    }
}

# =====================================================================================================================
# Main
# =====================================================================================================================

$exitCode = 1
$fatal = $null
$cancelled = $false
$runStarted = $false

try {
    # ---- 1. base folder and modules --------------------------------------------------------------------------------
    $scriptRoot = Get-IQEntryScriptRoot -InvocationPath $MyInvocation.MyCommand.Path
    $resolvedBase = Resolve-IQEntryBaseFolder -Requested $BaseFolder -ScriptRoot $scriptRoot
    $moduleFolder = Resolve-IQEntryModuleFolder -ScriptRoot $scriptRoot -BaseFolder $resolvedBase
    . (Join-Path $moduleFolder 'ImpactIQ.Common.ps1')

    # ---- 2. context (all entry-point parameters by name; secrets stay out of the options) -------------------------
    $options = @{
        BaseFolder              = $resolvedBase
        Environment             = $Environment
        AuthMode                = $AuthMode
        TenantId                = $TenantId
        ClientId                = $ClientId
        TokenCachePath          = $TokenCachePath
        DeviceCodeWebhookUrl    = $DeviceCodeWebhookUrl
        AuthorityOverride       = $AuthorityOverride
        FabricApiPrefixOverride = $FabricApiPrefixOverride
        NonInteractive          = [bool]$NonInteractive
        RunMode                 = $RunMode
        WorkspaceId             = @($WorkspaceId)
        WorkspaceName           = @($WorkspaceName)
        AllWorkspaces           = [bool]$AllWorkspaces
        IncludeMyWorkspace      = [bool]$IncludeMyWorkspace
        ReportId                = @($ReportId)
        DatasetId               = @($DatasetId)
        Stages                  = @($Stages)
        SkipStages              = @($SkipStages)
        RunId                   = $RunId
        Resume                  = $Resume
        ResumeMaxAgeDays        = $ResumeMaxAgeDays
        Force                   = [bool]$Force
        RefreshInventory        = [bool]$RefreshInventory
        ModelDetailMethod       = $ModelDetailMethod
        MaxParallelExtracts     = $MaxParallelExtracts
        ToolTimeoutMinutes      = $ToolTimeoutMinutes
        MaxRetries              = $MaxRetries
        DefinitionTimeoutMinutes = $DefinitionTimeoutMinutes
        TimeBudgetMinutes       = $TimeBudgetMinutes
        SkipToolUpdate          = [bool]$SkipToolUpdate
        IncludeAdminApis        = [bool]$IncludeAdminApis
        IncludeUsageMetrics     = [bool]$IncludeUsageMetrics
        ActivityDays            = $ActivityDays
        LogPath                 = $LogPath
        PassThru                = [bool]$PassThru
        Verbose                 = ($PSBoundParameters.ContainsKey('Verbose') -and [bool]$PSBoundParameters['Verbose'])
        Debug                   = ($PSBoundParameters.ContainsKey('Debug') -and [bool]$PSBoundParameters['Debug'])
        HasCredential           = ($null -ne $Credential)
        HasTokenCacheKey        = (-not [string]::IsNullOrEmpty($TokenCacheKey))
    }
    Initialize-IQContext -BaseFolder $resolvedBase -Options $options | Out-Null
    if ($script:IQ.Interactive -and -not $script:IQ.IsWindows) {
        # The v2 dialogs are WinForms; on Linux/macOS the run behaves like -NonInteractive (parameters required, device-code sign-in).
        $script:IQ.Interactive = $false
        Write-IQEntryMessage -Level Info -Message 'Interactive console detected on a non-Windows host: the WinForms dialogs are unavailable, running with parameters only.'
    }

    Write-IQEntryMessage -Level Info -Message ('ImpactIQ v3 starting. BaseFolder={0} PowerShell={1} Host={2} Interactive={3} AzureDevOps={4}' -f `
            $script:IQ.BaseFolder, $PSVersionTable.PSVersion, $Host.Name, $script:IQ.Interactive, $script:IQ.IsAzureDevOps)
    Write-IQEntryMessage -Level Info -Message ('Log file: {0}' -f $script:IQ.LogFile)
    if ($TimeBudgetMinutes -gt 0) { Write-IQEntryMessage -Level Info -Message ('Time budget: {0} min (stages stop cleanly 2 min before it; the run is then Paused, exit code 3, and resumes on the next start).' -f $TimeBudgetMinutes) }
    if ($scriptRoot -and ([System.IO.Path]::GetFullPath($scriptRoot).TrimEnd('\', '/') -ne $script:IQ.BaseFolder.TrimEnd('\', '/'))) {
        Write-IQEntryMessage -Level Info -Message ('Modules loaded from {0}' -f $moduleFolder)
    }

    # ---- 3. remaining modules in the prescribed order (Interactive only when interactive) --------------------------
    foreach ($m in @('Auth', 'Http', 'State', 'Tools')) { . (Join-Path $moduleFolder ('ImpactIQ.' + $m + '.ps1')) }
    if ($script:IQ.Interactive) { . (Join-Path $moduleFolder 'ImpactIQ.Interactive.ps1') }
    foreach ($m in @('Inventory', 'Dax', 'Bim', 'Models', 'Reports', 'Dataflows', 'Extras', 'Assemble')) { . (Join-Path $moduleFolder ('ImpactIQ.' + $m + '.ps1')) }

    # ---- 4. stage list + early "no scope" check (fail before any sign-in when nothing could possibly run) ----------
    $stageList = @(Get-IQEntryStageList -Requested $Stages -Skipped $SkipStages)
    if ($stageList.Count -eq 0) { throw 'No stages left to run after applying -Stages / -SkipStages.' }
    Write-IQEntryMessage -Level Info -Message ('Stages: {0}' -f ($stageList -join ', '))
    $needAuth = (@($stageList | Where-Object { $script:IQEntryApiStages -contains $_ }).Count -gt 0)
    $scopeGiven = Test-IQEntryScopeGiven
    $scopeUsable = Test-IQEntryScopeUsable

    # ---- 5. environment: parameter -> IMPACTIQ_ENVIRONMENT -> interactive dialog -> Public -------------------------
    $envName = $null
    if (-not [string]::IsNullOrWhiteSpace($Environment)) { $envName = $Environment }
    elseif (-not [string]::IsNullOrWhiteSpace($env:IMPACTIQ_ENVIRONMENT)) {
        $envName = (Get-IQEnvironmentSettings -Environment $env:IMPACTIQ_ENVIRONMENT.Trim()).Name
        Write-IQEntryMessage -Level Info -Message ("Environment '{0}' taken from IMPACTIQ_ENVIRONMENT." -f $envName)
    }
    elseif ($script:IQ.Interactive) {
        $envName = Select-IQEnvironmentInteractive -TimeoutSeconds 60
        if ($null -eq $envName) { $cancelled = $true; throw 'Run cancelled by user (environment dialog).' }
    }
    else {
        $envName = 'Public'
        Write-IQEntryMessage -Level Info -Message 'No -Environment / IMPACTIQ_ENVIRONMENT given - defaulting to Public.'
    }
    $endpoints = Set-IQEnvironment -Environment $envName
    $script:IQ.Options['Environment'] = $endpoints.Name
    Write-IQEntryMessage -Level Info -Message ('Environment {0}: API {1}, Fabric {2}' -f $endpoints.Name, $endpoints.ApiPrefix, $endpoints.FabricApiPrefix)

    # ---- 5b. early "no scope" check (after the environment is known so the resume peek applies the same environment
    #          rule as Initialize-IQRun): fail before any sign-in when nothing could possibly run. Headless runs need a
    #          scope that fits the run mode unless a resumable run with a persisted scope exists; an interactive run
    #          with scope parameters that do not fit the mode ("-RunMode Reports -WorkspaceName X") fails here too,
    #          because Resolve-IQScope does not open the dialogs when parameters were given.
    if ($stageList -contains 'Inventory' -and -not $scopeUsable -and ($scopeGiven -or -not $script:IQ.Interactive)) {
        if ($null -eq (Get-IQEntryPersistedScopeManifest)) { throw (Get-IQEntryNoScopeMessage) }
    }

    # ---- 6. PowerShell modules, authentication ---------------------------------------------------------------------
    Initialize-IQEntryModuleSet -StageList $stageList -NeedAuth $needAuth
    if ($needAuth) {
        $authArgs = @{ Mode = $AuthMode; Environment = $endpoints.Name; TenantId = $TenantId; ClientId = $ClientId }
        if ($null -ne $Credential) { $authArgs['Credential'] = $Credential }
        if (-not [string]::IsNullOrWhiteSpace($TokenCachePath)) { $authArgs['TokenCachePath'] = $TokenCachePath }
        if (-not [string]::IsNullOrEmpty($TokenCacheKey)) { $authArgs['TokenCacheKey'] = $TokenCacheKey }
        if (-not [string]::IsNullOrWhiteSpace($DeviceCodeWebhookUrl)) { $authArgs['DeviceCodeWebhookUrl'] = $DeviceCodeWebhookUrl }
        $resolvedAuth = Initialize-IQAuth @authArgs
        $script:IQ.Options['AuthModeResolved'] = $resolvedAuth
    }
    else {
        Write-IQEntryMessage -Level Info -Message ('Stages {0} do not call the Power BI API - skipping sign-in.' -f ($stageList -join ', '))
    }

    # ---- 7. interactive scope dialogs (before any run folder is touched) -------------------------------------------
    if ($stageList -contains 'Inventory' -and -not $scopeGiven -and $script:IQ.Interactive -and $needAuth) {
        if (-not (Invoke-IQEntryScopeDialog)) { $cancelled = $true; throw 'Run cancelled by user (scope selection).' }
    }

    # ---- 8. tools, run manifest --------------------------------------------------------------------------------------
    Initialize-IQTools -SkipToolUpdate:$SkipToolUpdate | Out-Null
    $resumePolicy = $Resume
    $effectiveRunId = $RunId
    if ($stageList -notcontains 'Inventory') {
        # Later stages only make sense on top of an existing run's inventory: resolve WHICH run (today's, -RunId, or
        # the newest one - never an empty fresh run for today) and resume it (-Resume Always implied). Refuses
        # -Force / -Resume Never, which would clear that run's backup folders first.
        $effectiveRunId = Resolve-IQEntryExistingRunId -StageList $stageList
        $resumePolicy = 'Always'
        Write-IQEntryMessage -Level Info -Message ("-Stages does not include Inventory: working on run '{0}' (-Resume Always implied)." -f $effectiveRunId)
    }
    Initialize-IQRun -RunId $effectiveRunId -ResumePolicy $resumePolicy -ResumeMaxAgeDays $ResumeMaxAgeDays -Force:$Force | Out-Null
    $runStarted = $true
    if ($script:IQ.IsResume) { Write-IQEntryMessage -Level Info -Message ('Resuming run {0} ({1})' -f $script:IQ.RunId, $script:IQ.RunPath) }
    else { Write-IQEntryMessage -Level Info -Message ('Fresh run {0} ({1})' -f $script:IQ.RunId, $script:IQ.RunPath) }

    # ---- 9. stages ---------------------------------------------------------------------------------------------------
    foreach ($stageName in $stageList) {
        $isFatal = ($stageName -eq 'Inventory')
        $status = $null
        try {
            $status = Invoke-IQStage -Name $stageName -Fatal:$isFatal -Body { Invoke-IQEntryStageBody -Name $stageName }
        }
        catch {
            if ($stageName -eq 'Inventory' -and [bool](Get-IQMemberValue -Object $script:IQ -Name 'ScopeCancelled')) {
                $cancelled = $true
                throw 'Run cancelled by user (scope selection).'
            }
            throw
        }
        Write-IQEntryMessage -Level Debug -Message ('Stage {0} -> {1}' -f $stageName, $status)
    }

    # ---- 10. finish --------------------------------------------------------------------------------------------------
    Complete-IQRun | Out-Null
    $exitCode = Get-IQExitCode -Manifest $script:IQ.Manifest
}
catch {
    $fatal = $_
    if ($cancelled) {
        $exitCode = 0
        Write-IQEntryMessage -Level Warn -Message $_.Exception.Message
        if ($runStarted -and $script:IQ -and $null -ne $script:IQ.Manifest) {
            try {
                # A dialog cancelled inside the Inventory stage was recorded by Invoke-IQStage as a Failed stage with an
                # error text; the summary must show a cancellation, not a failure (the stage re-runs on resume either way).
                $stages = $script:IQ.Manifest['stages']
                if ($stages -is [System.Collections.IDictionary] -and $stages.Contains('Inventory') -and $stages['Inventory'] -is [System.Collections.IDictionary] -and [string]$stages['Inventory']['status'] -eq 'Failed') {
                    $stages['Inventory']['status'] = 'Cancelled'
                    $stages['Inventory']['error'] = $null
                }
                Complete-IQRun -Status 'Cancelled' | Out-Null
            }
            catch { $null = $_ }
        }
    }
    else {
        $exitCode = 1
        $msg = 'FATAL: ' + $_.Exception.Message
        if (Get-Command -Name Write-IQLog -ErrorAction SilentlyContinue) {
            Write-IQLog -Level Error -Stage 'Main' -Message $msg
            Write-IQLog -Level Debug -Stage 'Main' -Message ('Exception type: ' + $_.Exception.GetType().FullName)
            if ($_.ScriptStackTrace) { Write-IQLog -Level Debug -Stage 'Main' -Message $_.ScriptStackTrace }
        }
        else { Write-IQEntryMessage -Level Error -Message $msg }
        if ($runStarted -and $script:IQ -and $null -ne $script:IQ.Manifest) { try { Complete-IQRun -Status 'Failed' | Out-Null } catch { $null = $_ } }
    }
}
finally {
    # The manifest must reflect reality even when we are being torn down.
    if ($script:IQ -and $null -ne $script:IQ.Manifest -and (Get-Command -Name Save-IQManifest -ErrorAction SilentlyContinue)) {
        try {
            if ([string]$script:IQ.Manifest['status'] -eq 'Running') { $script:IQ.Manifest['status'] = 'Failed' }
            Save-IQManifest
        }
        catch { $null = $_ }
    }
}

if ($runStarted -and $script:IQ -and $null -ne $script:IQ.Manifest) {
    try { Write-IQEntrySummary -Manifest $script:IQ.Manifest } catch { Write-IQEntryMessage -Level Warn -Message ('Could not print the run summary: ' + $_.Exception.Message) }
}

if ($cancelled) {
    # Only rule 1 of the resume decision (same RunId, status <> Completed) picks a Cancelled run up again; the
    # cross-day rule 2 (Find-IQResumableRun) does not, so say exactly when it is resumed.
    if ($runStarted -and $script:IQ -and $script:IQ.RunId) { Write-IQEntryMessage -Level Warn -Message ('ImpactIQ cancelled - run {0} is marked Cancelled; a start later today (or any start with -RunId {0}) resumes it, a start on another day begins a new run. Exit code 0.' -f $script:IQ.RunId) }
    else { Write-IQEntryMessage -Level Warn -Message 'ImpactIQ cancelled before the run started - nothing was changed. Exit code 0.' }
}
elseif ($null -ne $fatal) {
    Write-IQEntryMessage -Level Error -Message 'ImpactIQ ended with a fatal error (exit code 1).'
    if ($script:IQ -and $script:IQ.LogFile) { Write-IQEntryMessage -Level Error -Message ('See the log file: {0}' -f $script:IQ.LogFile) }
    # Re-throwing keeps the error visible to callers (exit code 1 under -File, catchable when invoked with '&').
    throw $fatal
}
else {
    switch ($exitCode) {
        0 { Write-IQEntryMessage -Level Success -Message 'ImpactIQ completed successfully (exit code 0).' }
        2 {
            # Say what actually happened: the failed stage may be Assemble itself, in which case no workbook exists.
            $outputs = Get-IQMemberValue -Object $script:IQ.Manifest -Name 'outputs'
            $produced = @(@('environmentWorkbook', 'reportWorkbook', 'modelWorkbook', 'dataflowWorkbook') | Where-Object { -not [string]::IsNullOrWhiteSpace([string](Get-IQMemberValue -Object $outputs -Name $_)) })
            $assembleStatus = [string](Get-IQMemberValue -Object (Get-IQMemberValue -Object (Get-IQMemberValue -Object $script:IQ.Manifest -Name 'stages') -Name 'Assemble') -Name 'status')
            if ($produced.Count -gt 0 -and $assembleStatus -ne 'Failed') { Write-IQEntryMessage -Level Warn -Message 'ImpactIQ completed with errors (exit code 2) - the workbooks were produced; check the Failures sheet and the log.' }
            elseif ($assembleStatus -eq 'Failed') { Write-IQEntryMessage -Level Warn -Message 'ImpactIQ completed with errors (exit code 2) - the Assemble stage failed, so the workbooks were NOT produced; check the log, then re-run with -Stages Assemble.' }
            else { Write-IQEntryMessage -Level Warn -Message 'ImpactIQ completed with errors (exit code 2) - no workbook was produced (Assemble did not run); check the Failures sheet and the log.' }
        }
        3 {
            $paused = @()
            $stageMap = Get-IQMemberValue -Object $script:IQ.Manifest -Name 'stages'
            foreach ($n in $script:IQEntryStageOrder) { if ([string](Get-IQMemberValue -Object (Get-IQMemberValue -Object $stageMap -Name $n) -Name 'status') -eq 'Paused') { $paused += $n } }
            $why = 'the time budget was reached in an earlier attempt'
            if ($script:IQ.ContainsKey('BudgetExceeded') -and [bool]$script:IQ['BudgetExceeded']) { $why = ('the time budget of {0} min was reached' -f $TimeBudgetMinutes) }
            Write-IQEntryMessage -Level Warn -Message ('ImpactIQ paused (exit code 3) - {0}; stages still pending: {1}. Partial workbooks were produced. Start it again (same parameters) to resume run {2}.' -f $why, ($paused -join ', '), $script:IQ.RunId)
        }
        default { Write-IQEntryMessage -Level Error -Message ('ImpactIQ ended with exit code {0}.' -f $exitCode) }
    }
}

if ($PassThru -and $script:IQ -and $null -ne $script:IQ.Manifest) { $script:IQ.Manifest }
exit $exitCode
