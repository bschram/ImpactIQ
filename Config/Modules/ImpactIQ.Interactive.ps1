#Requires -Version 5.1
<#
.SYNOPSIS
    ImpactIQ v3 - Interactive module: the WinForms dialogs of the v2 script plus the two selection wrappers the
    entry point calls (Select-IQEnvironmentInteractive, Select-IQScopeInteractive).

.DESCRIPTION
    Loaded by ImpactIQ.ps1 ONLY when the run is interactive ($script:IQ.Interactive). Nothing in here may run headless:
    every wrapper checks Assert-IQInteractiveHost first (Windows + interactive session + not -NonInteractive).

    Moved verbatim from "Final PS Script.txt" (bodies unchanged, only comment-based help and [CmdletBinding()] added):
      - lines 184-222  Read-HostWithTimeout            (+ audit C1-03: returns "" when console input is redirected)
      - lines 226-337  Show-EnvironmentSelectionDialog
      - lines 654-845  Show-WorkspacePicker            (+ audit C3-03: the 10-minute timer is stopped/disposed after ShowDialog;
                                                        + INT-03: the result also carries NothingChecked so the wrappers can tell
                                                        "OK with nothing ticked" from a real selection)
      - lines 850-915  Show-RunModeDialog
      - lines 920-975  Show-ReportScopeDialog
      - lines 980-1148 Show-ReportPicker               (+ audit C3-03: timer stopped/disposed after ShowDialog)
      - lines 1153-1208 Show-ModelScopeDialog
      - lines 1213-1379 Show-ModelPicker               (+ audit C3-03: timer stopped/disposed after ShowDialog)

    New (brief section 2.8): Select-IQEnvironmentInteractive and Select-IQScopeInteractive reproduce the dialog flow of
    monolith lines 340-380 and 1530-1860 but RETURN DATA instead of mutating globals; Cancel at the top level returns
    $null (the entry point exits 0 with "cancelled"). Workspace/report/model listings for the pickers go through
    Invoke-IQApi (ImpactIQ.Http.ps1) and are stored in the per-run inventory cache (Get-IQInventoryCache,
    ImpactIQ.Inventory.ps1) so Resolve-IQScope and the collectors do not list the same workspaces again (INT-02).
    All output goes through Write-IQLog. Windows PowerShell 5.1 compatible.
#>

function Assert-IQInteractiveHost {
    <#
    .SYNOPSIS
        Throws unless dialogs can be shown: Windows, an interactive session ($script:IQ.Interactive) and WinForms loadable.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][string]$Dialog = 'dialog')
    if ($null -eq $script:IQ) { throw "Cannot show the ${Dialog}: the ImpactIQ context is not initialised." }
    if (-not [bool]$script:IQ.Interactive) { throw "Cannot show the $Dialog in a non-interactive run; pass the equivalent parameters instead (see docs/Headless-and-Resume.md)." }
    if (-not [bool]$script:IQ.IsWindows) { throw "Cannot show the ${Dialog}: WinForms dialogs are only available on Windows; pass the equivalent parameters instead." }
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    }
    catch { throw "Cannot show the ${Dialog}: System.Windows.Forms could not be loaded ($($_.Exception.Message)); pass the equivalent parameters instead." }
}

function Select-IQEnvironmentInteractive {
    <#
    .SYNOPSIS
        Shows the environment dialog and returns the environment name (timeout / no selection => 'Public'; Cancel => $null).
    .DESCRIPTION
        Reproduces monolith lines 340-372: the display text is reduced to the environment name ("Public (Commercial)" ->
        Public), unknown values fall back to Public with a warning. The 60-second timeout defaults to Public exactly
        as before; pressing Cancel (or closing the window) now returns $null so the caller can stop the run.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $false)][int]$TimeoutSeconds = 60)
    Assert-IQInteractiveHost -Dialog 'environment dialog'

    $selectedEnv = Show-EnvironmentSelectionDialog -TimeoutSeconds $TimeoutSeconds
    $selectedText = [string]$selectedEnv

    if ($selectedText -eq 'Cancelled') {
        Write-IQLog -Level Warn -Message 'Environment selection cancelled by user.'
        return $null
    }
    if ($selectedText -eq 'Timeout' -or [string]::IsNullOrWhiteSpace($selectedText)) {
        Write-IQLog -Level Warn -Message 'No environment selected or timeout reached - defaulting to Public'
        return 'Public'
    }

    # Extract the environment name from the display text ("EnvironmentName" or "EnvironmentName (Description)")
    $envName = ($selectedText -replace ' \(.*\)', '').Trim()
    $result = 'Public'
    switch ($envName) {
        'Public' { $result = 'Public' }
        'Germany' { $result = 'Germany' }
        'USGovHigh' { $result = 'USGovHigh' }
        'USGovMil' { $result = 'USGovMil' }
        'USGov' { $result = 'USGov' }
        'China' { $result = 'China' }
        default {
            Write-IQLog -Level Warn -Message "Unrecognized environment '$selectedText'. Defaulting to 'Public'."
            $result = 'Public'
        }
    }
    Write-IQLog -Level Success -Message "Selected environment: $result"
    return $result
}

function ConvertTo-IQPickerWorkspaceRow {
    <#
    .SYNOPSIS
        Normalises a workspace row (raw API object or renamed inventory row) to the {id; name} shape the pickers read.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Workspace)
    $id = Get-IQMemberValue -Object $Workspace -Name 'id'
    if ($null -eq $id -or [string]::IsNullOrWhiteSpace([string]$id)) { $id = Get-IQMemberValue -Object $Workspace -Name 'WorkspaceId' }
    $name = Get-IQMemberValue -Object $Workspace -Name 'name'
    if ($null -eq $name -or [string]::IsNullOrWhiteSpace([string]$name)) { $name = Get-IQMemberValue -Object $Workspace -Name 'WorkspaceName' }
    return [pscustomobject]@{ id = [string]$id; name = [string]$name }
}

function Get-IQInteractiveCachedItemList {
    <#
    .SYNOPSIS
        Reads or stores a workspace's raw /reports or /datasets list in the per-run inventory cache (INT-02); no-op when the Inventory module is not loaded.
    .DESCRIPTION
        Without -Value: returns the cached array for the workspace, or $null when there is no entry. With -Value: stores
        the array. Only real responses are ever stored (INV-01: a failed call leaves the key absent so the Inventory
        stage re-fetches with error bookkeeping instead of treating a 4xx as an empty workspace).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$WorkspaceId,
        [Parameter(Mandatory = $true)][ValidateSet('Reports', 'Datasets')][string]$Kind,
        [Parameter(Mandatory = $false)][AllowNull()][AllowEmptyCollection()][array]$Value
    )
    if (-not (Get-Command -Name Get-IQInventoryCache -ErrorAction SilentlyContinue)) { return $null }
    $cache = Get-IQInventoryCache
    if ($null -eq $cache -or -not $cache.ContainsKey($Kind) -or $null -eq $cache[$Kind]) { return $null }
    $table = $cache[$Kind]
    if ($PSBoundParameters.ContainsKey('Value')) {
        if ($null -ne $Value) { $table[$WorkspaceId] = @($Value) }
        return $null
    }
    if ($table.ContainsKey($WorkspaceId)) { return , @($table[$WorkspaceId]) }
    return $null
}

function Get-IQInteractiveWorkspaceItemList {
    <#
    .SYNOPSIS
        Lists reports or datasets of the given workspaces as picker rows (monolith 1590-1614 / 1715-1737); failures per workspace are logged and skipped.
    .DESCRIPTION
        INT-02: each successful listing is stored in the per-run inventory cache (Get-IQInventoryCache) and an existing
        cache entry is reused, so Resolve-IQScope / Get-IQWorkspaceInventory (and a second pass through the pickers)
        do not list the same workspaces again.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Workspaces,
        [Parameter(Mandatory = $true)][ValidateSet('Reports', 'Models')][string]$Kind
    )
    $items = New-Object System.Collections.Generic.List[object]
    $total = @($Workspaces).Count
    $n = 0
    $label = 'reports'
    $cacheKind = 'Reports'
    $endpoint = 'reports'
    if ($Kind -eq 'Models') { $label = 'models'; $cacheKind = 'Datasets'; $endpoint = 'datasets' }
    Write-IQLog -Level Info -Message ("Fetching {0} for selection from {1} workspace(s)..." -f $label, $total)
    foreach ($ws in $Workspaces) {
        $n++
        $wsId = [string]$ws.id
        Write-IQLog -Level Debug -Message ("Scanning workspace {0} of {1}: {2}" -f $n, $total, $ws.name)
        try {
            $rows = Get-IQInteractiveCachedItemList -WorkspaceId $wsId -Kind $cacheKind
            if ($null -eq $rows) {
                $resp = Invoke-IQApi -Method GET -Path ("groups/{0}/{1}" -f $wsId, $endpoint) -AllowNotFound
                if ($null -ne $resp) {
                    $rows = @()
                    if ($null -ne (Get-IQMemberValue -Object $resp -Name 'value')) { $rows = @($resp.value | Where-Object { $null -ne $_ }) }
                    Get-IQInteractiveCachedItemList -WorkspaceId $wsId -Kind $cacheKind -Value $rows | Out-Null
                }
            }
            if ($Kind -eq 'Reports') {
                foreach ($rpt in @($rows)) {
                    if ($null -eq $rpt) { continue }
                    $items.Add([pscustomobject]@{
                            ReportId           = [string]$rpt.id
                            ReportName         = [string]$rpt.name
                            WorkspaceId        = $wsId
                            WorkspaceName      = [string]$ws.name
                            DatasetId          = [string](Get-IQMemberValue -Object $rpt -Name 'datasetId')
                            DatasetWorkspaceId = [string](Get-IQMemberValue -Object $rpt -Name 'datasetWorkspaceId')
                        })
                }
            }
            else {
                foreach ($ds in @($rows)) {
                    if ($null -eq $ds) { continue }
                    $items.Add([pscustomobject]@{
                            DatasetId     = [string]$ds.id
                            DatasetName   = [string]$ds.name
                            WorkspaceId   = $wsId
                            WorkspaceName = [string]$ws.name
                        })
                }
            }
        }
        catch {
            # Monolith swallowed these silently ("catch { }"); keep going but say why (audit C3-07 / C4-03).
            Write-IQLog -Level Warn -Item $ws.name -Message ("Could not list {0} of workspace '{1}': {2}" -f $label, $ws.name, $_.Exception.Message)
        }
    }
    Write-IQLog -Level Info -Message ("Found {0} {1} across {2} workspace(s)." -f $items.Count, $label, $total)
    return $items.ToArray()
}

function Test-IQInteractiveNothingChecked {
    <#
    .SYNOPSIS
        True when a Show-WorkspacePicker result says nothing was ticked (INT-03); falls back to "no ids" for a result without the flag.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][AllowNull()]$Selection,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$SelectedIds
    )
    $flag = Get-IQMemberValue -Object $Selection -Name 'NothingChecked'
    if ($null -ne $flag) { return [bool]$flag }
    return (@($SelectedIds).Count -eq 0)
}

function Select-IQInteractiveWorkspaceScope {
    <#
    .SYNOPSIS
        "Specific" branch of the report/model scope flow: shows the workspace picker; returns @{Cancelled; WorkspaceIds; Workspaces; TimedOut} (nothing ticked / timeout => all, as the monolith did).
    .DESCRIPTION
        INT-03: when nothing was ticked the result is WorkspaceIds = @() with every workspace in Workspaces ("all"), so
        the manifest does not carry an explicit list of every workspace id. INT-04: TimedOut reports the picker's
        10-minute timeout so the callers can log and persist why every workspace was scanned.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Workspaces,
        [Parameter(Mandatory = $true)][string]$ItemLabel
    )
    $result = @{ Cancelled = $false; WorkspaceIds = @(); Workspaces = @($Workspaces); TimedOut = $false }
    $scopeSelection = $null
    try {
        $scopeSelection = Show-WorkspacePicker -Workspaces $Workspaces
    }
    catch {
        # Only the picker's own cancel message means "cancel"; anything else is a real error (audit C4-03).
        if ($_.Exception.Message -notlike 'User cancelled*') { throw }
        $result.Cancelled = $true
        return $result
    }
    $result.TimedOut = [bool](Get-IQMemberValue -Object $scopeSelection -Name 'TimedOut')
    $scopeWorkspaceIds = @($scopeSelection.SelectedWorkspaceIds | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
    if ($result.TimedOut -and $scopeWorkspaceIds.Count -eq 0) {
        Write-IQLog -Level Warn -Message ("Workspace picker timed out - showing {0} from ALL workspaces (legacy behaviour)." -f $ItemLabel)
        return $result
    }
    if (Test-IQInteractiveNothingChecked -Selection $scopeSelection -SelectedIds $scopeWorkspaceIds) {
        # The verbatim picker substitutes every workspace id for an empty selection; report that as "all" (empty ids)
        # instead of persisting hundreds of explicit ids in manifest.scope (INT-03).
        Write-IQLog -Level Warn -Message ("No workspaces selected - showing {0} from ALL workspaces instead." -f $ItemLabel)
        return $result
    }
    $result.WorkspaceIds = $scopeWorkspaceIds
    $result.Workspaces = @($Workspaces | Where-Object { $scopeWorkspaceIds -contains [string]$_.id })
    return $result
}

function Select-IQScopeInteractive {
    <#
    .SYNOPSIS
        Runs the v2 dialog flow (run mode -> scope -> pickers) and returns @{RunMode; WorkspaceIds; IncludeMyWorkspace; ReportIds; DatasetIds; TimedOut}, or $null when cancelled.
    .DESCRIPTION
        Reproduces monolith lines 1530-1860 without globals: every sub-dialog Cancel steps back ONE level (scope dialog
        -> run-mode dialog, picker -> scope dialog), Cancel on the run-mode dialog returns $null. Reports mode returns
        the picked report ids (and their dataset ids) plus the workspace ids that were scanned (empty = all workspaces);
        Models mode returns the picked dataset ids the same way; Workspaces mode returns the picked workspace ids and the
        "Include My Workspace" flag (a timeout selects every workspace plus My Workspace, as before). Resolve-IQScope
        (ImpactIQ.Inventory.ps1) turns the result into the final scope exactly like the headless parameters.
    .PARAMETER Workspaces
        Accessible workspaces: raw API objects (id, name) or renamed inventory rows (WorkspaceId, WorkspaceName).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Workspaces)
    Assert-IQInteractiveHost -Dialog 'scope dialogs'

    $pickerWorkspaces = @()
    foreach ($w in @($Workspaces)) {
        if ($null -eq $w) { continue }
        $row = ConvertTo-IQPickerWorkspaceRow -Workspace $w
        if (-not [string]::IsNullOrWhiteSpace($row.id)) { $pickerWorkspaces += $row }
    }
    if ($pickerWorkspaces.Count -eq 0) {
        # A missing list must not look like "user cancelled" (audit C3-07 / C4-03).
        throw 'No workspaces are available for selection (the workspace listing returned nothing or access is missing).'
    }

    $result = $null
    $reportPickerCancelled = $true
    while ($reportPickerCancelled) {
        $reportPickerCancelled = $false
        $runMode = [string](Show-RunModeDialog)

        if ($runMode -eq 'Cancel' -or [string]::IsNullOrWhiteSpace($runMode)) {
            Write-IQLog -Level Warn -Message 'Run cancelled by user.'
            return $null
        }

        if ($runMode -eq 'Reports') {
            # ---- Reports Mode: choose scope, optionally pick workspaces, then pick reports ----
            $reportSelectionComplete = $false
            while (-not $reportSelectionComplete) {
                $reportScope = [string](Show-ReportScopeDialog)
                if ($reportScope -eq 'Cancel' -or [string]::IsNullOrWhiteSpace($reportScope)) {
                    Write-IQLog -Level Info -Message 'Report scope selection cancelled. Returning to run mode selection...'
                    $reportPickerCancelled = $true
                    break
                }

                $scopeWorkspaceIds = @()
                $workspacesToScan = $pickerWorkspaces
                $scopeTimedOut = $false
                if ($reportScope -eq 'Specific') {
                    $ws = Select-IQInteractiveWorkspaceScope -Workspaces $pickerWorkspaces -ItemLabel 'reports'
                    if ($ws.Cancelled) {
                        Write-IQLog -Level Info -Message 'Workspace selection cancelled. Returning to report scope selection...'
                        continue
                    }
                    $scopeWorkspaceIds = @($ws.WorkspaceIds)
                    $workspacesToScan = @($ws.Workspaces)
                    $scopeTimedOut = [bool]$ws.TimedOut   # INT-04
                }

                $allReportsForPicker = @(Get-IQInteractiveWorkspaceItemList -Workspaces $workspacesToScan -Kind Reports)
                if ($allReportsForPicker.Count -eq 0) {
                    Write-IQLog -Level Warn -Message 'No reports found in the selected workspace(s). Returning to report scope selection...'
                    continue
                }

                $selectedReports = Show-ReportPicker -Reports $allReportsForPicker
                if ($null -eq $selectedReports) {
                    Write-IQLog -Level Info -Message 'Report selection cancelled. Returning to report scope selection...'
                    continue
                }
                $selectedReports = @($selectedReports)
                if ($selectedReports.Count -eq 0) {
                    Write-IQLog -Level Warn -Message 'No reports selected. Returning to report scope selection...'
                    continue
                }

                $reportIds = @()
                $datasetIds = @()
                foreach ($rpt in $selectedReports) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$rpt.ReportId)) { $reportIds += [string]$rpt.ReportId }
                    if (-not [string]::IsNullOrWhiteSpace([string]$rpt.DatasetId)) { $datasetIds += [string]$rpt.DatasetId }
                    if ($rpt.DatasetWorkspaceId -and $rpt.DatasetWorkspaceId -ne $rpt.WorkspaceId) {
                        Write-IQLog -Level Info -Message "Report '$($rpt.ReportName)' uses a remote model in workspace $($rpt.DatasetWorkspaceId) - that workspace will be included."
                    }
                }
                $result = @{
                    RunMode            = 'Reports'
                    WorkspaceIds       = @($scopeWorkspaceIds | Select-Object -Unique)
                    IncludeMyWorkspace = $false
                    ReportIds          = @($reportIds | Select-Object -Unique)
                    DatasetIds         = @($datasetIds | Select-Object -Unique)
                    TimedOut           = $scopeTimedOut
                }
                Write-IQLog -Level Info -Message ("Selected {0} report(s) ({1} model(s))." -f $result.ReportIds.Count, $result.DatasetIds.Count)
                $reportSelectionComplete = $true
            }
            if ($reportPickerCancelled) { continue }
        }
        elseif ($runMode -eq 'Models') {
            # ---- Models Mode: choose scope, optionally pick workspaces, then pick models ----
            $modelSelectionComplete = $false
            while (-not $modelSelectionComplete) {
                $modelScope = [string](Show-ModelScopeDialog)
                if ($modelScope -eq 'Cancel' -or [string]::IsNullOrWhiteSpace($modelScope)) {
                    Write-IQLog -Level Info -Message 'Model scope selection cancelled. Returning to run mode selection...'
                    $reportPickerCancelled = $true
                    break
                }

                $scopeWorkspaceIds = @()
                $workspacesToScan = $pickerWorkspaces
                $scopeTimedOut = $false
                if ($modelScope -eq 'Specific') {
                    $ws = Select-IQInteractiveWorkspaceScope -Workspaces $pickerWorkspaces -ItemLabel 'models'
                    if ($ws.Cancelled) {
                        Write-IQLog -Level Info -Message 'Workspace selection cancelled. Returning to model scope selection...'
                        continue
                    }
                    $scopeWorkspaceIds = @($ws.WorkspaceIds)
                    $workspacesToScan = @($ws.Workspaces)
                    $scopeTimedOut = [bool]$ws.TimedOut   # INT-04
                }

                $allModelsForPicker = @(Get-IQInteractiveWorkspaceItemList -Workspaces $workspacesToScan -Kind Models)
                if ($allModelsForPicker.Count -eq 0) {
                    Write-IQLog -Level Warn -Message 'No models found in the selected workspace(s). Returning to model scope selection...'
                    continue
                }

                $selectedModels = Show-ModelPicker -Models $allModelsForPicker
                if ($null -eq $selectedModels) {
                    Write-IQLog -Level Info -Message 'Model selection cancelled. Returning to model scope selection...'
                    continue
                }
                $selectedModels = @($selectedModels)
                if ($selectedModels.Count -eq 0) {
                    Write-IQLog -Level Warn -Message 'No models selected. Returning to model scope selection...'
                    continue
                }

                $datasetIds = @()
                foreach ($mdl in $selectedModels) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$mdl.DatasetId)) { $datasetIds += [string]$mdl.DatasetId }
                }
                $result = @{
                    RunMode            = 'Models'
                    WorkspaceIds       = @($scopeWorkspaceIds | Select-Object -Unique)
                    IncludeMyWorkspace = $false
                    ReportIds          = @()
                    DatasetIds         = @($datasetIds | Select-Object -Unique)
                    TimedOut           = $scopeTimedOut
                }
                Write-IQLog -Level Info -Message ("Selected {0} model(s); connected reports in every accessible workspace are included automatically." -f $result.DatasetIds.Count)
                $modelSelectionComplete = $true
            }
            if ($reportPickerCancelled) { continue }
        }
        else {
            # ---- Workspaces Mode: existing behaviour ----
            $workspaceSelectionDone = $false
            while (-not $workspaceSelectionDone) {
                $selection = $null
                try {
                    $selection = Show-WorkspacePicker -Workspaces $pickerWorkspaces
                }
                catch {
                    if ($_.Exception.Message -notlike 'User cancelled*') { throw }
                    Write-IQLog -Level Info -Message 'Workspace selection cancelled. Returning to run mode selection...'
                    $reportPickerCancelled = $true
                    break
                }

                $selectedWorkspaceIds = @($selection.SelectedWorkspaceIds | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ })
                $includeMy = [bool]$selection.IncludeMyWorkspace
                $timedOut = [bool](Get-IQMemberValue -Object $selection -Name 'TimedOut')
                $nothingChecked = Test-IQInteractiveNothingChecked -Selection $selection -SelectedIds $selectedWorkspaceIds
                if ($timedOut) {
                    # INT-01: the picker forces "Include My Workspace" on timeout, so it never substitutes the workspace
                    # ids itself; do it here so the returned data really is "every accessible workspace + My Workspace"
                    # (what the log, the manifest and both consumers say a timeout means).
                    if ($selectedWorkspaceIds.Count -eq 0) {
                        $selectedWorkspaceIds = @($pickerWorkspaces | ForEach-Object { [string]$_.id })
                        Write-IQLog -Level Warn -Message 'Workspace picker timed out - running against every accessible workspace plus My Workspace (legacy behaviour).'
                    }
                    else {
                        Write-IQLog -Level Warn -Message ("Workspace picker timed out - running against the {0} ticked workspace(s) plus My Workspace (legacy behaviour)." -f $selectedWorkspaceIds.Count)
                    }
                }
                elseif ($nothingChecked -and -not $includeMy) {
                    # Audit C3-06 / INT-03: the verbatim picker turns "OK with nothing ticked" into every workspace id;
                    # an explicit empty selection must not widen the scope silently (INV-08) - ask again instead.
                    # "My Workspace only" (nothing checked + Include My Workspace) stays a valid selection.
                    Write-IQLog -Level Warn -Message 'No workspaces selected. Please select at least one workspace (or tick "Include My Workspace").'
                    continue
                }
                elseif ($nothingChecked) {
                    $selectedWorkspaceIds = @()
                }

                $result = @{
                    RunMode            = 'Workspaces'
                    WorkspaceIds       = @($selectedWorkspaceIds | Select-Object -Unique)
                    IncludeMyWorkspace = $includeMy
                    ReportIds          = @()
                    DatasetIds         = @()
                    TimedOut           = $timedOut
                }
                Write-IQLog -Level Info -Message ("Selected {0} workspace(s); My Workspace: {1}" -f $result.WorkspaceIds.Count, $includeMy)
                $workspaceSelectionDone = $true
            }
            if ($reportPickerCancelled) { continue }
        }
    }
    return $result
}

# =====================================================================================================================
# Verbatim WinForms functions from "Final PS Script.txt" (see the file header for the source lines and the two
# audit-driven one-liners). Bodies are otherwise unchanged on purpose.
# =====================================================================================================================

function Read-HostWithTimeout {
    <#
    .SYNOPSIS
        Console prompt with a timeout (legacy helper, verbatim from the monolith; returns "" when the timeout elapses or when no console input is available).
    #>
    [CmdletBinding()]
    param(
        [string]$Prompt,
        [int]$TimeoutSeconds = 60
    )

    # Audit C1-03: [Console]::KeyAvailable throws when input is redirected / no console; behave like the timeout instead.
    try { if ([Console]::IsInputRedirected) { return "" } } catch { return "" }
    Write-Host -NoNewline "$Prompt "
    Write-Host "(waiting up to $TimeoutSeconds seconds):" -ForegroundColor DarkGray

    $inputBuffer = ""
    $endTime = (Get-Date).AddSeconds($TimeoutSeconds)

    while ((Get-Date) -lt $endTime) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            # Enter pressed - stop reading
            if ($key.Key -eq "Enter") {
                Write-Host ""
                return $inputBuffer
            }
            # Backspace
            elseif ($key.Key -eq "Backspace") {
                if ($inputBuffer.Length -gt 0) {
                    $inputBuffer = $inputBuffer.Substring(0, $inputBuffer.Length - 1)
                    Write-Host -NoNewline "`b `b"
                }
            }
            else {
                $inputBuffer += $key.KeyChar
                Write-Host -NoNewline $key.KeyChar
            }
        }
        Start-Sleep -Milliseconds 100
    }

    Write-Host "Timeout reached - defaulting to Public" -ForegroundColor Yellow
    return ""
}

function Show-EnvironmentSelectionDialog {
    <#
    .SYNOPSIS
        WinForms list dialog for the Power BI cloud environment (verbatim from the monolith; returns the display text, "Timeout" or "Cancelled").
    #>
    [CmdletBinding()]
    param(
        [int]$TimeoutSeconds = 60
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Power BI Environment"
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(400, 320)
    $form.TopMost = $true
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    # Label
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select your Power BI cloud environment:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($label)

    # ListBox for environment selection
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(20, 50)
    $listBox.Size = New-Object System.Drawing.Size(340, 150)
    $listBox.SelectionMode = 'One'
    
    # Add environment options
    $environments = @(
        'Public (Commercial)',
        'Germany',
        'USGov',
        'China',
        'USGovHigh',
        'USGovMil'
    )
    
    foreach ($env in $environments) {
        [void]$listBox.Items.Add($env)
    }
    
    # Set default selection to Public
    $listBox.SelectedIndex = 0
    $form.Controls.Add($listBox)

    # Timeout label
    $timeoutLabel = New-Object System.Windows.Forms.Label
    $timeoutLabel.Text = "Timeout in $TimeoutSeconds seconds (defaults to Public)"
    $timeoutLabel.AutoSize = $true
    $timeoutLabel.Location = New-Object System.Drawing.Point(20, 210)
    $timeoutLabel.ForeColor = [System.Drawing.Color]::Gray
    $form.Controls.Add($timeoutLabel)

    # OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Size = New-Object System.Drawing.Size(75, 30)
    $okButton.Location = New-Object System.Drawing.Point(205, 240)
    $okButton.Add_Click({
        $form.Tag = $listBox.SelectedItem
        $form.Close()
    })
    $form.Controls.Add($okButton)

    # Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Size = New-Object System.Drawing.Size(75, 30)
    $cancelButton.Location = New-Object System.Drawing.Point(285, 240)
    $cancelButton.Add_Click({
        $form.Tag = 'Cancelled'
        $form.Close()
    })
    $form.Controls.Add($cancelButton)

    # Timer for timeout
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = $TimeoutSeconds * 1000
    $timer.Add_Tick({
        $timer.Stop()
        $form.Tag = 'Timeout'
        $form.Close()
    })

    # Add FormClosing event handler
    $form.Add_FormClosing({
        param($sender, $e)
        $timer.Stop()
        $timer.Dispose()
        
        # If Tag is not set (form closed without button click), mark as cancelled
        if (-not $form.Tag) {
            $form.Tag = 'Cancelled'
        }
    })

    # Handle double-click on list item (same as OK button)
    $listBox.Add_DoubleClick({
        $form.Tag = $listBox.SelectedItem
        $form.Close()
    })

    # Set default button and show
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $timer.Start()
    [void]$form.ShowDialog()

    return $form.Tag
}

function Show-WorkspacePicker {
    <#
    .SYNOPSIS
        WinForms workspace picker with search, Select All / Clear All and "Include My Workspace" (verbatim; throws on Cancel, 10-minute timeout).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Workspaces # array of PSCustomObjects with id + name
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Form
    $form               = New-Object System.Windows.Forms.Form
    $form.Text          = "Select Workspaces"
    $form.StartPosition = 'CenterScreen'
    $form.Size          = New-Object System.Drawing.Size(520,580)
    $form.TopMost       = $true

    # Instruction label
    $lbl                = New-Object System.Windows.Forms.Label
    $lbl.Text           = "Select the workspaces to run against:"
    $lbl.AutoSize       = $true
    $lbl.Location       = New-Object System.Drawing.Point(12,12)
    $form.Controls.Add($lbl)

    # Search label
    $lblSearch          = New-Object System.Windows.Forms.Label
    $lblSearch.Text     = "Search:"
    $lblSearch.AutoSize = $true
    $lblSearch.Location = New-Object System.Drawing.Point(12,40)
    $form.Controls.Add($lblSearch)

    # Search TextBox
    $txtSearch          = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(70,37)
    $txtSearch.Size     = New-Object System.Drawing.Size(422,23)
    $form.Controls.Add($txtSearch)

    # CheckedListBox
    $clb                = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location       = New-Object System.Drawing.Point(12,70)
    $clb.Size           = New-Object System.Drawing.Size(480,380)
    $clb.CheckOnClick   = $true
    $clb.Sorted         = $true

    # Store all workspace items for filtering
    $allWorkspaceItems = @()
    foreach ($ws in $Workspaces) {
        $display = "{0} ({1})" -f $ws.name, $ws.id
        $allWorkspaceItems += [pscustomobject]@{ Display=$display; Id=$ws.id; Name=$ws.name }
    }

    # Maintain persistent checked state across searches
    $persistentCheckedIds = @{}

    # Initial population
    foreach ($item in $allWorkspaceItems) {
        [void]$clb.Items.Add($item)
    }
    $clb.DisplayMember = 'Display'
    $form.Controls.Add($clb)

    # Track when items are checked/unchecked
    $clb.Add_ItemCheck({
        param($sender, $e)
        $item = $clb.Items[$e.Index]
        if ($e.NewValue -eq 'Checked') {
            $persistentCheckedIds[$item.Id] = $true
        } elseif ($e.NewValue -eq 'Unchecked') {
            $persistentCheckedIds.Remove($item.Id)
        }
    })

    # Search filter logic
    $txtSearch.Add_TextChanged({
        $searchText = $txtSearch.Text
        $clb.BeginUpdate()
        
        $clb.Items.Clear()
        
        # Escape search text once for reuse in loop
        $escapedSearchText = [regex]::Escape($searchText)
        
        # Filter and re-add items
        foreach ($item in $allWorkspaceItems) {
            if ([string]::IsNullOrWhiteSpace($searchText) -or 
                $item.Display -imatch $escapedSearchText) {
                $index = $clb.Items.Add($item)
                # Restore checked state from persistent store
                if ($persistentCheckedIds.ContainsKey($item.Id)) {
                    $clb.SetItemChecked($index, $true)
                }
            }
        }
        
        $clb.EndUpdate()
    })

    # "Select All" checkbox
    $chkAll            = New-Object System.Windows.Forms.CheckBox
    $chkAll.Text       = "Select All"
    $chkAll.AutoSize   = $true
    $chkAll.Location   = New-Object System.Drawing.Point(12,460)
    $chkAll.Add_CheckedChanged({
        for ($i=0; $i -lt $clb.Items.Count; $i++) {
            $clb.SetItemChecked($i, $chkAll.Checked)
        }
    })
    $form.Controls.Add($chkAll)

    # "Clear All" button
    $btnClearAll       = New-Object System.Windows.Forms.Button
    $btnClearAll.Text  = "Clear All"
    $btnClearAll.Width = 80
    $btnClearAll.Location = New-Object System.Drawing.Point(90,458)
    $btnClearAll.Add_Click({
        # Uncheck all items - ItemCheck events will update persistent store
        for ($i=0; $i -lt $clb.Items.Count; $i++) {
            $clb.SetItemChecked($i, $false)
        }
        $chkAll.Checked = $false
    })
    $form.Controls.Add($btnClearAll)

    # "Include My Workspace" checkbox
    $chkMy             = New-Object System.Windows.Forms.CheckBox
    $chkMy.Text        = "Include 'My Workspace'"
    $chkMy.AutoSize    = $true
    $chkMy.Location    = New-Object System.Drawing.Point(184,460)
    $chkMy.Checked     = $false
    $form.Controls.Add($chkMy)

    # OK button
    $okBtn             = New-Object System.Windows.Forms.Button
    $okBtn.Text        = "OK"
    $okBtn.Width       = 100
    $okBtn.Location    = New-Object System.Drawing.Point(286,500)
    $okBtn.Add_Click({
        $form.Tag = 'OK'
        $form.Close()
    })
    $form.Controls.Add($okBtn)

    # Cancel button
    $cancelBtn         = New-Object System.Windows.Forms.Button
    $cancelBtn.Text    = "Cancel"
    $cancelBtn.Width   = 100
    $cancelBtn.Location= New-Object System.Drawing.Point(392,500)
    $cancelBtn.Add_Click({
        $form.Tag = 'Cancel'
        $form.Close()
    })
    $form.Controls.Add($cancelBtn)

    # Timer (auto-close after 10 minutes)
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 600000
    $timer.Add_Tick({
        $timer.Stop()
        $form.Tag = 'Timeout'
        $form.Close()
    })
    $timer.Start()

    [void]$form.ShowDialog()

    # Audit C3-03: a started WinForms timer outlives the dialog and would close the NEXT dialog; stop it here.
    try { $timer.Stop(); $timer.Dispose() } catch { $null = $_ }

    # Gather selections from persistent store (to include items not currently visible due to search)
    $selectedIds = @($persistentCheckedIds.Keys)
    # INT-03: remember whether the user ticked anything BEFORE the "empty => all workspaces" substitution below.
    $nothingChecked = ($selectedIds.Count -eq 0)

    # Treat the Cancel button AND closing the window (the X) as a cancel so the
    # caller returns to the run-mode menu. Only an explicit OK proceeds, and the
    # timer still auto-continues on Timeout.
    if ($form.Tag -ne 'OK' -and $form.Tag -ne 'Timeout') {
        throw "User cancelled workspace selection."
    }

    $timedOut = ($form.Tag -eq 'Timeout')

    # Include My Workspace only if checked... except on timeout, then force include
    $includeMy = if ($timedOut) { $true } else { $chkMy.Checked }

    # If nothing selected from workspace list:
    # - If "Include My Workspace" is checked, process only My Workspace (empty array for regular workspaces)
    # - Otherwise, default to all workspaces (backwards compatibility with timeout/cancel behavior)
    if ($selectedIds.Count -eq 0 -and -not $includeMy) {
        $selectedIds = $Workspaces.id
    }

    return [pscustomobject]@{
        SelectedWorkspaceIds = $selectedIds
        IncludeMyWorkspace   = $includeMy
        TimedOut             = $timedOut
        NothingChecked       = $nothingChecked
    }
}

function Show-RunModeDialog {
    <#
    .SYNOPSIS
        WinForms run-mode dialog (verbatim): returns "Workspaces", "Reports", "Models" or "Cancel".
    #>
    [CmdletBinding()]
    param()
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "ImpactIQ - Select Run Mode"
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(440, 300)
    $form.TopMost = $true
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "How would you like to run the extraction?"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($label)

    $btnWorkspaces = New-Object System.Windows.Forms.Button
    $btnWorkspaces.Text = "Run against workspace(s)"
    $btnWorkspaces.Size = New-Object System.Drawing.Size(380, 40)
    $btnWorkspaces.Location = New-Object System.Drawing.Point(20, 60)
    $btnWorkspaces.Add_Click({
        $form.Tag = 'Workspaces'
        $form.Close()
    })
    $form.Controls.Add($btnWorkspaces)

    $btnReports = New-Object System.Windows.Forms.Button
    $btnReports.Text = "Run against specific reports"
    $btnReports.Size = New-Object System.Drawing.Size(380, 40)
    $btnReports.Location = New-Object System.Drawing.Point(20, 110)
    $btnReports.Add_Click({
        $form.Tag = 'Reports'
        $form.Close()
    })
    $form.Controls.Add($btnReports)

    $btnModels = New-Object System.Windows.Forms.Button
    $btnModels.Text = "Run against specific models"
    $btnModels.Size = New-Object System.Drawing.Size(380, 40)
    $btnModels.Location = New-Object System.Drawing.Point(20, 160)
    $btnModels.Add_Click({
        $form.Tag = 'Models'
        $form.Close()
    })
    $form.Controls.Add($btnModels)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(380, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(20, 210)
    $btnCancel.Add_Click({
        $form.Tag = 'Cancel'
        $form.Close()
    })
    $form.Controls.Add($btnCancel)

    $form.Add_FormClosing({
        param($sender, $e)
        if (-not $form.Tag) { $form.Tag = 'Cancel' }
    })

    [void]$form.ShowDialog()
    return $form.Tag
}

function Show-ReportScopeDialog {
    <#
    .SYNOPSIS
        WinForms report-scope dialog (verbatim): returns "All", "Specific" or "Cancel".
    #>
    [CmdletBinding()]
    param()
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "ImpactIQ - Report Selection Scope"
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(440, 250)
    $form.TopMost = $true
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Which reports would you like to choose from?"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($label)

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = "Reports from ALL workspaces"
    $btnAll.Size = New-Object System.Drawing.Size(380, 40)
    $btnAll.Location = New-Object System.Drawing.Point(20, 55)
    $btnAll.Add_Click({
        $form.Tag = 'All'
        $form.Close()
    })
    $form.Controls.Add($btnAll)

    $btnSpecific = New-Object System.Windows.Forms.Button
    $btnSpecific.Text = "Reports from specific workspace(s)"
    $btnSpecific.Size = New-Object System.Drawing.Size(380, 40)
    $btnSpecific.Location = New-Object System.Drawing.Point(20, 105)
    $btnSpecific.Add_Click({
        $form.Tag = 'Specific'
        $form.Close()
    })
    $form.Controls.Add($btnSpecific)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(380, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(20, 155)
    $btnCancel.Add_Click({
        $form.Tag = 'Cancel'
        $form.Close()
    })
    $form.Controls.Add($btnCancel)

    $form.Add_FormClosing({
        param($sender, $e)
        if (-not $form.Tag) { $form.Tag = 'Cancel' }
    })

    [void]$form.ShowDialog()
    return $form.Tag
}

function Show-ReportPicker {
    <#
    .SYNOPSIS
        WinForms report picker with search (verbatim): returns the checked report objects, or $null on Cancel (2-minute timeout).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Reports
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form               = New-Object System.Windows.Forms.Form
    $form.Text          = "Select Reports"
    $form.StartPosition = 'CenterScreen'
    $form.Size          = New-Object System.Drawing.Size(650,580)
    $form.TopMost       = $true

    # Instruction label
    $lbl                = New-Object System.Windows.Forms.Label
    $lbl.Text           = "Select the reports to run against (connected models will be included automatically):"
    $lbl.AutoSize       = $true
    $lbl.Location       = New-Object System.Drawing.Point(12,12)
    $form.Controls.Add($lbl)

    # Search label
    $lblSearch          = New-Object System.Windows.Forms.Label
    $lblSearch.Text     = "Search:"
    $lblSearch.AutoSize = $true
    $lblSearch.Location = New-Object System.Drawing.Point(12,40)
    $form.Controls.Add($lblSearch)

    # Search TextBox
    $txtSearch          = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(70,37)
    $txtSearch.Size     = New-Object System.Drawing.Size(552,23)
    $form.Controls.Add($txtSearch)

    # CheckedListBox
    $clb                = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location       = New-Object System.Drawing.Point(12,70)
    $clb.Size           = New-Object System.Drawing.Size(610,380)
    $clb.CheckOnClick   = $true
    $clb.Sorted         = $true

    # Build display items
    $allReportItems = @()
    foreach ($rpt in $Reports) {
        $display = "{0} - {1} ({2})" -f $rpt.WorkspaceName, $rpt.ReportName, $rpt.ReportId
        $item = [pscustomobject]@{
            Display            = $display
            ReportId           = $rpt.ReportId
            ReportName         = $rpt.ReportName
            WorkspaceId        = $rpt.WorkspaceId
            WorkspaceName      = $rpt.WorkspaceName
            DatasetId          = $rpt.DatasetId
            DatasetWorkspaceId = $rpt.DatasetWorkspaceId
        }
        $item | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.Display } -Force
        $allReportItems += $item
    }

    # Persistent checked state across search filtering
    $persistentCheckedIds = @{}

    # Initial population
    foreach ($item in $allReportItems) {
        [void]$clb.Items.Add($item)
    }
    $clb.DisplayMember = 'Display'
    $form.Controls.Add($clb)

    # Track check/uncheck
    $clb.Add_ItemCheck({
        param($sender, $e)
        $item = $clb.Items[$e.Index]
        if ($e.NewValue -eq 'Checked') {
            $persistentCheckedIds[$item.ReportId] = $true
        } elseif ($e.NewValue -eq 'Unchecked') {
            $persistentCheckedIds.Remove($item.ReportId)
        }
    })

    # Search filter logic
    $txtSearch.Add_TextChanged({
        $searchText = $txtSearch.Text
        $clb.BeginUpdate()
        $clb.Items.Clear()
        $escapedSearchText = [regex]::Escape($searchText)
        foreach ($item in $allReportItems) {
            if ([string]::IsNullOrWhiteSpace($searchText) -or
                $item.Display -imatch $escapedSearchText) {
                $index = $clb.Items.Add($item)
                if ($persistentCheckedIds.ContainsKey($item.ReportId)) {
                    $clb.SetItemChecked($index, $true)
                }
            }
        }
        $clb.EndUpdate()
    })

    # "Select All" checkbox
    $chkAll            = New-Object System.Windows.Forms.CheckBox
    $chkAll.Text       = "Select All"
    $chkAll.AutoSize   = $true
    $chkAll.Location   = New-Object System.Drawing.Point(12,460)
    $chkAll.Add_CheckedChanged({
        for ($i=0; $i -lt $clb.Items.Count; $i++) {
            $clb.SetItemChecked($i, $chkAll.Checked)
        }
    })
    $form.Controls.Add($chkAll)

    # "Clear All" button
    $btnClearAll       = New-Object System.Windows.Forms.Button
    $btnClearAll.Text  = "Clear All"
    $btnClearAll.Width = 80
    $btnClearAll.Location = New-Object System.Drawing.Point(100,458)
    $btnClearAll.Add_Click({
        for ($i=0; $i -lt $clb.Items.Count; $i++) {
            $clb.SetItemChecked($i, $false)
        }
        $chkAll.Checked = $false
    })
    $form.Controls.Add($btnClearAll)

    # OK button
    $okBtn             = New-Object System.Windows.Forms.Button
    $okBtn.Text        = "OK"
    $okBtn.Width       = 100
    $okBtn.Location    = New-Object System.Drawing.Point(416,500)
    $okBtn.Add_Click({
        $form.Tag = 'OK'
        $form.Close()
    })
    $form.Controls.Add($okBtn)

    # Cancel button
    $cancelBtn         = New-Object System.Windows.Forms.Button
    $cancelBtn.Text    = "Cancel"
    $cancelBtn.Width   = 100
    $cancelBtn.Location= New-Object System.Drawing.Point(522,500)
    $cancelBtn.Add_Click({
        $form.Tag = 'Cancel'
        $form.Close()
    })
    $form.Controls.Add($cancelBtn)

    # Timer (auto-close after 2 minutes)
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 120000
    $timer.Add_Tick({
        $timer.Stop()
        $form.Tag = 'Timeout'
        $form.Close()
    })
    $timer.Start()

    $form.AcceptButton = $okBtn
    $form.CancelButton = $cancelBtn
    [void]$form.ShowDialog()

    # Audit C3-03: a started WinForms timer outlives the dialog and would close the NEXT dialog; stop it here.
    try { $timer.Stop(); $timer.Dispose() } catch { $null = $_ }

    if ($form.Tag -eq 'Cancel') {
        return $null
    }

    # Gather selected reports from persistent store
    $selectedReportIds = @($persistentCheckedIds.Keys)
    $selectedReports = @($allReportItems | Where-Object { $selectedReportIds -contains $_.ReportId })

    return $selectedReports
}

function Show-ModelScopeDialog {
    <#
    .SYNOPSIS
        WinForms model-scope dialog (verbatim): returns "All", "Specific" or "Cancel".
    #>
    [CmdletBinding()]
    param()
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "ImpactIQ - Model Selection Scope"
    $form.StartPosition = 'CenterScreen'
    $form.Size = New-Object System.Drawing.Size(440, 250)
    $form.TopMost = $true
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Which models would you like to choose from?"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $form.Controls.Add($label)

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = "Models from ALL workspaces"
    $btnAll.Size = New-Object System.Drawing.Size(380, 40)
    $btnAll.Location = New-Object System.Drawing.Point(20, 55)
    $btnAll.Add_Click({
        $form.Tag = 'All'
        $form.Close()
    })
    $form.Controls.Add($btnAll)

    $btnSpecific = New-Object System.Windows.Forms.Button
    $btnSpecific.Text = "Models from specific workspace(s)"
    $btnSpecific.Size = New-Object System.Drawing.Size(380, 40)
    $btnSpecific.Location = New-Object System.Drawing.Point(20, 105)
    $btnSpecific.Add_Click({
        $form.Tag = 'Specific'
        $form.Close()
    })
    $form.Controls.Add($btnSpecific)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(380, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(20, 155)
    $btnCancel.Add_Click({
        $form.Tag = 'Cancel'
        $form.Close()
    })
    $form.Controls.Add($btnCancel)

    $form.Add_FormClosing({
        param($sender, $e)
        if (-not $form.Tag) { $form.Tag = 'Cancel' }
    })

    [void]$form.ShowDialog()
    return $form.Tag
}

function Show-ModelPicker {
    <#
    .SYNOPSIS
        WinForms model picker with search (verbatim): returns the checked model objects, or $null on Cancel (2-minute timeout).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Models
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form               = New-Object System.Windows.Forms.Form
    $form.Text          = "Select Models"
    $form.StartPosition = 'CenterScreen'
    $form.Size          = New-Object System.Drawing.Size(650,580)
    $form.TopMost       = $true

    # Instruction label
    $lbl                = New-Object System.Windows.Forms.Label
    $lbl.Text           = "Select the models to run against (reports using these models will be included automatically):"
    $lbl.AutoSize       = $true
    $lbl.Location       = New-Object System.Drawing.Point(12,12)
    $form.Controls.Add($lbl)

    # Search label
    $lblSearch          = New-Object System.Windows.Forms.Label
    $lblSearch.Text     = "Search:"
    $lblSearch.AutoSize = $true
    $lblSearch.Location = New-Object System.Drawing.Point(12,40)
    $form.Controls.Add($lblSearch)

    # Search TextBox
    $txtSearch          = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(70,37)
    $txtSearch.Size     = New-Object System.Drawing.Size(552,23)
    $form.Controls.Add($txtSearch)

    # CheckedListBox
    $clb                = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location       = New-Object System.Drawing.Point(12,70)
    $clb.Size           = New-Object System.Drawing.Size(610,380)
    $clb.CheckOnClick   = $true
    $clb.Sorted         = $true

    # Build display items
    $allModelItems = @()
    foreach ($mdl in $Models) {
        $display = "{0} - {1} ({2})" -f $mdl.WorkspaceName, $mdl.DatasetName, $mdl.DatasetId
        $item = [pscustomobject]@{
            Display       = $display
            DatasetId     = $mdl.DatasetId
            DatasetName   = $mdl.DatasetName
            WorkspaceId   = $mdl.WorkspaceId
            WorkspaceName = $mdl.WorkspaceName
        }
        $item | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.Display } -Force
        $allModelItems += $item
    }

    # Persistent checked state across search filtering
    $persistentCheckedIds = @{}

    # Initial population
    foreach ($item in $allModelItems) {
        [void]$clb.Items.Add($item)
    }
    $clb.DisplayMember = 'Display'
    $form.Controls.Add($clb)

    # Track check/uncheck
    $clb.Add_ItemCheck({
        param($sender, $e)
        $item = $clb.Items[$e.Index]
        if ($e.NewValue -eq 'Checked') {
            $persistentCheckedIds[$item.DatasetId] = $true
        } elseif ($e.NewValue -eq 'Unchecked') {
            $persistentCheckedIds.Remove($item.DatasetId)
        }
    })

    # Search filter logic
    $txtSearch.Add_TextChanged({
        $searchText = $txtSearch.Text
        $clb.BeginUpdate()
        $clb.Items.Clear()
        $escapedSearchText = [regex]::Escape($searchText)
        foreach ($item in $allModelItems) {
            if ([string]::IsNullOrWhiteSpace($searchText) -or
                $item.Display -imatch $escapedSearchText) {
                $index = $clb.Items.Add($item)
                if ($persistentCheckedIds.ContainsKey($item.DatasetId)) {
                    $clb.SetItemChecked($index, $true)
                }
            }
        }
        $clb.EndUpdate()
    })

    # "Select All" checkbox
    $chkAll            = New-Object System.Windows.Forms.CheckBox
    $chkAll.Text       = "Select All"
    $chkAll.AutoSize   = $true
    $chkAll.Location   = New-Object System.Drawing.Point(12,460)
    $chkAll.Add_CheckedChanged({
        for ($i=0; $i -lt $clb.Items.Count; $i++) {
            $clb.SetItemChecked($i, $chkAll.Checked)
        }
    })
    $form.Controls.Add($chkAll)

    # "Clear All" button
    $btnClearAll       = New-Object System.Windows.Forms.Button
    $btnClearAll.Text  = "Clear All"
    $btnClearAll.Width = 80
    $btnClearAll.Location = New-Object System.Drawing.Point(100,458)
    $btnClearAll.Add_Click({
        for ($i=0; $i -lt $clb.Items.Count; $i++) {
            $clb.SetItemChecked($i, $false)
        }
        $chkAll.Checked = $false
    })
    $form.Controls.Add($btnClearAll)

    # OK button
    $okBtn             = New-Object System.Windows.Forms.Button
    $okBtn.Text        = "OK"
    $okBtn.Width       = 100
    $okBtn.Location    = New-Object System.Drawing.Point(416,500)
    $okBtn.Add_Click({
        $form.Tag = 'OK'
        $form.Close()
    })
    $form.Controls.Add($okBtn)

    # Cancel button
    $cancelBtn         = New-Object System.Windows.Forms.Button
    $cancelBtn.Text    = "Cancel"
    $cancelBtn.Width   = 100
    $cancelBtn.Location= New-Object System.Drawing.Point(522,500)
    $cancelBtn.Add_Click({
        $form.Tag = 'Cancel'
        $form.Close()
    })
    $form.Controls.Add($cancelBtn)

    # Timer (auto-close after 2 minutes)
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 120000
    $timer.Add_Tick({
        $timer.Stop()
        $form.Tag = 'Timeout'
        $form.Close()
    })
    $timer.Start()

    $form.AcceptButton = $okBtn
    $form.CancelButton = $cancelBtn
    [void]$form.ShowDialog()

    # Audit C3-03: a started WinForms timer outlives the dialog and would close the NEXT dialog; stop it here.
    try { $timer.Stop(); $timer.Dispose() } catch { $null = $_ }

    if ($form.Tag -eq 'Cancel') {
        return $null
    }

    # Gather selected models from persistent store
    $selectedDatasetIds = @($persistentCheckedIds.Keys)
    $selectedModels = @($allModelItems | Where-Object { $selectedDatasetIds -contains $_.DatasetId })

    return $selectedModels
}
